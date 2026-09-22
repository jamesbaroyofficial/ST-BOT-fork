"use strict";

const fs = require("fs-extra");
const path = require("path");
const crypto = require("crypto");
const request = require("request");
const { spawn } = require("child_process");

const TEMP_ROOT = path.join(
  process.cwd(),
  "tmp",
  "doc2png"
);

const MAX_FILE_SIZE = 100 * 1024 * 1024;
const SEND_DELAY = 700;

const queues = new Map();

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: ["ignore", "pipe", "pipe"]
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", data => {
      stdout += data.toString();
    });

    child.stderr.on("data", data => {
      stderr += data.toString();
    });

    child.on("error", reject);

    child.on("close", code => {
      if (code === 0) {
        resolve({ stdout, stderr });
      } else {
        reject(
          new Error(
            `${command} exited with code ${code}\n${stderr || stdout}`
          )
        );
      }
    });
  });
}

function isSupported(name) {
  return /\.(pptx|docx)$/i.test(name || "");
}

/*
 * Queue documents from the same GC.
 */
function enqueue(threadID, task) {
  const previous =
    queues.get(threadID) || Promise.resolve();

  const next = previous
    .catch(() => {})
    .then(task)
    .finally(() => {
      if (queues.get(threadID) === next) {
        queues.delete(threadID);
      }
    });

  queues.set(threadID, next);

  return next;
}

/*
 * Create an authenticated request cookie jar
 * from ST-FCA's current Facebook session.
 */
function createFacebookJar(api, url) {
  if (
    !api ||
    typeof api.getAppState !== "function"
  ) {
    throw new Error(
      "ST-FCA getAppState() is unavailable."
    );
  }

  const appState = api.getAppState();

  if (
    !Array.isArray(appState) ||
    appState.length === 0
  ) {
    throw new Error(
      "Facebook session cookies are unavailable."
    );
  }

  const jar = request.jar();

  for (const item of appState) {
    if (
      !item ||
      !item.key ||
      item.value === undefined ||
      item.value === null
    ) {
      continue;
    }

    try {
      const cookie = request.cookie(
        `${item.key}=${item.value}`
      );

      /*
       * Do not print cookie values.
       */
      jar.setCookie(cookie, url);
    } catch (error) {
      console.error(
        `[DOC2PNG] Cookie ${item.key} could not be loaded:`,
        error.message
      );
    }
  }

  return jar;
}

/*
 * Download the actual Facebook attachment.
 */
function downloadFile(api, url, output) {
  return new Promise((resolve, reject) => {
    if (!url) {
      return reject(
        new Error("Attachment URL is missing.")
      );
    }

    let jar;

    try {
      jar = createFacebookJar(api, url);
    } catch (error) {
      return reject(error);
    }

    console.log(
      "[DOC2PNG] Downloading with authenticated Facebook session..."
    );

    const userAgent =
      (
        api &&
        api.ctx &&
        api.ctx.globalOptions &&
        api.ctx.globalOptions.userAgent
      ) ||
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36";

    const download = request.get({
      url: url,

      method: "GET",

      jar: jar,

      gzip: true,

      followRedirect: true,

      followAllRedirects: true,

      timeout: 180000,

      encoding: null,

      headers: {
        "User-Agent": userAgent,

        "Accept":
          "application/octet-stream,*/*",

        "Accept-Language":
          "en-US,en;q=0.9",

        "Referer":
          "https://www.facebook.com/",

        "Connection":
          "keep-alive"
      }
    });

    let responseReceived = false;
    let finished = false;

    function fail(error) {
      if (finished) return;

      finished = true;

      try {
        download.destroy();
      } catch (_) {}

      reject(error);
    }

    download.on("response", response => {
      responseReceived = true;

      console.log(
        `[DOC2PNG] HTTP status: ${response.statusCode}`
      );

      console.log(
        `[DOC2PNG] Content-Type: ${
          response.headers["content-type"] ||
          "unknown"
        }`
      );

      console.log(
        `[DOC2PNG] Content-Length: ${
          response.headers["content-length"] ||
          "unknown"
        }`
      );

      if (
        response.statusCode < 200 ||
        response.statusCode >= 400
      ) {
        return fail(
          new Error(
            `Facebook returned HTTP ${response.statusCode}.`
          )
        );
      }

      const contentType =
        String(
          response.headers["content-type"] || ""
        ).toLowerCase();

      /*
       * Do not save Facebook login/error HTML
       * as a PPTX/DOCX.
       */
      if (contentType.includes("text/html")) {
        return fail(
          new Error(
            "Facebook returned HTML instead of the document."
          )
        );
      }

      const writer =
        fs.createWriteStream(output);

      let bytes = 0;

      download.on("data", chunk => {
        bytes += chunk.length;

        if (bytes > MAX_FILE_SIZE) {
          fail(
            new Error(
              "Downloaded file exceeds 100 MB."
            )
          );
        }
      });

      download.on("error", error => {
        writer.destroy();
        fail(error);
      });

      writer.on("error", error => {
        fail(error);
      });

      writer.on("finish", async () => {
        if (finished) return;

        try {
          const stat =
            await fs.stat(output);

          console.log(
            `[DOC2PNG] Downloaded size: ${(stat.size / 1024 / 1024).toFixed(2)} MB`
          );

          if (stat.size <= 0) {
            return fail(
              new Error(
                "Downloaded attachment is 0 bytes."
              )
            );
          }

          if (stat.size > MAX_FILE_SIZE) {
            return fail(
              new Error(
                "Downloaded attachment exceeds 100 MB."
              )
            );
          }

          finished = true;
          resolve(stat.size);
        } catch (error) {
          fail(error);
        }
      });

      download.pipe(writer);
    });

    download.on("error", error => {
      if (!responseReceived) {
        fail(error);
      }
    });

    download.on("abort", () => {
      fail(
        new Error(
          "Facebook attachment download was aborted."
        )
      );
    });
  });
}

/*
 * DOCX/PPTX -> PDF -> PNG
 */
async function convertDocument(inputFile, workDir) {
  const pdfDir =
    path.join(workDir, "pdf");

  const pngDir =
    path.join(workDir, "png");

  const loProfile =
    path.join(workDir, "lo-profile");

  await fs.ensureDir(pdfDir);
  await fs.ensureDir(pngDir);
  await fs.ensureDir(loProfile);

  const profileURL =
    "file://" +
    loProfile.replace(/\\/g, "/");

  await run("libreoffice", [
    "--headless",
    "--nologo",
    "--nodefault",
    "--nofirststartwizard",

    `-env:UserInstallation=${profileURL}`,

    "--convert-to",
    "pdf:impress_pdf_Export",

    "--outdir",
    pdfDir,

    inputFile
  ]);

  const baseName =
    path.basename(
      inputFile,
      path.extname(inputFile)
    );

  const pdfFile =
    path.join(
      pdfDir,
      `${baseName}.pdf`
    );

  if (!await fs.pathExists(pdfFile)) {
    throw new Error(
      "LibreOffice did not create a PDF."
    );
  }

  const pdfStat =
    await fs.stat(pdfFile);

  if (pdfStat.size <= 0) {
    throw new Error(
      "LibreOffice created an empty PDF."
    );
  }

  const info =
    await run("pdfinfo", [pdfFile]);

  const match =
    info.stdout.match(
      /Pages:\s+(\d+)/i
    );

  const pageCount =
    match
      ? parseInt(match[1], 10)
      : 0;

  console.log(
    `[DOC2PNG] PDF pages detected: ${pageCount}`
  );

  if (pageCount < 1) {
    throw new Error(
      "PDF contains no pages."
    );
  }

  await run("pdftoppm", [
    "-png",
    "-r",
    "150",
    "-f",
    "1",
    "-l",
    String(pageCount),
    pdfFile,
    path.join(pngDir, "page")
  ]);

  const files =
    await fs.readdir(pngDir);

  const pages =
    files
      .filter(file =>
        /^page-\d+\.png$/i.test(file)
      )
      .sort((a, b) => {
        const A =
          parseInt(
            a.match(/\d+/)[0],
            10
          );

        const B =
          parseInt(
            b.match(/\d+/)[0],
            10
          );

        return A - B;
      })
      .map(file =>
        path.join(
          pngDir,
          file
        )
      );

  console.log(
    `[DOC2PNG] PNG pages generated: ${pages.length}`
  );

  if (pages.length !== pageCount) {
    throw new Error(
      `Expected ${pageCount} PNG pages but generated ${pages.length}.`
    );
  }

  return pages;
}

/*
 * Send one PNG at a time.
 */
function sendImage(api, file, threadID) {
  return new Promise((resolve, reject) => {
    const stream =
      fs.createReadStream(file);

    stream.on("error", reject);

    api.sendMessage(
      {
        attachment: stream
      },

      threadID,

      error => {
        if (error) {
          reject(error);
        } else {
          resolve();
        }
      }
    );
  });
}

/*
 * Process one document.
 */
async function processDocument(
  event,
  api,
  attachment
) {
  const threadID =
    event.threadID;

  const workDir =
    path.join(
      TEMP_ROOT,
      `${Date.now()}-${crypto
        .randomBytes(6)
        .toString("hex")}`
    );

  const originalName =
    attachment.name ||
    attachment.filename ||
    "document";

  const safeName =
    originalName.replace(
      /[^\w.\- ]/g,
      "_"
    );

  const inputFile =
    path.join(
      workDir,
      safeName
    );

  try {
    await fs.ensureDir(workDir);

    if (!attachment.url) {
      throw new Error(
        "Attachment URL is missing."
      );
    }

    console.log(
      `[DOC2PNG] Downloading ${originalName}`
    );

    await downloadFile(
      api,
      attachment.url,
      inputFile
    );

    const pages =
      await convertDocument(
        inputFile,
        workDir
      );

    console.log(
      `[DOC2PNG] ${originalName} converted: ${pages.length} page(s)`
    );

    for (
      let i = 0;
      i < pages.length;
      i++
    ) {
      console.log(
        `[DOC2PNG] Sending page ${i + 1}/${pages.length}`
      );

      await sendImage(
        api,
        pages[i],
        threadID
      );

      await sleep(SEND_DELAY);
    }

    console.log(
      `[DOC2PNG] Finished ${originalName}`
    );

  } catch (error) {
    console.error(
      `[DOC2PNG] ERROR ${originalName}:`,
      error.message
    );

  } finally {
    await fs
      .remove(workDir)
      .catch(() => {});
  }
}

module.exports = {
  config: {
    name: "doc2png",

    version: "8.0.0",

    author: "James Baroy",

    countDown: 0,

    role: 0,

    description: {
      en:
        "Automatically converts DOCX and PPTX files into PNG pages."
    },

    category: "utility"
  },

  onStart: async function () {},

  onChat: async function ({
    event,
    api
  }) {
    if (
      !event ||
      !event.threadID
    ) {
      return;
    }

    if (
      !Array.isArray(
        event.attachments
      )
    ) {
      return;
    }

    const documents =
      event.attachments.filter(
        attachment => {
          const name =
            attachment.name ||
            attachment.filename ||
            "";

          return isSupported(name);
        }
      );

    if (!documents.length) {
      return;
    }

    for (
      const attachment of documents
    ) {
      enqueue(
        event.threadID,
        () =>
          processDocument(
            event,
            api,
            attachment
          )
      );
    }
  }
};
