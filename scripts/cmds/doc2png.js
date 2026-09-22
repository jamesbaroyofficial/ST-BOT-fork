"use strict";

const fs = require("fs-extra");
const path = require("path");
const axios = require("axios");
const crypto = require("crypto");
const { spawn } = require("child_process");

const TEMP_ROOT = path.join(process.cwd(), "tmp", "doc2png");

const MAX_FILE_SIZE = 100 * 1024 * 1024;
const SEND_DELAY = 700;

const queues = new Map();

const sleep = ms =>
  new Promise(resolve => setTimeout(resolve, ms));

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
        resolve({
          stdout,
          stderr
        });
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
 * Keep documents from the same GC in order.
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
 * Convert api.getAppState() into a Cookie header.
 */
function buildCookieHeader(api) {
  try {
    if (!api || typeof api.getAppState !== "function") {
      return "";
    }

    const appState = api.getAppState();

    if (!Array.isArray(appState)) {
      return "";
    }

    return appState
      .filter(cookie => {
        return (
          cookie &&
          cookie.key &&
          cookie.value !== undefined &&
          cookie.value !== null
        );
      })
      .map(cookie => {
        return `${cookie.key}=${cookie.value}`;
      })
      .join("; ");
  } catch (error) {
    console.error(
      "[DOC2PNG] Could not read appState:",
      error.message
    );

    return "";
  }
}

/*
 * Download Facebook attachment using the
 * currently logged-in Facebook session.
 */
async function downloadFile(api, url, output) {
  if (!url) {
    throw new Error("Attachment URL is missing.");
  }

  const cookie = buildCookieHeader(api);

  if (!cookie) {
    throw new Error(
      "Facebook session cookies are unavailable."
    );
  }

  console.log(
    "[DOC2PNG] Downloading attachment with authenticated session..."
  );

  const response = await axios({
    method: "GET",
    url,
    responseType: "stream",

    timeout: 180000,

    maxRedirects: 10,

    maxContentLength: MAX_FILE_SIZE,
    maxBodyLength: MAX_FILE_SIZE,

    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36",

      "Accept":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/pdf,*/*",

      "Accept-Language":
        "en-US,en;q=0.9",

      "Referer":
        "https://www.facebook.com/",

      "Cookie": cookie
    },

    validateStatus: status =>
      status >= 200 && status < 400
  });

  console.log(
    `[DOC2PNG] HTTP status: ${response.status}`
  );

  console.log(
    `[DOC2PNG] Content-Type: ${
      response.headers["content-type"] || "unknown"
    }`
  );

  console.log(
    `[DOC2PNG] Content-Length: ${
      response.headers["content-length"] || "unknown"
    }`
  );

  await new Promise((resolve, reject) => {
    const writer = fs.createWriteStream(output);

    let bytes = 0;

    response.data.on("data", chunk => {
      bytes += chunk.length;

      if (bytes > MAX_FILE_SIZE) {
        response.data.destroy(
          new Error("File exceeds 100 MB limit.")
        );
      }
    });

    response.data.pipe(writer);

    writer.on("finish", () => {
      resolve();
    });

    writer.on("error", reject);

    response.data.on("error", reject);
  });

  const stat = await fs.stat(output);

  console.log(
    `[DOC2PNG] Downloaded size: ${(stat.size / 1024 / 1024).toFixed(2)} MB`
  );

  /*
   * This is important.
   * Never allow LibreOffice to process an empty file.
   */
  if (stat.size <= 0) {
    throw new Error(
      "Facebook returned an empty/0-byte attachment."
    );
  }

  if (stat.size > MAX_FILE_SIZE) {
    throw new Error(
      "Downloaded file exceeds 100 MB limit."
    );
  }

  return stat.size;
}

/*
 * Convert DOCX/PPTX -> PDF -> PNG pages.
 */
async function convertDocument(
  inputFile,
  workDir
) {
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

    path.join(
      pngDir,
      "page"
    )
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
 * Send one PNG and wait for FCA callback.
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
    await fs.ensureDir(
      workDir
    );

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

    /*
     * Send pages in exact order.
     */
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

      await sleep(
        SEND_DELAY
      );
    }

    console.log(
      `[DOC2PNG] Finished ${originalName}`
    );
  }

  catch (error) {
    console.error(
      `[DOC2PNG] ERROR ${originalName}:`,
      error.message
    );
  }

  finally {
    /*
     * Delete temporary files.
     */
    await fs
      .remove(workDir)
      .catch(() => {});
  }
}

module.exports = {

  config: {
    name: "doc2png",

    version: "6.0.0",

    author: "James Baroy",

    countDown: 0,

    role: 0,

    description: {
      en:
        "Automatically converts DOCX and PPTX files into PNG pages."
    },

    category: "utility"
  },

  /*
   * Required by ST-BOT command loader.
   */
  onStart: async function () {},

  /*
   * Detect incoming attachments.
   */
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

    if (
      !documents.length
    ) {
      return;
    }

    /*
     * Queue documents per GC.
     */
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
