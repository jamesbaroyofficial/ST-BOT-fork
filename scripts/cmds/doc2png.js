"use strict";

const fs = require("fs-extra");
const path = require("path");
const crypto = require("crypto");
const request = require("request");
const cheerio = require("cheerio");
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

  const next =
    previous
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
 * Create Facebook cookie jar from api.getAppState().
 *
 * IMPORTANT:
 * We do NOT put Facebook cookies on the
 * Facebook CDN/file host.
 */
function createFacebookJar(api) {
  if (
    !api ||
    typeof api.getAppState !== "function"
  ) {
    throw new Error(
      "ST-FCA getAppState() is unavailable."
    );
  }

  const appState =
    api.getAppState();

  if (
    !Array.isArray(appState) ||
    appState.length === 0
  ) {
    throw new Error(
      "Facebook session cookies are unavailable."
    );
  }

  const jar =
    request.jar();

  for (const item of appState) {
    if (
      !item ||
      !item.key ||
      item.value === undefined ||
      item.value === null
    ) {
      continue;
    }

    const cookieString =
      `${item.key}=${item.value}`;

    try {
      const cookie =
        request.cookie(
          cookieString
        );

      /*
       * If the saved cookie contains
       * a domain, preserve it.
       */
      const domain =
        item.domain ||
        "";

      if (
        domain.includes("messenger.com")
      ) {
        jar.setCookie(
          cookie,
          "https://www.messenger.com/"
        );
      } else {
        jar.setCookie(
          cookie,
          "https://www.facebook.com/"
        );
      }
    } catch (_) {
      /*
       * Never print cookie values.
       */
    }
  }

  return jar;
}

/*
 * Request helper.
 */
function requestBuffer(
  url,
  jar,
  userAgent
) {
  return new Promise(
    (resolve, reject) => {
      const req =
        request.get({
          url,

          jar,

          gzip: true,

          followRedirect: true,

          followAllRedirects: true,

          timeout: 180000,

          encoding: null,

          headers: {
            "User-Agent":
              userAgent,

            "Accept":
              "application/octet-stream,text/html,application/xhtml+xml,*/*",

            "Accept-Language":
              "en-US,en;q=0.9",

            "Referer":
              "https://www.facebook.com/",

            "Connection":
              "keep-alive"
          }
        });

      req.on(
        "response",
        response => {
          const chunks = [];

          response.on(
            "data",
            chunk => {
              chunks.push(chunk);

              const total =
                chunks.reduce(
                  (sum, item) =>
                    sum + item.length,
                  0
                );

              if (
                total >
                MAX_FILE_SIZE
              ) {
                req.destroy(
                  new Error(
                    "Response exceeds 100 MB."
                  )
                );
              }
            }
          );

          response.on(
            "end",
            () => {
              resolve({
                statusCode:
                  response.statusCode,

                headers:
                  response.headers,

                finalUrl:
                  response.request &&
                  response.request.uri
                    ? response.request.uri.href
                    : url,

                body:
                  Buffer.concat(
                    chunks
                  )
              });
            }
          );
        }
      );

      req.on(
        "error",
        reject
      );
    }
  );
}

/*
 * Find a real file/download URL inside
 * Facebook's attachment preview HTML.
 */
function findDownloadUrl(
  html,
  originalUrl,
  filename
) {
  const $ =
    cheerio.load(
      html
    );

  const candidates = [];

  /*
   * All anchor links.
   */
  $("a[href]").each(
    (_, element) => {
      const href =
        $(element).attr(
          "href"
        );

      if (href) {
        candidates.push(
          href
        );
      }
    }
  );

  /*
   * Meta URLs.
   */
  $(
    'meta[property="og:url"], meta[property="og:video"], meta[property="og:image"], meta[name="twitter:image"]'
  ).each(
    (_, element) => {
      const content =
        $(element).attr(
          "content"
        );

      if (content) {
        candidates.push(
          content
        );
      }
    }
  );

  /*
   * Raw HTML URLs.
   */
  const rawMatches =
    html.match(
      /https?:\\?\/\\?\/[^"'<> ]+/g
    ) || [];

  for (
    const value of rawMatches
  ) {
    candidates.push(
      value
        .replace(/\\\//g, "/")
        .replace(/\\"/g, '"')
    );
  }

  const decodedName =
    decodeURIComponent(
      filename || ""
    ).toLowerCase();

  /*
   * Normalize and remove duplicates.
   */
  const unique =
    [...new Set(
      candidates
        .map(value => {
          try {
            return new URL(
              value,
              originalUrl
            ).href;
          } catch (_) {
            return null;
          }
        })
        .filter(Boolean)
    )];

  /*
   * Prefer links containing the filename.
   */
  const filenameMatch =
    unique.find(url =>
      url
        .toLowerCase()
        .includes(
          decodedName
        )
    );

  if (
    filenameMatch
  ) {
    return filenameMatch;
  }

  /*
   * Prefer obvious download/file URLs.
   */
  const downloadMatch =
    unique.find(url => {
      const lower =
        url.toLowerCase();

      return (
        lower.includes(
          "download"
        ) ||
        lower.includes(
          "attachment"
        ) ||
        lower.includes(
          ".pptx"
        ) ||
        lower.includes(
          ".docx"
        )
      );
    });

  if (
    downloadMatch
  ) {
    return downloadMatch;
  }

  return null;
}

/*
 * Build Facebook's attachment-preview URL.
 */
function buildPreviewUrl(
  event,
  attachment
) {
  if (
    !event ||
    !event.messageID ||
    !event.threadID ||
    !attachment ||
    !attachment.ID
  ) {
    return null;
  }

  const params =
    new URLSearchParams();

  params.set(
    "mid",
    String(
      event.messageID
    )
  );

  params.set(
    "threadid",
    String(
      event.threadID
    )
  );

  params.set(
    "fbid",
    String(
      attachment.ID
    )
  );

  return (
    "https://m.facebook.com/messages/attachment_preview/?" +
    params.toString()
  );
}

/*
 * Try downloading the direct attachment URL.
 */
async function tryDirectDownload(
  api,
  attachment,
  output
) {
  const url =
    attachment.url;

  if (!url) {
    return {
      success: false,
      html: false
    };
  }

  const jar =
    createFacebookJar(
      api
    );

  const userAgent =
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) " +
    "AppleWebKit/537.36 (KHTML, like Gecko) " +
    "Chrome/131.0.0.0 Safari/537.36";

  console.log(
    "[DOC2PNG] Trying direct attachment URL..."
  );

  const response =
    await requestBuffer(
      url,
      jar,
      userAgent
    );

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

  const contentType =
    String(
      response.headers[
        "content-type"
      ] || ""
    ).toLowerCase();

  /*
   * HTML means Facebook gave us a page,
   * not the document.
   */
  if (
    contentType.includes(
      "text/html"
    )
  ) {
    console.log(
      "[DOC2PNG] Direct URL returned HTML."
    );

    return {
      success: false,
      html: true
    };
  }

  if (
    response.statusCode < 200 ||
    response.statusCode >= 400
  ) {
    return {
      success: false,
      html: false
    };
  }

  if (
    !response.body ||
    response.body.length <= 0
  ) {
    return {
      success: false,
      html: false
    };
  }

  await fs.writeFile(
    output,
    response.body
  );

  const stat =
    await fs.stat(
      output
    );

  if (
    stat.size <= 0
  ) {
    return {
      success: false,
      html: false
    };
  }

  console.log(
    `[DOC2PNG] Downloaded size: ${(stat.size / 1024 / 1024).toFixed(2)} MB`
  );

  return {
    success: true,
    html: false
  };
}

/*
 * Fallback:
 * Facebook attachment preview page ->
 * actual file URL.
 */
async function downloadFromPreview(
  api,
  event,
  attachment,
  output
) {
  const previewUrl =
    buildPreviewUrl(
      event,
      attachment
    );

  if (!previewUrl) {
    throw new Error(
      "Cannot build Facebook attachment preview URL."
    );
  }

  console.log(
    "[DOC2PNG] Trying Facebook attachment preview..."
  );

  console.log(
    `[DOC2PNG] Preview URL created for attachment ID ${attachment.ID}`
  );

  const jar =
    createFacebookJar(
      api
    );

  const userAgent =
    "Mozilla/5.0 (Linux; Android 14) " +
    "AppleWebKit/537.36 (KHTML, like Gecko) " +
    "Chrome/131.0.0.0 Mobile Safari/537.36";

  const preview =
    await requestBuffer(
      previewUrl,
      jar,
      userAgent
    );

  console.log(
    `[DOC2PNG] Preview HTTP status: ${preview.statusCode}`
  );

  const previewType =
    String(
      preview.headers[
        "content-type"
      ] || ""
    ).toLowerCase();

  console.log(
    `[DOC2PNG] Preview Content-Type: ${previewType || "unknown"}`
  );

  if (
    preview.statusCode < 200 ||
    preview.statusCode >= 400
  ) {
    throw new Error(
      `Facebook attachment preview returned HTTP ${preview.statusCode}.`
    );
  }

  const html =
    preview.body.toString(
      "utf8"
    );

  /*
   * If Facebook redirects the preview directly
   * to a binary file, save it.
   */
  if (
    !previewType.includes(
      "text/html"
    )
  ) {
    if (
      preview.body.length <= 0
    ) {
      throw new Error(
        "Facebook preview returned an empty file."
      );
    }

    await fs.writeFile(
      output,
      preview.body
    );

    return;
  }

  const fileUrl =
    findDownloadUrl(
      html,
      preview.finalUrl ||
        previewUrl,
      attachment.name ||
        attachment.filename ||
        ""
    );

  if (!fileUrl) {
    throw new Error(
      "Facebook preview page did not contain a downloadable document URL."
    );
  }

  console.log(
    "[DOC2PNG] Found document download URL."
  );

  const fileResponse =
    await requestBuffer(
      fileUrl,
      jar,
      userAgent
    );

  console.log(
    `[DOC2PNG] File HTTP status: ${fileResponse.statusCode}`
  );

  console.log(
    `[DOC2PNG] File Content-Type: ${
      fileResponse.headers[
        "content-type"
      ] || "unknown"
    }`
  );

  const fileContentType =
    String(
      fileResponse.headers[
        "content-type"
      ] || ""
    ).toLowerCase();

  if (
    fileContentType.includes(
      "text/html"
    )
  ) {
    throw new Error(
      "Facebook download URL returned HTML instead of the document."
    );
  }

  if (
    fileResponse.statusCode < 200 ||
    fileResponse.statusCode >= 400
  ) {
    throw new Error(
      `Facebook file download returned HTTP ${fileResponse.statusCode}.`
    );
  }

  if (
    !fileResponse.body ||
    fileResponse.body.length <= 0
  ) {
    throw new Error(
      "Facebook returned a 0-byte document."
    );
  }

  await fs.writeFile(
    output,
    fileResponse.body
  );

  const stat =
    await fs.stat(
      output
    );

  console.log(
    `[DOC2PNG] Downloaded size: ${(stat.size / 1024 / 1024).toFixed(2)} MB`
  );

  if (
    stat.size <= 0
  ) {
    throw new Error(
      "Downloaded document is 0 bytes."
    );
  }
}

/*
 * Main downloader.
 */
async function downloadFile(
  api,
  event,
  attachment,
  output
) {
  /*
   * First try the URL supplied by ST-FCA.
   */
  try {
    const direct =
      await tryDirectDownload(
        api,
        attachment,
        output
      );

    if (
      direct.success
    ) {
      return;
    }

    /*
     * If direct URL returned HTML,
     * use attachment preview.
     */
    if (
      direct.html
    ) {
      await downloadFromPreview(
        api,
        event,
        attachment,
        output
      );

      return;
    }
  } catch (error) {
    console.log(
      `[DOC2PNG] Direct download failed: ${error.message}`
    );
  }

  /*
   * Final fallback.
   */
  await downloadFromPreview(
    api,
    event,
    attachment,
    output
  );
}

/*
 * DOCX/PPTX -> PDF -> PNG
 */
async function convertDocument(
  inputFile,
  workDir
) {
  const pdfDir =
    path.join(
      workDir,
      "pdf"
    );

  const pngDir =
    path.join(
      workDir,
      "png"
    );

  const loProfile =
    path.join(
      workDir,
      "lo-profile"
    );

  await fs.ensureDir(
    pdfDir
  );

  await fs.ensureDir(
    pngDir
  );

  await fs.ensureDir(
    loProfile
  );

  const profileURL =
    "file://" +
    loProfile.replace(
      /\\/g,
      "/"
    );

  await run(
    "libreoffice",
    [
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
    ]
  );

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

  if (
    !await fs.pathExists(
      pdfFile
    )
  ) {
    throw new Error(
      "LibreOffice did not create a PDF."
    );
  }

  const pdfStat =
    await fs.stat(
      pdfFile
    );

  if (
    pdfStat.size <= 0
  ) {
    throw new Error(
      "LibreOffice created an empty PDF."
    );
  }

  const info =
    await run(
      "pdfinfo",
      [pdfFile]
    );

  const match =
    info.stdout.match(
      /Pages:\s+(\d+)/i
    );

  const pageCount =
    match
      ? parseInt(
          match[1],
          10
        )
      : 0;

  console.log(
    `[DOC2PNG] PDF pages detected: ${pageCount}`
  );

  if (
    pageCount < 1
  ) {
    throw new Error(
      "PDF contains no pages."
    );
  }

  await run(
    "pdftoppm",
    [
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
    ]
  );

  const files =
    await fs.readdir(
      pngDir
    );

  const pages =
    files
      .filter(file =>
        /^page-\d+\.png$/i.test(
          file
        )
      )
      .sort(
        (a, b) => {
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
        }
      )
      .map(file =>
        path.join(
          pngDir,
          file
        )
      );

  console.log(
    `[DOC2PNG] PNG pages generated: ${pages.length}`
  );

  if (
    pages.length !==
    pageCount
  ) {
    throw new Error(
      `Expected ${pageCount} PNG pages but generated ${pages.length}.`
    );
  }

  return pages;
}

/*
 * Send one PNG.
 */
function sendImage(
  api,
  file,
  threadID
) {
  return new Promise(
    (resolve, reject) => {
      const stream =
        fs.createReadStream(
          file
        );

      stream.on(
        "error",
        reject
      );

      api.sendMessage(
        {
          attachment:
            stream
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
    }
  );
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

    if (
      !attachment.url &&
      !attachment.ID
    ) {
      throw new Error(
        "Facebook attachment has no URL or ID."
      );
    }

    console.log(
      `[DOC2PNG] Downloading ${originalName}`
    );

    await downloadFile(
      api,
      event,
      attachment,
      inputFile
    );

    const stat =
      await fs.stat(
        inputFile
      );

    console.log(
      `[DOC2PNG] Final document size: ${(stat.size / 1024 / 1024).toFixed(2)} MB`
    );

    if (
      stat.size <= 0
    ) {
      throw new Error(
        "Final document is 0 bytes."
      );
    }

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

    version: "9.0.0",

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

          return isSupported(
            name
          );
        }
      );

    if (
      !documents.length
    ) {
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
