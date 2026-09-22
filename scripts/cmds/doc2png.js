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

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: ["ignore", "pipe", "pipe"]
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", d => {
      stdout += d.toString();
    });

    child.stderr.on("data", d => {
      stderr += d.toString();
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

function enqueue(threadID, task) {
  const previous = queues.get(threadID) || Promise.resolve();

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

async function downloadFile(url, output) {
  const response = await axios({
    method: "GET",
    url,
    responseType: "stream",
    timeout: 180000,
    maxContentLength: MAX_FILE_SIZE,
    maxBodyLength: MAX_FILE_SIZE
  });

  await new Promise((resolve, reject) => {
    const writer = fs.createWriteStream(output);

    response.data.pipe(writer);

    writer.on("finish", resolve);
    writer.on("error", reject);
    response.data.on("error", reject);
  });
}

async function convertDocument(inputFile, workDir) {
  const pdfDir = path.join(workDir, "pdf");
  const pngDir = path.join(workDir, "png");
  const loProfile = path.join(workDir, "lo-profile");

  await fs.ensureDir(pdfDir);
  await fs.ensureDir(pngDir);
  await fs.ensureDir(loProfile);

  const profileURL =
    "file://" + loProfile.replace(/\\/g, "/");

  /*
   * Use a completely separate LibreOffice profile.
   * This prevents another LibreOffice process from interfering
   * with the conversion.
   */
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

  const baseName = path.basename(
    inputFile,
    path.extname(inputFile)
  );

  const pdfFile = path.join(
    pdfDir,
    `${baseName}.pdf`
  );

  if (!await fs.pathExists(pdfFile)) {
    throw new Error("LibreOffice did not create a PDF.");
  }

  /*
   * Ask Poppler how many pages were actually created.
   */
  const info = await run("pdfinfo", [pdfFile]);

  const match = info.stdout.match(/Pages:\s+(\d+)/i);

  const pageCount = match
    ? parseInt(match[1], 10)
    : 0;

  console.log(
    `[DOC2PNG] PDF pages detected: ${pageCount}`
  );

  if (pageCount < 1) {
    throw new Error("PDF contains no pages.");
  }

  /*
   * PDF → PNG
   */
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

  const files = await fs.readdir(pngDir);

  const pages = files
    .filter(file => /^page-\d+\.png$/i.test(file))
    .sort((a, b) => {
      const A = parseInt(a.match(/\d+/)[0], 10);
      const B = parseInt(b.match(/\d+/)[0], 10);
      return A - B;
    })
    .map(file => path.join(pngDir, file));

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

async function processDocument(event, api, attachment) {
  const threadID = event.threadID;

  const workDir = path.join(
    TEMP_ROOT,
    `${Date.now()}-${crypto.randomBytes(6).toString("hex")}`
  );

  const originalName =
    attachment.name ||
    attachment.filename ||
    "document";

  const safeName = originalName.replace(
    /[^\w.\- ]/g,
    "_"
  );

  const inputFile = path.join(
    workDir,
    safeName
  );

  try {
    await fs.ensureDir(workDir);

    if (!attachment.url) {
      throw new Error("Attachment URL is missing.");
    }

    console.log(
      `[DOC2PNG] Downloading ${originalName}`
    );

    await downloadFile(
      attachment.url,
      inputFile
    );

    const stat = await fs.stat(inputFile);

    console.log(
      `[DOC2PNG] File size: ${(stat.size / 1024 / 1024).toFixed(2)} MB`
    );

    if (stat.size > MAX_FILE_SIZE) {
      throw new Error("File exceeds 100 MB limit.");
    }

    const pages = await convertDocument(
      inputFile,
      workDir
    );

    console.log(
      `[DOC2PNG] ${originalName} converted: ${pages.length} page(s)`
    );

    /*
     * Send every page in exact order.
     */
    for (let i = 0; i < pages.length; i++) {
      console.log(
        `[DOC2PNG] Sending page ${i + 1}/${pages.length}`
      );

      await api.sendMessage(
        {
          attachment: fs.createReadStream(
            pages[i]
          )
        },
        threadID
      );

      await sleep(SEND_DELAY);
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
    await fs.remove(workDir).catch(() => {});
  }
}

module.exports = {
  config: {
    name: "doc2png",
    version: "5.0.0",
    author: "James Baroy",
    countDown: 0,
    role: 0,
    description: {
      en: "Automatically converts DOCX and PPTX files into PNG pages."
    },
    category: "utility"
  },

  onStart: async function () {
    // Required by ST-BOT command loader.
  },

  onChat: async function ({ event, api }) {
    if (!event || !event.threadID) {
      return;
    }

    if (!Array.isArray(event.attachments)) {
      return;
    }

    const documents = event.attachments.filter(
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

    for (const attachment of documents) {
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
