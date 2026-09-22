"use strict";

const fs = require("fs-extra");
const path = require("path");
const axios = require("axios");
const crypto = require("crypto");
const { spawn } = require("child_process");

const TEMP_ROOT = path.join(process.cwd(), "tmp", "doc2png");
const MAX_FILE_SIZE = 100 * 1024 * 1024;
const SEND_DELAY = 500;

const queues = new Map();

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: ["ignore", "pipe", "pipe"]
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", d => stdout += d.toString());
    child.stderr.on("data", d => stderr += d.toString());

    child.on("error", reject);

    child.on("close", code => {
      if (code === 0) {
        resolve({ stdout, stderr });
      } else {
        reject(new Error(`${command} failed: ${stderr || stdout}`));
      }
    });
  });
}

function isDocument(name) {
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
    timeout: 120000,
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

async function convertToPNG(inputFile, workDir) {
  const pdfDir = path.join(workDir, "pdf");
  const pngDir = path.join(workDir, "png");

  await fs.ensureDir(pdfDir);
  await fs.ensureDir(pngDir);

  // DOCX/PPTX → PDF
  await run("libreoffice", [
    "--headless",
    "--convert-to",
    "pdf",
    "--outdir",
    pdfDir,
    inputFile
  ]);

  const baseName = path.basename(
    inputFile,
    path.extname(inputFile)
  );

  const pdfFile = path.join(pdfDir, `${baseName}.pdf`);

  if (!await fs.pathExists(pdfFile)) {
    throw new Error("PDF conversion failed.");
  }

  // PDF → PNG
  const prefix = path.join(pngDir, "page");

  await run("pdftoppm", [
    "-png",
    "-r",
    "120",
    pdfFile,
    prefix
  ]);

  const files = await fs.readdir(pngDir);

  return files
    .filter(file => /^page-\d+\.png$/i.test(file))
    .sort((a, b) => {
      const A = parseInt(a.match(/\d+/)[0]);
      const B = parseInt(b.match(/\d+/)[0]);
      return A - B;
    })
    .map(file => path.join(pngDir, file));
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

  const safeName = originalName.replace(/[^\w.\- ]/g, "_");
  const inputFile = path.join(workDir, safeName);

  try {
    await fs.ensureDir(workDir);

    if (!attachment.url) {
      throw new Error("Attachment URL not found.");
    }

    await downloadFile(attachment.url, inputFile);

    const stat = await fs.stat(inputFile);

    if (stat.size > MAX_FILE_SIZE) {
      throw new Error("File is larger than 100 MB.");
    }

    const pages = await convertToPNG(inputFile, workDir);

    if (!pages.length) {
      throw new Error("No pages were converted.");
    }

    // Send page 1, page 2, page 3... in exact order
    for (const page of pages) {
      await api.sendMessage(
        {
          attachment: fs.createReadStream(page)
        },
        threadID
      );

      await sleep(SEND_DELAY);
    }

    console.log(
      `[DOC2PNG] ${originalName} converted: ${pages.length} page(s)`
    );
  } catch (error) {
    console.error(
      `[DOC2PNG] ${originalName}:`,
      error.message
    );
  } finally {
    await fs.remove(workDir).catch(() => {});
  }
}

module.exports = {
  config: {
    name: "doc2png",
    version: "4.0.0",
    author: "James Baroy",
    countDown: 0,
    role: 0,
    description: {
      en: "Automatically converts DOCX and PPTX files to PNG."
    },
    category: "utility"
  },

  onChat: async function ({ event, api }) {
    if (!event || !event.threadID) return;

    if (!Array.isArray(event.attachments)) return;

    const documents = event.attachments.filter(attachment => {
      const name =
        attachment.name ||
        attachment.filename ||
        "";

      return isDocument(name);
    });

    if (!documents.length) return;

    for (const attachment of documents) {
      enqueue(event.threadID, () =>
        processDocument(event, api, attachment)
      );
    }
  }
};
