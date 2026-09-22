"use strict";

const fs = require("fs-extra");
const path = require("path");
const axios = require("axios");
const crypto = require("crypto");
const { spawn } = require("child_process");

const TEMP_ROOT = path.join(
	process.cwd(),
	"scripts",
	"events",
	"tmp",
	"doc2png"
);

const MAX_FILE_SIZE = 100 * 1024 * 1024;
const SEND_DELAY = 500;

// Prevent different documents from sending at the same time.
const queues = new Map();

function sleep(ms) {
	return new Promise(resolve => setTimeout(resolve, ms));
}

function runCommand(command, args) {
	return new Promise((resolve, reject) => {
		const child = spawn(command, args);

		let stderr = "";
		let stdout = "";

		child.stdout.on("data", data => {
			stdout += data.toString();
		});

		child.stderr.on("data", data => {
			stderr += data.toString();
		});

		child.on("error", reject);

		child.on("close", code => {
			if (code === 0)
				return resolve({ stdout, stderr });

			reject(
				new Error(
					`${command} exited with code ${code}\n${stderr}`
				)
			);
		});
	});
}

async function downloadFile(url, outputPath) {
	const response = await axios({
		method: "GET",
		url,
		responseType: "stream",
		timeout: 120000,
		maxContentLength: MAX_FILE_SIZE,
		maxBodyLength: MAX_FILE_SIZE
	});

	const contentLength = Number(
		response.headers["content-length"] || 0
	);

	if (contentLength > MAX_FILE_SIZE)
		throw new Error("File is larger than 100 MB.");

	await new Promise((resolve, reject) => {
		const writer = fs.createWriteStream(outputPath);

		response.data.pipe(writer);

		writer.on("finish", resolve);
		writer.on("error", reject);
		response.data.on("error", reject);
	});
}

async function convertToPNG(inputFile, jobDir) {
	const pdfDir = path.join(jobDir, "pdf");

	await fs.ensureDir(pdfDir);

	// PPTX/DOCX -> PDF
	await runCommand("libreoffice", [
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

	const pdfFile = path.join(
		pdfDir,
		`${baseName}.pdf`
	);

	if (!await fs.pathExists(pdfFile))
		throw new Error("PDF conversion failed.");

	// PDF -> PNG
	const outputPrefix = path.join(jobDir, "page");

	await runCommand("pdftoppm", [
		"-png",
		"-r",
		"120",
		pdfFile,
		outputPrefix
	]);

	const files = await fs.readdir(jobDir);

	const pngFiles = files
		.filter(file => /^page-\d+\.png$/i.test(file))
		.map(file => path.join(jobDir, file));

	// IMPORTANT:
	// Sort numerically, not alphabetically.
	// This guarantees:
	// page-1
	// page-2
	// page-3
	// ...
	// page-10
	// page-11
	// etc.

	pngFiles.sort((a, b) => {
		const numberA = parseInt(
			path.basename(a).match(/\d+/)[0],
			10
		);

		const numberB = parseInt(
			path.basename(b).match(/\d+/)[0],
			10
		);

		return numberA - numberB;
	});

	return pngFiles;
}

function getAttachmentName(attachment) {
	return (
		attachment.name ||
		attachment.filename ||
		attachment.fileName ||
		"document"
	);
}

function getExtension(name, url) {
	const nameExtension = path
		.extname(name || "")
		.toLowerCase();

	if (nameExtension)
		return nameExtension;

	try {
		return path
			.extname(url.split("?")[0])
			.toLowerCase();
	}
	catch {
		return "";
	}
}

function addToQueue(threadID, task) {
	const previous = queues.get(threadID) || Promise.resolve();

	const next = previous
		.catch(() => {})
		.then(task);

	queues.set(threadID, next);

	next.finally(() => {
		if (queues.get(threadID) === next)
			queues.delete(threadID);
	});

	return next;
}

async function processDocument({
	attachment,
	message,
	event
}) {
	const url = attachment.url;

	if (!url)
		return;

	const fileName = getAttachmentName(attachment);

	const extension = getExtension(
		fileName,
		url
	);

	if (![".pptx", ".docx"].includes(extension))
		return;

	const jobID =
		Date.now() +
		"-" +
		crypto.randomBytes(5).toString("hex");

	const jobDir = path.join(
		TEMP_ROOT,
		jobID
	);

	await fs.ensureDir(jobDir);

	const safeName = path
		.basename(fileName)
		.replace(
			/[<>:"/\\|?*\x00-\x1F]/g,
			"_"
		);

	const inputFile = path.join(
		jobDir,
		safeName.endsWith(extension)
			? safeName
			: `${safeName}${extension}`
	);

	try {
		// Download
		await downloadFile(
			url,
			inputFile
		);

		// Convert
		const pngFiles = await convertToPNG(
			inputFile,
			jobDir
		);

		if (!pngFiles.length)
			throw new Error(
				"No pages/slides were generated."
			);

		/*
		 * IMPORTANT:
		 * No status messages are sent.
		 *
		 * Only PNG images are sent.
		 */

		for (let i = 0; i < pngFiles.length; i++) {
			const pngFile = pngFiles[i];

			/*
			 * Empty body.
			 * The bot sends the picture only.
			 */
			await message.reply({
				attachment: fs.createReadStream(
					pngFile
				)
			});

			await sleep(SEND_DELAY);
		}
	}
	catch (error) {
		console.error(
			"[DOC2PNG]",
			error
		);

		/*
		 * We intentionally don't send
		 * "received/converting/complete"
		 * messages.
		 *
		 * Errors are logged on Render.
		 */
	}
	finally {
		await fs.remove(jobDir).catch(() => {});
	}
}

module.exports = {
	config: {
		name: "doc2png",
		version: "2.0.0",
		author: "James Baroy",
		description:
			"Automatically converts PPTX and DOCX files to PNG.",
		category: "events"
	},

	onStart: async function ({
		event,
		message
	}) {
		try {
			if (
				!event ||
				!event.attachments ||
				!event.attachments.length
			)
				return;

			const threadID =
				event.threadID ||
				event.senderID ||
				"unknown";

			/*
			 * Queue the whole document.
			 *
			 * If another PPTX/DOCX arrives while
			 * this one is being processed, it waits.
			 *
			 * This prevents:
			 *
			 * PPT A page 1
			 * PPT B page 1
			 * PPT A page 2
			 *
			 * from happening.
			 */

			for (
				const attachment
				of event.attachments
			) {
				const fileName =
					getAttachmentName(
						attachment
					);

				const extension =
					getExtension(
						fileName,
						attachment.url || ""
					);

				if (
					![".pptx", ".docx"]
						.includes(extension)
				)
					continue;

				addToQueue(
					threadID,
					() =>
						processDocument({
							attachment,
							message,
							event
						})
				);
			}
		}
		catch (error) {
			console.error(
				"[DOC2PNG EVENT]",
				error
			);
		}
	}
};
