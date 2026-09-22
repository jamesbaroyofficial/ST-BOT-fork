"use strict";

const fs = require("fs-extra");
const path = require("path");
const axios = require("axios");
const crypto = require("crypto");
const { spawn } = require("child_process");

const TEMP_ROOT = path.join(
	process.cwd(),
	"tmp",
	"doc2png"
);

const MAX_FILE_SIZE = 100 * 1024 * 1024;

// One queue per group.
// This prevents documents from different users
// from getting mixed together.
const queues = new Map();

function runCommand(command, args) {
	return new Promise((resolve, reject) => {
		const child = spawn(command, args);

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
			}
			else {
				reject(
					new Error(
						`${command} exited with code ${code}\n${stderr}`
					)
				);
			}
		});
	});
}

function sleep(ms) {
	return new Promise(resolve => setTimeout(resolve, ms));
}

function getFileName(attachment) {
	return (
		attachment.name ||
		attachment.filename ||
		attachment.fileName ||
		attachment.title ||
		"document"
	);
}

function getExtension(fileName, url) {
	const nameExt = path
		.extname(fileName || "")
		.toLowerCase();

	if (nameExt)
		return nameExt;

	try {
		return path
			.extname((url || "").split("?")[0])
			.toLowerCase();
	}
	catch {
		return "";
	}
}

async function downloadFile(url, outputFile) {
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

	if (
		contentLength &&
		contentLength > MAX_FILE_SIZE
	) {
		throw new Error(
			"File is larger than 100 MB."
		);
	}

	await new Promise((resolve, reject) => {
		const writer =
			fs.createWriteStream(outputFile);

		response.data.pipe(writer);

		writer.on("finish", resolve);
		writer.on("error", reject);

		response.data.on(
			"error",
			reject
		);
	});
}

async function convertDocument(
	inputFile,
	jobDir
) {
	const pdfDir = path.join(
		jobDir,
		"pdf"
	);

	await fs.ensureDir(pdfDir);

	/*
	 * STEP 1
	 * PPTX/DOCX -> PDF
	 */
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

	if (
		!await fs.pathExists(pdfFile)
	) {
		throw new Error(
			"LibreOffice failed to create PDF."
		);
	}

	/*
	 * STEP 2
	 * PDF -> PNG
	 */
	const outputPrefix = path.join(
		jobDir,
		"page"
	);

	await runCommand("pdftoppm", [
		"-png",
		"-r",
		"120",
		pdfFile,
		outputPrefix
	]);

	const files =
		await fs.readdir(jobDir);

	const pngFiles = files
		.filter(file =>
			/^page-\d+\.png$/i.test(file)
		)
		.map(file =>
			path.join(jobDir, file)
		);

	/*
	 * IMPORTANT:
	 * Numeric sorting.
	 *
	 * This prevents:
	 * 1, 10, 11, 12, 2...
	 *
	 * and guarantees:
	 * 1, 2, 3, 4...
	 */
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

function addToQueue(threadID, task) {
	const previous =
		queues.get(threadID) ||
		Promise.resolve();

	const next = previous
		.catch(() => {})
		.then(task);

	queues.set(
		threadID,
		next
	);

	next.finally(() => {
		if (
			queues.get(threadID) === next
		) {
			queues.delete(threadID);
		}
	});

	return next;
}

async function processDocument({
	attachment,
	event,
	api
}) {
	const url =
		attachment.url;

	if (!url)
		return;

	const fileName =
		getFileName(attachment);

	const extension =
		getExtension(
			fileName,
			url
		);

	if (
		extension !== ".pptx" &&
		extension !== ".docx"
	) {
		return;
	}

	const jobID =
		Date.now() +
		"-" +
		crypto
			.randomBytes(6)
			.toString("hex");

	const jobDir =
		path.join(
			TEMP_ROOT,
			jobID
		);

	await fs.ensureDir(
		jobDir
	);

	const safeName =
		path.basename(fileName)
			.replace(
				/[<>:"/\\|?*\x00-\x1F]/g,
				"_"
			);

	const inputFile =
		path.join(
			jobDir,
			safeName.endsWith(extension)
				? safeName
				: safeName + extension
		);

	try {
		/*
		 * DOWNLOAD
		 */
		await downloadFile(
			url,
			inputFile
		);

		/*
		 * CONVERT
		 */
		const pngFiles =
			await convertDocument(
				inputFile,
				jobDir
			);

		if (
			!pngFiles.length
		) {
			throw new Error(
				"No pages were generated."
			);
		}

		/*
		 * CREATE ALL ATTACHMENTS FIRST
		 *
		 * This is important.
		 *
		 * Instead of:
		 *
		 * send page 1
		 * send page 2
		 * send page 3
		 *
		 * we create ONE attachment array:
		 *
		 * [page1, page2, page3...]
		 *
		 * and send it as ONE message.
		 */
		const attachments =
			pngFiles.map(
				file =>
					fs.createReadStream(file)
			);

		/*
		 * SEND ONE BATCH
		 *
		 * No text/body.
		 * Pictures only.
		 */
		await api.sendMessage(
			{
				attachment: attachments
			},
			event.threadID
		);

		/*
		 * Small delay before cleanup
		 * to give the API enough time
		 * to consume the streams.
		 */
		await sleep(1000);
	}
	catch (error) {
		/*
		 * Do NOT send error/status messages
		 * to the group.
		 *
		 * Error is logged on Render.
		 */
		console.error(
			"[DOC2PNG ERROR]",
			error
		);
	}
	finally {
		/*
		 * Delete temporary files.
		 */
		await fs.remove(
			jobDir
		).catch(() => {});
	}
}

module.exports = {
	config: {
		name: "doc2png",
		version: "3.0.0",
		author: "James Baroy",
		countDown: 0,
		role: 0,

		description: {
			en: "Automatically converts PPTX and DOCX files to PNG images."
		},

		category: "events"
	},

	onEvent: async function ({
		event,
		api
	}) {
		try {
			if (
				!event ||
				!event.attachments ||
				!event.attachments.length
			) {
				return;
			}

			const threadID =
				event.threadID ||
				event.senderID;

			if (!threadID)
				return;

			/*
			 * Check every attachment.
			 */
			for (
				const attachment
				of event.attachments
			) {
				const fileName =
					getFileName(
						attachment
					);

				const extension =
					getExtension(
						fileName,
						attachment.url || ""
					);

				if (
					extension !== ".pptx" &&
					extension !== ".docx"
				) {
					continue;
				}

				/*
				 * QUEUE BY GROUP
				 *
				 * If multiple documents
				 * are sent quickly:
				 *
				 * Document A
				 *      ↓
				 * all A pages
				 *      ↓
				 * Document B
				 *      ↓
				 * all B pages
				 *
				 * Never:
				 *
				 * A1
				 * B1
				 * A2
				 * B2
				 */
				addToQueue(
					threadID,
					() =>
						processDocument({
							attachment,
							event,
							api
						})
				);
			}
		}
		catch (error) {
			console.error(
				"[DOC2PNG EVENT ERROR]",
				error
			);
		}
	}
};
