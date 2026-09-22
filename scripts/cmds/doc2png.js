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
const SEND_DELAY = 500;

// One queue for each group/chat.
const queues = new Map();

function sleep(ms) {
	return new Promise(resolve => setTimeout(resolve, ms));
}

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
				resolve({ stdout, stderr });
			} else {
				reject(
					new Error(
						`${command} exited with code ${code}\n${stderr}`
					)
				);
			}
		});
	});
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
	const extension = path
		.extname(fileName || "")
		.toLowerCase();

	if (extension) {
		return extension;
	}

	try {
		return path
			.extname((url || "").split("?")[0])
			.toLowerCase();
	} catch {
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
		throw new Error("File is larger than 100 MB.");
	}

	await new Promise((resolve, reject) => {
		const writer = fs.createWriteStream(outputFile);

		response.data.pipe(writer);

		writer.on("finish", resolve);
		writer.on("error", reject);

		response.data.on("error", reject);
	});
}

async function convertDocument(inputFile, jobDir) {
	const pdfDir = path.join(jobDir, "pdf");

	await fs.ensureDir(pdfDir);

	/*
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

	if (!await fs.pathExists(pdfFile)) {
		throw new Error(
			"LibreOffice failed to create PDF."
		);
	}

	/*
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

	const files = await fs.readdir(jobDir);

	const pngFiles = files
		.filter(file =>
			/^page-\d+\.png$/i.test(file)
		)
		.map(file =>
			path.join(jobDir, file)
		);

	/*
	 * NUMERIC SORT
	 *
	 * 1, 2, 3 ... 10, 11 ...
	 *
	 * NOT:
	 *
	 * 1, 10, 11, 2 ...
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

	queues.set(threadID, next);

	next.finally(() => {
		if (queues.get(threadID) === next) {
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
	const url = attachment.url;

	if (!url) {
		return;
	}

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

	const jobDir = path.join(
		TEMP_ROOT,
		jobID
	);

	await fs.ensureDir(jobDir);

	const safeName =
		path.basename(fileName).replace(
			/[<>:"/\\|?*\x00-\x1F]/g,
			"_"
		);

	const inputFile = path.join(
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

		if (!pngFiles.length) {
			throw new Error(
				"No pages/slides were generated."
			);
		}

		/*
		 * SEND ONE IMAGE AT A TIME.
		 *
		 * We deliberately use ONE ReadStream
		 * per send because this is the safest
		 * attachment format for the FCA API.
		 *
		 * The queue guarantees that the pages
		 * remain in the correct order.
		 */
		for (
			let i = 0;
			i < pngFiles.length;
			i++
		) {
			const pngFile = pngFiles[i];

			await api.sendMessage(
				{
					attachment:
						fs.createReadStream(
							pngFile
						)
				},
				event.threadID
			);

			await sleep(SEND_DELAY);
		}
	}
	catch (error) {
		/*
		 * Don't send error/status messages
		 * to the group.
		 *
		 * Errors appear in Render logs.
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
		version: "3.1.0",
		author: "James Baroy",
		countDown: 0,
		role: 0,

		description: {
			en: "Automatically converts PPTX and DOCX files to PNG."
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

			if (!threadID) {
				return;
			}

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

				/*
				 * Only PPTX and DOCX.
				 */
				if (
					extension !== ".pptx" &&
					extension !== ".docx"
				) {
					continue;
				}

				/*
				 * QUEUE PER GROUP
				 *
				 * Document A:
				 * 1 2 3 ... 20
				 *
				 * then Document B:
				 * 1 2 3 ... 20
				 *
				 * Never mixed.
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
