const allOnEvent = global.GoatBot.onEvent;

module.exports = {
	config: {
		name: "onEvent",
		version: "1.1",
		author: "NTKhang",
		description: "Loop to all event in global.GoatBot.onEvent and run when have new event",
		category: "events"
	},

	onStart: async ({
		api,
		args,
		message,
		event,
		threadsData,
		usersData,
		dashBoardData,
		threadModel,
		userModel,
		dashBoardModel,
		role,
		commandName
	}) => {

		// ==========================================
		// ALWAYS READ / AUTO READ
		// ==========================================
		if (event && event.threadID && event.type === "message") {
			if (typeof api.markAsRead === "function") {
				api.markAsRead(event.threadID, (err) => {
					if (err) {
						console.error(
							"[ALWAYS READ] Failed:",
							err
						);
					} else {
						console.log(
							"[ALWAYS READ] Success:",
							event.threadID
						);
					}
				});
			} else {
				console.error(
					"[ALWAYS READ] api.markAsRead is not available."
				);
			}
		}

		// ==========================================
		// RUN OTHER EVENTS
		// ==========================================
		for (const item of allOnEvent) {
			if (typeof item === "string")
				continue;

			item.onStart({
				api,
				args,
				message,
				event,
				threadsData,
				usersData,
				threadModel,
				dashBoardData,
				userModel,
				dashBoardModel,
				role,
				commandName
			});
		}
	}
};
