const CONFIG = {
	SHEET_ID: 'YOUR_SHEET_ID_HERE',
	SHEET_NAME: 'Accounting',
	PUSH_URL: 'https://api.line.me/v2/bot/message/push',
	CHANNEL_ACCESS_TOKEN: 'YOUR_CHANNEL_ACCESS_TOKEN_HERE',
	CHANNEL_SECRET: 'YOUR_CHANNEL_SECRET_HERE',
	ALLOWED_USERS: new Set([
		'USER_ID_1',
	]),
	TIMEZONE: 'Your_Timezone_Here',
	MONTHLY_MESSAGE_LIMIT: 100,
	TRIGGER_DELAY: 5000
};

/**
 * Post repust handler
 * @param {Object} e - event
 * @return {TextOutput} HTTP response
 */
function doPost(e) {
	try {
		const response = ContentService.createTextOutput('OK');

		// Save the event data to PropertiesService for later processing
		const eventData = JSON.stringify(e);
		// writeLog(LEVEL.DEBUG, 'main.gs:doPost', 'Get Json Format Event', eventData);
		PropertiesService.getScriptProperties().setProperty('EVENT', eventData);

		ScriptApp.newTrigger('processCachedEvents').timeBased().after(CONFIG.TRIGGER_DELAY).create();

		return response;
	} catch (error) {
		writeLog(LEVEL.EMERGENCY, 'main.gs:doPost', 'Respond', error);
		return ContentService.createTextOutput('Error');
	}
}

/**
 * Handle cached events (Triggered by doPost)
 */
function processCachedEvents() {
	try {
		const scriptProperties = PropertiesService.getScriptProperties();
		const eventData = scriptProperties.getProperty('EVENT');
		if (!eventData) {
			writeLog(LEVEL.ERROR, 'main.gs:processCachedEvents', 'Get Data From Properties', 'No "EVENT" found in script properties');
			return;
		}

		const event = JSON.parse(eventData);
		const body = JSON.parse(event.postData.contents);
		// writeLog(LEVEL.DEBUG, 'main.gs:processCachedEvents', 'Get body Data', body);
		// if (!verifySignature(event)) {
		// 	writeLog(LEVEL.ERROR, 'main.gs:processCachedEvents', 'Verify Signature', 'Invalid signature');
		// 	return;
		// }

		handleEventsAsync(body);

	} catch (error) {
		writeLog(LEVEL.ERROR, 'main.gs:processCachedEvents', 'Process Events', error);
	} finally {
		scriptProperties.deleteProperty('EVENT');
		cleanupTriggers('processCachedEvents');
	}
}

/**
 * @param {Object} body - request body
 */
function handleEventsAsync(body) {
	try {
		const events = body.events;
		if (!Array.isArray(events)) {
			writeLog(LEVEL.ERROR, 'main.gs:handleEventsAsync', 'Validate Events', {
				message: 'Events is not an array',
				events: events || 'undefined'
			});
			return;
		}

		const validEvents = events.filter(event => event?.source?.userId);
		const illegalAccessSheet = getSheet('Illegal');

		for (const event of validEvents) {
			const userId = event.source.userId;

			if (!CONFIG.ALLOWED_USERS.has(userId)) {
				if (!illegalAccessSheet) { return; }
				const timestamp = formatTimestamp(new Date());
				const userAction = getUserActionDescription(event);
				illegalAccessSheet.appendRow([timestamp, userId, userAction]);
				writeLog(LEVEL.INFO, 'main.gs:handleEventsAsync', 'Log Illegal Access', { userId: userId, requestType: userSend });
				continue;
			}

			if (event.type === 'message' && event.message?.type === 'text') {
				handleTextMessage(event, userId);
			}
		}

	} catch (error) {
		writeLog(LEVEL.EMERGENCY, 'main.gs:handleEventsAsync', 'Handle Events', error);

	}
}

/**
 * Get user action description based on event type
 * @param {Object} event
 * @return {string} action description
 */
function getUserActionDescription(event) {
	if (event.type === 'message') {
		const messageType = event.message?.type;
		if (messageType === 'text') {
			return `send: ${event.message.text || 'no text'}`;
		}
		return `send: ${messageType || 'unknown message type'}`;
	}
	return event.type || 'unknown event type';
}

/**
 * @param {Object} event
 * @param {string} userId
 */
function handleTextMessage(event, userId) {
	const userMsg = event.message.text.trim();
	const sheet = getSheet(CONFIG.SHEET_NAME);
	if (!sheet) {
		replyText(userId, 'X, cannot access accounting sheet');
		return;
	}

	try {
		const now = new Date();
		const recordResult = handleRecordCommand(userMsg, sheet, now);
		if (recordResult) return recordResult;
		const queryResult = handleQueryCommands(userMsg, sheet, now);
		const reply = queryResult || getHelpText();
		replyText(userId, reply);
	} catch (error) {
		writeLog(LEVEL.EMERGENCY, 'main.gs:handleTextMessage', 'Process User Message', error);
		replyText(userId, 'X, cannot process your request at the moment. Please try again later.');
	}
}

/**
 * @param {string} userMsg - User input message
 * @param {Sheet} sheet - Spreadsheet sheet for accounting
 * @param {Date} now - Current date and time
 * @return {string|null} Response message or null if no query command matched
 */
function handleRecordCommand(userMsg, sheet, now) {
	const parsed = parseRecord(userMsg);
	if (!parsed) return null;

	const timestamp = formatTimestamp(now);
	sheet.appendRow([timestamp, parsed.amount, parsed.desc, parsed.category]);

	const categoryText = parsed.category ? ` Category：${parsed.category}` : '';
	return `V, accounting successful!\n💰 Amount: ${parsed.amount}\n📝 Description: ${parsed.desc}${categoryText}\n⏰ Time: ${timestamp}`;
}

/**
 * @param {string} userMsg - User input message
 * @param {Sheet} sheet - Spreadsheet sheet for accounting
 * @param {Date} now - Current date and time
 * @return {string|null} Response message or null if no query command matched
 */
function handleQueryCommands(userMsg, sheet, now) {
	const queryHandlers = {
		':0': () => getDailyRecords(sheet, 0),
		':-1': () => getDailyRecords(sheet, -1),
		':-7': () => getPastNDaysRecords(sheet, 7),
	};
	if (queryHandlers[userMsg]) {
		return queryHandlers[userMsg]();
	}

	const monthMatch = userMsg.match(/^:(\d{1,2})$/);
	if (monthMatch) {
		const month = parseInt(monthMatch[1], 10);
		if (month >= 1 && month <= 12) {
			return getMonthlyStats(sheet, now.getFullYear(), month);
		}
	}

	const yearMatch = userMsg.match(/^:(\d{4})$/);
	if (yearMatch) {
		const year = parseInt(yearMatch[1], 10);
		return getYearlyStats(sheet, year);
	}
	return null;
}

/**
 * Parse user input message to extract record details
 * @param {string} text - User input message
 * @return {Object|null} {amount, desc, category} or null if parsing fails
 */
function parseRecord(text) {
	const cleaned = text.replace(/\s+/g, ' ').trim();

	const patterns = [
		/^(\d+(?:\.\d+)?)\s+(.+)\s+@(.+)$/,
		/^(\d+(?:\.\d+)?)\s+(.+)$/
	];

	for (const pattern of patterns) {
		const match = cleaned.match(pattern);
		if (match) {
			return {
				amount: parseFloat(match[1]),
				desc: match[2].trim(),
				category: match[3]?.trim() || ''
			};
		}
	}
	return null;
}

/**
 * Get daily records from the sheet
 * @param {Sheet} sheet - Spreadsheet sheet for accounting
 * @param {number} offsetDays
 * @return {string} Daily summary of expenses
 */
function getDailyRecords(sheet, offsetDays) {
	const targetDate = new Date(Date.now() + offsetDays * 86400000);
	const targetDateStr = formatDate(targetDate, 'yyyy/MM/dd');

	const records = getFilteredRecords(sheet, row => {
		const recordDate = formatDate(new Date(row[0]), 'yyyy/MM/dd');
		return recordDate === targetDateStr;
	});

	const label = offsetDays === 0 ? 'Today' : 'Yesterday';

	if (records.length === 0) {
		return `${label} no accounting records found. Please record your expenses!`;
	}

	return buildDailySummary(records, label);
}

/**
 * Get records from the past N days
 * @param {Sheet} sheet - Spreadsheet sheet for accounting
 * @param {number} days - Number of days to look back
 * @return {string} Daily summary of expenses for the past N days
 */
function getPastNDaysRecords(sheet, days) {
	const endDate = new Date();
	const startDate = new Date(endDate);
	startDate.setDate(endDate.getDate() - (days - 1));
	startDate.setHours(0, 0, 0, 0);
	endDate.setHours(23, 59, 59, 999);

	const records = getFilteredRecords(sheet, row => {
		const recordDate = new Date(row[0]);
		return recordDate >= startDate && recordDate <= endDate;
	});

	if (records.length === 0) {
		return `No accounting records found in the past ${days} days. Please record your expenses!`;
	}

	return buildDailySummary(records, `Past ${days} Days Summary`);
}

/**
 * Get monthly statistics from the sheet
 * @param {Sheet} sheet - Spreadsheet sheet for accounting
 * @param {number} year
 * @param {number} month
 * @return {string} Monthly statistics summary
 */
function getMonthlyStats(sheet, year, month) {
	const records = getFilteredRecords(sheet, row => {
		const recordDate = new Date(row[0]);
		return recordDate.getFullYear() === year && recordDate.getMonth() + 1 === month;
	});

	return buildStatsReply(records, `${year}/${month} Monthly Statistics`);
}

/**
 * Get yearly statistics from the sheet
 * @param {Sheet} sheet - Spreadsheet sheet for accounting
 * @param {number} year
 * @return {string} Yearly statistics summary
 */
function getYearlyStats(sheet, year) {
	const records = getFilteredRecords(sheet, row => {
		return new Date(row[0]).getFullYear() === year;
	});

	return buildStatsReply(records, `${year} Yearly Statistics`);
}

/**
 * Filter records from the sheet based on a custom filter function
 * @param {Sheet} sheet - Spreadsheet sheet to filter records from
 * @param {Function} filterFn - Function to filter records, should return true for records to keep
 * @return {Array}
 */
function getFilteredRecords(sheet, filterFn) {
	const allRows = sheet.getDataRange().getValues();
	if (allRows.length <= 1) return [];

	return allRows.slice(1).filter(row => row[0] && filterFn(row));
}

/**
 * Format daily summary of expenses
 * @param {Array} records - records array [[timestamp, amount, description, category], ...]
 * @param {string} label - description label
 * @return {string} Daily summary of expenses
 */
function buildDailySummary(records, label) {
	let total = 0;
	const details = records.map(([, amount, description, category]) => {
		const numAmount = parseFloat(amount) || 0;
		total += numAmount;
		const categoryText = category ? ` (@${category})` : '';
		return `・${description || 'no description'}：${numAmount}${categoryText}`;
	}).join('\n');

	return `[${label}]\nTotal Expenses: ${total}\nDetails:\n${details || 'No records found'}`;
}

/**
 * Format statistics summary from records
 * @param {Array} rows - records array [[timestamp, amount, description, category], ...]
 * @param {string} label - description label
 * @return {string} Statistics summary
 */
function buildStatsReply(rows, label) {
	if (!rows || rows.length === 0) {
		return `${label} no accounting records found. Please record your expenses!`;
	}

	let total = 0;
	const categoryMap = new Map();

	rows.forEach(([, amount, desc, category]) => {
		const numAmount = parseFloat(amount) || 0;
		total += numAmount;

		const c = category || 'Uncategorized';
		categoryMap.set(c, (categoryMap.get(c) || 0) + numAmount);
	});

	const summary = Array.from(categoryMap.entries())
		.map(([cat, amt]) => `・${cat}：${amt}`)
		.join('\n');

	return `${label} Total Expenses: ${total}\nCategory Summary:\n${summary || 'No records found'}`;
}

/**
 * Send a text message to a specific user
 * @param {string} userId
 * @param {string} message - Message to send
 */
function replyText(userId, message) {
	const usageNote = getSendUsage();
	const fullMessage = `${message}\n\n${usageNote}`;

	try {
		const response = UrlFetchApp.fetch(CONFIG.PUSH_URL, {
			method: 'POST',
			headers: {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': `Bearer ${CONFIG.CHANNEL_ACCESS_TOKEN}`
			},
			payload: JSON.stringify({
				to: userId,
				messages: [{
					type: 'text',
					text: fullMessage
				}]
			})
		});

		if (response.getResponseCode() !== 200) {
			writeLog(LEVEL.ERROR, 'main.gs:replyText', 'Send LINE Message', { statusCode: result.getResponseCode(), content: result.getContentText() });
			return;
		}
		recordSendTimes('Reply Text Message to specific user');
	} catch (error) {
		writeLog(LEVEL.EMERGENCY, 'main.gs:replyText', 'Send LINE Message', error);
	}
}


/**
 * @param {string} purpose - Purpose of the message sent
 */
function recordSendTimes(purpose) {
	const sheet = getSheet('Usage');
	if (!sheet) { return; }

	const timestamp = formatTimestamp(new Date());
	sheet.appendRow([timestamp, purpose]);
}

/**
 * Get the current month's message usage
 * @return {string} Message usage summary
 */
function getSendUsage() {
	try {
		const sheet = getSheet('Usage');
		if (!sheet) {
			return '';
		}

		const currentCount = getCurrentMonthMessageCount(sheet) + 1; // +1 for current message
		return `📤 This month has sent ${currentCount} / ${CONFIG.MONTHLY_MESSAGE_LIMIT} messages`;
	} catch (error) {
		writeLog(LEVEL.EMERGENCY, 'main.gs:getSendUsage', 'Calculate Usage', error);
		return '';
	}
}

/**
 * @param {Sheet} sheet - Spreadsheet sheet for counting messages
 * @return {number} Current month message count
 */
function getCurrentMonthMessageCount(sheet) {
	const lastRow = sheet.getLastRow();
	if (lastRow < 2) return 0;

	const data = sheet.getRange(2, 1, lastRow - 1).getValues();
	const now = new Date();
	const currentYear = now.getFullYear();
	const currentMonth = now.getMonth();

	return data.reduce((count, [timestamp]) => {
		const date = new Date(timestamp);
		return (date.getFullYear() === currentYear && date.getMonth() === currentMonth)
			? count + 1 : count;
	}, 0);
}

/**
 * @param {string} sheetName
 * @return {Sheet|null}
 */
function getSheet(sheetName) {
	try {
		return SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(sheetName);
	} catch (error) {
		writeLog(LEVEL.EMERGENCY, 'main.gs:getSheet', 'Open Sheet', `Failed to open sheet: ${sheetName}`);
		return null;
	}
}

/**
 * @param {Date} date
 * @return {string}
 */
function formatTimestamp(date) {
	return Utilities.formatDate(date, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
}

/**
 * @param {Date} date
 * @param {string} format - Date format string
 * @return {string}
 */
function formatDate(date, format) {
	return Utilities.formatDate(date, CONFIG.TIMEZONE, format);
}

/**
 * @param {string} functionName
 */
function cleanupTriggers(functionName) {
	const triggers = ScriptApp.getProjectTriggers();
	triggers.forEach(trigger => {
		if (trigger.getHandlerFunction() === functionName) {
			ScriptApp.deleteTrigger(trigger);
		}
	});
}

/**
 * @return {string}
 */
function getHelpText() {
	return `Format Error!
1️⃣ Accounting Format:
▪️Amount Description @Category
▪️Amount Description
2️⃣ Query Statistics:
▪️:0       ← Today
▪️:-1      ← Yesterday
▪️:-7      ← Past 7 Days
▪️:Number  ← This Month of the Year
▪️:Year    ← Specific Year`;
}