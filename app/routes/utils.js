// utils.js
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const Database = require("better-sqlite3");
const {app} = require("electron");

const {
	parseSimpleBiometric,
	isSimpleBiometricFormat,
} = require("./biometricParserSimple");

// Define the database file path

const userDataPath = app.getPath("userData");
const DB_FILE = path.join(userDataPath, "dtr.db");

// Ensure the data directory exists
const dataDir = path.dirname(DB_FILE);
if (!fs.existsSync(dataDir)) {
	fs.mkdirSync(dataDir, {
		recursive: true,
	});
}

let db; // Global database instance

/**
 * Initializes the SQLite database and creates tables if they don't exist.
 */
function initDb() {
	try {
		db = new Database(DB_FILE);
		console.log(`Connected to SQLite database: ${DB_FILE}`);

		// Enable WAL mode for better concurrency (recommended for web applications)
		db.pragma("journal_mode = WAL");

		// Create employees table with a composite unique key (userId, name, month)
		// This allows different people with the same userId to exist, or the same person
		// with the same userId/name for different months.
		db.exec(`
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                userId TEXT NOT NULL,
                name TEXT NOT NULL,
                department TEXT,
                month TEXT,
                attendanceDateRange TEXT,
                tablingDate TEXT,
                UNIQUE(userId, name, month) -- Composite unique key for employee records
            );
        `);

		// Create time_records table
		db.exec(`
            CREATE TABLE IF NOT EXISTS time_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                day INTEGER NOT NULL,
                dateWeekday TEXT,
                am_arrival TEXT,
                am_departure TEXT,
                pm_arrival TEXT,
                pm_departure TEXT,
                ot_arrival TEXT,
                ot_departure TEXT,
                total_hours TEXT,
                late TEXT,
                early_out TEXT,
                overtime TEXT,
                remarks TEXT,
                undertime_hours TEXT,
                undertime_minutes TEXT,
                FOREIGN KEY (employee_id) REFERENCES employees(id)
            );
        `);
		console.log("SQLite tables checked/created successfully.");
	} catch (error) {
		console.error("Error initializing SQLite database:", error);
		if (db) db.close(); // Ensure db is closed on error
		process.exit(1); // Exit if DB cannot be initialized
	}
}

// Initialize the database when utils.js is loaded
initDb();

// --- Constants for Excel Sheet Layout (used as search targets/relative offsets) ---
// Global headers (will be dynamically found)
const GLOBAL_DATE_RANGE_TEXT = "Date"; // Changed to match the image provided: "Date"
const TABLING_DATE_TEXT = "Tabling date:"; // Text to search for

// Key headers to search for within each employee block
const HEADER_NAMES = {
	DEPARTMENT: "Dept.",
	NAME: "Name",
	USER_ID: "User ID",
	DAY: "Date/Weekday", // Keep as 'Date/Weekday' but regex will be flexible
	AM_IN_OUT: "In", // Sub-header for AM In/Out
	PM_IN_OUT: "Out", // Sub-header for PM In/Out (using 'Out' for PM for now, to be dynamic later)
	OT_IN_OUT: "In", // Sub-header for OT In/Out
};

// Offsets relative to the found "Day" column for time data
const RELATIVE_OFFSETS = {
	AM_IN: 1, // B, Q, AF
	AM_OUT: 3, // D, S, AH
	PM_IN: 6, // G, V, AK
	PM_OUT: 8, // I, X, AM
	OT_IN: 10, // K, Z, AO
	OT_OUT: 12, // M, AA, AP
	LATE: 13, // N, AB, AQ
	UNDERTIME: 14, // O, AC, AR
	OVERTIME_HOURS: 19, // T, AI, AX
	TOTAL_HOURS: 20, // U, AJ, AY
};

// Row range where headers are expected to be found for dynamic block detection
const HEADER_SEARCH_START_ROW = 0; // Start from Row 1 (0-based) - sometimes headers can be very high
const HEADER_SEARCH_END_ROW = 100; // Increased search range for headers to cover more possibilities

/**
 * Helper function to convert Excel column letter (e.g., 'A', 'AA') to 0-based index.
 * @param {string} colLetter - The Excel column letter.
 * @returns {number} - The 0-based column index.
 */
function columnLetterToNumber(colLetter) {
	let col = 0;
	for (let i = 0; i < colLetter.length; i++) {
		col = col * 26 + (colLetter.charCodeAt(i) - "A".charCodeAt(0) + 1);
	}
	return col - 1; // Convert to 0-based
}

/**
 * Helper function to find a cell containing a specific value within a given range.
 * This version also attempts to find values within merged cells.
 * @param {object} sheet - The xlsx sheet object.
 * @param {string} value - The value to search for (case-insensitive, trimmed).
 * @param {object} searchRange - An object {s: {r, c}, e: {r, c}} defining the search area.
 * @returns {object|null} - {r: row_index, c: col_index} of the first match, or null if not found.
 */
function findCellByValue(sheet, value, searchRange) {
	const lowerCaseValue = value.trim().toLowerCase();

	// Create a more flexible regex for matching, accounting for different line breaks
	// and ensuring full string match after trimming.
	// This function specifically needs to handle the exact header text for Name, Dept, User ID
	// and also the start of global date strings.
	const regexValue = new RegExp(
		`^${value
			.trim()
			.toLowerCase()
			.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|\,\<\>\(\)]/g, "\\$&")
			.replace(/\s*\/\s*/g, "[\\s\\r\\n]*\\/[\\s\\r\\n]*")}$`,
		"i"
	);

	// 1. Check merged cells first
	if (sheet["!merges"]) {
		for (const merge of sheet["!merges"]) {
			// Check if the merged cell's top-left corner is within our search range
			if (
				merge.s.r >= searchRange.s.r &&
				merge.s.r <= searchRange.e.r &&
				merge.s.c >= searchRange.s.c &&
				merge.s.c <= searchRange.e.c
			) {
				const cellAddress = xlsx.utils.encode_cell(merge.s);
				const cell = sheet[cellAddress];
				const cellValueToTest = String(cell?.v || "")
					.trim()
					.toLowerCase();

				if (
					cell &&
					typeof cell.v === "string" &&
					regexValue.test(cellValueToTest)
				) {
					return merge.s; // Return top-left cell of the merged range
				}
			}
		}
	}

	// 2. If not found in merged cells, check individual cells within the search range
	for (let R = searchRange.s.r; R <= searchRange.e.r; ++R) {
		for (let C = searchRange.s.c; C <= searchRange.e.c; ++C) {
			const cellAddress = xlsx.utils.encode_cell({
				r: R,
				c: C,
			});
			const cell = sheet[cellAddress];
			const cellValueToTest = String(cell?.v || "")
				.trim()
				.toLowerCase();

			if (
				cell &&
				typeof cell.v === "string" &&
				regexValue.test(cellValueToTest)
			) {
				return {
					r: R,
					c: C,
				};
			}
		}
	}
	return null;
}

/**
 * Helper function to find ALL cells containing a specific value within a given range.
 * This version also attempts to find values within merged cells.
 * @param {object} sheet - The xlsx sheet object.
 * @param {string} value - The value to search for (case-insensitive, trimmed).
 * @param {object} searchRange - An object {s: {r, c}, e: {r, c}} defining the search area.
 * @returns {Array<object>} - An array of {r: row_index, c: col_index} for all matches.
 */
function findAllCellsByValue(sheet, value, searchRange) {
	let internalFoundCells = [];
	const lowerCaseSearchValue = value.trim().toLowerCase();
	let regexValue;

	// Special handling for the 'Date/Weekday' header to be flexible
	if (value === HEADER_NAMES.DAY) {
		// Matches "date" or "date/weekday" (with optional spaces/newlines around slash)
		// Also considers just "day" if it's in the correct context, but let's be more specific for now.
		// The regex now specifically handles "Date" or "Date/Weekday" with variations in whitespace.
		regexValue = new RegExp(`^date([\\s\\r\\n]*\\/[\\s\\r\\n]*weekday)?$`, "i");
	} else if (value === GLOBAL_DATE_RANGE_TEXT || value === TABLING_DATE_TEXT) {
		// For global date/tabling headers, match the start of the string
		regexValue = new RegExp(
			`^${lowerCaseSearchValue.replace(
				/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|\,\<\>\(\)]/g,
				"\\$&"
			)}`,
			"i"
		);
	} else if (
		value === HEADER_NAMES.AM_IN_OUT ||
		value === HEADER_NAMES.PM_IN_OUT ||
		value === HEADER_NAMES.OT_IN_OUT
	) {
		// For 'In' or 'Out' sub-headers, be exact
		regexValue = new RegExp(
			`^${lowerCaseSearchValue.replace(
				/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|\,\<\>\(\)]/g,
				"\\$&"
			)}$`,
			"i"
		);
	} else {
		// For other headers, use the standard exact match (case-insensitive)
		regexValue = new RegExp(
			`^${lowerCaseSearchValue.replace(
				/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|\,\<\>\(\)]/g,
				"\\$&"
			)}$`,
			"i"
		);
	}

	// Iterate through all cells in the search range
	for (let R = searchRange.s.r; R <= searchRange.e.r; ++R) {
		for (let C = searchRange.s.c; C <= searchRange.e.c; ++C) {
			const cellAddress = xlsx.utils.encode_cell({r: R, c: C});
			const cell = sheet[cellAddress];
			const cellValue = cell?.v; // Get raw value

			// Skip if cell has no value to reduce noise, as requested
			if (
				cellValue === undefined ||
				cellValue === null ||
				String(cellValue).trim() === ""
			) {
				continue;
			}

			const cellValueToTest = String(cellValue).trim().toLowerCase();

			const testResult = regexValue.test(cellValueToTest);

			if (typeof cellValue === "string" && testResult) {
				internalFoundCells.push({r: R, c: C});
			}
		}
	}
	return internalFoundCells;
}

/**
 * Parses an Excel date number into a JavaScript Date object.
 * Handles both numeric Excel dates and string dates.
 * @param {number|string} excelDate - The value from the Excel cell.
 * @returns {Date} - A JavaScript Date object.
 */
function parseExcelDate(excelDate) {
	if (typeof excelDate === "number") {
		const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
		const offsetMs = date.getTimezoneOffset() * 60 * 1000;
		return new Date(date.getTime() + offsetMs);
	}
	return new Date(excelDate);
}

/**
 * Helper function to calculate daily metrics (adapted for direct times)
 * @param {object} record - The daily time record object.
 * @param {number} year - The year of the record.
 * @param {number} monthNum - The 1-based month number.
 * @param {number} day - The day of the month.
 * @returns {object} - Calculated daily metrics.
 */
function calculateDailyMetricsFromTimes(record, year, monthNum, day) {
	const amArrival = record.am_arrival;
	const amDeparture = record.am_departure;
	const pmArrival = record.pm_arrival;
	const pmDeparture = record.pm_departure;
	const otArrival = record.ot_arrival;
	const otDeparture = record.ot_departure;

	let late = record.late_excel || "";
	let overtime = record.overtime_hours_excel || "";
	let totalHours = record.total_hours_excel || "";
	let earlyOut = "";

	let undertime_hours = "";
	let undertime_minutes = "";

	if (record.undertime_excel) {
		const utMatch = String(record.undertime_excel).match(/^(\d+):(\d+)$/);
		if (utMatch) {
			undertime_hours = parseInt(utMatch[1]);
			undertime_minutes = parseInt(utMatch[2]);
		}
	}

	const hasActualWorkEntries =
		amArrival ||
		amDeparture ||
		pmArrival ||
		pmDeparture ||
		otArrival ||
		otDeparture;

	let remarks = "";

	const dateString = `${year}-${String(monthNum).padStart(2, "0")}-${String(
		day
	).padStart(2, "0")}`;

	const amInTime = amArrival ? new Date(`${dateString} ${amArrival}`) : null;
	const amOutTime = amDeparture
		? new Date(`${dateString} ${amDeparture}`)
		: null;
	const pmInTime = pmArrival ? new Date(`${dateString} ${pmArrival}`) : null;
	const pmOutTime = pmDeparture
		? new Date(`${dateString} ${pmDeparture}`)
		: null;
	const otInTime = otArrival ? new Date(`${dateString} ${otArrival}`) : null;
	const otOutTime = otDeparture
		? new Date(`${dateString} ${otDeparture}`)
		: null;

	let morningDuration = 0;
	if (
		amInTime &&
		amOutTime &&
		!isNaN(amInTime) &&
		!isNaN(amInTime) &&
		amOutTime > amInTime
	) {
		morningDuration =
			(amOutTime.getTime() - amInTime.getTime()) / (1000 * 60 * 60);
	}

	let afternoonDuration = 0;
	if (
		pmInTime &&
		pmOutTime &&
		!isNaN(pmInTime) &&
		!isNaN(pmOutTime) &&
		pmOutTime > pmInTime
	) {
		afternoonDuration =
			(pmOutTime.getTime() - pmInTime.getTime()) / (1000 * 60 * 60);
	}

	let overtimeCalculatedDuration = 0;
	if (
		otInTime &&
		otOutTime &&
		!isNaN(otInTime) &&
		!isNaN(otOutTime) &&
		otOutTime > otInTime
	) {
		overtimeCalculatedDuration =
			(otOutTime.getTime() - otInTime.getTime()) / (1000 * 60 * 60);
	}

	const totalWorkDurationHours = morningDuration + afternoonDuration;

	const standardStartTime = new Date(`${dateString} 08:00`);
	const standardEndTime = new Date(`${dateString} 17:00`);
	const standardDailyHours = 8;
	const standardWorkMinutes = 480;

	const date = new Date(year, monthNum - 1, day);
	const dayOfWeek = date.getDay();

	if (undertime_hours === "" || undertime_minutes === "") {
		if (hasActualWorkEntries) {
			let workedMinutes = 0;

			if (
				amInTime &&
				amOutTime &&
				!isNaN(amInTime) &&
				!isNaN(amOutTime) &&
				amOutTime > amInTime
			) {
				workedMinutes +=
					(amOutTime.getTime() - amInTime.getTime()) / (1000 * 60);
			}
			if (
				pmInTime &&
				pmOutTime &&
				!isNaN(pmInTime) &&
				!isNaN(pmOutTime) &&
				pmOutTime > pmInTime
			) {
				workedMinutes +=
					(pmOutTime.getTime() - pmInTime.getTime()) / (1000 * 60);
			}

			const calculatedUndertimeMinutes = Math.max(
				standardWorkMinutes - workedMinutes,
				0
			);
			if (calculatedUndertimeMinutes > 0) {
				undertime_hours = Math.floor(calculatedUndertimeMinutes / 60);
				undertime_minutes = Math.round(calculatedUndertimeMinutes % 60);
			} else {
				undertime_hours = "";
				undertime_minutes = "";
			}
		} else {
			// If no time was entered
			if (dayOfWeek === 6) {
				// Saturday
				undertime_hours = "Saturday";
				undertime_minutes = "Saturday";
			} else if (dayOfWeek === 0) {
				// Sunday
				undertime_hours = "Sunday";
				undertime_minutes = "Sunday";
			} else {
				undertime_hours = "";
				undertime_minutes = "";
			}
		}
	}

	if (!late && amInTime && !isNaN(amInTime) && amInTime > standardStartTime) {
		const lateMinutes =
			(amInTime.getTime() - standardStartTime.getTime()) / (1000 * 60);
		late = `${Math.floor(lateMinutes / 60)}h ${String(
			Math.round(lateMinutes % 60)
		).padStart(2, "0")}m`;
		remarks = (remarks ? remarks + ", " : "") + "Late";
	}

	if (
		!earlyOut &&
		pmOutTime &&
		!isNaN(pmOutTime) &&
		pmOutTime < standardEndTime
	) {
		const earlyOutMinutes =
			(standardEndTime.getTime() - pmOutTime.getTime()) / (1000 * 60);
		earlyOut = `${Math.floor(earlyOutMinutes / 60)}h ${String(
			Math.round(earlyOutMinutes % 60)
		).padStart(2, "0")}m`;
		remarks = (remarks ? remarks + ", " : "") + "Early Out";
	}

	if (!overtime) {
		let totalWorkHoursForOTCalc =
			totalWorkDurationHours + overtimeCalculatedDuration;
		if (totalWorkHoursForOTCalc > standardDailyHours) {
			const otCalculated = totalWorkHoursForOTCalc - standardDailyHours;
			overtime = `${Math.floor(otCalculated)}h ${String(
				Math.round((otCalculated % 1) * 60)
			).padStart(2, "0")}m`;
			remarks = (remarks ? remarks + ", " : "") + "Overtime";
		} else {
			overtime = "";
		}
	}

	if (
		!totalHours &&
		hasActualWorkEntries &&
		totalWorkDurationHours + overtimeCalculatedDuration > 0
	) {
		let totalWorkHoursIncludingOT =
			totalWorkDurationHours + overtimeCalculatedDuration;
		totalHours = `${Math.floor(totalWorkHoursIncludingOT)}h ${String(
			Math.round((totalWorkHoursIncludingOT % 1) * 60)
		).padStart(2, "0")}m`;
	} else if (!totalHours) {
		totalHours = "";
	}

	return {
		am_arrival: amArrival,
		am_departure: amDeparture,
		pm_arrival: pmArrival,
		pm_departure: pmDeparture,
		ot_arrival: otArrival,
		ot_departure: otDeparture,
		total_hours: totalHours,
		late: late,
		early_out: earlyOut,
		overtime: overtime,
		remarks: remarks,
		undertime_hours: String(undertime_hours), // Ensure these are strings for consistency with DB schema
		undertime_minutes: String(undertime_minutes), // Ensure these are strings for consistency with DB schema
	};
}

/**
 * Parses a single "Employee Attendance Table" sheet and prepares data for DB.
 * This function now dynamically locates all employee blocks on the sheet.
 * @param {object} sheet - The xlsx sheet object.
 * @returns {Array<object>} - An array of parsed employee data from this sheet, ready for DB insertion.
 */
function parseSingleEmployeeAttendanceSheet(sheet) {
	const parsedEmployees = [];
	// Log the full range of the sheet to verify its bounds
	console.log(`DEBUG Sheet '${sheet.name}': Full range: ${sheet["!ref"]}`);
	const sheetFullRange = xlsx.utils.decode_range(sheet["!ref"]);

	console.log(`--- Parsing Sheet: ${sheet.name || "Unnamed Sheet"} ---`);

	const getCellCoords = (cellAddr) => {
		const col = cellAddr.charCodeAt(0) - 65;
		const row = parseInt(cellAddr.substring(1)) - 1;
		return {
			r: row,
			c: col,
		};
	};

	// --- Extract global attendance date range and tabling date dynamically ---
	let attendanceDateRange = "";
	let tablingDate = "";

	const globalHeaderSearchRange = {
		// Search for global headers in a reasonable top-left area
		s: {r: 0, c: 0},
		e: {r: HEADER_SEARCH_END_ROW, c: sheetFullRange.e.c},
	};

	// Find Attendance Date Range
	const attendanceDateRangeCell = findAllCellsByValue(
		sheet,
		GLOBAL_DATE_RANGE_TEXT,
		globalHeaderSearchRange
	);
	if (attendanceDateRangeCell.length > 0) {
		// Take the value from the cell next to the found "Attendance date:" text
		const cellRef = xlsx.utils.encode_cell({
			r: attendanceDateRangeCell[0].r,
			c: attendanceDateRangeCell[0].c + 1,
		});
		attendanceDateRange = String(sheet[cellRef]?.v || "").trim();
	}
	console.log(`Attendance Date Range: ${attendanceDateRange}`);

	// Find Tabling Date
	const tablingDateCell = findAllCellsByValue(
		sheet,
		TABLING_DATE_TEXT,
		globalHeaderSearchRange
	);
	if (tablingDateCell.length > 0) {
		// Corrected: First get the cell object, then apply optional chaining to its 'v' property.
		const cellAddressForTablingDate = xlsx.utils.encode_cell({
			r: tablingDateCell[0].r,
			c: tablingDateCell[0].c + 1,
		});
		const tablingDateRaw = sheet[cellAddressForTablingDate]?.v;

		if (typeof tablingDateRaw === "number") {
			tablingDate = parseExcelDate(tablingDateRaw).toISOString().split("T")[0];
		} else {
			tablingDate = String(tablingDateRaw || "").trim();
		}
	}
	console.log(`Tabling Date: ${tablingDate}`);

	let reportMonth = "";
	// NEW REGEX: More flexible to capture YYYY-MM from YYYY-MM-DD~YYYY-MM-DD format
	// This will now correctly capture "2025-04" from "2025-04-01~2025-04-08"
	const dateRangeMatch = attendanceDateRange.match(
		/^(\d{4}-\d{2})-\d{2}~\d{4}-\d{2}-\d{2}$/
	);
	if (dateRangeMatch && dateRangeMatch[1]) {
		reportMonth = dateRangeMatch[1];
	} else {
		// Fallback to current month if detection fails
		const today = new Date();
		reportMonth = `${today.getFullYear()}-${String(
			today.getMonth() + 1
		).padStart(2, "0")}`;
	}
	console.log(`Report Month: ${reportMonth}`);

	// --- Dynamic Employee Block Discovery ---
	const discoveredDayColumnIndices = new Set();
	const allHeaderLocations = []; // To help determine data start row

	const dayHeaderSearchRange = {
		s: {r: HEADER_SEARCH_START_ROW, c: 0}, // Start from column 0
		e: {r: sheetFullRange.e.r, c: sheetFullRange.e.c}, // Search all columns
	};
	console.log(
		`Searching for '${HEADER_NAMES.DAY}' in range: ${xlsx.utils.encode_range(
			dayHeaderSearchRange
		)}`
	);

	// Call findAllCellsByValue to get all instances of the Day header
	const dayHeaderLocations = findAllCellsByValue(
		sheet,
		HEADER_NAMES.DAY,
		dayHeaderSearchRange
	);

	if (dayHeaderLocations.length === 0) {
		console.warn(
			`No '${HEADER_NAMES.DAY}' headers found in sheet '${sheet.name}'. Skipping sheet processing.`
		);
		return [];
	}

	// Populate discoveredDayColumnIndices and find maxDayHeaderRow
	let maxDayHeaderRow = -1;
	dayHeaderLocations.forEach((loc) => {
		discoveredDayColumnIndices.add(loc.c);
		allHeaderLocations.push(loc);
		if (loc.r > maxDayHeaderRow) {
			maxDayHeaderRow = loc.r;
		}
	});

	const sortedDayColumnIndices = Array.from(discoveredDayColumnIndices).sort(
		(a, b) => a - b
	);
	console.log(
		`Discovered Day Column Indices: ${sortedDayColumnIndices
			.map((c) => xlsx.utils.encode_col(c))
			.join(", ")}`
	);

	if (sortedDayColumnIndices.length === 0) {
		console.warn(
			`No '${HEADER_NAMES.DAY}' headers found in sheet '${sheet.name}'. Skipping sheet processing.`
		);
		return [];
	}

	// Process each discovered employee block
	sortedDayColumnIndices.forEach((dayColIndex, index) => {
		console.log(
			`--- Processing Discovered Employee Block ${
				index + 1
			} (Day Column: ${xlsx.utils.encode_col(dayColIndex)}) ---`
		);

		// Determine the column search range for headers within this specific block
		const nextBlockDayColIndex =
			index + 1 < sortedDayColumnIndices.length
				? sortedDayColumnIndices[index + 1]
				: sheetFullRange.e.c + 1; // If last block, search till end of sheet

		const blockHeaderSearchRange = {
			s: {r: HEADER_SEARCH_START_ROW, c: dayColIndex},
			e: {r: HEADER_SEARCH_END_ROW, c: nextBlockDayColIndex - 1},
		};
		console.log(
			`Block Header search range for block ${
				index + 1
			}: ${xlsx.utils.encode_range(blockHeaderSearchRange)}`
		);

		let employeeData = {
			department: "",
			name: "",
			userId: "",
			attendanceDateRange: attendanceDateRange,
			tablingDate: tablingDate,
			month: reportMonth,
			timeCard: [],
		};

		// Find Name, User ID, Department headers within this block's range
		const nameHeaderCell = findCellByValue(
			sheet,
			HEADER_NAMES.NAME,
			blockHeaderSearchRange
		);
		const userIdHeaderCell = findCellByValue(
			sheet,
			HEADER_NAMES.USER_ID,
			blockHeaderSearchRange
		);
		const departmentHeaderCell = findCellByValue(
			sheet,
			HEADER_NAMES.DEPARTMENT,
			blockHeaderSearchRange
		);

		// Extract employee details
		if (nameHeaderCell) {
			employeeData.name = String(
				sheet[
					xlsx.utils.encode_cell({r: nameHeaderCell.r, c: nameHeaderCell.c + 1})
				]?.v || ""
			).trim();
			console.log(
				`Found Name header at ${xlsx.utils.encode_cell(
					nameHeaderCell
				)}. Extracted Name: '${employeeData.name}'`
			);
		} else {
			console.warn(
				`'Name' header not found for block ${
					index + 1
				}. Employee name will be empty.`
			);
		}

		if (userIdHeaderCell) {
			employeeData.userId = String(
				sheet[
					xlsx.utils.encode_cell({
						r: userIdHeaderCell.r,
						c: userIdHeaderCell.c + 1,
					})
				]?.v || ""
			).trim();
			console.log(
				`Found User ID header at ${xlsx.utils.encode_cell(
					userIdHeaderCell
				)}. Extracted User ID: '${employeeData.userId}'`
			);
		} else {
			console.warn(
				`'User ID' header not found for block ${
					index + 1
				}. User ID will be empty.`
			);
		}

		if (departmentHeaderCell) {
			employeeData.department = String(
				sheet[
					xlsx.utils.encode_cell({
						r: departmentHeaderCell.r,
						c: departmentHeaderCell.c + 1,
					})
				]?.v || ""
			).trim();
			console.log(
				`Found Department header at ${xlsx.utils.encode_cell(
					departmentHeaderCell
				)}. Extracted Department: '${employeeData.department}'`
			);
		} else {
			console.warn(
				`'Department' header not found for block ${
					index + 1
				}. Department will be empty.`
			);
		}

		// Validate essential info
		if (!employeeData.name || !employeeData.userId) {
			console.warn(
				`SKIPPING EMPLOYEE BLOCK ${index + 1}: Missing Name ('${
					employeeData.name
				}') or User ID ('${employeeData.userId}').`
			);
			return; // Skip this block
		}

		// Determine actual start of timecard data by finding the lowest row of 'In' or 'Out' sub-headers
		let timecardDataStartRow = -1; // Declared here for proper scope
		const subHeaderSearchRange = {
			s: {r: maxDayHeaderRow, c: dayColIndex}, // Start searching from the row of 'Date/Weekday' header
			e: {r: maxDayHeaderRow + 5, c: nextBlockDayColIndex - 1}, // Search a few rows below
		};

		const inHeaders = findAllCellsByValue(
			sheet,
			HEADER_NAMES.AM_IN_OUT,
			subHeaderSearchRange
		);
		const outHeaders = findAllCellsByValue(
			sheet,
			HEADER_NAMES.PM_IN_OUT,
			subHeaderSearchRange
		);

		const allSubHeaders = [...inHeaders, ...outHeaders];

		if (allSubHeaders.length > 0) {
			timecardDataStartRow =
				allSubHeaders.reduce((maxRow, cell) => Math.max(maxRow, cell.r), -1) +
				1;
		} else {
			// Fallback if 'In'/'Out' headers are not found, use maxDayHeaderRow + fixed offset
			console.warn(
				`'In' or 'Out' sub-headers not found for Block ${
					index + 1
				}. Falling back to maxDayHeaderRow + 2.`
			);
			timecardDataStartRow = maxDayHeaderRow + 2;
		}

		console.log(
			`Timecard data expected to start at row: ${
				timecardDataStartRow + 1
			} (0-based: ${timecardDataStartRow}) for Block ${index + 1}.`
		);

		const dailyExcelRecords = new Map();
		for (let R = timecardDataStartRow; R <= sheetFullRange.e.r; ++R) {
			const dayCell = sheet[xlsx.utils.encode_cell({r: R, c: dayColIndex})];
			const dateWeekdayRaw = dayCell?.v;

			// Stop processing if we hit an empty "Day" cell, indicating end of data for this block
			// Also ensure it's not the very first row of data, as that might legitimately be empty
			if (!dateWeekdayRaw && R > timecardDataStartRow) {
				console.log(
					`Detected end of timecard data for Block ${index + 1} at row ${
						R + 1
					} (empty day cell).`
				);
				break;
			}

			let day = null;
			if (dateWeekdayRaw) {
				const dayMatch = String(dateWeekdayRaw).match(/^(\d+)/);
				if (dayMatch) {
					day = parseInt(dayMatch[1]);
				} else if (typeof dateWeekdayRaw === "number") {
					day = dateWeekdayRaw;
				}
			}

			if (
				day !== null &&
				day >= 1 &&
				day <=
					new Date(
						parseInt(reportMonth.split("-")[0]),
						parseInt(reportMonth.split("-")[1]),
						0
					).getDate()
			) {
				const amIn =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.AM_IN,
						})
					]?.v || "";
				const amOut =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.AM_OUT,
						})
					]?.v || "";
				const pmIn =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.PM_IN,
						})
					]?.v || "";
				const pmOut =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.PM_OUT,
						})
					]?.v || "";
				const otIn =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.OT_IN,
						})
					]?.v || "";
				const otOut =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.OT_OUT,
						})
					]?.v || "";
				const lateExcel =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.LATE,
						})
					]?.v || "";
				const undertimeExcel =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.UNDERTIME,
						})
					]?.v || "";
				const overtimeHoursExcel =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.OVERTIME_HOURS,
						})
					]?.v || "";
				const totalHoursExcel =
					sheet[
						xlsx.utils.encode_cell({
							r: R,
							c: dayColIndex + RELATIVE_OFFSETS.TOTAL_HOURS,
						})
					]?.v || "";

				dailyExcelRecords.set(day, {
					dateWeekday: String(dateWeekdayRaw || "").trim(),
					am_arrival: String(amIn).trim(),
					am_departure: String(amOut).trim(),
					pm_arrival: String(pmIn).trim(),
					pm_departure: String(pmOut).trim(),
					ot_arrival: String(otIn).trim(),
					ot_departure: String(otOut).trim(),
					late_excel: String(lateExcel).trim(),
					undertime_excel: String(undertimeExcel).trim(),
					overtime_hours_excel: String(overtimeHoursExcel).trim(),
					total_hours_excel: String(totalHoursExcel).trim(),
				});
			} else if (dateWeekdayRaw) {
				console.warn(
					`Invalid day extracted '${dateWeekdayRaw}' for row ${
						R + 1
					} in Block ${index + 1}. Skipping record.`
				);
			}
		}

		const [yearStr, monthStr] = employeeData.month.split("-");
		const year = parseInt(yearStr);
		const monthNum = parseInt(monthStr);
		const daysInMonth = new Date(year, monthNum, 0).getDate();

		for (let day = 1; day <= daysInMonth; day++) {
			const excelRecord = dailyExcelRecords.get(day) || {};
			const dailyMetrics = calculateDailyMetricsFromTimes(
				// Corrected call: removed 'exports.'
				excelRecord,
				year,
				monthNum,
				day
			);

			const formattedDateWeekday = `${String(day).padStart(2, "0")} ${new Date(
				year,
				monthNum - 1,
				day
			).toLocaleString("en-US", {weekday: "short"})}`;

			employeeData.timeCard.push({
				day: day,
				dateWeekday: formattedDateWeekday,
				am_arrival: dailyMetrics.am_arrival,
				am_departure: dailyMetrics.am_departure,
				pm_arrival: dailyMetrics.pm_arrival,
				pm_departure: dailyMetrics.pm_departure,
				ot_arrival: dailyMetrics.ot_arrival,
				ot_departure: dailyMetrics.ot_departure,
				total_hours: dailyMetrics.total_hours,
				late: dailyMetrics.late,
				early_out: dailyMetrics.early_out,
				overtime: dailyMetrics.overtime,
				remarks: dailyMetrics.remarks,
				undertime_hours: dailyMetrics.undertime_hours,
				undertime_minutes: dailyMetrics.undertime_minutes,
			});
		}
		parsedEmployees.push(employeeData);
		console.log(
			`Successfully parsed employee: '${employeeData.name}' (User ID: ${
				employeeData.userId
			}) from Block ${index + 1}.`
		);
	});
	return parsedEmployees;
}

/**
 * Main function to load Excel file and parse data from specified sheets,
 * then store it into the SQLite database.
 * @param {string} filePath - The path to the Excel file.
 */
// function loadMultiSheetEmployeeAttendance(filePath) {
//     try {
//         const workbook = xlsx.readFile(filePath);
//         const allParsedEmployeesData = [];

//         // Loop through sheets from index 0 to 19 (for 1-20 sheets as specified)
//         for (let i = 0; i < 20; i++) {
//             const sheetName = workbook.SheetNames[i];
//             if (!sheetName) {
//                 console.warn(`No sheet found at index ${i}. Stopping sheet processing.`);
//                 break;
//             }
//             const sheet = workbook.Sheets[sheetName];
//             sheet.name = sheetName;

//             const parsedDataFromSheet = parseSingleEmployeeAttendanceSheet(sheet);
//             allParsedEmployeesData.push(...parsedDataFromSheet);
//         }
function loadMultiSheetEmployeeAttendance(filePath) {
	try {
		const workbook = xlsx.readFile(filePath);
		let allParsedEmployeesData = []; // Changed to 'let' as its value can be reassigned

		const firstSheetName = workbook.SheetNames[0];
		if (!firstSheetName) {
			console.error(`Error: No sheets found in Excel file: ${filePath}`);
			return;
		}
		const firstSheet = workbook.Sheets[firstSheetName];
		firstSheet.name = firstSheetName; // Assign name for logging

		// Detect format
		// HERE: isSimpleBiometricFormat is called
		if (isSimpleBiometricFormat(firstSheet)) {
			console.log(`Detected Simple Biometric Format for file: ${filePath}`);
			// HERE: parseSimpleBiometric is called if it's the simple format
			allParsedEmployeesData = parseSimpleBiometric(filePath);
			// The daily metrics are now applied within parseSimpleBiometric before it returns.
			// So, no need to apply them here again.
		} else {
			console.log(
				`Detected Multi-Sheet/Complex Biometric Format for file: ${filePath}`
			);
			// Existing loop for old format, assuming it might span multiple sheets
			for (let i = 0; i < workbook.SheetNames.length && i < 20; i++) {
				const sheetName = workbook.SheetNames[i];
				if (!sheetName) {
					console.warn(
						`No sheet found at index ${i}. Stopping sheet processing.`
					);
					break;
				}
				const sheet = workbook.Sheets[sheetName];
				sheet.name = sheetName;
				const parsedDataFromSheet = parseSingleEmployeeAttendanceSheet(sheet);
				allParsedEmployeesData.push(...parsedDataFromSheet);
			}
		}
		const insertEmployeeStmt = db.prepare(`
            INSERT INTO employees (userId, name, department, month, attendanceDateRange, tablingDate)
            VALUES (?, ?, ?, ?, ?, ?)
        `);
		const updateEmployeeStmt = db.prepare(`
            UPDATE employees
            SET department = ?, attendanceDateRange = ?, tablingDate = ?
            WHERE id = ?
        `);
		const insertTimeRecordStmt = db.prepare(`
            INSERT INTO time_records (employee_id, day, dateWeekday, am_arrival, am_departure, pm_arrival, pm_departure, ot_arrival, ot_departure, total_hours, late, early_out, overtime, remarks, undertime_hours, undertime_minutes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `);

		db.transaction(() => {
			for (const employee of allParsedEmployeesData) {
				let employeeId;
				// Try to find an existing employee record based on composite key
				const existingEmployee = db
					.prepare(
						"SELECT id FROM employees WHERE userId = ? AND name = ? AND month = ?"
					)
					.get(employee.userId, employee.name, employee.month);

				if (existingEmployee) {
					employeeId = existingEmployee.id;
					// Update the existing record if found
					updateEmployeeStmt.run(
						employee.department,
						employee.attendanceDateRange,
						employee.tablingDate,
						employeeId
					);
					console.log(
						`Updated existing data for employee: ${employee.name} (User ID: ${employee.userId}) for month ${employee.month} (ID: ${employeeId})`
					);
				} else {
					// Insert a new record if no matching composite key is found
					const info = insertEmployeeStmt.run(
						employee.userId,
						employee.name,
						employee.department,
						employee.month,
						employee.attendanceDateRange,
						employee.tablingDate
					);
					employeeId = info.lastInsertRowid;
					console.log(
						`Inserted new data for employee: ${employee.name} (User ID: ${employee.userId}) for month ${employee.month} (ID: ${employeeId})`
					);
				}

				// Delete old time records for this employee_id (now uniquely identified by composite key)
				// This ensures we always have the latest daily records for the current employee/month.
				db.prepare("DELETE FROM time_records WHERE employee_id = ?").run(
					employeeId
				);

				for (const record of employee.timeCard) {
					// Only insert if there's actual data for the day beyond just day/weekday
					if (
						record.am_arrival ||
						record.am_departure ||
						record.pm_arrival ||
						record.pm_departure ||
						record.ot_arrival ||
						record.ot_departure
					) {
						insertTimeRecordStmt.run(
							employeeId,
							record.day,
							record.dateWeekday,
							record.am_arrival,
							record.am_departure,
							record.pm_arrival,
							record.pm_departure,
							record.ot_arrival,
							record.ot_departure,
							record.total_hours,
							record.late,
							record.early_out,
							record.overtime,
							record.remarks,
							record.undertime_hours,
							record.undertime_minutes
						);
					}
				}
				console.log(
					`Processed timecard entries for employee: ${employee.name} (User ID: ${employee.userId}) for month ${employee.month}`
				);
			}
		})();

		console.log(
			"Biometric data loaded and stored successfully in SQLite from:",
			filePath
		);

		// DIAGNOSTIC: Dump all employees from DB after load
		try {
			const allEmployeesInDb = db
				.prepare("SELECT id, userId, name, month FROM employees")
				.all();
			console.log("--- DB CONTENTS AFTER INITIAL LOAD ---");
			console.log(allEmployeesInDb);
			console.log("------------------------------------");

			// NEW DIAGNOSTIC: Dump time_records for the first employee
			if (allEmployeesInDb.length > 0) {
				const firstEmployeeId = allEmployeesInDb[0].id;
				const sampleTimeRecords = db
					.prepare(
						"SELECT day, am_arrival, am_departure, pm_arrival, pm_departure, ot_arrival, ot_departure, undertime_hours, undertime_minutes FROM time_records WHERE employee_id = ? ORDER BY day ASC LIMIT 5"
					)
					.all(firstEmployeeId);
				console.log(
					`--- SAMPLE TIME RECORDS FOR EMPLOYEE ID ${firstEmployeeId} ---`
				);
				console.log(sampleTimeRecords);
				console.log("------------------------------------");
			}
		} catch (e) {
			console.error("Error dumping DB contents after load:", e);
		}
	} catch (err) {
		console.error(
			"Error loading or parsing Excel file and storing to DB:",
			err
		);
	}
}

/**
 * Retrieves employee data and their time records from the database.
 * @param {string} empNameQuery - The name of the employee to search for (case-insensitive, partial match).
 * @returns {Promise<Array<object>>} - An array of employee data objects, each including timeCard.
 */
// utils.js (corrected getEmployeeDataFromDb function)
async function getEmployeeDataFromDb(nameQuery = '', monthQuery = '') {
    try {
        console.log(`DEBUG DB Search: Received nameQuery: '${nameQuery}', monthQuery: '${monthQuery}'`);
        let employees = [];
        // Start with a base query that can be extended
        let query = "SELECT * FROM employees WHERE 1=1"; 
        const params = [];

        // Add name filter if nameQuery is provided
        if (nameQuery) {
            query += " AND name LIKE ? COLLATE NOCASE";
            params.push(`%${nameQuery}%`);
        }

        // Add month filter if monthQuery is provided
        if (monthQuery) {
            query += " AND month = ?";
            params.push(monthQuery);
        }
        
        // If no specific queries, return all employees (for initial dropdown population)
        if (!nameQuery && !monthQuery) {
            query = "SELECT * FROM employees";
        }

        const stmt = db.prepare(query);
        console.log(`DEBUG DB Search: Executing SQL query: "${query}" with parameters:`, params);
        employees = stmt.all(...params); // Spread parameters for stmt.all
        console.log(`DEBUG DB Search: Raw DB results for name='${nameQuery}', month='${monthQuery}':`, employees);

        const results = [];
        for (const emp of employees) {
            // REVERTED: Query time_records ONLY by employee_id, as 'month' column doesn't exist here.
            const timeRecordsStmt = db.prepare(`
                SELECT * FROM time_records
                WHERE employee_id = ?
                ORDER BY day ASC
            `);
            // Pass only emp.id to the time_records query
            const timeCards = timeRecordsStmt.all(emp.id); 

            const dailyRecordMap = new Map();
            timeCards.forEach((record) => {
                dailyRecordMap.set(record.day, record);
            });

            // The rest of the logic for parsing emp.month and generating formDataForDisplay
            // remains the same, as 'emp.month' comes from the 'employees' table which is filtered correctly.
            const [yearStr, monthStr] = (emp.month || "").split("-");
            const year = parseInt(yearStr);
            const monthNum = parseInt(monthStr);

            let daysInMonth = 31;
            if (!isNaN(year) && !isNaN(monthNum) && monthNum >= 1 && monthNum <= 12) {
                daysInMonth = new Date(year, monthNum, 0).getDate();
            } else {
                console.warn(
                    `Invalid month format '${emp.month}' for employee ${emp.name}. Defaulting to 31 days.`
                );
            }

            const formDataForDisplay = {};
            for (let day = 1; day <= daysInMonth; day++) {
                const currentDate = new Date(year, monthNum - 1, day);
                const dayName = currentDate.toLocaleString("en-US", {weekday: "long"});
                const shortDayName = currentDate.toLocaleString("en-US", {
                    weekday: "short",
                });
                const isWeekend =
                    currentDate.getDay() === 0 || currentDate.getDay() === 6;

                formDataForDisplay[day] = dailyRecordMap.get(day) || {
                    day: day,
                    dateWeekday: `${String(day).padStart(2, "0")} ${shortDayName}`,
                    am_arrival: "",
                    am_departure: "",
                    pm_arrival: "",
                    pm_departure: "",
                    ot_arrival: "",
                    ot_departure: "",
                    total_hours: "",
                    late: "",
                    early_out: "",
                    overtime: "",
                    remarks: "",
                    undertime_hours: isWeekend ? dayName : "", 
                    undertime_minutes: isWeekend ? dayName : "", 
                };
            }

            results.push({
                empId: emp.userId,
                empName: emp.name,
                department: emp.department,
                months: [emp.month], 
                formData: {
                    [emp.month]: formDataForDisplay,
                },
                attendanceDateRange: emp.attendanceDateRange,
                tablingDate: emp.tablingDate,
            });
        }
        console.log(
            `DEBUG DB Search: Final processed results count for name='${nameQuery}', month='${monthQuery}': ${results.length}`
        );
        return results;
    } catch (error) {
        console.error("Error retrieving employee data from DB:", error);
        return [];
    }
}

function convertTo12HourFormat(time24hr) {
    if (!time24hr || time24hr.trim() === '') {
        return ''; // Return empty for empty or null times
    }

    try {
        const [hours, minutes] = time24hr.split(':').map(Number);
        if (isNaN(hours) || isNaN(minutes)) {
            return time24hr; // Return original if not a valid time string
        }

        const period = hours >= 12 ? 'PM' : 'AM';
        const displayHours = hours % 12 || 12; // Convert 0 to 12 for 12 AM
        const displayMinutes = String(minutes).padStart(2, '0');

        return `${displayHours}:${displayMinutes} ${period}`;
    } catch (error) {
        console.error("Error converting time:", time24hr, error);
        return time24hr; // Return original in case of unexpected error
    }
}

module.exports = {
	convertTo12HourFormat,
	parseExcelDate,
	parseSingleEmployeeAttendanceSheet,
	loadMultiSheetEmployeeAttendance,
	calculateDailyMetricsFromTimes,
	getEmployeeDataFromDb,
};
