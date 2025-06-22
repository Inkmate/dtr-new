/**
 * controllers/attendanceController.js
 *
 * This controller manages the business logic for parsing attendance data
 * and calculating time metrics. It maintains the application's state.
 */
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// This variable will hold our in-memory data. It's encapsulated within this module.
let biometricData = {};

// --- All your helper and parsing functions are moved here ---

function parseExcelDate(excelDate) {
    // ... (implementation from your original file)
	if (typeof excelDate === "number") {
		const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
		const offsetMs = date.getTimezoneOffset() * 60 * 1000;
		return new Date(date.getTime() + offsetMs);
	}
	return new Date(excelDate);
}


function calculateDailyMetricsFromTimes(record, year, monthNum, day) {
    // ... (implementation from your original file)
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
    // ... rest of the function
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
		remarks: "", // Simplified for brevity
		undertime_hours: undertime_hours,
		undertime_minutes: undertime_minutes,
	};
}


function parseSingleEmployeeAttendanceSheet(sheet) {
    // ... (implementation from your original file)
    const sheetData = {};
    // ... logic to parse a sheet
    return sheetData;
}

/**
 * Loads and parses the multi-sheet Excel file, updating the internal biometricData state.
 * @param {string} filePath - The path to the Excel file.
 */
function loadMultiSheetEmployeeAttendance(filePath) {
	try {
		const workbook = xlsx.readFile(filePath);
		const allParsedEmployeesData = {};
		for (let i = 2; i <= 4; i++) {
			const sheetName = workbook.SheetNames[i];
			if (!sheetName) break;
			const sheet = workbook.Sheets[sheetName];
			sheet.name = sheetName;
			const parsedDataFromSheet = parseSingleEmployeeAttendanceSheet(sheet);
			Object.assign(allParsedEmployeesData, parsedDataFromSheet);
		}
		biometricData = allParsedEmployeesData;
        console.log("Biometric data reloaded successfully.");
	} catch (err) {
        console.error("Error loading Excel file:", err);
		biometricData = {};
	}
}


// --- EXPORTED FUNCTIONS ---
// These are the public methods that our route files will use.

module.exports = {
    // Loads the data from the default path at startup
    initialLoad: () => {
        const defaultExcelPath = path.join(__dirname, "../uploads/DTR.xlsx");
        if (fs.existsSync(defaultExcelPath)) {
            loadMultiSheetEmployeeAttendance(defaultExcelPath);
        } else {
            console.warn("Default DTR.xlsx not found in uploads folder.");
        }
    },

    // Function to handle reloading data after an upload
    reloadData: (filePath) => {
        loadMultiSheetEmployeeAttendance(filePath);
    },

    // Gets all currently loaded biometric data
    getBiometricData: () => biometricData,

    // Gets a list of all employees for dropdowns etc.
    getAllEmployees: () => {
        return Object.values(biometricData).map((emp) => ({
			userId: emp.userId,
			name: emp.name,
		}));
    },

    // Finds a single employee by ID or Name
    findEmployee: ({ empId, empName }) => {
        if (empId) {
            return Object.values(biometricData).find(e => e.userId === empId) || null;
        }
        if (empName) {
            return Object.values(biometricData).find(e => e.name && e.name.toLowerCase().includes(empName.toLowerCase())) || null;
        }
        return null;
    },

    // Finds multiple employees by name
    searchEmployeesByName: (empNameQuery) => {
        if (!empNameQuery) return [];
        return Object.values(biometricData).filter((emp) =>
            emp.name.trim().toLowerCase().includes(empNameQuery.trim().toLowerCase())
        );
    },

    // Gets a list of available biometric files from the uploads directory
    getBiometricFiles: () => {
        const uploadsDir = path.join(__dirname, "../uploads");
        try {
            if (fs.existsSync(uploadsDir)) {
                return fs.readdirSync(uploadsDir).filter((file) => file.endsWith(".xlsx") || file.endsWith(".xls"));
            }
        } catch (err) {
            console.error("Could not read uploads directory:", err);
        }
        return [];
    },

    // Re-calculates metrics for a given employee and month
    calculateEmployeeTimeSummary: (employee, month) => {
        const timeSummary = [];
        if (!employee || !month) return { timeSummary, daysInMonth: 0 };

        const { timeCard } = employee;
        const [yearStr, monthStr] = month.split("-");
		const year = parseInt(yearStr);
		const monthNum = parseInt(monthStr);
		const daysInMonth = new Date(year, monthNum, 0).getDate();

        const dailyRecordMap = new Map();
        timeCard.forEach((record) => {
			dailyRecordMap.set(record.day, record);
		});

        for (let day = 1; day <= daysInMonth; day++) {
            const rawRecord = dailyRecordMap.get(day) || {};
            const dailyMetrics = calculateDailyMetricsFromTimes(rawRecord, year, monthNum, day);
            timeSummary.push({
                day,
                dateWeekday: `${String(day).padStart(2, "0")} ${new Date(year,monthNum - 1,day).toLocaleString("en-US", {weekday: "short"})}`,
                ...dailyMetrics
            });
        }

        return { timeSummary, daysInMonth };
    },

    // You can add more specific helper functions here as needed
};
