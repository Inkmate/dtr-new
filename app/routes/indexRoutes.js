// indexRoutes.js
const express = require("express");
const router = express.Router();
const path = require("path");
const fs = require("fs");

// Import shared utility functions and data
const {
	loadMultiSheetEmployeeAttendance,
	calculateDailyMetricsFromTimes,
	getEmployeeDataFromDb, // Import the new DB retrieval function
} = require("./utils");

// Dynamically find and load all Excel files at startup
const uploadsDir = path.join(__dirname, "../uploads");

try {
	if (fs.existsSync(uploadsDir)) {
		const biometricFiles = fs
			.readdirSync(uploadsDir)
			.filter((file) => file.endsWith(".xlsx") || file.endsWith(".xls"));

		if (biometricFiles.length > 0) {
			console.log(
				`Found ${biometricFiles.length} Excel file(s) in uploads/ folder. Loading all of them.`
			);
			biometricFiles.forEach((file) => {
				const excelFilePath = path.join(uploadsDir, file);
				console.log(`Loading attendance data from: ${excelFilePath}`);
				loadMultiSheetEmployeeAttendance(excelFilePath);
			});
		} else {
			console.warn(
				"No Excel file (.xlsx or .xls) found in the uploads/ folder. Please upload an attendance file."
			);
		}
	} else {
		console.warn(
			"The uploads/ directory does not exist. Please create it and upload an attendance file."
		);
		fs.mkdirSync(uploadsDir, {recursive: true}); // Attempt to create the directory
	}
} catch (err) {
	console.error("Error during initial Excel file load:", err);
}

router.get("/", (req, res) => {
	res.render("home"); // Create home.ejs
});

// GET / - Main form route
router.get("/form", async (req, res) => {
	let empId = req.query.empId || "";
	let empName = req.query.empName || "";
	let month = req.query.month || "";

	let foundEmployeeData = null;
	let allEmployeesForDropdown = [];

	// Fetch all employees and their months for the dropdown.
	const allDbEmployees = await getEmployeeDataFromDb();
	if (Array.isArray(allDbEmployees)) {
		allDbEmployees.forEach((emp) => {
			if (emp.months && Array.isArray(emp.months)) {
				emp.months.forEach((m) => {
					allEmployeesForDropdown.push({
						userId: emp.empId,
						name: emp.empName,
						display: `${emp.empName}-${m}`, // Concatenated string for datalist option value
					});
				});
			} else {
				allEmployeesForDropdown.push({
					userId: emp.empId,
					name: emp.empName,
					month: "",
					display: emp.empName,
				});
			}
		});
	} else {
		console.error(
			"getEmployeeDataFromDb did not return an array for all employees during dropdown load:",
			allDbEmployees
		);
		allEmployeesForDropdown = [];
	}

	// --- NEW LOG HERE: Check if searchQuery exists at all ---
	console.log(
		`DEBUG indexRoutes: Raw req.query.searchQuery: '${req.query.searchQuery}'`
	);

	// Logic to determine which employee and month to load based on query parameters
	if (req.query.searchQuery) {
		const searchQuery = req.query.searchQuery.trim();
		const parts = searchQuery.split("-");

		// --- KEEP ALL THESE DEBUG LOGS ---
		console.log(`DEBUG indexRoutes: searchQuery (trimmed): '${searchQuery}'`);
		console.log(`DEBUG indexRoutes: parts:`, parts);
		console.log(`DEBUG indexRoutes: parts.length: ${parts.length}`);

		// --- REFINED: Add .trim() to individual parts for robustness ---
		const potentialMonthYear = parts[parts.length - 2]?.trim();
		const potentialMonthNum = parts[parts.length - 1]?.trim();

		console.log(
			`DEBUG indexRoutes: potentialMonthYear (trimmed): '${potentialMonthYear}'`
		);
		console.log(
			`DEBUG indexRoutes: potentialMonthNum (trimmed): '${potentialMonthNum}'`
		);

		const isMonthFormat =
			potentialMonthYear &&
			potentialMonthNum &&
			`${potentialMonthYear}-${potentialMonthNum}`.match(/^\d{4}-\d{2}$/);
		console.log(`DEBUG indexRoutes: isMonthFormat: ${isMonthFormat}`);
		// --- END DEBUG LOGS ---

		let queriedName = "";
		let queriedMonth = "";
		let specificSearchAttempted = false;

		if (parts.length >= 3 && isMonthFormat) {
			queriedMonth = `${potentialMonthYear}-${potentialMonthNum}`;
			queriedName = parts
				.slice(0, parts.length - 2)
				.join("-")
				.trim();
			specificSearchAttempted = true;

			const searchResults = await getEmployeeDataFromDb(
				queriedName,
				queriedMonth
			);
			if (searchResults.length > 0) {
				foundEmployeeData = searchResults[0];
				empId = foundEmployeeData.empId;
				empName = foundEmployeeData.empName;
				month = queriedMonth;
			}
		}

		if (!foundEmployeeData) {
			queriedName = specificSearchAttempted ? queriedName : searchQuery;
			const nameOnlySearchResults = await getEmployeeDataFromDb(queriedName);
			if (nameOnlySearchResults.length > 0) {
				foundEmployeeData = nameOnlySearchResults[0];
				empId = foundEmployeeData.empId;
				empName = foundEmployeeData.empName;
				month = foundEmployeeData.months[0] || "";
			}
		}
	} else if (empId || empName) {
		let searchResults = [];
		if (empId) {
			searchResults = await getEmployeeDataFromDb("", "");
			foundEmployeeData = searchResults.find((emp) => emp.empId === empId);
		} else if (empName) {
			searchResults = await getEmployeeDataFromDb(empName);
			if (searchResults.length > 0) {
				foundEmployeeData = searchResults[0];
			}
		}

		if (foundEmployeeData) {
			empId = foundEmployeeData.empId;
			empName = foundEmployeeData.empName;
			month = req.query.month || foundEmployeeData.months[0] || "";
		}
	}

	if (!foundEmployeeData) {
		empId = "";
		empName = req.query.searchQuery || req.query.empName || "";
		month = "";
	}

	const selectedFile = req.query.biometricFile || "";
	let formData = {};
	let timeSummary = [];
	let daysInMonth = 0;

	if (foundEmployeeData && month) {
		const currentMonthData =
			foundEmployeeData.formData && foundEmployeeData.formData[month];
		if (currentMonthData) {
			const [yearStr, monthNumStr] = month.split("-");
			const year = parseInt(yearStr);
			const monthNum = parseInt(monthNumStr);
			daysInMonth = new Date(year, monthNum, 0).getDate();

			for (let day = 1; day <= daysInMonth; day++) {
				const dailyRecord = currentMonthData[day] || {
					day: day,
					dateWeekday: `${String(day).padStart(2, "0")} ${new Date(
						year,
						monthNum - 1,
						day
					).toLocaleString("en-US", {weekday: "short"})}`,
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
					undertime_hours: "",
					undertime_minutes: "",
				};
				formData[day] = dailyRecord;
				timeSummary.push(dailyRecord);
			}
		}
	}

	let biometricFiles = [];
	try {
		if (fs.existsSync(uploadsDir)) {
			biometricFiles = fs
				.readdirSync(uploadsDir)
				.filter((file) => file.endsWith(".xlsx") || file.endsWith(".xls"));
		}
	} catch (err) {
		console.error("Error reading uploads directory for form rendering:", err);
	}

	res.render("form", {
		empId: empId,
		empName: empName,
		month: month,
		formData,
		biometricFiles,
		selectedFile,
		timeSummary,
		daysInMonth,
		allEmployees: allEmployeesForDropdown,
	});
});

// GET /search-employee - This route is now primarily for AJAX calls from other parts if needed.
// The main form submission to / will handle the primary search.
router.get("/search-employee", async (req, res, next) => {
	const empNameQuery = req.query.empName;
	console.log(`[Search] Received empNameQuery: '${empNameQuery}'`);

	if (!empNameQuery) {
		return res.status(400).json({
			status: "error",
			message: "Employee name is required for search.",
		});
	}

	const searchResults = await getEmployeeDataFromDb(empNameQuery);

	if (searchResults.length === 0) {
		console.warn(`[Search] No employees found matching '${empNameQuery}'.`);
		return res.status(404).json({
			status: "error",
			message: `No employees found matching '${empNameQuery}'.`,
		});
	}

	console.log(`[Search] Found ${searchResults.length} matches.`);
	return res.json({
		status: "success",
		results: searchResults,
		count: searchResults.length,
	});
});

module.exports = router;
