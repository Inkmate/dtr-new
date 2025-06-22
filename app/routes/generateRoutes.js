// generateRoutes.js
const express = require("express");
const router = express.Router();
const path = require("path");

// Import shared utility functions and data
const {
	calculateDailyMetricsFromTimes,
	getEmployeeDataFromDb,
} = require("./utils");

/**
 * Converts a 24-hour time string (e.g., "08:00" or "14:30") to 12-hour format (e.g., "8:00 AM" or "2:30 PM").
 * Handles empty or invalid inputs gracefully.
 * @param {string} time24hr - The time string in "HH:MM" format.
 * @returns {string} The formatted 12-hour time string, or an empty string if input is invalid/empty.
 */
function convertTo12HourFormat(time24hr) {
	if (!time24hr || typeof time24hr !== "string" || time24hr.trim() === "") {
		return ""; // Return empty for empty, null, or non-string inputs
	}

	try {
		const [hours, minutes] = time24hr.split(":").map(Number);

		// Basic validation for numbers
		if (
			isNaN(hours) ||
			isNaN(minutes) ||
			hours < 0 ||
			hours > 23 ||
			minutes < 0 ||
			minutes > 59
		) {
			return time24hr; // Return original if not a valid time string (e.g., "Invalid Time")
		}

		const period = hours >= 12 ? "" : "";
		const displayHours = hours % 12 || 12; // Convert 0 (midnight) to 12 for 12 AM

		// Ensure minutes are always two digits
		const displayMinutes = String(minutes).padStart(2, "0");

		return `${displayHours}:${displayMinutes} ${period}`;
	} catch (error) {
		console.error("Error converting time '", time24hr, "':", error);
		return time24hr; // Return original in case of unexpected error during conversion
	}
}

// POST /generate - Handle single DTR generation (now duplicates)
router.post("/generate", async (req, res) => {
	const {empId, empName, month} = req.body;

	if (!empId || !empName || !month) {
		// Redirect back to form with an error or show a toast
		return res.redirect(
			`/form?error=Missing employee data for DTR generation.`
		);
	}

	// Fetch the employee's data from the database using the employee ID
	const searchResults = await getEmployeeDataFromDb(empName); // Assuming empName is unique enough or you use empId for more precise lookup
	const employeeData = searchResults.find((emp) => emp.empId === empId);

	if (!employeeData) {
		return res.redirect(`/form?error=Employee not found for DTR generation.`);
	}

	const [yearStr, monthNumStr] = month.split("-");
	const year = parseInt(yearStr);
	const monthNum = parseInt(monthNumStr);
	const daysInMonth = new Date(year, monthNum, 0).getDate();

	let timeSummary = [];
	let totalUndertimeMinutes = 0;
	let totalUndertimeHours = 0;

	for (let i = 1; i <= daysInMonth; i++) {
		const am_arrival = req.body[`am_arrival_${i}`] || "";
		const am_departure = req.body[`am_departure_${i}`] || "";
		const pm_arrival = req.body[`pm_arrival_${i}`] || "";
		const pm_departure = req.body[`pm_departure_${i}`] || "";
		const ot_arrival = req.body[`ot_arrival_${i}`] || ""; // Assuming OT fields if they exist
		const ot_departure = req.body[`ot_departure_${i}`] || ""; // Assuming OT fields if they exist

		// Crucial: Get undertime directly from the submitted hidden input fields
		let undertime_hours_submitted = req.body[`undertime_hours_${i}`];
		let undertime_minutes_submitted = req.body[`undertime_minutes_${i}`];

		const date = new Date(year, monthNum - 1, i);
		const dayOfWeek = date.getDay(); // 0 = Sunday, 6 = Saturday
		const dateWeekday = `${String(i).padStart(2, "0")} ${date.toLocaleString(
			"en-US",
			{weekday: "short"}
		)}`;

		let displayUndertimeHours;
		let displayUndertimeMinutes;

		// If it's a weekend, force undertime to be blank for display
		// Allow "Saturday" or "Sunday" as string labels from the form
		if (dayOfWeek === 0 || dayOfWeek === 6) {
			displayUndertimeHours = undertime_hours_submitted;
			displayUndertimeMinutes = undertime_minutes_submitted;
		} else {
			// If it's a weekday, first try submitted values.
			// If submitted values are empty, then recalculate.
			if (
				undertime_hours_submitted === "" ||
				undertime_minutes_submitted === ""
			) {
				const recalculatedMetrics = calculateDailyMetricsFromTimes(
					am_arrival,
					am_departure,
					pm_arrival,
					pm_departure,
					ot_arrival,
					ot_departure
				);
				displayUndertimeHours = recalculatedMetrics.undertime_hours;
				displayUndertimeMinutes = recalculatedMetrics.undertime_minutes;
			} else {
				displayUndertimeHours = undertime_hours_submitted;
				displayUndertimeMinutes = undertime_minutes_submitted;
			}
		}

		// Final formatting: Ensure values are empty string if they are empty or non-numeric (except for 0)
		// For numbers, parse them.
		const isWeekendLabel = (val) => val === "Saturday" || val === "Sunday";

		displayUndertimeHours = isWeekendLabel(displayUndertimeHours)
			? displayUndertimeHours
			: displayUndertimeHours === "" || isNaN(parseInt(displayUndertimeHours))
			? ""
			: parseInt(displayUndertimeHours);

		displayUndertimeMinutes = isWeekendLabel(displayUndertimeMinutes)
			? displayUndertimeMinutes
			: displayUndertimeMinutes === "" ||
			  isNaN(parseInt(displayUndertimeMinutes))
			? ""
			: String(parseInt(displayUndertimeMinutes) || 0).padStart(2, "0");

		// For total calculation, always treat empty or non-numeric as 0
		const parsedForTotalHours = parseInt(displayUndertimeHours) || 0;
		const parsedForTotalMinutes = parseInt(displayUndertimeMinutes) || 0;

		totalUndertimeHours += parsedForTotalHours;
		totalUndertimeMinutes += parsedForTotalMinutes;

		timeSummary.push({
			day: i,
			dateWeekday,
            am_arrival_12hr: convertTo12HourFormat(am_arrival),
            am_departure_12hr: convertTo12HourFormat(am_departure),
            pm_arrival_12hr: convertTo12HourFormat(pm_arrival),
            pm_departure_12hr: convertTo12HourFormat(pm_departure),
            ot_arrival_12hr: convertTo12HourFormat(ot_arrival),
            ot_departure_12hr: convertTo12HourFormat(ot_departure),
			// Assuming these are also derived from calculateDailyMetricsFromTimes if needed
			total_hours: "",
			late: "",
			early_out: "",
			overtime: "",
			undertime_hours: displayUndertimeHours, // Use the display value
			undertime_minutes: displayUndertimeMinutes, // Use the display value
			remarks: "", // No remarks in form yet
		});
	}

	// Final adjustment for total undertime minutes
	totalUndertimeHours += Math.floor(totalUndertimeMinutes / 60);
	totalUndertimeMinutes = totalUndertimeMinutes % 60;

	// Construct the data for a single DTR instance
	const singleDtrData = {
		empId,
		empName,
		month,
		timeSummary, // Array of daily records for this DTR
		totalUndertimeHours,
		totalUndertimeMinutes: String(totalUndertimeMinutes).padStart(2, "0"), // Format minutes
	};

	// Duplicate the DTR data for display
	const dtrsToDisplay = [
		singleDtrData,
		singleDtrData, // Duplicate the data
	];

	res.render("display", {
		dtrs: dtrsToDisplay, // Pass the array of DTR objects
	});
});

// POST /generate-batch - Handle combined DTR generation
router.post("/generate-batch", async (req, res) => {
	const selectedDTRs = JSON.parse(req.body.dtrs); // Array of { empId, empName, month }

	const combinedDTRsData = [];

	for (const dtrInfo of selectedDTRs) {
		const {empId, empName, month} = dtrInfo;

		// Fetch the employee's data from the database
		const searchResults = await getEmployeeDataFromDb(empName);
		const employeeData = searchResults.find((emp) => emp.empId === empId);

		if (!employeeData || !employeeData.formData[month]) {
			console.warn(`Skipping DTR for ${empName} (${month}): data not found.`);
			continue; // Skip to the next DTR if data is missing
		}

		const [yearStr, monthNumStr] = month.split("-");
		const year = parseInt(yearStr);
		const monthNum = parseInt(monthNumStr);
		const daysInMonth = new Date(year, monthNum, 0).getDate();

		let timeSummary = [];
		let totalUndertimeMinutes = 0;
		let totalUndertimeHours = 0;

		const currentMonthData = employeeData.formData[month];

		for (let i = 1; i <= daysInMonth; i++) {
			const dailyRecord = currentMonthData[i] || {}; // Use stored data
			const am_arrival = dailyRecord.am_arrival || "";
			const am_departure = dailyRecord.am_departure || "";
			const pm_arrival = dailyRecord.pm_arrival || "";
			const pm_departure = dailyRecord.pm_departure || "";
			const ot_arrival = dailyRecord.ot_arrival || "";
			const ot_departure = dailyRecord.ot_departure || "";

			const date = new Date(year, monthNum - 1, i);
			const dayOfWeek = date.getDay();
			const dateWeekday = `${String(i).padStart(2, "0")} ${date.toLocaleString(
				"en-US",
				{weekday: "short"}
			)}`;

			let displayUndertimeHours;
			let displayUndertimeMinutes;

			// If it's a weekend, force undertime to be blank for display
			// Allow "Saturday" or "Sunday" as string labels from the form
			// ✅ This works correctly in batch:
			if (dayOfWeek === 0 || dayOfWeek === 6) {
				displayUndertimeHours = dailyRecord.undertime_hours;
				displayUndertimeMinutes = dailyRecord.undertime_minutes;
			} else {
				// For batch, undertime should typically come from stored data.
				// If stored undertime is empty for a weekday, recalculate it for the batch
				if (
					(dailyRecord.undertime_hours === "" ||
						dailyRecord.undertime_minutes === "") &&
					dayOfWeek !== 0 &&
					dayOfWeek !== 6
				) {
					const recalculatedMetrics = calculateDailyMetricsFromTimes(
						am_arrival,
						am_departure,
						pm_arrival,
						pm_departure,
						ot_arrival,
						ot_departure
					);
					displayUndertimeHours = recalculatedMetrics.undertime_hours;
					displayUndertimeMinutes = recalculatedMetrics.undertime_minutes;
				} else {
					displayUndertimeHours = dailyRecord.undertime_hours;
					displayUndertimeMinutes = dailyRecord.undertime_minutes;
				}
			}

			// Ensure values for display are formatted correctly (empty if non-numeric/empty, or parsed number)
			const isWeekendLabel = (val) => val === "Saturday" || val === "Sunday";

			displayUndertimeHours = isWeekendLabel(displayUndertimeHours)
				? displayUndertimeHours
				: displayUndertimeHours === "" || isNaN(parseInt(displayUndertimeHours))
				? ""
				: parseInt(displayUndertimeHours);

			displayUndertimeMinutes = isWeekendLabel(displayUndertimeMinutes)
				? displayUndertimeMinutes
				: displayUndertimeMinutes === "" ||
				  isNaN(parseInt(displayUndertimeMinutes))
				? ""
				: String(parseInt(displayUndertimeMinutes) || 0).padStart(2, "0");

			// For total calculation, always treat empty or non-numeric as 0
			const parsedForTotalHours = parseInt(displayUndertimeHours) || 0;
			const parsedForTotalMinutes = parseInt(displayUndertimeMinutes) || 0;

			totalUndertimeHours += parsedForTotalHours;
			totalUndertimeMinutes += parsedForTotalMinutes;

			timeSummary.push({
				day: i,
				dateWeekday,
				am_arrival_12hr: convertTo12HourFormat(am_arrival),
				am_departure_12hr: convertTo12HourFormat(am_departure),
				pm_arrival_12hr: convertTo12HourFormat(pm_arrival),
				pm_departure_12hr: convertTo12HourFormat(pm_departure),
				ot_arrival_12hr: convertTo12HourFormat(ot_arrival),
				ot_departure_12hr: convertTo12HourFormat(ot_departure),
				// Assuming these are also derived from calculateDailyMetricsFromTimes if needed
				total_hours: "",
				late: "",
				early_out: "",
				overtime: "",
				undertime_hours: displayUndertimeHours,
				undertime_minutes: displayUndertimeMinutes,
				remarks: dailyRecord.remarks || "",
			});
		}

		totalUndertimeHours += Math.floor(totalUndertimeMinutes / 60);
		totalUndertimeMinutes = totalUndertimeMinutes % 60;

		const processedDtr = {
			empId,
			empName,
			month,
			timeSummary,
			totalUndertimeHours,
			totalUndertimeMinutes: String(totalUndertimeMinutes).padStart(2, "0"),
		};
		// Duplicate each processed DTR
		combinedDTRsData.push(processedDtr);
		combinedDTRsData.push(processedDtr); // Duplicate
	}

	res.render("display", {
		dtrs: combinedDTRsData, // Pass the array of DTR objects for batch
	});
});

module.exports = router;
