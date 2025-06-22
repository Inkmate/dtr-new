// app/routes/biometricParserSimple.js
const xlsx = require("xlsx");

/**
 * Parses a date string in DD/MM/YYYY format into a JavaScript Date object.
 * This is needed because `new Date()` can be inconsistent with "DD/MM/YYYY".
 * @param {string} dateString - The date string (e.g., "13/06/2025").
 * @returns {Date|null} - A Date object if successful, null otherwise.
 */
function parseDdMmYyyyDateString(dateString) {
    const parts = dateString.split('/');
    if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10); // Month is 1-based
        const year = parseInt(parts[2], 10);

        // Basic validation for numbers and reasonable ranges
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) &&
            month >= 1 && month <= 12 &&
            day >= 1 && day <= 31 && // Broad day check, specific month days handled by Date constructor
            year >= 1900 && year <= 2100) // Reasonable year range
        {
            // Note: Month in Date constructor is 0-based
            const date = new Date(year, month - 1, day);
            // Verify if the constructed date is valid and matches input to prevent invalid dates like Feb 30
            if (date.getFullYear() === year && (date.getMonth() + 1) === month && date.getDate() === day) {
                return date;
            }
        }
    }
    return null; // Return null for invalid date strings
}


/**
 * Helper function to calculate daily metrics (adapted for direct times)
 * This is a copy of calculateDailyMetricsFromTimes from utils.js,
 * necessary because parseSimpleBiometric needs it and should be self-contained.
 * If you ever update calculateDailyMetricsFromTimes in utils.js, remember to sync this.
 * @param {object} record - The daily time record object.
 * @param {number} year - The year of the record.
 * @param {number} monthNum - The 1-based month number.
 * @param {number} day - The day of the month.
 * @returns {object} - Calculated daily metrics.
 */
function calculateDailyMetricsFromTimesForSimpleParser(record, year, monthNum, day) {
    const amArrival = record.am_arrival;
    const amDeparture = record.am_departure;
    const pmArrival = record.pm_arrival;
    const pmDeparture = record.pm_departure;
    const otArrival = record.ot_arrival;
    const otDeparture = record.ot_departure;

    // These fields are not typically in the raw simple biometric data, so default them
    // unless you expect them to be provided in some form (e.g., calculated by the excel itself)
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
    const amOutTime = amDeparture ?
        new Date(`${dateString} ${amDeparture}`) :
        null;
    const pmInTime = pmArrival ? new Date(`${dateString} ${pmArrival}`) : null;
    const pmOutTime = pmDeparture ?
        new Date(`${dateString} ${pmDeparture}`) :
        null;
    const otInTime = otArrival ? new Date(`${dateString} ${otArrival}`) : null;
    const otOutTime = otDeparture ?
        new Date(`${dateString} ${otDeparture}`) :
        null;

    let morningDuration = 0;
    if (
        amInTime &&
        amOutTime &&
        !isNaN(amInTime) &&
        !isNaN(amOutTime) &&
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

    if (
        (undertime_hours === "" || undertime_minutes === "") &&
        dayOfWeek !== 0 &&
        dayOfWeek !== 6
    ) {
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
                !isNaN(pmInTime) &&
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
            undertime_hours = "";
            undertime_minutes = "";
        }
    } else if (dayOfWeek === 0 || dayOfWeek === 6) {
        undertime_hours = "";
        undertime_minutes = "";
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
 * Parses a simple biometric Excel file with flat structure:
 * Name | Date | Timetable (AM/PM) | Clock In | Clock Out
 * @param {string} filePath - Path to the Excel file
 * @returns {Array<object>} Parsed employee attendance data
 */
function parseSimpleBiometric(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]]; // Assume data is on the first sheet
    
    // Add raw: false to ensure values are not raw numbers for dates/times.
    // Add cellDates: true to get proper JS Date objects if Excel stores them as dates.
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: "", raw: false, cellDates: true });

    const employeeMap = new Map();

    for (const row of rows) {
        console.log("DEBUG: Processing row:", JSON.stringify(row)); // Crucial debug log

        const name = (row["Name"] || "").trim();
        let dateValue = row["Date"]; // Get raw value, could be string or Date object
        const timetable = (row["Timetable"] || "").trim().toUpperCase();
        const clockIn = (row["Clock In"] || "").trim();
        const clockOut = (row["Clock Out"] || "").trim();

        if (!name || !dateValue || !timetable) {
            console.warn(`Skipping row due to missing essential data: Name='${name}', Date='${dateValue}', Timetable='${timetable}'`);
            continue;
        }

        let date;
        let dateStrForLogging = String(dateValue).trim(); // For initial logging

        if (dateValue instanceof Date) {
            // If cellDates: true parsed it as a Date object directly
            date = dateValue;
            dateStrForLogging = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
            console.log(`DEBUG: Date was already a Date object: ${dateStrForLogging}`);
        } else {
            // It's a string, try custom DD/MM/YYYY parser first
            date = parseDdMmYyyyDateString(dateStrForLogging);
            if (!date) {
                // If custom DD/MM/YYYY parser failed, try generic Date constructor as a fallback
                date = new Date(dateStrForLogging);
                if (isNaN(date.getTime())) { // Check if new Date() resulted in "Invalid Date"
                    console.warn(`Skipping row for '${name}' due to invalid date format: '${dateStrForLogging}'`);
                    continue;
                }
                console.log(`DEBUG: Date parsed by generic new Date(): ${dateStrForLogging}`);
            } else {
                 console.log(`DEBUG: Date parsed by parseDdMmYyyyDateString: ${dateStrForLogging}`);
            }
        }
        
        const year = date.getFullYear();
        const monthNum = date.getMonth() + 1; // 1-based month
        const month = String(monthNum).padStart(2, "0");
        const day = date.getDate();
        const key = `${name}-${year}-${month}`;

        if (!employeeMap.has(key)) {
            // Initialize timeCard with enough slots for the max days in any month (31)
            const initialTimeCard = Array(31).fill(null).map((_, idx) => {
                // Determine weekday for each day
                const currentDayDate = new Date(year, monthNum - 1, idx + 1);
                const dateWeekday = `${String(idx + 1).padStart(2, "0")} ${currentDayDate.toLocaleString("en-US", { weekday: "short" })}`;
                return {
                    day: idx + 1,
                    dateWeekday: dateWeekday,
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
            });

            employeeMap.set(key, {
                userId: name, // Using name as userId for now; adjust if actual ID exists
                name,
                department: "", // New format doesn't provide department, keep empty
                month: `${year}-${month}`,
                attendanceDateRange: "", // New format doesn't provide global range, keep empty
                tablingDate: "", // New format doesn't provide tabling date, keep empty
                timeCard: initialTimeCard
            });
        }

        const employeeData = employeeMap.get(key);
        // Ensure the day is within valid bounds for the timeCard array
        if (day >= 1 && day <= employeeData.timeCard.length) {
            const record = employeeData.timeCard[day - 1];

            if (timetable === "AM") {
                record.am_arrival = clockIn;
                record.am_departure = clockOut;
            } else if (timetable === "PM") {
                record.pm_arrival = clockIn;
                record.pm_departure = clockOut;
            }
            // Add other timetables (e.g., "OT") if the new format supports it
            // else if (timetable === "OT") {
            //     record.ot_arrival = clockIn;
            //     record.ot_departure = clockOut;
            // }
        } else {
            console.warn(`Day ${day} out of bounds for month ${monthNum} in record for ${name}. Skipping.`);
        }
    }

    // Convert Map values to Array, apply daily metrics calculation to each record
    return Array.from(employeeMap.values()).map(employee => {
        const [yearStr, monthStr] = (employee.month || '').split('-');
        const year = parseInt(yearStr);
        const monthNum = parseInt(monthStr);

        // Filter out nulls and apply calculations
        employee.timeCard = employee.timeCard
            .filter(record => record !== null) // Keep this filter
            .map(record => {
                // Ensure day and dateWeekday are preserved
                const day = record.day;
                const dateWeekday = record.dateWeekday;

                const calculatedMetrics = calculateDailyMetricsFromTimesForSimpleParser(record, year, monthNum, day);
                // Merge original record properties with the newly calculated metrics
                return { ...record, ...calculatedMetrics }; 
            });
        return employee;
    });
}
                                                                                                                                                                                                                                                                                                                                                                                   

/**
 * Helper function to determine if the Excel sheet matches the new simple biometric format.
 * @param {object} sheet - The xlsx sheet object.
 * @returns {boolean} - True if it matches the simple format, false otherwise.
 */
function isSimpleBiometricFormat(sheet) {
    // Attempt to read the first row as headers
    // Using raw: false, cellDates: true for header detection consistency
    const rawHeaders = xlsx.utils.sheet_to_json(sheet, { header: 1, range: "A1:Z1", raw: false, cellDates: true });
    
    if (!rawHeaders || rawHeaders.length === 0 || !Array.isArray(rawHeaders[0])) {
        return false; // No headers found or not in expected format
    }
    
    const headers = rawHeaders[0].map(h => String(h || '').trim().toLowerCase());
    
    const requiredSimpleHeaders = ["name", "date", "timetable", "clock in", "clock out"];
    
    // Check if all required headers are present
    return requiredSimpleHeaders.every(reqHeader => headers.includes(reqHeader));
}

module.exports = {
    parseSimpleBiometric,
    isSimpleBiometricFormat,
    parseDdMmYyyyDateString // Export for testing if needed
};
