// uploadRoutes.js
const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs"); // Import fs for existsSync
const os = require("os");

const router = express.Router();

// Import shared utility functions and data
const {loadMultiSheetEmployeeAttendance} = require("./utils");

const uploadsDir = path.join(process.cwd(), "uploads"); // Use working dir

// Ensure the uploads directory exists
if (!fs.existsSync(uploadsDir)) {
	fs.mkdirSync(uploadsDir, {recursive: true});
}

const storage = multer.diskStorage({
	destination: (req, file, cb) => {
		cb(null, uploadsDir);
	},
	filename: (req, file, cb) => {
		cb(null, Date.now() + "-" + file.originalname);
	},
});

const upload = multer({
	storage,
});

// POST /upload - Handle file uploads
router.post("/upload", upload.single("excelFile"), (req, res, next) => {
	if (!req.file) {
		res.status(400);
		return next(
			new Error("No file uploaded. Please select an Excel file to upload.")
		);
	}
	loadMultiSheetEmployeeAttendance(req.file.path);
	res.redirect("/form");
});

module.exports = router;
