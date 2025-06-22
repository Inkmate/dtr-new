// app.js (or server.js)
const express = require("express");
const path = require("path");
const bodyParser = require("body-parser"); // For parsing POST request bodies
const fs = require('fs'); // Import fs to check for existence of error view files

// Import the route files
const indexRoutes = require("./routes/indexRoutes");
const uploadRoutes = require("./routes/uploadRoutes");
const generateRoutes = require("./routes/generateRoutes");

const app = express();
// The port should now be managed by main.js or a separate config, not here.
// const port = process.env.PORT || 3000; // REMOVED from here

// Set EJS as the view engine
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views")); // Assuming your views are in a 'views' directory

// Middleware to parse URL-encoded bodies (for form data)
app.use(bodyParser.urlencoded({
    extended: true
}));
// Middleware to parse JSON bodies (for API requests if any)
app.use(bodyParser.json());

// Serve static files (e.g., CSS, JS, images)
app.use(express.static(path.join(__dirname, "public")));
app.use("/uploads", express.static(path.join(__dirname, "uploads"))); // Serve uploaded files if needed

// Mount the route files
app.use("/", indexRoutes);
app.use("/", uploadRoutes);
app.use("/", generateRoutes);

// Catch 404 and forward to error handler
app.use((req, res, next) => {
    const err = new Error('Not Found');
    err.statusCode = 404; // Set status code for 404 errors
    next(err);
});

// Centralized Error Handling Middleware
app.use((err, req, res, next) => {
    // Log the error for debugging purposes (in development)
    console.error(err.stack);

    const statusCode = err.statusCode || 500; // Default to 500 if no specific status code is set
    res.status(statusCode);

    let viewName = 'error'; // Default error view
    const specificErrorViewPath = path.join(__dirname, 'views', `${statusCode}.ejs`);

    // Check if a specific EJS file for the status code exists
    if (fs.existsSync(specificErrorViewPath)) {
        viewName = String(statusCode); // Use the status code as the view name (e.g., '404', '500')
    } else {
        // Fallback to generic error.ejs if specific one doesn't exist
        console.warn(`Specific error view for status ${statusCode}.ejs not found. Falling back to error.ejs.`);
    }

    // Render the appropriate error view
    res.render(viewName, {
        message: err.message || `An unexpected error occurred (Status: ${statusCode}).`,
        error: process.env.NODE_ENV === "development" ? err : {}, // Only show error details in development
    });
});


// EXPORT the Express app instance so main.js can start it.
module.exports = app; // ADDED this line
// REMOVED app.listen() from here, as main.js will now call it.
