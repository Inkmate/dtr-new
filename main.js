const { app, BrowserWindow } = require('electron');
const path = require('path');

// Import your Express application instance
// Assuming your main Express app file is named 'app.js' and is inside a folder named 'app'
// CORRECTED PATH: Changed from './app' to './app/app'
const expressApp = require('./app/app'); // This will load your app/app.js

function createWindow() {
    const mainWindow = new BrowserWindow({
        width: 1200, // Adjust initial window width as needed
        height: 900, // Adjust initial window height as needed
        webPreferences: {
            // IMPORTANT: Be cautious with these settings in a production environment.
            // For a simple internal tool, they might be acceptable, but they can pose security risks.
            nodeIntegration: true, // Allows Node.js APIs in the renderer process
            contextIsolation: false, // Disables context isolation (less secure)
            preload: path.join(__dirname, 'preload.js') // Optional: A more secure way to expose Node.js APIs
        },
    });

    // Start the Express server on a specific port
    // You can choose any available port. Consider making this configurable if needed.
    const PORT = 3000; 
    let server = null;

    // Ensure the server starts only once
    if (!server) {
        server = expressApp.listen(PORT, () => {
            console.log(`Express app listening on http://localhost:${PORT}`);
            // Load the Express app's URL in the Electron window
            mainWindow.loadURL(`http://localhost:${PORT}`);
        });

        server.on('error', (err) => {
            console.error('Failed to start Express server:', err);
            // Handle error, e.g., show a message to the user and quit
            app.quit();
        });
    }

    // Optional: Open DevTools for debugging during development
    // mainWindow.webContents.openDevTools();
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        // On macOS it's common to re-create a window in the app when the
        // dock icon is clicked and there are no other windows open.
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

// Quit when all windows are closed, except on macOS. There's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

