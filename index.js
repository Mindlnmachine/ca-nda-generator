require('dotenv').config();

const express = require('express');
const passport = require('passport');
const session = require('express-session');
const googleStrategy = require('passport-google-oauth20').Strategy;
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs').promises;

const app = express();
const port = process.env.PORT || 3000;
const EXCEL_FILE_PATH = path.join(__dirname, 'saved_texts.xlsx');
// NEW: Path to the worked_for.xlsx file
const WORKED_FOR_EXCEL_PATH = path.join(__dirname, 'worked_for.xlsx');

app.use(express.json());

app.use(session({
  secret: process.env.SESSION_SECRET || 'your_secret_key',
  resave: false,
  saveUninitialized: true,
  cookie: { secure: false }
}));

app.use(passport.initialize());
app.use(passport.session());

passport.use(new googleStrategy({
    clientID: process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
    callbackURL: "http://localhost:3000/auth/google/callback"
  },
  function(accessToken, refreshToken, profile, done) {
    return done(null, profile);
  }
));

passport.serializeUser(function(user, done) {
  done(null, user);
});
passport.deserializeUser(function(user, done) {
  done(null, user);
});
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    // If the user is already authenticated, redirect them to the profile page
    if (req.isAuthenticated()) {
        return res.redirect('/profile.html');
    }

    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>NDA Generator</title>
            <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
            <style>
                body {
                    font-family: 'Roboto', sans-serif;
                    margin: 0;
                    padding: 0;
                    background-color: #f0f2f5;
                    color: #333;
                    display: flex;
                    flex-direction: column;
                    min-height: 100vh;
                }
                .header-container {
                    background-color: #ffffff;
                    padding: 15px 50px;
                    border-bottom: 1px solid #eee;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                }
                .logo {
                    display: flex;
                    align-items: center;
                    gap: 10px; /* Space between brand, separator, and project name */
                }
                .company-brand {
                    font-size: 14px; /* Smaller font for "bharat valley" */
                    line-height: 1.1;
                    color: #666; /* Lighter color */
                    text-align: left;
                    font-weight: 500;
                }
                .company-brand .bharat-text {
                    color: #4285F4; /* A touch of color for 'bharat' */
                    font-weight: 700;
                }
                .separator {
                    font-size: 24px; /* Size of the vertical bar */
                    color: #ccc; /* Lighter separator color */
                    font-weight: 300;
                }
                .project-name {
                    font-size: 24px; /* Larger for the project name */
                    font-weight: 700;
                    color: #333; /* Darker color for prominence */
                }

                .nav-buttons {
                    display: flex;
                    gap: 15px;
                }
                .nav-buttons a {
                    text-decoration: none;
                    padding: 8px 18px;
                    border-radius: 5px;
                    font-weight: 600;
                    transition: all 0.2s ease;
                    white-space: nowrap; /* Prevent buttons from wrapping */
                }
                .nav-login-btn {
                    color: #4285F4;
                    border: 1px solid #4285F4;
                    background-color: transparent;
                }
                .nav-login-btn:hover {
                    background-color: #e0eaff;
                }
                .nav-get-started-btn {
                    background-color: #4285F4;
                    color: white;
                    border: 1px solid #4285F4;
                }
                .nav-get-started-btn:hover {
                    background-color: #357ae8;
                }
                .main-content {
                    flex-grow: 1;
                    display: flex;
                    flex-direction: column;
                    justify-content: center;
                    align-items: center;
                    text-align: center;
                    padding: 50px 20px;
                }
                .main-heading {
                    font-size: 48px;
                    font-weight: 700;
                    margin-bottom: 15px;
                    line-height: 1.2;
                    color: #333;
                }
                .main-heading .highlight {
                    color: #4285F4;
                }
                .sub-text {
                    font-size: 18px;
                    color: #666;
                    max-width: 600px;
                    margin-bottom: 40px;
                }
                .cta-buttons {
                    display: flex;
                    gap: 20px;
                }
                .cta-buttons a {
                    text-decoration: none;
                    padding: 15px 30px;
                    border-radius: 8px;
                    font-size: 18px;
                    font-weight: 600;
                    transition: all 0.2s ease;
                }
                .cta-primary-btn {
                    background-color: #4285F4;
                    color: white;
                    border: none;
                }
                .cta-primary-btn:hover {
                    background-color: #357ae8;
                }
                .cta-secondary-btn {
                    background-color: transparent;
                    color: #4285F4;
                    border: 1px solid #4285F4;
                }
                .cta-secondary-btn:hover {
                    background-color: #e0eaff;
                }
            </style>
        </head>
        <body>
            <div class="header-container">
                <div class="logo">
                    <div class="company-brand">
                        <span class="bharat-text">bharat</span><br>
                        <span>valley</span>
                    </div>
                    <span class="separator">|</span>
                    <span class="project-name">NDA Generator</span>
                </div>
                <nav class="nav-buttons">
                    <a href="/auth/google" class="nav-login-btn">Login</a>
                    <a href="/auth/google" class="nav-get-started-btn">Get Started</a>
                </nav>
            </div>

            <div class="main-content">
                <h1 class="main-heading">Automate your <span class="highlight">Non-Disclosure Agreements</span></h1>
                <p class="sub-text">Effortlessly generate, customize, and manage confidentiality agreements with integrated Google authentication and data export.</p>
                <div class="cta-buttons">
                    <a href="/auth/google" class="cta-primary-btn">Create Free Account</a>
                    <a href="/auth/google" class="cta-secondary-btn">Login</a>
                </div>
            </div>
        </body>
        </html>
    `);
});

app.get('/auth/google',
  passport.authenticate('google', { scope: ['profile', 'email'] })
);

app.get('/auth/google/callback',
  passport.authenticate('google', { failureRedirect: '/' }),
  (req, res) => {
    res.redirect('/profile.html');
  }
);

app.get('/profile', (req, res) => {
    if (!req.isAuthenticated()) {
        return res.redirect('/');
    }
    res.sendFile(path.join(__dirname, 'public', 'profile.html'));
});

app.get('/profile-data', (req, res) => {
    if (!req.isAuthenticated()) {
        return res.status(401).json({ message: 'Unauthorized' });
    }
    res.json({ user: {
        displayName: req.user.displayName,
        email: req.user.emails && req.user.emails[0] ? req.user.emails[0].value : 'N/A'
    }});
});

app.post('/api/save-text', async (req, res) => {
    if (!req.isAuthenticated()) {
        return res.status(401).json({ success: false, message: 'Unauthorized' });
    }

    // Accept both spaced headers (Excel-friendly) and camelCase payloads from the client
    const body = req.body || {};
    const userEmail = body['User Email'] || body.userEmail || (req.user?.emails?.[0]?.value ?? '');
    const date = body['Date'] || body.date;
    const companyNameA = body['Company Name A'] || body.companyNameA;
    const llpin = body['LLPIN'] || body.llpin;
    const addressA = body['Address A'] || body.addressA;
    const caFirmName = body['CA FIRM NAME'] || body.caFirmName;
    const caName = body['CA NAME'] || body.caName;
    const memberRegNo = body['MEMBER REG NO'] || body.memberRegNo;
    const partnerProprietor = body['PARTNER/PROPRIETOR'] || body.partnerProprietor;
    const companyNameB = body['Company Name B'] || body.companyNameB;
    const cin = body['CIN'] || body.cin;
    const addressB = body['Address B'] || body.addressB;

    // Ensure these are extracting from the correct body fields
    // Updated validation to include new fields in the check
    if (!date && !companyNameA && !llpin && !addressA && !caFirmName && !caName && !memberRegNo && !partnerProprietor && !companyNameB && !cin && !addressB) {
        return res.status(400).json({ success: false, message: 'No data provided to save.' });
    }

    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        await fs.access(EXCEL_FILE_PATH);
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        worksheet = workbook.getWorksheet('User Details');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('User Details');
            // UPDATED: Add new headers for CA Firm details
            worksheet.addRow(['Timestamp', 'User Email', 'Date', 'Company Name A', 'LLPIN', 'Address A', 'CA FIRM NAME', 'CA NAME', 'MEMBER REG NO', 'PARTNER/PROPRIETOR', 'Company Name B', 'CIN', 'Address B']);
        }
    } catch (error) {
        worksheet = workbook.addWorksheet('User Details');
        // UPDATED: Add new headers for CA Firm details
        worksheet.addRow(['Timestamp', 'User Email', 'Date', 'Company Name A', 'LLPIN', 'Address A', 'CA FIRM NAME', 'CA NAME', 'MEMBER REG NO', 'PARTNER/PROPRIETOR', 'Company Name B', 'CIN', 'Address B']);
    }

    // CRITICAL: Make sure the variables in this array are in the EXACT SAME ORDER
    // as your column headers in saved_texts.xlsx.
    worksheet.addRow([
        new Date().toISOString(), // 1. Timestamp
        userEmail,               // 2. User Email
        (function parseDate(d) {
            if (!d) return '';
            // Try ISO/standard parse first
            const dt = new Date(d);
            if (!isNaN(dt)) {
                // strip time portion and return date-only at local midnight
                return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
            }
            // Try common dd/mm/yyyy or dd-mm-yyyy
            const m = String(d).trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
            if (m) {
                const day = parseInt(m[1], 10);
                const month = parseInt(m[2], 10) - 1;
                let year = parseInt(m[3], 10);
                if (year < 100) year += 2000;
                const dt2 = new Date(year, month, day);
                if (!isNaN(dt2)) return new Date(dt2.getFullYear(), dt2.getMonth(), dt2.getDate());
            }
            // fallback: return original string so we don't lose data
            return String(d);
        })(date),
        companyNameA,
        llpin,
        addressA,
        caFirmName,          // NEW
        caName,              // NEW
        memberRegNo,         // NEW
        partnerProprietor,   // NEW
        companyNameB, // <<< Company Name B value
        cin,          // <<< CIN value
        addressB      // <<< Address B value
    ]);

    try {
        await workbook.xlsx.writeFile(EXCEL_FILE_PATH);
        res.json({ success: true, message: 'Details saved to Excel.' });
    } catch (error) {
        console.error('Error writing to Excel file:', error);
        res.status(500).json({ success: false, message: 'Failed to save details to Excel.' });
    }
    // After adding row, set the Date cell number format so Excel shows it as a date
    try {
        // find the header index for 'Date'
        let dateCol = null;
        if (worksheet.getRow(1)) {
            worksheet.getRow(1).eachCell((cell, colNumber) => {
                if (cell && String(cell.value).trim() === 'Date') dateCol = colNumber;
            });
        }
        if (dateCol) {
            const lastRowNumber = worksheet.actualRowCount;
            const lastRow = worksheet.getRow(lastRowNumber);
            const cell = lastRow.getCell(dateCol);
            // If the cell contains a Date object, set Excel number format
            if (cell && cell.value instanceof Date) {
                cell.numFmt = 'dd-mm-yyyy';
            }
        }
    } catch (e) {
        console.warn('Could not set Date column format:', e && e.message ? e.message : e);
    }
});

app.get('/api/get-all-data', async (req, res) => {
    if (!req.isAuthenticated()) {
        return res.status(401).json({ success: false, message: 'Unauthorized' });
    }

    const workbook = new ExcelJS.Workbook();
    let data = { headers: [], rows: [] };

    try {
        await fs.access(EXCEL_FILE_PATH);
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        const worksheet = workbook.getWorksheet('User Details');

        if (worksheet && worksheet.actualRowCount > 1) {
            worksheet.getRow(1).eachCell((cell) => {
                data.headers.push(cell.value);
            });

            for (let i = 2; i <= worksheet.actualRowCount; i++) {
                const row = worksheet.getRow(i);
                const rowData = [];
                row.eachCell((cell) => {
                    rowData.push(cell.value);
                });
                data.rows.push(rowData);
            }
        } else if (worksheet && worksheet.actualRowCount === 1) {
            worksheet.getRow(1).eachCell((cell) => {
                data.headers.push(cell.value);
            });
        }
        res.json({ success: true, ...data });

    } catch (error) {
        if (error.code === 'ENOENT') {
            return res.json({ success: true, message: 'No Excel file found yet.', headers: [], rows: [] });
        }
        console.error('Error reading Excel file:', error);
        res.status(500).json({ success: false, message: 'Failed to read data from Excel.' });
    }
});

// NEW: API endpoint to get only the last entry from the Excel file
app.get('/api/get-last-excel-entry', async (req, res) => {
    if (!req.isAuthenticated()) {
        return res.status(401).json({ success: false, message: 'Unauthorized' });
    }

    const workbook = new ExcelJS.Workbook();
    let lastEntry = { headers: [], row: null }; // Changed 'rows' to 'row' for single entry

    try {
        await fs.access(EXCEL_FILE_PATH);
        await workbook.xlsx.readFile(EXCEL_FILE_PATH);
        const worksheet = workbook.getWorksheet('User Details');

        if (worksheet && worksheet.actualRowCount > 1) { // Check if there's actual data beyond headers
            worksheet.getRow(1).eachCell((cell) => {
                lastEntry.headers.push(cell.value);
            });

            const lastRow = worksheet.getRow(worksheet.actualRowCount);
            const rowData = {};
            lastEntry.headers.forEach((header, index) => {
                rowData[header] = lastRow.getCell(index + 1).value; // ExcelJS is 1-indexed for cells
            });
            lastEntry.row = rowData;
        } else if (worksheet && worksheet.actualRowCount === 1) { // Only headers exist
             worksheet.getRow(1).eachCell((cell) => {
                lastEntry.headers.push(cell.value);
            });
            return res.json({ success: true, message: 'Only headers found in Excel file.', ...lastEntry });
        } else { // No headers or data
            return res.json({ success: true, message: 'No data found in the Excel file yet.', ...lastEntry });
        }
        res.json({ success: true, ...lastEntry });

    } catch (error) {
        if (error.code === 'ENOENT') {
            return res.json({ success: true, message: 'No Excel file found yet.', ...lastEntry });
        }
        console.error('Error reading last Excel entry:', error);
        res.status(500).json({ success: false, message: 'Failed to read last entry from Excel.' });
    }
});

// NEW: API endpoint to get options for Party A from worked_for.xlsx
app.get('/api/get-party-a-options', async (req, res) => {
    if (!req.isAuthenticated()) {
        return res.status(401).json({ success: false, message: 'Unauthorized' });
    }

    const workbook = new ExcelJS.Workbook();
    let options = [];

    try {
        await fs.access(WORKED_FOR_EXCEL_PATH);
        await workbook.xlsx.readFile(WORKED_FOR_EXCEL_PATH);
        const worksheet = workbook.getWorksheet(1); // Assuming data is in the first worksheet

        if (worksheet && worksheet.actualRowCount > 1) {
            const headers = [];
            worksheet.getRow(1).eachCell((cell) => {
                headers.push(cell.value);
            });

            for (let i = 2; i <= worksheet.actualRowCount; i++) {
                const row = worksheet.getRow(i);
                const rowData = {};
                row.eachCell((cell, colNumber) => {
                    const header = headers[colNumber - 1];
                    rowData[header] = cell.value;
                });
                options.push(rowData);
            }
        }
        res.json({ success: true, options });

    } catch (error) {
        if (error.code === 'ENOENT') {
            return res.json({ success: true, message: 'Worked For Excel file not found.', options: [] });
        }
        console.error('Error reading worked_for.xlsx:', error);
        res.status(500).json({ success: false, message: 'Failed to read options for Party A.' });
    }
});

app.get('/logout', (req, res) => {
    req.logout((err) => {
        if (err) { return next(err); }
        res.redirect('/');
    });
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});