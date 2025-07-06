// Import the Google Sheets API client
const { google } = require('googleapis');

// The ID of your Google Sheet.
const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';

// The exact name of the sheet (tab) you want to write to.
const SHEET_NAME = 'allexpense';
const LOG_SHEET_NAME = 'UploadLog'; // The new sheet for logging

exports.handler = async function (event, context) {
    if (event.httpMethod !== 'POST') {
        return { statusCode: 405, body: JSON.stringify({ message: 'Method Not Allowed' }) };
    }

    // Destructure new properties from the body
    const { rows: incomingRows, uploader, fileName } = JSON.parse(event.body);

    if (!incomingRows || !Array.isArray(incomingRows)) {
        return { statusCode: 400, body: JSON.stringify({ message: 'Bad Request: Invalid "rows" data.' }) };
    }

    try {
        if (!process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || !process.env.GOOGLE_PRIVATE_KEY) {
            throw new Error('Missing Google credentials in environment variables.');
        }

        const auth = new google.auth.GoogleAuth({
            credentials: {
                client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
            },
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const sheets = google.sheets({ version: 'v4', auth });

        // --- Step 1: Clear old data ---
        await sheets.spreadsheets.values.clear({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!A2:AC`,
        });

        // --- Step 2: Write new data ---
        let updatedRowsCount = 0;
        if (incomingRows.length > 0) {
            const writeRequest = {
                spreadsheetId: SPREADSHEET_ID,
                range: `'${SHEET_NAME}'!A2`,
                valueInputOption: 'RAW',
                resource: { values: incomingRows },
            };
            const writeResponse = await sheets.spreadsheets.values.update(writeRequest);
            updatedRowsCount = writeResponse.data.updatedRows || 0;
        }

        // --- Step 3: Write to UploadLog sheet ---
        const logEntry = [
            uploader || 'N/A',
            fileName || 'N/A',
            new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' })
        ];
        await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${LOG_SHEET_NAME}'!A1`,
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: [logEntry] }
        });

        // --- Step 4: Update timestamp in AB2 ---
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = String(now.getFullYear()).slice(-2);
        const formattedDate = `${day}/${month}/${year}`; // Use slash format for consistency
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!AB2`,
            valueInputOption: 'RAW',
            resource: { values: [[formattedDate]] },
        });

        return {
            statusCode: 200,
            body: JSON.stringify({ message: `อัปโหลดข้อมูลทับของเดิมสำเร็จ! เขียนข้อมูลใหม่ ${updatedRowsCount} แถว`, updatedRows: updatedRowsCount }),
        };

    } catch (err) {
        console.error('ERROR during sheet operation:', err.response ? JSON.stringify(err.response.data, null, 2) : err.message);
        const detailedError = err.response?.data?.error?.message || err.message || 'An unknown error occurred.';
        return {
            statusCode: 500,
            body: JSON.stringify({ message: `เกิดข้อผิดพลาดในการเขียนข้อมูลลง Google Sheet. Reason: ${detailedError}`, error: detailedError }),
        };
    }
};
