// Import the Google Sheets API client
const { google } = require('googleapis');

// The ID of your Google Sheet.
const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';

// The exact name of the sheet (tab) you want to write to.
const SHEET_NAME = 'All expense';

exports.handler = async function (event, context) {
    // === DEBUGGING LOGS START ===
    // These logs will help us verify the environment variables in Netlify.
    console.log('[DEBUG] Function invoked.');
    console.log(`[DEBUG] GOOGLE_SERVICE_ACCOUNT_EMAIL type: ${typeof process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL}`);
    console.log(`[DEBUG] GOOGLE_SERVICE_ACCOUNT_EMAIL value (first 15 chars): ${String(process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL).substring(0, 15)}`);
    console.log(`[DEBUG] GOOGLE_PRIVATE_KEY type: ${typeof process.env.GOOGLE_PRIVATE_KEY}`);
    console.log(`[DEBUG] GOOGLE_PRIVATE_KEY value (first 30 chars): ${String(process.env.GOOGLE_PRIVATE_KEY).substring(0, 30)}`);
    // === DEBUGGING LOGS END ===

    if (event.httpMethod !== 'POST') {
        return { statusCode: 405, body: JSON.stringify({ message: 'Method Not Allowed' }) };
    }

    const { rows } = JSON.parse(event.body);

    if (!rows || !Array.isArray(rows) || rows.length === 0) {
        return { statusCode: 400, body: JSON.stringify({ message: 'Bad Request: Missing or empty "rows" data.' }) };
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

        const request = {
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!A1`,
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: rows },
        };

        const response = await sheets.spreadsheets.values.append(request);
        const updatedRows = response.data.updates.updatedRows || 0;
        
        console.log(`Successfully appended ${updatedRows} rows.`);

        return {
            statusCode: 200,
            body: JSON.stringify({ message: 'Data uploaded successfully!', updatedRows: updatedRows }),
        };
    } catch (err) {
        console.error('ERROR appending to sheet:', err.response ? JSON.stringify(err.response.data, null, 2) : err.message);
        const detailedError = err.response?.data?.error?.message || err.message || 'An unknown error occurred.';
        return {
            statusCode: 500,
            body: JSON.stringify({ message: `Failed to append data to Google Sheet. Reason: ${detailedError}`, error: detailedError }),
        };
    }
};
