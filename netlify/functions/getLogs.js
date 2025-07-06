const { google } = require('googleapis');

const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';
const LOG_SHEET_NAME = 'UploadLog'; // The new sheet for logging

exports.handler = async function (event, context) {
    if (event.httpMethod !== 'GET') {
        return { statusCode: 405, body: JSON.stringify({ message: 'Method Not Allowed' }) };
    }

    try {
        const auth = new google.auth.GoogleAuth({
            credentials: {
                client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
                private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
            },
            // Use read-only scope as we are only reading data
            scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
        });

        const sheets = google.sheets({ version: 'v4', auth });

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${LOG_SHEET_NAME}'!A:C`, // Read columns A, B, C
        });

        const rows = response.data.values || [];

        // Reverse the array to get the latest logs first, then take the top 5
        const latestLogs = rows.reverse().slice(0, 5);

        return {
            statusCode: 200,
            body: JSON.stringify(latestLogs),
        };

    } catch (err) {
        console.error('ERROR fetching logs:', err);
        // If the sheet doesn't exist, return an empty array instead of an error
        if (err.code === 400 && err.message.includes('Unable to parse range')) {
             return {
                statusCode: 200,
                body: JSON.stringify([]), // Return empty log if sheet not found
            };
        }
        return {
            statusCode: 500,
            body: JSON.stringify({ message: 'Failed to fetch logs.', error: err.message }),
        };
    }
};
