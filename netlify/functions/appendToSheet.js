// Import the Google Sheets API client
const { google } = require('googleapis');

// The ID of your Google Sheet.
// Found in the URL: https://docs.google.com/spreadsheets/d/THIS_IS_THE_ID/edit
const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';

exports.handler = async function (event, context) {
    // 1. Check for POST request and body
    if (event.httpMethod !== 'POST') {
        return {
            statusCode: 405,
            body: JSON.stringify({ message: 'Method Not Allowed' }),
        };
    }

    const { rows } = JSON.parse(event.body);

    if (!rows || !Array.isArray(rows) || rows.length === 0) {
        return {
            statusCode: 400,
            body: JSON.stringify({ message: 'Bad Request: Missing or empty "rows" data.' }),
        };
    }

    // 2. Authenticate with Google Sheets API
    // Credentials are set as environment variables in the Netlify UI
    const auth = new google.auth.GoogleAuth({
        credentials: {
            client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
            // The private key must have newlines replaced with \\n in the Netlify UI
            private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
        },
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const sheets = google.sheets({ version: 'v4', auth });

    // 3. Prepare the request to append data
    const request = {
        spreadsheetId: SPREADSHEET_ID,
        // The range to append to. A1 notation. Appending after the last row with data.
        // 'All expense' is the name of your sheet/tab.
        range: "'All expense'!A1",
        valueInputOption: 'USER_ENTERED', // Interprets values as if a user typed them.
        insertDataOption: 'INSERT_ROWS', // Inserts new rows for the data.
        resource: {
            values: rows,
        },
    };

    // 4. Execute the request
    try {
        const response = await sheets.spreadsheets.values.append(request);
        const updatedRows = response.data.updates.updatedRows || 0;
        
        console.log(`Successfully appended ${updatedRows} rows.`);

        return {
            statusCode: 200,
            body: JSON.stringify({ 
                message: 'Data uploaded successfully!',
                updatedRows: updatedRows 
            }),
        };
    } catch (err) {
        console.error('ERROR appending to sheet:', err.response ? err.response.data : err.message);
        return {
            statusCode: 500,
            body: JSON.stringify({ 
                message: 'Failed to append data to Google Sheet.',
                error: err.response ? err.response.data.error.message : err.message,
            }),
        };
    }
};
