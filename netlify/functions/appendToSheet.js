// Import the Google Sheets API client
const { google } = require('googleapis');

// The ID of your Google Sheet.
const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';

// The exact name of the sheet (tab) you want to write to.
const SHEET_NAME = 'allexpense';

exports.handler = async function (event, context) {
    if (event.httpMethod !== 'POST') {
        return { statusCode: 405, body: JSON.stringify({ message: 'Method Not Allowed' }) };
    }

    const { rows: incomingRows } = JSON.parse(event.body);

    // Allow an empty array, which means the user wants to clear the sheet.
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

        // --- ขั้นตอนที่ 1: ล้างข้อมูลเก่าทั้งหมด (ยกเว้นหัวตาราง) ---
        console.log('Clearing existing data from the sheet...');
        await sheets.spreadsheets.values.clear({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!A2:AC`, // ระบุช่วงที่จะล้างข้อมูลให้ครอบคลุมทั้งหมด
        });
        console.log('Sheet cleared successfully.');

        // --- ขั้นตอนที่ 2: เขียนข้อมูลใหม่ทั้งหมดจากไฟล์ Excel (ถ้ามี) ---
        let updatedRowsCount = 0;
        if (incomingRows.length > 0) {
            console.log(`Writing ${incomingRows.length} new rows to the sheet...`);
            const writeRequest = {
                spreadsheetId: SPREADSHEET_ID,
                range: `'${SHEET_NAME}'!A2`, // เริ่มเขียนข้อมูลที่แถว A2
                // !! แก้ไขสำคัญ: ใช้ RAW เพื่อป้องกัน Google Sheets แปลง format อัตโนมัติ !!
                valueInputOption: 'RAW',
                resource: { values: incomingRows },
            };
            // ใช้ .update เพื่อเขียนทับที่ตำแหน่งที่ระบุ
            const writeResponse = await sheets.spreadsheets.values.update(writeRequest);
            updatedRowsCount = writeResponse.data.updatedRows || 0;
            console.log(`Successfully wrote ${updatedRowsCount} rows.`);
        } else {
            console.log('Uploaded file is empty. The sheet is now cleared.');
        }


        // --- ขั้นตอนที่ 3: อัปเดตวันที่ในเซลล์ AB2 ---
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = String(now.getFullYear()).slice(-2);
        const formattedDate = `${day}${month}${year}`;

        const updateTimestampRequest = {
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!AB2`,
            // ใช้ RAW ที่นี่ด้วยเพื่อความสอดคล้องกัน
            valueInputOption: 'RAW',
            resource: { values: [[formattedDate]] },
        };
        await sheets.spreadsheets.values.update(updateTimestampRequest);
        console.log(`Successfully updated cell AB2 with timestamp ${formattedDate}.`);

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
