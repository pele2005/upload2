// Import the Google Sheets API client
const { google } = require('googleapis');

// The ID of your Google Sheet.
const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';

// The exact name of the sheet (tab) you want to write to.
const SHEET_NAME = 'allexpense';

/**
 * สร้าง Key ที่ไม่ซ้ำกันสำหรับแต่ละแถวเพื่อใช้ในการเปรียบเทียบ
 * @param {Array<string>} row - แถวข้อมูลจากชีต
 * @returns {string} - Key ที่ไม่ซ้ำกันสำหรับแถวนั้น
 */
function createRowKey(row) {
    // รวมข้อมูลจากคอลัมน์สำคัญเพื่อสร้าง Key
    // คอลัมน์: 0:Date, 3:Team, 4:Cost_Center, 7:Account, 8:Hospital, 10:Doctor, 14:Request_Amount
    if (!row || row.length < 15) {
        return null; // ถ้าแถวข้อมูลไม่สมบูรณ์ ให้ข้ามไป
    }
    const keyParts = [
        row[0],  // Date
        row[3],  // Team
        row[4],  // Cost_Center
        row[7],  // Account
        row[8],  // Hospital
        row[10], // Doctor
        row[14]  // Request_Amount
    ];
    return keyParts.join('|');
}


exports.handler = async function (event, context) {
    if (event.httpMethod !== 'POST') {
        return { statusCode: 405, body: JSON.stringify({ message: 'Method Not Allowed' }) };
    }

    const { rows: incomingRows } = JSON.parse(event.body);

    if (!incomingRows || !Array.isArray(incomingRows) || incomingRows.length === 0) {
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

        // --- ขั้นตอนที่ 1: ดึงข้อมูลที่มีอยู่ทั้งหมดจาก Google Sheet ---
        console.log('Fetching existing data from Google Sheet...');
        const getRowsResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!A2:Z`, // อ่านตั้งแต่ A2 เพื่อข้ามหัวตาราง
        });
        const existingRows = getRowsResponse.data.values || [];
        
        // --- ขั้นตอนที่ 2: สร้าง Set ของ Key ที่มีอยู่เพื่อการค้นหาที่รวดเร็ว ---
        const existingKeys = new Set(existingRows.map(createRowKey).filter(key => key !== null));
        console.log(`Found ${existingKeys.size} existing unique keys in the sheet.`);

        // --- ขั้นตอนที่ 3: กรองข้อมูลใหม่ที่ยังไม่มีในชีต ---
        const newRowsToAppend = incomingRows.filter(row => {
            const rowKey = createRowKey(row);
            return rowKey && !existingKeys.has(rowKey);
        });

        console.log(`Received ${incomingRows.length} rows from upload, found ${newRowsToAppend.length} new rows to append.`);

        // --- ขั้นตอนที่ 4: เพิ่มเฉพาะข้อมูลใหม่และอัปเดตวันที่ (ถ้ามี) ---
        if (newRowsToAppend.length === 0) {
            console.log('No new data to append.');
            return {
                statusCode: 200,
                body: JSON.stringify({ message: 'ไม่พบข้อมูลใหม่ให้อัปเดต', updatedRows: 0 }),
            };
        }

        // เพิ่มข้อมูลใหม่
        const appendRequest = {
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!A1`,
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: newRowsToAppend },
        };
        const appendResponse = await sheets.spreadsheets.values.append(appendRequest);
        const updatedRowsCount = appendResponse.data.updates.updatedRows || 0;
        console.log(`Successfully appended ${updatedRowsCount} new rows.`);

        // อัปเดตวันที่ในเซลล์ AB2
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = String(now.getFullYear()).slice(-2);
        const formattedDate = `${day}${month}${year}`;

        const updateRequest = {
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!AB2`,
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[formattedDate]] },
        };
        await sheets.spreadsheets.values.update(updateRequest);
        console.log(`Successfully updated cell AB2 with timestamp ${formattedDate}.`);

        return {
            statusCode: 200,
            body: JSON.stringify({ message: `อัปเดตข้อมูลสำเร็จ! เพิ่มข้อมูลใหม่ ${updatedRowsCount} แถว`, updatedRows: updatedRowsCount }),
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
