// Import the Google Sheets API client
const { google } = require('googleapis');

// The ID of your Google Sheet.
const SPREADSHEET_ID = '1iQ18yGtavcRAlD0Gu3Igr2qpCuFGT4dl4b32lWBTOdY';

// =================================================================
// !! สำคัญมาก !!
// ชื่อของ "แท็บ" (Sheet Name) ที่ต้องการให้บันทึกข้อมูล
// อัปเดตเป็น 'allexpense' ตามที่ร้องขอ
// =================================================================
const SHEET_NAME = 'allexpense';

exports.handler = async function (event, context) {
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

        // --- ขั้นตอนที่ 1: เพิ่มข้อมูลแถวใหม่ (Append Rows) ---
        const appendRequest = {
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!A1`,
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: rows },
        };

        const appendResponse = await sheets.spreadsheets.values.append(appendRequest);
        const updatedRows = appendResponse.data.updates.updatedRows || 0;
        
        console.log(`Successfully appended ${updatedRows} rows.`);

        // --- ขั้นตอนที่ 2: อัปเดตวันที่ในเซลล์ AB2 (Update Timestamp Cell) ---
        // สร้างวันที่ปัจจุบันในรูปแบบ DDMMYY
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = String(now.getFullYear()).slice(-2); // เอาแค่ 2 ตัวท้ายของปี
        const formattedDate = `${day}${month}${year}`;

        const updateRequest = {
            spreadsheetId: SPREADSHEET_ID,
            range: `'${SHEET_NAME}'!AB2`, // ระบุเซลล์เป้าหมายคือ AB2
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[formattedDate]] // ข้อมูลวันที่ที่ต้องการเขียน
            }
        };

        await sheets.spreadsheets.values.update(updateRequest);
        console.log(`Successfully updated cell AB2 with timestamp ${formattedDate}.`);


        return {
            statusCode: 200,
            body: JSON.stringify({ message: 'อัปโหลดข้อมูลและอัปเดตวันที่สำเร็จ!', updatedRows: updatedRows }),
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
