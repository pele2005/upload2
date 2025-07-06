// DOM Elements
const uploaderSelect = document.getElementById('uploader');
const fileInput = document.getElementById('file-upload');
const uploadButton = document.getElementById('upload-button');
const statusMessage = document.getElementById('status-message');
const fileNameDisplay = document.getElementById('file-name-display');
const fileUploadUI = document.getElementById('file-upload-ui');
const buttonText = document.getElementById('button-text');
const buttonSpinner = document.getElementById('button-spinner');

// Store file data globally
let workbookData = null;
let selectedFileName = '';

// === EVENT LISTENERS ===
fileInput.addEventListener('change', handleFileSelect);
fileUploadUI.addEventListener('dragenter', handleDragEnter, false);
fileUploadUI.addEventListener('dragleave', handleDragLeave, false);
fileUploadUI.addEventListener('dragover', handleDragOver, false);
fileUploadUI.addEventListener('drop', handleDrop, false);
uploadButton.addEventListener('click', handleUpload);


// === DRAG AND DROP HANDLERS ===
function handleDragEnter(e) {
    e.preventDefault();
    e.stopPropagation();
    fileUploadUI.classList.add('border-blue-500', 'bg-blue-50');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    fileUploadUI.classList.remove('border-blue-500', 'bg-blue-50');
}

function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    fileUploadUI.classList.remove('border-blue-500', 'bg-blue-50');
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files && files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(event) {
    const files = event.target.files;
    if (files && files.length > 0) {
        processFile(files[0]);
    }
}

function processFile(file) {
    if (!file) return;
    selectedFileName = file.name;
    fileNameDisplay.textContent = `ไฟล์ที่เลือก: ${selectedFileName}`;
    fileUploadUI.classList.add('border-green-500');
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            workbookData = workbook;
            console.log("File read successfully.");
            showStatusMessage('อ่านไฟล์สำเร็จแล้ว กดปุ่มอัปโหลดได้เลย', 'success');
        } catch (error) {
            console.error("Error reading file:", error);
            showStatusMessage('เกิดข้อผิดพลาดในการอ่านไฟล์: ' + error.message, 'error');
            resetFileState();
        }
    };
    reader.onerror = function(error) {
        console.error("FileReader error:", error);
        showStatusMessage('เกิดข้อผิดพลาด: ไม่สามารถอ่านไฟล์ได้', 'error');
        resetFileState();
    };
    reader.readAsArrayBuffer(file);
}

async function handleUpload() {
    if (!workbookData) {
        showStatusMessage('กรุณาเลือกไฟล์ Excel ก่อนครับ', 'warning');
        return;
    }
    setLoadingState(true);
    try {
        const processedData = processWorkbook(workbookData);
        if (!processedData) { // Can be an empty array if file is empty
            throw new Error("ไม่พบข้อมูลในแท็บ 'Data' หรือไฟล์มีรูปแบบไม่ถูกต้อง");
        }
        
        console.log(`Processed ${processedData.length} rows. Sending to server...`);
        const response = await fetch('/.netlify/functions/appendToSheet', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ rows: processedData }),
        });
        const result = await response.json();
        if (!response.ok) {
            throw new Error(result.message || `Server error: ${response.statusText}`);
        }
        showStatusMessage(result.message, 'success');
        resetFileState();
    } catch (error) {
        console.error("Upload failed:", error);
        showStatusMessage(`อัปโหลดล้มเหลว: ${error.message}`, 'error');
    } finally {
        setLoadingState(false);
    }
}

function processWorkbook(workbook) {
    const sheetName = "Data";
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) return null;

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    if (data.length < 2) return [];

    const headers = data.shift().map(h => String(h).trim());
    const headerMap = {
        date: headers.indexOf("Date"),
        month: headers.indexOf("Month"),
        year: headers.indexOf("Year"),
        team: headers.indexOf("Team"),
        costCenter: headers.indexOf("Cost Center"),
        type: headers.indexOf("Type"),
        accountGroup: headers.indexOf("Account Group"),
        account: headers.indexOf("Account"),
        hospital: headers.indexOf("Hospital"),
        doctor: headers.indexOf("Doctor"),
        event: headers.indexOf("Event"),
        request: headers.indexOf("Request"),
        requestAmount: headers.indexOf("Request Amount"),
        payby: headers.indexOf("Payby"),
        payee: headers.indexOf("Payee"),
        status: headers.indexOf("Status"),
        clearingDate: headers.indexOf("Clearing Date"),
        clearingAmount: headers.indexOf("Clearing Amount"),
        plan: headers.indexOf("Plan"),
        createdAt: headers.indexOf("Created At"),
        updatedBy1: headers.indexOf("Updated By"),
        updatedAt1: headers.indexOf("Updated At"),
        updatedBy2: headers.indexOf("Updated By", headers.indexOf("Updated By") + 1),
        updatedAt2: headers.indexOf("Updated At", headers.indexOf("Updated At") + 1),
    };

    const uploaderName = uploaderSelect.value;

    return data.map(row => {
        if (row.every(cell => cell === "")) return null;

        const dateRaw = row[headerMap.date];
        const clearingDateRaw = row[headerMap.clearingDate];
        const createdAtRaw = row[headerMap.createdAt];
        const updatedAt1Raw = headerMap.updatedAt1 !== -1 ? row[headerMap.updatedAt1] : "";
        const updatedAt2Raw = headerMap.updatedAt2 !== -1 ? row[headerMap.updatedAt2] : "";

        // Format dates according to target specification
        const formattedDate = formatDate(dateRaw, 'MM/DD/YYYY');
        const formattedClearingDate = formatDate(clearingDateRaw, 'DD/MM/YYYY');
        const formattedCreatedAt = formatDate(createdAtRaw, 'MM/DD/YYYY');
        const formattedUpdatedAt1 = formatDate(updatedAt1Raw, 'DD/MM/YYYY');
        const formattedUpdatedAt2 = formatDate(updatedAt2Raw, 'DD/MM/YYYY');

        return [
            formattedDate, // A
            row[headerMap.month], // B
            row[headerMap.year], // C
            row[headerMap.team], // D
            row[headerMap.costCenter], // E
            row[headerMap.type], // F
            row[headerMap.accountGroup], // G
            row[headerMap.account], // H
            row[headerMap.hospital], // I
            "", // J: Hospital_Remark
            row[headerMap.doctor], // K
            row[headerMap.event], // L
            "", // M: Description
            row[headerMap.request], // N
            row[headerMap.requestAmount], // O
            row[headerMap.payby], // P
            row[headerMap.payee], // Q
            row[headerMap.status], // R
            formattedClearingDate, // S
            row[headerMap.clearingAmount], // T
            row[headerMap.plan], // U
            uploaderName, // V: Created_By
            formattedCreatedAt, // W
            headerMap.updatedBy1 !== -1 ? row[headerMap.updatedBy1] : "", // X
            formattedUpdatedAt1, // Y
            headerMap.updatedBy2 !== -1 ? row[headerMap.updatedBy2] : "", // Z
            formattedUpdatedAt2, // AA
            "", // AB: Updated_date - ปล่อยให้ว่างเพื่อให้ Server จัดการ
        ];
    }).filter(row => row !== null);
}

/**
 * Formats a date value from various possible Excel inputs into the target format.
 * @param {string|number|Date} rawValue - The raw cell value from SheetJS.
 * @param {'MM/DD/YYYY'|'DD/MM/YYYY'} targetFormat - The desired output format with slashes.
 * @returns {string} The formatted date string, or an empty string if input is invalid.
 */
function formatDate(rawValue, targetFormat) {
    if (!rawValue) return "";

    let dateObj;

    if (rawValue instanceof Date && !isNaN(rawValue)) {
        dateObj = rawValue;
    }
    else if (typeof rawValue === 'number' && rawValue > 1) {
        const date = new Date((rawValue - 25569) * 86400000);
        const tzOffset = date.getTimezoneOffset() * 60000;
        dateObj = new Date(date.getTime() + tzOffset);
    }
    else if (typeof rawValue === 'string') {
        const parts = rawValue.trim().split(/[\/\-\.]/);
        if (parts.length === 3) {
            let day = parts[0];
            let month = parts[1];
            let year = parts[2];
            if (year.length === 2) year = '20' + year;
            dateObj = new Date(`${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T00:00:00`);
        }
    }

    if (dateObj instanceof Date && !isNaN(dateObj)) {
        const day = String(dateObj.getDate()).padStart(2, '0');
        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
        const year = dateObj.getFullYear();

        if (targetFormat === 'MM/DD/YYYY') {
            return `${month}/${day}/${year}`;
        }
        if (targetFormat === 'DD/MM/YYYY') {
            return `${day}/${month}/${year}`;
        }
    }

    console.warn(`Could not parse date: ${rawValue}`);
    return "";
}


function resetFileState() {
    fileInput.value = '';
    workbookData = null;
    selectedFileName = '';
    fileNameDisplay.textContent = 'คลิกเพื่อเลือกไฟล์ หรือลากไฟล์มาวาง';
    fileUploadUI.classList.remove('border-green-500', 'border-blue-500', 'bg-blue-50');
}

function showStatusMessage(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = 'text-center p-4 rounded-lg';
    let typeClasses = '';
    switch (type) {
        case 'success': typeClasses = 'bg-green-100 text-green-800'; break;
        case 'error': typeClasses = 'bg-red-100 text-red-800'; break;
        case 'warning': typeClasses = 'bg-yellow-100 text-yellow-800'; break;
    }
    if (typeClasses) {
        statusMessage.classList.add(...typeClasses.split(' '));
    }
}

function setLoadingState(isLoading) {
    uploadButton.disabled = isLoading;
    if (isLoading) {
        buttonText.classList.add('hidden');
        buttonSpinner.classList.remove('hidden');
    } else {
        buttonText.classList.remove('hidden');
        buttonSpinner.classList.add('hidden');
    }
}
