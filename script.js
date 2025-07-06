// DOM Elements
const uploaderSelect = document.getElementById('uploader');
const fileInput = document.getElementById('file-upload');
const uploadButton = document.getElementById('upload-button');
const statusMessage = document.getElementById('status-message');
const fileNameDisplay = document.getElementById('file-name-display');
const fileUploadUI = document.getElementById('file-upload-ui');
const buttonText = document.getElementById('button-text');
const buttonSpinner = document.getElementById('button-spinner');
const logContainer = document.getElementById('log-container');
const logStatus = document.getElementById('log-status');


// === EVENT LISTENERS ===
// Fetch logs when the page is fully loaded
document.addEventListener('DOMContentLoaded', fetchLogs);
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

// === MODIFIED UPLOAD FUNCTION ===
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
            // Send additional data for logging
            body: JSON.stringify({
                rows: processedData,
                uploader: uploaderSelect.value,
                fileName: selectedFileName
            }),
        });
        const result = await response.json();
        if (!response.ok) {
            throw new Error(result.message || `Server error: ${response.statusText}`);
        }
        showStatusMessage(result.message, 'success');
        resetFileState();
        fetchLogs(); // Refresh logs after a successful upload
    } catch (error) {
        console.error("Upload failed:", error);
        showStatusMessage(`อัปโหลดล้มเหลว: ${error.message}`, 'error');
    } finally {
        setLoadingState(false);
    }
}

// === LOG FUNCTIONS ===
/**
 * Fetches the latest upload logs from the server.
 */
async function fetchLogs() {
    if(logStatus) logStatus.textContent = 'กำลังโหลดประวัติ...';
    
    try {
        const response = await fetch('/.netlify/functions/getLogs');
        if (!response.ok) {
            throw new Error('ไม่สามารถโหลดประวัติการอัปโหลดได้');
        }
        const logs = await response.json();
        renderLogs(logs);
    } catch (error) {
        console.error('Failed to fetch logs:', error);
        if(logStatus) logStatus.textContent = 'เกิดข้อผิดพลาดในการโหลดประวัติ';
    }
}

/**
 * Renders the fetched logs into the log container.
 * @param {Array<Array<string>>} logs - An array of log entries.
 */
function renderLogs(logs) {
    // Clear previous content
    logContainer.innerHTML = '';

    if (!logs || logs.length === 0) {
        logContainer.innerHTML = '<p class="text-gray-500 text-center py-4">ยังไม่มีประวัติการอัปโหลด</p>';
        return;
    }

    const logList = document.createElement('ul');
    logList.className = 'space-y-3';

    logs.forEach(log => {
        const uploader = log[0] || 'ไม่ระบุ';
        const fileName = log[1] || 'ไม่ระบุชื่อไฟล์';
        const timestamp = log[2] || 'ไม่มีข้อมูลเวลา';

        const listItem = document.createElement('li');
        listItem.className = 'flex flex-col sm:flex-row sm:items-center sm:justify-between p-3 bg-gray-50 rounded-md transition hover:bg-gray-100';

        const mainInfo = document.createElement('div');
        mainInfo.innerHTML = `
            <p class="font-semibold text-gray-800 truncate" title="${fileName}">${fileName}</p>
            <p class="text-sm text-gray-500">อัปโหลดโดย: ${uploader}</p>
        `;

        const timeInfo = document.createElement('p');
        timeInfo.className = 'text-sm text-gray-600 mt-2 sm:mt-0 flex-shrink-0 sm:ml-4';
        timeInfo.textContent = timestamp;

        listItem.appendChild(mainInfo);
        listItem.appendChild(timeInfo);
        logList.appendChild(listItem);
    });

    logContainer.appendChild(logList);
}


// === EXISTING FUNCTIONS (with updates) ===
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

        const formattedDate = formatDate(dateRaw, 'MM/DD/YYYY');
        const formattedClearingDate = formatDate(clearingDateRaw, 'DD/MM/YYYY');
        const formattedCreatedAt = formatDate(createdAtRaw, 'MM/DD/YYYY');
        const formattedUpdatedAt1 = formatDate(updatedAt1Raw, 'DD/MM/YYYY');
        const formattedUpdatedAt2 = formatDate(updatedAt2Raw, 'DD/MM/YYYY');

        return [
            formattedDate, row[headerMap.month], row[headerMap.year], row[headerMap.team],
            row[headerMap.costCenter], row[headerMap.type], row[headerMap.accountGroup],
            row[headerMap.account], row[headerMap.hospital], "", row[headerMap.doctor],
            row[headerMap.event], "", row[headerMap.request], row[headerMap.requestAmount],
            row[headerMap.payby], row[headerMap.payee], row[headerMap.status],
            formattedClearingDate, row[headerMap.clearingAmount], row[headerMap.plan],
            uploaderName, formattedCreatedAt,
            headerMap.updatedBy1 !== -1 ? row[headerMap.updatedBy1] : "", formattedUpdatedAt1,
            headerMap.updatedBy2 !== -1 ? row[headerMap.updatedBy2] : "", formattedUpdatedAt2,
            "", // AB: Updated_date - Leave blank
        ];
    }).filter(row => row !== null);
}

function formatDate(rawValue, targetFormat) {
    if (!rawValue) return "";
    let dateObj;
    if (rawValue instanceof Date && !isNaN(rawValue)) {
        dateObj = rawValue;
    } else if (typeof rawValue === 'number' && rawValue > 1) {
        const date = new Date((rawValue - 25569) * 86400000);
        const tzOffset = date.getTimezoneOffset() * 60000;
        dateObj = new Date(date.getTime() + tzOffset);
    } else if (typeof rawValue === 'string') {
        const parts = rawValue.trim().split(/[\/\-\.]/);
        if (parts.length === 3) {
            let day = parts[0]; let month = parts[1]; let year = parts[2];
            if (year.length === 2) year = '20' + year;
            dateObj = new Date(`${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T00:00:00`);
        }
    }
    if (dateObj instanceof Date && !isNaN(dateObj)) {
        const day = String(dateObj.getDate()).padStart(2, '0');
        const month = String(dateObj.getMonth() + 1).padStart(2, '0');
        const year = dateObj.getFullYear();
        if (targetFormat === 'MM/DD/YYYY') return `${month}/${day}/${year}`;
        if (targetFormat === 'DD/MM/YYYY') return `${day}/${month}/${year}`;
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
