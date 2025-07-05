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
// Listen for file selection via the click-to-browse button
fileInput.addEventListener('change', handleFileSelect);

// Listen for drag and drop events on the upload UI
fileUploadUI.addEventListener('dragenter', handleDragEnter, false);
fileUploadUI.addEventListener('dragleave', handleDragLeave, false);
fileUploadUI.addEventListener('dragover', handleDragOver, false);
fileUploadUI.addEventListener('drop', handleDrop, false);

// Listen for the main upload button click
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
    // This is necessary to prevent the browser's default behavior
    e.preventDefault();
    e.stopPropagation();
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    fileUploadUI.classList.remove('border-blue-500', 'bg-blue-50');

    // Get the files from the drop event
    const dt = e.dataTransfer;
    const files = dt.files;

    if (files && files.length > 0) {
        // Pass the first file to the main file processing function
        processFile(files[0]);
    }
}

/**
 * Handles the file selection event from the hidden file input.
 * @param {Event} event - The file input change event.
 */
function handleFileSelect(event) {
    const files = event.target.files;
    if (files && files.length > 0) {
        processFile(files[0]);
    }
}

/**
 * Main function to process the selected file (from either drop or click).
 * Reads the file using SheetJS.
 * @param {File} file - The file object to process.
 */
function processFile(file) {
    if (!file) {
        return;
    }

    selectedFileName = file.name;
    fileNameDisplay.textContent = `ไฟล์ที่เลือก: ${selectedFileName}`;
    fileUploadUI.classList.add('border-green-500');

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            workbookData = workbook; // Store workbook for later processing
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


/**
 * Handles the upload button click. Processes and sends data.
 */
async function handleUpload() {
    if (!workbookData) {
        showStatusMessage('กรุณาเลือกไฟล์ Excel ก่อนครับ', 'warning');
        return;
    }

    setLoadingState(true);

    try {
        const processedData = processWorkbook(workbookData);
        if (!processedData || processedData.length === 0) {
            throw new Error("ไม่พบข้อมูลในแท็บ 'Data' หรือไฟล์มีรูปแบบไม่ถูกต้อง");
        }
        
        console.log(`Processed ${processedData.length} rows. Sending to server...`);

        // Send data to the Netlify serverless function
        const response = await fetch('/.netlify/functions/appendToSheet', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ rows: processedData }),
        });

        const result = await response.json();

        if (!response.ok) {
            throw new Error(result.message || `Server error: ${response.statusText}`);
        }
        
        showStatusMessage(`อัปโหลดข้อมูลสำเร็จ! เพิ่มข้อมูล ${result.updatedRows} แถวเรียบร้อยแล้ว`, 'success');
        resetFileState();

    } catch (error) {
        console.error("Upload failed:", error);
        showStatusMessage(`อัปโหลดล้มเหลว: ${error.message}`, 'error');
    } finally {
        setLoadingState(false);
    }
}

/**
 * Processes the workbook to extract, map, and format data.
 * @param {Object} workbook - The workbook object from SheetJS.
 * @returns {Array<Array<string>>} - An array of rows ready for Google Sheets.
 */
function processWorkbook(workbook) {
    const sheetName = "Data"; // Target sheet name
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        return null;
    }

    // Convert sheet to array of arrays (more robust for header mapping)
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    if (data.length < 2) {
        return []; // No data rows
    }

    const headers = data.shift().map(h => String(h).trim()); // Get and clean headers
    
    // Find indices of columns, handling duplicates
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
        // Handle duplicate columns
        updatedBy1: headers.indexOf("Updated By"),
        updatedAt1: headers.indexOf("Updated At"),
        updatedBy2: headers.indexOf("Updated By", headers.indexOf("Updated By") + 1),
        updatedAt2: headers.indexOf("Updated At", headers.indexOf("Updated At") + 1),
    };

    const uploaderName = uploaderSelect.value;
    const uploadTimestamp = new Date().toLocaleString('en-GB'); // DD/MM/YYYY, HH:MM:SS

    return data.map(row => {
        // Skip empty rows
        if (row.every(cell => cell === "")) {
            return null;
        }

        const dateRaw = row[headerMap.date];
        const clearingDateRaw = row[headerMap.clearingDate];
        const createdAtRaw = row[headerMap.createdAt];

        // Format dates
        const formattedDate = formatDate(dateRaw, 'MMDDYYYY');
        const formattedClearingDate = formatDate(clearingDateRaw, 'DDMMYYYY');
        const formattedCreatedAt = formatDate(createdAtRaw, 'MMDDYYYY');

        // Target Google Sheet Column Order:
        return [
            formattedDate, // Date (MMDDYYYY)
            row[headerMap.month], // Month
            row[headerMap.year], // Year
            row[headerMap.team], // Team
            row[headerMap.costCenter], // Cost_Center
            row[headerMap.type], // Type
            row[headerMap.accountGroup], // Account_Group
            row[headerMap.account], // Account
            row[headerMap.hospital], // Hospital
            "", // Hospital_Remark (leave blank)
            row[headerMap.doctor], // Doctor
            row[headerMap.event], // Event
            "", // Description (leave blank)
            row[headerMap.request], // Request
            row[headerMap.requestAmount], // Request_Amount
            row[headerMap.payby], // Payby
            row[headerMap.payee], // Payee
            row[headerMap.status], // Status
            formattedClearingDate, // Clearing_Date (DDMMYYYY)
            row[headerMap.clearingAmount], // Clearing_Amount
            row[headerMap.plan], // Plan
            uploaderName, // Created_By (from dropdown)
            formattedCreatedAt, // Created_At (MMDDYYYY)
            headerMap.updatedBy1 !== -1 ? row[headerMap.updatedBy1] : "", // Updated_By
            headerMap.updatedAt1 !== -1 ? formatDate(row[headerMap.updatedAt1], 'DDMMYYYY') : "", // Update_At
            headerMap.updatedBy2 !== -1 ? row[headerMap.updatedBy2] : "", // Updated_By_2
            headerMap.updatedAt2 !== -1 ? formatDate(row[headerMap.updatedAt2], 'DDMMYYYY') : "", // Updated_At_2
            uploadTimestamp, // Updated_date
        ];
    }).filter(row => row !== null); // Remove empty rows
}

/**
 * Formats a date string from DDMMYY to the target format.
 * Handles various input types from SheetJS.
 * @param {string|number|Date} rawValue - The raw cell value.
 * @param {'MMDDYYYY'|'DDMMYYYY'} targetFormat - The desired output format.
 * @returns {string} The formatted date string.
 */
function formatDate(rawValue, targetFormat) {
    if (!rawValue) return "";

    let dateStr = String(rawValue);
    
    // If it's a date object from SheetJS
    if (rawValue instanceof Date) {
        const day = String(rawValue.getDate()).padStart(2, '0');
        const month = String(rawValue.getMonth() + 1).padStart(2, '0');
        const year = rawValue.getFullYear();
        if (targetFormat === 'MMDDYYYY') return `${month}${day}${year}`;
        if (targetFormat === 'DDMMYYYY') return `${day}${month}${year}`;
    }

    // Handle Excel's numeric date format
    if (typeof rawValue === 'number' && rawValue > 10000) {
       const dateObj = XLSX.SSF.parse_date_code(rawValue);
       const day = String(dateObj.d).padStart(2, '0');
       const month = String(dateObj.m).padStart(2, '0');
       const year = dateObj.y;
       if (targetFormat === 'MMDDYYYY') return `${month}${day}${year}`;
       if (targetFormat === 'DDMMYYYY') return `${day}${month}${year}`;
    }
    
    // Handle DDMMYY string format
    dateStr = dateStr.replace(/[^0-9]/g, ''); // Clean non-numeric chars
    if (dateStr.length !== 6) return dateStr; // Return original if not in expected format

    const day = dateStr.substring(0, 2);
    const month = dateStr.substring(2, 4);
    const year = '20' + dateStr.substring(4, 6);

    if (targetFormat === 'MMDDYYYY') return `${month}${day}${year}`;
    if (targetFormat === 'DDMMYYYY') return `${day}${month}${year}`;
    
    return dateStr; // Fallback
}


/**
 * Resets the file input and related UI elements.
 */
function resetFileState() {
    fileInput.value = ''; // Clear the file input
    workbookData = null;
    selectedFileName = '';
    fileNameDisplay.textContent = 'คลิกเพื่อเลือกไฟล์ หรือลากไฟล์มาวาง';
    fileUploadUI.classList.remove('border-green-500', 'border-blue-500', 'bg-blue-50');
}

/**
 * Shows a status message to the user.
 * @param {string} message - The message to display.
 * @param {'success'|'error'|'warning'} type - The type of message.
 */
function showStatusMessage(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = 'text-center p-4 rounded-lg'; // Reset classes to a base state

    let typeClasses = '';
    switch (type) {
        case 'success':
            typeClasses = 'bg-green-100 text-green-800';
            break;
        case 'error':
            typeClasses = 'bg-red-100 text-red-800';
            break;
        case 'warning':
            typeClasses = 'bg-yellow-100 text-yellow-800';
            break;
    }
    
    // Correctly add multiple classes by splitting the string and using the spread operator.
    if (typeClasses) {
        statusMessage.classList.add(...typeClasses.split(' '));
    }
}

/**
 * Toggles the loading state of the upload button.
 * @param {boolean} isLoading - Whether to show the loading state.
 */
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
