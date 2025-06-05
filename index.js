
/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

// For actual Excel parsing, you would import a library like SheetJS:
// import * as XLSX from 'xlsx'; // Or whatever the library's import mechanism is.
// Since we cannot add external libraries in this environment without user action,
// we'll use a placeholder and a very simple CSV-like parser if no library is found.
var XLSX; // Declare XLSX to avoid TypeScript errors if the library is loaded globally.


document.addEventListener('DOMContentLoaded', () => {
    const xmlFileInput = document.getElementById('xmlFile');
    const excelFileInput = document.getElementById('excelFile');
    const processFileButton = document.getElementById('processFileButton');
    const resultContainer = document.getElementById('resultContainer');
    
    let employeeData = null; // Stores Cert -> Name

    if (!xmlFileInput || !excelFileInput || !processFileButton || !resultContainer) {
        console.error('Required HTML elements are missing.');
        if (resultContainer) {
            resultContainer.innerHTML = '<p class="error-message">Lỗi: Không thể khởi tạo giao diện người dùng ứng dụng. Thiếu các yếu tố bắt buộc.</p>';
        }
        return;
    }

    processFileButton.addEventListener('click', handleFileProcessing);
    
    xmlFileInput.addEventListener('change', () => {
        resetPartialUIForNewFiles();
        processFileButton.disabled = !xmlFileInput.files || xmlFileInput.files.length === 0;
    });
    
    excelFileInput.addEventListener('change', () => {
        employeeData = null; // Clear previous employee data if a new file is selected/deselected
        if (excelFileInput.files && excelFileInput.files.length > 0) {
            // Optionally, you could trigger a pre-parse or validation here
        }
        resetPartialUIForNewFiles(); // Reset messages, but don't disable button if XML is still there
    });
    
    processFileButton.disabled = true; 

    function resetPartialUIForNewFiles() {
        // Only reset the result container, not the button's disabled state fully yet
        // as it depends on XML input primarily.
        const existingTable = resultContainer.querySelector('.results-table');
        const existingDownloadButton = resultContainer.querySelector('#downloadExcelButton');
        const existingMessage = resultContainer.querySelector('.error-message') || resultContainer.querySelector('.info-message');

        if (existingTable || existingDownloadButton || existingMessage) {
            resultContainer.innerHTML = '';
            const initialMsgP = document.createElement('p');
            initialMsgP.id = 'initialMessage';
            initialMsgP.innerHTML = '1. Tải lên một hoặc nhiều tệp XML chứa dữ liệu chấm công.<br>2. (Tùy chọn) Tải lên tệp Excel danh sách nhân viên.<br>3. Nhấp "Tạo Bảng Chấm Công".';
            resultContainer.appendChild(initialMsgP);
        }
    }
    
    function resetUI() {
        resultContainer.innerHTML = '';
        const initialMsgP = document.createElement('p');
        initialMsgP.id = 'initialMessage';
        initialMsgP.innerHTML = '1. Tải lên một hoặc nhiều tệp XML chứa dữ liệu chấm công.<br>2. (Tùy chọn) Tải lên tệp Excel danh sách nhân viên.<br>3. Nhấp "Tạo Bảng Chấm Công".';
        resultContainer.appendChild(initialMsgP);
        xmlFileInput.value = '';
        excelFileInput.value = '';
        employeeData = null;
        processFileButton.disabled = true;
    }

    function parseNgayYl(ngayYlStr) {
        if (!ngayYlStr || ngayYlStr.length !== 12) return null; // yyyyMMddHHmm
        const year = parseInt(ngayYlStr.substring(0, 4), 10);
        const month = parseInt(ngayYlStr.substring(4, 6), 10) - 1; // JS months are 0-indexed
        const day = parseInt(ngayYlStr.substring(6, 8), 10);
        const hours = parseInt(ngayYlStr.substring(8, 10), 10);
        const minutes = parseInt(ngayYlStr.substring(10, 12), 10);

        if (isNaN(year) || isNaN(month) || isNaN(day) || isNaN(hours) || isNaN(minutes)) {
            return null;
        }
        const date = new Date(year, month, day, hours, minutes);
        if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day || date.getHours() !== hours || date.getMinutes() !== minutes) {
            return null;
        }
        return date;
    }

    function formatTime(date) {
        const hours = date.getHours().toString().padStart(2, '0');
        const minutes = date.getMinutes().toString().padStart(2, '0');
        return `${hours}:${minutes}`;
    }

    async function parseExcelData(file) {
        const MAX_HEADER_SCAN_ROWS = 10; // Scan up to 10 rows for headers
        
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (event) => {
                try {
                    const data = event.target?.result;
                    const newEmployeeData = new Map();
                    let parsedSuccessfully = false;

                    const canUseXLSX = typeof XLSX !== 'undefined';
                    const isExcelFileByMimeOrExt = file.type.startsWith('application/vnd.ms-excel') ||
                                               file.type.startsWith('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') ||
                                               file.name.toLowerCase().endsWith('.xls') ||
                                               file.name.toLowerCase().endsWith('.xlsx');
                    const isCsvOrTxtFileByMimeOrExt = file.type === 'text/csv' ||
                                                 file.name.toLowerCase().endsWith('.csv') ||
                                                 file.name.toLowerCase().endsWith('.txt');

                    if (canUseXLSX && isExcelFileByMimeOrExt && data instanceof ArrayBuffer) {
                        const workbook = XLSX.read(data, { type: 'array' });
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        
                        let headerRowIndex = -1;
                        let nameColIdx = -1;
                        let certColIdx = -1;

                        if (jsonData && jsonData.length > 0) {
                            for (let i = 0; i < Math.min(jsonData.length, MAX_HEADER_SCAN_ROWS); i++) {
                                const currentRow = (jsonData[i]).map(h => String(h === null || h === undefined ? "" : h).toLowerCase().trim());
                                const tempNameColIdx = currentRow.findIndex(h => 
                                    typeof h === 'string' && (h.includes('tên nhân viên') || h.includes('ten nhan vien') || h.includes('họ tên') || h.includes('ho ten'))
                                );
                                const tempCertColIdx = currentRow.findIndex(h => 
                                    typeof h === 'string' && (h.includes('chứng chỉ hành nghề') || h.includes('cchn') || h.includes('ma_bac_si') || h.includes('mã bác sĩ'))
                                );
                                if (tempNameColIdx !== -1 && tempCertColIdx !== -1) {
                                    headerRowIndex = i;
                                    nameColIdx = tempNameColIdx;
                                    certColIdx = tempCertColIdx;
                                    break;
                                }
                            }
                        }

                        if (headerRowIndex === -1 || nameColIdx === -1 || certColIdx === -1) {
                            displayMessage(`Không tìm thấy cột "Tên nhân viên" và "Chứng chỉ hành nghề" trong ${MAX_HEADER_SCAN_ROWS} hàng đầu tiên của tệp Excel "${file.name}".`, 'error-message');
                            resolve(newEmployeeData); 
                            return;
                        }

                        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                            const row = jsonData[i];
                            if (row && row.length > Math.max(nameColIdx, certColIdx)) { 
                                const name = row[nameColIdx]?.trim();
                                const cert = row[certColIdx]?.trim();
                                if (name && cert) {
                                    newEmployeeData.set(cert, name);
                                }
                            }
                        }
                        parsedSuccessfully = true;
                    } else if (isCsvOrTxtFileByMimeOrExt && typeof data === 'string') {
                        const text = data;
                        const lines = text.split(/\r\n|\n/);
                        let headerRowIndex = -1;
                        let nameColIdx = -1;
                        let certColIdx = -1;

                        if (lines.length > 0) {
                            for (let i = 0; i < Math.min(lines.length, MAX_HEADER_SCAN_ROWS); i++) {
                                const currentRow = lines[i].split(',').map(h => String(h === null || h === undefined ? "" : h).toLowerCase().trim());
                                const tempNameColIdx = currentRow.findIndex(h => 
                                    typeof h === 'string' && (h.includes('tên nhân viên') || h.includes('ten nhan vien') || h.includes('họ tên') || h.includes('ho ten'))
                                );
                                const tempCertColIdx = currentRow.findIndex(h => 
                                    typeof h === 'string' && (h.includes('chứng chỉ hành nghề') || h.includes('cchn') || h.includes('ma_bac_si') || h.includes('mã bác sĩ'))
                                );
                                if (tempNameColIdx !== -1 && tempCertColIdx !== -1) {
                                    headerRowIndex = i;
                                    nameColIdx = tempNameColIdx;
                                    certColIdx = tempCertColIdx;
                                    break;
                                }
                            }
                        }

                        if (headerRowIndex === -1 || nameColIdx === -1 || certColIdx === -1) {
                             displayMessage(`Không tìm thấy cột "Tên nhân viên" và "Chứng chỉ hành nghề" trong ${MAX_HEADER_SCAN_ROWS} dòng đầu tiên của tệp CSV/TXT "${file.name}".`, 'error-message');
                             resolve(newEmployeeData);
                             return;
                        }
                        for (let i = headerRowIndex + 1; i < lines.length; i++) {
                            const columns = lines[i].split(',');
                            if (columns.length > Math.max(nameColIdx, certColIdx)) {
                                const name = columns[nameColIdx]?.trim();
                                const cert = columns[certColIdx]?.trim();
                                if (name && cert) {
                                    newEmployeeData.set(cert, name);
                                }
                            }
                        }
                        parsedSuccessfully = true;
                    }

                    if (!parsedSuccessfully && (isExcelFileByMimeOrExt || isCsvOrTxtFileByMimeOrExt) ) {
                         displayMessage(`Không thể trích xuất dữ liệu nhân viên từ ${file.name}. Tệp có thể trống hoặc không có các cột dự kiến.`, 'info-message', true);
                    } else if (newEmployeeData.size > 0) {
                        displayMessage(`Đã tải ${newEmployeeData.size} mục nhân viên từ ${file.name}.`, 'info-message', true);
                    } else if (parsedSuccessfully && newEmployeeData.size === 0) {
                         displayMessage(`Không tìm thấy dữ liệu nhân viên hợp lệ (tên và chứng chỉ) trong ${file.name}, hoặc tệp trống.`, 'info-message', true);
                    }

                    resolve(newEmployeeData);
                } catch (e) {
                    console.error("Lỗi phân tích cú pháp tệp Excel/CSV: ", e);
                    displayMessage(`Lỗi khi phân tích cú pháp nội dung tệp nhân viên ${file.name}. Chi tiết: ${e.message}`, 'error-message');
                    resolve(new Map()); 
                }
            };

            reader.onerror = () => {
                reject(new Error(`Lỗi đọc tệp nhân viên: ${file.name}`));
            };

            const canUseXLSXLib = typeof XLSX !== 'undefined';
            const isExcelFile = file.name.toLowerCase().endsWith('.xls') ||
                                file.name.toLowerCase().endsWith('.xlsx') ||
                                file.type.startsWith('application/vnd.ms-excel') ||
                                file.type.startsWith('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            const isCsvTxtFile = file.name.toLowerCase().endsWith('.csv') ||
                                 file.name.toLowerCase().endsWith('.txt') ||
                                 file.type === 'text/csv';

            if (canUseXLSXLib && isExcelFile) {
                reader.readAsArrayBuffer(file);
            } else if (isCsvTxtFile) {
                reader.readAsText(file);
            } else {
                 let message = `Loại tệp không được hỗ trợ cho danh sách nhân viên: ${file.name}. Vui lòng sử dụng .xls, .xlsx, .csv hoặc .txt.`;
                 if (isExcelFile && !canUseXLSXLib) { 
                     message = `Để xử lý tệp Excel ("${file.name}"), thư viện XLSX (SheetJS) cần được tải. Vui lòng đảm bảo thư viện được kích hoạt trong tệp HTML. Hiện tại, chỉ có thể xử lý .csv hoặc .txt.`;
                 }
                 displayMessage(message, 'error-message');
                 resolve(new Map());
            }
        });
    }


    async function handleFileProcessing() {
        resultContainer.innerHTML = ''; 

        const xmlFiles = xmlFileInput.files;
        if (!xmlFiles || xmlFiles.length === 0) {
            displayMessage('Vui lòng chọn một hoặc nhiều tệp XML trước.', 'error-message');
            processFileButton.disabled = true;
            return;
        }

        displayMessage('Đang xử lý các tệp...', 'info-message');
        processFileButton.disabled = true;
        xmlFileInput.disabled = true;
        excelFileInput.disabled = true;

        let currentEmployeeData = null;
        if (excelFileInput.files && excelFileInput.files.length > 0) {
            try {
                currentEmployeeData = await parseExcelData(excelFileInput.files[0]);
                employeeData = currentEmployeeData; 
            } catch (error) {
                console.error('Lỗi xử lý tệp Excel:', error);
                displayMessage(`Lỗi xử lý tệp nhân viên: ${error.message}`, 'error-message', true);
            }
        } else {
            employeeData = null; 
        }


        const xmlDocuments = [];
        const fileReadPromises = [];
        const fileNames = [];

        for (const file of Array.from(xmlFiles)) {
            if (file.type !== 'text/xml' && !file.name.toLowerCase().endsWith('.xml')) {
                displayMessage(`Loại tệp không hợp lệ: ${file.name}. Vui lòng chỉ chọn tệp XML (.xml).`, 'error-message', true);
                xmlFileInput.value = ''; 
                processFileButton.disabled = false;
                xmlFileInput.disabled = false;
                excelFileInput.disabled = false;
                return;
            }
            fileNames.push(file.name);
            fileReadPromises.push(new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (event) => resolve(event.target?.result);
                reader.onerror = () => reject(new Error(`Lỗi đọc tệp: ${file.name}`));
                reader.readAsText(file);
            }));
        }

        try {
            const fileContents = await Promise.all(fileReadPromises);

            for (let i = 0; i < fileContents.length; i++) {
                const fileContent = fileContents[i];
                const fileName = fileNames[i];
                if (!fileContent) {
                    throw new Error(`Nội dung tệp ${fileName} trống hoặc không thể đọc được.`);
                }

                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(fileContent, 'text/xml');
                const parserError = xmlDoc.getElementsByTagName('parsererror');

                if (parserError.length > 0) {
                    let errorMessageText = `Tệp XML ${fileName} không hợp lệ. Vui lòng kiểm tra nội dung tệp.`;
                    const errorDetails = parserError[0].textContent;
                     if (errorDetails) {
                       const specificErrorMatch = errorDetails.match(/Error: (.*?) at line/s) || errorDetails.match(/^(.*?)\nLocation:/s);
                       if (specificErrorMatch && specificErrorMatch[1]) {
                           errorMessageText = `XML không hợp lệ trong ${fileName}: ${specificErrorMatch[1].trim()}`;
                       }
                    }
                    throw new Error(errorMessageText);
                }
                xmlDocuments.push(xmlDoc);
            }
            
            generateTimesheet(xmlDocuments, employeeData);

        } catch (error) {
            console.error('Lỗi xử lý các tệp XML:', error);
            const errorMessage = (error instanceof Error) ? error.message : 'Đã xảy ra lỗi không mong muốn khi xử lý tệp.';
            displayMessage(`Lỗi XML: ${errorMessage}`, 'error-message', !employeeData); 
        } finally {
            processFileButton.disabled = false;
            xmlFileInput.disabled = false;
            excelFileInput.disabled = false;
        }
    }

    function generateTimesheet(xmlDocs, currentEmployeeData) {
        const allEntries = [];

        const processElementsFromDoc = (xmlDoc, entries) => {
            const parentTagNames = ['CHI_TIET_THUOC', 'CHI_TIET_DVKT'];
        
            parentTagNames.forEach(parentTagName => {
                const elements = xmlDoc.getElementsByTagName(parentTagName);
                for (const el of Array.from(elements)) {
                    // Check for MA_BAC_SI entry
                    const maBacSiEl = el.getElementsByTagName('MA_BAC_SI')[0];
                    const ngayYlEl = el.getElementsByTagName('NGAY_YL')[0];
                    const maBacSi = maBacSiEl?.textContent?.trim();
                    const ngayYlStr = ngayYlEl?.textContent?.trim();
        
                    if (maBacSi && ngayYlStr) {
                        const dateTime = parseNgayYl(ngayYlStr);
                        if (dateTime) {
                            entries.push({ id: maBacSi, dateTime, isNthEntry: false });
                        }
                    }
        
                    // Check for NGUOI_THUC_HIEN entry
                    const nguoiThucHienEl = el.getElementsByTagName('NGUOI_THUC_HIEN')[0];
                    const ngayThYlEl = el.getElementsByTagName('NGAY_TH_YL')[0];
                    const nguoiThucHienRaw = nguoiThucHienEl?.textContent?.trim();
                    const ngayThYlStr = ngayThYlEl?.textContent?.trim();
        
                    if (nguoiThucHienRaw && ngayThYlStr) {
                        const dateTime = parseNgayYl(ngayThYlStr);
                        if (dateTime) {
                            if (nguoiThucHienRaw.includes(';')) {
                                const individualIds = nguoiThucHienRaw.split(';');
                                for (const idStr of individualIds) {
                                    const trimmedId = idStr.trim();
                                    if (trimmedId) { // Ensure non-empty ID after trim
                                        entries.push({ id: trimmedId, dateTime, isNthEntry: true });
                                    }
                                }
                            } else {
                                entries.push({ id: nguoiThucHienRaw, dateTime, isNthEntry: true });
                            }
                        }
                    }
                }
            });
        };

        for (const xmlDoc of xmlDocs) {
            processElementsFromDoc(xmlDoc, allEntries);
        }
        
        if (allEntries.length === 0) {
            displayMessage('Không tìm thấy dữ liệu chấm công (MA_BAC_SI, NGAY_YL, NGUOI_THUC_HIEN, NGAY_TH_YL) hợp lệ trong các tệp XML đã chọn.', 'info-message', !currentEmployeeData);
            return;
        }

        allEntries.sort((a, b) => a.dateTime.getTime() - b.dateTime.getTime()); 
        const firstEntryDate = allEntries[0].dateTime;
        const targetYear = firstEntryDate.getFullYear();
        const targetMonth = firstEntryDate.getMonth(); // 0-indexed

        const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
        const timesheetData = {};

        for (const entry of allEntries) {
            const { id, dateTime, isNthEntry } = entry;

            if (dateTime.getFullYear() !== targetYear || dateTime.getMonth() !== targetMonth) {
                continue; 
            }

            const day = dateTime.getDate();
            const hours = dateTime.getHours();
            const minutes = dateTime.getMinutes();
            const timeValue = hours * 100 + minutes;

            if (!timesheetData[id]) {
                timesheetData[id] = {};
            }
            if (!timesheetData[id][day]) {
                timesheetData[id][day] = { S: false, C: false, sTimes: [], cTimes: [], S_isNth: false, C_isNth: false };
            }

            if (timeValue >= 700 && timeValue < 1130) { // Morning shift
                if (!timesheetData[id][day].S) {
                     timesheetData[id][day].S = true;
                }
                if (isNthEntry) timesheetData[id][day].S_isNth = true;
                timesheetData[id][day].sTimes.push(dateTime);
            } else if (timeValue >= 1330 && timeValue < 1700) {  // Afternoon shift
                if (!timesheetData[id][day].C) {
                    timesheetData[id][day].C = true;
                }
                if (isNthEntry) timesheetData[id][day].C_isNth = true;
                timesheetData[id][day].cTimes.push(dateTime);
            }
        }
        
        displayTimesheetTable(timesheetData, targetYear, targetMonth, daysInMonth, currentEmployeeData);
    }
    
    function getDayOfWeek(year, month, day) {
        // month is 0-indexed
        return new Date(year, month, day).getDay(); // 0 for Sunday, 6 for Saturday
    }

    function getDayOfWeekString(year, month, day) {
        const dayIndex = getDayOfWeek(year, month, day);
        const days = ["CN", "T2", "T3", "T4", "T5", "T6", "T7"];
        return days[dayIndex];
    }

    function displayTimesheetTable(
        timesheetData,
        year,
        month, // 0-indexed
        daysInMonth,
        currentEmployeeData
    ) {
        const existingMessages = resultContainer.innerHTML;
        let clearContainer = true;
        if (typeof existingMessages === 'string' && (existingMessages.includes('info-message') || existingMessages.includes('error-message'))) {
            clearContainer = false; 
        }

        if (clearContainer) {
            resultContainer.innerHTML = '';
        } else {
            const hr = document.createElement('hr');
            hr.style.margin = "20px 0";
            resultContainer.appendChild(hr);
        }
        
        const existingDownloadButton = document.getElementById('downloadExcelButton');
        if (existingDownloadButton) {
            existingDownloadButton.remove();
        }

        const table = document.createElement('table');
        table.className = 'results-table timesheet-table'; 
        table.setAttribute('aria-label', `Bảng chấm công tháng ${month + 1}/${year}`);

        const caption = table.createCaption();
        caption.textContent = `Bảng chấm công tháng ${month + 1}/${year}`;

        const thead = table.createTHead();
        
        const headerRowDays = thead.insertRow();
        const thId = document.createElement('th');
        thId.scope = 'col';
        thId.rowSpan = 2; 
        thId.textContent = 'ID / TÊN NHÂN VIÊN';
        headerRowDays.appendChild(thId);

        for (let day = 1; day <= daysInMonth; day++) {
            const thDayNumber = document.createElement('th');
            thDayNumber.scope = 'col';
            thDayNumber.textContent = day.toString();
            const dayOfWeek = getDayOfWeek(year, month, day);
            if (dayOfWeek === 0 || dayOfWeek === 6) { // 0 for Sunday, 6 for Saturday
                thDayNumber.classList.add('weekend-header');
            }
            headerRowDays.appendChild(thDayNumber);
        }

        const headerRowDayNames = thead.insertRow();
        for (let day = 1; day <= daysInMonth; day++) {
            const thDayName = document.createElement('th');
            thDayName.scope = 'col';
            thDayName.textContent = getDayOfWeekString(year, month, day);
            const dayOfWeek = getDayOfWeek(year, month, day);
            if (dayOfWeek === 0 || dayOfWeek === 6) {
                thDayName.classList.add('weekend-header');
            }
            headerRowDayNames.appendChild(thDayName);
        }

        const tbody = table.createTBody();
        const uniqueIds = Object.keys(timesheetData).sort();

        if (uniqueIds.length === 0) {
            displayMessage(`Không có dữ liệu chấm công cho tháng ${month + 1}/${year} trong các tệp XML.`, 'info-message', !currentEmployeeData);
            return;
        }

        uniqueIds.forEach(id => {
            const row = tbody.insertRow();
            const cellId = row.insertCell();
            
            let idIsNthSource = false;
            for (let dayIdx = 1; dayIdx <= daysInMonth; dayIdx++) {
                const dayAtt = timesheetData[id]?.[dayIdx];
                if (dayAtt && (dayAtt.S_isNth || dayAtt.C_isNth)) {
                    idIsNthSource = true;
                    break;
                }
            }

            const employeeName = currentEmployeeData?.get(id);
            cellId.textContent = employeeName ? `${employeeName} (${id})` : id;
            if (idIsNthSource) {
                cellId.classList.add('nguoi-thuc-hien-name');
            }


            for (let day = 1; day <= daysInMonth; day++) {
                const cell = row.insertCell();
                cell.classList.add('attendance-mark');
                const dayOfWeek = getDayOfWeek(year, month, day);
                if (dayOfWeek === 0 || dayOfWeek === 6) {
                    cell.classList.add('weekend-data-cell');
                }
                
                const dayAttendance = timesheetData[id]?.[day];
                if (dayAttendance) {
                    let attendanceMark = '';
                    if (dayAttendance.S && dayAttendance.C) {
                        attendanceMark = '+';
                    } else if (dayAttendance.S) {
                        attendanceMark = 'S';
                    } else if (dayAttendance.C) {
                        attendanceMark = 'C';
                    }
                    cell.textContent = attendanceMark;

                    let cellIsNthDerived = false;
                    if (attendanceMark === '+' && (dayAttendance.S_isNth || dayAttendance.C_isNth)) {
                        cellIsNthDerived = true;
                    } else if (attendanceMark === 'S' && dayAttendance.S_isNth) {
                        cellIsNthDerived = true;
                    } else if (attendanceMark === 'C' && dayAttendance.C_isNth) {
                        cellIsNthDerived = true;
                    }

                    if (cellIsNthDerived) {
                        cell.classList.add('nguoi-thuc-hien-date');
                    }


                    const allDayTimes = [...dayAttendance.sTimes, ...dayAttendance.cTimes];
                    if (allDayTimes.length > 0) {
                        allDayTimes.sort((a, b) => a.getTime() - b.getTime());
                        const earliest = allDayTimes[0];
                        const latest = allDayTimes[allDayTimes.length - 1];
                        if (allDayTimes.length === 1) {
                            cell.title = `Thời gian: ${formatTime(earliest)}`;
                        } else {
                            cell.title = `Sớm nhất: ${formatTime(earliest)}\nMuộn nhất: ${formatTime(latest)}`;
                        }
                    }

                } else {
                    cell.textContent = '';
                }
            }
        });
        resultContainer.appendChild(table);

        if (uniqueIds.length > 0 && typeof XLSX !== 'undefined') {
            const downloadButton = document.createElement('button');
            downloadButton.id = 'downloadExcelButton';
            downloadButton.textContent = 'Tải Xuống Bảng Chấm Công (XLSX)';
            downloadButton.onclick = () => {
                try {
                    const filename = `BangChamCong_Thang_${month + 1}_${year}.xlsx`;
                    const excelSheetData = [];

                    // Define styles
                    const thinBorderSide = { style: "thin", color: { rgb: "D3D3D3" } }; 
                    const allBorders = { top: thinBorderSide, bottom: thinBorderSide, left: thinBorderSide, right: thinBorderSide };
                    
                    const headerFontBase = { name: 'Arial', sz: 11, bold: true }; // Base for header fonts
                    const headerFill = { fgColor: { rgb: "4F81BD" } }; // Blue
                    const headerAlign = { horizontal: "center", vertical: "center", wrapText: true };
                    const baseHeaderStyle = { font: {...headerFontBase, color: { rgb: "FFFFFF" }}, fill: headerFill, alignment: headerAlign, border: allBorders };

                    const dayOfWeekFont = { name: 'Arial', sz: 9, bold: false, color: { rgb: "000000" } };
                    const dayOfWeekFill = { fgColor: { rgb: "E0E0E0" } }; // Light gray
                    const baseDayOfWeekStyle = { font: dayOfWeekFont, fill: dayOfWeekFill, alignment: headerAlign, border: allBorders };
                    
                    const idNameFont = { name: 'Arial', sz: 10 };
                    const idNameAlign = { horizontal: "left", vertical: "center" };
                    const baseIdNameStyle = { font: idNameFont, alignment: idNameAlign, border: allBorders };
                    
                    const idNameNthFont = { ...idNameFont, color: { rgb: "FF0000" } }; // Red
                    const baseIdNameNthStyle = { font: idNameNthFont, alignment: idNameAlign, border: allBorders };

                    const markFont = { name: 'Arial', sz: 10, bold: true };
                    const markAlign = { horizontal: "center", vertical: "center" };
                    const baseMarkStyle = { font: markFont, alignment: markAlign, border: allBorders };

                    const markNthFont = { ...markFont, color: { rgb: "FF0000" } }; // Red
                    const baseMarkNthStyle = { font: markNthFont, alignment: markAlign, border: allBorders };
                    
                    const defaultFontForEmptyCells = { name: 'Arial', sz: 10 };
                    const baseEmptyCellStyle = { 
                        font: defaultFontForEmptyCells, 
                        border: allBorders, 
                        alignment: markAlign 
                    };


                    const weekendCellFill = {
                        patternType: "solid", 
                        fgColor: { rgb: "F5F5F5" } // LightGray for weekends
                    };
                    const blackFontColor = { color: { rgb: "000000" } };

                    // Header Row 1: ID/Name (placeholder, will be styled) and Day Numbers
                    const headerRow1 = ["ID / TÊN NHÂN VIÊN"];
                    for (let day = 1; day <= daysInMonth; day++) {
                        headerRow1.push(day);
                    }
                    excelSheetData.push(headerRow1);

                    // Header Row 2: Empty (for merged ID/Name) and Day of Week
                    const headerRow2 = [""]; // Placeholder for merged cell
                    for (let day = 1; day <= daysInMonth; day++) {
                        headerRow2.push(getDayOfWeekString(year, month, day));
                    }
                    excelSheetData.push(headerRow2);
                    
                    // Data Rows
                    uniqueIds.forEach(id => {
                        const dataRow = [];
                        let idIsNthSource = false;
                        for (let dayIdx = 1; dayIdx <= daysInMonth; dayIdx++) {
                            const dayAtt = timesheetData[id]?.[dayIdx];
                            if (dayAtt && (dayAtt.S_isNth || dayAtt.C_isNth)) {
                                idIsNthSource = true;
                                break;
                            }
                        }
                        const employeeName = currentEmployeeData?.get(id);
                        dataRow.push(employeeName ? `${employeeName} (${id})` : id);

                        for (let day = 1; day <= daysInMonth; day++) {
                            const dayAttendance = timesheetData[id]?.[day];
                            let attendanceMark = '';
                            if (dayAttendance) {
                                if (dayAttendance.S && dayAttendance.C) attendanceMark = '+';
                                else if (dayAttendance.S) attendanceMark = 'S';
                                else if (dayAttendance.C) attendanceMark = 'C';
                            }
                            dataRow.push(attendanceMark);
                        }
                        excelSheetData.push(dataRow);
                    });

                    const worksheet = XLSX.utils.aoa_to_sheet(excelSheetData);

                    // Apply styles
                    // Header Row 1 (Day Numbers)
                    worksheet[XLSX.utils.encode_cell({r:0, c:0})].s = baseHeaderStyle; // Main ID/Name header
                    for (let c = 1; c <= daysInMonth; c++) {
                        const dayOfMonth = c; // Day number is column index c (1-based for day)
                        const dayOfWeek = getDayOfWeek(year, month, dayOfMonth);
                        let currentHeaderStyle = {...baseHeaderStyle}; 
                        
                        if (dayOfWeek === 0 || dayOfWeek === 6) { 
                           currentHeaderStyle.fill = weekendCellFill; 
                           currentHeaderStyle.font = { ...baseHeaderStyle.font, ...blackFontColor }; 
                        }
                        worksheet[XLSX.utils.encode_cell({r:0, c:c})].s = currentHeaderStyle;
                    }
                    // Header Row 2 (Day of Week Names)
                    for (let c = 1; c <= daysInMonth; c++) {
                        const dayOfMonth = c;
                        const dayOfWeek = getDayOfWeek(year, month, dayOfMonth);
                        let currentDayNameStyle = {...baseDayOfWeekStyle}; 
                         if (dayOfWeek === 0 || dayOfWeek === 6) { 
                           currentDayNameStyle.fill = weekendCellFill; 
                        }
                        worksheet[XLSX.utils.encode_cell({r:1, c:c})].s = currentDayNameStyle;
                    }

                    // Data Rows Styling
                    uniqueIds.forEach((id, rowIndex) => {
                        const r = rowIndex + 2; // Data starts from Excel row index 2

                        let idIsNthSource = false;
                         for (let dayIdx = 1; dayIdx <= daysInMonth; dayIdx++) {
                            const dayAtt = timesheetData[id]?.[dayIdx];
                            if (dayAtt && (dayAtt.S_isNth || dayAtt.C_isNth)) {
                                idIsNthSource = true;
                                break;
                            }
                        }
                        worksheet[XLSX.utils.encode_cell({r:r, c:0})].s = idIsNthSource ? baseIdNameNthStyle : baseIdNameStyle;
                        
                        for (let day = 1; day <= daysInMonth; day++) {
                            const excelColIdx = day; 
                            const dayOfWeek = getDayOfWeek(year, month, day);
                            const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

                            const dayAttendance = timesheetData[id]?.[day];
                            let cellIsNthDerived = false;
                            // Ensure .v (value) exists before checking it for attendanceMarkValue
                            const cellObject = worksheet[XLSX.utils.encode_cell({r:r, c:excelColIdx})];
                            const attendanceMarkValue = cellObject ? cellObject.v : ''; 

                            if (dayAttendance) {
                                if (attendanceMarkValue === '+' && (dayAttendance.S_isNth || dayAttendance.C_isNth)) cellIsNthDerived = true;
                                else if (attendanceMarkValue === 'S' && dayAttendance.S_isNth) cellIsNthDerived = true;
                                else if (attendanceMarkValue === 'C' && dayAttendance.C_isNth) cellIsNthDerived = true;
                            }
                            
                            let currentCellStyle;
                            if (attendanceMarkValue) { // Check if there's a mark
                                currentCellStyle = cellIsNthDerived ? {...baseMarkNthStyle} : {...baseMarkStyle};
                            } else {
                                currentCellStyle = {...baseEmptyCellStyle};
                            }
                            if (isWeekend) {
                                // Important: Create a new object for the style to avoid unintended mutations
                                currentCellStyle = { ...currentCellStyle, fill: weekendCellFill };
                            }
                             if (cellObject) { // Ensure cell object exists before assigning style
                                cellObject.s = currentCellStyle;
                            }
                        }
                    });
                    
                    // Merges
                    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }]; // For "ID / TÊN NHÂN VIÊN"

                    // Column Widths
                    const colWidths = [{ wch: 35 }]; // ID/Name column
                    for (let i = 0; i < daysInMonth; i++) {
                        colWidths.push({ wch: 4 }); // Day columns
                    }
                    worksheet['!cols'] = colWidths;
                    
                    const rowHeights = [ {hpx: 30}, {hpx: 25} ]; 
                    worksheet['!rows'] = rowHeights;


                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, `Tháng ${month + 1}-${year}`);
                    XLSX.writeFile(workbook, filename);

                } catch (e) {
                    console.error("Lỗi khi tạo tệp Excel:", e);
                    displayMessage("Đã xảy ra lỗi khi cố gắng tạo tệp Excel để tải xuống. Chi tiết: " + e.message, "error-message", true);
                }
            };
            resultContainer.appendChild(downloadButton);
        }
    }

    function displayMessage(message, className = '', append = false) {
        const initialMessagePresent = resultContainer.querySelector('#initialMessage');
        
        if (!append || (initialMessagePresent && resultContainer.childElementCount === 1) ) {
            resultContainer.innerHTML = ''; 
        }
        
        const messageElement = document.createElement('p');
        messageElement.innerHTML = message; 
        if (className) {
            messageElement.className = className;
        }
        
        if (append && resultContainer.childElementCount > 0 && !initialMessagePresent ) {
             if (! (resultContainer.lastChild && resultContainer.lastChild.nodeName === 'BR')) {
                resultContainer.appendChild(document.createElement('br'));
             }
        }
        resultContainer.appendChild(messageElement);
    }

    resetUI();
});
