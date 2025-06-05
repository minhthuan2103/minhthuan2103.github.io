
/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

document.addEventListener('DOMContentLoaded', () => {
    console.log('DOMContentLoaded event fired.');
    console.log('Checking window.XLSX right after DOMContentLoaded:', typeof window.XLSX, window.XLSX);

    // --- State Variables (formerly in index.tsx and tableRenderer.ts) ---
    let employeeData_b = null;
    let uniqueIdsForTable_b = [];
    let draggedItemId_b = null; // For table row dragging

    // --- Constants (formerly in fileUtils.ts) ---
    const MAX_HEADER_SCAN_ROWS_b = 10;

    // --- DOM Utility Functions (originally from domUtils.ts) ---
    function displayMessage_b(
        resultContainer,
        message,
        className = '',
        append = false
    ) {
        const initialMsgElement = resultContainer.querySelector('#initialMessage');
        
        if (!append || (initialMsgElement && resultContainer.childNodes.length === 1 && resultContainer.firstChild === initialMsgElement)) {
            resultContainer.innerHTML = ''; 
        } else if (initialMsgElement && initialMsgElement.parentNode === resultContainer) {
            initialMsgElement.remove();
        }
        
        const messageElement = document.createElement('p');
        messageElement.innerHTML = message; 
        if (className) messageElement.className = className;
        resultContainer.appendChild(messageElement);
    }

    function getRequiredDOMelements_b() {
        const xmlFileInput = document.getElementById('xmlFile');
        const excelFileInput = document.getElementById('excelFile');
        const processFileButton = document.getElementById('processFileButton');
        const resultContainer = document.getElementById('resultContainer');

        if (!xmlFileInput || !excelFileInput || !processFileButton || !resultContainer) {
            console.error('One or more required HTML elements are missing from the DOM.');
            const body = document.body;
            if (body && !resultContainer) {
                try {
                    const p = document.createElement('p');
                    p.className = 'error-message';
                    p.style.cssText = "padding: 10px; background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; margin: 10px;";
                    p.textContent = 'Lỗi nghiêm trọng: Không thể tìm thấy vùng chứa kết quả ("resultContainer").';
                    body.insertBefore(p, body.firstChild);
                } catch (e) {
                    console.error("Fallback DOM manipulation failed", e);
                }
            } else if (resultContainer) { // If resultContainer exists, display error there
                 displayMessage_b(resultContainer, 'Lỗi nghiêm trọng: Thiếu một hoặc nhiều thành phần HTML chính.', 'error-message');
            }
            return null;
        }
        return { 
            xmlFileInput: xmlFileInput, 
            excelFileInput: excelFileInput, 
            processFileButton: processFileButton, 
            resultContainer: resultContainer 
        };
    }
    
    function resetUI_b(xmlFileInput, excelFileInput, processFileButton, resultContainer) {
        resultContainer.innerHTML = ''; 
        const initialMsgP = document.createElement('p');
        initialMsgP.id = 'initialMessage';
        initialMsgP.innerHTML = '1. Tải lên một hoặc nhiều tệp XML chứa dữ liệu chấm công.<br>2. (Tùy chọn) Tải lên tệp Excel danh sách nhân viên.<br>3. Nhấp "Tạo Bảng Chấm Công".';
        resultContainer.appendChild(initialMsgP);
        
        if (xmlFileInput) xmlFileInput.value = '';
        if (excelFileInput) excelFileInput.value = '';
        
        employeeData_b = null;
        uniqueIdsForTable_b = []; 

        if (processFileButton) processFileButton.disabled = true;
    }

    function resetPartialUIForNewFiles_b(resultContainer) {
        const hasResultsTable = !!resultContainer.querySelector('.results-table');
        const hasDownloadButton = !!resultContainer.querySelector('#downloadExcelButton');
        const hasSpecificMessages = !!resultContainer.querySelector('.error-message, .info-message');

        if (hasResultsTable || hasDownloadButton || hasSpecificMessages) {
            resultContainer.innerHTML = ''; 
            const initialMsgP = document.createElement('p');
            initialMsgP.id = 'initialMessage';
            initialMsgP.innerHTML = '1. Tải lên một hoặc nhiều tệp XML chứa dữ liệu chấm công.<br>2. (Tùy chọn) Tải lên tệp Excel danh sách nhân viên.<br>3. Nhấp "Tạo Bảng Chấm Công".';
            resultContainer.appendChild(initialMsgP);
        }
    }

    // --- File Parsing Functions (originally from fileUtils.ts) ---
    async function parseExcelData_b(file, resultContainerForMessages) {
        console.log('Inside parseExcelData_b. Checking window.XLSX:', typeof window.XLSX, window.XLSX);
        return new Promise((resolve) => { 
            const reader = new FileReader();
            reader.onload = (event) => {
                const newEmployeeData = new Map();
                try {
                    const data = event.target?.result;
                    let parsedSuccessfully = false;
                    const canUseXLSX = typeof window.XLSX !== 'undefined' && window.XLSX;
                    const isExcelFileByMimeOrExt = file.type.startsWith('application/vnd.ms-excel') || file.type.startsWith('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') || file.name.toLowerCase().endsWith('.xls') || file.name.toLowerCase().endsWith('.xlsx');
                    const isCsvOrTxtFileByMimeOrExt = file.type === 'text/csv' || file.name.toLowerCase().endsWith('.csv') || file.name.toLowerCase().endsWith('.txt');

                    if (canUseXLSX && isExcelFileByMimeOrExt && data instanceof ArrayBuffer) {
                        const workbook = window.XLSX.read(data, { type: 'array' });
                        const firstSheetName = workbook.SheetNames[0];
                        if (!firstSheetName) {
                            displayMessage_b(resultContainerForMessages, `Tệp Excel "${file.name}" không chứa trang tính (sheet) nào.`, 'error-message', true);
                            resolve(newEmployeeData); return;
                        }
                        const worksheet = workbook.Sheets[firstSheetName];
                        const jsonData = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                        let headerRowIndex = -1, nameColIdx = -1, certColIdx = -1;

                        if (jsonData && jsonData.length > 0) {
                            for (let i = 0; i < Math.min(jsonData.length, MAX_HEADER_SCAN_ROWS_b); i++) {
                                const currentRow = (jsonData[i]).map(h => String(h ?? "").toLowerCase().trim());
                                const tempNameColIdx = currentRow.findIndex(h => typeof h === 'string' && (h.includes('tên nhân viên') || h.includes('ten nhan vien') || h.includes('họ tên') || h.includes('ho ten')));
                                const tempCertColIdx = currentRow.findIndex(h => typeof h === 'string' && (h.includes('chứng chỉ hành nghề') || h.includes('cchn') || h.includes('ma_bac_si') || h.includes('mã bác sĩ')));
                                if (tempNameColIdx !== -1 && tempCertColIdx !== -1) {
                                    headerRowIndex = i; nameColIdx = tempNameColIdx; certColIdx = tempCertColIdx; break;
                                }
                            }
                        }
                        if (headerRowIndex === -1 || nameColIdx === -1 || certColIdx === -1) {
                            displayMessage_b(resultContainerForMessages, `Không tìm thấy cột "Tên nhân viên" và "Chứng chỉ hành nghề" trong ${MAX_HEADER_SCAN_ROWS_b} hàng đầu tiên của tệp Excel "${file.name}". Vui lòng kiểm tra lại tiêu đề cột.`, 'error-message', true);
                            resolve(newEmployeeData); return;
                        }
                        for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                            const row = jsonData[i];
                            if (row && row.length > Math.max(nameColIdx, certColIdx)) {
                                const name = row[nameColIdx]?.trim();
                                const cert = row[certColIdx]?.trim();
                                if (name && cert) newEmployeeData.set(cert, name);
                            }
                        }
                        parsedSuccessfully = true;
                    } else if (isCsvOrTxtFileByMimeOrExt && typeof data === 'string') {
                        const lines = data.split(/\r\n|\n/);
                        let headerRowIndex = -1, nameColIdx = -1, certColIdx = -1;
                        if (lines.length > 0) {
                            for (let i = 0; i < Math.min(lines.length, MAX_HEADER_SCAN_ROWS_b); i++) {
                                const currentRow = lines[i].split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/) 
                                                     .map(h => String(h ?? "").toLowerCase().trim().replace(/^"|"$/g, ''));
                                const tempNameColIdx = currentRow.findIndex(h => typeof h === 'string' && (h.includes('tên nhân viên') || h.includes('ten nhan vien') || h.includes('họ tên') || h.includes('ho ten')));
                                const tempCertColIdx = currentRow.findIndex(h => typeof h === 'string' && (h.includes('chứng chỉ hành nghề') || h.includes('cchn') || h.includes('ma_bac_si') || h.includes('mã bác sĩ')));
                                if (tempNameColIdx !== -1 && tempCertColIdx !== -1) {
                                    headerRowIndex = i; nameColIdx = tempNameColIdx; certColIdx = tempCertColIdx; break;
                                }
                            }
                        }
                        if (headerRowIndex === -1 || nameColIdx === -1 || certColIdx === -1) {
                             displayMessage_b(resultContainerForMessages, `Không tìm thấy cột "Tên nhân viên" và "Chứng chỉ hành nghề" trong ${MAX_HEADER_SCAN_ROWS_b} dòng đầu tiên của tệp CSV/TXT "${file.name}". Vui lòng kiểm tra lại tiêu đề cột.`, 'error-message', true);
                             resolve(newEmployeeData); return;
                        }
                        for (let i = headerRowIndex + 1; i < lines.length; i++) {
                             if (!lines[i] || lines[i].trim() === '') continue; 
                            const columns = lines[i].split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/)
                                              .map(c => c.trim().replace(/^"|"$/g, ''));
                            if (columns.length > Math.max(nameColIdx, certColIdx)) {
                                const name = columns[nameColIdx]?.trim();
                                const cert = columns[certColIdx]?.trim();
                                if (name && cert) newEmployeeData.set(cert, name);
                            }
                        }
                        parsedSuccessfully = true;
                    }

                    if (!parsedSuccessfully && (isExcelFileByMimeOrExt || isCsvOrTxtFileByMimeOrExt)) {
                        displayMessage_b(resultContainerForMessages, `Không thể trích xuất dữ liệu nhân viên từ ${file.name}. Định dạng có thể không được hỗ trợ hoặc tệp bị hỏng.`, 'info-message', true);
                    } else if (newEmployeeData.size > 0) {
                        displayMessage_b(resultContainerForMessages, `Đã tải thành công ${newEmployeeData.size} mục nhân viên từ tệp ${file.name}.`, 'info-message', true);
                    } else if (parsedSuccessfully && newEmployeeData.size === 0) {
                        displayMessage_b(resultContainerForMessages, `Không tìm thấy dữ liệu nhân viên hợp lệ trong ${file.name}, mặc dù tệp đã được xử lý.`, 'info-message', true);
                    }
                    resolve(newEmployeeData);

                } catch (e) {
                    console.error("Lỗi phân tích cú pháp tệp Excel/CSV: ", e);
                    displayMessage_b(resultContainerForMessages, `Lỗi nghiêm trọng khi phân tích cú pháp tệp ${file.name}. Chi tiết: ${e.message}`, 'error-message', true);
                    resolve(newEmployeeData); 
                }
            };
            reader.onerror = () => {
                displayMessage_b(resultContainerForMessages, `Lỗi không thể đọc tệp nhân viên: ${file.name}.`, 'error-message', true);
                resolve(new Map());
            };

            const canUseXLSXLib = typeof window.XLSX !== 'undefined' && window.XLSX;
            const isExcelFile = file.name.toLowerCase().endsWith('.xls') || file.name.toLowerCase().endsWith('.xlsx') || file.type.startsWith('application/vnd.ms-excel') || file.type.startsWith('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            const isCsvTxtFile = file.name.toLowerCase().endsWith('.csv') || file.name.toLowerCase().endsWith('.txt') || file.type === 'text/csv';

            if (canUseXLSXLib && isExcelFile) {
                reader.readAsArrayBuffer(file);
            } else if (isCsvTxtFile) {
                reader.readAsText(file);
            } else {
                 let message = `Loại tệp không được hỗ trợ cho danh sách nhân viên: "${file.name}". Chỉ chấp nhận .xls, .xlsx, .csv, hoặc .txt.`;
                 if (isExcelFile && !canUseXLSXLib) {
                    message = `Thư viện XLSX (SheetJS) cần thiết để xử lý các tệp Excel "${file.name}" nhưng chưa được tải. Hiện tại, chỉ có thể xử lý các tệp .csv hoặc .txt mà không cần thư viện này.`;
                 }
                 displayMessage_b(resultContainerForMessages, message, 'error-message', true);
                 resolve(new Map());
            }
        });
    }

    async function parseXmlFiles_b(xmlFiles, resultContainerForMessages) {
        const xmlDocuments = [];
        const fileErrors = [];
        const fileReadPromises = [];

        for (const file of Array.from(xmlFiles)) {
            if (file.type !== 'text/xml' && !file.name.toLowerCase().endsWith('.xml')) {
                fileErrors.push({ fileName: file.name, message: `Loại tệp không hợp lệ. Chỉ được phép chọn tệp XML.` });
                continue; 
            }
            fileReadPromises.push(new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = (event) => resolve({ content: event.target?.result, name: file.name });
                reader.onerror = () => resolve({ errorMsg: `Lỗi đọc tệp hệ thống.`, name: file.name });
                reader.readAsText(file);
            }));
        }

        const fileReadResults = await Promise.all(fileReadPromises);

        for (const result of fileReadResults) {
            if (result.errorMsg) { // Check if errorMsg property exists
                fileErrors.push({ fileName: result.name, message: result.errorMsg });
                continue;
            }

            const { content: fileContent, name: fileName } = result;
            if (!fileContent || String(fileContent).trim() === '') {
                fileErrors.push({ fileName, message: `Nội dung tệp XML trống hoặc không hợp lệ.` });
                continue;
            }

            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(String(fileContent), 'text/xml');
            const parserError = xmlDoc.getElementsByTagName('parsererror');

            if (parserError.length > 0) {
                let errorMessageText = `Tệp XML không hợp lệ hoặc có lỗi cấu trúc.`;
                const errorDetails = parserError[0].textContent;
                if (errorDetails) {
                   const specificErrorMatch = errorDetails.match(/Error: (.*?) at line \d+ column \d+/s) || 
                                              errorDetails.match(/^(.*?)\nLocation:/s) ||
                                              errorDetails.match(/<sourcetext>.*?<\/sourcetext>\s*(.*)/s);
                   if (specificErrorMatch?.[1] && specificErrorMatch[1].trim().length > 5) {
                        errorMessageText = `Lỗi phân tích XML: ${specificErrorMatch[1].trim()}`;
                   } else if (errorDetails.includes("no element found")) {
                        errorMessageText = `Lỗi phân tích XML: Không tìm thấy phần tử nào, tệp có thể trống hoặc không phải XML.`;
                   }
                }
                fileErrors.push({ fileName, message: errorMessageText });
            } else {
                xmlDocuments.push(xmlDoc);
            }
        }
        return { xmlDocuments, fileErrors };
    }

    // --- Timesheet Logic Functions (originally from timesheetLogic.ts) ---
    function parseNgayYl_b(ngayYlStr) {
        if (!ngayYlStr || ngayYlStr.length !== 12) return null;
        const year = parseInt(ngayYlStr.substring(0, 4), 10);
        const month = parseInt(ngayYlStr.substring(4, 6), 10) - 1;
        const day = parseInt(ngayYlStr.substring(6, 8), 10);
        const hours = parseInt(ngayYlStr.substring(8, 10), 10);
        const minutes = parseInt(ngayYlStr.substring(10, 12), 10);

        if (isNaN(year) || isNaN(month) || isNaN(day) || isNaN(hours) || isNaN(minutes)) return null;
        if (month < 0 || month > 11 || day < 1 || day > 31 || hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
        
        const date = new Date(year, month, day, hours, minutes);
        if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day || 
            date.getHours() !== hours || date.getMinutes() !== minutes) return null;
        return date;
    }

    function formatTime_b(date) {
        const hours = date.getHours().toString().padStart(2, '0');
        const minutes = date.getMinutes().toString().padStart(2, '0');
        return `${hours}:${minutes}`;
    }

    function getDayOfWeek_b(year, month, day) { // month is 0-indexed
        return new Date(year, month, day).getDay(); 
    }

    function getDayOfWeekString_b(year, month, day) { // month is 0-indexed
        const dayIndex = getDayOfWeek_b(year, month, day);
        return ["CN", "T2", "T3", "T4", "T5", "T6", "T7"][dayIndex];
    }

    function processXmlDataToTimesheet_b(xmlDocuments) {
        const allEntries = [];
        const processElementsFromDoc = (xmlDoc, entries) => {
            const parentTagNames = ['CHI_TIET_THUOC', 'CHI_TIET_DVKT'];
            parentTagNames.forEach(parentTagName => {
                const elements = xmlDoc.getElementsByTagName(parentTagName);
                for (const el of Array.from(elements)) {
                    const maBacSiEl = el.getElementsByTagName('MA_BAC_SI')[0];
                    const ngayYlEl = el.getElementsByTagName('NGAY_YL')[0];
                    if (maBacSiEl?.textContent?.trim() && ngayYlEl?.textContent?.trim()) {
                        const dateTime = parseNgayYl_b(ngayYlEl.textContent.trim());
                        if (dateTime) entries.push({ id: maBacSiEl.textContent.trim(), dateTime, isNthEntry: false });
                    }
                    const nguoiThucHienEl = el.getElementsByTagName('NGUOI_THUC_HIEN')[0];
                    const ngayThYlEl = el.getElementsByTagName('NGAY_TH_YL')[0];
                    if (nguoiThucHienEl?.textContent?.trim() && ngayThYlEl?.textContent?.trim()) {
                        const dateTime = parseNgayYl_b(ngayThYlEl.textContent.trim());
                        if (dateTime) {
                            const ids = nguoiThucHienEl.textContent.trim().split(';').map(s => s.trim()).filter(s => s);
                            ids.forEach(id => entries.push({ id, dateTime, isNthEntry: true }));
                        }
                    }
                }
            });
        };

        for (const xmlDoc of xmlDocuments) processElementsFromDoc(xmlDoc, allEntries);

        if (allEntries.length === 0) return { timesheetData: {}, allEntriesCount: 0 };

        allEntries.sort((a, b) => a.dateTime.getTime() - b.dateTime.getTime());
        const firstEntryDate = allEntries[0].dateTime;
        const targetYear = firstEntryDate.getFullYear();
        const targetMonth = firstEntryDate.getMonth();
        const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();
        const timesheetData = {};

        for (const entry of allEntries) {
            const { id, dateTime, isNthEntry } = entry;
            if (dateTime.getFullYear() !== targetYear || dateTime.getMonth() !== targetMonth) continue;
            const day = dateTime.getDate();
            const timeValue = dateTime.getHours() * 100 + dateTime.getMinutes();
            if (!timesheetData[id]) timesheetData[id] = {};
            if (!timesheetData[id][day]) timesheetData[id][day] = { S: false, C: false, sTimes: [], cTimes: [], S_isNth: false, C_isNth: false };
            const dayAtt = timesheetData[id][day];
            if (timeValue >= 700 && timeValue < 1130) { 
                if (!dayAtt.S) dayAtt.S = true;
                if (isNthEntry) dayAtt.S_isNth = true;
                dayAtt.sTimes.push(new Date(dateTime));
            } else if (timeValue >= 1330 && timeValue < 1700) { 
                if (!dayAtt.C) dayAtt.C = true;
                if (isNthEntry) dayAtt.C_isNth = true;
                dayAtt.cTimes.push(new Date(dateTime));
            }
        }
        return { timesheetData, year: targetYear, month: targetMonth, daysInMonth, allEntriesCount: allEntries.length };
    }

    // --- Table Rendering Functions (originally from tableRenderer.ts) ---
    function populateTableBody_b(tbody, timesheetData, year, month, daysInMonth, currentEmployeeData, currentOrderedIds, onRowOrderChangeCallback) {
        tbody.innerHTML = '';
        currentOrderedIds.forEach(id => {
            const row = tbody.insertRow();
            row.draggable = true;
            row.dataset.id = id;
            row.setAttribute('aria-grabbed', 'false');

            row.addEventListener('dragstart', (event) => {
                draggedItemId_b = id;
                if (event.dataTransfer) {
                    event.dataTransfer.effectAllowed = 'move';
                    event.dataTransfer.setData('text/plain', id);
                }
                const currentTargetRow = event.currentTarget;
                currentTargetRow.classList.add('dragging');
                currentTargetRow.setAttribute('aria-grabbed', 'true');
                tbody.querySelectorAll('tr').forEach(tr => {
                    if (tr !== currentTargetRow) tr.classList.add('drop-candidate');
                });
            });
            row.addEventListener('dragend', (event) => {
                const currentTargetRow = event.currentTarget;
                currentTargetRow.classList.remove('dragging');
                currentTargetRow.setAttribute('aria-grabbed', 'false');
                tbody.querySelectorAll('tr').forEach(tr => {
                    tr.classList.remove('drag-over-target');
                    tr.classList.remove('drop-candidate');
                });
                draggedItemId_b = null;
            });
            row.addEventListener('dragover', (event) => {
                event.preventDefault();
                const targetRow = event.currentTarget;
                if (!draggedItemId_b || targetRow.dataset.id === draggedItemId_b) return;
                if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';
                targetRow.classList.add('drag-over-target');
            });
            row.addEventListener('dragleave', (event) => {
                event.currentTarget.classList.remove('drag-over-target');
            });
            row.addEventListener('drop', (event) => {
                event.preventDefault();
                const targetRow = event.currentTarget;
                targetRow.classList.remove('drag-over-target');
                const droppedOnItemId = targetRow.dataset.id;
                if (draggedItemId_b && droppedOnItemId && draggedItemId_b !== droppedOnItemId) {
                    const newOrder = [...currentOrderedIds];
                    const draggedIndex = newOrder.indexOf(draggedItemId_b);
                    const droppedOnIndex = newOrder.indexOf(droppedOnItemId);
                    if (draggedIndex !== -1 && droppedOnIndex !== -1) {
                        const [movedItem] = newOrder.splice(draggedIndex, 1);
                        newOrder.splice(droppedOnIndex, 0, movedItem);
                        onRowOrderChangeCallback(newOrder);
                    }
                }
                draggedItemId_b = null;
            });

            const cellId = row.insertCell();
            const employeeName = currentEmployeeData?.get(id);
            cellId.textContent = employeeName ? `${employeeName} (${id})` : id;

            for (let day = 1; day <= daysInMonth; day++) {
                const cell = row.insertCell();
                cell.classList.add('attendance-mark');
                const dayOfWeekNum = getDayOfWeek_b(year, month, day);
                if (dayOfWeekNum === 0 || dayOfWeekNum === 6) cell.classList.add('weekend-data-cell');
                const dayAttendance = timesheetData[id]?.[day];
                if (dayAttendance) {
                    let attendanceMark = '';
                    if (dayAttendance.S && dayAttendance.C) attendanceMark = '+';
                    else if (dayAttendance.S) attendanceMark = 'S';
                    else if (dayAttendance.C) attendanceMark = 'C';
                    cell.textContent = attendanceMark;
                    const allDayTimes = [...dayAttendance.sTimes, ...dayAttendance.cTimes].sort((a,b) => a.getTime() - b.getTime());
                    if (allDayTimes.length > 0) {
                        const earliest = allDayTimes[0];
                        const latest = allDayTimes[allDayTimes.length - 1];
                        cell.title = allDayTimes.length === 1 
                            ? `Thời gian: ${formatTime_b(earliest)}` 
                            : `Sớm nhất: ${formatTime_b(earliest)}\nMuộn nhất: ${formatTime_b(latest)}`;
                    }
                } else {
                    cell.textContent = '';
                }
            }
        });
    }

    function displayTimesheetTable_b(resultContainer, timesheetData, year, month, daysInMonth, currentEmployeeData, currentOrderedIds, onRowOrderChangeCallback) {
        let table = resultContainer.querySelector('.results-table.timesheet-table');
        let tbody;

        const hrElement = resultContainer.querySelector('hr.results-separator');
        if (!hrElement && (resultContainer.querySelector('.info-message') || resultContainer.querySelector('.error-message'))) {
            const hr = document.createElement('hr');
            hr.className = 'results-separator';
            hr.style.margin = "20px 0";
            const firstMessage = resultContainer.querySelector('.info-message, .error-message');
            if (firstMessage && firstMessage.nextSibling) resultContainer.insertBefore(hr, firstMessage.nextSibling);
            else if (firstMessage) resultContainer.appendChild(hr);
        }

        if (!table) {
            const initialMsg = resultContainer.querySelector('#initialMessage');
            if (initialMsg && initialMsg.parentNode === resultContainer) initialMsg.remove();
            table = document.createElement('table');
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
                if ([0,6].includes(getDayOfWeek_b(year, month, day))) thDayNumber.classList.add('weekend-header');
                headerRowDays.appendChild(thDayNumber);
            }
            const headerRowDayNames = thead.insertRow();
            for (let day = 1; day <= daysInMonth; day++) {
                const thDayName = document.createElement('th');
                thDayName.scope = 'col';
                thDayName.textContent = getDayOfWeekString_b(year, month, day);
                if ([0,6].includes(getDayOfWeek_b(year, month, day))) thDayName.classList.add('weekend-header');
                headerRowDayNames.appendChild(thDayName);
            }
            tbody = table.createTBody();
            resultContainer.appendChild(table); 
        } else {
            tbody = table.tBodies[0] || table.createTBody();
            const caption = table.caption;
            if (caption) caption.textContent = `Bảng chấm công tháng ${month + 1}/${year}`;
        }

        if (currentOrderedIds.length === 0) {
            if (table && table.parentNode === resultContainer) {
                const noDataMsgExists = Array.from(resultContainer.querySelectorAll('.info-message')).some(el => el.textContent?.includes("Không có dữ liệu chấm công"));
                if (!noDataMsgExists) displayMessage_b(resultContainer, `Không có dữ liệu chấm công cho tháng ${month + 1}/${year}.`, 'info-message', !currentEmployeeData);
                table.remove();
            } else if (!resultContainer.querySelector('.info-message, .error-message')) {
                displayMessage_b(resultContainer, `Không có dữ liệu chấm công cho tháng ${month + 1}/${year}.`, 'info-message', !currentEmployeeData);
            }
            return; 
        }
        populateTableBody_b(tbody, timesheetData, year, month, daysInMonth, currentEmployeeData, currentOrderedIds, onRowOrderChangeCallback);
    }

    // --- Excel Export Function (originally from excelExport.ts) ---
    function setupExcelDownloadButton_b(button, currentOrderedIds, timesheetData, year, month, daysInMonth, currentEmployeeData, resultContainerForMessages) {
        button.onclick = () => {
            if (typeof window.XLSX === 'undefined' || !window.XLSX) {
                displayMessage_b(resultContainerForMessages, "Lỗi: Không thể tạo tệp Excel. Thư viện XLSX (SheetJS) cần thiết chưa được tải hoặc không khả dụng...", "error-message", true);
                button.disabled = true;
                return;
            }
            try {
                const filename = `BangChamCong_Thang_${month + 1}_${year}.xlsx`;
                const excelSheetData = [];
                const thinBorderStyle = { style: "thin", color: { rgb: "E0E0E0" } };
                const allCellBorders = { top: thinBorderStyle, bottom: thinBorderStyle, left: thinBorderStyle, right: thinBorderStyle };
                const centerAlignment = { horizontal: "center", vertical: "center", wrapText: false };
                const centerWrapAlignment = { horizontal: "center", vertical: "center", wrapText: true };
                const leftAlignment = { horizontal: "left", vertical: "center", wrapText: false };
                const dayNumberHeaderFont = { name: 'Arial', sz: 10, bold: true, color: { rgb: "495057" } }; 
                const dayNumberHeaderFill = { patternType: "solid", fgColor: { rgb: "E9ECEF" } };
                const baseDayNumberHeaderStyle = { font: dayNumberHeaderFont, fill: dayNumberHeaderFill, alignment: centerWrapAlignment, border: allCellBorders };
                const dayNameHeaderFont = { name: 'Arial', sz: 9, bold: false, color: { rgb: "555555" } };
                const dayNameHeaderFill = { patternType: "solid", fgColor: { rgb: "F8F9FA" } };
                const baseDayNameHeaderStyle = { font: dayNameHeaderFont, fill: dayNameHeaderFill, alignment: centerWrapAlignment, border: allCellBorders };
                const weekendHeaderSharedFontColor = { color: { rgb: "333333" } };
                const weekendHeaderSharedFill = { patternType: "solid", fgColor: { rgb: "F0F2F5" } };
                const idNameCellFont = { name: 'Arial', sz: 10, bold: false };
                const baseIdNameCellStyle = { font: idNameCellFont, alignment: leftAlignment, border: allCellBorders };
                const markCellFont = { name: 'Arial', sz: 10, bold: true };
                const baseMarkCellStyle = { font: markCellFont, alignment: centerAlignment, border: allCellBorders };
                const emptyDataCellFont = { name: 'Arial', sz: 10, bold: false };
                const baseEmptyDataCellStyle = { font: emptyDataCellFont, alignment: centerAlignment, border: allCellBorders };
                const evenDataRowFill = { patternType: "solid", fgColor: { rgb: "F8F9FA" } };
                const weekendDataCellFill = { patternType: "solid", fgColor: { rgb: "F0F2F5" } };

                const idNameMergedHeaderStyle = JSON.parse(JSON.stringify(baseDayNumberHeaderStyle));
                const headerRow1Values = [{v: "ID / TÊN NHÂN VIÊN", s: idNameMergedHeaderStyle}];
                for (let day = 1; day <= daysInMonth; day++) {
                    let dayStyle = JSON.parse(JSON.stringify(baseDayNumberHeaderStyle));
                    if ([0,6].includes(getDayOfWeek_b(year, month, day))) {
                        dayStyle.fill = { ...weekendHeaderSharedFill }; 
                        if(dayStyle.font) dayStyle.font.color = weekendHeaderSharedFontColor.color;
                        else dayStyle.font = { ...dayNumberHeaderFont, ...weekendHeaderSharedFontColor };
                    }
                    headerRow1Values.push({v: day, s: dayStyle});
                }
                excelSheetData.push(headerRow1Values);
                
                const emptyCellUnderMergedHeaderStyle = JSON.parse(JSON.stringify(baseDayNameHeaderStyle)); // Match style of row
                const headerRow2Values = [{v: "", s: emptyCellUnderMergedHeaderStyle }]; 
                for (let day = 1; day <= daysInMonth; day++) {
                    let dayNameStyle = JSON.parse(JSON.stringify(baseDayNameHeaderStyle));
                     if ([0,6].includes(getDayOfWeek_b(year, month, day))) {
                        dayNameStyle.fill = { ...weekendHeaderSharedFill };
                         if(dayNameStyle.font) dayNameStyle.font.color = weekendHeaderSharedFontColor.color;
                         else dayNameStyle.font = { ...dayNameHeaderFont, ...weekendHeaderSharedFontColor };
                    }
                    headerRow2Values.push({v: getDayOfWeekString_b(year, month, day), s: dayNameStyle});
                }
                excelSheetData.push(headerRow2Values);
                
                currentOrderedIds.forEach((id, dataRowZeroIndexed) => {
                    const dataRowValues = [];
                    const employeeName = currentEmployeeData?.get(id);
                    const idNameDisplayValue = employeeName ? `${employeeName} (${id})` : id;
                    const currentIdNameCellStyle = JSON.parse(JSON.stringify(baseIdNameCellStyle));
                    const isEvenVisualDataRow = dataRowZeroIndexed % 2 === 1; 
                    if (isEvenVisualDataRow) currentIdNameCellStyle.fill = { ...(currentIdNameCellStyle.fill || {}), ...evenDataRowFill };
                    dataRowValues.push({v: idNameDisplayValue, s: currentIdNameCellStyle });

                    for (let day = 1; day <= daysInMonth; day++) {
                        const dayAttendance = timesheetData[id]?.[day];
                        let attendanceMark = '';
                        if (dayAttendance) {
                            if (dayAttendance.S && dayAttendance.C) attendanceMark = '+';
                            else if (dayAttendance.S) attendanceMark = 'S';
                            else if (dayAttendance.C) attendanceMark = 'C';
                        }
                        const cellIsMarked = attendanceMark !== '';
                        const baseStyleForCell = cellIsMarked ? baseMarkCellStyle : baseEmptyDataCellStyle;
                        let currentDayCellStyle = JSON.parse(JSON.stringify(baseStyleForCell));
                        if (isEvenVisualDataRow) currentDayCellStyle.fill = { ...(currentDayCellStyle.fill || {}), ...evenDataRowFill };
                        if ([0,6].includes(getDayOfWeek_b(year, month, day))) currentDayCellStyle.fill = { ...(currentDayCellStyle.fill || {}), ...weekendDataCellFill };
                        dataRowValues.push({v: attendanceMark, s: currentDayCellStyle});
                    }
                    excelSheetData.push(dataRowValues);
                });

                const worksheet = window.XLSX.utils.aoa_to_sheet([]); 
                window.XLSX.utils.sheet_add_aoa(worksheet, excelSheetData, {origin: "A1", cellStyles: true});
                if (!worksheet['!merges']) worksheet['!merges'] = [];
                worksheet['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } });
                if (worksheet['A1']) worksheet['A1'].s = idNameMergedHeaderStyle;

                worksheet['!cols'] = [{ wch: 35 }, ...Array(daysInMonth).fill({ wch: 5.5 })];
                worksheet['!rows'] = [ {hpx: 22, level:0}, {hpx: 20, level:0} ];
                currentOrderedIds.forEach((_, idx) => {
                    if (!worksheet['!rows']) worksheet['!rows'] = [];
                    worksheet['!rows'][idx + 2] = { hpx: 18, level:0 };
                });

                const workbook = window.XLSX.utils.book_new();
                window.XLSX.utils.book_append_sheet(workbook, worksheet, `Tháng ${month + 1}-${year}`);
                window.XLSX.writeFile(workbook, filename);
            } catch (e) {
                displayMessage_b(resultContainerForMessages, "Lỗi trong quá trình tạo tệp Excel: " + e.message, "error-message", true);
                console.error("Error during Excel generation:", e);
            }
        };
    }

    // --- Main Application Logic (adapted from index.tsx) ---
    const elements = getRequiredDOMelements_b();
    if (!elements) {
        // Error already handled by getRequiredDOMelements_b
        return;
    }
    const { xmlFileInput, excelFileInput, processFileButton, resultContainer } = elements;
    
    async function handleFileProcessing_main() {
        const initialMsgP = document.createElement('p');
        initialMsgP.id = 'initialMessage';
        initialMsgP.innerHTML = '1. Tải lên một hoặc nhiều tệp XML...<br>2. (Tùy chọn) Tải lên tệp Excel...<br>3. Nhấp "Tạo Bảng Chấm Công".';
        resultContainer.innerHTML = ''; 
        resultContainer.appendChild(initialMsgP);
        uniqueIdsForTable_b = [];

        const xmlFiles = xmlFileInput.files;
        if (!xmlFiles || xmlFiles.length === 0) {
            displayMessage_b(resultContainer, 'Vui lòng chọn một hoặc nhiều tệp XML trước.', 'error-message');
            return;
        }

        displayMessage_b(resultContainer, 'Đang xử lý các tệp...', 'info-message', true);
        processFileButton.disabled = true;
        xmlFileInput.disabled = true;
        excelFileInput.disabled = true;

        let currentEmployeeDataForRun = null;
        if (excelFileInput.files && excelFileInput.files.length > 0) {
            try {
                currentEmployeeDataForRun = await parseExcelData_b(excelFileInput.files[0], resultContainer);
                employeeData_b = currentEmployeeDataForRun; 
            } catch (error) {
                displayMessage_b(resultContainer, `Lỗi không mong muốn khi xử lý tệp nhân viên: ${error.message}`, 'error-message', true);
            }
        } else {
            employeeData_b = null; 
        }

        try {
            const { xmlDocuments, fileErrors } = await parseXmlFiles_b(xmlFiles, resultContainer);
            let infoMessageShown = !!employeeData_b;

            if (fileErrors.length > 0) {
                fileErrors.forEach(err => displayMessage_b(resultContainer, `Lỗi tệp ${err.fileName}: ${err.message}`, 'error-message', true));
                infoMessageShown = true;
            }
            
            if (xmlDocuments.length === 0) {
                if (!infoMessageShown) {
                     displayMessage_b(resultContainer, 'Không có tệp XML hợp lệ nào được phân tích hoặc không tìm thấy dữ liệu.', 
                                   fileErrors.length > 0 ? 'error-message' : 'info-message', false);
                } else if (fileErrors.length === 0) {
                     displayMessage_b(resultContainer, 'Không tìm thấy tệp XML hợp lệ để xử lý.', 'info-message', true);
                }
                return; 
            }

            const { timesheetData, year, month, daysInMonth, allEntriesCount } = processXmlDataToTimesheet_b(xmlDocuments);

            if (allEntriesCount === 0) {
                displayMessage_b(resultContainer, 'Không tìm thấy dữ liệu chấm công hợp lệ trong các tệp XML đã xử lý.', 'info-message', !employeeData_b);
                return;
            }
            if (year === undefined || month === undefined || daysInMonth === undefined) {
                displayMessage_b(resultContainer, 'Không thể xác định tháng và năm từ dữ liệu chấm công.', 'error-message', !employeeData_b);
                return;
            }
            
            uniqueIdsForTable_b = Object.keys(timesheetData).sort();

            const updateDisplayedTableAndDownloadButton_logic = () => {
                displayTimesheetTable_b(
                    resultContainer, timesheetData, year, month, daysInMonth,
                    employeeData_b, uniqueIdsForTable_b,
                    (newOrder) => { 
                        uniqueIdsForTable_b = newOrder;
                        updateDisplayedTableAndDownloadButton_logic(); 
                    }
                );
                
                let downloadButton = resultContainer.querySelector('#downloadExcelButton');
                if (!downloadButton && uniqueIdsForTable_b.length > 0 && typeof window.XLSX !== 'undefined') {
                    downloadButton = document.createElement('button');
                    downloadButton.id = 'downloadExcelButton';
                    downloadButton.textContent = 'Tải Xuống Bảng Chấm Công (XLSX)';
                    const tableEl = resultContainer.querySelector('.results-table');
                    if (tableEl && tableEl.nextSibling) resultContainer.insertBefore(downloadButton, tableEl.nextSibling);
                    else if (tableEl) resultContainer.appendChild(downloadButton);
                    else resultContainer.appendChild(downloadButton);
                }
                
                if (downloadButton) {
                    if (uniqueIdsForTable_b.length > 0 && typeof window.XLSX !== 'undefined') {
                        setupExcelDownloadButton_b(downloadButton, uniqueIdsForTable_b, timesheetData, year, month, daysInMonth, employeeData_b, resultContainer);
                        downloadButton.style.display = '';
                    } else {
                        downloadButton.style.display = 'none';
                    }
                }

                const xlsxMissingErrorMsg = resultContainer.querySelector('#xlsxMissingError');
                if (uniqueIdsForTable_b.length > 0 && typeof window.XLSX === 'undefined') {
                    if (!xlsxMissingErrorMsg) {
                        const errorP = document.createElement('p');
                        errorP.id = 'xlsxMissingError';
                        errorP.className = 'error-message';
                        errorP.textContent = 'Lỗi: Không thể tạo nút tải xuống Excel do thư viện XLSX (SheetJS) chưa sẵn sàng.';
                        const tableEl = resultContainer.querySelector('.results-table');
                        if (tableEl && tableEl.nextSibling) resultContainer.insertBefore(errorP, tableEl.nextSibling);
                        else if (tableEl) resultContainer.appendChild(errorP);
                        else resultContainer.appendChild(errorP);
                    }
                } else if (xlsxMissingErrorMsg) {
                    xlsxMissingErrorMsg.remove();
                }
            };
            
            updateDisplayedTableAndDownloadButton_logic();

        } catch (error) {
            const errorMessage = (error instanceof Error) ? error.message : 'Lỗi không xác định khi xử lý tệp.';
            const appendError = !!resultContainer.querySelector('.info-message, .error-message, .results-table');
            displayMessage_b(resultContainer, `Lỗi nghiêm trọng trong quá trình xử lý: ${errorMessage}`, 'error-message', appendError);
        } finally {
            processFileButton.disabled = false;
            xmlFileInput.disabled = false;
            excelFileInput.disabled = false;
        }
    }
    
    processFileButton.addEventListener('click', handleFileProcessing_main);
    xmlFileInput.addEventListener('change', () => {
        resetPartialUIForNewFiles_b(resultContainer);
        processFileButton.disabled = !xmlFileInput.files || xmlFileInput.files.length === 0;
    });
    excelFileInput.addEventListener('change', () => {
        employeeData_b = null; 
        resetPartialUIForNewFiles_b(resultContainer); 
    });
    
    resetUI_b(xmlFileInput, excelFileInput, processFileButton, resultContainer);
});
