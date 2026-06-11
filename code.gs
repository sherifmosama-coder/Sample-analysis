// Complete CODE.gs - With External Spreadsheet Copy for Portal A
function doGet() {
  return HtmlService.createTemplateFromFile('INDEX')
    .evaluate()
    .setTitle('نظام إدارة العينات')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// --- UNIFIED RBAC AUTHENTICATION ---
function loginUser(passcode, requestedPortal) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('App_Users');
    
    if (!usersSheet) {
      throw new Error('Sheet "App_Users" not found. Please create it with the specified columns.');
    }

    const lastRow = usersSheet.getLastRow();
    if (lastRow < 2) return { valid: false, message: 'No users configured in the system.' };

    // Fetch Columns A through F
    const data = usersSheet.getRange(2, 1, lastRow - 1, 6).getValues(); 
    
    // Map portal ID to the respective Checkbox Column Index (0-based array)
    // Col C (Index 2) = Portal A
    // Col D (Index 3) = Portal B
    // Col E (Index 4) = PROD
    // Col F (Index 5) = TDS
    const portalColMap = {
      'A': 2,
      'B': 3,
      'PROD': 4,
      'TDS': 5
    };
    const targetColIndex = portalColMap[requestedPortal];

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][1]) === String(passcode)) { // Found the PIN in Col B
        const userName = data[i][0]; // Name in Col A
        const hasAccess = data[i][targetColIndex] === true; // Checkbox value
        
        if (hasAccess) {
          return { valid: true, userName: userName };
        } else {
          return { valid: false, message: 'ليس لديك صلاحية للدخول إلى هذه البوابة.' };
        }
      }
    }
    return { valid: false, message: 'الرمز السري غير صحيح.' };
  } catch (error) {
    return { valid: false, message: 'System Error: ' + error.message };
  }
}

// --- HEARTBEAT & ONLINE TRACKING ---
function updateHeartbeat(userName, portalName) {
  if (!userName) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('App_Users');
    if (!usersSheet || usersSheet.getLastRow() < 2) return;
    
    const data = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userName) {
        const rowIndex = i + 2;
        const timestamp = new Date();
        // Col G = Last Ping, Col H = Location
        usersSheet.getRange(rowIndex, 7).setValue(timestamp);
        usersSheet.getRange(rowIndex, 8).setValue(portalName);
        break;
      }
    }
  } catch (e) {
    // Silent fail to avoid disrupting user experience
  }
}

function getOnlineUsers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('App_Users');
    if (!usersSheet || usersSheet.getLastRow() < 2) return [];

    // Fetch Name (A) to Location (H)
    const data = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 8).getValues();
    const now = new Date().getTime();
    const onlineUsers = [];

    for (let i = 0; i < data.length; i++) {
      const userName = String(data[i][0] || '').trim();
      const lastPingDate = data[i][6]; // Col G
      const location = String(data[i][7] || '').trim(); // Col H

      if (userName && lastPingDate instanceof Date) {
        const diffMs = now - lastPingDate.getTime();
        // Consider "Online" if pinged within the last 2 minutes (120,000 ms)
        if (diffMs <= 120000 && location && !location.includes('Idle')) {
           // Get first 2 characters for initials
           const initials = userName.substring(0, 2).toUpperCase();
           onlineUsers.push({
             name: userName,
             initials: initials,
             location: location
           });
        }
      }
    }
    return onlineUsers;
  } catch (e) {
    return [];
  }
}
// Get material options from index sheet
function getMaterialOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = ss.getSheetByName('index');
    if (!indexSheet) {
      throw new Error('Sheet "index" not found');
    }

    const options = indexSheet.getRange('A2:A5').getValues();
    return options.map(function (row) { return row[0]; }).filter(function (val) { return val !== ''; });
  } catch (error) {
    throw new Error('Error loading options: ' + error.message);
  }
}
// Get product names for Portal B
function getProductNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const indexSheet = ss.getSheetByName('index');
    if (!indexSheet) {
      throw new Error('Sheet "index" not found');
    }

    const lastRow = indexSheet.getLastRow();
    if (lastRow < 2) return [];

    const products = indexSheet.getRange('B2:B' + lastRow).getValues();
    return products.map(function (row) { return row[0]; }).filter(function (val) { return val !== ''; });
  } catch (error) {
    throw new Error('Error loading products: ' + error.message);
  }
}
// [UPDATED] Get tank serial numbers (Merges both sheets)
function getTankSerialNumbers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let allSerials = [];
    // Get from NEW sheet
    const newTankSheet = ss.getSheetByName('Tank');
    if (newTankSheet) {
      allSerials = allSerials.concat(getSheetSerials(newTankSheet));
    }

    // Get from OLD sheet
    const oldTankSheet = ss.getSheetByName('Tank_2025');
    if (oldTankSheet) {
      allSerials = allSerials.concat(getSheetSerials(oldTankSheet));
    }

    // Remove duplicates and sort descending
    // Note: If Serial #1 exists in both, this will only show one #1. 
    // The system prioritizes the NEW sheet (2026) in getTankInfo.
    const uniqueSerials = [...new Set(allSerials)];
    uniqueSerials.sort(function (a, b) { return b - a; });

    return uniqueSerials;
  } catch (error) {
    throw new Error('Error loading tank serials: ' + error.message);
  }
}
function getSheetSerials(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const serials = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  return serials.map(function (row) { return row[0]; })
    .filter(function (val) { return val !== '' && !isNaN(val); });
}
// [UPDATED] Get tank info by serial number (Searches 2026 then 2025)
function getTankInfo(serialNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 1. Try to find in the NEW sheet first
    const newTankSheet = ss.getSheetByName('Tank');
    if (newTankSheet) {
      const result = searchSheetForTank(newTankSheet, serialNumber);
      if (result) return result;
    }

    // 2. If not found, try the OLD sheet
    const oldTankSheet = ss.getSheetByName('Tank_2025');
    if (oldTankSheet) {
      const result = searchSheetForTank(oldTankSheet, serialNumber);
      if (result) return result;
    }

    throw new Error('Serial number not found in current or archived records');
  } catch (error) {
    throw new Error('Error loading tank info: ' + error.message);
  }
}
// Helper function to avoid duplicate code
function searchSheetForTank(sheet, serialNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  // Column 13 contains Serial Numbers (Index 12)
  const serials = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  for (let i = 0; i < serials.length; i++) {
    if (serials[i][0] == serialNumber) {
      const rowNum = i + 2;
      const data = sheet.getRange(rowNum, 1, 1, 18).getValues()[0];
      return {
        timestamp: data[0] ? Utilities.formatDate(new Date(data[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        deviceReading1: data[10] || '',
        vvPercent: data[13] || '',
        previousDeviceReading: data[11] || '',
        previousVVPercent: data[14] || '',
        previousLabResult: data[15] || '',
        deviceAnalysisTime: data[16] ? Utilities.formatDate(new Date(data[16]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : '',
        labAnalysisTime: data[17] ? Utilities.formatDate(new Date(data[17]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : ''
      };
    }
  }
  return null;
}
// Save Tank Analysis (Concurrent Workflow Upsert)
function saveTankAnalysis(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tankSheet = ss.getSheetByName('Tank');
    if (!tankSheet) {
      tankSheet = ss.insertSheet('Tank');
      const headers = [
        'Timestamp', 'User Email', 'التاريخ', '', 'نوع الخامة', 'طريقة التحضير',
        'العدد (A2)', 'العدد (A3)', 'العدد (A4)', 'العدد (A5)',
        'Old Device', 'Old Prev Dev', 'Serial Number', 'Old v/v%', 'Old Prev v/v%',
        'نتيجة التحليل المعملي', 'وقت التحليل المعملي', 'صورة التحليل المعملي'
      ];
      tankSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      tankSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    const lastRow = tankSheet.getLastRow();
    let rowIndex = -1;

    if (lastRow >= 2) {
      const table = tankSheet.getRange(2, 1, lastRow - 1, 18).getValues();
      for (let i = 0; i < table.length; i++) {
        if (table[i][12] == data.serialNumber) { // Col M
          rowIndex = i + 2; break;
        }
      }
    }

    const timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const labResStr = (data.labResult / 100);

    if (rowIndex > -1) {
      // UPDATE EXISTING ROW (Started by Portal A)
      tankSheet.getRange(rowIndex, 16).setValue(labResStr); // Col P
      tankSheet.getRange(rowIndex, 17).setValue(timestampStr); // Col Q
      if (data.imageUrl) tankSheet.getRange(rowIndex, 18).setValue(data.imageUrl); // Col R
    } else {
      // INSERT NEW ROW (Portal B First)
      const row = new Array(18).fill('');
      row[12] = data.serialNumber; // Col M
      row[15] = labResStr; // Col P
      row[16] = timestampStr; // Col Q
      if (data.imageUrl) row[17] = data.imageUrl; // Col R

      tankSheet.appendRow(row);
    }

    return { success: true, message: 'تم حفظ نتيجة التحليل بنجاح' };
  } catch (error) {
    throw new Error('Error updating tank: ' + error.message);
  }
}
// Save product analysis (Portal B - Product path)
function saveProductAnalysis(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let analysisSheet = ss.getSheetByName('Analysis');
    if (!analysisSheet) {
      analysisSheet = ss.insertSheet('Analysis');
      const headers = ['Timestamp', 'Product Name', 'Analysis Method', 'Device Reading', 'v/v%', 'Lab Result', 'Reference Number', 'Image URL'];
      analysisSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      analysisSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    const row = new Array(8).fill('');
    row[0] = new Date();
    row[1] = data.productName;
    row[2] = 'معملي'; // Always Lab
    row[5] = data.labResult / 100;
    row[6] = data.referenceNumber;
    row[7] = data.imageUrl || '';

    analysisSheet.appendRow(row);
    return { success: true, message: 'تم حفظ البيانات بنجاح' };
  } catch (error) {
    throw new Error('Error saving analysis: ' + error.message);
  }
}
// Upload image to Google Drive
function uploadImageToDrive(imageData, fileName) {
  try {
    const folderId = '1L3c3hS4fBdQ6Z6eFXOzzplkmV6pXkOei';
    const folder = DriveApp.getFolderById(folderId);
    const contentType = imageData.substring(5, imageData.indexOf(';'));
    const bytes = Utilities.base64Decode(imageData.substr(imageData.indexOf('base64,') + 7));
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      success: true,
      url: file.getUrl(),
      id: file.getId()
    };
  } catch (error) {
    throw new Error('Error uploading image: ' + error.message);
  }
}
// Delete image from Google Drive
function deleteImageFromDrive(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return { success: true };
  } catch (error) {
    throw new Error('Error deleting image: ' + error.message);
  }
}
// Validate serial number (Concurrent Context-Aware Workflow)
function validateSerialNumber(serialNumber, portal) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tankSheet = ss.getSheetByName('Tank');
    if (!tankSheet || tankSheet.getLastRow() < 2) {
      if (serialNumber === 1) {
        return {
          valid: true,
          message: portal === 'B' ? "يرجي الابلاغ بالنتيجة للتجهيز.. جاهز لإدخال نتيجة التحليل." : "لم يتم التحليل بعد، يرجي مراجعة المعمل قبل الارسال",
          type: portal === 'B' ? 'success' : 'warning'
        };
      } else {
        return { valid: false, message: 'رقم التانك يجب أن يبدأ من 1', type: 'error' };
      }
    }

    const data = tankSheet.getRange(2, 1, tankSheet.getLastRow() - 1, 18).getValues();
    let maxSerial = 0;
    let existingRow = null;

    for (let i = 0; i < data.length; i++) {
      let s = data[i][12]; // Col M (index 12)
      if (s !== '' && !isNaN(s)) {
        let num = parseInt(s);
        if (num > maxSerial) maxSerial = num;
        if (num === serialNumber) existingRow = data[i];
      }
    }

    if (existingRow) {
      const hasMaterial = existingRow[4] !== ''; // Col E
      const hasLabResult = existingRow[15] !== ''; // Col P

      if (hasMaterial && hasLabResult) { // Case 1
        return { valid: false, message: 'هذا التانك مكتمل التسجيل والتحليل (مكرر).', type: 'error' };
      }
      if (portal === 'A' && hasMaterial) { // Case 2
        return { valid: false, message: 'تم تسجيل الإنتاج مسبقاً لهذا التانك.', type: 'error' };
      }
      if (portal === 'B' && hasLabResult) { // Case 3
        return { valid: false, message: 'تم تحليل هذا التانك مسبقاً.', type: 'error' };
      }
      if (portal === 'A' && hasLabResult && !hasMaterial) { // Case 4
        return { valid: true, message: 'تم التحليل معملياً.. جاهز لاستكمال بيانات الإنتاج.', type: 'success' };
      }
      if (portal === 'B' && hasMaterial && !hasLabResult) { // Case 5
        return { valid: true, message: 'تم ارسال التانك.. جاهز لإدخال نتيجة التحليل.', type: 'info' };
      }
    }

    if (serialNumber === maxSerial + 1) {
      if (portal === 'B') { // Case 6
        return { valid: true, message: 'يرجي الابلاغ بالنتيجة للتجهيز.. جاهز لإدخال نتيجة التحليل.', type: 'success' };
      } else { // Case 7
        return { valid: true, message: 'لم يتم التحليل بعد، يرجي مراجعة المعمل قبل الارسال', type: 'warning' };
      }
    } else {
      return { valid: false, message: 'رقم التانك يجب أن يكون ' + (maxSerial + 1), type: 'error' };
    }
  } catch (error) {
    throw new Error('Error validating serial: ' + error.message);
  }
}
// Save Portal A data (Concurrent Workflow Upsert)
function savePortalAData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tankSheet = ss.getSheetByName('Tank');
    if (!tankSheet) {
      tankSheet = ss.insertSheet('Tank');
      const headers = [
        'Timestamp', 'User Email', 'التاريخ', '', 'نوع الخامة', 'طريقة التحضير',
        'العدد (A2)', 'العدد (A3)', 'العدد (A4)', 'العدد (A5)',
        'Old Device', 'Old Prev Dev', 'Serial Number', 'Old v/v%', 'Old Prev v/v%',
        'نتيجة التحليل المعملي', 'وقت التحليل المعملي', 'صورة التحليل المعملي'
      ];
      tankSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      tankSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    const quantityCol = data.materialIndex + 4;
    let rowIndex = -1;

    if (!data.isBlossomOrRose) {
      const lastRow = tankSheet.getLastRow();
      if (lastRow >= 2) {
        const table = tankSheet.getRange(2, 1, lastRow - 1, 18).getValues();
        for (let i = 0; i < table.length; i++) {
          if (table[i][12] == data.serialNumber) { // Col M (Index 12)
            rowIndex = i + 2; break;
          }
        }
      }
    }

    const timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    if (rowIndex > -1) {
      // UPDATE EXISTING ROW (Started by Portal B)
      tankSheet.getRange(rowIndex, 1).setValue(timestampStr);
      tankSheet.getRange(rowIndex, 2).setValue(Session.getActiveUser().getEmail());
      tankSheet.getRange(rowIndex, 3).setValue('تاريخ اليوم');
      tankSheet.getRange(rowIndex, 5).setValue(data.material); // Col E
      tankSheet.getRange(rowIndex, 6).setValue(data.preparation || ''); // Col F
      tankSheet.getRange(rowIndex, quantityCol + 1).setValue(data.quantity);
    } else {
      // INSERT NEW ROW (Portal A First)
      const row = new Array(18).fill('');
      row[0] = timestampStr;
      row[1] = Session.getActiveUser().getEmail();
      row[2] = 'تاريخ اليوم';
      row[4] = data.material;
      row[5] = data.preparation || '';
      row[quantityCol] = data.quantity;
      if (!data.isBlossomOrRose) row[12] = data.serialNumber; // Col M

      tankSheet.appendRow(row);
    }

    // Export to external spreadsheet (Strictly 14 Columns)
    try {
      const extRow = new Array(14).fill('');
      extRow[0] = timestampStr;
      extRow[1] = Session.getActiveUser().getEmail();
      extRow[2] = 'تاريخ اليوم';
      extRow[4] = data.material;
      extRow[5] = data.preparation || '';
      extRow[quantityCol] = data.quantity;
      if (!data.isBlossomOrRose) extRow[12] = data.serialNumber; // Col M
      saveToExternalSpreadsheet(extRow);
    } catch (err) {
      Logger.log('Error saving external: ' + err.message);
    }

    return { success: true, message: 'تم حفظ بيانات الإنتاج بنجاح' };
  } catch (error) {
    throw new Error('Error saving data: ' + error.message);
  }
}
// [UPDATED] Save to external spreadsheet with NEW ID
function saveToExternalSpreadsheet(rowData) {
  try {
    // UPDATED: New Spreadsheet ID for 2026
    const externalSpreadsheetId = '1AATh8aLorXF428tbExFU6LjWAdsn8ybQjp4933BeTzU';
    // Open external spreadsheet
    const externalSS = SpreadsheetApp.openById(externalSpreadsheetId);

    // Get the first sheet
    const externalSheet = externalSS.getSheetById(1416876654);

    // Append the row to external spreadsheet
    externalSheet.appendRow(rowData);

    Logger.log('Successfully saved to external spreadsheet');
  } catch (error) {
    throw new Error('Failed to save to external spreadsheet: ' + error.message);
  }
}
// Portal C - Get report data
function getReportData(reportType, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const start = new Date(startDate);
    start.setHours(0, 0, 0, 0);
    const end = new Date(endDate);
    end.setHours(23, 59, 59, 999);

    if (reportType === 'tank') {
      return getTankReportData(ss, start, end);
    } else if (reportType === 'analysis') {
      return getAnalysisReportData(ss, start, end);
    }

    throw new Error('نوع التقرير غير صحيح');
  } catch (error) {
    throw new Error('خطأ في تحميل البيانات: ' + error.message);
  }
}
// [UPDATED] Get Tank Report Data (Includes Lab Result & Merges Sheets)
function getTankReportData(ss, startDate, endDate) {
  let allData = [];
  // 1. Get data from NEW sheet
  const newTankSheet = ss.getSheetByName('Tank');
  if (newTankSheet) {
    allData = allData.concat(extractSheetData(newTankSheet));
  }
  // 2. Get data from OLD sheet
  const oldTankSheet = ss.getSheetByName('Tank_2025');
  if (oldTankSheet) {
    allData = allData.concat(extractSheetData(oldTankSheet));
  }
  const filtered = [];
  allData.forEach(function (row) {
    const timestamp = new Date(row[0]);
    if (timestamp >= startDate && timestamp <= endDate) {
      filtered.push({
        date: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
        serial: row[12] || 'N/A',
        material: row[4] || 'N/A',
        preparation: row[5] || '',
        quantity: row[6] || row[7] || row[8] || row[9] || 'N/A',
        reading: row[10] ? parseFloat(row[10]).toFixed(2) : 'N/A',
        vv: row[13] ? (parseFloat(row[13]) * 100).toFixed(2) : 'N/A',
        // NEW: Lab Result from Column P (Index 15)
        labResult: row[15] ? (parseFloat(row[15]) * 100).toFixed(2) + '%' : '-',
        // NEW: Image URL from Column R (Index 17)
        imageUrl: row[17] ? String(row[17]) : ''
      });
    }
  });
  filtered.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
  return filtered;
}
// [UPDATED] Helper to extract raw data (Widened range to 18 columns)
function extractSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // UPDATED: Now reads 18 columns (A to R) instead of 14 to include Lab Data
  return sheet.getRange(2, 1, lastRow - 1, 18).getValues();
}
// [UPDATED] Get Analysis Report Data (Combines Products + Tanks from 2026 & 2025)
function getAnalysisReportData(ss, startDate, endDate) {
  const filtered = [];
  // --- PART 1: Get Product Analysis (Existing Logic) ---
  const analysisSheet = ss.getSheetByName('Analysis');
  if (analysisSheet && analysisSheet.getLastRow() >= 2) {
    const data = analysisSheet.getRange(2, 1, analysisSheet.getLastRow() - 1, 8).getValues();
    data.forEach(function (row) {
      const timestamp = new Date(row[0]);
      if (timestamp >= startDate && timestamp <= endDate) {
        filtered.push({
          date: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
          product: row[1] || 'N/A',
          method: row[2] || 'N/A',
          deviceReading: row[3] ? parseFloat(row[3]).toFixed(2) : 'N/A',
          vv: row[4] ? (parseFloat(row[4]) * 100).toFixed(2) + '%' : 'N/A',
          labResult: row[5] ? (parseFloat(row[5]) * 100).toFixed(2) + '%' : 'N/A',
          refNumber: row[6] || 'N/A',
          imageUrl: row[7] || '',
          type: 'product' // Internal marker
        });
      }
    });
  }
  // --- PART 2: Get Tank Analysis (New Logic) ---
  // We check both current Tank sheet and archived Tank_2025
  const sheetsToCheck = ['Tank', 'Tank_2025'];
  sheetsToCheck.forEach(function (sheetName) {
    const tankSheet = ss.getSheetByName(sheetName);
    if (!tankSheet || tankSheet.getLastRow() < 2) return;
    const tankData = tankSheet.getRange(2, 1, tankSheet.getLastRow() - 1, 18).getValues();

    tankData.forEach(function (row) {
      const serial = row[12];

      if (sheetName === 'Tank_2025') {
        // Old 2025 Structure: Col Q (16) = Device Time, Col R (17) = Lab Time
        const deviceTime = row[16] ? new Date(row[16]) : null;
        if (deviceTime && deviceTime >= startDate && deviceTime <= endDate) {
          filtered.push({
            date: Utilities.formatDate(deviceTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
            product: 'Tank ' + serial,
            method: 'الجهاز',
            deviceReading: row[10] ? parseFloat(row[10]).toFixed(2) : 'N/A',
            vv: row[13] ? (parseFloat(row[13]) * 100).toFixed(2) + '%' : 'N/A',
            labResult: 'N/A',
            refNumber: serial,
            imageUrl: '',
            type: 'tank'
          });
        }
        const labTime = row[17] ? new Date(row[17]) : null;
        if (labTime && labTime >= startDate && labTime <= endDate) {
          filtered.push({
            date: Utilities.formatDate(labTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
            product: 'Tank ' + serial,
            method: 'معملي',
            deviceReading: 'N/A',
            vv: 'N/A',
            labResult: row[15] ? (parseFloat(row[15]) * 100).toFixed(2) + '%' : 'N/A',
            refNumber: serial,
            imageUrl: '',
            type: 'tank'
          });
        }
      } else {
        // New 2026 Structure: Col Q (16) = Lab Time, Col R (17) = Image URL
        const labTime = row[16] ? new Date(row[16]) : null;
        if (labTime && labTime >= startDate && labTime <= endDate) {
          filtered.push({
            date: Utilities.formatDate(labTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
            product: 'Tank ' + serial,
            method: 'معملي',
            deviceReading: 'N/A',
            vv: 'N/A',
            labResult: row[15] ? (parseFloat(row[15]) * 100).toFixed(2) + '%' : 'N/A',
            refNumber: serial,
            imageUrl: row[17] ? String(row[17]) : '',
            type: 'tank'
          });
        }
      }
    });
  });
  // Sort by Date Descending
  filtered.sort(function (a, b) { return new Date(b.date) - new Date(a.date); });
  return filtered;
}
// Export report to Excel
function exportReportToExcel(reportType, reportData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const sheetName = reportType === 'tank' ? 'تقرير_تانكات_' + timestamp : 'تقرير_تحليلات_' + timestamp;
    const reportSheet = ss.insertSheet(sheetName);

    const title = reportType === 'tank' ? 'تقرير التانكات' : 'تقرير التحليلات';
    reportSheet.getRange(1, 1).setValue(title);
    reportSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');

    const exportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    reportSheet.getRange(2, 1).setValue('تاريخ التصدير: ' + exportDate);
    reportSheet.getRange(2, 1).setFontSize(10).setFontStyle('italic');

    if (reportType === 'tank') {
      const headers = [['التاريخ', 'رقم العينة', 'نوع الخامة', 'طريقة التحضير', 'العدد', 'نتيجة المعمل', 'صورة']];
      reportSheet.getRange(4, 1, 1, 7).setValues(headers);
      reportSheet.getRange(4, 1, 1, 7).setFontWeight('bold').setBackground('#3b82f6').setFontColor('#ffffff');
      if (reportData.length > 0) {
        const rows = reportData.map(function (row) {
          return [
            row.date,
            row.serial,
            row.material,
            row.preparation || '-',
            row.quantity || '-',
            row.labResult,
            row.imageUrl || '-'
          ];
        });
        reportSheet.getRange(5, 1, rows.length, 7).setValues(rows);
        reportSheet.getRange(4, 1, rows.length + 1, 7).setBorder(true, true, true, true, true, true);
      }

      for (let i = 1; i <= 7; i++) {
        reportSheet.autoResizeColumn(i);
      }
    } else {
      const headers = [['التاريخ', 'اسم المنتج', 'نتيجة المعمل', 'رقم التشغيلة', 'صورة']];
      reportSheet.getRange(4, 1, 1, 5).setValues(headers);
      reportSheet.getRange(4, 1, 1, 5).setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');

      if (reportData.length > 0) {
        const rows = reportData.map(function (row) {
          return [
            row.date,
            row.product,
            row.labResult !== 'N/A' ? row.labResult : '-',
            row.refNumber,
            row.imageUrl || '-'
          ];
        });
        reportSheet.getRange(5, 1, rows.length, 5).setValues(rows);
        reportSheet.getRange(4, 1, rows.length + 1, 5).setBorder(true, true, true, true, true, true);
      }

      for (let i = 1; i <= 5; i++) {
        reportSheet.autoResizeColumn(i);
      }
    }

    const summaryRow = reportSheet.getLastRow() + 2;
    reportSheet.getRange(summaryRow, 1).setValue('إجمالي السجلات: ' + reportData.length);
    reportSheet.getRange(summaryRow, 1).setFontWeight('bold');

    return ss.getUrl() + '#gid=' + reportSheet.getSheetId();

  } catch (error) {
    throw new Error('خطأ في التصدير: ' + error.message);
  }
}
// Dashboard Functions
function getTodayTanksDashboardData(targetDateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tankSheet = ss.getSheetByName('Tank');
    if (!tankSheet || tankSheet.getLastRow() < 2) return getEmptyTanksDashboard();

    const tz = Session.getScriptTimeZone();
    let today;
    if (targetDateStr) {
      today = new Date(targetDateStr);
    } else {
      today = new Date();
    }
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const data = tankSheet.getRange(2, 1, tankSheet.getLastRow() - 1, 18).getValues();

    let todayData = [];
    data.forEach(function (row) {
      // A tank can be initiated by Production (Col A/0) OR by the Lab (Col Q/16)
      let regTime = row[0] ? new Date(row[0]) : null;
      let labTime = row[16] ? new Date(row[16]) : null;
      
      let isValidDay = false;
      if (regTime && regTime >= today && regTime < tomorrow) isValidDay = true;
      if (labTime && labTime >= today && labTime < tomorrow) isValidDay = true;

      if (isValidDay) {
        let serial = parseInt(row[12]);
        if (!isNaN(serial)) {
          todayData.push({
            // Show registration time if it exists, otherwise fallback to lab analysis time
            timestamp: regTime ? regTime : labTime, 
            serial: serial,
            labResult: row[15] !== '' ? parseFloat(row[15]) * 100 : null,
            productType: row[4] || 'غير مسجل', // Provide a fallback label if Col E is empty
            isRegistered: regTime !== null
          });
        }
      }
    });

    const serials = todayData.map(function(i) { return i.serial; });
    const startSerial = serials.length > 0 ? Math.min.apply(null, serials) : 0;
    const endSerial = serials.length > 0 ? Math.max.apply(null, serials) : 0;
    
    // Find the highest serial number actually registered by production
    const registeredSerials = todayData.filter(function(i) { return i.isRegistered; }).map(function(i) { return i.serial; });
    const maxRegisteredSerial = registeredSerials.length > 0 ? Math.max.apply(null, registeredSerials) : 0;

    const concData = todayData.filter(function(i) { return i.labResult !== null; });
    let avg = 0, highest = 0, lowest = 0;
    let highestSerials = [], lowestSerials = [];

    if (concData.length > 0) {
      let sum = 0;
      concData.forEach(function(i) { sum += i.labResult; });
      avg = sum / concData.length;
      
      const vals = concData.map(function(i) { return i.labResult; });
      highest = Math.max.apply(null, vals);
      lowest = Math.min.apply(null, vals);

      highestSerials = concData.filter(function(i) { return i.labResult === highest; }).map(function(i) { return i.serial; });
      lowestSerials = concData.filter(function(i) { return i.labResult === lowest; }).map(function(i) { return i.serial; });
    }

    // Fetch Handover Data safely (Foolproof ISO String Matching)
    let handover = {};
    let prevShiftLastTank = {};
    try {
      const hoSheet = ss.getSheetByName("Shift_Handovers");
      if (hoSheet && hoSheet.getLastRow() >= 2) {
         const hoData = hoSheet.getRange(2, 1, hoSheet.getLastRow() - 1, 4).getValues();
         let hoRow = null;

         // Establish target date string (always yyyy-MM-dd format)
         const targetDate = targetDateStr || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

         // Look backwards for the first handover timestamp that is strictly BEFORE the target date
         for (let i = hoData.length - 1; i >= 0; i--) {
           let hoDateStr = '';
           if (hoData[i][0]) {
             try {
               hoDateStr = Utilities.formatDate(new Date(hoData[i][0]), tz, 'yyyy-MM-dd');
             } catch (e) {
               // Safe fallback if row is read as a raw string format 'YYYY-MM-DD HH:mm:ss'
               hoDateStr = String(hoData[i][0]).substring(0, 10);
             }
           }

           if (hoDateStr && hoDateStr < targetDate) {
             hoRow = hoData[i];
             break;
           }
         }

         // Global Fallback: If no previous row satisfies the condition, load the absolute last available entry
         if (!hoRow && hoData.length > 0) {
            hoRow = hoData[hoData.length - 1];
         }

         if (hoRow) {
           handover = {
              time: String(hoRow[0]),
              lastTankNum: parseInt(hoRow[1]),
              expectedNext: parseInt(hoRow[2]),
              image: String(hoRow[3])
           };
           // Scan historical data for the actual last tank to get its real timestamp and image
           for(let i = data.length - 1; i >= 0; i--) {
              if(parseInt(data[i][12]) === handover.lastTankNum) {
                 prevShiftLastTank = {
                    time: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), tz, 'yyyy-MM-dd HH:mm') : '',
                    image: String(data[i][17] || '')
                 };
                 break;
              }
           }
         }
      }
    } catch(e) {}

    // Format timestamps and lab results for the frontend chip popovers
    const formattedTanks = {};
    todayData.forEach(function(t) {
      formattedTanks[t.serial] = {
        time: Utilities.formatDate(t.timestamp, Session.getScriptTimeZone(), 'hh:mm a'),
        labResult: t.labResult !== null ? t.labResult.toFixed(2) + '%' : '-',
        productType: t.productType,
        isRegistered: t.isRegistered
      };
    });

    return {
      totalCount: todayData.length,
      startSerial: startSerial,
      endSerial: endSerial,
      maxRegisteredSerial: maxRegisteredSerial,
      avgConc: avg.toFixed(2),
      highestConc: highest.toFixed(2),
      lowestConc: lowest.toFixed(2),
      highestSerialUnique: highestSerials.length === 1 ? highestSerials[0] : null,
      lowestSerialUnique: lowestSerials.length === 1 ? lowestSerials[0] : null,
      hasConcentrationData: concData.length > 0,
      handover: handover,
      prevShiftLastTank: prevShiftLastTank,
      tanksMap: formattedTanks
    };
  } catch (error) {
    return getEmptyTanksDashboard();
  }
}

function getEmptyTanksDashboard() {
  return {
    totalCount: 0, startSerial: 0, endSerial: 0, maxRegisteredSerial: 0,
    avgConc: '0.00', highestConc: '0.00', lowestConc: '0.00',
    highestSerialUnique: null, lowestSerialUnique: null,
    hasConcentrationData: false,
    handover: {}, prevShiftLastTank: {},
    tanksMap: {}
  };
}
function getTodayAnalysisDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const analysisSheet = ss.getSheetByName('Analysis');
    const tankSheet = ss.getSheetByName('Tank');
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);

    let deviceProducts = 0;
    let deviceTanks = 0;
    let labProducts = 0;
    let labTanks = 0;
    let allConcentrations = [];

    if (analysisSheet) {
      const lastRow = analysisSheet.getLastRow();
      if (lastRow >= 2) {
        const data = analysisSheet.getRange(2, 1, lastRow - 1, 7).getValues();

        data.forEach(function (row, index) {
          const timestamp = new Date(row[0]);
          if (timestamp >= today && timestamp < tomorrow) {
            const method = row[2];
            const vv = row[4];
            const labResult = row[5];
            const refNumber = row[6];

            if (method === 'الجهاز') {
              deviceProducts++;
              if (vv && !isNaN(vv)) {
                allConcentrations.push({
                  value: parseFloat(vv) * 100,
                  id: refNumber,
                  label: 'منتج: ' + refNumber,
                  rowIndex: index + 2
                });
              }
            } else if (method === 'معملي') {
              labProducts++;
              if (labResult && !isNaN(labResult)) {
                allConcentrations.push({
                  value: parseFloat(labResult) * 100,
                  id: refNumber,
                  label: 'منتج: ' + refNumber,
                  rowIndex: index + 2
                });
              }
            }
          }
        });
      }
    }

    if (tankSheet) {
      const lastRow = tankSheet.getLastRow();
      if (lastRow >= 2) {
        const data = tankSheet.getRange(2, 1, lastRow - 1, 18).getValues();

        data.forEach(function (row, index) {
          const deviceAnalysisTime = row[16];
          const labAnalysisTime = row[17];
          const serial = row[12];

          if (deviceAnalysisTime) {
            const analysisDate = new Date(deviceAnalysisTime);
            if (analysisDate >= today && analysisDate < tomorrow) {
              deviceTanks++;
              const vv = row[14];
              if (vv && !isNaN(vv)) {
                allConcentrations.push({
                  value: parseFloat(vv) * 100,
                  id: 'T' + serial,
                  label: 'تانك: ' + serial,
                  rowIndex: index + 2
                });
              }
            }
          }

          if (labAnalysisTime) {
            const analysisDate = new Date(labAnalysisTime);
            if (analysisDate >= today && analysisDate < tomorrow) {
              labTanks++;
              const labResult = row[15];
              if (labResult && !isNaN(labResult)) {
                allConcentrations.push({
                  value: parseFloat(labResult) * 100,
                  id: 'T' + serial,
                  label: 'تانك: ' + serial,
                  rowIndex: index + 2
                });
              }
            }
          }
        });
      }
    }

    let highest = 0;
    let lowest = 0;
    let highestId = '0';
    let lowestId = '0';
    let highestLabel = '';
    let lowestLabel = '';

    if (allConcentrations.length > 0) {
      const sortedConc = allConcentrations.sort(function (a, b) { return b.value - a.value; });

      highest = sortedConc[0].value;
      highestId = sortedConc[0].id;
      highestLabel = sortedConc[0].label;

      lowest = sortedConc[sortedConc.length - 1].value;
      lowestId = sortedConc[sortedConc.length - 1].id;
      lowestLabel = sortedConc[sortedConc.length - 1].label;
    }

    return {
      deviceTotal: deviceProducts + deviceTanks,
      deviceProducts: deviceProducts,
      deviceTanks: deviceTanks,
      labTotal: labProducts + labTanks,
      labProducts: labProducts,
      labTanks: labTanks,
      highest: highest.toFixed(2),
      lowest: lowest.toFixed(2),
      highestId: highestId,
      lowestId: lowestId,
      highestLabel: highestLabel,
      lowestLabel: lowestLabel
    };

  } catch (error) {
    Logger.log('Error in getTodayAnalysisDashboardData: ' + error.message);
    return getEmptyAnalysisDashboard();
  }
}

function getEmptyAnalysisDashboard() {
  return {
    deviceTotal: 0,
    deviceProducts: 0,
    deviceTanks: 0,
    labTotal: 0,
    labProducts: 0,
    labTanks: 0,
    highest: '0.00',
    lowest: '0.00',
    highestId: '0',
    lowestId: '0',
    highestLabel: '',
    lowestLabel: ''
  };
}
// --- TDS GENERATOR FUNCTIONS ---
// Get Product Options from TDS_Settings
function getTDSProducts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TDS_Settings');
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Get Product Names from Column A
    const products = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return products.map(function (row) { return row[0]; }).filter(function (val) { return val !== ''; });
  } catch (error) {
    throw new Error('Error loading TDS products: ' + error.message);
  }
}
// Get Fixed Settings for a specific product (Safe Version)
function getTDSDetails(productName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TDS_Settings');
    const lastRow = sheet.getLastRow();
    // Get all data A2:K
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == productName) {
        // Convert EVERYTHING to String to avoid "Complex Object" errors
        return {
          product: String(data[i][0] || ''),
          standard: String(data[i][1] || ''),
          appearance: String(data[i][2] || ''),
          odor: String(data[i][3] || ''),
          density: String(data[i][4] || ''),
          alcohol: String(data[i][5] || ''),
          solids: String(data[i][6] || ''),
          micro: String(data[i][7] || ''),
          shelfLife: String(data[i][8] || ''),
          storage: String(data[i][9] || ''),
          logoUrl: String(data[i][10] || '')
        };
      }
    }
    throw new Error('Product settings not found');
  } catch (error) {
    throw new Error('Error fetching TDS details: ' + error.message);
  }
}
// --- TDS GENERATION & LOGGING ---
function generateAndLogTDS(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('TDS_Logs');
    if (!logSheet) throw new Error('Sheet "TDS_Logs" not found');
    // 1. Generate Unique ID
    const year = new Date().getFullYear();
    const lastRow = Math.max(logSheet.getLastRow(), 1);
    const sequence = lastRow; // Starts at 1 if sheet is empty
    const padSequence = ('000' + sequence).slice(-3);
    const uniqueID = 'TDS-' + year + '-' + padSequence;

    // 2. Fetch Fixed Settings for the PDF
    const settings = getTDSDetails(data.productName);

    // 3. Prepare Data for Template
    data.documentId = uniqueID; // Add ID to data object

    const template = HtmlService.createTemplateFromFile('TDSTEMPLATE');
    template.data = data;
    template.settings = settings;

    // 4. Generate PDF Blob
    const htmlOutput = template.evaluate();
    const pdfBlob = htmlOutput.getAs('application/pdf');
    pdfBlob.setName(uniqueID + '_' + data.productName + '.pdf');

    // 5. Save to Google Drive (Folder: "TDS Certificates")
    const folderName = "TDS Certificates";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    // 6. Save Log to Spreadsheet (Col K = URL)
    const row = [
      uniqueID,                  // A: ID
      new Date(),                // B: Timestamp
      data.batchNumber,          // C: Batch
      data.productName,          // D: Product
      data.prodDate,             // E: Prod
      data.expDate,              // F: Exp
      data.acidity,              // G: Acidity
      data.oxidation,            // H: Oxidation
      data.ph,                   // I: pH
      Session.getActiveUser().getEmail(), // J: User
      fileUrl                    // K: URL
    ];

    logSheet.appendRow(row);

    return { success: true, documentId: uniqueID, url: fileUrl };

  } catch (error) {
    throw new Error('Error generating TDS: ' + error.message);
  }
}
// --- GENERIC TDS GENERATION ---
function generateAndLogGenericTDS(productName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('TDS_Logs');
    // 1. Generate ID (use "SPEC" prefix instead of "TDS")
    const year = new Date().getFullYear();
    const lastRow = Math.max(logSheet.getLastRow(), 1);
    const uniqueID = 'SPEC-' + year + '-' + ('000' + lastRow).slice(-3);

    // 2. Fetch Fixed Settings
    const settings = getTDSDetails(productName);

    // 3. Prepare Data (Flag isGeneric = true)
    const data = {
      productName: productName,
      documentId: uniqueID,
      isGeneric: true  // <--- This controls the template
    };

    // 4. Generate PDF
    const template = HtmlService.createTemplateFromFile('TDSTEMPLATE');
    template.data = data;
    template.settings = settings;

    const htmlOutput = template.evaluate();
    const pdfBlob = htmlOutput.getAs('application/pdf');
    pdfBlob.setName(uniqueID + '_' + productName + '_Spec.pdf');

    // 5. Save to Drive
    const folderName = "TDS Certificates";
    const folders = DriveApp.getFoldersByName(folderName);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    // 6. Save Log (Simple log for generic specs)
    const row = [
      uniqueID,
      new Date(),
      'GENERIC',     // Batch
      productName,
      '-',           // Prod Date
      '-',           // Exp Date
      '-', '-', '-', // Results
      Session.getActiveUser().getEmail(),
      fileUrl
    ];
    logSheet.appendRow(row);

    return { success: true, documentId: uniqueID, url: fileUrl };

  } catch (error) {
    throw new Error('Error generating Generic Spec: ' + error.message);
  }
}

function getProdSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Production_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('Production_Logs');
      sheet.appendRow(['Session ID (Master)', 'Timestamp', 'User Name', 'Date', 'Product Name', 'Start Time', 'End Time', 'Duration (Mins)', 'Quantity', 'Avg Time/Unit (Mins)', 'Notes', 'Image URL', 'Status', 'Segment ID']);
      sheet.getRange(1, 1, 1, 14).setFontWeight('bold');
    }
    const idxSheet = ss.getSheetByName('index');
    let products = [];
    if (idxSheet && idxSheet.getLastRow() >= 2) {
      const data = idxSheet.getRange(2, 2, idxSheet.getLastRow() - 1, 1).getValues();
      products = data.map(function (r) { return r[0]; }).filter(function (v) { return v !== ''; });
    }
    return { products: products };
  } catch (e) { throw new Error(e.message); }
}
function startProdSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const sessionId = 'PRD-' + new Date().getTime();
    const segmentId = sessionId + '-1'; // Phase 1: Initialize Segment ID
    sheet.appendRow([
      sessionId, new Date(), data.user, data.date, data.product,
      data.startTime, '', '', '', '', '', '', 'Active', segmentId
    ]);
    syncWasteTimes(data.date);
    return { success: true };
  } catch (e) { throw new Error(e.message); }
}
function endProdSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('لا توجد بيانات');
    const table = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    let rowIndex = -1;

    // Find the ACTIVE row for this session (searching from bottom ensures we get the latest resumed segment)
    for (let i = table.length - 1; i >= 0; i--) {
      if (table[i][0] === data.sessionId && table[i][12] !== 'Completed' && table[i][12] !== 'Paused' && table[i][12] !== 'Cancelled') {
        rowIndex = i + 2; break;
      }
    }
    if (rowIndex === -1) throw new Error('لم يتم العثور على الجلسة النشطة');

    const startTimeStr = table[rowIndex - 2][5];
    const duration = calculateDuration(startTimeStr, data.endTime);
    const tz = Session.getScriptTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');

    // Determine Action Logic
    let statusStr = 'Completed';
    let notePrefix = '';

    if (data.action === 'pause') {
      statusStr = 'Paused';
      notePrefix = '[إيقاف مؤقت] ';
    } else if (data.action === 'cancel') {
      statusStr = 'Cancelled';
      notePrefix = '[إلغاء] ';
    }

    // Save Current Segment Data
    sheet.getRange(rowIndex, 7).setValue(data.endTime);
    sheet.getRange(rowIndex, 8).setValue(duration);
    let existingNote = String(table[rowIndex - 2][10] || '');
    let finalNote = existingNote + (existingNote ? '\n' : '') + notePrefix + (data.notes || '');
    sheet.getRange(rowIndex, 11).setValue(finalNote);
    if (data.imageUrl) sheet.getRange(rowIndex, 12).setValue(data.imageUrl);
    sheet.getRange(rowIndex, 13).setValue(statusStr);
    addCellNote(sheet, rowIndex, 13, `[${timestamp}] Session ${statusStr} by ${data.user}.`);

    // --- PROPORTIONAL ACCRUAL ACCOUNTING (If Completed) ---
    if (data.action === 'complete') {
      const updatedTable = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
      let segments = [];
      let totalDur = 0;

      // Find ALL segments belonging to this Session ID (including Resumed ones)
      for (let i = 0; i < updatedTable.length; i++) {
        let stat = updatedTable[i][12];
        if (updatedTable[i][0] === data.sessionId && (stat === 'Paused' || stat === 'Resumed' || stat === 'Completed')) {
          segments.push(i + 2);
          totalDur += (parseFloat(updatedTable[i][7]) || 0); // Index 7 is Duration (Col H)
        }
      }

      // Distribute Quantity Proportionally across all days/segments
      const totalQty = parseFloat(data.quantity) || 0;
      if (totalDur > 0 && totalQty > 0) {
        segments.forEach(rIdx => {
          let segDur = parseFloat(sheet.getRange(rIdx, 8).getValue()) || 0;
          let proportion = segDur / totalDur;
          // Strictly round quantity to 1 decimal place, stripping trailing zeroes
          let segQty = parseFloat((proportion * totalQty).toFixed(1));
          let segAvg = segQty > 0 ? (segDur / segQty).toFixed(2) : 0;

          sheet.getRange(rIdx, 9).setValue(segQty);
          sheet.getRange(rIdx, 10).setValue(segAvg);
          sheet.getRange(rIdx, 13).setValue('Completed'); // Force status to Completed
        });
      }
    }

    let dateStr = table[rowIndex - 2][3] instanceof Date ? Utilities.formatDate(table[rowIndex - 2][3], tz, 'yyyy-MM-dd') : String(table[rowIndex - 2][3]);
    syncWasteTimes(dateStr);

    return { success: true };
  } catch (e) { throw new Error(e.message); }
}
function resumeProdSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('لا توجد بيانات');
    const table = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    let parentRow = null;
    let parentRowIndex = -1;

    // Find the last paused row for this session to copy its data
    for (let i = table.length - 1; i >= 0; i--) {
      if (table[i][0] === data.sessionId && table[i][12] === 'Paused') {
        parentRow = table[i];
        parentRowIndex = i + 2;
        break;
      }
    }
    if (!parentRow) throw new Error('لم يتم العثور على تشغيلة معلقة لهذا المعرف');

    // Update the old Paused row to 'Resumed' so it disappears from the frontend queue
    sheet.getRange(parentRowIndex, 13).setValue('Resumed');

    const tz = Session.getScriptTimeZone();
    const newDate = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'); // Resume happens TODAY

    // Create new row segment linked to the same sessionId (Master_Batch_ID)
    const segmentId = parentRow[0] + '-' + new Date().getTime(); // Phase 1: Unique Segment ID
    sheet.appendRow([
      parentRow[0], // A: sessionId (Master)
      new Date(),   // B: Timestamp
      data.user,    // C: User
      newDate,      // D: Date
      parentRow[4], // E: Product
      data.resumeTime, // F: Start Time
      '', '', '', '', '', '', 'Active', // Force status to Active
      segmentId     // N: Segment ID
    ]);

    return { success: true };
  } catch (e) { throw new Error(e.message); }
}
function getTodayProdSessions(targetDateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    if (!sheet || sheet.getLastRow() < 2) return [];
    const todayStr = targetDateStr || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // Auto-Sync Waste Times BEFORE fetching the data to ensure the dashboard timeline is 100% accurate on load
    syncWasteTimes(todayStr);

    // Re-fetch the Last Row because syncWasteTimes might have added or deleted rows!
    const finalLastRow = sheet.getLastRow();
    const data = finalLastRow < 2 ? [] : sheet.getRange(2, 1, finalLastRow - 1, 14).getValues();
    const notesData = finalLastRow < 2 ? [] : sheet.getRange(2, 1, finalLastRow - 1, 14).getNotes();

    const filtered = [];
    for (let i = 0; i < data.length; i++) {
      // Ensure the cell date is formatted to match todayStr regardless of how Sheets saved it
      let rowDateStr = '';
      if (data[i][3]) {
        try {
          rowDateStr = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
          rowDateStr = String(data[i][3]).trim();
        }
      }

      if (rowDateStr === todayStr || data[i][12] === 'Paused') {
        let sTime = data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'HH:mm') : String(data[i][5] || '');
        let eTime = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'HH:mm') : String(data[i][6] || '');

        // Combine notes from Start, End, Qty, Notes, and Image columns (Indices 5, 6, 8, 10, 11)
        let combinedNotes = [notesData[i][5], notesData[i][6], notesData[i][8], notesData[i][10], notesData[i][11]]
          .filter(function (n) { return n !== ''; })
          .join('\n');

        filtered.push({
          id: data[i][0], user: data[i][2], date: rowDateStr, product: data[i][4],
          startTime: sTime, endTime: eTime, duration: data[i][7],
          quantity: data[i][8], avg: data[i][9], notes: data[i][10], image: data[i][11], status: data[i][12],
          editHistory: combinedNotes
        });
      }
    }
    return filtered.reverse();
// Newest first
  } catch (e) { throw new Error(e.message);
}
}

// ============================================================================
// --- PHASE 2: NEW UNIFIED DUAL-MATH ENGINE & CONCURRENCY PROFILER ---
// ============================================================================

function getUnifiedProdData(targetDateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    if (!sheet || sheet.getLastRow() < 2) return { timeline: [], masterCards: [] };
    
    const tz = Session.getScriptTimeZone();
    const todayStr = targetDateStr || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    
    // Auto-Sync Waste Times for the visual timeline
    syncWasteTimes(todayStr);
    
    const finalLastRow = sheet.getLastRow();
    const data = finalLastRow < 2 ? [] : sheet.getRange(2, 1, finalLastRow - 1, 14).getValues();
    const notesData = finalLastRow < 2 ? [] : sheet.getRange(2, 1, finalLastRow - 1, 14).getNotes();

    const touchedMasterIds = new Set();
    const timelineSegments = [];
    const breaksTimeline = [];
    
    // PASS 1: Identify all Master IDs touched today & build Factory Minute Map
    for (let i = 0; i < data.length; i++) {
      let rowDateStr = data[i][3] instanceof Date ? Utilities.formatDate(new Date(data[i][3]), tz, 'yyyy-MM-dd') : String(data[i][3]).trim();
      let status = data[i][12];
      let product = data[i][4];
      
      // If it belongs to today, or is currently hanging in Paused/Active state
      if (rowDateStr === todayStr || status === 'Paused' || status === 'Active') {
         if (product !== 'وقت مهدر' && product !== 'وقت الراحة') {
            touchedMasterIds.add(data[i][0]); // Add Master_Batch_ID
         }
      }
      
      // Collect purely today's physical blocks for the UI Timeline & Concurrency Mapping
      if (rowDateStr === todayStr) {
         let sTime = data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], tz, 'HH:mm') : String(data[i][5] || '');
         let eTime = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], tz, 'HH:mm') : String(data[i][6] || '');
         
         let segmentObj = {
            masterId: data[i][0], segmentId: data[i][13], user: data[i][2], date: rowDateStr, product: product,
            startTime: sTime, endTime: eTime, duration: data[i][7], status: status
         };
         
         timelineSegments.push(segmentObj);
         if (product === 'وقت الراحة') breaksTimeline.push(segmentObj);
      }
    }

    // BUILD FACTORY MINUTE MAP (For Concurrency Math)
    // 0 = Idle, 1 = Solo run, 2+ = Simultaneous Batches
    const minuteMap = new Array(1440).fill(0);
    timelineSegments.forEach(seg => {
       if (seg.product !== 'وقت مهدر' && seg.product !== 'وقت الراحة' && seg.status !== 'Cancelled') {
          let sM = timeToMins(seg.startTime);
          let eM = seg.endTime ? timeToMins(seg.endTime) : timeToMins(Utilities.formatDate(new Date(), tz, 'HH:mm'));
          if (eM < sM) eM += 1440; // Midnight rollover
          
          for (let m = sM; m < eM; m++) {
             // Ensure this minute doesn't overlap with a global break
             let isBreak = breaksTimeline.some(b => m >= timeToMins(b.startTime) && m < (b.endTime ? timeToMins(b.endTime) : 1440));
             if (!isBreak) minuteMap[m % 1440]++;
          }
       }
    });

    // PASS 2: Assemble Complete Multi-Day History for Master Cards
    const masterBatches = {};
    for (let i = 0; i < data.length; i++) {
       let masterId = data[i][0];
       if (touchedMasterIds.has(masterId)) {
          if (!masterBatches[masterId]) {
             masterBatches[masterId] = {
                masterId: masterId,
                product: data[i][4],
                totalQuantity: 0,
                finalStatus: 'Active',
                finalImage: '',
                history: [],
                metrics: { operationalDur: 0, financialDur: 0, concurrencyCounts: { "1": 0, "2": 0, "3": 0, "4+": 0 } }
             };
          }
          
          let rowDateStr = data[i][3] instanceof Date ? Utilities.formatDate(new Date(data[i][3]), tz, 'yyyy-MM-dd') : String(data[i][3]).trim();
          let sTime = data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], tz, 'HH:mm') : String(data[i][5] || '');
          let eTime = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], tz, 'HH:mm') : String(data[i][6] || '');
          let dur = parseFloat(data[i][7]) || 0;
          let qty = parseFloat(data[i][8]) || 0;
          let status = data[i][12];
          
          let combinedNotes = [notesData[i][5], notesData[i][6], notesData[i][8], notesData[i][10], notesData[i][11]].filter(n => n !== '').join('\n');
          
          // Push Segment to History
          masterBatches[masterId].history.push({
             segmentId: data[i][13] || masterId,
             date: rowDateStr,
             startTime: sTime,
             endTime: eTime,
             duration: dur,
             quantityAllocated: qty,
             status: status,
             notes: data[i][10],
             editHistory: combinedNotes
          });
          
          // Update Master Batch Global Variables
          masterBatches[masterId].metrics.operationalDur += dur;
          masterBatches[masterId].totalQuantity += qty;
          masterBatches[masterId].finalStatus = status; // Will overwrite sequentially to latest status
          if (data[i][11]) masterBatches[masterId].finalImage = data[i][11]; // Grab latest uploaded image
          
          // PASS 3: Calculate Diluted Financial Math & Concurrency for TODAY's segments
          if (rowDateStr === todayStr && status !== 'Cancelled') {
             let sM = timeToMins(sTime);
             let eM = eTime ? timeToMins(eTime) : timeToMins(Utilities.formatDate(new Date(), tz, 'HH:mm'));
             if (eM < sM) eM += 1440;
             
             for (let m = sM; m < eM; m++) {
                let isBreak = breaksTimeline.some(b => m >= timeToMins(b.startTime) && m < (b.endTime ? timeToMins(b.endTime) : 1440));
                if (!isBreak) {
                   let concurrentActive = minuteMap[m % 1440];
                   if (concurrentActive > 0) {
                      masterBatches[masterId].metrics.financialDur += (1 / concurrentActive);
                      
                      // Log concurrency context
                      if (concurrentActive === 1) masterBatches[masterId].metrics.concurrencyCounts["1"]++;
                      else if (concurrentActive === 2) masterBatches[masterId].metrics.concurrencyCounts["2"]++;
                      else if (concurrentActive === 3) masterBatches[masterId].metrics.concurrencyCounts["3"]++;
                      else masterBatches[masterId].metrics.concurrencyCounts["4+"]++;
                   }
                }
             }
          }
       }
    }

    // Convert Object to Array and Reverse (Newest First)
    const masterCardsArray = Object.values(masterBatches).reverse();

    return {
       timeline: timelineSegments,
       masterCards: masterCardsArray
    };
    
  } catch (e) { 
    throw new Error(e.message);
  }
}
// ============================================================================

function editProdSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const table = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
    
    let targetRowIndex = -1;
    let masterIdForRipple = '';
    let targetData = [];

    // 1. Locate Target Row (Segment ID [Col 13] or Master ID [Col 0])
    for (let i = table.length - 1; i >= 0; i--) {
       if (table[i][13] === data.targetId || table[i][0] === data.targetId) {
           targetRowIndex = i + 2;
           targetData = table[i];
           masterIdForRipple = table[i][0];
           break;
       }
    }
    if (targetRowIndex === -1) throw new Error('تعذر العثور على السجل في قاعدة البيانات.');

    const tz = Session.getScriptTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
    
    // 2. Apply Time Edits (Segment Level)
    if (data.editTime) {
       let oldSTime = targetData[5] instanceof Date ? Utilities.formatDate(targetData[5], tz, 'HH:mm') : String(targetData[5] || '');
       sheet.getRange(targetRowIndex, 6).setValue(data.newStartTime);
       if (oldSTime !== data.newStartTime) addCellNote(sheet, targetRowIndex, 6, `[${timestamp}] ${data.user}: تعديل وقت البدء إلى ${data.newStartTime}`);
       
       if (data.newEndTime && targetData[12] !== 'Active') {
           let oldETime = targetData[6] instanceof Date ? Utilities.formatDate(targetData[6], tz, 'HH:mm') : String(targetData[6] || '');
           sheet.getRange(targetRowIndex, 7).setValue(data.newEndTime);
           
           const newDuration = calculateDuration(data.newStartTime, data.newEndTime);
           sheet.getRange(targetRowIndex, 8).setValue(newDuration);
           
           if (oldETime !== data.newEndTime) addCellNote(sheet, targetRowIndex, 7, `[${timestamp}] ${data.user}: تعديل وقت الإنهاء إلى ${data.newEndTime}`);
       }
    }

    // 3. Apply Notes & Image Edits (Universal for both Segment and Master)
    let oldNotes = targetData[10];
    sheet.getRange(targetRowIndex, 11).setValue(data.newNotes || '');
    if (String(oldNotes) !== String(data.newNotes)) addCellNote(sheet, targetRowIndex, 11, `[${timestamp}] ${data.user}: تم تعديل الملاحظات`);
    
    if (data.newImageUrl) {
       sheet.getRange(targetRowIndex, 12).setValue(data.newImageUrl);
       addCellNote(sheet, targetRowIndex, 12, `[${timestamp}] ${data.user}: تم تحديث الصورة`);
    }

    // 4. RIPPLE EFFECT: Proportional Math for Production Batches
    if (targetData[4] !== 'وقت مهدر' && targetData[4] !== 'وقت الراحة') {
       const updatedTable = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
       let segments = [];
       let totalBatchDur = 0;
       let globalQty = data.editQuantity ? data.newTotalQuantity : 0;
       
       for (let i = 0; i < updatedTable.length; i++) {
           let stat = updatedTable[i][12];
           if (updatedTable[i][0] === masterIdForRipple && (stat === 'Paused' || stat === 'Resumed' || stat === 'Completed' || stat === 'Active')) {
               segments.push(i + 2);
               totalBatchDur += (parseFloat(updatedTable[i][7]) || 0); // Accumulate updated durations
               if (!data.editQuantity) globalQty += (parseFloat(updatedTable[i][8]) || 0); // Retain existing global qty if not editing it
           }
       }

       if (totalBatchDur > 0 && globalQty > 0) {
           segments.forEach(rIdx => {
               let sDur = parseFloat(sheet.getRange(rIdx, 8).getValue()) || 0;
               let proportion = sDur / totalBatchDur;
               let sQty = parseFloat((proportion * globalQty).toFixed(1));
               let sAvg = sQty > 0 ? (sDur / sQty).toFixed(2) : 0;

               sheet.getRange(rIdx, 9).setValue(sQty);
               sheet.getRange(rIdx, 10).setValue(sAvg);
           });
           if (data.editQuantity) addCellNote(sheet, targetRowIndex, 9, `[${timestamp}] ${data.user}: إعادة توزيع إجمالي الدفعة (${globalQty})`);
       }
    }

    // 5. SYNC WASTE & JSON ACROSS ALL AFFECTED DAYS
    let affectedDates = new Set();
    
    // Always add the base date of the specific row being edited
    let baseDate = targetData[3] instanceof Date ? Utilities.formatDate(targetData[3], tz, 'yyyy-MM-dd') : String(targetData[3]).trim();
    if (baseDate) affectedDates.add(baseDate);

    // If this is a multi-segment Master Batch, scan the original table for ANY other dates it touched
    if (masterIdForRipple && targetData[4] !== 'وقت مهدر' && targetData[4] !== 'وقت الراحة') {
       for (let i = 0; i < table.length; i++) {
           if (table[i][0] === masterIdForRipple) {
               let dStr = table[i][3] instanceof Date ? Utilities.formatDate(table[i][3], tz, 'yyyy-MM-dd') : String(table[i][3]).trim();
               if (dStr) affectedDates.add(dStr);
           }
       }
    }

    // Trigger the recalculation and ledger sync for EVERY affected day
    affectedDates.forEach(dateStr => {
       syncWasteTimes(dateStr);
       triggerRetroactiveSync(dateStr, data.user);
    });
    
    return { success: true };
  } catch (e) { throw new Error(e.message); }
}
function calculateDuration(startStr, endStr) {
  let sStr = startStr instanceof Date ? Utilities.formatDate(startStr, Session.getScriptTimeZone(), 'HH:mm') : String(startStr);
  let eStr = endStr instanceof Date ? Utilities.formatDate(endStr, Session.getScriptTimeZone(), 'HH:mm') : String(endStr);
  const s = sStr.split(':');
  const e = eStr.split(':');
  const d1 = new Date(); d1.setHours(parseInt(s[0] || 0), parseInt(s[1] || 0), 0);
  const d2 = new Date(); d2.setHours(parseInt(e[0] || 0), parseInt(e[1] || 0), 0);
  let diff = Math.round((d2 - d1) / 60000);
  if (diff < 0) diff += 24 * 60; // Over midnight
  return diff;
}
function addCellNote(sheet, row, col, newNoteLine) {
  const cell = sheet.getRange(row, col);
  const existingNote = cell.getNote();
  const fullNote = existingNote ? existingNote + '\n' + newNoteLine : newNoteLine;
  cell.setNote(fullNote);
}
// --- NEW TIMELINE ENGINE & BREAK LOGIC ---
function addBreakSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const sessionId = 'BRK-' + new Date().getTime();
    const duration = data.endTime ? calculateDuration(data.startTime, data.endTime) : '';
    const status = data.endTime ? 'Completed' : 'Active';
    const segmentId = sessionId + '-1'; // Phase 1: Initialize Segment ID
    sheet.appendRow([
      sessionId, new Date(), data.user, data.date, 'وقت الراحة',
      data.startTime, data.endTime || '', duration, '', '', data.notes || '', '', status, segmentId
    ]);
    syncWasteTimes(data.date);
    return { success: true };
  } catch (e) {
    throw new Error(e.message);
  }
}
function endBreakSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('لا توجد بيانات');
    const table = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
    let rowIndex = -1;
    // Search backward to guarantee we grab the active segment
    for (let i = table.length - 1; i >= 0; i--) {
      if (table[i][0] === data.sessionId) { rowIndex = i + 2; break; }
    }
    if (rowIndex === -1) throw new Error('لم يتم العثور على الجلسة');
    const startTimeStr = table[rowIndex - 2][5];
    const duration = calculateDuration(startTimeStr, data.endTime);

    let statusStr = 'Completed';
    if (data.action === 'cancel') statusStr = 'Cancelled';

    sheet.getRange(rowIndex, 7).setValue(data.endTime);
    sheet.getRange(rowIndex, 8).setValue(duration);

    if (data.action === 'cancel' && data.notes) {
      let existingNote = String(table[rowIndex - 2][10] || '');
      let finalNote = existingNote + (existingNote ? '\n' : '') + '[إلغاء الراحة] ' + data.notes;
      sheet.getRange(rowIndex, 11).setValue(finalNote);
    }

    sheet.getRange(rowIndex, 13).setValue(statusStr);

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    addCellNote(sheet, rowIndex, 13, `[${timestamp}] Break ${statusStr} by ${data.user}.`);

    let dateStr = table[rowIndex - 2][3] instanceof Date ? Utilities.formatDate(table[rowIndex - 2][3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(table[rowIndex - 2][3]);
    syncWasteTimes(dateStr);

    return { success: true };
  } catch (e) { throw new Error(e.message); }
}
function timeToMins(timeStr) {
  if (!timeStr) return 0;
  let p = String(timeStr).split(':');
  return parseInt(p[0] || 0) * 60 + parseInt(p[1] || 0);
}
function minsToTime(mins) {
  let h = Math.floor((mins % 1440) / 60);
  let m = (mins % 1440) % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}
function forceSyncWasteTimes(dateStr) {
  syncWasteTimes(dateStr);
  return { success: true };
}
function syncWasteTimes(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Logs');
  if (!sheet) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  let timeline = new Array(1440).fill(false);
  let todayRows = [];
  data.forEach(function (r, i) {
    let rowDate = r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[3] || '').trim();
    let status = String(r[12] || '').trim();
    // CRITICAL: Completely ignore Cancelled sessions so they correctly turn into Waste time
    if (rowDate === dateStr && status !== 'Cancelled' && status !== 'ملغي') {
      let type = (r[4] === 'وقت مهدر') ? 'Waste' : 'Busy';

      // Safely extract HH:mm regardless of how Google Sheets formatted the cell
      let sTime = r[5] instanceof Date ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), 'HH:mm') : String(r[5] || '').trim();
      let eTime = r[6] instanceof Date ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'HH:mm') : String(r[6] || '').trim();

      // Handle cases where Sheets appends seconds to string representations
      if (sTime.length > 5 && sTime.includes(':')) sTime = sTime.substring(0, 5);
      if (eTime.length > 5 && eTime.includes(':')) eTime = eTime.substring(0, 5);

      todayRows.push({ rowIndex: i + 2, type: type, start: sTime, end: eTime, status: r[12] });
    }
  });
  const shiftStart = 480; // 08:00 AM
  const shiftEnd = 990;   // 16:30 PM (4:30 PM)
  let isToday = dateStr === Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let nowTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  let nowMin = timeToMins(nowTimeStr);
  // Do not project waste into the future if the shift is still ongoing
  let latestEvalTime = isToday ? Math.min(shiftEnd, nowMin) : shiftEnd;
  if (isToday && nowMin < shiftStart) return; // Shift hasn't started yet
  todayRows.forEach(function (r) {
    if (r.type === 'Waste' || !r.start) return;
    let sMin = timeToMins(r.start);
    if (r.status === 'Active') {
      for (let m = sMin; m < nowMin; m++) timeline[m] = true;
    } else if (r.end) {
      let eMin = timeToMins(r.end);
      if (eMin < sMin) eMin += 1440;
      for (let m = sMin; m < eMin; m++) timeline[m] = true;
    }
  });
  let gaps = []; let inGap = false; let gapStart = 0;
  // Scan strictly within the shift boundaries
  for (let m = shiftStart; m <= latestEvalTime; m++) {
    if (!timeline[m] && m !== latestEvalTime) {
      if (!inGap) { inGap = true; gapStart = m; }
    } else {
      if (inGap) {
        if (m - gapStart > 0) gaps.push({ start: gapStart, end: m });
        inGap = false;
      }
    }
  }
  let wasteRows = todayRows.filter(function (r) { return r.type === 'Waste'; }).sort(function (a, b) { return timeToMins(a.start) - timeToMins(b.start); });
  let maxI = Math.max(gaps.length, wasteRows.length);
  let rowsToDelete = [];
  for (let i = 0; i < maxI; i++) {
    if (i < gaps.length && i < wasteRows.length) {
      sheet.getRange(wasteRows[i].rowIndex, 6).setValue(minsToTime(gaps[i].start));
      sheet.getRange(wasteRows[i].rowIndex, 7).setValue(minsToTime(gaps[i].end));
      sheet.getRange(wasteRows[i].rowIndex, 8).setValue(gaps[i].end - gaps[i].start);
    } else if (i < gaps.length) {
      const sessionId = 'WST-' + new Date().getTime() + '-' + i;
      const segmentId = sessionId + '-1';
      sheet.appendRow([
        sessionId, new Date(), 'النظام', dateStr, 'وقت مهدر',
        minsToTime(gaps[i].start), minsToTime(gaps[i].end), gaps[i].end - gaps[i].start,
        '', '', '', '', 'Completed', segmentId
      ]);
    } else {
      rowsToDelete.push(wasteRows[i].rowIndex);
    }
  }
  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sheet.deleteRow(r); });
}
// --- END OF DAY AUTOMATION & DAILY SUMMARY ---
function endOfDayRoutine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Logs');
  if (!sheet) return;
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  let updatedActive = false;
  // 1. Force-close any uncompleted Active sessions to 16:30
  for (let i = 0; i < data.length; i++) {
    let rowDate = data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], tz, 'yyyy-MM-dd') : String(data[i][3]).trim();
    if (rowDate === todayStr && data[i][12] === 'Active') {
      let sTime = data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], tz, 'HH:mm') : String(data[i][5]).trim();
      let dur = calculateDuration(sTime, '16:30');
      let oldNotes = data[i][10] || '';
      let newNotes = oldNotes ? oldNotes + '\nإغلاق تلقائي بواسطة النظام' : 'إغلاق تلقائي بواسطة النظام';
      let rowIndex = i + 2;
      sheet.getRange(rowIndex, 7).setValue('16:30'); // Force close at 4:30 PM
      sheet.getRange(rowIndex, 8).setValue(dur);
      sheet.getRange(rowIndex, 11).setValue(newNotes);
      sheet.getRange(rowIndex, 13).setValue('Completed');
      updatedActive = true;
    }
  }
  // 2. Synchronize Waste Times completely for the day
  syncWasteTimes(todayStr);
  // 3. Gather Updated Data for Metric Calculation
  const updatedData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  let actualStart = 1440; let actualEnd = 0;
  let totalProd = 0; let totalBreak = 0; let totalWaste = 0;
  let hasData = false;
  let endDayBreaks = [];
  updatedData.forEach(function (r) {
    let rowDate = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3]).trim();
    if (rowDate === todayStr && r[12] === 'Completed' && r[4] === 'وقت الراحة') {
      let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
      let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
      let dur = eM - sM; if (dur < 0) dur += 1440;
      endDayBreaks.push({ start: sM, end: sM + dur });
    }
  });
  updatedData.forEach(function (r) {
    let rowDate = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3]).trim();
    if (rowDate === todayStr && r[12] === 'Completed') {
      hasData = true;
      let product = r[4];
      let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
      let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
      let dur = eM - sM; if (dur < 0) dur += 1440;
      let absEM = sM + dur;
      if (sM < actualStart) actualStart = sM;
      if (eM > actualEnd) actualEnd = eM;

      if (product === 'وقت مهدر') {
        totalWaste += dur;
      } else if (product === 'وقت الراحة') {
        totalBreak += dur;
      } else {
        let netDur = dur;
        endDayBreaks.forEach(b => {
          let oStart = Math.max(sM, b.start);
          let oEnd = Math.min(absEM, b.end);
          if (oStart < oEnd) netDur -= (oEnd - oStart);
        });
        totalProd += netDur;
      }
    }
  });
  if (!hasData) return; // Exit if no activity happened today
  // Apply Flex-Time / Overtime Math
  let shiftStart = (actualStart < 480) ? actualStart : 480;
  let shiftEnd = shiftStart + 510; // 8.5 hours
  let deficit = (actualStart > 480) ? actualStart - 480 : 0;
  let extraMins = actualEnd > shiftEnd ? actualEnd - shiftEnd : 0;
  let compMins = Math.min(deficit, extraMins);
  let overtimeMins = extraMins - compMins;
  function formatDurSummary(mins) {
    if (!mins) return '0 د';
    let h = Math.floor(mins / 60);
    let m = mins % 60;
    if (h > 0 && m > 0) return h + ' س و ' + m + ' د';
    if (h > 0) return h + ' س';
    return m + ' د';
  }
  // 4. Update or Create Daily_Summary Sheet
  let summarySheet = ss.getSheetByName('Daily_Summary');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('Daily_Summary');
    summarySheet.appendRow(['التاريخ', 'وقت الإنتاج الفعلي', 'وقت الراحة', 'الوقت المهدر', 'الوقت الإضافي']);
    summarySheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f4f6');
  }
  // Check if today already exists to prevent duplicate rows
  let sumData = summarySheet.getDataRange().getValues();
  let rowIndexToUpdate = -1;
  for (let i = 1; i < sumData.length; i++) {
    let sumDate = sumData[i][0] instanceof Date ? Utilities.formatDate(sumData[i][0], tz, 'yyyy-MM-dd') : String(sumData[i][0]).trim();
    if (sumDate === todayStr) {
      rowIndexToUpdate = i + 1;
      break;
    }
  }
  let rowValues = [
    todayStr,
    formatDurSummary(totalProd),
    formatDurSummary(totalBreak),
    formatDurSummary(totalWaste),
    formatDurSummary(overtimeMins)
  ];
  if (rowIndexToUpdate > -1) {
    summarySheet.getRange(rowIndexToUpdate, 1, 1, 5).setValues([rowValues]);
  } else {
    summarySheet.appendRow(rowValues);
  }

  // 5. AUTO-SYNC: Capture snapshot into Daily_Product_Performance JSON API
  triggerRetroactiveSync(todayStr, 'النظام التلقائي (Automated)');
}
function setupDailyTrigger() {
  // 1. Clean up any existing triggers to prevent duplicates
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'endOfDayRoutine') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // 2. Create the daily trigger to run at 11:55 PM (23:55)
  ScriptApp.newTrigger('endOfDayRoutine')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(55)
    .create();
}
// --- EXECUTIVE REPORTING ENGINE (PORTAL C) ---
function getProductionDateRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Logs');
  if (!sheet) return { min: '', max: '' };
  // Fetch Dates (Col D/Index 3) and Status (Col M/Index 12)
  const data = sheet.getRange(2, 4, sheet.getLastRow() - 1, 10).getValues();
  const tz = Session.getScriptTimeZone();
  let minDate = null;
  let maxDate = null;
  data.forEach(r => {
    if (r[9] !== 'Completed') return; // Only count productive days
    let dObj = new Date(r[0]);
    if (isNaN(dObj.getTime())) return;
    if (!minDate || dObj < minDate) minDate = dObj;
    if (!maxDate || dObj > maxDate) maxDate = dObj;
  });
  return {
    min: minDate ? Utilities.formatDate(minDate, tz, 'yyyy-MM-dd') : '',
    max: maxDate ? Utilities.formatDate(maxDate, tz, 'yyyy-MM-dd') : ''
  };
}
function getProductionReportData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Logs');
  if (!sheet) throw new Error("Sheet not found");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
  const tz = Session.getScriptTimeZone();
  let kpis = { prod: 0, break: 0, waste: 0, overtime: 0 };
  let prodMap = {};
  let trendMap = {};
  let productsSet = new Set();
  let usersSet = new Set();
  let dayGroups = {};

  // 1. Filter & Group Chronologically
  data.forEach(r => {
    if (r[12] !== 'Completed') return; // Only process finished sessions

    let dateStr = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3]).trim();
    let product = r[4];
    let user = r[2];

    if (product !== 'وقت مهدر' && product !== 'وقت الراحة') productsSet.add(product);
    if (user && user !== 'النظام') usersSet.add(user);

    if (filters.startDate && dateStr < filters.startDate) return;
    if (filters.endDate && dateStr > filters.endDate) return;

    if (filters.user && filters.user !== 'الكل' && user !== filters.user) return;

    if (filters.product && filters.product !== 'الكل') {
      if (product !== filters.product && product !== 'وقت مهدر' && product !== 'وقت الراحة') return;
      // Hide general waste/breaks if a specific product is filtered to avoid skewed data
      if (product === 'وقت مهدر' || product === 'وقت الراحة') return;
    }

    if (!dayGroups[dateStr]) dayGroups[dateStr] = [];
    dayGroups[dateStr].push(r);
  });
  // 2. Process Daily Flex-Time Logic
  for (let d in dayGroups) {
    let dayRows = dayGroups[d];
    let actualStart = 1440; let actualEnd = 0;
    let dProd = 0; let dBreak = 0; let dWaste = 0;
    // Collect breaks for the day
    let dayBreaks = [];
    dayRows.forEach(r => {
      if (r[4] === 'وقت الراحة') {
        let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
        let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
        let dur = eM - sM; if (dur < 0) dur += 1440;
        dayBreaks.push({ start: sM, end: sM + dur });
      }
    });

    let prodSessions = [];
    let minuteMap = new Array(1440).fill(null).map(() => []);

    dayRows.forEach(r => {
      let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
      let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
      let dur = eM - sM; if (dur < 0) dur += 1440;
      let absEM = sM + dur;
      let product = r[4];

      if (sM < actualStart) actualStart = sM;
      if (eM > actualEnd) actualEnd = eM;

      if (product === 'وقت مهدر') {
        dWaste += dur; kpis.waste += dur;
      } else if (product === 'وقت الراحة') {
        dBreak += dur; kpis.break += dur;
      } else {
        prodSessions.push({
          product: product,
          start: sM,
          end: absEM,
          qty: parseFloat(r[8]) || 0,
          allocatedDur: 0
        });
      }
    });

    // Build the minute map for valid production times
    prodSessions.forEach((session, idx) => {
      for (let m = session.start; m < session.end; m++) {
        let minuteOfDay = m % 1440;

        // Check if this minute overlaps with ANY break (ensuring breaks don't count towards production time)
        let isBreak = false;
        for (let b of dayBreaks) {
          if (m >= b.start && m < b.end) {
            isBreak = true; break;
          }
        }

        if (!isBreak) {
          minuteMap[minuteOfDay].push(idx);
        }
      }
    });

    // Calculate fractional durations and absolute factory production time
    for (let m = 0; m < 1440; m++) {
      let activeCount = minuteMap[m].length;
      if (activeCount > 0) {
        dProd += 1; // Exactly 1 minute of net factory uptime
        kpis.prod += 1;

        // Dilute the minute equally across all products running at that exact time
        let fraction = 1 / activeCount;
        minuteMap[m].forEach(idx => {
          prodSessions[idx].allocatedDur += fraction;
        });
      }
    }

    // Aggregate the perfectly balanced times back into the product maps
    prodSessions.forEach(session => {
      let p = session.product;
      if (!prodMap[p]) prodMap[p] = { qty: 0, dilutedDur: 0, rawDur: 0, breakdown: {} };
      prodMap[p].qty += session.qty;
      prodMap[p].dilutedDur += session.allocatedDur;
      prodMap[p].rawDur += (session.end - session.start); // Add raw physical duration

      if (!prodMap[p].breakdown[d]) prodMap[p].breakdown[d] = { qty: 0, dilutedDur: 0, rawDur: 0 };
      prodMap[p].breakdown[d].qty += session.qty;
      prodMap[p].breakdown[d].dilutedDur += session.allocatedDur;
      prodMap[p].breakdown[d].rawDur += (session.end - session.start);
    });

    let shiftStart = (actualStart < 480) ? actualStart : 480;
    let shiftEnd = shiftStart + 510;
    let deficit = (actualStart > 480) ? actualStart - 480 : 0;
    let extraMins = actualEnd > shiftEnd ? actualEnd - shiftEnd : 0;
    let compMins = Math.min(deficit, extraMins);
    let overtimeMins = extraMins - compMins;

    // RULE: If NO actual production happened, ignore waste and overtime for this day
    if (dProd === 0) {
      kpis.waste -= dWaste; // Remove this day's waste from the global KPI tally
      dWaste = 0;           // Zero it out for the daily chart trend
      overtimeMins = 0;     // Zero out daily overtime
    }

    kpis.overtime += overtimeMins;
    trendMap[d] = { date: d, prod: dProd, waste: dWaste, break: dBreak, ot: overtimeMins };
  }
  let trendArray = Object.values(trendMap).sort((a, b) => a.date.localeCompare(b.date));
  let productsArray = Object.keys(prodMap).map(p => {
    let d = prodMap[p];
    
    // NEW DUAL MATH:
    // Operational Speed strictly uses Raw Duration. Financial Cost uses Diluted Duration.
    let opAvgTime = d.qty > 0 ? (d.rawDur / d.qty).toFixed(2) : 0;
    let finLoadedTime = d.qty > 0 ? (d.dilutedDur / d.qty).toFixed(2) : 0;
    
    return { 
      name: p, 
      qty: d.qty, 
      dur: d.dilutedDur, // Keep 'dur' tied to financial for backward compatibility
      rawDur: d.rawDur,
      dilutedDur: d.dilutedDur,
      opAvg: opAvgTime,
      finAvg: finLoadedTime,
      breakdown: d.breakdown 
    };
  }).sort((a, b) => b.dur - a.dur);
  return {
    kpis: kpis,
    trend: trendArray,
    products: productsArray,
    productNames: Array.from(productsSet).sort(),
    users: Array.from(usersSet).sort()
  };
}

// --- RETROACTIVE PERFORMANCE & JSON SYNC ENGINE ---

function formatDurationBackend(mins) {
  if (!mins || isNaN(mins) || mins === Infinity) return '00:00:00';
  let totalSeconds = Math.round(mins * 60);
  let h = Math.floor(totalSeconds / 3600);
  let m = Math.floor((totalSeconds % 3600) / 60);
  let s = totalSeconds % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m + ':' + (s < 10 ? '0' : '') + s;
}

function triggerRetroactiveSync(startDateStr, triggerSource) {
  try {
    let tz = Session.getScriptTimeZone();
    let todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    let currentDate = new Date(startDateStr);
    let endDate = new Date(todayStr);

    if (isNaN(currentDate.getTime()) || isNaN(endDate.getTime())) throw new Error("تاريخ غير صالح");

    // Ripple Effect: Sync the edited day and every single day after it up to today
    while (currentDate <= endDate) {
      let dStr = Utilities.formatDate(currentDate, tz, 'yyyy-MM-dd');
      syncPerformanceData(dStr, triggerSource);
      currentDate.setDate(currentDate.getDate() + 1);
    }
    return { success: true, message: 'تمت مزامنة بيانات الأداء بنجاح!' };
  } catch (e) {
    throw new Error("خطأ في المزامنة: " + e.message);
  }
}

function syncPerformanceData(dateStr, triggerSource) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Daily_Product_Performance');
  if (!sheet) {
    sheet = ss.insertSheet('Daily_Product_Performance');
    sheet.appendRow(['التاريخ', 'المنتج', 'الكمية المنتجة', 'وقت الإنتاج الصافي', 'متوسط وقت الإنتاج التشغيلي', 'الوقت المالي المخفف', 'الوقت المحمل للوحدة', 'نسبة التكلفة المجمعة', 'وزن تكلفة الوحدة', 'Daily JSON', 'Monthly JSON', 'ملف التزامن', 'وقت الحفظ', 'بواسطة']);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#f3f4f6');
  }

  // 1. Get Daily Data via your existing bulletproof reporting engine
  let report = getProductionReportData({ startDate: dateStr, endDate: dateStr, product: 'الكل', user: 'الكل' });
  let totalFactoryTime = report.kpis.prod;

  // 2. Delete existing rows for dateStr (to allow clean Upsert)
  let data = sheet.getDataRange().getValues();
  let rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(data[i][0]).trim();
    if (rowDate === dateStr) rowsToDelete.push(i + 1);
  }
  rowsToDelete.reverse().forEach(r => sheet.deleteRow(r));

  // 3. Prepare new rows & JSONs
  let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  let newRows = [];
  let startOfMonth = dateStr.substring(0, 8) + '01'; // 'YYYY-MM-01'

  // Pre-fetch MTD Factory Time to avoid loop queries
  let mtdFactoryReport = getProductionReportData({ startDate: startOfMonth, endDate: dateStr, product: 'الكل', user: 'الكل' });
  let mtdTotalFactoryTime = mtdFactoryReport.kpis.prod;

  // We need the Unified Data to grab the exact concurrency breakdown JSON
  let unifiedData = getUnifiedProdData(dateStr);
  
  report.products.forEach(p => {
    let aggCostPct = totalFactoryTime > 0 ? (p.dilutedDur / totalFactoryTime) * 100 : 0;
    let unitWeightPct = p.qty > 0 ? (aggCostPct / p.qty) : 0;
    let loadedTimePerUnit = p.qty > 0 ? (p.dilutedDur / p.qty) : 0;
    let operationalTimePerUnit = p.qty > 0 ? (p.rawDur / p.qty) : 0;

    // Find the master card to extract the concurrency profile
    let masterCard = unifiedData.masterCards.find(m => m.product === p.name);
    let concurrencyProfileJSON = masterCard && masterCard.metrics ? JSON.stringify(masterCard.metrics.concurrencyCounts) : '{}';

    let dailyJson = {
      date: dateStr,
      product: p.name,
      qty: p.qty,
      rawProdTimeMins: parseFloat(p.rawDur.toFixed(2)),
      dilutedProdTimeMins: parseFloat(p.dilutedDur.toFixed(2)),
      opTimePerUnitMins: parseFloat(operationalTimePerUnit.toFixed(2)),
      finTimePerUnitMins: parseFloat(loadedTimePerUnit.toFixed(2)),
      aggCostPct: parseFloat(aggCostPct.toFixed(4)),
      unitWeightPct: parseFloat(unitWeightPct.toFixed(6))
    };

    let mtdProduct = mtdFactoryReport.products.find(x => x.name === p.name);
    let mtdJson = {
      month: dateStr.substring(0, 7),
      product: p.name,
      mtdQty: mtdProduct ? mtdProduct.qty : 0,
      mtdRawProdTimeMins: mtdProduct ? parseFloat(mtdProduct.rawDur.toFixed(2)) : 0,
      mtdDilutedProdTimeMins: mtdProduct ? parseFloat(mtdProduct.dilutedDur.toFixed(2)) : 0,
      mtdTotalFactoryTimeMins: mtdTotalFactoryTime
    };

    newRows.push([
      dateStr,
      p.name,
      p.qty,
      formatDurationBackend(p.rawDur),        // D: Raw Operational Time
      formatDurationBackend(operationalTimePerUnit), // E: Avg Operational Time / Unit
      formatDurationBackend(p.dilutedDur),    // F: Financial Diluted Time
      formatDurationBackend(loadedTimePerUnit),// G: Financial Time / Unit
      aggCostPct.toFixed(2) + '%',
      unitWeightPct.toFixed(4) + '%',
      JSON.stringify(dailyJson),
      JSON.stringify(mtdJson),
      concurrencyProfileJSON,                 // L: Concurrency State Array
      timestamp,
      triggerSource
    ]);
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 14).setValues(newRows);
  }
}

// --- END OF SHIFT (HANDOVER) ENGINE ---

function getLastTankNumber() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Strictly target the active 'Tank' sheet
    const tankSheet = ss.getSheetByName('Tank');
    if (!tankSheet || tankSheet.getLastRow() < 2) return 0;

    // Fetch exclusively Column M (Column Index 13)
    const data = tankSheet.getRange(2, 13, tankSheet.getLastRow() - 1, 1).getValues();
    
    let maxSerial = 0;
    // Scan all rows to find the absolute highest registered serial number
    for (let i = 0; i < data.length; i++) {
      let val = parseInt(data[i][0]);
      if (!isNaN(val) && val > maxSerial) {
        maxSerial = val;
      }
    }
    return maxSerial;
  } catch(e) {
    return 0;
  }
}

function processEndOfShift(inputNumberStr, imageBase64) {
  try {
    const lastTank = getLastTankNumber();
    const expected = lastTank + 1;
    const input = parseInt(inputNumberStr);

    if (isNaN(input)) {
      return { success: false, message: "الرقم المدخل غير صالح. يرجى إدخال أرقام فقط." };
    }

    // Sequence Verification Logic (Arabic Only)
    if (input > expected) {
      return { success: false, message: `الرقم المدخل (${input}) أكبر من المتوقع (${expected}).\n\nيبدو أن هناك ملصق مفقود! يرجى المراجعة.` };
    } else if (input < expected) {
      return { success: false, message: `الرقم المدخل (${input}) أصغر من المتوقع (${expected}).\n\nيبدو أن هناك تانك لم يتم تسجيله في النظام أو تم تكرار الرقم! يرجى المراجعة.` };
    }

    // Exact Match Success - Save Data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let handoverSheet = ss.getSheetByName("Shift_Handovers");
    if (!handoverSheet) {
      handoverSheet = ss.insertSheet("Shift_Handovers");
      handoverSheet.appendRow(["وقت التسليم", "آخر تانك", "الرقم القادم المتوقع", "رابط الصورة"]);
      handoverSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#f3f4f6");
    }

    // Save Image to Drive
    let imageUrl = "";
    if (imageBase64) {
      const folderName = "Shift_Handovers_Images";
      let folders = DriveApp.getFoldersByName(folderName);
      let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      
      let base64Data = imageBase64.split(',')[1];
      let blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'Handover_Next_' + expected + '_' + new Date().getTime() + '.jpg');
      let file = folder.createFile(blob);
      imageUrl = file.getUrl();
    }

    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    handoverSheet.appendRow([timestamp, lastTank, expected, imageUrl]);

    return { success: true, message: "تطابق التسلسل! تم إنهاء الوردية وتسليم العهدة بنجاح." };
  } catch(e) {
    return { success: false, message: "خطأ في النظام: " + e.message };
  }
}

function getLatestHandover() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ws = ss.getSheetByName("Shift_Handovers");
    if (!ws || ws.getLastRow() < 2) return null;
    
    // Get the absolute last row saved
    const lastRow = ws.getRange(ws.getLastRow(), 1, 1, 4).getValues()[0];
    return {
      time: String(lastRow[0]),
      number: lastRow[2], 
      image: String(lastRow[3])
    };
  } catch(e) { return null; }
}

// ============================================================================
// --- PHASE 6: ONE-TIME HISTORICAL DATA MIGRATION SCRIPT ---
// ============================================================================

function runHistoricalDataMigration() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Production_Logs');
    if (!logSheet || logSheet.getLastRow() < 2) {
      return "عذراً، لم يتم العثور على أي بيانات في جدول Production_Logs.";
    }
    
    const lastRow = logSheet.getLastRow();
    const range = logSheet.getRange(2, 1, lastRow - 1, 14);
    const values = range.getValues();
    
    const tz = Session.getScriptTimeZone();
    const sessionCounters = {};
    const uniqueDates = new Set();
    let segmentsUpdated = 0;
    
    // PASS 1: Assign Retroactive Segment IDs and Collect Unique Logged Dates
    for (let i = 0; i < values.length; i++) {
      let sessionId = values[i][0];
      if (!sessionId) continue;
      
      // Extract the chronological date string
      let dateStr = values[i][3] instanceof Date 
        ? Utilities.formatDate(values[i][3], tz, 'yyyy-MM-dd') 
        : String(values[i][3]).trim();
        
      if (dateStr && dateStr !== 'وقت مهدر' && dateStr !== 'وقت الراحة' && values[i][4]) {
        uniqueDates.add(dateStr);
      }
      
      // Track serialized counters for segment grouping
      if (!sessionCounters[sessionId]) {
        sessionCounters[sessionId] = 0;
      }
      sessionCounters[sessionId]++;
      
      // If Column N (Segment ID) is completely blank, retro-assign identity
      if (!values[i][13] || String(values[i][13]).trim() === "") {
        values[i][13] = sessionId + '-' + sessionCounters[sessionId];
        segmentsUpdated++;
      }
    }
    
    // Save generated segment IDs back to Column N in bulk execution
    if (segmentsUpdated > 0) {
      const segmentIdColumnValues = values.map(row => [row[13]]);
      logSheet.getRange(2, 14, segmentIdColumnValues.length, 1).setValues(segmentIdColumnValues);
    }
    
    // PASS 2: Sequential Backfill via the Live Dual-Math Engine
    let datesArray = Array.from(uniqueDates).sort();
    let datesSynced = 0;
    
    datesArray.forEach(dateStr => {
      try {
        // Triggers full deletion of old row + complete recalculation of the 14 columns
        syncPerformanceData(dateStr, 'الهجرة التاريخية (Data Migration)');
        datesSynced++;
      } catch (err) {
        Logger.log(`خطأ أثناء معالجة تاريخ ${dateStr}: ${err.message}`);
      }
    });
    
    return `🎉 تمت عملية الهجرة وتحديث البيانات التاريخية بنجاح!\n\n` +
           `- تم إنشاء معرّفات تشغيل جزئية (Segment IDs): ${segmentsUpdated} سجل.\n` +
           `- تم إعادة احتساب ومزامنة الأداء المالي والتشغيلي لـ: ${datesSynced} يومًا بالكامل.`;
           
  } catch (e) {
    throw new Error("فشلت عملية الهجرة: " + e.message);
  }
}
