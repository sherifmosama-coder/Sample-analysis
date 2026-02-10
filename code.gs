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

// Check password
function checkPassword(portal, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const systemSheet = ss.getSheetByName('System data');
    
    if (!systemSheet) {
      throw new Error('Sheet "System data" not found');
    }
    
    let correctPassword = '';
    
    if (portal === 'A') {
      correctPassword = systemSheet.getRange('A2').getValue().toString();
    } else if (portal === 'B') {
      correctPassword = systemSheet.getRange('B2').getValue().toString();
    }
    
    return password === correctPassword;
  } catch (error) {
    throw new Error('Error checking password: ' + error.message);
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
    return options.map(function(row) { return row[0]; }).filter(function(val) { return val !== ''; });
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
    return products.map(function(row) { return row[0]; }).filter(function(val) { return val !== ''; });
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
    uniqueSerials.sort(function(a, b) { return b - a; });
    
    return uniqueSerials;
  } catch (error) {
    throw new Error('Error loading tank serials: ' + error.message);
  }
}

function getSheetSerials(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const serials = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  return serials.map(function(row) { return row[0]; })
                .filter(function(val) { return val !== '' && !isNaN(val); });
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

// [UPDATED] Update tank analysis (Checks 2026 then 2025)
function updateTankAnalysis(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Try updating NEW sheet
    const newTankSheet = ss.getSheetByName('Tank');
    if (newTankSheet) {
      if (tryUpdateSheet(newTankSheet, data)) {
        return { success: true, message: 'تم تحديث البيانات بنجاح (2026)' };
      }
    }
    
    // 2. Try updating OLD sheet
    const oldTankSheet = ss.getSheetByName('Tank_2025');
    if (oldTankSheet) {
      if (tryUpdateSheet(oldTankSheet, data)) {
        return { success: true, message: 'تم تحديث البيانات بنجاح (أرشيف 2025)' };
      }
    }
    
    throw new Error('Serial number not found');
  } catch (error) {
    throw new Error('Error updating tank: ' + error.message);
  }
}

// Helper function to perform the update
function tryUpdateSheet(sheet, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  
  const serials = sheet.getRange(2, 13, lastRow - 1, 1).getValues();
  
  for (let i = 0; i < serials.length; i++) {
    if (serials[i][0] == data.serialNumber) {
      const rowNum = i + 2;
      
      if (data.analysisMethod === 'device') {
        sheet.getRange(rowNum, 12).setValue(data.deviceReading);
        sheet.getRange(rowNum, 15).setValue(data.vvPercent);
        sheet.getRange(rowNum, 17).setValue(new Date());
      } else {
        sheet.getRange(rowNum, 16).setValue(data.labResult / 100);
        sheet.getRange(rowNum, 18).setValue(new Date());
      }
      return true; // Update successful
    }
  }
  return false; // Serial not found in this sheet
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
    row[2] = data.analysisMethod === 'device' ? 'الجهاز' : 'معملي';
    
    if (data.analysisMethod === 'device') {
      row[3] = data.deviceReading;
      row[4] = data.vvPercent;
    } else {
      row[5] = data.labResult / 100;
    }
    
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

// Validate serial number
function validateSerialNumber(serialNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tankSheet = ss.getSheetByName('Tank');
    
    if (!tankSheet) {
      return { valid: serialNumber === 1, message: serialNumber === 1 ? '' : 'رقم العينة غير صحيح' };
    }
    
    const lastRow = tankSheet.getLastRow();
    
    if (lastRow < 2) {
      return { valid: serialNumber === 1, message: serialNumber === 1 ? '' : 'رقم العينة يجب أن يبدأ من 1' };
    }
    
    const existingSerials = tankSheet.getRange(2, 13, lastRow - 1, 1).getValues().flat();
    
    if (existingSerials.includes(serialNumber)) {
      return { valid: false, message: 'رقم العينة مكرر' };
    }
    
    const maxSerial = Math.max.apply(null, existingSerials.filter(function(n) { return typeof n === 'number' && !isNaN(n); }));
    
    if (serialNumber !== maxSerial + 1) {
      return { valid: false, message: 'رقم العينة يجب أن يكون ' + (maxSerial + 1) };
    }
    
    return { valid: true, message: '' };
  } catch (error) {
    throw new Error('Error validating serial: ' + error.message);
  }
}

// Calculate v/v% from pH (returns as fraction, not percentage)
function calculateVV(pH) {
  try {
    const Ka = 0.000018;
    const molarMass = 60.05;
    
    const hPlus = Math.pow(10, -pH);
    
    const molarity = hPlus + (Math.pow(hPlus, 2) / Ka);
    
    const percentage = (molarity * molarMass) / 10;
    
    const fraction = percentage / 100;
    
    return fraction.toFixed(6);
  } catch (error) {
    throw new Error('Error calculating v/v%: ' + error.message);
  }
}

// Save Portal A data (with external spreadsheet copy)
function savePortalAData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tankSheet = ss.getSheetByName('Tank');
    
    if (!tankSheet) {
      tankSheet = ss.insertSheet('Tank');
      const headers = [
        'Timestamp', 'User Email', 'التاريخ', '', 'نوع الخامة', 'طريقة التحضير',
        'العدد (A2)', 'العدد (A3)', 'العدد (A4)', 'العدد (A5)',
        'قراءة الجهاز', '', 'Serial Number', 'v/v%'
      ];
      tankSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      tankSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    const row = new Array(14).fill('');
    
    row[0] = new Date();
    row[1] = Session.getActiveUser().getEmail();
    row[2] = 'تاريخ اليوم';
    row[3] = '';
    row[4] = data.material;
    row[5] = data.preparation || '';
    
    const quantityCol = data.materialIndex + 4;
    row[quantityCol] = data.quantity;
    
    // For blossom/rose: save blank values for readings, serial, and v/v%
    if (data.isBlossomOrRose) {
      row[10] = '';  // Device reading blank
      row[11] = '';
      row[12] = '';  // Serial number blank
      row[13] = '';  // v/v% blank
    } else {
      row[10] = data.deviceReading;
      row[11] = '';
      row[12] = data.serialNumber;
      row[13] = data.calculatedValue;
    }
    
    // Save to main Tank sheet
    tankSheet.appendRow(row);
    
    // Save copy to external spreadsheet (including blossom/rose)
    try {
      saveToExternalSpreadsheet(row);
    } catch (externalError) {
      Logger.log('Error saving to external spreadsheet: ' + externalError.message);
      // Continue even if external save fails - don't block main save
    }
    
    return { success: true, message: 'تم حفظ البيانات بنجاح' };
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
  allData.forEach(function(row) {
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
        labResult: row[15] ? (parseFloat(row[15]) * 100).toFixed(2) + '%' : '-'
      });
    }
  });
  
  filtered.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
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
    
    data.forEach(function(row) {
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
  
  sheetsToCheck.forEach(function(sheetName) {
    const tankSheet = ss.getSheetByName(sheetName);
    if (!tankSheet || tankSheet.getLastRow() < 2) return;

    const tankData = tankSheet.getRange(2, 1, tankSheet.getLastRow() - 1, 18).getValues();
    
    tankData.forEach(function(row) {
      const serial = row[12];
      
      // A. Check for Device Analysis on Tank
      const deviceTime = row[16] ? new Date(row[16]) : null;
      if (deviceTime && deviceTime >= startDate && deviceTime <= endDate) {
        filtered.push({
          date: Utilities.formatDate(deviceTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
          product: 'Tank ' + serial,  // Display Tank Serial as Product Name
          method: 'الجهاز',
          deviceReading: row[10] ? parseFloat(row[10]).toFixed(2) : 'N/A',
          vv: row[13] ? (parseFloat(row[13]) * 100).toFixed(2) + '%' : 'N/A',
          labResult: 'N/A',
          refNumber: serial,
          imageUrl: '',
          type: 'tank'
        });
      }

      // B. Check for Lab Analysis on Tank (THIS IS THE MISSING PART)
      const labTime = row[17] ? new Date(row[17]) : null;
      if (labTime && labTime >= startDate && labTime <= endDate) {
        filtered.push({
          date: Utilities.formatDate(labTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
          product: 'Tank ' + serial, // Display Tank Serial as Product Name
          method: 'معملي',
          deviceReading: 'N/A',
          vv: 'N/A',
          labResult: row[15] ? (parseFloat(row[15]) * 100).toFixed(2) + '%' : 'N/A', // Column P (Index 15)
          refNumber: serial,
          imageUrl: '',
          type: 'tank'
        });
      }
    });
  });

  // Sort by Date Descending
  filtered.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });
  
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
      // UPDATED: Added 'نتيجة المعمل' to headers
      const headers = [['التاريخ', 'رقم العينة', 'نوع الخامة', 'طريقة التحضير', 'العدد', 'قراءة الجهاز', 'v/v%', 'نتيجة المعمل']];
      
      // UPDATED: Range width increased to 8
      reportSheet.getRange(4, 1, 1, 8).setValues(headers);
      reportSheet.getRange(4, 1, 1, 8).setFontWeight('bold').setBackground('#3b82f6').setFontColor('#ffffff');
      
      if (reportData.length > 0) {
        const rows = reportData.map(function(row) {
          return [
            row.date,
            row.serial,
            row.material,
            row.preparation || '-',
            row.quantity || '-',
            row.reading !== 'N/A' ? row.reading + ' pH' : '-',
            row.vv !== 'N/A' ? row.vv + '%' : '-',
            // NEW: Add Lab Result to row
            row.labResult
          ];
        });
        
        // UPDATED: Range width increased to 8
        reportSheet.getRange(5, 1, rows.length, 8).setValues(rows);
        reportSheet.getRange(4, 1, rows.length + 1, 8).setBorder(true, true, true, true, true, true);
      }
      
      // UPDATED: Loop 8 times for auto-resize
      for (let i = 1; i <= 8; i++) {
        reportSheet.autoResizeColumn(i);
      }
    } else {
      const headers = [['التاريخ', 'اسم المنتج', 'طريقة التحليل', 'قراءة الجهاز', 'v/v%', 'نتيجة المعمل', 'رقم التشغيلة']];
      reportSheet.getRange(4, 1, 1, 7).setValues(headers);
      reportSheet.getRange(4, 1, 1, 7).setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
      
      if (reportData.length > 0) {
        const rows = reportData.map(function(row) {
          return [
            row.date,
            row.product,
            row.method,
            row.deviceReading !== 'N/A' ? row.deviceReading + ' pH' : '-',
            row.vv !== 'N/A' ? row.vv : '-',
            row.labResult !== 'N/A' ? row.labResult : '-',
            row.refNumber
          ];
        });
        reportSheet.getRange(5, 1, rows.length, 7).setValues(rows);
        
        reportSheet.getRange(4, 1, rows.length + 1, 7).setBorder(true, true, true, true, true, true);
      }
      
      for (let i = 1; i <= 7; i++) {
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

function getTodayTanksDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tankSheet = ss.getSheetByName('Tank');
    
    if (!tankSheet) {
      return getEmptyTanksDashboard();
    }
    
    const lastRow = tankSheet.getLastRow();
    if (lastRow < 2) {
      return getEmptyTanksDashboard();
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    const data = tankSheet.getRange(2, 1, lastRow - 1, 14).getValues();
    
    const todayData = [];
    data.forEach(function(row) {
      const timestamp = new Date(row[0]);
      if (timestamp >= today && timestamp < tomorrow) {
        todayData.push({
          timestamp: timestamp,
          serial: row[12],
          vv: row[13]
        });
      }
    });
    
    if (todayData.length === 0) {
      return getEmptyTanksDashboard();
    }
    
    const serials = todayData.map(function(item) { return item.serial; }).filter(function(s) { return s !== '' && !isNaN(s); });
    const vvValues = todayData.map(function(item) { return parseFloat(item.vv) * 100; }).filter(function(v) { return !isNaN(v); });
    
    const startSerial = serials.length > 0 ? Math.min.apply(null, serials) : 0;
    const endSerial = serials.length > 0 ? Math.max.apply(null, serials) : 0;
    
    let highest = 0;
    let lowest = 0;
    let highestSerial = 0;
    let lowestSerial = 0;
    
    if (vvValues.length > 0) {
      highest = Math.max.apply(null, vvValues);
      lowest = Math.min.apply(null, vvValues);
      
      todayData.forEach(function(item) {
        const vvPercent = parseFloat(item.vv) * 100;
        if (vvPercent === highest) {
          highestSerial = item.serial;
        }
        if (vvPercent === lowest) {
          lowestSerial = item.serial;
        }
      });
    }
    
    return {
      totalCount: todayData.length,
      startSerial: startSerial,
      endSerial: endSerial,
      highest: highest.toFixed(2),
      lowest: lowest.toFixed(2),
      highestSerial: highestSerial,
      lowestSerial: lowestSerial
    };
    
  } catch (error) {
    Logger.log('Error in getTodayTanksDashboardData: ' + error.message);
    return getEmptyTanksDashboard();
  }
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
        
        data.forEach(function(row, index) {
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
        
        data.forEach(function(row, index) {
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
      const sortedConc = allConcentrations.sort(function(a, b) { return b.value - a.value; });
      
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

function getEmptyTanksDashboard() {
  return {
    totalCount: 0,
    startSerial: 0,
    endSerial: 0,
    highest: '0.00',
    lowest: '0.00',
    highestSerial: 0,
    lowestSerial: 0
  };
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
