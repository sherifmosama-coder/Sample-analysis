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
    } else if (portal === 'TDS') {
      correctPassword = systemSheet.getRange('C2').getValue().toString();
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
    return products.map(function(row) { return row[0]; }).filter(function(val) { return val !== ''; });
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
          product:    String(data[i][0] || ''),
          standard:   String(data[i][1] || ''),
          appearance: String(data[i][2] || ''),
          odor:       String(data[i][3] || ''),
          density:    String(data[i][4] || ''),
          alcohol:    String(data[i][5] || ''),
          solids:     String(data[i][6] || ''),
          micro:      String(data[i][7] || ''),
          shelfLife:  String(data[i][8] || ''),
          storage:    String(data[i][9] || ''),
          logoUrl:    String(data[i][10] || '')
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

// --- PRODUCTION PORTAL FUNCTIONS ---

function authenticateUser(passcode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { valid: false, message: 'لم يتم العثور على شيت Users. الرجاء إنشائه.' };
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { valid: false, message: 'قائمة المستخدمين فارغة' };
    
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1].toString() === passcode.toString()) {
        return { valid: true, userName: data[i][0] };
      }
    }
    return { valid: false, message: 'رمز المرور غير صحيح' };
  } catch(e) { return { valid: false, message: e.message }; }
}

function getProdSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Production_Logs');
    if (!sheet) {
      sheet = ss.insertSheet('Production_Logs');
      sheet.appendRow(['Session ID', 'Timestamp', 'User Name', 'Date', 'Product Name', 'Start Time', 'End Time', 'Duration (Mins)', 'Quantity', 'Avg Time/Unit (Mins)', 'Notes', 'Image URL', 'Status']);
      sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
    }
    
    const idxSheet = ss.getSheetByName('index');
    let products = [];
    if (idxSheet && idxSheet.getLastRow() >= 2) {
      const data = idxSheet.getRange(2, 2, idxSheet.getLastRow() - 1, 1).getValues();
      products = data.map(function(r) { return r[0]; }).filter(function(v) { return v !== ''; });
    }
    return { products: products };
  } catch (e) { throw new Error(e.message); }
}

function startProdSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const sessionId = 'PRD-' + new Date().getTime();
    
    sheet.appendRow([
      sessionId, new Date(), data.user, data.date, data.product, 
      data.startTime, '', '', '', '', '', '', 'Active'
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
    
    const table = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    let rowIndex = -1;
    for (let i = 0; i < table.length; i++) {
      if (table[i][0] === data.sessionId) { rowIndex = i + 2; break; }
    }
    if (rowIndex === -1) throw new Error('لم يتم العثور على الجلسة');
    
    const startTimeStr = table[rowIndex - 2][5];
    const duration = calculateDuration(startTimeStr, data.endTime);
    const avg = data.quantity > 0 ? (duration / data.quantity).toFixed(2) : 0;
    
    sheet.getRange(rowIndex, 7).setValue(data.endTime);
    sheet.getRange(rowIndex, 8).setValue(duration);
    sheet.getRange(rowIndex, 9).setValue(data.quantity);
    sheet.getRange(rowIndex, 10).setValue(avg);
    sheet.getRange(rowIndex, 11).setValue(data.notes || '');
    sheet.getRange(rowIndex, 12).setValue(data.imageUrl || '');
    sheet.getRange(rowIndex, 13).setValue('Completed');
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    addCellNote(sheet, rowIndex, 13, `[${timestamp}] Session ended by ${data.user}.`);
    
    let dateStr = table[rowIndex - 2][3] instanceof Date ? Utilities.formatDate(table[rowIndex - 2][3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(table[rowIndex - 2][3]);
    syncWasteTimes(dateStr);
    
    return { success: true };
  } catch (e) { throw new Error(e.message); }
}

function getTodayProdSessions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    const notesData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getNotes();
    
    const filtered = [];
    for (let i = 0; i < data.length; i++) {
      // Ensure the cell date is formatted to match todayStr regardless of how Sheets saved it
      let rowDateStr = '';
      if (data[i][3]) {
        try {
          rowDateStr = Utilities.formatDate(new Date(data[i][3]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch(e) {
          rowDateStr = String(data[i][3]).trim();
        }
      }
      
      if (rowDateStr === todayStr) {
        let sTime = data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), 'HH:mm') : String(data[i][5] || '');
        let eTime = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), 'HH:mm') : String(data[i][6] || '');
        
        // Combine notes from Start, End, Qty, Notes, and Image columns (Indices 5, 6, 8, 10, 11)
        let combinedNotes = [notesData[i][5], notesData[i][6], notesData[i][8], notesData[i][10], notesData[i][11]]
          .filter(function(n) { return n !== ''; })
          .join('\n');
        
        filtered.push({
          id: data[i][0], user: data[i][2], date: rowDateStr, product: data[i][4], 
          startTime: sTime, endTime: eTime, duration: data[i][7], 
          quantity: data[i][8], avg: data[i][9], notes: data[i][10], image: data[i][11], status: data[i][12],
          editHistory: combinedNotes
        });
      }
    }
    return filtered.reverse(); // Newest first
  } catch (e) { throw new Error(e.message); }
}

function editProdSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const lastRow = sheet.getLastRow();
    
    const table = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    let rowIndex = -1; let oldData = [];
    for (let i = 0; i < table.length; i++) {
      if (table[i][0] === data.sessionId) { rowIndex = i + 2; oldData = table[i]; break; }
    }
    if (rowIndex === -1) throw new Error('Session not found');
    
    const oldStartTime = oldData[5] instanceof Date ? Utilities.formatDate(oldData[5], Session.getScriptTimeZone(), 'HH:mm') : String(oldData[5] || '');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');

    if (data.isActiveEdit) {
      sheet.getRange(rowIndex, 6).setValue(data.newStartTime);
      if (oldStartTime !== data.newStartTime) {
        addCellNote(sheet, rowIndex, 6, `[${timestamp}] ${data.user}: تعديل وقت البدء من ${oldStartTime} إلى ${data.newStartTime}`);
      }
      return { success: true };
    }
    
    const oldEndTime = oldData[6] instanceof Date ? Utilities.formatDate(oldData[6], Session.getScriptTimeZone(), 'HH:mm') : String(oldData[6] || '');
    const oldQty = oldData[8];
    const oldNotes = oldData[10];
    
    const duration = calculateDuration(data.newStartTime, data.newEndTime);
    const newAvg = data.newQuantity > 0 ? (duration / data.newQuantity).toFixed(2) : 0;
    
    sheet.getRange(rowIndex, 6).setValue(data.newStartTime);
    sheet.getRange(rowIndex, 7).setValue(data.newEndTime);
    sheet.getRange(rowIndex, 8).setValue(duration);
    sheet.getRange(rowIndex, 9).setValue(data.newQuantity);
    sheet.getRange(rowIndex, 10).setValue(newAvg);
    sheet.getRange(rowIndex, 11).setValue(data.newNotes || '');
    if (data.newImageUrl) {
      sheet.getRange(rowIndex, 12).setValue(data.newImageUrl);
    }
    
    if (oldStartTime !== data.newStartTime) {
      addCellNote(sheet, rowIndex, 6, `[${timestamp}] ${data.user}: تعديل وقت البدء من ${oldStartTime} إلى ${data.newStartTime}`);
    }
    if (oldEndTime !== data.newEndTime) {
      addCellNote(sheet, rowIndex, 7, `[${timestamp}] ${data.user}: تعديل وقت الإنهاء من ${oldEndTime} إلى ${data.newEndTime}`);
    }
    if (String(oldQty) !== String(data.newQuantity)) {
      addCellNote(sheet, rowIndex, 9, `[${timestamp}] ${data.user}: تعديل الكمية من ${oldQty} إلى ${data.newQuantity}`);
    }
    if (String(oldNotes) !== String(data.newNotes)) {
      addCellNote(sheet, rowIndex, 11, `[${timestamp}] ${data.user}: تم تعديل الملاحظات`);
    }
    if (data.newImageUrl) {
      addCellNote(sheet, rowIndex, 12, `[${timestamp}] ${data.user}: تم تحديث الصورة`);
    }
    
    let dateStr = oldData[3] instanceof Date ? Utilities.formatDate(oldData[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(oldData[3]);
    syncWasteTimes(dateStr);
    
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
    sheet.appendRow([
      sessionId, new Date(), data.user, data.date, 'وقت الراحة',
      data.startTime, data.endTime || '', duration, '', '', data.notes || '', '', status
    ]);
    syncWasteTimes(data.date);
    return { success: true };
  } catch (e) { throw new Error(e.message);
  }
}

function endBreakSession(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Production_Logs');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('لا توجد بيانات');
    const table = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    let rowIndex = -1;
    for (let i = 0; i < table.length; i++) {
      if (table[i][0] === data.sessionId) { rowIndex = i + 2; break; }
    }
    if (rowIndex === -1) throw new Error('لم يتم العثور على الجلسة');
    
    const startTimeStr = table[rowIndex - 2][5];
    const duration = calculateDuration(startTimeStr, data.endTime);
    
    sheet.getRange(rowIndex, 7).setValue(data.endTime);
    sheet.getRange(rowIndex, 8).setValue(duration);
    sheet.getRange(rowIndex, 13).setValue('Completed');
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    addCellNote(sheet, rowIndex, 13, `[${timestamp}] Break ended by ${data.user}.`);
    
    let dateStr = table[rowIndex - 2][3] instanceof Date ? Utilities.formatDate(table[rowIndex - 2][3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(table[rowIndex - 2][3]);
    syncWasteTimes(dateStr);
    
    return { success: true };
  } catch (e) { throw new Error(e.message); }
}

function timeToMins(timeStr) {
  if (!timeStr) return 0;
  let p = String(timeStr).split(':');
  return parseInt(p[0]||0) * 60 + parseInt(p[1]||0);
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
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();

  let timeline = new Array(1440).fill(false);
  let todayRows = [];

  data.forEach(function(r, i) {
    let rowDate = r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[3]||'').trim();
    if (rowDate === dateStr) {
      let type = (r[4] === 'وقت مهدر') ? 'Waste' : 'Busy';
      
      // Safely extract HH:mm regardless of how Google Sheets formatted the cell
      let sTime = r[5] instanceof Date ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), 'HH:mm') : String(r[5]||'').trim();
      let eTime = r[6] instanceof Date ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'HH:mm') : String(r[6]||'').trim();
      
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

  todayRows.forEach(function(r) {
    if (r.type === 'Waste' || !r.start) return;
    let sMin = timeToMins(r.start);

    if (r.status === 'Active') {
      for(let m = sMin; m < nowMin; m++) timeline[m] = true;
    } else if (r.end) {
      let eMin = timeToMins(r.end);
      if (eMin < sMin) eMin += 1440; 
      for(let m = sMin; m < eMin; m++) timeline[m] = true;
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

  let wasteRows = todayRows.filter(function(r) { return r.type === 'Waste'; }).sort(function(a,b) { return timeToMins(a.start) - timeToMins(b.start); });
  let maxI = Math.max(gaps.length, wasteRows.length);
  let rowsToDelete = [];

  for (let i = 0; i < maxI; i++) {
    if (i < gaps.length && i < wasteRows.length) {
      sheet.getRange(wasteRows[i].rowIndex, 6).setValue(minsToTime(gaps[i].start));
      sheet.getRange(wasteRows[i].rowIndex, 7).setValue(minsToTime(gaps[i].end));
      sheet.getRange(wasteRows[i].rowIndex, 8).setValue(gaps[i].end - gaps[i].start);
    } else if (i < gaps.length) {
      const sessionId = 'WST-' + new Date().getTime() + '-' + i;
      sheet.appendRow([
        sessionId, new Date(), 'النظام', dateStr, 'وقت مهدر',
        minsToTime(gaps[i].start), minsToTime(gaps[i].end), gaps[i].end - gaps[i].start,
        '', '', '', '', 'Completed'
      ]);
    } else {
      rowsToDelete.push(wasteRows[i].rowIndex);
    }
  }

  rowsToDelete.sort(function(a,b) { return b - a; }).forEach(function(r) { sheet.deleteRow(r); });
}

// --- END OF DAY AUTOMATION & DAILY SUMMARY ---

function endOfDayRoutine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Logs');
  if (!sheet) return;
  
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
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
  const updatedData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  let actualStart = 1440; let actualEnd = 0;
  let totalProd = 0; let totalBreak = 0; let totalWaste = 0;
  let hasData = false;
  
  let endDayBreaks = [];
  updatedData.forEach(function(r) {
    let rowDate = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3]).trim();
    if (rowDate === todayStr && r[12] === 'Completed' && r[4] === 'وقت الراحة') {
      let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
      let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
      let dur = eM - sM; if(dur < 0) dur += 1440;
      endDayBreaks.push({start: sM, end: sM + dur});
    }
  });

  updatedData.forEach(function(r) {
    let rowDate = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3]).trim();
    if (rowDate === todayStr && r[12] === 'Completed') {
      hasData = true;
      let product = r[4];
      let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
      let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
      let dur = eM - sM; if(dur < 0) dur += 1440;
      let absEM = sM + dur;
      
      if(sM < actualStart) actualStart = sM;
      if(eM > actualEnd) actualEnd = eM;
      
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

function getProductionReportData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Production_Logs');
  if (!sheet) throw new Error("Sheet not found");

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const tz = Session.getScriptTimeZone();

  let kpis = { prod: 0, break: 0, waste: 0, overtime: 0 };
  let prodMap = {}; 
  let trendMap = {}; 
  let usersSet = new Set();
  let productsSet = new Set();
  let dayGroups = {};

  // 1. Filter & Group Chronologically
  data.forEach(r => {
    if (r[12] !== 'Completed') return; // Only process finished sessions
    
    let dateStr = r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3]).trim();
    let user = r[2];
    let product = r[4];

    usersSet.add(user);
    if (product !== 'وقت مهدر' && product !== 'وقت الراحة') productsSet.add(product);

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
        let dur = eM - sM; if(dur < 0) dur += 1440;
        dayBreaks.push({start: sM, end: sM + dur});
      }
    });

    dayRows.forEach(r => {
      let sM = timeToMins(r[5] instanceof Date ? Utilities.formatDate(r[5], tz, 'HH:mm') : String(r[5]));
      let eM = timeToMins(r[6] instanceof Date ? Utilities.formatDate(r[6], tz, 'HH:mm') : String(r[6]));
      let dur = eM - sM; if(dur < 0) dur += 1440;
      let absEM = sM + dur;
      let product = r[4];

      if(sM < actualStart) actualStart = sM;
      if(eM > actualEnd) actualEnd = eM;

      if (product === 'وقت مهدر') {
         dWaste += dur; kpis.waste += dur;
      } else if (product === 'وقت الراحة') {
         dBreak += dur; kpis.break += dur;
      } else {
         let netDur = dur;
         // Subtract overlapping breaks
         dayBreaks.forEach(b => {
             let overlapStart = Math.max(sM, b.start);
             let overlapEnd = Math.min(absEM, b.end);
             if (overlapStart < overlapEnd) netDur -= (overlapEnd - overlapStart);
         });
         dProd += netDur; kpis.prod += netDur;
         if (!prodMap[product]) prodMap[product] = { qty: 0, dur: 0 };
         prodMap[product].qty += (parseFloat(r[8]) || 0);
         prodMap[product].dur += netDur;
      }
    });

    let shiftStart = (actualStart < 480) ? actualStart : 480;
    let shiftEnd = shiftStart + 510;
    let deficit = (actualStart > 480) ? actualStart - 480 : 0;
    let extraMins = actualEnd > shiftEnd ? actualEnd - shiftEnd : 0;
    let compMins = Math.min(deficit, extraMins);
    let overtimeMins = extraMins - compMins;

    kpis.overtime += overtimeMins;
    trendMap[d] = { date: d, prod: dProd, waste: dWaste, break: dBreak, ot: overtimeMins };
  }

  let trendArray = Object.values(trendMap).sort((a, b) => a.date.localeCompare(b.date));
  let productsArray = Object.keys(prodMap).map(p => {
     let d = prodMap[p];
     let avg = d.qty > 0 ? (d.dur / d.qty).toFixed(2) : 0;
     return { name: p, qty: d.qty, dur: d.dur, avg: avg };
  }).sort((a,b) => b.dur - a.dur);

  return {
     kpis: kpis,
     trend: trendArray,
     products: productsArray,
     users: Array.from(usersSet).sort(),
     productNames: Array.from(productsSet).sort()
  };
}
