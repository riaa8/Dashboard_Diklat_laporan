// Spreadsheet ID
const SPREADSHEET_ID = '10HQEZi1yw7sA5g8z46iDKbkI1ebAsn5Sipl8Ghsc8eU';

// ===== CACHING HELPERS =====
function getCache() {
  return CacheService.getScriptCache();
}

function getCachedData(key) {
  const cache = getCache();
  const cachedStr = cache.get(key);
  if (cachedStr) {
    try {
      return JSON.parse(cachedStr);
    } catch (e) {
      Logger.log('Error parsing cache for key ' + key + ': ' + e.toString());
      return null;
    }
  }
  return null;
}

function setCachedData(key, data, expirationInSeconds) {
  const cache = getCache();
  try {
    const dataStr = JSON.stringify(data);
    // Maksimal value panjang string 100KB, kita perlu handle jika terlalu besar.
    if (dataStr.length < 100000) {
       cache.put(key, dataStr, expirationInSeconds);
    } else {
       Logger.log('Data for cache key ' + key + ' is too large (' + dataStr.length + ')');
    }
  } catch (e) {
    Logger.log('Error setting cache for key ' + key + ': ' + e.toString());
  }
}

function invalidateCache(keys) {
  const cache = getCache();
  if (Array.isArray(keys)) {
    cache.removeAll(keys);
    Logger.log('Invalidated cache keys: ' + keys.join(', '));
  } else {
    cache.remove(keys);
    Logger.log('Invalidated cache key: ' + keys);
  }
}
// ==========================
// Template Document ID for Surat Balasan
const TEMPLATE_SURAT_BALASAN_ID = '1lfEp1UJ0gNgaey-7sMjBYLY1R6oZWTU5dTzjEzszeyc';
// Template Document ID for Surat Penolakan
const TEMPLATE_SURAT_PENOLAKAN_ID = '1V-sMTjnjuHCi0h4Io2v7ScRBmPbo_qY3Bt_WZtL-DJs';
// Template Document ID for Surat Selesai
const TEMPLATE_SURAT_SELESAI_ID = '1XUlwNrNUUgaXfn5z6myFuTos027Tjqrnxv2refPNtVw'; // New ID from user



// ===== REST API ENDPOINTS (FOR REACT SPA) =====

function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    let params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      return responseJson({ success: false, message: 'No payload data provided' });
    }
    
    const action = params.action;
    const args = params.args || [];
    
    // Dynamic function invocation (Bypasses the need to map 50+ functions manually)
    if (typeof this[action] === 'function') {
      const result = this[action].apply(this, args);
      // Return wrapped in standard JSON response
      return responseJson({ success: true, data: result });
    } else {
      return responseJson({ success: false, message: 'Server function not found: ' + action });
    }
    
  } catch (error) {
    return responseJson({ success: false, message: error.toString(), stack: error.stack });
  }
}

// Function to serve HTML (Backward compatibility & GET API)
function doGet(e) {
  // 1. Handle GET requests as API if action parameter exists
  if (e && e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    let args = [];
    if (e.parameter.args) {
      try { args = JSON.parse(e.parameter.args); } catch(err) {} 
    }
    if (typeof this[action] === 'function') {
      try {
        const result = this[action].apply(this, args);
        return responseJson({ success: true, data: result });
      } catch (err) {
        return responseJson({ success: false, message: err.toString() });
      }
    }
    return responseJson({ success: false, message: 'Server function not found' });
  }

  // 2. Legacy Monolithic SSR Fallback (Akan dimatikan total di akhir Fase 7)
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Dashboard Diklat - Management System')
    .setFaviconUrl('https://upload.wikimedia.org/wikipedia/id/d/de/Semen_Tonasa_logo.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Include function for template (Legacy)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get Spreadsheet
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ===== TEST FUNCTIONS (for debugging) =====
function testConnection() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets().map(s => s.getName());
    Logger.log('Connected successfully!');
    Logger.log('Available sheets: ' + sheets.join(', '));
    return { success: true, sheets: sheets };
  } catch (error) {
    Logger.log('Connection error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ===== UTILITY FUNCTION =====
// Function to fix KET column header font size
function fixKETHeaderFontSize() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Magang" tidak ditemukan' };
    }
    
    // Get the KET header cell (assume it's in column B, row 1)
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const ketColumnIndex = headers.indexOf('KET');
    
    if (ketColumnIndex !== -1) {
      const headerCell = sheet.getRange(1, ketColumnIndex + 1);
      
      // Set font size to 10 (same as default)
      headerCell.setFontSize(10);
      
      Logger.log('KET header font size updated to 10');
      return { success: true, message: 'Font size berhasil diubah ke 10' };
    } else {
      return { success: false, message: 'Kolom KET tidak ditemukan' };
    }
  } catch (error) {
    Logger.log('Error fixing font size: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Function to clear all formatting from KET column
function clearKETColumnFormatting() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Magang" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const ketColumnIndex = headers.indexOf('KET');
    
    if (ketColumnIndex === -1) {
      return { success: false, message: 'Kolom KET tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, message: 'Tidak ada data untuk dibersihkan' };
    }
    
    // Get all cells in KET column (excluding header)
    const ketRange = sheet.getRange(2, ketColumnIndex + 1, lastRow - 1, 1);
    
    // Clear all formatting: background, font color, and font weight
    ketRange.setBackground(null);  // Reset to default white
    ketRange.setFontColor(null);   // Reset to default black
    ketRange.setFontWeight('normal');
    
    Logger.log(`Cleared formatting for ${lastRow - 1} cells in KET column`);
    return { success: true, message: `Format berhasil dibersihkan untuk ${lastRow - 1} baris` };
  } catch (error) {
    Logger.log('Error clearing formatting: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Function to apply highlight to all existing rows with REKOMENDASI content
function highlightExistingRekomendasi() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Magang" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rekomendasiIndex = headers.map(h => String(h).toUpperCase()).findIndex(h => h.includes('REKOMENDASI'));
    
    if (rekomendasiIndex === -1) {
      return { success: false, message: 'Kolom REKOMENDASI tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, message: 'Tidak ada data untuk di-highlight' };
    }
    
    // Get all values in REKOMENDASI column (excluding header)
    const rekomendasiColumn = sheet.getRange(2, rekomendasiIndex + 1, lastRow - 1, 1);
    const values = rekomendasiColumn.getValues();
    
    let highlightedCount = 0;
    
    // Loop through each row and apply highlight to entire row if REKOMENDASI has content
    for (let i = 0; i < values.length; i++) {
      const rowRange = sheet.getRange(i + 2, 1, 1, headers.length);
      const value = values[i][0];
      
      if (value && value.toString().trim() !== '') {
        // Has content - apply highlight to entire row
        rowRange.setBackground('#ffff00'); // Bright yellow
        rowRange.setFontColor('#000000'); // Black text
        highlightedCount++;
      } else {
        // Empty - remove highlight from entire row
        rowRange.setBackground(null);
        rowRange.setFontColor(null);
      }
    }
    
    Logger.log(`Applied highlight to ${highlightedCount} rows in Data Magang`);
    return { 
      success: true, 
      message: `Berhasil highlight ${highlightedCount} dari ${lastRow - 1} baris yang memiliki rekomendasi` 
    };
  } catch (error) {
    Logger.log('Error highlighting rekomendasi: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function testGetKaryawanData() {
  const result = getKaryawanData();
  Logger.log('Karyawan Data Result:');
  Logger.log(JSON.stringify(result));
  return result;
}

function testGetMagangData() {
  const result = getMagangData();
  Logger.log('Magang Data Result:');
  Logger.log(JSON.stringify(result));
  return result;
}

// TEST FUNCTION - Debug auto-numbering
function testAutoNumbering() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const lastRow = sheet.getLastRow();
    
    Logger.log('Headers: ' + JSON.stringify(headers));
    
    let noIndex = headers.indexOf('No');
    if (noIndex === -1) noIndex = headers.indexOf('NO');
    if (noIndex === -1) noIndex = headers.indexOf('no');
    
    Logger.log('No column index: ' + noIndex);
    Logger.log('Last row: ' + lastRow);
    
    if (noIndex !== -1 && lastRow > 1) {
      const noColumn = sheet.getRange(2, noIndex + 1, lastRow - 1, 1).getValues();
      Logger.log('No column values: ' + JSON.stringify(noColumn));
      
      const numbers = noColumn.map(row => {
        const val = row[0];
        const num = typeof val === 'number' ? val : parseInt(val);
        return isNaN(num) ? 0 : num;
      }).filter(n => n > 0);
      
      Logger.log('Parsed numbers: ' + JSON.stringify(numbers));
      
      if (numbers.length > 0) {
        const maxNo = Math.max(...numbers);
        Logger.log('Max No found: ' + maxNo);
        Logger.log('Next No should be: ' + (maxNo + 1));
        return { maxNo: maxNo, nextNo: maxNo + 1 };
      }
    }
    
    return { message: 'No index not found or no data' };
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return { error: error.toString() };
  }
}

// UTILITY FUNCTION - Initialize all empty No fields with sequential numbers
// Run this ONCE to fix existing data
function initializeMagangNumbers() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Magang" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, message: 'Tidak ada data untuk di-initialize' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find No column
    let noIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'No' || headers[i] === 'NO' || headers[i].toString().toUpperCase() === 'NO') {
        noIndex = i;
        break;
      }
    }
    
    if (noIndex === -1) {
      return { success: false, message: 'Kolom No tidak ditemukan' };
    }
    
    // Generate sequential numbers 1, 2, 3, ...
    const numbers = [];
    for (let i = 0; i < lastRow - 1; i++) {
      numbers.push([i + 1]);
    }
    
    // Fill kolom No (skip header)
    sheet.getRange(2, noIndex + 1, lastRow - 1, 1).setValues(numbers);
    
    Logger.log('Successfully initialized ' + (lastRow - 1) + ' numbers');
    return { 
      success: true, 
      message: 'Berhasil mengisi nomor untuk ' + (lastRow - 1) + ' baris (1 sampai ' + (lastRow - 1) + ')'
    };
  } catch (error) {
    Logger.log('Error initializing numbers: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// DEBUG FUNCTION - Check sheet info for troubleshooting
function debugMagangSheet() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    
    if (!sheet) {
      Logger.log('ERROR: Sheet "Data Magang" tidak ditemukan');
      return { success: false, message: 'Sheet not found' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    // Check if sheet is empty
    if (lastCol === 0) {
       return { success: true, message: 'Sheet empty' };
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    Logger.log('=== DEBUG MAGANG SHEET ===');
    Logger.log('Sheet Name: Data Magang');
    Logger.log('Last Row: ' + lastRow);
    Logger.log('Last Column: ' + lastCol);
    Logger.log('Headers: ' + JSON.stringify(headers));
    
    return {
      success: true,
      lastRow: lastRow,
      lastCol: lastCol,
      headers: headers
    };
  } catch (error) {
    Logger.log('ERROR in debugMagangSheet: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ===== FIX DATE FORMATS UTILITY =====
function fixMagangDateFormats() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Activity Log', 'Login/register', 'Notifikasi'];
    
    let fixedCount = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (excludedSheets.includes(name)) return;
      
      Logger.log('Scanning sheet: ' + name);
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      // Identify Date Columns
      const dateColIndices = [];
      headers.forEach((h, i) => {
        const header = h ? h.toString().toUpperCase() : '';
        if (header.includes('TGL') || header.includes('DATE') || header.includes('WAKTU')) {
          dateColIndices.push(i);
        }
      });
      
      if (dateColIndices.length === 0) return;
      
      // Process each date column
      dateColIndices.forEach(colIndex => {
        // Get full column data (skipping header)
        // Check if there are rows to get
        if (lastRow < 2) return;
        
        const range = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
        const values = range.getValues();
        let changed = false;
        
        const newValues = values.map(row => {
          let val = row[0];
          
          // Logic to fix strings -> Date Objects (Noon)
          if (typeof val === 'string' && val.trim() !== '') {
             // CASE A: YYYY-MM-DD
             const parts = val.match(/^(\d{4})-(\d{2})-(\d{2})$/);
             if (parts) {
                 val = new Date(parseInt(parts[1]), parseInt(parts[2]) - 1, parseInt(parts[3]), 12, 0, 0);
                 changed = true;
                 fixedCount++;
             } 
             // CASE B: ISO or other string
             else {
                 const d = new Date(val);
                 if (!isNaN(d.getTime())) {
                     d.setHours(12, 0, 0, 0);
                     val = d;
                     changed = true;
                     fixedCount++;
                 }
             }
          } else if (val instanceof Date) {
              const h = val.getHours();
              if (h === 0 || h === 23) {
                 val.setHours(12, 0, 0, 0);
                 changed = true; 
              }
          }
           return [val];
        });
        
        // Write back values if any parsing happened
        if (changed) {
           range.setValues(newValues);
        }
        
        // ALWAYS Apply Number Format 'd mmmm yyyy'
        range.setNumberFormat('d mmmm yyyy'); 
      });
    });
    
    return { success: true, message: 'Berhasil memperbaiki format tanggal. Total fixed columns/cells: ' + fixedCount };
  } catch (e) {
    Logger.log('Error fix date: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

/**
 * Utility function to hide all sheets that have "Lookers" in their name.
 * Run this function once manually or via trigger.
 */
function hideLookersSheets() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    let hiddenCount = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (name.toLowerCase().includes('lookers')) {
        sheet.hideSheet();
        hiddenCount++;
      }
    });
    
    invalidateCache(['magang_batch_names']);
    return { success: true, message: `Berhasil menyembunyikan ${hiddenCount} sheet Lookers.` };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}


// ===== MAINTENANCE MODE FUNCTIONS =====
function setMaintenanceMode(active, message) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('MAINTENANCE_ACTIVE', active ? 'true' : 'false');
    props.setProperty('MAINTENANCE_MESSAGE', message || 'Sistem sedang dalam pemeliharaan rutin. Silakan kembali lagi nanti.');
    return { success: true, message: `Maintenance mode ${active ? 'diaktifkan' : 'dimatikan'}` };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getMaintenanceStatus() {
  try {
    const props = PropertiesService.getScriptProperties();
    return {
      active: props.getProperty('MAINTENANCE_ACTIVE') === 'true',
      message: props.getProperty('MAINTENANCE_MESSAGE') || 'Sistem sedang maintenance.'
    };
  } catch (e) {
    return { active: false, message: '' };
  }
}

// ===== AUDIT LOG FUNCTIONS =====
function getAuditLogs(limit = 100) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Activity Log');
    if (!sheet) return { success: false, message: 'Sheet "Activity Log" tidak ditemukan' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, logs: [] };
    
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const logs = data.map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).reverse();
    return { success: true, logs: logs };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== AUTOCOMPLETE FUNCTIONS =====

// Get unique values from a specific column for autocomplete
function getUniqueColumnValues(columnName, sheetName) {
  try {
    const ss = getSpreadsheet();
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    // Get headers and normalize checks
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Fuzzy search for column index (case-insensitive, trimmed)
    const targetCol = columnName.toUpperCase().trim();
    const colIndex = headers.findIndex(h => h && h.toString().toUpperCase().trim() === targetCol);
    
    if (colIndex === -1) {
      Logger.log('Column "' + columnName + '" not found (Fuzzy search)');
      return { success: true, values: [] }; // Return empty instead of error
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, values: [] };
    }
    
    // Get all values in column
    const columnData = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
    
    // Get unique non-empty values
    const uniqueValues = [...new Set(
      columnData
        .map(row => row[0])
        .filter(val => val && val.toString().trim() !== '')
        .map(val => val.toString().trim())
    )].sort();
    
    Logger.log('Found ' + uniqueValues.length + ' unique values for ' + columnName);
    
    return { success: true, values: uniqueValues };
  } catch (error) {
    Logger.log('Error in getUniqueColumnValues: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Get autocomplete data for all relevant fields
function getMagangAutocompleteData() {
  try {
    const CACHE_KEY = 'magang_autocomplete_data';
    const cachedResult = getCachedData(CACHE_KEY);
    if (cachedResult) return cachedResult;

    // Hanya load autocomplete untuk UNIVERSITAS dan JURUSAN
    const fieldsToAutocomplete = [
      'UNIVERSITAS',
      'JURUSAN',
      // Added based on user request for better autocomplete coverage
      'PENGIRIM',
      'JABATAN',
      'UNIT KERJA',
      'PEMBIMBING',
      'NOMOR SURAT' 
    ];
    
    const autocompleteData = {};
    
    for (const field of fieldsToAutocomplete) {
      const result = getUniqueColumnValues(field);
      if (result.success) {
        autocompleteData[field] = result.values;
      } else {
        autocompleteData[field] = [];
      }
    }
    
    Logger.log('Autocomplete data loaded for ' + Object.keys(autocompleteData).length + ' fields');
    
    const finalResult = { success: true, data: autocompleteData };
    setCachedData(CACHE_KEY, finalResult, 3600); // Cache for 1 hour

    return finalResult;
  } catch (error) {
    Logger.log('Error in getMagangAutocompleteData: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


// ===== DATA KARYAWAN FUNCTIONS =====

// Get all employee data
function getKaryawanData() {
  try {
    const CACHE_KEY = 'karyawan_data';
    const cachedResult = getCachedData(CACHE_KEY);
    if (cachedResult) return cachedResult;

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Karyawan');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Karyawan" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) {
      return { success: true, data: [], headers: [] };
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Filter out completely empty rows
    const validRows = rows.filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined));
    
    const formattedData = validRows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        let value = row[index];
        
        // Handle Date objects - Format Indonesia
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd MMMM yyyy');
        }
        
        // Convert to string and handle empty values
        obj[header] = value !== null && value !== undefined && value !== '' ? String(value) : '';
      });
      return obj;
    });
    
    const finalResult = { success: true, data: formattedData, headers: headers };
    setCachedData(CACHE_KEY, finalResult, 3600); // Cache for 1 hour
    
    return finalResult;
  } catch (error) {
    Logger.log('Error in getKaryawanData: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Add new employee
function addKaryawan(data, userLog) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Karyawan');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Karyawan" tidak ditemukan' };
    }
    
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // [NEW] Check Created By Column
    let createdByIdx = headers.indexOf('Created By');
    if (createdByIdx === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Created By');
      headers.push('Created By'); // Update local headers array
    }

    // [NEW] Inject User Log
    if (userLog) data['Created By'] = userLog;
    
    const newRow = headers.map(header => data[header] || '');
    
    sheet.appendRow(newRow);
    
    invalidateCache(['karyawan_data', 'dashboard_stats', 'dashboard_widgets_data']);
    
    return { success: true, message: 'Data karyawan berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Update employee data
function updateKaryawan(rowIndex, data) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Karyawan');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Karyawan" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const updatedRow = headers.map(header => data[header] || '');
    
    // rowIndex + 2 karena: +1 untuk header, +1 karena index mulai dari 0
    sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([updatedRow]);
    
    invalidateCache(['karyawan_data']);
    
    return { success: true, message: 'Data karyawan berhasil diupdate' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Delete employee
function deleteKaryawan(rowIndex) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Karyawan');
    
    if (!sheet) {
      return { success: false, message: 'Sheet "Data Karyawan" tidak ditemukan' };
    }
    
    // rowIndex + 2 karena: +1 untuk header, +1 karena index mulai dari 0
    sheet.deleteRow(rowIndex + 2);
    
    invalidateCache(['karyawan_data', 'dashboard_stats', 'dashboard_widgets_data']);

    return { success: true, message: 'Data karyawan berhasil dihapus' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Get next available No for magang
function getNextMagangNo(sheetName) {
  try {
    const ss = getSpreadsheet();
    // Default to 'Data Magang' if not provided
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    
    // If no data yet, start from 1
    if (lastRow <= 1) {
      return { success: true, nextNo: 1 };
    }
    
    // Next number = current row count (for preview only)
    // Actual number will be auto-generated by addMagang()
    const nextNo = lastRow; // Row 2 = no 1, Row 3 = no 2, etc.
    
    Logger.log('Next preview No for ' + targetSheet + ': ' + nextNo);
    
    return { success: true, nextNo: nextNo };
  } catch (error) {
    Logger.log('Error in getNextMagangNo: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ===== DATA MAGANG FUNCTIONS =====

// [NEW] Get list of Magang Batch Names ONLY (FAST)
function getMagangBatchNames() {
  try {
    const CACHE_KEY = 'magang_batch_names';
    const cachedResult = getCachedData(CACHE_KEY);
    if (cachedResult) return cachedResult;

    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    
    // System sheets to exclude
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Data Magang', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Activity Log', 'Login/register', 'Notifikasi'];
    
    const batchNames = [];
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // Exclude system sheets and Lookers sheets
      const isSystemSheet = excludedSheets.includes(name);
      const isLookersSheet = name.toLowerCase().includes('lookers');
      const isHidden = sheet.isSheetHidden();
      
      if (!isSystemSheet && !isLookersSheet && !isHidden) {
        batchNames.push(name);
      }
    });

    // Simple reverse sort (assuming newer sheets are added last or we want them first?)
    // Or just reverse them to show latest created (usually) first
    // batchNames.reverse(); 
    
    const finalResult = { success: true, batchNames: batchNames };
    setCachedData(CACHE_KEY, finalResult, 3600); // Cache for 1 hour

    return finalResult;
  } catch (error) {
     return { success: false, message: error.toString() };
  }
}

/**
 * [NEW] Mengambil daftar tahun unik dari semua batch magang
 * Digunakan untuk mengisi sidebar laporan secara dinamis
 */
function getMagangYears() {
  try {
    const res = getMagangBatchNames();
    if (!res.success) return res;
    
    const years = new Set();
    const currentYear = new Date().getFullYear();
    
    // Tambahkan tahun saat ini sebagai default
    years.add(currentYear);
    
    res.batchNames.forEach(name => {
      // Asumsi format nama sheet: "Bulan Tahun" (Contoh: "Januari 2026")
      const parts = name.split(' ');
      const year = parseInt(parts[parts.length - 1]);
      if (!isNaN(year) && year > 2000 && year < 2100) {
        years.add(year);
      }
    });
    
    // Urutkan tahun dari yang terbaru
    const sortedYears = Array.from(years).sort((a, b) => b - a);
    return { success: true, years: sortedYears };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// Wrapper function for backward compatibility
function getMagangBatches() {
  const result = getMagangBatchNames();
  if (result.success) {
    return { success: true, batches: result.batchNames };
  }
  return result;
}

// [NEW] Get Metadata for a SINGLE Batch (Used for Lazy Loading)
function getBatchMetadata(batchName) {
  try {
     const CACHE_KEY = 'batch_metadata_v4_' + batchName; // Force refresh again
     const cachedResult = getCachedData(CACHE_KEY);
     if (cachedResult) return cachedResult;

     const ss = getSpreadsheet();
     const sheet = ss.getSheetByName(batchName);
     if (!sheet) return { success: false, message: 'Sheet not found' };

      // Calculate Metadata
      const lastRow = sheet.getLastRow();
      const count = Math.max(0, lastRow - 1); // Subtract header
      
      let activeCount = 0;
      let finishedCount = 0;
      let earliestDate = null;
      let latestDate = null;
      let status = 'Finished'; // Default
      
      if (count > 0) {
        // 1. Read ONLY headers
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        
        const activeHeaders = headers.map(h => h ? h.toString().trim().toUpperCase() : '');
        
        // [Refined] Find TGL MASUK
        // Priority 1: Exact Standard Names
        let tglMasukIdx = activeHeaders.findIndex(h => ['TGL MASUK', 'TANGGAL MASUK', 'START DATE', 'TGL. MASUK', 'TANGGAL MULAI'].includes(h));
        
        if (tglMasukIdx === -1) {
             // Priority 2: Contains MASUK/MULAI but NOT confusing terms
             tglMasukIdx = activeHeaders.findIndex(h => 
                (h.includes('MASUK') || h.includes('MULAI')) && 
                !h.includes('SURAT') && 
                !h.includes('BERKAS') && 
                !h.includes('PROPOSAL') &&
                !h.includes('DOKUMEN')
             );
        }
        
        // [Refined] Find TGL SELESAI
        let tglSelesaiIdx = activeHeaders.findIndex(h => ['TGL SELESAI', 'TANGGAL SELESAI', 'END DATE', 'TGL. SELESAI', 'TANGGAL BERAKHIR'].includes(h));
        
        if (tglSelesaiIdx === -1) {
             tglSelesaiIdx = activeHeaders.findIndex(h => 
                (h.includes('SELESAI') || h.includes('BERAKHIR')) && 
                !h.includes('SURAT')
             );
        }
        
        // [Refined] Find Validation/Status Column
        // Priority 1: Exact matches for known status columns
        let validationIdx = activeHeaders.findIndex(h => ['KET', 'KETERANGAN STATUS', 'STATUS', 'VALIDASI', 'KET.'].includes(h));
        
        // Priority 2: Fuzzy match but exclude "KETERANGAN" if it stands alone (likely notes)
        if (validationIdx === -1) {
             validationIdx = activeHeaders.findIndex(h => 
                (h.includes('STATUS') || h.includes('VALIDASI') || h.includes('KET')) && 
                h !== 'KETERANGAN' && // Avoid generic notes
                !h.includes('SURAT')
             );
        }
        
        // Priority 3: Fallback to KETERANGAN if nothing else found (risky but legacy support)
        if (validationIdx === -1) {
             validationIdx = activeHeaders.findIndex(h => h === 'KETERANGAN');
        }
        
        // 2. Read ONLY necessary data columns
        const today = new Date();
        today.setHours(0,0,0,0);
        
        let tglSelesaiValues = [];
        let tglMasukValues = [];
        let validationValues = [];
        
        if (tglSelesaiIdx !== -1) {
             tglSelesaiValues = sheet.getRange(2, tglSelesaiIdx + 1, count, 1).getValues();
        }
        
        if (tglMasukIdx !== -1) {
             tglMasukValues = sheet.getRange(2, tglMasukIdx + 1, count, 1).getValues();
        }
 
        if (validationIdx !== -1) {
             validationValues = sheet.getRange(2, validationIdx + 1, count, 1).getValues();
        }
        
        for (let i = 0; i < count; i++) {
            // [UPDATED] Check Validation Status (Strict "Diteruskan" only)
            let isValid = true; // Default to true if no column found (backward compatibility)
            
            if (validationIdx !== -1) {
                 const statusVal = validationValues[i][0] ? String(validationValues[i][0]).toLowerCase().trim() : '';
                 // User Requirement: "Ambil dari validation yang sudah di teruskan"
                 // inclusive check for "diteruskan" AND other valid statuses (prevent excluding finished interns)
                 if (!statusVal.includes('diteruskan') && 
                     !statusVal.includes('selesai') && 
                     !statusVal.includes('lulus') && 
                     !statusVal.includes('aktif') &&
                     !statusVal.includes('tamat')) {
                      isValid = false; 
                 }
            }
 
            if (!isValid) continue; // Skip if not "Diteruskan"


            // Check TGL SELESAI
            // [UPDATED] Check Status FIRST. If "Selesai" or "Lulus", count as finished regardless of date.
            const statusVal = validationIdx !== -1 && validationValues[i][0] ? validationValues[i][0].toString().toLowerCase() : '';
            
            if (statusVal.includes('selesai') || statusVal.includes('lulus') || statusVal.includes('tamat')) {
                 finishedCount++;
            } else if (tglSelesaiIdx !== -1) {
                const cellVal = tglSelesaiValues[i][0];
                if (cellVal) {
                    const tglSelesai = new Date(cellVal);
                    if (!isNaN(tglSelesai.getTime())) { 
                         if (tglSelesai >= today) {
                            activeCount++;
                         } else {
                            finishedCount++;
                         }
                         if (!latestDate || tglSelesai > latestDate) latestDate = tglSelesai;
                    } else {
                       activeCount++; // Invalid date but not rejected? Assume active for safety or custom handling
                    }
                } else {
                     activeCount++; // No end date = active
                }
            } else {
               // If no date column, assume everyone is active (fallback)
               activeCount++;
            }
            
            // Check TGL MASUK
            if (tglMasukIdx !== -1) {
                const cellVal = tglMasukValues[i][0];
                if (cellVal) {
                    const tglMasuk = new Date(cellVal);
                    if (!isNaN(tglMasuk.getTime())) {
                        if (!earliestDate || tglMasuk < earliestDate) earliestDate = tglMasuk;
                    }
                }
            }
        }
      }
      
      if (activeCount > 0) status = 'Active';
      const formatDate = (d) => {
        if (!d) return '-';
        // [FIX] Add buffer to prevent -1 day timezone issue
        // Clone date to not affect original object if used elsewhere (though here it's fine)
        const safeDate = new Date(d);
        safeDate.setHours(12, 0, 0, 0); // Set to noon to be safe
        return Utilities.formatDate(safeDate, Session.getScriptTimeZone(), 'd MMM yyyy');
      };

      // Return the metadata object
      const finalResult = {
        success: true,
        data: {
            name: batchName,
            count: count,
            activeCount: activeCount,
            finishedCount: finishedCount,
            status: status,
            dateRange: (earliestDate ? formatDate(earliestDate) : '?') + ' - ' + (latestDate ? formatDate(latestDate) : '?'),
            rawDate: latestDate ? latestDate.getTime() : 0
        }
      };
      
      setCachedData(CACHE_KEY, finalResult, 300); // Reduce cache to 5 minutes for better consistency
      
      return finalResult;

  } catch (e) {
      return { success: false, message: e.toString() };
  }
}


// [NEW] Create a new Magang Batch (Sheet) from Template
function createNewMagangBatch(batchName) {
  try {
    const ss = getSpreadsheet();
    
    // Check if sheet exists
    if (ss.getSheetByName(batchName)) {
      return { success: false, message: 'Batch "' + batchName + '" sudah ada.' };
    }
    
    // Get Template (Data Magang is now the Master Template)
    const templateSheet = ss.getSheetByName('Data Magang');
    if (!templateSheet) {
      return { success: false, message: 'Template "Data Magang" tidak ditemukan.' };
    }
    
    // Duplicate
    const newSheet = templateSheet.copyTo(ss);
    newSheet.setName(batchName);
    
    // Clear Content (Keep Headers & Validation)
    const lastRow = newSheet.getLastRow();
    
    if (lastRow > 1) {
      // 1. Clear content of Row 2 (this preserves Data Validation/Dropdowns)
      const lastCol = newSheet.getLastColumn();
      if (lastCol > 0) {
        newSheet.getRange(2, 1, 1, lastCol).clearContent();
      }
      
      // 2. If there are extra rows (Row 3 onwards), delete them
      if (lastRow > 2) {
        newSheet.deleteRows(3, lastRow - 2);
      }
      
      Logger.log('Batch created with clean row 2 (validation preserved)');
    } else {
      // Edge case: Template has only header. Add 1 blank row.
      newSheet.insertRowAfter(1);
    }
    
    // Make it active
    newSheet.activate();
    
    // Invalidate caches
    invalidateCache(['magang_batch_names', 'magang_batch_names_v3', 'dashboard_stats', 'dashboard_widgets_data']);

    return { success: true, message: 'Batch "' + batchName + '" berhasil dibuat.' };
  } catch (error) {
    Logger.log('Error create batch: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Get all internship data (supports dynamic sheet)
function getMagangData(sheetName) {
  try {
    const ss = getSpreadsheet();
    // Default to 'Data Magang' if no sheet provided (legacy safety)
    const targetSheet = sheetName || 'Data Magang'; 
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) {
      return { success: true, data: [], headers: [] };
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Filter out completely empty rows
    const validRows = rows.filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined));
    
    // OPTIMIZATION: Return raw array of arrays with explicit string conversion
    const cleanData = validRows.map(row => row.map(cell => {
      if (cell instanceof Date) return cell.toISOString();
      return (cell !== null && cell !== undefined) ? String(cell) : '';
    }));
    
    return { success: true, data: cleanData, headers: headers };
  } catch (error) {
    Logger.log('Error in getMagangData: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Add new internship
function addMagang(data, sheetName, userLog) {
  try {
    const ss = getSpreadsheet();
    // Default to 'Data Magang' for backward compatibility
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    // Check if data is valid
    if (!data || typeof data !== 'object') {
      Logger.log('ERROR: Data is undefined or invalid');
      return { success: false, message: 'Data tidak valid' };
    }
    
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // [NEW] Check Created By Column
    let createdByIdx = headers.indexOf('Created By');
    if (createdByIdx === -1) {
      sheet.getRange(1, headers.length + 1).setValue('Created By');
      headers.push('Created By');
    }
    
    // [NEW] Inject User Log
    if (userLog) data['Created By'] = userLog;
    const lastRow = sheet.getLastRow();
    
    Logger.log('=== ADD MAGANG START ===');
    Logger.log('Target Sheet: ' + targetSheet);
    Logger.log('Data received: ' + JSON.stringify(data));
    
    const ketIndex = headers.indexOf('KET');
    const kehadiranIndex = headers.indexOf('KEHADIRAN');
    const rekomendasiIndex = headers.indexOf('REKOMENDASI');
    
    // Insert a new blank row after the last row
    sheet.insertRowAfter(lastRow);
    const newRowIndex = lastRow + 1;
    
    // Copy formatting & validation from MASTER Template ('Data Magang') to ensure consistency
    // This allows "healing" of broken batches that lost their validation.
    const masterSheet = ss.getSheetByName('Data Magang');
    if (masterSheet) {
      const sourceRange = masterSheet.getRange(2, 1, 1, headers.length);
      const targetRange = sheet.getRange(newRowIndex, 1, 1, headers.length);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    } else if (lastRow > 1) {
       // Fallback to previous row if Master not found (unlikely)
       const sourceRange = sheet.getRange(lastRow, 1, 1, headers.length);
       const targetRange = sheet.getRange(newRowIndex, 1, 1, headers.length);
       sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
       sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    
    // AUTO-GENERATE nomor berdasarkan row position
    // If empty sheet (lastRow=1 header only), autoNumber = 1
    const autoNumber = lastRow; 
    Logger.log('AUTO-GENERATED No: ' + autoNumber);
    
    // Build new row data
    const newRowData = headers.map((header, index) => {
      // Auto-generate No column
      if (header === 'No' || header === 'NO' || header.toUpperCase() === 'NO') {
        Logger.log('Setting auto-generated No at index ' + index + ': ' + autoNumber);
        return autoNumber;
      }
      
      // Leave empty for dropdown columns
      if (index === ketIndex || index === kehadiranIndex) {
        return '';
      }
      
      // Get data from user input with SMART HEADER MAPPING
      let val = data[header];
      
      // Fallback: Check specific known problematic fields
      const h = header.toUpperCase();
      
      // 1. STAMBUK / NIM variations
      if (!val && (h.includes('STAMBUK') || h.includes('NIM'))) {
        val = data['STAMBUK / NIM'] || data['STAMBUK'] || data['NIM'] || data['STAMBUK/NIM'];
        if (val) Logger.log('Smart mapped STAMBUK for header "' + header + '" -> ' + val);
      }
      
      // 2. NAMA MAHASISWA variations
      if (!val && h.includes('NAMA') && h.includes('MAHASISWA')) {
        val = data['NAMA MAHASISWA'] || data['NAMA'] || data['MAHASISWA'];
      }

      // 3. REKOMENDASI variations
      if (!val && h.includes('REKOMENDASI')) {
        val = data['REKOMENDASI'] || data['REKOMENDASI / Catatan'];
      }

      // 4. DATE Formatting (Fix Format: d MMMM yyyy)
      // Check if this looks like a date column
      if (h.includes('TGL') || h.includes('DATE') || h.includes('WAKTU')) {
         // Current value might be String (YYYY-MM-DD) or Date
         if (val) {
             // If string YYYY-MM-DD, convert to Date
             if (typeof val === 'string') {
                 // CASE A: YYYY-MM-DD (Standard Input)
                 const parts = val.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                 if (parts) {
                     // Create date at noon to avoid timezone shift
                     val = new Date(parseInt(parts[1]), parseInt(parts[2]) - 1, parseInt(parts[3]), 12, 0, 0);
                 } 
                 // CASE B: ISO String (e.g. 2025-09-28T16:00:00.000Z) or other formats
                 else {
                     const d = new Date(val);
                     if (!isNaN(d.getTime())) {
                         // Normalize to Noon to stay safe from TZ shifts
                         d.setHours(12, 0, 0, 0);
                         val = d;
                     }
                 }
             }
         }
      }
      
      return val || '';
    });
    
    Logger.log('New row data: ' + JSON.stringify(newRowData));
    
    // Set all values at once
    sheet.getRange(newRowIndex, 1, 1, headers.length).setValues([newRowData]);
    
    // APPLY DATE FORMATTING (d mmmm yyyy)
    headers.forEach((header, index) => {
        const h = header.toUpperCase();
        if (h.includes('TGL') || h.includes('DATE')) {
             // 1-based index is index + 1
             sheet.getRange(newRowIndex, index + 1).setNumberFormat('d mmmm yyyy');
        }
    });

    // Apply highlight if REKOMENDASI has content
    const upperHeaders = headers.map(h => String(h).toUpperCase());
    const rekomendasiIndexSmart = upperHeaders.findIndex(h => h.includes('REKOMENDASI'));
    
    if (rekomendasiIndexSmart !== -1) {
      const rowRange = sheet.getRange(newRowIndex, 1, 1, headers.length);
      const val = newRowData[rekomendasiIndexSmart];
      
      // Strict check for content
      if (val && String(val).trim().length > 0) {
        rowRange.setBackground('#ffff00');
        rowRange.setFontColor('#000000');
        Logger.log('Applied yellow highlight for Recommendation');
      } else {
        // IMPORTANT: Clear formatting if empty (prevent inheriting from previous row)
        // But preserve border/alignment if needed? Usually null clears everything. 
        // Better to just clear Back/Font color
        rowRange.setBackground(null);
        rowRange.setFontColor(null);
        Logger.log('Cleared highlight (empty Recommendation)');
      }
    }
    
    // [NEW] Log Activity
    const studentName = data['NAMA MAHASISWA'] || data['NAMA'] || 'Mahasiswa Baru';
    logActivity('Menambahkan Data Magang', studentName, userLog, `Unit: ${data['UNIT KERJA'] || '-'}`);
    
    invalidateCache(['batch_metadata_' + targetSheet, 'magang_autocomplete_data', 'dashboard_stats', 'dashboard_widgets_data', 'analytics_data_all']);

    Logger.log('=== ADD MAGANG SUCCESS ===');
    return { success: true, message: '🎉 DATA BARU BERHASIL DITAMBAHKAN (No: ' + autoNumber + ')' };
  } catch (error) {
    Logger.log('=== ADD MAGANG ERROR ===');
    Logger.log('Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}


// Add batch magang (multiple mahasiswa dengan 1 surat yang sama)
function addBatchMagang(submissionData) {
  try {
    const ss = getSpreadsheet();
    // Extract sheetName from submissionData
    const sheetName = submissionData.sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + sheetName + '" tidak ditemukan' };
    }
    
    const commonData = submissionData.commonData;
    const mahasiswaList = submissionData.mahasiswaList;
    
    if (!mahasiswaList || mahasiswaList.length === 0) {
      return { success: false, message: 'Tidak ada data mahasiswa yang dikirim' };
    }
    
    Logger.log('=== ADD BATCH MAGANG START ===');
    Logger.log('Available in sheet: ' + sheetName);
    Logger.log('Common data: ' + JSON.stringify(commonData));
    Logger.log('Mahasiswa count: ' + mahasiswaList.length);
    
    const successCount = [];
    const failedCount = [];
    
    // Loop setiap mahasiswa dan create entry terpisah
    for (var i = 0; i < mahasiswaList.length; i++) {
      try {
        var mahasiswa = mahasiswaList[i];
        
        // Gabungkan commonData dengan data mahasiswa individual
        var fullData = {};
        
        // Copy all common data
        for (var key in commonData) {
          if (commonData.hasOwnProperty(key)) {
            fullData[key] = commonData[key];
          }
        }
        
        // Add mahasiswa-specific data
        fullData['NAMA MAHASISWA'] = mahasiswa.nama;
        fullData['STAMBUK'] = mahasiswa.stambuk;
        
        Logger.log('Processing mahasiswa ' + (i + 1) + ': ' + mahasiswa.nama);
        
        // Call existing addMagang function with explicit sheetName
        var result = addMagang(fullData, sheetName);
        
        if (result.success) {
          successCount.push(mahasiswa.nama);
          Logger.log('SUCCESS: ' + mahasiswa.nama);
        } else {
          failedCount.push(mahasiswa.nama);
          Logger.log('FAILED: ' + mahasiswa.nama + ' - ' + result.message);
        }
      } catch (error) {
        Logger.log('Error adding mahasiswa ' + mahasiswa.nama + ': ' + error.toString());
        failedCount.push(mahasiswa.nama);
      }
    }
    
    Logger.log('=== ADD BATCH MAGANG COMPLETE ===');
    Logger.log('Success: ' + successCount.length + ', Failed: ' + failedCount.length);
    
    if (failedCount.length === 0) {
      return { 
        success: true, 
        message: successCount.length + ' mahasiswa berhasil ditambahkan ke ' + sheetName
      };
    } else {
      return {
        success: false,
        message: successCount.length + ' berhasil, ' + failedCount.length + ' gagal. Gagal: ' + failedCount.join(', ')
      };
    }
    
  } catch (error) {
    Logger.log('=== ADD BATCH MAGANG ERROR ===');
    Logger.log('Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Update internship data
// Update internship data
function updateMagang(rowIndex, data, sheetName, userLog) {
  try {
    // FORCE Integer conversation to prevent string concatenation bugs
    const rIndex = parseInt(rowIndex);
    
    Logger.log('=== UPDATE MAGANG START ===');
    const targetSheet = sheetName || 'Data Magang';
    Logger.log('Target Sheet: ' + targetSheet);
    Logger.log('Received rowIndex: ' + rowIndex + ' (parsed: ' + rIndex + ')');
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Calculate 1-based row number (Header is row 1, index 0 is row 2)
    const targetRow = rIndex + 2;
    Logger.log('Target Row in Sheet: ' + targetRow);
    
    if (targetRow > sheet.getLastRow()) {
      Logger.log('WARNING: Target row ' + targetRow + ' is beyond last row ' + sheet.getLastRow());
    }

    // [NEW] READ OLD DATA FOR DIFFING
    const oldRowValues = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
    const oldData = {};
    headers.forEach((h, i) => oldData[h] = oldRowValues[i]);

    const updatedRow = headers.map((header, index) => {
      // Handle both "No" and "NO" variants for the number column
      if (header === 'NO' || header === 'No') {
        return data['No'] || data['NO'] || '';
      }
      
      // Get data with SMART MAPPING
      let val = data[header];
      const h = header.toUpperCase();
      
      if (!val && (h.includes('STAMBUK') || h.includes('NIM'))) {
        val = data['STAMBUK / NIM'] || data['STAMBUK'] || data['NIM'] || data['STAMBUK/NIM'];
      }
      
      if (!val && h.includes('NAMA') && h.includes('MAHASISWA')) {
        val = data['NAMA MAHASISWA'] || data['NAMA'] || data['MAHASISWA'];
      }

      if (!val && h.includes('REKOMENDASI')) {
        val = data['REKOMENDASI'] || data['REKOMENDASI / Catatan'];
      }
      
      // Special handling for Status/Validation columns that might be missing in form
      // If header is KET or KEHADIRAN and data[header] is undefined (not just empty string), keep old value
      if ((h === 'KET' || h === 'KEHADIRAN') && data[header] === undefined) {
          return oldRowValues[index];
      }

      // 3. DATE Formatting (Fix Format: d MMMM yyyy)
      // Check if this looks like a date column
      if (h.includes('TGL') || h.includes('DATE') || h.includes('WAKTU')) {
         // Current value might be String (YYYY-MM-DD) or Date
         if (val) {
             // If string YYYY-MM-DD, convert to Date
             if (typeof val === 'string') {
                 // CASE A: YYYY-MM-DD (Standard Input)
                 const parts = val.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                 if (parts) {
                     // Create date at noon to avoid timezone shift
                     val = new Date(parseInt(parts[1]), parseInt(parts[2]) - 1, parseInt(parts[3]), 12, 0, 0);
                 } 
                 // CASE B: ISO String (e.g. 2025-09-28T16:00:00.000Z) or other formats
                 else {
                     const d = new Date(val);
                     if (!isNaN(d.getTime())) {
                         // CHECK FOR TIMEZONE ARTIFACTS
                         // If time is late (e.g. > 15:00 UTC), it often means it belongs to the next day in Indonesia (UTC+7/8)
                         // But we can't be sure of the original intention if we don't know the offset.
                         // However, if we assume "Date Only" was intended:
                         // Simple check: if hours are not 0 or 12, reset to 12.
                         
                         // Fix for "2 becoming 1":
                         // If the date was parsing to "Previous Day 16:00" because of UTC... 
                         // We should probably rely on the Sheet's display value, but we only have raw value here.
                         // Let's force NOON Local.
                         d.setHours(12, 0, 0, 0);
                         val = d;
                     }
                 }
             }
         }
      }
      
      return val || '';
    });
    
    // === FIX VALIDATION: Ensure the new Status is allowed in Dropdown ===
    const ketIndex = headers.indexOf('KET');
    if (ketIndex !== -1 && data['KET']) {
         const cell = sheet.getRange(targetRow, ketIndex + 1);
         ensureValidationOption(cell, data['KET']);
    }

    // Update the row data only, no formatting
    const rowRange = sheet.getRange(targetRow, 1, 1, headers.length);
    rowRange.setValues([updatedRow]);

    // APPLY DATE FORMATTING (d mmmm yyyy)
    headers.forEach((header, index) => {
        const h = header.toUpperCase();
        if (h.includes('TGL') || h.includes('DATE')) {
             // 1-based index is index + 1
             sheet.getRange(targetRow, index + 1).setNumberFormat('d mmmm yyyy');
        }
    });
    
    // FORCE LEFT ALIGNMENT for the updated row
    rowRange.setHorizontalAlignment("left");
    
    // Apply/remove highlight to ENTIRE ROW based on REKOMENDASI content
    const upperHeadersMap = headers.map(h => String(h).toUpperCase());
    const rekomendasiIndexSmart = upperHeadersMap.findIndex(h => h.includes('REKOMENDASI'));
    
    if (rekomendasiIndexSmart !== -1) {
      const rowRangeUpdate = sheet.getRange(targetRow, 1, 1, headers.length);
      const val = updatedRow[rekomendasiIndexSmart];
      
      if (val && String(val).trim() !== '') {
        // Has content - apply highlight to entire row
        rowRangeUpdate.setBackground('#ffff00'); // Bright yellow
        rowRangeUpdate.setFontColor('#000000'); // Black text
      } else {
        // Empty - remove highlight from entire row
        rowRangeUpdate.setBackground(null);
        rowRangeUpdate.setFontColor(null);
      }
    }
    
    // [NEW] DETECT CHANGES (DIFF LOGGING)
    let changes = [];
    headers.forEach((header, index) => {
        // Skip technical columns
        if (header === 'No' || header === 'NO' || header === 'Created By') return;

        const oldVal = oldData[header] ? String(oldData[header]) : '';
        const newVal = updatedRow[index] ? String(updatedRow[index]) : '';
        
        // Simple comparison (ignoring date format differences for now, assuming string uniformity if possible)
        // For Dates: updatedRow might be Date object or String. oldVal is Date object often.
        // Let's rely on loose string comparison for now.
        if (oldVal.trim() !== newVal.trim()) {
             // Basic formatting for Dates
             let displayOld = oldVal.length > 50 ? oldVal.substring(0,47)+'...' : oldVal;
             let displayNew = newVal.length > 50 ? newVal.substring(0,47)+'...' : newVal;
             changes.push(`${header}: "${displayOld}" -> "${displayNew}"`);
        }
    });

    let details = `Row: ${targetRow}`;
    if (changes.length > 0) {
        // Limit details length
        details = changes.join(', ');
        if (details.length > 300) details = details.substring(0, 300) + '...';
    } else {
        details = 'Tidak ada perubahan data (Save ulang)';
    }

    // [NEW] Log Activity with DETAILS + METADATA for Interactivity
    // Metadata format: { sheet: '...', rowIndex: ... } encoded in Subject or a separate field?
    // Current logActivity takes (Action, Subject, Creator, Details).
    // We can embed metadata in Subject via a delimiter if we want to hack it without DB schema change.
    // Better: Stick to standard fields. Subject = Name. Details = Diff.
    // For interactivity, frontend needs to know Sheet & Row. 
    // Let's sneak it into Details inside hidden markers or just append it.
    // Or better: Let's assume Subject is just the Name, and we search for Name/ID when clicking? 
    // No, duplicate names exist. 
    // Let's append `[REF:${targetSheet}:${rIndex}]` to Details. Frontend can parse it.
    
    const metaTag = `[REF:${targetSheet}:${rIndex}]`;
    logActivity('Mengedit Data Magang', data['NAMA MAHASISWA'] || data['NAMA'] || 'Mahasiswa', userLog || 'User', `${details} ${metaTag}`);

    invalidateCache(['batch_metadata_' + targetSheet, 'dashboard_stats', 'dashboard_widgets_data', 'analytics_data_all']);

    return { success: true, message: '✅ DATA LAMA BERHASIL DIUPDATE' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Delete internship
function deleteMagang(rowIndex, sheetName, userLog) {
  try {
    const ss = getSpreadsheet();
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameIdx = headers.findIndex(h => h.toUpperCase().includes('NAMA'));
    const subjectName = (nameIdx !== -1) ? sheet.getRange(rowIndex + 2, nameIdx + 1).getValue() : 'Row ' + (rowIndex + 1);

    sheet.deleteRow(rowIndex + 2);
    
    // [NEW] Log Activity
    logActivity('Menghapus Data Magang', subjectName, userLog, `Sheet: ${targetSheet}`);
    
    // RENUMBER all rows after delete to maintain sequential numbering
    renumberMagangRows(targetSheet);
    
    invalidateCache(['batch_metadata_' + targetSheet, 'dashboard_stats', 'dashboard_widgets_data', 'analytics_data_all']);

    return { success: true, message: 'Data magang berhasil dihapus dan nomor telah di-update' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Renumber all rows in Data Magang sheet
function renumberMagangRows(sheetName) {
  try {
    const ss = getSpreadsheet();
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      Logger.log('Sheet "' + targetSheet + '" tidak ditemukan');
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    
    // If no data, nothing to renumber
    if (lastRow <= 1) {
      Logger.log('No data to renumber');
      return { success: true, message: 'Tidak ada data untuk di-renumber' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find No column index
    let noIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === 'No' || headers[i] === 'NO' || headers[i].toString().toUpperCase() === 'NO') {
        noIndex = i;
        break;
      }
    }
    
    if (noIndex === -1) {
      return { success: false, message: 'Kolom No tidak ditemukan' };
    }
    
    // Generate sequential numbers 1, 2, 3, ...
    const numbers = [];
    for (let i = 0; i < lastRow - 1; i++) {
      numbers.push([i + 1]);
    }
    
    // Fill kolom No (skip header)
    sheet.getRange(2, noIndex + 1, lastRow - 1, 1).setValues(numbers);
    
    Logger.log('Successfully initialized ' + (lastRow - 1) + ' numbers');
    return { 
      success: true, 
      message: 'Berhasil mengisi nomor untuk ' + (lastRow - 1) + ' baris (1 sampai ' + (lastRow - 1) + ')'
    };
  } catch (error) {
    Logger.log('Error initializing numbers: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// DEBUG FUNCTION - Check sheet info for troubleshooting
function debugMagangSheet() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    
    if (!sheet) {
      Logger.log('ERROR: Sheet "Data Magang" tidak ditemukan');
      return { success: false, message: 'Sheet not found' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    // Check if sheet is empty
    if (lastCol === 0) {
       return { success: true, message: 'Sheet empty' };
    }
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    Logger.log('=== DEBUG MAGANG SHEET ===');
    Logger.log('Sheet Name: Data Magang');
    Logger.log('Last Row: ' + lastRow);
    Logger.log('Last Column: ' + lastCol);
    Logger.log('Headers: ' + JSON.stringify(headers));
    
    return {
      success: true,
      lastRow: lastRow,
      lastCol: lastCol,
      headers: headers
    };
  } catch (error) {
    Logger.log('ERROR in debugMagangSheet: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ===== FIX DATE FORMATS UTILITY =====
function fixMagangDateFormats() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Activity Log', 'Login/register', 'Notifikasi'];
    
    let fixedCount = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (excludedSheets.includes(name)) return;
      
      Logger.log('Scanning sheet: ' + name);
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      // Identify Date Columns
      const dateColIndices = [];
      headers.forEach((h, i) => {
        const header = h ? h.toString().toUpperCase() : '';
        if (header.includes('TGL') || header.includes('DATE') || header.includes('WAKTU')) {
          dateColIndices.push(i);
        }
      });
      
      if (dateColIndices.length === 0) return;
      
      // Process each date column
      dateColIndices.forEach(colIndex => {
        // Need to read RAW values (strings if possible)
        const range = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
        const values = range.getValues();
        let changed = false;
        
        const newValues = values.map(row => {
          let val = row[0];
          
          if (typeof val === 'string' && val.trim() !== '') {
             // CASE A: YYYY-MM-DD
             const parts = val.match(/^(\d{4})-(\d{2})-(\d{2})$/);
             if (parts) {
                 val = new Date(parseInt(parts[1]), parseInt(parts[2]) - 1, parseInt(parts[3]), 12, 0, 0);
                 changed = true;
                 fixedCount++;
             } 
             // CASE B: ISO string
             else {
                 const d = new Date(val);
                 if (!isNaN(d.getTime())) {
                     const utch = d.getUTCHours();
                     if (utch >= 14) { d.setDate(d.getDate() + 1); }
                     d.setHours(12, 0, 0, 0);
                     val = d;
                     changed = true;
                     fixedCount++;
                 }
             }
          } 
          else if (val instanceof Date) {
              // Existing Date Object.
              // Logic Check: If the hour is late (e.g. 15:00 - 23:00), it's likely a timezone shift from the "Next Day 00:00"
              // Example: 23:00 WIB (previous day) should be 00:00 (next day)
              
              const h = val.getHours();
              if (h >= 15) {
                  // Assume it fell back to previous day
                  val.setDate(val.getDate() + 1);
                  changed = true;
              }
              
              // Normalize to noon
              if (h !== 12) {
                 val.setHours(12, 0, 0, 0);
                 changed = true;
              }
          }
           return [val];
        });
        
        if (changed) {
           range.setValues(newValues);
        }
        
        // Apply Number Format
        range.setNumberFormat('d mmmm yyyy'); 
      });
    });
    
    return { success: true, message: 'Berhasil memperbaiki format tanggal. Total fixed: ' + fixedCount };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== DASHBOARD STATISTICS =====

// ===== DASHBOARD STATISTICS =====

// ===== DASHBOARD WIDGETS DATA =====
function getDashboardWidgetsData() {
  try {
    const CACHE_KEY = 'dashboard_widgets_data';
    const cachedResult = getCachedData(CACHE_KEY);
    if (cachedResult) return cachedResult;

    Logger.log('Starting getDashboardWidgetsData'); // LOG
    const ss = getSpreadsheet();
    const allSheets = ss.getSheets();
    
    // 1. Calculate Stats & Composition
    let activeCount = 0;
    let finishedCount = 0;
    let expiringInterns = [];
    
    const today = new Date();
    // Normalize today to midnight for comparison
    today.setHours(0,0,0,0);
    
    const sevenDaysLater = new Date(today);
    sevenDaysLater.setDate(today.getDate() + 7);
    
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Activity Log', 'Login/register', 'Notifikasi'];
    
    allSheets.forEach(sheet => {
      const name = sheet.getName();
      // Skip excluded sheets AND check if it looks like a Magang Batch (usually Month Year)
      // or simply include everything not in excluded list
      if (name === 'Data Magang' || !excludedSheets.includes(name)) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
          const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim().toUpperCase());
          
          // Identify columns
          const nameIdx = 1; // Assuming Name is usually col 2 (Index 1)
          const ketIdx = headers.indexOf('KET');
          const tglSelesaiIdx = headers.findIndex(h => h.includes('TGL SELESAI') || h.includes('SELESAI'));
          const unitIdx = headers.findIndex(h => h.includes('UNIT KERJA') || h.includes('UNIT'));
          
          values.forEach(row => {
            const ketValue = (ketIdx !== -1 && row[ketIdx]) ? row[ketIdx].toString().trim() : '';
            
            // STRICT ACTIVE COUNTING: Only "Sudah Diteruskan"
            if (ketValue === 'Sudah Diteruskan') {
               activeCount++;
               
               // Check Expiration ONLY for Active Interns
               if (tglSelesaiIdx !== -1 && row[tglSelesaiIdx]) {
                 let endDate = null;
                 if (row[tglSelesaiIdx] instanceof Date) {
                   endDate = row[tglSelesaiIdx];
                 } else if (typeof row[tglSelesaiIdx] === 'string' && row[tglSelesaiIdx].match(/^\d{4}-\d{2}-\d{2}$/)) {
                   endDate = new Date(row[tglSelesaiIdx]);
                 }
                 
                 if (endDate) {
                   endDate.setHours(0,0,0,0);
                   if (endDate >= today && endDate <= sevenDaysLater) {
                      const diffTime = Math.abs(endDate - today);
                      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
                      
                      expiringInterns.push({
                         name: row[nameIdx] || 'Tanpa Nama',
                         unit: (unitIdx !== -1 && row[unitIdx]) ? row[unitIdx] : '-',
                         endDate: Utilities.formatDate(endDate, ss.getSpreadsheetTimeZone(), 'dd MMM yyyy'),
                         daysLeft: diffDays
                      });
                   }
                 }
               }
               
            } else if (['selesai', 'finished', 'resign'].includes(ketValue.toLowerCase())) {
               finishedCount++;
            }
          });
        }
      }
    });

    // Sort expiring by days left (ascending)
    expiringInterns.sort((a, b) => a.daysLeft - b.daysLeft);
    
    const result = {
      success: true,
      data: {
        composition: { 
          active: activeCount, 
          finished: finishedCount 
        },
        expiring: expiringInterns.slice(0, 5) // Limit to top 5
      }
    };
    Logger.log('Widgets Data Result: ' + JSON.stringify(result)); // LOG
    
    setCachedData(CACHE_KEY, result, 900); // cache 15 min

    return result;
    
  } catch (error) {
    Logger.log('Error widgets: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

function getDashboardStats() {
  try {
    const CACHE_KEY = 'dashboard_stats';
    const cachedResult = getCachedData(CACHE_KEY);
    // Kita harus selalu load Activity Log yang fresh, jadi ini triknya:
    // cache hanya return sebagian kalau ada. Karena returnnya kompleks, kita build ualng if ncessary.
    // Tapi untuk simplicity, cache semua dan log direfresh jika ada invalidasi.
    if (cachedResult) return cachedResult;

    const ss = getSpreadsheet();
    
    // 1. Get ALL sheets at once (Optimized)
    const allSheets = ss.getSheets();
    
    // 2. Count Karyawan
    const karyawanSheet = allSheets.find(s => s.getName() === 'Data Karyawan');
    const karyawanCount = karyawanSheet ? Math.max(0, karyawanSheet.getLastRow() - 1) : 0;
    
    // 3. Count Magang (Aggregate)
    let magangCount = 0;
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Activity Log', 'Login/register', 'Notifikasi'];
    
    allSheets.forEach(sheet => {
      const name = sheet.getName();
      // Include 'Data Magang' and all Batch Sheets
      if (name === 'Data Magang' || !excludedSheets.includes(name)) {
         magangCount += Math.max(0, sheet.getLastRow() - 1);
      }
    });

    // 4. Get recent activity (Optimized call)
    // Pass sheets to avoid reopening? No, logic inside is simple enough or we accept 2nd open.
    // For now we keep it simple.
    const activity = getRecentActivity();
    
    // 5. Get Widgets Data (Merged call for efficiency if needed, but separate is fine too)
    // Actually user asked for separate or part of overview.
    // We can call getDashboardWidgetsData HERE and return it all!
    
    // [REMOVED] Widgets removed per user request
    // const widgetData = getDashboardWidgetsData();
    const widgetData = { success: true, data: null }; // Return empty success
    
    const finalResult = {
      success: true,
      stats: {
        karyawan: karyawanCount,
        magang: magangCount,
        total: karyawanCount + magangCount
      },
      activity: activity,
      widgets: widgetData.success ? widgetData.data : null,
      onlineUsers: getOnlineUsers().users // [NEW] Return online users
    };

    setCachedData(CACHE_KEY, finalResult, 300); // 5 menit

    return finalResult;

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// [NEW] Mock Online Users (In real app, we'd use ScriptProperties or CacheService)
function getOnlineUsers() {
    // Return random 1-5 users for demo
    // Or just return the 'Admin'
    return { 
        success: true, 
        users: [
            { name: 'Admin', role: 'Administrator', initial: 'A' },
            // { name: 'User 2', role: 'Staff', initial: 'B' } 
        ]
    };
}


// Update KEHADIRAN status for magang
function updateMagangKehadiran(rowIndex, newKehadiran, sheetName) {
  try {
    const ss = getSpreadsheet();
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const kehadiranIndex = headers.indexOf('KEHADIRAN');
    
    if (kehadiranIndex === -1) {
      return { success: false, message: 'Kolom KEHADIRAN tidak ditemukan' };
    }
    
    // Update KEHADIRAN cell (rowIndex + 2 karena: +1 header, +1 zero-index)
    const targetRow = rowIndex + 2;
    sheet.getRange(targetRow, kehadiranIndex + 1).setValue(newKehadiran);
    
    invalidateCache(['batch_metadata_' + targetSheet]);

    return { success: true, message: 'Status kehadiran berhasil diupdate' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Update KET (Keterangan) status for magang
function updateMagangKet(rowIndex, newKet, sheetName) {
  try {
    const ss = getSpreadsheet();
    const targetSheet = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheet);
    
    if (!sheet) {
      return { success: false, message: 'Sheet "' + targetSheet + '" tidak ditemukan' };
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const ketIndex = headers.indexOf('KET');
    
    if (ketIndex === -1) {
      return { success: false, message: 'Kolom KET tidak ditemukan' };
    }
    
    const targetRow = rowIndex + 2;
    sheet.getRange(targetRow, ketIndex + 1).setValue(newKet);
    
    invalidateCache(['batch_metadata_' + targetSheet]);

    return { success: true, message: 'Keterangan berhasil diupdate' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}


// Get recent activity (Real Log)
function getRecentActivity() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Activity Log');
    
    if (!logSheet) return [];
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) return [];
    
    // Get last 20 rows to ensure we have enough after filtering
    const startRow = Math.max(2, lastRow - 19);
    const numRows = lastRow - startRow + 1;
    const data = logSheet.getRange(startRow, 1, numRows, 5).getValues(); // [Timestamp, Action, Subject, Creator, Details]
    
    const activities = [];
    
    // Process reverse (Newest first)
    for (let i = data.length - 1; i >= 0; i--) {
        const row = data[i];
        const creator = row[3] || 'System';
        
        // [NEW] Skip generic system usernames
        if (creator === 'System' || creator === 'Admin' || creator === 'User') {
            continue; // Skip this entry
        }
        
        let timeStr = 'Baru saja';
        const timestamp = row[0];
        if (timestamp instanceof Date) {
            // Simple relative time or formatting
            const diff = (new Date() - timestamp) / 1000;
            if (diff < 60) timeStr = 'Baru saja';
            else if (diff < 3600) timeStr = Math.floor(diff / 60) + ' menit lalu';
            else if (diff < 86400) timeStr = Math.floor(diff / 3600) + ' jam lalu';
            else timeStr = Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), 'dd MMM');
        }
        
        activities.push({
            type: 'magang', // Icon type
            name: creator,
            action: `${row[1]} <strong>${row[2]}</strong>`, // Action + Subject
            time: timeStr,
            details: row[4] || '' // [NEW] Return Details
        });
    }
    
    return activities.slice(0, 5);
  } catch (error) {
    Logger.log('Error getRecentActivity: ' + error.toString());
    return [];
  }
}

// [NEW] Get All Activity Logs for Dedicated Page (Moved to Top Level)
function getAllActivityLogs() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Activity Log');
    
    // Auto-create if missing
    if (!sheet) {
      sheet = ss.insertSheet('Activity Log');
      sheet.appendRow(['Timestamp', 'Action', 'Subject', 'User', 'Details']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      return { success: true, data: [] };
    }
    
    // Read all data
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
       return { success: true, data: [] };
    }
    
    // Get values (descending order - newest first)
    // We'll reverse in JS for simplicity, or loop backwards
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    // Map to objects
    const logs = values.map(row => {
       const d = new Date(row[0]);
       const timeStr = Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'dd MMM yyyy HH:mm');
       
       return {
           timestampRaw: d.getTime(), // For sorting/filtering
           time: timeStr,
           action: row[1],
           subject: row[2],
           user: row[3],
           details: row[4] || ''
       };
    });
    
    // Sort Newest First
    logs.sort((a, b) => b.timestampRaw - a.timestampRaw);
    
    return { success: true, data: logs };
    
  } catch (error) {
    return { success: false, message: error.toString() }; 
  }
}

// [NEW] Delete Activity Logs (Timezone Safe)
function deleteActivityLogs(rangeType) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Activity Log');
    
    if (!sheet) {
      return { success: false, message: 'Sheet Activity Log tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, message: 'Tidak ada log untuk dihapus' };
    }
    
    if (rangeType === 'all') {
      // Clear all except header
      sheet.deleteRows(2, lastRow - 1);
      return { success: true, message: 'Semua log berhasil dihapus' };
    }
    
    // Timezone Safe Filter
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const keptValues = [];
    
    const timeZone = ss.getSpreadsheetTimeZone();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
    
    const yesterday = new Date(now);
    yesterday.setDate(now.getDate() - 1);
    const yesterdayStr = Utilities.formatDate(yesterday, timeZone, 'yyyy-MM-dd');
    
    let deletedCount = 0;
    
    for (let i = 0; i < values.length; i++) {
        const rowDateVal = values[i][0];
        let rowDateStr = '';
        
        if (rowDateVal instanceof Date) {
            rowDateStr = Utilities.formatDate(rowDateVal, timeZone, 'yyyy-MM-dd');
        } else {
             // Fallback for string dates?
             // If manual parsing needed:
             try {
                const d = new Date(rowDateVal);
                rowDateStr = Utilities.formatDate(d, timeZone, 'yyyy-MM-dd');
             } catch(e) {
                rowDateStr = 'INVALID';
             }
        }
        
        let shouldDelete = false;
        
        if (rangeType === 'today') {
            if (rowDateStr === todayStr) shouldDelete = true;
        } else if (rangeType === 'yesterday') {
             if (rowDateStr === yesterdayStr) shouldDelete = true;
        }
        
        if (!shouldDelete) {
            keptValues.push(values[i]);
        } else {
            deletedCount++;
        }
    }
    
    if (deletedCount === 0) {
        return { success: true, message: 'Tidak ada log yang sesuai kriteria untuk dihapus' };
    }
    
    // Clear old data
    sheet.deleteRows(2, lastRow - 1);
    
    // Write back kept data (if any)
    if (keptValues.length > 0) {
        sheet.getRange(2, 1, keptValues.length, 5).setValues(keptValues);
    }
    
    return { success: true, message: `Berhasil menghapus ${deletedCount} baris log` };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Log Activity Helper
function logActivity(action, subject, creator, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Activity Log');
    
    // Create and Hide if not exists
    if (!sheet) {
      sheet = ss.insertSheet('Activity Log');
      sheet.appendRow(['Timestamp', 'Action', 'Subject', 'Creator', 'Details']);
      sheet.hideSheet();
    }
    
    // Force Hide (ensure it stays hidden)
    if (!sheet.isSheetHidden()) sheet.hideSheet();
    
    const timestamp = new Date();
    sheet.appendRow([
        timestamp,
        action,
        subject,
        creator || 'System',
        details || ''
    ]);
    
    invalidateCache(['dashboard_stats', 'dashboard_activity_log']);

  } catch (e) {
    Logger.log('Error logging activity: ' + e.toString());
  }
}

// Search function
function searchData(sheetName, query) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan` };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const filteredRows = rows.filter(row => {
      return row.some(cell => {
        return cell.toString().toLowerCase().includes(query.toLowerCase());
      });
    });
    
    const formattedData = filteredRows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    return { success: true, data: formattedData };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Export to CSV
function exportToCSV(sheetName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: `Sheet "${sheetName}" tidak ditemukan` };
    }
    
    const data = sheet.getDataRange().getValues();
    const csv = data.map(row => row.join(',')).join('\n');
    
    return { success: true, csv: csv };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ===== MULTIPLE MAHASISWA FEATURE =====
// Add Multiple Magang (Multiple Students from 1 Letter)
function addMultipleMagang(payload) {
  try {
    // Verify Target Sheet Logic
    let targetSheetName = payload.targetSheet;
    
    // Safety check: if targetSheet is missing or empty, warn but default to 'Data Magang' only as last resort
    if (!targetSheetName || targetSheetName.trim() === '') {
        Logger.log('WARNING: targetSheet is empty in payload. Defaulting to "Data Magang"');
        targetSheetName = 'Data Magang';
    }

    Logger.log('Targeting Sheet: ' + targetSheetName);

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) {
      Logger.log('ERROR: Sheet "' + targetSheetName + '" not found.');
      return { success: false, message: 'Sheet "' + targetSheetName + '" tidak ditemukan' };
    }

    // Validate payload
    if (!payload || !payload.sharedData || !payload.students || payload.students.length === 0) {
      Logger.log('ERROR: Invalid payload structure');
      return { success: false, message: 'Data tidak valid' };
    }

    const sharedData = payload.sharedData;
    const students = payload.students;

    Logger.log('=== ADD MULTIPLE MAGANG START ===');
    Logger.log('Payload received: ' + JSON.stringify(payload));
    Logger.log('Shared Data: ' + JSON.stringify(sharedData));
    Logger.log('Number of students: ' + students.length);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('Sheet headers: ' + JSON.stringify(headers));
    Logger.log('Number of columns: ' + headers.length);
    
    const lastRow = sheet.getLastRow();
    const ketIndex = headers.indexOf('KET');
    const kehadiranIndex = headers.indexOf('KEHADIRAN');
    
    // Gunakan pencarian cerdas untuk menghindari bug jika nama header huruf kecil / memiliki tambahan
    const upperHeadersMap = headers.map(h => String(h).toUpperCase());
    const rekomendasiIndexSmart = upperHeadersMap.findIndex(h => h.includes('REKOMENDASI'));

    Logger.log('KET index: ' + ketIndex);
    Logger.log('KEHADIRAN index: ' + kehadiranIndex);
    Logger.log('REKOMENDASI index: ' + rekomendasiIndexSmart);

    let addedCount = 0;
    let errors = [];

    // Loop through each student and create a new row
    for (let i = 0; i < students.length; i++) {
      try {
        const student = students[i];
        Logger.log('Processing student ' + (i + 1) + ': ' + student['NAMA MAHASISWA']);

        // Insert a new blank row after the last row
        const currentLastRow = sheet.getLastRow();
        sheet.insertRowAfter(currentLastRow);
        const newRowIndex = currentLastRow + 1;

        // Copy formatting & validation from MASTER Template ('Data Magang')
        const masterSheet = ss.getSheetByName('Data Magang');
        if (masterSheet) {
          const sourceRange = masterSheet.getRange(2, 1, 1, headers.length);
          const targetRange = sheet.getRange(newRowIndex, 1, 1, headers.length);
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        } else if (currentLastRow > 1) {
          const sourceRange = sheet.getRange(currentLastRow, 1, 1, headers.length);
          const targetRange = sheet.getRange(newRowIndex, 1, 1, headers.length);
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
          sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
        }

        // AUTO-GENERATE nomor berdasarkan row position
        const autoNumber = currentLastRow; // Row 2 = no 1, Row 3 = no 2, etc.
        Logger.log('AUTO-GENERATED No for student ' + (i + 1) + ': ' + autoNumber);

        // Merge shared data + individual student data
        const combinedData = Object.assign({}, sharedData, student);
        
        Logger.log('Student ' + (i + 1) + ' combinedData: ' + JSON.stringify(combinedData));

        // Build new row data
        const newRowData = headers.map((header, index) => {
          Logger.log('  Mapping header "' + header + '" (index ' + index + ')');
          
          // Auto-generate No column
          if (header === 'No' || header === 'NO' || header.toUpperCase() === 'NO') {
            Logger.log('    â†’ Auto-generated No: ' + autoNumber);
            return autoNumber;
          }

          // Leave empty for dropdown columns
          if (index === ketIndex || index === kehadiranIndex) {
            Logger.log('    â†’ Dropdown column, leaving empty');
            return '';
          }

          // Get data from combined data
          let value = combinedData[header] || '';

          // [FIX] SMART MAPPING for STAMBUK & NAME (Missing in Bulk Upload)
          const h = header.toUpperCase();
          
          if (!value && (h.includes('STAMBUK') || h.includes('NIM'))) {
              value = combinedData['STAMBUK / NIM'] || combinedData['STAMBUK'] || combinedData['NIM'] || combinedData['STAMBUK/NIM'];
              Logger.log('    â†’ Smart mapped STAMBUK for header "' + header + '" -> ' + value);
          }
          
          if (!value && h.includes('NAMA') && h.includes('MAHASISWA')) {
              value = combinedData['NAMA MAHASISWA'] || combinedData['NAMA'] || combinedData['MAHASISWA'];
              Logger.log('    â†’ Smart mapped NAMA for header "' + header + '" -> ' + value);
          }
          
          if (!value && h.includes('REKOMENDASI')) {
            value = combinedData['REKOMENDASI'] || combinedData['REKOMENDASI / Catatan'];
            Logger.log('    â†’ Smart mapped REKOMENDASI for header "' + header + '" -> ' + value);
          }
          
          // DATE Formatting (Fix Format: d MMMM yyyy)
          // Check if this looks like a date column
          // const h = header.toUpperCase(); // REMOVED (Declared above)
          if (h.includes('TGL') || h.includes('DATE') || h.includes('WAKTU')) {
             if (value) {
                 // If string YYYY-MM-DD, convert to Date
                 if (typeof value === 'string') {
                     const parts = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                     if (parts) {
                         // Create date at noon to avoid timezone shift
                         value = new Date(parseInt(parts[1]), parseInt(parts[2]) - 1, parseInt(parts[3]), 12, 0, 0);
                     } else {
                         // Try standard parsing
                         const d = new Date(value);
                         if (!isNaN(d.getTime())) value = d;
                     }
                 }
             }
          }

          Logger.log('    â†’ Value: "' + value + '"');
          return value;
        });

        Logger.log('Row data for student ' + (i + 1) + ': ' + JSON.stringify(newRowData));

        // Set all values at once
        sheet.getRange(newRowIndex, 1, 1, headers.length).setValues([newRowData]);

        // APPLY DATE FORMATTING (d mmmm yyyy)
        headers.forEach((header, index) => {
            const h = header.toUpperCase();
            if (h.includes('TGL') || h.includes('DATE')) {
                 sheet.getRange(newRowIndex, index + 1).setNumberFormat('d mmmm yyyy');
            }
        });

        if (rekomendasiIndexSmart !== -1 && newRowData[rekomendasiIndexSmart] && String(newRowData[rekomendasiIndexSmart]).trim() !== '') {
          const rowRange = sheet.getRange(newRowIndex, 1, 1, headers.length);
          rowRange.setBackground('#ffff00');
          rowRange.setFontColor('#000000');
          Logger.log('Applied yellow highlight for student ' + (i + 1));
        } else {
          Logger.log('NO highlight - REKOMENDASI is empty');
        }

        addedCount++;
        
        // [NEW] Log Activity
        const userLog = payload.userLog || 'System';
        logActivity('Menambahkan Data Magang', student['NAMA MAHASISWA'], userLog, `Sheet: ${targetSheetName}`);
        
        Logger.log('Successfully added student ' + (i + 1) + ': ' + student['NAMA MAHASISWA']);

      } catch (studentError) {
        Logger.log('Error adding student ' + (i + 1) + ': ' + studentError.toString());
        errors.push('Student ' + (i + 1) + ': ' + studentError.toString());
      }
    }

    Logger.log('=== ADD MULTIPLE MAGANG COMPLETE ===');
    Logger.log('Successfully added: ' + addedCount + ' students');

    if (addedCount === 0) {
      return { 
        success: false, 
        message: 'Gagal menambahkan data mahasiswa. Errors: ' + errors.join('; ') 
      };
    }

    if (errors.length > 0) {
      invalidateCache(['batch_metadata_' + targetSheetName, 'magang_autocomplete_data', 'dashboard_stats', 'dashboard_widgets_data', 'analytics_data_all']);
      return { 
        success: true, 
        message: 'Berhasil menambahkan ' + addedCount + ' dari ' + students.length + ' mahasiswa. Beberapa error: ' + errors.join('; ')
      };
    }

    invalidateCache(['batch_metadata_' + targetSheetName, 'magang_autocomplete_data', 'dashboard_stats', 'dashboard_widgets_data', 'analytics_data_all']);

    return { 
      success: true, 
      message: 'Berhasil menambahkan ' + addedCount + ' mahasiswa dari surat "' + sharedData['NOMOR SURAT'] + '"'
    };

  } catch (error) {
    Logger.log('=== ADD MULTIPLE MAGANG ERROR ===');
    Logger.log('Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ===== DEBUG FUNCTION =====
function debugTestMapping() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Data Magang');
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('===== DEBUG TEST MAPPING =====');
  Logger.log('Total headers: ' + headers.length);
  Logger.log('All headers:');
  headers.forEach(function(header, index) {
    Logger.log('  [' + index + '] Column ' + String.fromCharCode(65 + index) + ': "' + header + '"');
  });
  
  const namaIdx = headers.indexOf('NAMA MAHASISWA');
  const stambukIdx = headers.indexOf('STAMBUK / NIM');
  const nomorIdx = headers.indexOf('NOMOR SURAT');
  
  Logger.log('');
  Logger.log('Key indexes:');
  Logger.log('  NAMA MAHASISWA: index ' + namaIdx + ' (column ' + String.fromCharCode(65 + namaIdx) + ')');
  Logger.log('  STAMBUK / NIM: index ' + stambukIdx + ' (column ' + String.fromCharCode(65 + stambukIdx) + ')');
  Logger.log('  NOMOR SURAT: index ' + nomorIdx + ' (column ' + String.fromCharCode(65 + nomorIdx) + ')');
  
  return { success: true, namaIdx: namaIdx, stambukIdx: stambukIdx };
}

// ===== ANALYTICS FUNCTIONS =====
// ===== ANALYTICS FUNCTIONS =====
// ===== ANALYTICS FUNCTIONS =====
function getAnalyticsData(yearFilter) {
  try {
    const CACHE_KEY = 'analytics_data_' + (yearFilter || 'all');
    const cachedResult = getCachedData(CACHE_KEY);
    if (cachedResult) return cachedResult;

    const ss = getSpreadsheet();
    
    // 1. Get All Magang Data across accessible sheets
    let sheetsToScan = [];
    const batchResult = getMagangBatchNames(); 
    
    if (batchResult.success) {
      sheetsToScan = batchResult.batchNames;
    }
    // [UPDATED] Do NOT include 'Data Magang' per user request (Analytics only for Batches)
    // if (!sheetsToScan.includes('Data Magang')) {
    //   sheetsToScan.push('Data Magang');
    // }
    
    // Aggregation Variables
    const trendMap = {}; 
    const univMap = {}; 
    const majorMap = {}; 
    const unitMap = {}; 
    
    let totalAllTime = 0;
    let activeParticipants = 0;
    
    const now = new Date();
    
    sheetsToScan.forEach(sheetName => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return;
        
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return;
        
        const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => h.toString().toUpperCase().trim());
        
        // Identify Column Indices (Robust Matching)
        const idxUniv = headers.findIndex(h => h === 'UNIVERSITAS' || h === 'ASAL SEKOLAH' || h.includes('UNIVERSITAS'));
        const idxJurusan = headers.findIndex(h => h === 'JURUSAN' || h === 'PRODI' || h.includes('JURUSAN'));
        const idxTglMasuk = headers.findIndex(h => h === 'TGL MASUK' || h === 'START DATE');
        const idxTglSelesai = headers.findIndex(h => h === 'TGL SELESAI' || h === 'END DATE');
        
        // Use inclusive matching for Unit Kerja
        const idxUnit = headers.findIndex(h => h === 'UNIT KERJA' || h === 'PENEMPATAN' || h === 'DEPARTEMEN' || h === 'UNIT');
        const idxKet = headers.indexOf('KET'); // [ADDED] For Strict Validation
        
        data.forEach(row => {
             // Basic Validation
             if (!row[0]) return; // Skip empty rows if No is empty
             
             let includeRow = true;
             let entryYear = null;

             // Date Processing
             if (idxTglMasuk !== -1 && row[idxTglMasuk] instanceof Date) {
                 entryYear = row[idxTglMasuk].getFullYear();
                 if (yearFilter && yearFilter !== 'all') {
                     if (entryYear.toString() !== yearFilter.toString()) includeRow = false;
                 }
             } else if (yearFilter && yearFilter !== 'all') {
                 includeRow = false; 
             }

             if (!includeRow) return;

             totalAllTime++;

             // Active Check: STRICT (Hanya "Sudah Diteruskan")
             if (idxKet !== -1 && row[idxKet] && row[idxKet].toString().trim() === 'Sudah Diteruskan') {
                 activeParticipants++;
             }

             // 1. University Stats
             if (idxUniv !== -1 && row[idxUniv]) {
                 const univ = row[idxUniv].toString().trim();
                 if (univ) univMap[univ] = (univMap[univ] || 0) + 1;
             }
             
             // 2. Major Stats
             if (idxJurusan !== -1 && row[idxJurusan]) {
                 const major = row[idxJurusan].toString().trim();
                 if (major) majorMap[major] = (majorMap[major] || 0) + 1;
             }
             
             // 3. Unit Kerja Stats
             if (idxUnit !== -1 && row[idxUnit]) {
                 const unit = row[idxUnit].toString().trim();
                 // Filter out extremely short strings (bad data)
                 if (unit && unit.length > 2) unitMap[unit] = (unitMap[unit] || 0) + 1;
             }

             // 4. Trend (Monthly)
             if (idxTglMasuk !== -1 && row[idxTglMasuk] instanceof Date) {
                const date = row[idxTglMasuk];
                const year = date.getFullYear();
                const month = date.getMonth(); 
                const key = year + '-' + String(month + 1).padStart(2, '0'); 
                trendMap[key] = (trendMap[key] || 0) + 1;
             }
        });
    });
    
    // Sort & Format Data
    const trendLabels = Object.keys(trendMap).sort();
    const trendData = trendLabels.map(k => trendMap[k]);
    
    const getTopN = (map, n) => {
        return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, n);
    };

    const topUniv = getTopN(univMap, 5);
    const topMajor = getTopN(majorMap, 8);
    
    // Log unit map for debugging
    Logger.log('Unit Map Keys: ' + Object.keys(unitMap).join(', '));
    const topUnit = getTopN(unitMap, 8);

    const finalResult = {
        success: true,
        stats: {
           total: totalAllTime,
           active: activeParticipants,
           univCount: Object.keys(univMap).length,
        },
        trend: { labels: trendLabels, data: trendData },
        university: { labels: topUniv.map(i=>i[0]), data: topUniv.map(i=>i[1]) },
        major: { labels: topMajor.map(i=>i[0]), data: topMajor.map(i=>i[1]) },
        unit: { labels: topUnit.map(i=>i[0]), data: topUnit.map(i=>i[1]) }
    };
    
    setCachedData(CACHE_KEY, finalResult, 3600); // 1 jam

    return finalResult;
    
  } catch (error) {
    Logger.log('Error getting analytics: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// ===== TEST FUNCTION =====
function testAddMagang() {
  Logger.log('===== TESTING addMagang() =====');
  
  const testData = {
    'NOMOR SURAT': 'TEST-BACKEND-001',
    'UNIVERSITAS': 'Test University Backend',
    'JURUSAN': 'Test Jurusan Backend',
    'NAMA MAHASISWA': 'Test Student Backend',
    'STAMBUK / NIM': 'BACKEND-99999',
    'TGL MASUK': '2026-01-13',
    'TGL SELESAI': '2026-03-13',
    'PENGIRIM': 'Test Backend Sender',
    'REKOMENDASI': ''  // Empty - should NOT highlight
  };
  
  Logger.log('Test data: ' + JSON.stringify(testData));
  
  const result = addMagang(testData);
  
  Logger.log('Result: ' + JSON.stringify(result));
  
  if (result.success) {
    Logger.log('âœ… SUCCESS! Check Google Sheets for new row with No: ' + result.autoNumber);
    Logger.log('Verify: NAMA MAHASISWA column has "Test Student Backend"');
    Logger.log('Verify: STAMBUK / NIM column has "BACKEND-99999"');
  } else {
    Logger.log('âŒ FAILED! Message: ' + result.message);
  }
  
  return result;
}

/**
 * Get Magang Data by NOMOR SURAT (Specific Sheet or Global Search)
 */
function getDataByNomorSurat(nomorSurat, sheetName) {
  try {
    Logger.log('getDataByNomorSurat called with: ' + nomorSurat + ', Sheet: ' + (sheetName || 'ALL'));
    
    const ss = getSpreadsheet();
    
    // Determine which sheets to scan
    let sheetsToScan = [];
    
    if (sheetName) {
      // [SPECIFIC SEARCH] If sheetName is provided, only scan that sheet
      sheetsToScan = [sheetName];
      Logger.log('Scanning SPECIFIC sheet: ' + sheetName);
    } else {
      // [GLOBAL SEARCH - Legacy Behavior] Scan ALL Magang Sheets + Data Magang
      const batchResult = getMagangBatches();
      if (batchResult.success) {
        sheetsToScan = batchResult.batches;
      }
      // Always include the Master/Default sheet as fallback or primary
      if (!sheetsToScan.includes('Data Magang')) {
        sheetsToScan.push('Data Magang');
      }
      Logger.log('Scanning GLOBAL sheets: ' + sheetsToScan.join(', '));
    }
    
    let foundSheet = null;
    let sheetData = null;
    let filteredRows = [];
    let headers = [];
    
    // LOOP SEARCH
    for (const sName of sheetsToScan) {
      const sheet = ss.getSheetByName(sName);
      if (!sheet) {
        Logger.log('Skip: Sheet not found ' + sName);
        continue;
      }
      
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length <= 1) continue;
      
      const currHeaders = data[0];
      const currRows = data.slice(1);
      
      // Find NOMOR SURAT column (Robust)
      let nsIdx = -1;
      const candidates = ['NOMOR SURAT', 'NO. SURAT', 'NO SURAT', 'NOSURAT'];
      for (let i = 0; i < currHeaders.length; i++) {
        if (candidates.includes(currHeaders[i].toString().toUpperCase().trim())) {
          nsIdx = i;
          break;
        }
      }
      
      if (nsIdx === -1) continue;
      
      const searchStr = String(nomorSurat).trim();
      // Filter rows that match the nomor surat
      const matches = currRows.filter(row => String(row[nsIdx]).trim() === searchStr);
      
      if (matches.length > 0) {
        foundSheet = sName;
        sheetData = data;
        filteredRows = matches;
        headers = currHeaders; // Capture headers of the found sheet
        Logger.log('FOUND DATA in sheet: ' + sName);
        break; // Stop scanning once found
      }
    }
    
    if (!foundSheet || filteredRows.length === 0) {
       return {
         success: false,
         message: 'Data dengan Nomor Surat "' + nomorSurat + '" tidak ditemukan' + (sheetName ? ' di sheet ' + sheetName : '')
       };
    }
    
    // --- Helper to map row array to object based on headers ---
    const mapRowToObj = (row, headers) => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h) obj[h] = row[i]; 
      });
      return obj;
    };
    
    // Helper to find value insensitive
    const findVal = (item, ...candidates) => {
        const keys = Object.keys(item);
        for (const c of candidates) {
            const k = keys.find(key => key.toUpperCase().trim() === c.toUpperCase().trim());
            if (k && item[k]) return item[k];
        }
        return '';
    };

    // 1. Process Student Data & Normalize Keys for PDF Generation (CRITICAL FIX)
    const mahasiswa = filteredRows.map(row => {
        const rawObj = mapRowToObj(row, headers);
        
        // Add normalized keys that generateSuratBalasanPDF expects
        rawObj.nama = findVal(rawObj, 'NAMA MAHASISWA', 'NAMA', 'NAME');
        rawObj.nis = findVal(rawObj, 'STAMBUK', 'NIS', 'NIM', 'NO. INDUK');
        rawObj.jurusan = findVal(rawObj, 'JURUSAN', 'PRODI', 'PROGRAM STUDI');
        
        return rawObj;
    });
    
    // 2. Extract Common Data (from first match)
    const firstMatch = mahasiswa[0];
    
    const commonData = {
      nomorSurat: findVal(firstMatch, 'NOMOR SURAT', 'NO. SURAT'),
      tglSurat: findVal(firstMatch, 'TGL MSK SURAT', 'TANGGAL MASUK SURAT', 'TGL SURAT'),
      perihal: findVal(firstMatch, 'PERIHAL', 'HAL'),
      pengirim: findVal(firstMatch, 'PENGIRIM', 'NAMA PENGIRIM'),
      
      universitas: findVal(firstMatch, 'UNIVERSITAS', 'ASAL SEKOLAH', 'INSTITUSI'),
      jurusan: findVal(firstMatch, 'JURUSAN', 'PRODI'),
      
      tglMulai: findVal(firstMatch, 'TGL MASUK', 'TANGGAL MULAI', 'START DATE'),
      tglSelesai: findVal(firstMatch, 'TGL SELESAI', 'TANGGAL SELESAI', 'END DATE'),
      
      unit: findVal(firstMatch, 'UNIT KERJA', 'PENEMPATAN', 'TEMBUSAN'), // Added TEMBUSAN candidate
      pembimbing: findVal(firstMatch, 'PEMBIMBING')
    };
    
    return {
      success: true,
      data: {
        mahasiswa: mahasiswa,
        common: commonData,
        totalCount: mahasiswa.length
      },
      warning: mahasiswa.length > 8 ? 'Jumlah mahasiswa (' + mahasiswa.length + ') melebihi kapasitas template standar (8 orang). Beberapa data mungkin tidak muat.' : null
    };
    
  } catch (error) {
    Logger.log('Error in getDataByNomorSurat: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.toString()
    };
  }
}


// ===== SURAT BALASAN GENERATION FUNCTIONS =====

/**
 * Get unique NOMOR SURAT values for dropdown (Filtered by Sheet)
 */

function getMagangSuratOptions(sheetName, requireValidation) {
  try {
    Logger.log('=== GET MAGANG SURAT OPTIONS START ===');
    Logger.log('Target Sheet: ' + (sheetName || 'Default (Data Magang)') + ', Filter Validation: ' + requireValidation);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Use provided sheetName or fallback to default 'Data Magang'
    const targetSheetName = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) {
      Logger.log('ERROR: Sheet not found: ' + targetSheetName);
      return { success: false, message: 'Sheet "' + targetSheetName + '" tidak ditemukan' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('No data rows in ' + targetSheetName);
      return { success: true, data: [] };
    }
    
    // Get headers to find NOMOR SURAT column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Find column index (case-insensitive robust search)
    let nomorSuratIndex = -1;
    let validationIndex = -1;
    const candidates = ['NOMOR SURAT', 'NO. SURAT', 'NO SURAT', 'NOSURAT'];
    const validCandidates = ['VALIDATION', 'VALIDASI', 'STATUS VALIDASI', 'KET']; // Filter Column Candidates
    
    for (let i = 0; i < headers.length; i++) {
        const h = headers[i].toString().toUpperCase().trim();
        if (nomorSuratIndex === -1 && candidates.includes(h)) {
            nomorSuratIndex = i;
        }
        if (validationIndex === -1 && validCandidates.includes(h)) {
             validationIndex = i;
        }
    }
    
    if (nomorSuratIndex === -1) {
      Logger.log('ERROR: NOMOR SURAT column not found in ' + targetSheetName);
      return { success: false, message: 'Kolom NOMOR SURAT tidak ditemukan di sheet ' + targetSheetName };
    }
    
    // Get all data (we need full rows now to check validation)
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Count occurrences
    const suratMap = {};
    data.forEach(row => {
      const nomor = row[nomorSuratIndex];
      // Validation Check
      let isValid = true;
      
      // ONLY check validation if requested AND column exists
      if (requireValidation) {
          if (validationIndex !== -1) {
              const valStatus = (row[validationIndex] || '').toString().toLowerCase().trim();
              // Filter Condition: Must be explicitly "sudah diteruskan"
              if (!valStatus.includes('sudah diteruskan')) {
                  isValid = false;
              }
          } else {
              // Strict mode: If validation required but column missing -> warn but allow? 
              // Or block? To be safe, usually better to allow or log warning.
              // For now, let's just log and allow to avoid broken lists if someone renamed column.
              // But logically if I want "Sudah Diteruskan", missing column = "unknown" = maybe invalid.
              // Let's stick to safe: If column found, filter. If not, allow (fallback).
          }
      }

      if (isValid && nomor && nomor.toString().trim() !== '' && nomor.toString().trim() !== '-') {
        const key = nomor.toString().trim();
        suratMap[key] = (suratMap[key] || 0) + 1;
      }
    });
    
    // Convert to array
    const options = Object.keys(suratMap).map(nomor => ({
      nomorSurat: nomor,
      count: suratMap[nomor]
    })).sort((a, b) => a.nomorSurat.localeCompare(b.nomorSurat));
    
    Logger.log('Found ' + options.length + ' unique nomor surat in ' + targetSheetName);
    return { success: true, data: options };
  } catch (error) {
    Logger.log('Error in getMagangSuratOptions: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Generate Surat Balasan PDF from template
 */
function generateSuratBalasanPDF(nomorSurat, headerData, sheetName) {
  try {
    Logger.log('=== GENERATE SURAT BALASAN PDF (FAST CLONE MODE) ===');
    Logger.log('Nomor: ' + nomorSurat + ', Sheet: ' + (sheetName || 'Global Search'));
    
    // 1. Get data (Pass sheetName for specific lookup)
    const dataResult = getDataByNomorSurat(nomorSurat, sheetName);
    if (!dataResult.success) return dataResult;
    
    const { mahasiswa, common } = dataResult.data;
    
    // 3. Access Template File & Check Type
    let templateFile;
    try {
      templateFile = DriveApp.getFileById(TEMPLATE_SURAT_BALASAN_ID);
      const mimeType = templateFile.getMimeType();
      Logger.log('Template MimeType: ' + mimeType);
      
      if (mimeType === MimeType.MICROSOFT_WORD || mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
         return { 
           success: false, 
           message: 'File Template Anda berformat Microsoft Word (.docx). Sistem hanya bisa memproses Google Docs. Silakan buka file template di Google Drive -> File -> Simpan sebagai Google Docs, lalu gunakan ID file baru tersebut.' 
         };
      }
      
      if (mimeType !== MimeType.GOOGLE_DOCS) {
         return { 
           success: false, 
           message: 'Format File Template bukan Google Docs (' + mimeType + '). Silakan convert file .docx Anda ke Google Docs terlebih dahulu.' 
         };
      }
    } catch (e) {
      return { success: false, message: 'Gagal mengakses template. Error: ' + e.toString() };
    }
    
    // 4. COPY Template (Preserves Layout 100%)
    let copiedDoc;
    let docId;
    // Custom Filename: SB a.n. [Name] [, DKK]
    const firstStudentName = (mahasiswa && mahasiswa.length > 0 && mahasiswa[0].nama) ? mahasiswa[0].nama : 'Mahasiswa';
    const suffix = (mahasiswa && mahasiswa.length > 1) ? ', DKK' : '';
    const templateName = 'SB a.n. ' + firstStudentName + suffix;
    
    try {
      copiedDoc = templateFile.makeCopy(templateName);
      docId = copiedDoc.getId();
      
      // Unlock permissions immediately
      try { copiedDoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      
    } catch (e) {
      return { success: false, message: 'Gagal menyalin template. Error: ' + e.toString() };
    }

    // 5. Open Document with RETRY (Fixes Race Condition)
    let doc;
    const maxRetries = 5; 
    let waitTime = 2000; 
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        Utilities.sleep(waitTime);
        doc = DocumentApp.openById(docId);
        if (doc) break; 
      } catch (e) {
        Logger.log('Attempt ' + attempt + ' failed: ' + e.toString());
        if (attempt === maxRetries) {
          // Cleanup if failed
          try { copiedDoc.setTrashed(true); } catch(err) {} 
          return { success: false, message: 'Timeout: Gagal membuka dokumen baru. Google Drive sedang lambat. Silakan coba lagi.' };
        }
        waitTime *= 2; 
      }
    }

    // 6. Replace Text
    const body = doc.getBody();
    try {
      body.replaceText('«Nomor»', headerData.nomorBalasan || '');
      body.replaceText('«Lampiran»', headerData.lampiran || '-');
      body.replaceText('«Perihal»', headerData.perihal || '');
      body.replaceText('«Kepada»', headerData.kepada || '');
      body.replaceText('«Jabatan»', headerData.jabatan || '');
      body.replaceText('«Sekolah»', common.universitas || '');
      body.replaceText('«Sheet1Jurusan_»', common.jurusan || '');
      
      // Helper to format date string to Indonesian format with Timezone safety
      function safeFormatDate(dateInput, addBuffer = true) {
        if (!dateInput) return '';
        try {
           let d;
           if (dateInput instanceof Date) {
             d = new Date(dateInput.getTime());
           } else {
             d = new Date(dateInput);
           }
           
           if (!isNaN(d.getTime())) {
              // 1. Timezone Safety Fix
              // Only add buffer if it's a Sheet date (usually midnight) to avoid -1 day rollback
              // For "Now" (new Date()), we trust the system time.
              if (addBuffer) {
                d.setHours(d.getHours() + 12);
              }
              
              // 2. Format to English first
              const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMMM yyyy');
              
              // 3. Translate to Indonesian
              const months = {
                'January': 'Januari',
                'February': 'Februari',
                'March': 'Maret',
                'April': 'April',
                'May': 'Mei',
                'June': 'Juni',
                'July': 'Juli',
                'August': 'Agustus',
                'September': 'September',
                'October': 'Oktober',
                'November': 'November',
                'December': 'Desember'
              };
              
              let indoDate = dateStr;
              Object.keys(months).forEach(eng => {
                indoDate = indoDate.replace(eng, months[eng]);
              });
              
              return indoDate;
           }
        } catch(e) {
          Logger.log('Date format error: ' + e.toString());
        }
        return dateInput;
      }

      // 1. TGL SURAT (Body Text - "Yellow Box")
      const formattedTglSurat = safeFormatDate(common.tglSurat);
      body.replaceText('«TanggalMasukSurat»', formattedTglSurat || ''); // Exact match from template image
      body.replaceText('«Tanggal»', formattedTglSurat || ''); // Fallback for legacy
      
      // 2. NOMOR SURAT (Body Text - e.g. "Perihal 123")
      body.replaceText('«NomorSurat»', common.nomorSurat || ''); // Exact match
      body.replaceText('«Nomor_Surat»', common.nomorSurat || ''); // Fallback
      
      // 3. JADWAL PELAKSANAAN ("Red Box" - Tgl Masuk s/d Tgl Selesai)
      const tglMulaiStr = safeFormatDate(common.tglMulai);
      const tglSelesaiStr = safeFormatDate(common.tglSelesai);
      const periode = tglMulaiStr && tglSelesaiStr ? tglMulaiStr + ' s/d ' + tglSelesaiStr : tglMulaiStr || tglSelesaiStr || '';
      body.replaceText('«Periode»', periode);
      
      body.replaceText('«Tgl_Pembekalan»', tglMulaiStr);
      
      // Fix: Tembusan = Unit Kerja
      body.replaceText('«Tembusan»', common.unit || '-');
      
      const today = safeFormatDate(new Date(), false); // Disable buffer for "Today"
      body.replaceText('«Tgl_Dibuat_Surat»', today);
      body.replaceText('«AutoMergeField»', '');
      
      // --- DYNAMIC TABLE HANDLING ---
      // 1. Find the table containing student data
      const tables = body.getTables();
      let studentTable = null;
      let startRowIndex = -1;
      
      // Look for the table with the first placeholder
      for (let t = 0; t < tables.length; t++) {
        const table = tables[t];
        const numRows = table.getNumRows();
        for (let r = 0; r < numRows; r++) {
          const rowText = table.getRow(r).getText();
          if (rowText.includes('«Nama_Satu»')) {
            studentTable = table;
            startRowIndex = r; // This is where the student list starts
            break;
          }
        }
        if (studentTable) break;
      }
      
      if (studentTable && startRowIndex !== -1) {
        Logger.log('Found student table at row ' + startRowIndex);
        
        // 2. Identify rows to keep vs remove
        // We assume the template has 8 rows prepared sequentially: StartRow -> StartRow+7
        // number of students to display:
        const count = mahasiswa.length; 
        
        // We only deleting PROSPECTIVE empty rows that are actually in the template
        // placeholders are: Nama_Satu (idx 0) ... Nama_Delapan (idx 7)
        // If we have 2 students, we keep row 0 and 1 (relative to start).
        // We delete row 2, 3, 4, 5, 6, 7.
        
        // Delete from bottom up to preserve indices
        // Max placeholders = 8.
        const maxSlots = 8;
        let rowsRemoved = 0; // Track removed rows to add spacing back
        
        for (let i = maxSlots - 1; i >= count; i--) {
           // i is the student index (0-based) we want to remove. 
           // e.g. if count=2, we remove i=7, 6, 5, 4, 3, 2.
           const targetRowIdx = startRowIndex + i;
           
           // Safety check: ensure row exists and contains the expected placeholder
           // This prevents deleting headers or footers if the table structure is unexpected
           if (targetRowIdx < studentTable.getNumRows()) {
             const row = studentTable.getRow(targetRowIdx);
             const text = row.getText();
             // Check if it really holds the placeholder we expect (e.g. Nama_Tiga for i=2)
             const expectedPlaceholder = '«Nama_' + getNameSuffix(i) + '»'; // Helper check
             
             // We'll trust the index mostly, but check bounds
             try {
                studentTable.removeRow(targetRowIdx);
                rowsRemoved++; // Increment counter
                Logger.log('Removed unused row ' + targetRowIdx + ' (Slot ' + (i+1) + ')');
             } catch(e) {
                Logger.log('Could not remove row ' + targetRowIdx + ': ' + e.toString());
             }
           }
        }
        
        // LAYOUT FIX: COMPACT MODE (AGGRESSIVE)
        // If we have many students (> 4), the table is tall.
        // We should REMOVE ALL empty lines after the table to pull the footer up
        // and ensure everything fits on 1 page.
        if (count > 4) {
           try {
             const tableIndex = body.getChildIndex(studentTable);
             let attempts = 0;
             let linesRemoved = 0;
             
             // Aggressively remove up to 10 empty paragraphs immediately following the table
             // limit iterations to avoid infinite loops
             while (attempts < 20 && linesRemoved < 10) {
                 // Check if next element exists
                 if (tableIndex + 1 < body.getNumChildren()) {
                     const nextSibling = body.getChild(tableIndex + 1);
                     // Check if it's an empty paragraph
                     if (nextSibling.getType() === DocumentApp.ElementType.PARAGRAPH) {
                         const p = nextSibling.asParagraph();
                         // Robust regex to detect empty/whitespace-only paragraphs
                         const text = p.getText();
                         const isEmpty = !text || text.replace(/[\s\u00A0\u200B]+/g, '').length === 0;

                         if (isEmpty) {
                             nextSibling.removeFromParent();
                             linesRemoved++;
                             // Don't increment index, next element takes this spot
                         } else {
                             // Hit non-empty content (e.g. text or another table), stop
                             break; 
                         }
                     } else {
                         break;
                     }
                 } else {
                     break;
                 }
                 attempts++;
             }
             Logger.log('Compact Mode: Removed ' + linesRemoved + ' empty lines to fit page.');
           } catch(e) {
             Logger.log('Error compacting layout: ' + e.toString());
           }
        }

        // NEW: Special handling for 1-3 students (Compact Mode)
        // Center align text in table cells and remove extra spacing after table
        if (count <= 3) {
           try {
              // Iterate through EACH student row to apply centering
              // We have 'count' students, so we process 'count' rows starting from startRowIndex
              for (let i = 0; i < count; i++) {
                  const currentRowIndex = startRowIndex + i;
                  if (currentRowIndex < studentTable.getNumRows()) {
                       const row = studentTable.getRow(currentRowIndex);
                       for (let c = 0; c < row.getNumCells(); c++) {
                          const cell = row.getCell(c);
                          // Determine Alignment: Name (Column 0) = Left, Others = Center
                          const alignment = (c === 0) ? DocumentApp.HorizontalAlignment.LEFT : DocumentApp.HorizontalAlignment.CENTER;
                          
                          // Clear SpacingAfter for paragraphs
                          for (let pIdx = 0; pIdx < cell.getNumChildren(); pIdx++) {
                              const child = cell.getChild(pIdx);
                              if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                                  child.asParagraph().setSpacingAfter(0);
                                  child.asParagraph().setAlignment(alignment);
                              }
                          }
                          // Vertical Alignment (Always Center)
                          cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
                       }
                  }
              }
              Logger.log('Applied CENTER alignment for ' + count + ' student(s).');
              
              // NEW: Remove extra spacing below table for single student
              const tableIndex = body.getChildIndex(studentTable);
              let attempts = 0;
              let linesRemoved = 0;
              
              // Remove up to 20 empty lines after table (AGGRESSIVE with Regex)
              while (attempts < 20) {
                  if (tableIndex + 1 < body.getNumChildren()) {
                      const nextSibling = body.getChild(tableIndex + 1);
                      if (nextSibling.getType() === DocumentApp.ElementType.PARAGRAPH) {
                          const p = nextSibling.asParagraph();
                          // Regex handles regular space, tabs, non-breaking space (\u00A0), etc.
                          const text = p.getText();
                          const isEmpty = !text || text.replace(/[\s\u00A0\u200B]+/g, '').length === 0;
                          
                          if (isEmpty) {
                              nextSibling.removeFromParent();
                              linesRemoved++;
                          } else {
                              // Found non-empty content!
                              // Pull it up by removing top margin
                              p.setSpacingBefore(0);
                              Logger.log('Removed top spacing from following paragraph.');
                              break;
                          }
                      } else {
                          // Not a paragraph (maybe a list item or table), stop but try to close gap
                          break;
                      }
                  } else {
                      break;
                  }
                  attempts++;
              }
              Logger.log('Single Student Mode: Removed ' + linesRemoved + ' empty lines after table.');
              
              // Add exactly 1 empty line as requested ("satu spasi saja")
              body.insertParagraph(tableIndex + 1, "");

           } catch(e) {
              Logger.log('Error applying center alignment / spacing: ' + e.toString());
           }
        }
        
        // [REMOVED] SPACER MODE for 1-3 Students
        // This was conflicting with the single-student compact logic.
        // We now rely on the aggressive whitespace removal above logic for balanced layout.
      }
      
      // 3. Fill data for EXISTING students (Standard replaceText works fine now)
      const namaFields = ['Nama_Satu', 'Nama_dua', 'Nama_tiga', 'Nama_Empat', 'Nama_Lima', 'Nama_Enam', 'Nama_Tujuh', 'Nama_Delapan'];
      // FIX: Template uses "Nis_" (Title Case), not "NIS_" (All Caps) for some fields
      // We'll update to match the likely template format or try both.
      // Based on screenshot "Nis_Tujuh", let's use Title Case for the array or just replaceText broadly?
      // Replacing the array definition to match user's template:
      const nisFields = ['NIS_Satu', 'NIS_dua', 'NIS_tiga', 'NIS_Empat', 'NIS_Lima', 'NIS_Enam', 'Nis_Tujuh', 'NIS_Delapan'];
      // Ideally we should handle all, but let's stick to what we see or try both variants in loop
      
      for (var i = 0; i < mahasiswa.length; i++) {
          body.replaceText('«' + namaFields[i] + '»', mahasiswa[i].nama || '');
          
          // Try both NIS variants to be safe
          body.replaceText('«' + nisFields[i] + '»', mahasiswa[i].nis || ''); // Logic existing
          body.replaceText('«Nis_' + getNameSuffix(i) + '»', mahasiswa[i].nis || ''); // Fallback for "Nis_"
      }
      // Note: We don't need to clear "else" placeholders because we deleted those rows!
      
      doc.saveAndClose();
      
    } catch (e) {
       return { success: false, message: 'Gagal edit text. ' + e.toString() };
    }
    
    // 7. Export PDF & Cleanup
    let downloadUrl;
    let viewUrl;
    try {
      Utilities.sleep(1000); // Flush
      
      const pdfBlob = copiedDoc.getAs('application/pdf');
      pdfBlob.setName(templateName + '.pdf');
      
      const pdfFile = DriveApp.createFile(pdfBlob);
      viewUrl = pdfFile.getUrl();
      downloadUrl = 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId();
      // viewUrl already assigned above
      
      // Share PDF Publicly so user can download
      try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}

      // DELETE TEMP DOC (Clean up)
      // DISABLED for DOCX Download support - File must remain for link to work
      // try { copiedDoc.setTrashed(true); } catch(e) {}
      
    } catch (e) {
      return { success: false, message: 'Gagal export PDF. ' + e.toString() };
    }
    
    // Construct preview URL safely
    const previewUrl = (viewUrl && viewUrl.includes('/view')) ? viewUrl.replace('/view', '/preview') : viewUrl;
    const docxUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=docx';

    return {
      success: true,
      message: 'Berhasil! (Standard Mode)',
      pdfUrl: downloadUrl,
      previewUrl: previewUrl, 
      docUrl: viewUrl, 
      fileName: templateName + '.pdf',
      docxUrl: docxUrl,
      warning: dataResult.warning
    };
    
  } catch (error) {
    Logger.log('Fatal Error: ' + error.toString());
    return { success: false, message: 'Fatal Error: ' + error.toString() };
  }
}

// TEST FUNCTION
function testGenerateSuratBalasan() {
  // Get first nomor surat for testing
  const options = getMagangSuratOptions();
  if (options.success && options.data.length > 0) {
    const testNomor = options.data[0].nomorSurat;
    Logger.log('Testing with: ' + testNomor);
    const result = generateSuratBalasanPDF(testNomor);
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  } else {
    Logger.log('No nomor surat found for testing');
    return { success: false, message: 'No data to test' };
  }
}


// Helper for dynamic table row deletion
function getNameSuffix(index) {
  const suffixes = ['Satu', 'dua', 'tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan'];
  return suffixes[index] || '';
}

/**
 * TEST FUNCTION - Run this to trigger Drive permission authorization
 * This will request permission to access Google Drive for template copying
 */
function testDriveAccessNew() {
  try {
    Logger.log('=== TESTING DRIVE ACCESS ===');
    
    // Test 1: Access template file
    const templateFile = DriveApp.getFileById(TEMPLATE_SURAT_BALASAN_ID);
    Logger.log('âœ“ Template file accessible: ' + templateFile.getName());
    
    // Test 2: Access spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('âœ“ Spreadsheet accessible: ' + ss.getName());
    
    // Test 3: Try to copy template (will be deleted immediately)
    const testCopy = templateFile.makeCopy('TEST - DELETE ME');
    Logger.log('âœ“ Can create copies: ' + testCopy.getName());
    
    // Clean up test file
    testCopy.setTrashed(true);
    Logger.log('âœ“ Test copy deleted');
    
    Logger.log('=== ALL PERMISSIONS OK ===');
    return { success: true, message: 'All permissions granted successfully!' };
    
  } catch (error) {
    Logger.log('=== PERMISSION ERROR ===');
    Logger.log('Error: ' + error.toString());
    return { 
      success: false, 
      message: 'Permission error: ' + error.toString(),
      instructions: 'Click "Review Permissions" and authorize the app'
    };
  }
}

/**
 * DEBUG FUNCTION - Check what permissions are actually active
 * Run this to see if "https://www.googleapis.com/auth/drive" is present
 */
function checkActiveScopes() {
  const token = ScriptApp.getOAuthToken();
  const parts = token.split('.');
  if (parts.length > 1) {
    const payload = JSON.parse(Utilities.newBlob(Utilities.base64Decode(parts[1])).getDataAsString());
    Logger.log('=== CURRENT ACTIVE SCOPES ===');
    const scopes = payload.scope.split(' ');
    scopes.forEach(s => Logger.log('- ' + s));
    
    if (scopes.some(s => s.includes('drive'))) {
      Logger.log('âœ… DRIVE SCOPE FOUND!');
    } else {
      Logger.log('âŒ DRIVE SCOPE MISSING (Manifest not applied?)');
    }
  } else {
    Logger.log('Could not decode token');
  }
}
// MAINTENANCE FUNCTION: Fix All Date Formats
function fixAllMagangDates() {
  try {
    const ss = getSpreadsheet();
    
    // Get all sheets
    const allSheets = ss.getSheets();
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Notifikasi'];
    
    let totalFixed = 0;
    let fixedSheets = [];
    
    allSheets.forEach(sheet => {
      const name = sheet.getName();
      
      // Process if it's Data Magang or a Batch Sheet
      if (name === 'Data Magang' || !excludedSheets.includes(name)) {
        Logger.log('Scanning sheet: ' + name);
        
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        
        if (lastRow > 1 && lastCol > 0) {
          const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
          
          // Identify Date Columns
          const dateIndices = [];
          headers.forEach((h, i) => {
             const header = h ? h.toString().toUpperCase() : '';
             if (header.includes('TGL') || header.includes('DATE') || header.includes('WAKTU')) {
               dateIndices.push(i);
             }
          });
          
          if (dateIndices.length > 0) {
            let sheetFixedCount = 0;
            let dirty = false;
            
            // Loop through each date column
            dateIndices.forEach(colIndex => {
               // 1-based index
               const range = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
               const values = range.getValues();
               let columnChanged = false;
               
               const fixedValues = values.map(row => {
                 let val = row[0];
                 if (val && typeof val === 'string') {
                    // Check for YYYY-MM-DD
                    const parts = val.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                    if (parts) {
                       // Convert to Date Object
                       val = new Date(parseInt(parts[1]), parseInt(parts[2]) - 1, parseInt(parts[3]), 12, 0, 0);
                       columnChanged = true;
                       sheetFixedCount++;
                    }
                 }
                 return [val];
               });
               
               if (columnChanged) {
                 // Write back values NOT AS STRINGS, but objects
                 range.setValues(fixedValues);
                 dirty = true;
               }
               
               // ALWAYS Apply Format
               range.setNumberFormat('d mmmm yyyy');
            });
            
            if (dirty) {
               fixedSheets.push(`${name} (${sheetFixedCount} repaired)`);
               totalFixed += sheetFixedCount;
            }
          }
        }
      }
    });
    
    return { 
      success: true, 
      message: `Selesai! Memperbaiki ${totalFixed} data tanggal di: ${fixedSheets.join(', ')}` 
    };
    
  } catch (error) {
    Logger.log('Error fix dates: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Get Batch Metadata (Stats, Date Range)
 */
function getBatchMetadata(batchName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(batchName);
    if (!sheet) return { success: false, message: 'Batch not found' };

    const data = sheet.getDataRange().getValues();
    
    // Default Stats
    const stats = { total: 0, active: 0, finished: 0 };
    let dateRange = batchName; // Fallback
    
    if (data.length > 1) {
        const headers = data[0].map(h => h.toString().trim().toUpperCase());
        const ketIndex = headers.indexOf('KET');
        
        // Find Date Columns (flexible matching)
        const tglMasukIndex = headers.findIndex(h => h.includes('TGL MASUK') || h.includes('TANGGAL MASUK') || h.includes('MULAI'));
        const tglSelesaiIndex = headers.findIndex(h => h.includes('TGL SELESAI') || h.includes('TANGGAL SELESAI') || h.includes('BERAKHIR'));
        
        const rows = data.slice(1);
        stats.total = rows.length;
        
        let minDate = null;
        let maxDate = null;
        
        rows.forEach(r => {
            // Stats Logic
            if (ketIndex !== -1) {
                const status = (r[ketIndex] || '').toString().trim();
                
                // Active Statuses: STRICT (Hanya "Sudah Diteruskan" Saja)
                if (status === 'Sudah Diteruskan') {
                    stats.active++;
                } 
                // Finished Statuses (DITOLAK tidak dihitung sebagai selesai)
                else if (status.toLowerCase() === 'selesai' || 
                         status.toLowerCase() === 'finished' || 
                         status.toLowerCase() === 'resign') {
                    stats.finished++;
                }
            } else {
                // If no KET column, do NOT count as active by default to be safe
            }
            
            // Date Logic
            if (tglMasukIndex !== -1) {
                const d = r[tglMasukIndex];
                if (d instanceof Date && !isNaN(d)) {
                    if (!minDate || d < minDate) minDate = d;
                }
            }
            
            if (tglSelesaiIndex !== -1) {
                const d = r[tglSelesaiIndex];
                if (d instanceof Date && !isNaN(d)) {
                    if (!maxDate || d > maxDate) maxDate = d;
                }
            }
        });
        
        // Format Date Range
        if (minDate && maxDate) {
            const timeZone = ss.getSpreadsheetTimeZone();
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
            
            const formatDate = (d) => {
                // Use Utilities.formatDate to strictly respect the Sheet's timezone
                // "d" = day of month (1-31), "M" = month in year (1-12)
                const day = Utilities.formatDate(d, timeZone, "d"); 
                const monthIndex = parseInt(Utilities.formatDate(d, timeZone, "M")) - 1;
                const year = Utilities.formatDate(d, timeZone, "yyyy");
                return `${day} ${months[monthIndex]} ${year}`;
            };
            
            dateRange = `${formatDate(minDate)} - ${formatDate(maxDate)}`;
        }
    }

    // Determine Status
    let status = 'Active';
    // Logic: If current date is past end date -> Finished?
    // Or if all students finished -> Finished
    if (stats.total > 0 && stats.active === 0) {
        status = 'Finished';
    }

    return {
        success: true,
        data: {
            name: batchName,
            stats: stats,
            status: status,
            date_range: dateRange
        }
    };

  } catch (e) {
      Logger.log('getBatchMetadata Error: ' + e.toString());
      return { success: false, message: e.toString() };
  }
}

/**
 * Rename Magang Sheet (Batch)
 */
function renameMagangSheet(oldName, newName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(oldName);
    
    if (!sheet) {
      return { success: false, message: 'Batch tidak ditemukan.' };
    }
    
    // Check if new name exists
    if (ss.getSheetByName(newName)) {
      return { success: false, message: 'Nama batch sudah ada.' };
    }
    
    sheet.setName(newName);
    return { success: true };
    
  } catch (error) {
    Logger.log('Rename Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Delete Magang Sheet (Batch)
 */
function deleteMagangSheet(batchName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(batchName);
    
    if (!sheet) {
      return { success: false, message: 'Batch tidak ditemukan.' };
    }
    
    ss.deleteSheet(sheet);
    return { success: true };
    
  } catch (error) {
    Logger.log('Delete Error: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Helper to ensure a value exists in the cell's Data Validation rule (Dropdown).
 * If missing, it adds the value to the allowed list.
 */
function ensureValidationOption(cell, value) {
  try {
    const rule = cell.getDataValidation();
    if (rule) {
      const criteria = rule.getCriteriaType();
      if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        const args = rule.getCriteriaValues();
        let allowedValues = args[0]; // Array of values
        
        // Normalize checking (Case Insensitive)
        const valStr = String(value).trim().toLowerCase();
        const exists = allowedValues.some(v => String(v).trim().toLowerCase() === valStr);
        
        if (!exists) {
          Logger.log('Validation: Adding "' + value + '" to allowed list.');
          // Add original value casing
          allowedValues.push(String(value).trim());
          
          // Rebuild validation rule
          const newRule = rule.copy().withCriteria(criteria, [allowedValues]).build();
          cell.setDataValidation(newRule);
        }
      }
    }
  } catch (e) {
    Logger.log('Error ensureValidationOption: ' + e.toString());
  }
}

// ===== AUTHENTICATION & SECURITY =====

// ===== GEMINI API INTEGRATION =====

function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function encodePassword(password) {
  if (!password) return '';
  const prefix = "diklat_";
  return Utilities.base64Encode(Utilities.newBlob(prefix + password).getBytes());
}

function decodePassword(encoded) {
  if (!encoded) return '';
  encoded = String(encoded);
  // Kompatibilitas untuk hash lama SHA-256
  if (encoded.length === 64 && /^[0-9a-f]{64}$/i.test(encoded)) return encoded;
  
  try {
    const raw = Utilities.newBlob(Utilities.base64Decode(encoded)).getDataAsString();
    if (raw.startsWith("diklat_")) {
      return raw.slice(7);
    }
  } catch (e) {
    // Lewati dan return text aslinya jika bukan base64 valid
  }
  return encoded;
}

function doLogin(email, password) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    if (!sheet) return { success: false, message: 'Database login tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const passIdx = headers.indexOf('password');
    const statusIdx = headers.indexOf('status');
    const roleIdx = headers.indexOf('role');
    const nameIdx = headers.indexOf('user'); 
    
    if (emailIdx === -1 || passIdx === -1) return { success: false, message: 'Struktur database login tidak valid' };
    
    const hashedPassword = hashPassword(password);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sheetEmail = String(row[emailIdx] || '').trim().toLowerCase();
      const inputEmail = String(email || '').trim().toLowerCase();
      const sheetPasswordRaw = String(row[passIdx] || '');
      const sheetPasswordPlain = decodePassword(sheetPasswordRaw);
      
      // Check if password matches either plain text OR the legacy hashed version
      if (sheetEmail === inputEmail && (sheetPasswordPlain === password || sheetPasswordRaw === hashedPassword)) {
        if (row[statusIdx] && row[statusIdx].toLowerCase() !== 'active') {
          return { success: false, message: 'Akun Anda non-aktif. Hubungi Admin.' };
        }
        
        // Update Last Active
        const lastActiveIdx = headers.indexOf('last active');
        if (lastActiveIdx !== -1) {
          sheet.getRange(i + 1, lastActiveIdx + 1).setValue(new Date());
        } else {
             let realLastActiveIdx = data[0].indexOf('Last Active');
             if (realLastActiveIdx === -1) {
                 sheet.getRange(1, data[0].length + 1).setValue('Last Active');
                 realLastActiveIdx = data[0].length;
             }
             sheet.getRange(i + 1, realLastActiveIdx + 1).setValue(new Date());
        }

        // Assign Developer role to wildan if no role exists
        let finalRole = row[roleIdx] || 'User';
        if (!row[roleIdx] && sheetEmail.includes('wildan')) {
          finalRole = 'Developer';
        }

        return {
          success: true,
          user: {
            name: row[nameIdx] || 'User',
            email: row[emailIdx],
            role: finalRole
          }
        };
      }
    }
    
    return { success: false, message: 'Email atau password salah' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function registerUser(formData) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Login/register');
    
    // Create sheet if not exists
    if (!sheet) {
      sheet = ss.insertSheet('Login/register');
      sheet.appendRow(['TimeStamp', 'Email', 'Password', 'User', 'Status', 'Time', 'Role', 'Last Active']);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailIdx = headers.map(h => h.toString().toLowerCase()).indexOf('email');
    
    // Check duplicates
    if (emailIdx !== -1 && data.length > 1) {
      const emailExists = data.slice(1).some(row => row[emailIdx] === formData.email);
      if (emailExists) return { success: false, message: 'Email sudah terdaftar' };
    }
    
    const timeStamp = new Date();
    
    // Create an empty array with the same length as headers
    const newRow = new Array(headers.length).fill('');
    
    const lowerHeaders = headers.map(h => h.toString().toLowerCase());
    const passIdx = lowerHeaders.indexOf('password');
    const userIdx = lowerHeaders.indexOf('user');
    const statusIdx = lowerHeaders.indexOf('status');
    const timeIdx = lowerHeaders.indexOf('time');
    const roleIdx = lowerHeaders.indexOf('role');
    const lastActiveIdx = lowerHeaders.indexOf('last active');
    const tsIdx = lowerHeaders.indexOf('timestamp');

    if (tsIdx !== -1) newRow[tsIdx] = timeStamp;
    if (emailIdx !== -1) newRow[emailIdx] = formData.email;
    if (passIdx !== -1) newRow[passIdx] = encodePassword(formData.password); // Encoded base64
    if (userIdx !== -1) newRow[userIdx] = formData.name;
    if (statusIdx !== -1) newRow[statusIdx] = 'Active';
    if (timeIdx !== -1) newRow[timeIdx] = Utilities.formatDate(timeStamp, ss.getSpreadsheetTimeZone(), 'HH:mm');
    if (roleIdx !== -1) newRow[roleIdx] = formData.role || 'User';
    
    // Assign last active if column exists, or push it later
    if (lastActiveIdx !== -1) {
        newRow[lastActiveIdx] = timeStamp;
    } else {
        // If 'Last Active' is magically not in lowerHeaders but in sheet
        newRow.push(timeStamp);
    }
    
    sheet.appendRow(newRow);
    return { success: true, message: 'Registrasi berhasil' };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function sendHeartbeat(email) {
  if (!email) return;
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailIdx = headers.indexOf('email');
    
    // Find Header 'Last Active' or create it
    let lastActiveIdx = headers.indexOf('last active');
    if (lastActiveIdx === -1) {
        // Create if missing
        sheet.getRange(1, headers.length + 1).setValue('Last Active');
        lastActiveIdx = headers.length;
    }
    
    // Find row
    for(let i=1; i<data.length; i++) {
        if(data[i][emailIdx] === email) {
            sheet.getRange(i+1, lastActiveIdx+1).setValue(new Date());
            return;
        }
    }
  } catch(e) {}
}

function getOnlineUsers() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    if (!sheet) return { success: true, users: [] };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const nameIdx = headers.indexOf('user');
    const lastActiveIdx = headers.indexOf('last active');
    const roleIdx = headers.indexOf('role');
    
    if (lastActiveIdx === -1) return { success: true, users: [] };
    
    const now = new Date();
    const cutoff = new Date(now.getTime() - 2 * 60 * 1000); // 2 minutes ago
    
    const onlineUsers = [];
    
    for(let i=1; i<data.length; i++) {
        const lastActive = data[i][lastActiveIdx];
        if (lastActive instanceof Date && lastActive > cutoff) {
            onlineUsers.push({
                name: data[i][nameIdx] || 'User',
                role: data[i][roleIdx] || 'User',
                initial: (data[i][nameIdx] || 'U').charAt(0).toUpperCase()
            });
        }
    }
    
    return { success: true, users: onlineUsers };
    
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getDashboardStats() {
  try {
    const ss = getSpreadsheet();
    const magangSheet = ss.getSheetByName('Data Magang');
    const karyawanSheet = ss.getSheetByName('Data Karyawan'); // Adjust if sheet name differs
    
    // Default Stats
    const stats = {
        karyawan: 0,
        magang: 0,
        total: 0
    };
    
    // 1. Get Magang Count & Activity (Use Real Log)
    const activity = getRecentActivity(); // [FIXED] Use real activity log instead of mock data
    let magangExpiring = [];
    let magangActive = 0;
    let magangFinished = 0;
    
    // Distributions
    const unitMap = {};
    
    if (magangSheet) {
        // Fix: Use getLastRow - 1 for data count (excluding header)
        const lastRow = magangSheet.getLastRow();
        if (lastRow > 1) {
            stats.magang = lastRow - 1;
            
            // Get Data for Activity & stats
            // Columns: No(0), Nama(1), Unit(2), Start(3), End(4)... based on common structure
            // We need to hunt headers to be sure
            const headers = magangSheet.getRange(1, 1, 1, magangSheet.getLastColumn()).getValues()[0];
            const nameIdx = headers.findIndex(h => ['NAMA', 'NAMA MAHASISWA', 'NAME'].includes(h.toString().toUpperCase()));
            const unitIdx = headers.findIndex(h => ['UNIT', 'UNIT KERJA', 'PENEMPATAN'].includes(h.toString().toUpperCase()));
            const statusIdx = headers.findIndex(h => ['STATUS', 'KET'].includes(h.toString().toUpperCase()));
            const startIdx = headers.findIndex(h => ['TGL MASUK', 'TANGGAL MULAI'].includes(h.toString().toUpperCase()));
            const endIdx = headers.findIndex(h => ['TGL SELESAI', 'TANGGAL SELESAI'].includes(h.toString().toUpperCase()));
            
            const rawData = magangSheet.getRange(2, 1, Math.min(lastRow - 1, 100), headers.length).getValues();
            
            // Calculate Active/Finished & Expiring
            const today = new Date();
            const sevenDaysInfo = new Date(today);
            sevenDaysInfo.setDate(today.getDate() + 7);
            
            rawData.forEach(row => {
                const status = (row[statusIdx] || '').toString().toLowerCase();
                const endDate = row[endIdx] instanceof Date ? row[endIdx] : new Date(row[endIdx]);
                
                // Active/Finished Logic
                if (status.includes('selesai') || status.includes('tamat')) {
                    magangFinished++;
                } else {
                    magangActive++;
                    
                    // Expiring Check
                    if (!isNaN(endDate.getTime())) {
                       if (endDate >= today && endDate <= sevenDaysInfo) {
                           const timeDiff = endDate.getTime() - today.getTime();
                           const daysLeft = Math.ceil(timeDiff / (1000 * 3600 * 24));
                           if (daysLeft >= 0) {
                               magangExpiring.push({
                                   name: row[nameIdx],
                                   unit: row[unitIdx] || '-',
                                   daysLeft: daysLeft,
                                   endDate: Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'dd MMM')
                               });
                           }
                       }
                    }
                }
                
                // Unit Distribution
                const unit = (row[unitIdx] || 'Lainnya').toString().trim();
                if(unit) unitMap[unit] = (unitMap[unit] || 0) + 1;
            });
        }
    }
    
    // 2. Get Karyawan Count
    if (karyawanSheet) {
         const lastRow = karyawanSheet.getLastRow();
         if (lastRow > 1) {
             stats.karyawan = lastRow - 1;
         }
    }
    
    stats.total = stats.karyawan + stats.magang;
    
    // Sort Activity by "time" is fake, but list is already built recent-first roughly
    
    // Online Users
    const onlineResult = getOnlineUsers(); // Reuse existing
    
    // Chart Data
    const unitLabels = Object.keys(unitMap).sort((a,b) => unitMap[b] - unitMap[a]).slice(0, 5); // Top 5
    const unitData = unitLabels.map(l => unitMap[l]);
    
    return {
        success: true,
        stats: stats,
        activity: activity,
        onlineUsers: onlineResult.users || [],
        widgets: {
            expiring: magangExpiring,
            composition: {
                active: magangActive,
                finished: magangFinished
            },
            unit: {
                labels: unitLabels,
                data: unitData
            }
        }
    };
    
  } catch (e) {
    Logger.log('Error in getDashboardStats: ' + e.toString());
    return { 
        success: false, 
        message: e.toString(),
        stats: { karyawan: 0, magang: 0, total: 0 } 
    };
  }
}

// ===== EXPORT SURAT PENGANTAR =====

function generateSuratPengantar(payload) {
  try {
    const nomorSurat = payload.nomorSurat;
    const tglSurat = payload.tglSurat || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yyyy');
    const customNomorSurat = payload.nomorSuratKeluar; // Custom format string
    // inputTembusan is ignored, we map to Unit
    const templateId = '19aHmpSrzjsVKaEIl0K7KgCxkiV2xu5xMSVueL6lxSCU'; // Fixed Template ID
    
    // 1. Fetch Students by Nomor Surat
    const ss = getSpreadsheet();
    const sheetName = payload.sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found: ' + sheetName };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find relevant columns
    const idxNomor = headers.findIndex(h => h.toString().toUpperCase() === 'NOMOR SURAT');
    const idxUnit = headers.findIndex(h => h.toString().toUpperCase() === 'UNIT KERJA');
    const idxNama = headers.findIndex(h => h.toString().toUpperCase().includes('NAMA'));
    const idxStambuk = headers.findIndex(h => h.toString().toUpperCase().includes('STAMBUK') || h.toString().toUpperCase().includes('NIM'));
    const idxJurusan = headers.findIndex(h => h.toString().toUpperCase() === 'JURUSAN');
    const idxPerihal = headers.findIndex(h => h.toString().toUpperCase() === 'PERIHAL');
    const idxTglMulai = headers.findIndex(h => h.toString().toUpperCase().includes('TGL MASUK') || h.toString().toUpperCase().includes('TGL MULAI'));
    const idxTglSelesai = headers.findIndex(h => h.toString().toUpperCase().includes('TGL SELESAI'));
    
    if (idxNomor === -1) return { success: false, message: 'Kolom NOMOR SURAT tidak ditemukan' };
    
    // 2. PREPARE ALL UNITS TO PROCESS (From Single or Multiple Letters)
    const nomorSuratList = payload.nomorSuratList || [nomorSurat];
    const allProcessingItems = []; // Array of { unitName, students, nomorSurat: "..." }

    // Loop through each requested Nomor Surat
    nomorSuratList.forEach(currentNomor => {
        // Filter rows for this specific letter
        const students = data.slice(1).filter(row => row[idxNomor] === currentNomor);
        
        if (students.length > 0) {
            // Group by UNIT KERJA for this letter
            const groupedData = {};
            students.forEach(row => {
               let unit = row[idxUnit] || 'Tanpa Unit';
               // Normalize spaces: Trim and replace multiple spaces/tabs with single space
               unit = unit.toString().trim().replace(/\s+/g, ' ');
               if (!groupedData[unit]) groupedData[unit] = [];
               groupedData[unit].push(row);
            });
            
            // Add to master list
            Object.entries(groupedData).forEach(([uName, uStudents]) => {
                allProcessingItems.push({
                    unitName: uName,
                    unitStudents: uStudents,
                    nomorSurat: currentNomor
                });
            });
        }
    });
    
    // 3. Process each Item (Unit within a Letter)
    if (allProcessingItems.length === 0) return { success: false, message: 'Tidak ada data unit ditemukan untuk surat yang dipilih.' };

    // --- MERGE STRATEGY: Master Doc is the FIRST Item's Doc ---
    // We will append subsequent items to this Master Doc
    
    // Use the FIRST item to initialize the Master Doc
    const firstItem = allProcessingItems[0];
    // Template File
    const templateFile = DriveApp.getFileById(templateId);
    
    // Create Master Doc (Filename based on first or generic if multiple)
    const masterFileName = (nomorSuratList.length > 1) 
        ? `Surat_Pengantar_Gabungan_${nomorSuratList.length}_Surat` 
        : `Surat_Pengantar_${firstItem.nomorSurat}`;
        
    const masterFile = templateFile.makeCopy(masterFileName);
    const masterDocId = masterFile.getId();
    const masterDoc = DocumentApp.openById(masterDocId);
    let masterBody = masterDoc.getBody();
    
    // Helper to process a Body (Replace text + Format Table)
    // NOW ACCEPTS dynamic 'currentNomor' argument
    const processBodyForUnit = (bodyToProcess, unitName, unitStudents, currentNomor) => {
       const perihal = (idxPerihal !== -1 && unitStudents[0][idxPerihal] != null) ? String(unitStudents[0][idxPerihal]) : 'Permohonan Magang';
       
       bodyToProcess.replaceText('«Nomor»', String(customNomorSurat || currentNomor || ''));
       bodyToProcess.replaceText('«Perihal»', String(perihal));
       // Unit Name Replacements
       bodyToProcess.replaceText('«Tembusan»', String(unitName || '')); 
       bodyToProcess.replaceText('«Unit»', String(unitName || ''));
       bodyToProcess.replaceText('«Tgl_Dibuat_Surat»', String(tglSurat || ''));
       
       const idxUniv = headers.findIndex(h => h.toString().toUpperCase() === 'UNIVERSITAS');
       const sekolah = (idxUniv !== -1 && unitStudents[0][idxUniv] != null) ? String(unitStudents[0][idxUniv]) : '';
       bodyToProcess.replaceText('«Sekolah»', String(sekolah));
       
       const program = (String(perihal).toLowerCase().includes('penelitian')) ? 'Penelitian' : 'Praktek Kerja Lapangan';
       bodyToProcess.replaceText('«Program»', program);

       // 4. Handle Dynamic Periode (From Tgl Mulai & Tgl Selesai)
       let periodeStr = '-';
       if (idxTglMulai !== -1 && idxTglSelesai !== -1 && unitStudents.length > 0) {
           const firstStudent = unitStudents[0];
           const dStart = firstStudent[idxTglMulai];
           const dEnd = firstStudent[idxTglSelesai];
           
           const formatIndo = (d) => {
               if (!d || !(d instanceof Date)) return '-';
               const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
               return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
           };
           
           if (dStart && dEnd) {
               periodeStr = `${formatIndo(dStart)} s.d ${formatIndo(dEnd)}`;
           }
       }
       bodyToProcess.replaceText('«Periode»', periodeStr);
       
       // Handle Table (Identical Logic as optimized)
       const tables = bodyToProcess.getTables();
       for (let t = 0; t < tables.length; t++) {
         const table = tables[t];
         if (table.getText().includes('«Nama_Satu»') || table.getText().includes('Nama')) {
           const numRows = table.getNumRows();
           let templateRow = null;

           // Clone Row 1
           if (numRows > 1) {
              templateRow = table.getRow(1).copy(); 
           }
           // Remove existing rows
           for (let r = numRows - 1; r > 0; r--) {
              table.removeRow(r);
           }
           if (!templateRow) templateRow = table.appendTableRow().copy();

           // Add Students
           unitStudents.forEach((studentRow) => {
              const row = table.appendTableRow(templateRow.copy());
              row.setMinimumHeight(0); // Force tight layout

              const updateCell = (colIndex, text, alignType) => {
                  if (colIndex < row.getNumCells()) {
                      const cell = row.getCell(colIndex);
                      cell.setText(text || '-');
                      if (cell.getNumChildren() > 0) {
                          const p = cell.getChild(0).asParagraph();
                          p.setAlignment(alignType);
                          
                          // Fix artifacts: Ensure content is not Bold/Underlined (inherited from header/placeholder)
                          const textElement = p.editAsText();
                          if (textElement) {
                              textElement.setBold(false);
                              textElement.setItalic(false);
                              textElement.setUnderline(false);
                          }
                      }
                      cell.setPaddingTop(1);
                      cell.setPaddingBottom(0);
                  }
              };
              updateCell(0, studentRow[idxNama], DocumentApp.HorizontalAlignment.LEFT);
              updateCell(1, studentRow[idxStambuk], DocumentApp.HorizontalAlignment.CENTER);
              updateCell(2, studentRow[idxJurusan], DocumentApp.HorizontalAlignment.CENTER);
           });
           

         }
       }
    };

    // 1. Process First Item (Master Doc)
    processBodyForUnit(masterBody, firstItem.unitName, firstItem.unitStudents, firstItem.nomorSurat);

    // 2. Process Subsequent Items (Append to Master)
    // Iterate from index 1 (since 0 is already processed)
    for (let i = 1; i < allProcessingItems.length; i++) {
        const item = allProcessingItems[i];
        
        // Append Page Break
        masterBody.appendPageBreak();
        
        // Create Temp Doc for this unit to get fresh layout
        const tempFile = templateFile.makeCopy('TEMP_MERGE_' + i);
        const tempDoc = DocumentApp.openById(tempFile.getId());
        const tempBody = tempDoc.getBody();
        
        // Process the Temp Doc Content FIRST (PASSING SPECIFIC ITEM NOMOR)
        processBodyForUnit(tempBody, item.unitName, item.unitStudents, item.nomorSurat);
        
        // Copy processed elements to Master
        const totalChildren = tempBody.getNumChildren();
        let lastWasTable = false;
        
        for (let j = 0; j < totalChildren; j++) {
            const child = tempBody.getChild(j);
            const type = child.getType();
            
            // Append based on type to preserve formatting
            if (type === DocumentApp.ElementType.PARAGRAPH) {
                const p = child.asParagraph();
                const text = p.getText();
                // Skip empty paragraph immediately following a table (fix spacing double-up)
                if (lastWasTable && (!text || text.trim().length === 0)) {
                   lastWasTable = false; // Reset
                   continue;
                }
                masterBody.appendParagraph(p.copy());
                lastWasTable = false;
            } else if (type === DocumentApp.ElementType.TABLE) {
                masterBody.appendTable(child.copy());
                lastWasTable = true;
            } else if (type === DocumentApp.ElementType.LIST_ITEM) {
                masterBody.appendListItem(child.copy());
                lastWasTable = false;
            } else {
                 try {} catch(e) {}
                 lastWasTable = false;
            }
        }
        
        // Trash temp
        tempFile.setTrashed(true);
    }
    
    masterDoc.saveAndClose();
    
    // Export Master PDF
    const pdfBlob = masterFile.getAs(MimeType.PDF);
    const pdfName = masterFileName + '.pdf'; // Name set above
    pdfBlob.setName(pdfName);
    
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // FIX: Share the Master Doc so the export link works for everyone
    try { masterFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}

    // Export Master DOCX (Direct Link Strategy)
    // We provide the direct Google Drive export link.
    // IMPORTANT: We must NOT trash the file immediately, otherwise the link breaks.
    const docxDownloadUrl = `https://docs.google.com/document/d/${masterFile.getId()}/export?format=docx`;
    
    // Trash master doc after export? 
    // NO: We keep it so the user can download the DOCX.
    // masterFile.setTrashed(true); 
    
    return { 
        success: true, 
        files: [{
            name: pdfName,
            url: pdfFile.getUrl(),
            downloadUrl: pdfFile.getDownloadUrl(),
            docxUrl: docxDownloadUrl
        }]
    };
    
    // -- LEGACY CODE BELOW REMOVED BY REFACTOR --
    /*

    // --- MERGE STRATEGY: Master Doc is the FIRST Unit's Doc ---
    // We will append subsequent units to this Master Doc
    
    const masterUnitName = unitEntries[0][0];
    const templateFile = DriveApp.getFileById(templateId);
    
    // Create Master Doc
    const masterFile = templateFile.makeCopy('Surat_Pengantar_Gabungan_' + nomorSurat);
    const masterDocId = masterFile.getId();
    const masterDoc = DocumentApp.openById(masterDocId);
    let masterBody = masterDoc.getBody();
    
    // Helper to process a Body (Replace text + Format Table)
    const processBodyForUnit = (bodyToProcess, unitName, unitStudents) => {
       const perihal = (idxPerihal !== -1) ? unitStudents[0][idxPerihal] : 'Permohonan Magang';
       
       bodyToProcess.replaceText('«Nomor»', nomorSurat);
       bodyToProcess.replaceText('«Perihal»', perihal);
       // Unit Name Replacements
       bodyToProcess.replaceText('«Tembusan»', unitName); 
       bodyToProcess.replaceText('«Unit»', unitName);
       bodyToProcess.replaceText('«Periode»', inputPeriode);
       bodyToProcess.replaceText('«Tgl_Dibuat_Surat»', tglSurat);
       
       const idxUniv = headers.findIndex(h => h.toString().toUpperCase() === 'UNIVERSITAS');
       const sekolah = (idxUniv !== -1) ? unitStudents[0][idxUniv] : '';
       bodyToProcess.replaceText('«Sekolah»', sekolah);
       
       const program = (perihal.toLowerCase().includes('penelitian')) ? 'Penelitian' : 'Praktek Kerja Lapangan';
       bodyToProcess.replaceText('«Program»', program);
       
       // Handle Table
       const tables = bodyToProcess.getTables();
       for (let t = 0; t < tables.length; t++) {
         const table = tables[t];
         if (table.getText().includes('«Nama_Satu»') || table.getText().includes('Nama')) {
           const numRows = table.getNumRows();
           let templateRow = null;

           // Clone Row 1
           if (numRows > 1) {
              templateRow = table.getRow(1).copy(); 
           }
           // Remove existing rows
           for (let r = numRows - 1; r > 0; r--) {
              table.removeRow(r);
           }
           if (!templateRow) templateRow = table.appendTableRow().copy();

           // Add Students
           unitStudents.forEach((studentRow) => {
              const row = table.appendTableRow(templateRow.copy());
              row.setMinimumHeight(0); // Force tight layout

              const updateCell = (colIndex, text, alignType) => {
                  if (colIndex < row.getNumCells()) {
                      const cell = row.getCell(colIndex);
                      cell.setText(text || '-');
                      if (cell.getNumChildren() > 0) {
                          const p = cell.getChild(0).asParagraph();
                          p.setSpacingAfter(0);
                          p.setSpacingBefore(0);
                          p.setLineSpacing(1);
                          p.setIndentStart(0);
                          p.setIndentEnd(0);
                          p.setIndentFirstLine(0);
                          p.setAlignment(alignType);
                          const textElement = p.editAsText();
                          if (textElement) {
                              textElement.setFontFamily('Arial');
                              textElement.setFontSize(9);
                              textElement.setBold(false);
                              textElement.setItalic(false);
                              textElement.setUnderline(false);
                          }
                      }
                      cell.setPaddingTop(1);
                      cell.setPaddingBottom(0);
                  }
              };
              updateCell(0, studentRow[idxNama], DocumentApp.HorizontalAlignment.LEFT);
              updateCell(1, studentRow[idxStambuk], DocumentApp.HorizontalAlignment.CENTER);
              updateCell(2, studentRow[idxJurusan], DocumentApp.HorizontalAlignment.CENTER);
           });
         }
       }
    };

    // 1. Process First Unit (Master Doc)
    processBodyForUnit(masterBody, unitEntries[0][0], unitEntries[0][1]);

    // 2. Process Subsequent Units (Append to Master)
    for (let i = 1; i < unitEntries.length; i++) {
        const [uName, uStudents] = unitEntries[i];
        
        // Append Page Break
        masterBody.appendPageBreak();
        
        // Create Temp Doc for this unit to get fresh layout
        const tempFile = templateFile.makeCopy('TEMP_MERGE_' + i);
        const tempDoc = DocumentApp.openById(tempFile.getId());
        const tempBody = tempDoc.getBody();
        
        // Process the Temp Doc Content FIRST
        processBodyForUnit(tempBody, uName, uStudents);
        
        // Copy processed elements to Master
        const totalChildren = tempBody.getNumChildren();
        for (let j = 0; j < totalChildren; j++) {
            const child = tempBody.getChild(j);
            const type = child.getType();
            
            // Append based on type to preserve formatting
            if (type === DocumentApp.ElementType.PARAGRAPH) {
                masterBody.appendParagraph(child.copy());
            } else if (type === DocumentApp.ElementType.TABLE) {
                masterBody.appendTable(child.copy());
            } else if (type === DocumentApp.ElementType.LIST_ITEM) {
                masterBody.appendListItem(child.copy());
            } else {
                // Fallback for other types
                try {
                     // Some elements cannot be appended directly to body like inline images without wrapper
                     // But copy() usually duplicates the element structure.
                     // Helper: generic append
                     // We check specific append methods
                } catch(e) {}
            }
        }
        
        // Trash temp
        tempFile.setTrashed(true);
    }
    
    masterDoc.saveAndClose();
    
    // Export Master PDF
    const pdfBlob = masterFile.getAs(MimeType.PDF);
    const mergedName = `Surat_Pengantar_Gabungan_${nomorSurat}.pdf`;
    pdfBlob.setName(mergedName);
    
    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Trash master doc after export
    masterFile.setTrashed(true);
    
    return { 
        success: true, 
        files: [{
            name: mergedName,
            url: pdfFile.getUrl(),
            downloadUrl: pdfFile.getDownloadUrl()
        }]
    */
  } catch (error) {
    Logger.log('Error generating surat: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}

// Helper to pre-fill modal
function getSuratPengantarDefaults(nomorSurat) {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Data Magang');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idxNomor = headers.findIndex(h => h.toString().toUpperCase() === 'NOMOR SURAT');
    
    if (idxNomor === -1) return { success: false };
    
    // Find first row
    const row = data.slice(1).find(r => r[idxNomor] === nomorSurat);
    if (!row) return { success: false };
    
    const idxPerihal = headers.findIndex(h => h.toString().toUpperCase() === 'PERIHAL');
    
    return { 
      success: true, 
      perihal: (idxPerihal !== -1) ? row[idxPerihal] : '',
    };
}

// New Helper for Preview
function getSuratPengantarPreviewData(sheetName, nomorSurat) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName || 'Data Magang');
    if (!sheet) return { success: false, message: 'Sheet not found' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idxNomor = headers.findIndex(h => h.toString().toUpperCase() === 'NOMOR SURAT');
    const idxNama = headers.findIndex(h => h.toString().toUpperCase().includes('NAMA'));
    const idxStambuk = headers.findIndex(h => h.toString().toUpperCase().includes('STAMBUK') || h.toString().toUpperCase().includes('NIM'));
    const idxUnit = headers.findIndex(h => h.toString().toUpperCase() === 'UNIT KERJA');
    const idxPerihal = headers.findIndex(h => h.toString().toUpperCase() === 'PERIHAL');
    const idxTglMasuk = headers.findIndex(h => h.toString().toUpperCase() === 'TGL MASUK' || h.toString().toUpperCase() === 'TANGGAL MASUK');
    const idxTglSelesai = headers.findIndex(h => h.toString().toUpperCase() === 'TGL SELESAI' || h.toString().toUpperCase() === 'TANGGAL SELESAI');
    
    if (idxNomor === -1) return { success: false, message: 'Column Nomor Surat not found' };
    
    // Filter rows
    const students = data.slice(1).filter(row => row[idxNomor] === nomorSurat).map(row => ({
      nama: row[idxNama] || '-',
      stambuk: row[idxStambuk] || '-',
      unit: row[idxUnit] || '-',
    }));
    
    // Get Perihal & Periode from first matching row
    const firstRow = data.slice(1).find(row => row[idxNomor] === nomorSurat);
    const perihal = (firstRow && idxPerihal !== -1 && firstRow[idxPerihal] != null) ? String(firstRow[idxPerihal]) : '';
    
    let periode = '';
    if (firstRow) {
        // Helper to format date (Force Jakarta Timezone)
        const formatDateIndo = (d) => {
            if (!d || !(d instanceof Date)) return '';
            const day = Utilities.formatDate(d, 'Asia/Jakarta', 'd');
            const month = Utilities.formatDate(d, 'Asia/Jakarta', 'M');
            const year = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy');
            const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
            return `${day} ${months[parseInt(month) - 1]} ${year}`;
        };
        
        const tglMasuk = (idxTglMasuk !== -1) ? firstRow[idxTglMasuk] : null;
        const tglSelesai = (idxTglSelesai !== -1) ? firstRow[idxTglSelesai] : null;
        
        if (tglMasuk instanceof Date && tglSelesai instanceof Date) {
            periode = `${formatDateIndo(tglMasuk)} s.d ${formatDateIndo(tglSelesai)}`;
        } else if (tglMasuk instanceof Date) {
            periode = `Mulai ${formatDateIndo(tglMasuk)}`;
        }
    }
    
    return {
      success: true,
      students: students,
      perihal: perihal,
      periode: periode
    };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ===== SURAT PENOLAKAN GENERATION =====
function generateSuratPenolakanPDF(nomorSurat, headerData, sheetName) {
  try {
    Logger.log('=== GENERATE SURAT PENOLAKAN PDF ===');
    
    // 1. Get data
    const dataResult = getDataByNomorSurat(nomorSurat, sheetName);
    if (!dataResult.success) return dataResult;
    
    const { mahasiswa, common } = dataResult.data;
    
    // 3. Access Template File
    let templateFile;
    try {
      templateFile = DriveApp.getFileById(TEMPLATE_SURAT_PENOLAKAN_ID);
    } catch (e) {
      return { success: false, message: 'Gagal mengakses template Surat Penolakan. Pastikan ID Template benar. Error: ' + e.toString() };
    }
    
    // 4. COPY Template
    const firstStudentName = (mahasiswa && mahasiswa.length > 0 && mahasiswa[0].nama) ? mahasiswa[0].nama : 'Mahasiswa';
    const suffix = (mahasiswa && mahasiswa.length > 1) ? ', DKK' : '';
    const templateName = 'Surat Penolakan - ' + firstStudentName + suffix;
    
    let copiedDoc;
    let docId;
    try {
      copiedDoc = templateFile.makeCopy(templateName);
      docId = copiedDoc.getId();
      try { copiedDoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT); } catch (e) {}
    } catch (e) {
      return { success: false, message: 'Gagal menyalin template. Error: ' + e.toString() };
    }
    
    // 5. Open Doc
    let doc = DocumentApp.openById(docId);
    let body = doc.getBody();
    
    // 6. Replace Text
    // Header Data
    body.replaceText('«Nomor»', headerData.nomorSurat || '');
    body.replaceText('«Lampiran»', headerData.lampiran || '-');
    body.replaceText('«Perihal»', headerData.perihal || '');
    body.replaceText('«Kepada»', headerData.kepada || '');
    body.replaceText('«Jabatan»', headerData.jabatan || '');
    body.replaceText('«Di_Tempat»', headerData.di_tempat || ''); // Hidden field but might be used
    body.replaceText('«Tembusan»', headerData.tembusan || ''); // New Tembusan field
    
    // Common Data from Sheet
    body.replaceText('«Universitas»', common.universitas || '');
    body.replaceText('«Sekolah»', common.universitas || ''); // Map Sekolah to Universitas
    body.replaceText('«Jurusan»', common.jurusan || '');
    body.replaceText('«Surat»', nomorSurat || ''); // Map Surat to Incoming Reference Number (nomorSurat argument)
    
    // Date formatting (Indonesia)
    const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const today = new Date();
    const tglSuratStr = today.getDate() + ' ' + months[today.getMonth()] + ' ' + today.getFullYear();
    body.replaceText('«Tanggal»', tglSuratStr);
    body.replaceText('«Tgl_Dibuat_Surat»', tglSuratStr); // Fix for User's template specific key
    
    // Tembusan
    // If input is empty, ensure it doesn't leave floating "1." if the template handles numbering,
    // but usually template has "1. <<Tembusan>>". If empty, we might want to put "-" or empty.
    body.replaceText('«Tembusan»', headerData.tembusan || '-');

    // Handle Student Table (Identical logic to Balasan)
    const tables = body.getTables();
    let studentTable = null;
    
    // Find table with placeholders
    for (let t = 0; t < tables.length; t++) {
        if (tables[t].getText().includes('«Nama_Satu»') || tables[t].getText().includes('Nama')) {
            studentTable = tables[t];
            break;
        }
    }
    
    if (studentTable) {
        // ... (Compact logic mirroring generateSuratBalasanPDF) ...
        const startRowIndex = 1; // Assuming header is row 0
        const totalRows = studentTable.getNumRows();
        const count = mahasiswa.length;
        
        // Remove unused rows (reverse order)
        if (totalRows > count + 1) {
            for (let r = totalRows - 1; r >= count + 1; r--) {
                studentTable.removeRow(r);
            }
        }
    }
    
    // Fill data
    const namaFields = ['Nama_Satu', 'Nama_dua', 'Nama_tiga', 'Nama_Empat', 'Nama_Lima', 'Nama_Enam', 'Nama_Tujuh', 'Nama_Delapan'];
    const nisFields = ['NIS_Satu', 'NIS_dua', 'NIS_tiga', 'NIS_Empat', 'NIS_Lima', 'NIS_Enam', 'Nis_Tujuh', 'NIS_Delapan'];
    
    for (var i = 0; i < mahasiswa.length; i++) {
        body.replaceText('«' + namaFields[i] + '»', mahasiswa[i].nama || '');
        body.replaceText('«' + nisFields[i] + '»', mahasiswa[i].nis || '');
        body.replaceText('«Nis_' + getNameSuffix(i) + '»', mahasiswa[i].nis || '');
        // Specific weird placeholder fix
        body.replaceText('«Sheet1Jurusan_»', mahasiswa[i].jurusan || common.jurusan || '');
        // Also support cleaner placeholder just in case
        body.replaceText('«Jurusan_Satu»', mahasiswa[i].jurusan || '');
    }
    
    doc.saveAndClose();
    
    // 7. Export PDF
    Utilities.sleep(1000);
    const pdfBlob = copiedDoc.getAs('application/pdf');
    pdfBlob.setName(templateName + '.pdf');
    const pdfFile = DriveApp.createFile(pdfBlob);
    
    // Set Sharing for PDF (CRITICAL FIX)
    try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT); } catch (e) {}

    // Cleanup
    // DISABLED for DOCX Download support - File must remain for link to work
    // try { copiedDoc.setTrashed(true); } catch(e) {}

    const docxUrl = 'https://docs.google.com/document/d/' + docId + '/export?format=docx';
    
    return {
        success: true,
        message: 'Berhasil Generate Surat Penolakan!',
        pdfUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId(),
        previewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/preview',
        fileName: templateName + '.pdf',
        docxUrl: docxUrl
    };
    
  } catch (error) {
    Logger.log('Fatal Error: ' + error.toString());
    return { success: false, message: 'Fatal Error: ' + error.toString() };
  }
}

// ===== GENERATE SURAT SELESAI =====
function generateSuratSelesai(token, rowIndex, nomorSurat, sheetName) {
  try {
    const ss = getSpreadsheet();
    const activeSheetName = sheetName || 'Data Magang';
    const sheet = ss.getSheetByName(activeSheetName);
    if (!sheet) return { success: false, message: 'Sheet ' + activeSheetName + ' not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Direct row access via exact index mapping
    const internRow = data.slice(1)[rowIndex];
    
    if (!internRow) return { success: false, message: 'Data Magang tidak ditemukan pada baris tersebut.' };

    // Helper to get value by column name (fuzzy search)
    const getVal = (colName) => {
        const idx = headers.findIndex(h => h && h.toString().toUpperCase().includes(colName.toUpperCase()));
        return idx !== -1 ? internRow[idx] : '';
    };

    const info = {
       nama: getVal('NAMA'),
       perguruan: getVal('INSTANSI') || getVal('PERGURUAN') || getVal('UNIVERSITAS'),
       jurusan: getVal('JURUSAN'),
       unit: getVal('UNIT'),
       mulai: getVal('MULAI') || getVal('MASUK'),
       selesai: getVal('SELESAI') || getVal('BERAKHIR')
    };
    
    // Format Dates
    const formatDate = (d) => {
        if (!d) return "";
        if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMMM yyyy");
        return d;
    };

    // Helper for Indonesian Month names manual override (Utilities.formatDate uses English by default often)
    const formatIndoDate = (d) => {
        if (!d || !(d instanceof Date)) return formatDate(d);
        const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
        const day = d.getDate();
        const month = months[d.getMonth()];
        const year = d.getFullYear();
        return `${day} ${month} ${year}`;
    };
    
    const mulaiStr = formatIndoDate(info.mulai);
    const selesaiStr = formatIndoDate(info.selesai);
    const todayStr = formatIndoDate(new Date());

    // Create Doc
    const templateId = TEMPLATE_SURAT_SELESAI_ID;
    const docName = `Surat Keterangan Selesai - ${info.nama}`;
    const templateFile = DriveApp.getFileById(templateId);
    const folder = templateFile.getParents().next();
    const copy = templateFile.makeCopy(docName, folder);
    
    // [FIX] Ensure the file is accessible for preview
    copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const docId = copy.getId();
    
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Replace Placeholders
    body.replaceText('«Nama»', info.nama || '');
    body.replaceText('«Perguruan»', info.perguruan || '');
    body.replaceText('«Jurusan»', info.jurusan || '');
    body.replaceText('«Unit»', info.unit || '');
    body.replaceText('«Priode»', `${mulaiStr} s.d ${selesaiStr}`);
    
    // Manual Input & Auto Date
    body.replaceText('«nomor»', nomorSurat || `001/M/${new Date().getFullYear()}`);
    body.replaceText('«Tgl_Surat_di_Buat»', todayStr);
    
    body.replaceText('«Prihal»', 'Kerja Praktik');

    doc.saveAndClose();

    // Return ID and URL for Preview
    return {
        success: true, 
        docId: docId, 
        docUrl: copy.getUrl(),
        fileName: docName
    };

  } catch (e) {
      return { success: false, message: e.toString() };
  }
}

// ===== BATCH METADATA & STATS =====
function getBatchMetadata(batchName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(batchName);
    
    if (!sheet) {
      return { success: false, message: 'Batch not found' };
    }
    
    // Get Data
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
       return { 
           success: true, 
           data: { 
               name: batchName, 
               stats: { total: 0, active: 0, finished: 0 },
               status: 'Empty',
               date_range: '-'
           } 
       };
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().toUpperCase().trim());
    
    // Identify Columns
    const idxKet = headers.indexOf('KET');
    const idxTglMasuk = headers.findIndex(h => h === 'TGL MASUK' || h === 'START DATE');
    const idxTglSelesai = headers.findIndex(h => h === 'TGL SELESAI' || h === 'END DATE');
    
    let total = 0;
    let active = 0;
    let finished = 0;
    
    let minDate = null;
    let maxDate = null;
    
    const today = new Date();
    today.setHours(0,0,0,0);
    
    data.forEach(row => {
        // Skip empty rows (check if first col is empty based on typical pattern)
        if (!row[0]) return;
        
        total++;
        
        const status = (idxKet !== -1 && row[idxKet]) ? row[idxKet].toString().trim() : '';
        const endDate = (idxTglSelesai !== -1) ? row[idxTglSelesai] : null;
        const startDate = (idxTglMasuk !== -1) ? row[idxTglMasuk] : null;
        
        // Date Range Calculation
        if (startDate instanceof Date) {
            if (!minDate || startDate < minDate) minDate = startDate;
        }
        if (endDate instanceof Date) {
            if (!maxDate || endDate > maxDate) maxDate = endDate;
        }
        
        // Stats Logic
        let isExpired = false;
        if (endDate instanceof Date) {
            // Normalize end date
            const d = new Date(endDate);
            d.setHours(0,0,0,0);
            if (d < today) isExpired = true;
        }
        
        // Filter Logic
        if (status === 'DITOLAK' || status === 'Belum Dibuatkan SB') {
            // Do not count as Active OR Finished (effectively ignored or just Total)
        } else if (status.toLowerCase().includes('selesai') || status.toLowerCase().includes('finished') || status.toLowerCase().includes('resign')) {
            finished++;
        } else if (status === 'Sudah Diteruskan') {
            if (isExpired) {
                finished++; // Auto-move to finished if expired
            } else {
                active++;
            }
        } else {
            // Pending or Unknown status - Count as active if not rejected
             active++;
        }
    });
    
    // Format Date Range
    const formatDate = (d) => {
        if (!d) return '';
        const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
        return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
    };
    
    const dateRange = (minDate && maxDate) ? `${formatDate(minDate)} - ${formatDate(maxDate)}` : '-';
    
    // Overall Status
    let batchStatus = 'Active';
    if (active === 0 && finished > 0) batchStatus = 'Finished';
    if (total === 0) batchStatus = 'Empty';
    
    return {
        success: true,
        data: {
            name: batchName,
            stats: {
                total: total,
                active: active,
                finished: finished
            },
            status: batchStatus,
            date_range: dateRange
        }
    };
    
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== SUPER ADMIN USER MANAGEMENT =====
function getAllUsersData(adminEmail) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    if (!sheet) return { success: false, message: 'Database login tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const nameIdx = headers.indexOf('user');
    const roleIdx = headers.indexOf('role');
    const statusIdx = headers.indexOf('status');
    const passIdx = headers.indexOf('password');
    
    let hasAccess = false;
    for(let r=1; r<data.length; r++) {
       const role = data[r][roleIdx];
       if(data[r][emailIdx] === adminEmail && (role === 'Super Admin' || role === 'Developer')) {
           hasAccess = true; break;
       }
    }
    if(!hasAccess) return { success: false, message: 'Akses Ditolak: Hanya Super Admin atau Developer.' };
    
    let users = [];
    for(let r=1; r<data.length; r++) {
       if(!data[r][emailIdx]) continue;
       users.push({
           email: data[r][emailIdx],
           name: data[r][nameIdx],
           role: data[r][roleIdx] || 'User',
           status: data[r][statusIdx] || 'Active',
           password: decodePassword(data[r][passIdx])
       });
    }
    return { success: true, users: users };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function updateUserRoleBySuperAdmin(adminEmail, targetEmail, newRole) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    if (!sheet) return { success: false, message: 'Data tidak ditemukan' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const roleIdx = headers.indexOf('role');
    
    let hasAccess = false;
    for(let r=1; r<data.length; r++) {
       const role = data[r][roleIdx];
       if(data[r][emailIdx] === adminEmail && (role === 'Super Admin' || role === 'Developer')) {
           hasAccess = true; break;
       }
    }
    if(!hasAccess) return { success: false, message: 'Akses Ditolak: Hanya Super Admin atau Developer.' };
    
    for(let r=1; r<data.length; r++) {
       if(data[r][emailIdx] === targetEmail) {
          sheet.getRange(r+1, roleIdx+1).setValue(newRole);
          return { success: true, message: 'Role berhasil diperbarui' };
       }
    }
    return { success: false, message: 'Target user tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function deleteUserAccountBySuperAdmin(adminEmail, targetEmail) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const roleIdx = headers.indexOf('role');
    
    let hasAccess = false;
    for(let r=1; r<data.length; r++) {
       const role = data[r][roleIdx];
       if(data[r][emailIdx] === adminEmail && (role === 'Super Admin' || role === 'Developer')) {
           hasAccess = true; break;
       }
    }
    if(!hasAccess) return { success: false, message: 'Akses Ditolak: Hanya Super Admin atau Developer.' };
    if(targetEmail === adminEmail) return { success: false, message: 'Tidak dapat menghapus akun sendiri.' };
    
    for(let r=1; r<data.length; r++) {
       if(data[r][emailIdx] === targetEmail) {
          sheet.deleteRow(r+1);
          return { success: true, message: 'Akun berhasil dihapus' };
       }
    }
    return { success: false, message: 'Target user tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function resetUserPasswordBySuperAdmin(adminEmail, targetEmail, newPassword) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const roleIdx = headers.indexOf('role');
    const passIdx = headers.indexOf('password');
    
    let hasAccess = false;
    for(let r=1; r<data.length; r++) {
       const role = data[r][roleIdx];
       if(data[r][emailIdx] === adminEmail && (role === 'Super Admin' || role === 'Developer')) {
           hasAccess = true; break;
       }
    }
    if(!hasAccess) return { success: false, message: 'Akses Ditolak: Hanya Super Admin atau Developer.' };
    
    for(let r=1; r<data.length; r++) {
       if(data[r][emailIdx] === targetEmail) {
          // Menyimpan password yang tersandikan (encoded) agar aman di sheet namun bisa dibaca Dashboard
          sheet.getRange(r+1, passIdx+1).setValue(encodePassword(newPassword));
          return { success: true, message: 'Password berhasil direset' };
       }
    }
    return { success: false, message: 'Target user tidak ditemukan' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function upgradeWildanToSuperAdmin() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Login/register');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const nameIdx = headers.indexOf('user');
    const roleIdx = headers.indexOf('role');
    
    let successCount = 0;
    for(let r=1; r<data.length; r++) {
       if(data[r][nameIdx] && data[r][nameIdx].toString().toLowerCase().includes('wildan')) {
          sheet.getRange(r+1, roleIdx+1).setValue('Super Admin');
          successCount++;
       }
    }
    return { success: successCount > 0, message: successCount > 0 ? 'Akun Wildan berhasil diupgrade ke Super Admin.' : 'Gagal menemukan akun Wildan.' };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ===== GLOBAL ANNOUNCEMENT FUNCTIONS (REMOVED) =====

function getGlobalAnnouncement(userEmail) {
  try {
    const props = PropertiesService.getScriptProperties();
    const jokeTarget = props.getProperty('joke_target');
    const jokeType = props.getProperty('joke_type');
    
    let result = {
      announcements: [],
      joke_target: jokeTarget || null,
      joke_type: jokeType || 'evasive',
      forceReload: false,
      maintenance: getMaintenanceStatus()
    };

    // Check for Force Reload
    const reloadAllTime = props.getProperty('FORCE_RELOAD_ALL');
    const reloadUserTime = props.getProperty('FORCE_RELOAD_' + userEmail);
    const lastReloadCheck = PropertiesService.getUserProperties().getProperty('LAST_RELOAD_CHECK');

    if (reloadAllTime || reloadUserTime) {
      const targetTime = Math.max(
        parseInt(reloadAllTime || '0'), 
        parseInt(reloadUserTime || '0')
      );
      
      if (!lastReloadCheck || parseInt(lastReloadCheck) < targetTime) {
        result.forceReload = true;
        PropertiesService.getUserProperties().setProperty('LAST_RELOAD_CHECK', targetTime.toString());
      }
    }

    return result;
  } catch (e) {
    Logger.log("Error in getGlobalAnnouncement: " + e.toString());
    return null;
  }
}

function setJokeTarget(email, type) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('joke_target', email);
    props.setProperty('joke_type', type || 'evasive');
    return { success: true, message: 'Target joke (' + type + ') berhasil diatur ke: ' + email };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function clearJokeTarget() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('joke_target');
    props.deleteProperty('joke_type');
    return { success: true, message: 'Fitur joke dinonaktifkan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== REMOTE RELOAD FUNCTIONS =====
function forceRemoteReload(email, forAll = false) {
  try {
    const props = PropertiesService.getScriptProperties();
    const timestamp = new Date().getTime().toString();
    if (forAll) {
      props.setProperty('FORCE_RELOAD_ALL', timestamp);
    } else {
      props.setProperty('FORCE_RELOAD_' + email, timestamp);
    }
    return { success: true, message: 'Instruksi reload dikirim ke ' + (forAll ? 'semua user' : email) };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== MAHASISWA DASHBOARD FUNCTIONS =====

function doLoginMahasiswa(nim) {
  try {
    if (!nim || String(nim).trim() === '') return { success: false, message: 'NIM tidak boleh kosong' };
    
    const targetNim = String(nim).trim().toLowerCase();
    const ss = getSpreadsheet();
    
    // Ambil semua sheet yang ada di spreadsheet
    const allSheets = ss.getSheets();
    
    // Tentukan sheet mana saja yang BUKAN data magang (sheet sistem)
    const excludeSheets = ['Login/register', 'Sistem', 'Dashboard', 'Notifikasi', 'Pengaturan'];
    let studentFound = false;
    
    // Scan di setiap sheet magang batch satu per satu
    for (let sIdx = 0; sIdx < allSheets.length; sIdx++) {
      const sheet = allSheets[sIdx];
      const sheetName = sheet.getName();
      
      // Lewati sheet sistem
      if (excludeSheets.includes(sheetName)) continue;
      
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1 || lastCol === 0) continue;
      
      const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      // Cari baris header yang sebenarnya (bisa di row 1, 2, atau 3)
      let headerRowIdx = -1;
      let headers = [];
      
      for (let r = 0; r < Math.min(5, data.length); r++) {
        // Bersihkan whitespace berlebih & line break (enter) di dalam judul kolom
        const rowVals = data[r].map(h => String(h).toUpperCase().replace(/\r?\n|\r/g, ' ').replace(/\s+/g, ' ').trim());
        
        // Strategi ketat menemukan Header Row sejati: 
        // Baris judul kemungkinan berisi banyak cell seperti 'NIM', 'NAMA', 'JURUSAN'
        let matchCount = 0;
        const exactHeaders = ['NIM', 'NO BP', 'STAMBUK', 'NO MAHASISWA', 'NO. MAHASISWA', 'NAMA', 'NAMA MAHASISWA', 'NAMA LENGKAP', 'UNIVERSITAS', 'ASAL SEKOLAH', 'JURUSAN', 'PRODI', 'PROGRAM STUDI', 'UNIT KERJA', 'PEMBIMBING', 'TGL MASUK', 'TANGGAL MASUK'];
        
        rowVals.forEach(h => {
            if (exactHeaders.includes(h)) {
                matchCount++;
            } else if (h.includes('KEHADIRAN') || h.includes('ABSEN')) {
                matchCount++;
            }
        });
        
        // Minimal harus ada 2 kolom yang mirip dengan pakem nama tabel 
        // (contoh: di baris ada sel "NIM" dan "NAMA" terpisah)
        if (matchCount >= 2) {
           headerRowIdx = r;
           headers = rowVals;
           break;
        }
      }
      
      if (headerRowIdx === -1) continue; // Skip jika tidak ditemukan baris judul tabel
      
      // Helper function untuk mencari header dengan lebih akurat (Exact first, lalu Partial)
      const findBestHeader = (headers, exactList, partialList) => {
        let idx = headers.findIndex(h => exactList.includes(h));
        if (idx === -1) {
          idx = headers.findIndex(h => partialList.some(p => h.includes(p)));
        }
        return idx;
      };

      const nimColIdx = findBestHeader(headers, 
         ['NIM', 'NO BP', 'STAMBUK', 'NO MAHASISWA', 'NO. MAHASISWA'], 
         ['NIM', 'STAMBUK', 'BP ']
      );
      if (nimColIdx === -1) continue; // Skip jika sheet ini tidak punya kolom NIM

      const namaColIdx = findBestHeader(headers,
         ['NAMA', 'NAMA MAHASISWA', 'NAMA LENGKAP', 'PESERTA', 'NAMA PESERTA'],
         ['NAMA', 'PESERTA']
      );

      const univColIdx = findBestHeader(headers,
         ['UNIVERSITAS', 'KAMPUS', 'INSTANSI', 'ASAL SEKOLAH', 'ASAL UNIVERSITAS', 'ASAL KAMPUS', 'ASAL INSTANSI', 'NAMA SEKOLAH/KAMPUS', 'ASAL SEKOLAH / UNIVERSITAS'],
         ['UNIV', 'KAMPUS', 'INSTANSI', 'SEKOLAH']
      );

      const jurusanColIdx = findBestHeader(headers,
         ['JURUSAN', 'PROGRAM STUDI', 'PRODI', 'JURUSAN/PRODI', 'JURSAN'],
         ['JURUSAN', 'PRODI', 'STUDI', 'JURSAN']
      );

      const kehadiranColIdx = findBestHeader(headers,
         ['KEHADIRAN', 'STATUS KEHADIRAN', 'ABSEN', 'ABSENSI'],
         ['KEHADIRAN', 'ABSEN']
      );

      const statusColIdx = findBestHeader(headers,
         ['KET', 'STATUS', 'KETERANGAN'],
         ['KET', 'STATUS']
      );

      const tglMasukColIdx = findBestHeader(headers,
         ['TGL MASUK', 'TANGGAL MASUK', 'MULAI', 'TANGGAL MULAI', 'TGL MULAI'],
         ['MASUK', 'MULAI', 'AWAL']
      );

      const tglSelesaiColIdx = findBestHeader(headers,
         ['TGL SELESAI', 'TANGGAL SELESAI', 'SELESAI', 'AKHIR', 'PENARIKAN', 'TANGGAL PENARIKAN', 'TGL. SELESAI MAGANG/KKN'],
         ['SELESAI', 'AKHIR', 'PENARIKAN', 'KELUAR']
      );

      const pembimbingColIdx = findBestHeader(headers,
         ['PEMBIMBING', 'MENTOR', 'PENDAMPING', 'PEMBIMBING LAPANGAN'],
         ['PEMBIMBING', 'MENTOR']
      );

      const unitColIdx = findBestHeader(headers,
         ['UNIT', 'UNIT KERJA', 'DEPARTEMEN', 'PENEMPATAN', 'BAGIAN'],
         ['UNIT', 'DEPARTEMEN', 'PENEMPATAN']
      );
      
      // Look for the student row
      for (let i = headerRowIdx + 1; i < data.length; i++) {
        const rowNumStr = String(data[i][nimColIdx] || '').trim().toLowerCase();
        
        if (rowNumStr === targetNim || rowNumStr.includes(targetNim)) {
          // CEK STATUS DITOLAK ATAU SELESAI
          if (statusColIdx !== -1) {
            const statusStr = String(data[i][statusColIdx] || '').trim().toUpperCase();
            if (statusStr === 'DITOLAK') {
               return { success: false, message: 'Akses ditolak. Status Pengajuan Magang Anda Ditolak.' };
            }
            if (statusStr.includes('SELESAI') || statusStr.includes('LULUS') || statusStr.includes('TAMAT')) {
               return { success: false, message: 'Status magang Anda sudah tidak aktif (Selesai). Anda tidak dapat login lagi.' };
            }
          }

          // CEK MASA SELESAI (berdasarkan tanggal akhir magang)
          if (tglSelesaiColIdx !== -1) {
             let rawTglSelesai = data[i][tglSelesaiColIdx];
             if (rawTglSelesai instanceof Date && !isNaN(rawTglSelesai)) {
                 const today = new Date();
                 today.setHours(0,0,0,0);
                 if (today > rawTglSelesai) {
                     return { success: false, message: 'Masa magang Anda telah berakhir sesuai tanggal selesai. Anda tidak dapat login lagi.' };
                 }
             } else if (rawTglSelesai && typeof rawTglSelesai === 'string') {
                 let parsedDt = new Date(rawTglSelesai.trim());
                 if (!isNaN(parsedDt.getTime())) {
                     const today = new Date();
                     today.setHours(0,0,0,0);
                     if (today > parsedDt) {
                         return { success: false, message: 'Masa magang Anda telah berakhir sesuai tanggal selesai. Anda tidak dapat login lagi.' };
                     }
                 }
             }
          }

          // Format Date helper (Bahasa Indonesia)
          const formatDt = (val) => {
             if (!val) return '-';
             
             const bulanIndo = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
             
             try {
                // Formatting function using string padding and GMT+8 timezone to prevent "date shifting" bug
                const formatIndo = (d) => {
                   let formattedStr = Utilities.formatDate(d, "GMT+8", "d-M-yyyy");
                   let parts = formattedStr.split('-');
                   let date = String(parts[0]).padStart(2, '0');
                   let monthIdx = parseInt(parts[1], 10) - 1;
                   let month = bulanIndo[monthIdx];
                   let year = parts[2];
                   return `${date} ${month} ${year}`;
                };

                // Valid Date Object Process
                if (val instanceof Date && !isNaN(val)) {
                   return formatIndo(val);
                }
                
                // Try parsing string to Date
                let strVal = String(val).trim();
                if (strVal && (strVal.includes('/') || strVal.includes('-'))) {
                   let parsedDate = new Date(strVal);
                   if (!isNaN(parsedDate.getTime())) {
                       return formatIndo(parsedDate);
                   }
                }
                return strVal; // Fallback to raw string
             } catch(e) {
                return String(val);
             }
          };

          // Cek masa aktif mahasiswa (apakah jadwal magang sudah dimulai)
          let hasStarted = true; // Default aktif
          if (tglMasukColIdx !== -1) {
             let rawTglMasuk = data[i][tglMasukColIdx];
             if (rawTglMasuk instanceof Date && !isNaN(rawTglMasuk)) {
                 const today = new Date();
                 today.setHours(0,0,0,0);
                 if (rawTglMasuk > today) {
                     hasStarted = false;
                 }
             } else if (rawTglMasuk && typeof rawTglMasuk === 'string') {
                 // Try parse from string
                 let parsedDt = new Date(rawTglMasuk.trim());
                 if (!isNaN(parsedDt.getTime())) {
                     const today = new Date();
                     today.setHours(0,0,0,0);
                     if (parsedDt > today) {
                         hasStarted = false;
                     }
                 }
             }
          }

          const resMap = {
            nim: String(data[i][nimColIdx] || '-'),
            nama: namaColIdx !== -1 ? String(data[i][namaColIdx] || '-') : '-',
            univ: univColIdx !== -1 ? String(data[i][univColIdx] || '-') : '-',
            jurusan: jurusanColIdx !== -1 ? String(data[i][jurusanColIdx] || '-') : '-',
            kehadiran: kehadiranColIdx !== -1 ? String(data[i][kehadiranColIdx] || '') : '',
            hasStarted: hasStarted,
            isAbsenGlobalActive: getGlobalAbsensiState(),
            tglMasuk: tglMasukColIdx !== -1 ? formatDt(data[i][tglMasukColIdx]) : '-',
            tglSelesai: tglSelesaiColIdx !== -1 ? formatDt(data[i][tglSelesaiColIdx]) : '-',
            pembimbing: pembimbingColIdx !== -1 ? String(data[i][pembimbingColIdx] || '-') : '-',
            unit: unitColIdx !== -1 ? String(data[i][unitColIdx] || '-') : '-',
            sheetName: sheetName,
            rowIndex: i - 1 // Ini selaras dengan rowIndex base data yg dikirim di getMagangData()
          };
          
          return { success: true, data: resMap };
        }
      }
    }
    
    return { success: false, message: 'NIM tidak ditemukan. Pastikan Anda sudah terdaftar di Data Magang.' };
  } catch (e) {
    Logger.log('ERROR in doLoginMahasiswa: ' + e.toString());
    return { success: false, message: 'Terjadi kesalahan sistem: ' + e.toString() };
  }
}

function markAttendancePembekalan(nim, sheetName, rowIndex) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet batch tidak ditemukan' };
    
    // Verifikasi headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let kehadiranColIdx = headers.map(h => String(h).toUpperCase().trim()).indexOf('KEHADIRAN');
    
    // Jika tidak ada kolom KEHADIRAN, kita bisa menolak atau menambahkannya. 
    // Diasumsikan sebelumnya sistem sudah punya kolom KEHADIRAN (Dari getMagangData validation)
    if (kehadiranColIdx === -1) {
      return { success: false, message: 'Sistem absensi pembekalan belum diaktifkan (Kolom Kehadiran tidak ada pada Batch ini)' };
    }
    
    // rowIdx yang didapatkan adalah base array index (0 is data array 1, which is sheet row 2).
    // Jadi sheetRow = rowIndex + 2
    const sheetRow = rowIndex + 2;
    
    // Double check that we are setting for the right NIM
    // Tentu NIM Header bisa di mana saja, tapi this is just a quick safety check
    
    // Set "Hadir"
    sheet.getRange(sheetRow, kehadiranColIdx + 1).setValue('Hadir');
    
    // Clear the specific cache to force reload on Admin side
    invalidateCache(['magang_data_' + sheetName, 'dashboard_stats']);
    
    return { success: true, message: 'Kehadiran berhasil dicatat' };
  } catch(e) {
    Logger.log('ERROR in markAttendancePembekalan: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// --- GLOBAL ABSENSI TOGGLE SETTINGS ---
function getGlobalAbsensiState() {
  try {
    const props = PropertiesService.getScriptProperties();
    // Default is false (button hidden/locked) until Admin turns it on
    const state = props.getProperty('GLOBAL_ABSENSI_ACTIVE');
    return state === 'true'; 
  } catch (e) {
    return false;
  }
}

function setGlobalAbsensiState(isActive) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('GLOBAL_ABSENSI_ACTIVE', isActive ? 'true' : 'false');
    return { success: true, message: 'Status tombol absensi berhasil diperbarui.' };
  } catch (e) {
    return { success: false, message: 'Gagal memperbarui status: ' + e.toString() };
  }
}

// ======================================
// ===== ACCEPTANCE RATE REPORT ========
// ======================================

function getAcceptanceReportData(year) {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const excludedSheets = ['Dashboard', 'Data Karyawan', 'Template', 'TEMPLATE', 'Archive', 'Setting', 'Settings', 'Activity Log', 'Login/register'];
    
    // Inisialisasi struktur data:
    // result[month][jenis] = { accepted: 0, rejected: 0 }
    // month: 0-11
    // jenis: 'Magang', 'PKL', 'Penelitian'
    
    const result = {};
    for (let m = 0; m < 12; m++) {
      result[m] = {
        'Magang': { accepted: 0, rejected: 0 },
        'PKL': { accepted: 0, rejected: 0 },
        'Penelitian': { accepted: 0, rejected: 0 }
      };
    }
    
    // Status mappings
    const acceptedStatuses = ['diterima', 'diteruskan', 'aktif', 'selesai', 'lulus', 'tamat'];
    const rejectedStatuses = ['ditolak', 'penolakan', 'gagal'];
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (excludedSheets.includes(sheetName)) return; // Skip excluded sheets
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const headerMap = headers.map(h => h ? h.toString().trim().toUpperCase() : '');
      
      // Find Columns
      let tglMasukIdx = headerMap.findIndex(h => ['TGL MASUK', 'TANGGAL MASUK', 'START DATE'].includes(h) || (h.includes('MASUK') && !h.includes('SURAT')));
      let jenisIdx = headerMap.findIndex(h => h.includes('JENIS') || ['PROGRAM', 'KATEGORI'].includes(h));
      let statusIdx = headerMap.findIndex(h => ['KET', 'KETERANGAN STATUS', 'STATUS', 'VALIDASI', 'KETERANGAN', 'KET.'].includes(h));
      
      // Strict fallback for Tgl Masuk if not found
      if (tglMasukIdx === -1) {
          tglMasukIdx = headerMap.findIndex(h => h.includes('DATE') || h.includes('TGL'));
      }
      
      if (tglMasukIdx === -1 || jenisIdx === -1 || statusIdx === -1) return; // Cannot process without these 3 columns
      
      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      const data = dataRange.getValues();
      
      data.forEach(row => {
        const rawTgl = row[tglMasukIdx];
        const rawJenis = row[jenisIdx] ? row[jenisIdx].toString().trim().toUpperCase() : '';
        const rawStatus = row[statusIdx] ? row[statusIdx].toString().trim().toLowerCase() : '';
        
        let targetJenis = null;
        if (rawJenis.includes('MAGANG')) targetJenis = 'Magang';
        else if (rawJenis.includes('PKL') || rawJenis.includes('PRAKTEK')) targetJenis = 'PKL';
        else if (rawJenis.includes('PENELITIAN') || rawJenis.includes('RISET')) targetJenis = 'Penelitian';
        
        if (!targetJenis || !rawTgl) return; // Skip invalid rows
        
        let tglDate;
        if (rawTgl instanceof Date) {
          tglDate = rawTgl;
        } else if (typeof rawTgl === 'string') {
          tglDate = new Date(rawTgl);
          if (isNaN(tglDate.getTime())) return;
        } else {
          return;
        }
        
        // Filter by Year if specified
        if (year && year !== 'all') {
          if (tglDate.getFullYear().toString() !== year.toString()) return;
        }
        
        const month = tglDate.getMonth(); // 0-11
        
        let isAccepted = false;
        let isRejected = false;
        
        if (acceptedStatuses.some(s => rawStatus.includes(s))) isAccepted = true;
        else if (rejectedStatuses.some(s => rawStatus.includes(s))) isRejected = true;
        
        if (isAccepted) {
           result[month][targetJenis].accepted++;
        } else if (isRejected) {
           result[month][targetJenis].rejected++;
        }
      });
    });
    
    return { success: true, data: result };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ======================================
// ===== REKAP MAGANG REPORT (NEW) ======
// ======================================
/**
 * Menghasilkan rekapitulasi tawaran (surat masuk) vs diterima
 * per kategori (Magang, PKL, KKN Profesi, Penelitian) per bulan
 * Format output sama dengan sheet "Over Acceptance Rate"
 */
function getRekapMagangReport(year) {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    // Tambahkan 'Data Magang' ke daftar excluded berkaca pada fungsinya sebagai summary
    const excludedSheets = [
      'Dashboard', 'Data Karyawan', 'Data Magang', 'Template', 'TEMPLATE',
      'Archive', 'Setting', 'Settings', 'Activity Log',
      'Login/register', 'Notifikasi'
    ];

    const BULAN = ['Januari','Februari','Maret','April','Mei','Juni',
                   'Juli','Agustus','September','Oktober','November','Desember'];
    const KATEGORI = ['Magang','PKL','KKN Profesi','Penelitian'];

    // Inisialisasi struktur result
    const result = {};
    KATEGORI.forEach(k => {
      result[k] = {};
      for (let m = 0; m < 12; m++) {
        result[k][m] = { tawaran: 0, diterima: 0 };
      }
    });

    // Status mapping yang lebih luas
    const acceptedStatuses = ['diterima', 'diteruskan', 'aktif', 'selesai', 'lulus', 'tamat', 'sudah dibuatkan', 'disetujui', 'acc'];
    const rejectedStatuses  = ['ditolak', 'penolakan', 'gagal', 'batal'];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      // Filter sheet sistem
      if (excludedSheets.includes(sheetName)) return;
      // Filter sheet Lookers (asumsi data archive/looker report)
      if (sheetName.toLowerCase().includes('lookers')) return;
      // Jika sheet disembunyikan, biasanya bukan batch aktif yang ingin direkap (kecuali user minta sebaliknya)
      if (sheet.isSheetHidden()) return;

      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;

      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;

      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const headerMap = headers.map(h => h ? h.toString().trim().toUpperCase() : '');

      // Identifikasi kolom dengan fuzzy matching lebih baik
      let tglIdx = headerMap.findIndex(h =>
        ['TGL MASUK','TANGGAL MASUK','START DATE','TGL. MASUK','TANGGAL MULAI'].includes(h) ||
        (h.includes('MASUK') && !h.includes('SURAT') && !h.includes('NOMOR') && !h.includes('KET'))
      );
      
      if (tglIdx === -1) {
        tglIdx = headerMap.findIndex(h => h === 'TGL' || h === 'TANGGAL' || h === 'DATE' || h.includes('WAKTU MASUK'));
      }

      let jenisIdx = headerMap.findIndex(h =>
        h.includes('JENIS') || h === 'PROGRAM' || h === 'KATEGORI' || h.includes('PENGAJUAN')
      );

      let statusIdx = headerMap.findIndex(h =>
        ['KET','STATUS','KETERANGAN','KETERANGAN STATUS','KET.','VALIDASI','HASIL'].includes(h)
      );

      // Metadata dari nama sheet (Contoh: "Januari 2026")
      let sheetMonth = -1;
      let sheetYear = -1;
      const nameParts = sheetName.split(' ');
      if (nameParts.length >= 2) {
          const mIdx = BULAN.indexOf(nameParts[0]);
          if (mIdx !== -1) sheetMonth = mIdx;
          const y = parseInt(nameParts[nameParts.length - 1]);
          if (!isNaN(y)) sheetYear = y;
      }

      // Ambil data satu kali (Bulk read)
      const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
      const data = range.getValues();

      data.forEach(row => {
        // Tentukan tanggal
        let tglDate = null;
        if (tglIdx !== -1 && row[tglIdx]) {
           tglDate = (row[tglIdx] instanceof Date) ? row[tglIdx] : new Date(row[tglIdx]);
        }
        
        // Fallback ke metadata sheet jika kolom tanggal kosong/tidak ada
        if ((!tglDate || isNaN(tglDate.getTime())) && sheetMonth !== -1 && sheetYear !== -1) {
            tglDate = new Date(sheetYear, sheetMonth, 1);
        }
        
        if (!tglDate || isNaN(tglDate.getTime())) return;

        // Filter tahun
        if (year && year !== 'all') {
          if (tglDate.getFullYear().toString() !== year.toString()) return;
        }

        const month = tglDate.getMonth(); // 0-11

        // Tentukan kategori (Logic yang lebih cerdas)
        let kat = null;
        
        // 1. Cek dari kolom JENIS/KATEGORI jika ditemukan
        if (jenisIdx !== -1 && row[jenisIdx]) {
            const rawJenis = row[jenisIdx].toString().trim().toUpperCase();
            if (rawJenis.includes('KKN')) kat = 'KKN Profesi';
            else if (rawJenis.includes('PKL') || rawJenis.includes('PRAKTEK') || rawJenis.includes('KERJA LAPANGAN') || rawJenis.includes('Praktek Kerja Lapangan')) kat = 'PKL';
            else if (rawJenis.includes('PENELITIAN') || rawJenis.includes('RISET')) kat = 'Penelitian';
            else if (rawJenis.includes('MAGANG')) kat = 'Magang';
        }
        
        // 2. Jika belum ketemu, cari di seluruh sel dalam baris tersebut (Fuzzy Search)
        if (!kat) {
            const rowStr = row.join(' ').toUpperCase();
            if (rowStr.includes('KKN')) kat = 'KKN Profesi';
            else if (rowStr.includes('PKL') || rowStr.includes('PRAKTEK') || rowStr.includes('KERJA LAPANGAN')) kat = 'PKL';
            else if (rowStr.includes('PENELITIAN') || rowStr.includes('RISET')) kat = 'Penelitian';
            else if (rowStr.includes('MAGANG')) kat = 'Magang';
        }
        
        // 3. Jika masih belum ketemu, cek dari nama Sheet
        if (!kat) {
            const sn = sheetName.toUpperCase();
            if (sn.includes('KKN')) kat = 'KKN Profesi';
            else if (sn.includes('PKL')) kat = 'PKL';
            else if (sn.includes('PENELITIAN')) kat = 'Penelitian';
            else if (sn.includes('MAGANG')) kat = 'Magang';
        }
        
        // 4. Fallback terakhir jika benar-benar tidak ada petunjuk
        if (!kat) kat = 'Magang';

        if (!result[kat]) return;


        const rawStatus = statusIdx !== -1 && row[statusIdx] ? row[statusIdx].toString().trim().toLowerCase() : '';

        // Hitung Tawaran (Semua data yang masuk)
        result[kat][month].tawaran++;

        // Hitung Diterima (Accepted)
        const isAccepted = acceptedStatuses.some(status => rawStatus.includes(status));
        if (isAccepted) {
          result[kat][month].diterima++;
        }
      });
    });


    // Kalkulasi Ringkasan dan Grand Total
    const summary = {};
    let grandTotalTawaran = 0;
    let grandTotalDiterima = 0;

    KATEGORI.forEach(k => {
      let catTawaran = 0;
      let catDiterima = 0;
      for (let m = 0; m < 12; m++) {
        catTawaran += result[k][m].tawaran;
        catDiterima += result[k][m].diterima;
      }
      summary[k] = {
        totalTawaran: catTawaran,
        totalDiterima: catDiterima,
        rate: catTawaran > 0 ? Math.round((catDiterima / catTawaran) * 100) : 0
      };
      grandTotalTawaran += catTawaran;
      grandTotalDiterima += catDiterima;
    });

    const grandRate = grandTotalTawaran > 0 ? Math.round((grandTotalDiterima / grandTotalTawaran) * 100) : 0;

    return {
      success: true,
      data: result,
      summary: summary,
      grandTotal: {
        totalTawaran: grandTotalTawaran,
        totalDiterima: grandTotalDiterima,
        rate: grandRate
      },
      bulan: BULAN,
      kategori: KATEGORI,
      year: year || 'all'
    };

  } catch (error) {
    Logger.log('Error in getRekapMagangReport: ' + error.toString());
    return { success: false, message: error.toString() };
  }
}
