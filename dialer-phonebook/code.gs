/**
 * Dialer PhoneBook - Google Apps Script
 * ==============================================
 * סקריפט זה משרת את אפליקציית ספר הטלפונים.
 * יש להדביק אותו ב-Google Sheets "Dialer Contacts" → Extensions → Apps Script
 * ולבצע Deploy → Web app (Anyone can access)
 *
 * גליונות נדרשים בקובץ:
 *   - Contacts (נתוני אנשי קשר)
 *   - הרשאות (טלפון/מייל/קוד + שם + הרשאה)
 *   - מכשירים (device_id, name, token, ua, lastSeen, permType)
 *   - ארועים (תאריך, משתמש, פעולה, שדה, ערך קודם, ערך חדש, CUS_NUMBER)
 */

function doGet(e) {
  var action = (e.parameter.action || '').trim();
  
  if (action === 'getCSV') return getCSV_();
  if (action === 'auth') return doAuth_(e.parameter);
  if (action === 'deviceAuth') return doDeviceAuth_(e.parameter);
  
  // Default: return notes/fields (like loadCloudNotes in LOGISTY2026)
  return ContentService.createTextOutput('{}').setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || '';
    
    if (action === 'registerDevice') return registerDevice_(data);
    if (action === 'logLogin') return logLogin_(data);
    if (action === 'logEvent') return logEvent_(data);
    if (action === 'updateContact') return updateContact_(data);
    
    return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
  } catch(ex) {
    return ContentService.createTextOutput('{"error":"' + ex.message + '"}').setMimeType(ContentService.MimeType.JSON);
  }
}

// === getCSV: שליפת גליון Contacts כ-CSV ===
function getCSV_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet) return ContentService.createTextOutput('{"csv":""}').setMimeType(ContentService.MimeType.JSON);
  
  var data = sheet.getDataRange().getValues();
  var csv = data.map(function(row) {
    return row.map(function(cell) {
      var s = String(cell === null || cell === undefined ? '' : cell);
      if (s.indexOf(',') > -1 || s.indexOf('"') > -1 || s.indexOf('\n') > -1) {
        return '"' + s.replace(/"/g, '""') + '"';
      }
      return s;
    }).join(',');
  }).join('\n');
  
  return ContentService.createTextOutput(JSON.stringify({csv: csv})).setMimeType(ContentService.MimeType.JSON);
}

// === auth: אימות טלפון/מייל/קוד ===
function doAuth_(params) {
  var type = (params.type || '').trim();
  var value = (params.value || '').trim();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('הרשאות');
  if (!sheet) return jsonOut_({authorized: false});
  
  var data = sheet.getDataRange().getValues();
  // הרשאות: עמודה A=טלפון, B=מייל, C=קוד, D=שם, E=הרשאה
  for (var i = 1; i < data.length; i++) {
    var phone = String(data[i][0] || '').replace(/[\s\-]/g, '');
    var email = String(data[i][1] || '').trim().toLowerCase();
    var code = String(data[i][2] || '').trim();
    var name = String(data[i][3] || '').trim();
    var perm = String(data[i][4] || '').trim();
    
    var match = false;
    if (type === 'phone' && phone === value.replace(/[\s\-]/g, '')) match = true;
    if (type === 'email' && email === value.toLowerCase()) match = true;
    if (type === 'code' && code === value) match = true;
    
    if (match) {
      return jsonOut_({authorized: true, name: name, permType: perm});
    }
  }
  
  return jsonOut_({authorized: false});
}

// === deviceAuth: אימות לפי מזהה מכשיר ===
function doDeviceAuth_(params) {
  var deviceId = (params.deviceId || '').trim();
  if (!deviceId) return jsonOut_({authorized: false});
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('מכשירים');
  if (!sheet) return jsonOut_({authorized: false});
  
  var data = sheet.getDataRange().getValues();
  // מכשירים: A=deviceId, B=name, C=token, D=ua, E=lastSeen, F=permType
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === deviceId) {
      // Update lastSeen
      sheet.getRange(i + 1, 5).setValue(new Date());
      var name = String(data[i][1] || '');
      var perm = String(data[i][5] || '');
      return jsonOut_({authorized: true, name: name, permType: perm});
    }
  }
  
  return jsonOut_({authorized: false});
}

// === registerDevice: רישום מכשיר חדש ===
function registerDevice_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('מכשירים');
  if (!sheet) return jsonOut_({ok: false});
  
  var deviceId = data.deviceId || '';
  var name = data.name || '';
  var token = data.token || '';
  var ua = data.ua || '';
  var permType = data.permType || '';
  
  // Check if device already exists
  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim() === deviceId) {
      // Update existing
      sheet.getRange(i + 1, 2).setValue(name);
      sheet.getRange(i + 1, 5).setValue(new Date());
      sheet.getRange(i + 1, 6).setValue(permType);
      return jsonOut_({ok: true});
    }
  }
  
  // Add new device
  sheet.appendRow([deviceId, name, token, ua, new Date(), permType]);
  return jsonOut_({ok: true});
}

// === logLogin: רישום כניסה בגליון ארועים ===
function logLogin_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ארועים');
  if (!sheet) return jsonOut_({ok: false});
  
  sheet.appendRow([new Date(), data.user || '', 'כניסה', '', '', '', '']);
  return jsonOut_({ok: true});
}

// === logEvent: רישום אירוע בגליון ארועים ===
function logEvent_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ארועים');
  if (!sheet) return jsonOut_({ok: false});
  
  // ארועים: תאריך, משתמש, פעולה, שדה, ערך קודם, ערך חדש, CUS_NUMBER
  sheet.appendRow([
    new Date(),
    data.user || '',
    data.action_type || data.field ? 'עדכון שדה' : 'פעולה',
    data.field || '',
    data.oldVal || '',
    data.newVal || '',
    data.cus_number || ''
  ]);
  return jsonOut_({ok: true});
}

// === updateContact: עדכון שדה באיש קשר ===
function updateContact_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Contacts');
  if (!sheet) return jsonOut_({ok: false});
  
  var contactId = String(data.contactId || '');
  var field = data.field || '';
  var value = data.value || '';
  
  if (!contactId || !field) return jsonOut_({ok: false, error: 'missing params'});
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIdx = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === field) { colIdx = c; break; }
  }
  if (colIdx === -1) return jsonOut_({ok: false, error: 'field not found'});
  
  // Find row by ContactID (column A)
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]).trim() === contactId) {
      sheet.getRange(i + 1, colIdx + 1).setValue(value);
      return jsonOut_({ok: true});
    }
  }
  
  return jsonOut_({ok: false, error: 'contact not found'});
}

// === Helper ===
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
