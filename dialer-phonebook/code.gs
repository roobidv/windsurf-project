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
 *   - הודעות (ID, Target, MSG, CreatedDate, ReadDate)
 */

function doGet(e) {
  var action = (e.parameter.action || '').trim();

  if (action === 'getCSV') return getCSV_();
  if (action === 'auth') return doAuth_(e.parameter);
  if (action === 'deviceAuth') return doDeviceAuth_(e.parameter);
  if (action === 'checkMessages') return checkMessages_(e.parameter);
  if (action === 'getOfferPdf') return getOfferPdf_(e.parameter);

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
    if (action === 'addNote') return addNote_(data);
    if (action === 'markMessageRead') return markMessageRead_(data);

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

// === auth: אימות לפי קוד כניסה אחיד + הגבלת ניסיונות ===
function doAuth_(params) {
  var type = (params.type || '').trim();
  var value = (params.value || '').trim();

  var ACCESS_CODE = '583995';
  var MAX_ATTEMPTS = 5;
  var WINDOW_MINUTES = 5;
  var LOCKOUT_HOURS = 4;

  var cache = CacheService.getScriptCache();

  // בדוק חסימה
  var lockout = cache.get('auth_lockout');
  if (lockout) {
    return jsonOut_({authorized: false, locked: true, message: 'locked'});
  }

  // כניסה לפי קוד אחיד
  if (type === 'code' && value === ACCESS_CODE) {
    cache.remove('auth_fails');
    return jsonOut_({authorized: true, name: 'משתמש מורשה', permType: 'מנהל'});
  }

  // ניסיון כושל - עדכן מונה
  var failsRaw = cache.get('auth_fails');
  var fails = failsRaw ? JSON.parse(failsRaw) : {count: 0, first: Date.now()};

  // אם עבר חלון הזמן - אפס
  if (Date.now() - fails.first > WINDOW_MINUTES * 60 * 1000) {
    fails = {count: 0, first: Date.now()};
  }

  fails.count++;

  if (fails.count >= MAX_ATTEMPTS) {
    // חסום ל-4 שעות
    cache.put('auth_lockout', 'true', LOCKOUT_HOURS * 3600);
    cache.remove('auth_fails');
    return jsonOut_({authorized: false, locked: true, message: 'locked'});
  }

  cache.put('auth_fails', JSON.stringify(fails), WINDOW_MINUTES * 60);
  var remaining = MAX_ATTEMPTS - fails.count;
  return jsonOut_({authorized: false, remaining: remaining, message: 'wrong_code'});
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

// === addNote: שמירת הערה בגליון ארועים ===
function addNote_(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ארועים');
  if (!sheet) return jsonOut_({ok: false});

  sheet.appendRow([
    new Date(),
    data.user || '',
    'הערה חדשה',
    'הערה',
    '',
    data.text || '',
    data.contactId || ''
  ]);
  return jsonOut_({ok: true});
}

// === checkMessages: בדיקת הודעות למחשב מסוים ===
function checkMessages_(params) {
  var target = (params.target || '').trim();
  if (!target) return jsonOut_({messages: []});

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('הודעות');
  if (!sheet) return jsonOut_({messages: []});

  var data = sheet.getDataRange().getValues();
  // הודעות: A=ID, B=Target, C=MSG, D=CreatedDate, E=ReadDate
  var results = [];
  for (var i = 1; i < data.length; i++) {
    var rowTarget = String(data[i][1] || '').trim().toUpperCase();
    var msg = String(data[i][2] || '').trim();
    var readDate = String(data[i][4] || '').trim();

    // הודעה לא נקראה + מיועדת למחשב הזה או לכולם
    if (msg && !readDate && (rowTarget === target.toUpperCase() || rowTarget === 'ALL')) {
      results.push({id: String(data[i][0] || i), msg: msg, row: i + 1});
    }
  }

  return jsonOut_({messages: results});
}

// === markMessageRead: סימון הודעה כנקראה ===
function markMessageRead_(data) {
  var row = parseInt(data.row);
  if (!row || row < 2) return jsonOut_({ok: false});

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('הודעות');
  if (!sheet) return jsonOut_({ok: false});

  // עמודה E = ReadDate
  sheet.getRange(row, 5).setValue(new Date());
  return jsonOut_({ok: true});
}

// === getOfferPdf: שליפת קישור זמני ל-PDF הצעת מחיר מ-Dropbox ===
function getOfferPdf_(params) {
  var serial = (params.serial || '').trim();
  if (!serial) return jsonOut_({error: 'missing serial number'});

  try {
    var link = getDropboxTempLink_(serial);
    return jsonOut_({link: link});
  } catch(e) {
    return jsonOut_({error: e.message});
  }
}

function getDropboxTempLink_(serialNumber) {
  var props = PropertiesService.getScriptProperties();
  var refreshToken = props.getProperty('DROPBOX_REFRESH_TOKEN');
  var appKey = props.getProperty('DROPBOX_APP_KEY');
  var appSecret = props.getProperty('DROPBOX_APP_SECRET');

  if (!refreshToken || !appKey || !appSecret) {
    throw new Error('Dropbox credentials not configured in Script Properties');
  }

  // Get fresh access token using refresh token
  var tokenResponse = UrlFetchApp.fetch('https://api.dropbox.com/oauth2/token', {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: appKey,
      client_secret: appSecret
    }
  });
  var accessToken = JSON.parse(tokenResponse.getContentText()).access_token;

  // List files in offers folder and find matching serial number
  var listResponse = UrlFetchApp.fetch('https://api.dropboxapi.com/2/files/list_folder', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({path: '/Ness/offers', limit: 100}),
    muteHttpExceptions: true
  });

  var listCode = listResponse.getResponseCode();
  if (listCode !== 200) {
    throw new Error('שגיאה בגישה לתיקיית ההצעות');
  }

  var listResult = JSON.parse(listResponse.getContentText());
  var matchedPath = null;

  for (var i = 0; i < listResult.entries.length; i++) {
    var entry = listResult.entries[i];
    if (entry.name.indexOf('_' + serialNumber + '.pdf') > -1) {
      matchedPath = entry.path_lower;
      break;
    }
  }

  if (!matchedPath) {
    throw new Error('הצעה מספר ' + serialNumber + ' לא נמצאה');
  }

  // Get temporary link using the exact path from listing
  var response = UrlFetchApp.fetch('https://api.dropboxapi.com/2/files/get_temporary_link', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + accessToken,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({path: matchedPath}),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var result = JSON.parse(response.getContentText());

  if (code !== 200) {
    throw new Error(result.error_summary || 'Dropbox API error');
  }

  return result.link;
}

// === Test: run this once to grant UrlFetchApp permissions ===
function testDropboxAuth() {
  var link = getDropboxTempLink_('1900384');
  Logger.log('Link: ' + link);
}

// === Helper ===
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
