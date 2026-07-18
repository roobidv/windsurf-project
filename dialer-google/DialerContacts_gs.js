// ============================================================================
// Google Apps Script for Dialer - Contacts & tblGLOBAL_PHONE_BOOK
// Sheet ID: 10BBvLgvzDEMNfXkallqvHrml4MqRK-zsoFaX1eZVKyo
//
// Contacts tab columns (A-K):
//   ContactID, ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Address, Notes, DateAdded, CallCount
//
// tblGLOBAL_PHONE_BOOK tab columns (A-K):
//   ContactID, ContactName, FamlyName, Tital, Landline, PhoneNumber, DateAdded, Email, Notes, Address, CallCount
//
// Deploy URL: https://script.google.com/macros/s/AKfycbzVe67O5Qoxc2Z4o73FsoK6tRUgzHXo-U-hYTxWe3RRofOZJEXJIjtKw5q5IBOpyWIKDQ/exec
// ============================================================================

var SHEET_ID = '10BBvLgvzDEMNfXkallqvHrml4MqRK-zsoFaX1eZVKyo';
var VALID_TABLES = ['Contacts', 'tblGLOBAL_PHONE_BOOK'];
var LOG_SHEET = 'EventLog';

// PhoneNumber column index (0-based) per table
var PHONE_COL = { 'Contacts': 4, 'tblGLOBAL_PHONE_BOOK': 5 };

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(tableName) {
  if (VALID_TABLES.indexOf(tableName) === -1) return null;
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(tableName);
}

// ---- GET: read all rows from a table ----
function doGet(e) {
  var table = (e && e.parameter) ? (e.parameter.table || e.parameter.action || '') : '';
  var fmt = (e && e.parameter) ? (e.parameter.format || '') : '';

  // Support legacy ?action=getContacts
  if (table === 'getContacts') table = 'Contacts';
  if (table === 'getPhoneBook') table = 'tblGLOBAL_PHONE_BOOK';

  var sheet = getSheet(table);
  if (!sheet) {
    if (fmt === 'tsv') return ContentService.createTextOutput('ERROR\tUnknown table').setMimeType(ContentService.MimeType.PLAIN_TEXT);
    return json({status:'error', message:'Unknown table: ' + table});
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  // --- TSV format: tab-separated, no JSON needed ---
  if (fmt === 'tsv') {
    var lines = [];
    lines.push(headers.join('\t'));
    for (var i = 1; i < data.length; i++) {
      var vals = [];
      for (var j = 0; j < headers.length; j++) {
        var v = data[i][j];
        if (v === null || v === undefined) v = '';
        v = String(v).replace(/\t/g, ' ').replace(/\r?\n/g, ' ');
        // Format dates as yyyy-mm-dd
        if (v && data[i][j] instanceof Date) {
          v = Utilities.formatDate(data[i][j], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        vals.push(v);
      }
      lines.push(vals.join('\t'));
    }
    return ContentService.createTextOutput(lines.join('\n')).setMimeType(ContentService.MimeType.PLAIN_TEXT);
  }

  // --- JSON format (default) ---
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j] !== null && data[i][j] !== undefined ? data[i][j] : '';
    }
    rows.push(row);
  }
  return json({status:'ok', table: table, count: rows.length, rows: rows});
}

// ---- POST: add / update / delete ----
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var action = data.action || '';
  var table = data.table || 'Contacts';

  var sheet = getSheet(table);
  if (!sheet) return json({status:'error', message:'Unknown table: ' + table});

  var phoneCol = PHONE_COL[table];  // 0-based index of PhoneNumber column

  // --- add ---
  if (action === 'add') {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Auto-increment ContactID
    var lastRow = sheet.getLastRow();
    var nextId = 1;
    if (lastRow > 1) {
      var lastId = sheet.getRange(lastRow, 1).getValue();
      nextId = (parseInt(lastId) || 0) + 1;
    }
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      if (headers[j] === 'ContactID') {
        row.push(nextId);
      } else {
        row.push(data[headers[j]] !== undefined ? data[headers[j]] : '');
      }
    }
    sheet.appendRow(row);
    return json({status:'ok', ContactID: nextId});
  }

  // --- update: find by PhoneNumber ---
  if (action === 'update') {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var values = sheet.getDataRange().getValues();
    var phone = String(data.PhoneNumber || '').replace(/[-\s]/g, '');
    for (var i = 1; i < values.length; i++) {
      var rowPhone = String(values[i][phoneCol] || '').replace(/[-\s]/g, '');
      if (rowPhone === phone) {
        for (var j = 0; j < headers.length; j++) {
          if (headers[j] === 'ContactID') continue;  // never change ID
          if (data[headers[j]] !== undefined) {
            sheet.getRange(i+1, j+1).setValue(data[headers[j]]);
          }
        }
        return json({status:'ok'});
      }
    }
    return json({status:'not_found'});
  }

  // --- delete: find by PhoneNumber ---
  if (action === 'delete') {
    var values = sheet.getDataRange().getValues();
    var phone = String(data.PhoneNumber || '').replace(/[-\s]/g, '');
    for (var i = 1; i < values.length; i++) {
      var rowPhone = String(values[i][phoneCol] || '').replace(/[-\s]/g, '');
      if (rowPhone === phone) {
        sheet.deleteRow(i+1);
        return json({status:'ok'});
      }
    }
    return json({status:'not_found'});
  }

  // --- log: append event row to EventLog sheet ---
  if (action === 'log') {
    var logSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET);
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SHEET_ID).insertSheet(LOG_SHEET);
      logSheet.appendRow(['Timestamp', 'ComputerName', 'UserName', 'EventType', 'Details']);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([
      ts,
      data.ComputerName || '',
      data.UserName || '',
      data.EventType || '',
      data.Details || ''
    ]);
    return json({status:'ok'});
  }

  return json({status:'error', message:'Unknown action: ' + action});
}
