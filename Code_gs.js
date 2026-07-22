var GEMINI_KEY = 'AIzaSyCF24GDp93Kn8f0--hmKB-Qvj5h9gHqeRY';
var MESSAGE_READERS = ['\u05e8\u05d5\u05d1\u05d9 \u05d3\u05d1\u05d9\u05e8', '\u05dc\u05d9\u05de\u05d5\u05e8 \u05d7\u05d1\u05e1', '\u05d7\u05d9\u05d4 \u05e9\u05d5\u05d5\u05e8\u05e5'];

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = data.action || '';

  if (action === 'logLogin') {
    var evSheet = getOrCreateSheet(ss, '\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd', ['\u05ea\u05d0\u05e8\u05d9\u05da','\u05de\u05e9\u05ea\u05de\u05e9','\u05e4\u05e2\u05d5\u05dc\u05d4','\u05e9\u05d3\u05d4','\u05e2\u05e8\u05da \u05e7\u05d5\u05d3\u05dd','\u05e2\u05e8\u05da \u05d7\u05d3\u05e9','\u05e1\u05de\u05dc \u05de\u05d5\u05e1\u05d3']);
    evSheet.appendRow([new Date(), data.user, '\u05db\u05e0\u05d9\u05e1\u05d4', '', '', '', '']);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'logEvent') {
    var evSheet = getOrCreateSheet(ss, '\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd', ['\u05ea\u05d0\u05e8\u05d9\u05da','\u05de\u05e9\u05ea\u05de\u05e9','\u05e4\u05e2\u05d5\u05dc\u05d4','\u05e9\u05d3\u05d4','\u05e2\u05e8\u05da \u05e7\u05d5\u05d3\u05dd','\u05e2\u05e8\u05da \u05d7\u05d3\u05e9','\u05e1\u05de\u05dc \u05de\u05d5\u05e1\u05d3']);
    evSheet.appendRow([new Date(), data.user, '\u05e9\u05d9\u05e0\u05d5\u05d9', data.field, data.oldVal, data.newVal, data.code]);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'addNote') {
    var nSheet = getOrCreateSheet(ss, '\u05d4\u05e2\u05e8\u05d5\u05ea', ['\u05e1\u05de\u05dc \u05de\u05d5\u05e1\u05d3','\u05e9\u05dc\u05d1','\u05d4\u05e2\u05e8\u05d4','\u05de\u05e9\u05ea\u05de\u05e9','\u05ea\u05d0\u05e8\u05d9\u05da']);
    nSheet.appendRow([data.code, data.level, data.text, data.user, data.date]);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'addMessage') {
    var msgSheet = getOrCreateSheet(ss, '\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd', ['\u05ea\u05d0\u05e8\u05d9\u05da','\u05de\u05e9\u05ea\u05de\u05e9','\u05e4\u05e2\u05d5\u05dc\u05d4','\u05e9\u05d3\u05d4','\u05e2\u05e8\u05da \u05e7\u05d5\u05d3\u05dd','\u05e2\u05e8\u05da \u05d7\u05d3\u05e9','\u05e1\u05de\u05dc \u05de\u05d5\u05e1\u05d3']);
    msgSheet.appendRow([new Date(), data.user, '\u05d4\u05d5\u05d3\u05e2\u05d4', Utilities.getUuid(), '', data.text, '']);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'deleteMessage') {
    var deleteSheet = ss.getSheetByName('\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd');
    var messageRow = parseInt(data.row, 10);
    if (!deleteSheet || !messageRow || messageRow < 2 || String(deleteSheet.getRange(messageRow, 3).getValue()).trim() !== '\u05d4\u05d5\u05d3\u05e2\u05d4') {
      return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
    }
    if (!data.id || String(deleteSheet.getRange(messageRow, 4).getValue()).trim() !== String(data.id).trim()) {
      return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
    }
    deleteSheet.deleteRow(messageRow);
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'markMessagesRead') {
    var reader = String(data.user || '').trim();
    var messageIds = Array.isArray(data.messageIds) ? data.messageIds : [];
    if (!isMessageReader(reader) || messageIds.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({status:'error'})).setMimeType(ContentService.MimeType.JSON);
    }
    var readsSheet = getOrCreateSheet(ss, '\u05e7\u05e8\u05d9\u05d0\u05d5\u05ea \u05d4\u05d5\u05d3\u05e2\u05d5\u05ea', ['\u05de\u05d6\u05d4\u05d4 \u05d4\u05d5\u05d3\u05e2\u05d4','\u05de\u05e9\u05ea\u05de\u05e9','\u05ea\u05d0\u05e8\u05d9\u05da']);
    var existingReads = readsSheet.getDataRange().getValues();
    var readKeys = {};
    for (var i = 1; i < existingReads.length; i++) {
      readKeys[String(existingReads[i][0] || '') + '|' + String(existingReads[i][1] || '')] = true;
    }
    var newReads = [];
    for (var j = 0; j < messageIds.length; j++) {
      var messageId = String(messageIds[j] || '').trim();
      var readKey = messageId + '|' + reader;
      if (messageId && !readKeys[readKey]) {
        newReads.push([messageId, reader, new Date()]);
        readKeys[readKey] = true;
      }
    }
    if (newReads.length) {
      readsSheet.getRange(readsSheet.getLastRow() + 1, 1, newReads.length, 3).setValues(newReads);
    }
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'registerDevice') {
    var dSheet = getOrCreateSheet(ss, '\u05de\u05db\u05e9\u05d9\u05e8\u05d9\u05dd', ['deviceId','\u05e9\u05dd','\u05d8\u05d5\u05e7\u05df','UA','\u05ea\u05d0\u05e8\u05d9\u05da \u05e8\u05d9\u05e9\u05d5\u05dd','\u05e4\u05e2\u05d9\u05dc']);
    var dRows = dSheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < dRows.length; i++) {
      if (String(dRows[i][0]) === data.deviceId) {
        dSheet.getRange(i+1, 2).setValue(data.name);
        dSheet.getRange(i+1, 4).setValue(data.ua);
        found = true;
        break;
      }
    }
    if (!found) {
      dSheet.appendRow([data.deviceId, data.name, data.token, data.ua, new Date(), '\u05db\u05df']);
    }
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }

  // Save field directly to Sheet1 (main data)
  var fieldColMap = {name_sec:12, phone_sec:13, name_mgr:14, phone_mgr:15, it_name:16, it_phone:17, email:18, address:19, status:21, close_amount:22, status_date:23};
  var col = fieldColMap[data.field];
  if (col) {
    var mainSheet = ss.getSheets()[0];
    var values = mainSheet.getDataRange().getValues();
    var code = String(data.code).trim();
    var level = String(data.level || '').replace(/"/g,'').trim();
    for (var i = 1; i < values.length; i++) {
      var rowCode = String(values[i][1]).trim();
      var rowLevel = String(values[i][4]).replace(/"/g,'').trim();
      if (rowCode === code && (!level || rowLevel === level)) {
        mainSheet.getRange(i+1, col).setValue(data.value || '');
        // Write timestamp to column W when status changes
        if (data.field === 'status' || data.field === 'close_amount') {
          mainSheet.getRange(i+1, 23).setValue(new Date());
        }
        break;
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (e && e.parameter) {

    if (e.parameter.action === 'getCSV') {
      var sheet = ss.getSheets()[0];
      var values = sheet.getDataRange().getValues();
      var csv = values.map(function(row) {
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

    if (e.parameter.action === 'auth') {
      return doAuth(e.parameter.type, e.parameter.value);
    }

    if (e.parameter.action === 'geminiKey') {
      return ContentService.createTextOutput(JSON.stringify({key: 'SERVER_SIDE'})).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === 'geminiSearch') {
      var user = decodeURIComponent(e.parameter.user || '').toLowerCase().trim();
      if (user.indexOf('roobid') === -1 && user.indexOf('roobi') === -1 && user.indexOf('רובי') === -1) {
        return ContentService.createTextOutput(JSON.stringify({result: 'ERROR: Gemini access restricted'})).setMimeType(ContentService.MimeType.JSON);
      }
      var prompt = decodeURIComponent(e.parameter.prompt || '');
      var apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + GEMINI_KEY;
      var payload = JSON.stringify({contents:[{parts:[{text: prompt}]}]});
      try {
        var resp = UrlFetchApp.fetch(apiUrl, {method:'post', contentType:'application/json', payload: payload, muteHttpExceptions: true});
        var json = JSON.parse(resp.getContentText());
        var answer = '';
        if (json.candidates && json.candidates[0] && json.candidates[0].content) {
          answer = json.candidates[0].content.parts[0].text;
        } else if (json.error) {
          answer = 'ERROR: ' + json.error.message;
        }
        return ContentService.createTextOutput(JSON.stringify({result: answer})).setMimeType(ContentService.MimeType.JSON);
      } catch(err) {
        return ContentService.createTextOutput(JSON.stringify({result: 'ERROR: ' + err.message})).setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (e.parameter.action === 'getUserEmail') {
      var userName = decodeURIComponent(e.parameter.user || '');
      var authSheet = ss.getSheetByName('\u05d4\u05e8\u05e9\u05d0\u05d5\u05ea');
      var email = '';
      if (authSheet) {
        var authRows = authSheet.getDataRange().getValues();
        for (var i = 1; i < authRows.length; i++) {
          if (String(authRows[i][2]).trim() === userName) {
            // Email is in the rightmost column (column after פעיל)
            email = String(authRows[i][4] || '').trim();
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({email: email})).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === 'getMessages') {
      var messagesSheet = ss.getSheetByName('\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd');
      var messages = [];
      var messageReads = {};
      var readsSheet = ss.getSheetByName('\u05e7\u05e8\u05d9\u05d0\u05d5\u05ea \u05d4\u05d5\u05d3\u05e2\u05d5\u05ea');
      if (readsSheet) {
        var readRows = readsSheet.getDataRange().getValues();
        for (var r = 1; r < readRows.length; r++) {
          var readMessageId = String(readRows[r][0] || '').trim();
          var readUser = String(readRows[r][1] || '').trim();
          if (readMessageId && isMessageReader(readUser)) {
            if (!messageReads[readMessageId]) messageReads[readMessageId] = [];
            if (messageReads[readMessageId].indexOf(readUser) === -1) messageReads[readMessageId].push(readUser);
          }
        }
      }
      if (messagesSheet) {
        var messageRows = messagesSheet.getDataRange().getValues();
        for (var i = messageRows.length - 1; i >= 1; i--) {
          if (String(messageRows[i][2]).trim() !== '\u05d4\u05d5\u05d3\u05e2\u05d4') continue;
          var messageDate = new Date(messageRows[i][0]);
          var messageId = String(messageRows[i][3] || '');
          if (!messageId) {
            messageId = Utilities.getUuid();
            messagesSheet.getRange(i + 1, 4).setValue(messageId);
          }
          messages.push({
            id: messageId,
            row: i + 1,
            timestamp: isNaN(messageDate.getTime()) ? 0 : messageDate.getTime(),
            readBy: messageReads[messageId] || [],
            date: isNaN(messageDate.getTime()) ? String(messageRows[i][0] || '') : Utilities.formatDate(messageDate, 'Asia/Jerusalem', 'dd.MM.yy HH:mm'),
            user: String(messageRows[i][1] || ''),
            text: String(messageRows[i][5] || '')
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify({messages: messages})).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === 'getReport') {
      var reportUser = decodeURIComponent(e.parameter.user || '');
      var reportDays = parseInt(e.parameter.days) || 7;
      var evSheet = ss.getSheetByName('\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd');
      var events = [];
      if (evSheet) {
        var evRows = evSheet.getDataRange().getValues();
        var cutoff = new Date();
        cutoff.setDate(cutoff.getDate() - reportDays);
        for (var i = 1; i < evRows.length; i++) {
          var evDate = new Date(evRows[i][0]);
          var evUser = String(evRows[i][1]).trim();
          if (evUser === reportUser && evDate >= cutoff) {
            events.push({
              date: Utilities.formatDate(evDate, 'Asia/Jerusalem', 'dd.MM.yy HH:mm'),
              action: String(evRows[i][2] || ''),
              field: String(evRows[i][3] || ''),
              oldVal: String(evRows[i][4] || ''),
              newVal: String(evRows[i][5] || ''),
              code: String(evRows[i][6] || '')
            });
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({events: events})).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === 'getMgrReport') {
      var mgrDays = parseInt(e.parameter.days) || 7;
      var evSheet = ss.getSheetByName('\u05d0\u05e8\u05d5\u05e2\u05d9\u05dd');
      // Fixed salesperson keys
      var counts = {
        '\u05e8\u05d5\u05d1\u05d9': {'\u05e9\u05d9\u05d7\u05d4':0, '\u05d4\u05e6\u05e2\u05d4':0, '\u05d0\u05d1\u05d5\u05d3':0, '\u05e1\u05d2\u05d9\u05e8\u05d4':0},
        '\u05d0\u05d9\u05e8\u05d9\u05ea': {'\u05e9\u05d9\u05d7\u05d4':0, '\u05d4\u05e6\u05e2\u05d4':0, '\u05d0\u05d1\u05d5\u05d3':0, '\u05e1\u05d2\u05d9\u05e8\u05d4':0}
      };
      if (evSheet) {
        var evRows = evSheet.getDataRange().getValues();
        var cutoff = new Date();
        cutoff.setDate(cutoff.getDate() - mgrDays);
        for (var i = 1; i < evRows.length; i++) {
          var evDate = new Date(evRows[i][0]);
          if (evDate < cutoff) continue;
          var evUser = String(evRows[i][1]).trim();
          var field = String(evRows[i][3] || '').trim();
          var newVal = String(evRows[i][5] || '').trim();
          // Map user to salesperson key (partial match)
          var spKey = '';
          if (evUser.indexOf('\u05e8\u05d5\u05d1\u05d9') >= 0) spKey = '\u05e8\u05d5\u05d1\u05d9';
          else if (evUser.indexOf('\u05d0\u05d9\u05e8\u05d9\u05ea') >= 0) spKey = '\u05d0\u05d9\u05e8\u05d9\u05ea';
          if (!spKey) continue;
          // Count status changes (field='סטטוס') and closings (field='סגירה')
          if (field === '\u05e1\u05d8\u05d8\u05d5\u05e1' || field === '\u05e1\u05d2\u05d9\u05e8\u05d4') {
            if (field === '\u05e1\u05d2\u05d9\u05e8\u05d4') {
              counts[spKey]['\u05e1\u05d2\u05d9\u05e8\u05d4']++;
            } else if (counts[spKey][newVal] !== undefined) {
              counts[spKey][newVal]++;
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({counts: counts})).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === 'getContacts') {
      var cSheet = ss.getSheetByName('\u05db\u05ea\u05d5\u05d1\u05d5\u05ea');
      var contacts = [];
      if (cSheet) {
        var cRows = cSheet.getDataRange().getValues();
        for (var i = 1; i < cRows.length; i++) {
          contacts.push({code: String(cRows[i][0]).trim(), desc: String(cRows[i][1]||'').trim(), value: String(cRows[i][2]||'').trim()});
        }
      }
      return ContentService.createTextOutput(JSON.stringify({contacts: contacts})).setMimeType(ContentService.MimeType.JSON);
    }

    if (e.parameter.action === 'deviceAuth') {
      var did = e.parameter.deviceId;
      var dSheet = ss.getSheetByName('\u05de\u05db\u05e9\u05d9\u05e8\u05d9\u05dd');
      if (dSheet) {
        var dRows = dSheet.getDataRange().getValues();
        for (var i = 1; i < dRows.length; i++) {
          if (String(dRows[i][0]) === did && String(dRows[i][5]).trim() === '\u05db\u05df') {
            // Look up permType from auth sheet using token
            var devToken = String(dRows[i][2]||'').replace(/[\s\-()]/g,'').toLowerCase();
            var permType = 'phone';
            var aSheet = ss.getSheetByName('\u05d4\u05e8\u05e9\u05d0\u05d5\u05ea');
            if (aSheet) {
              var aRows = aSheet.getDataRange().getValues();
              for (var j = 1; j < aRows.length; j++) {
                var aid = String(aRows[j][0]).replace(/[\s\-()]/g,'').toLowerCase();
                if (aid === devToken && String(aRows[j][3]).trim() === '\u05db\u05df') {
                  permType = String(aRows[j][1]).trim();
                  break;
                }
              }
            }
            return ContentService.createTextOutput(JSON.stringify({authorized:true, name:dRows[i][1]||'', permType:permType})).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({authorized:false})).setMimeType(ContentService.MimeType.JSON);
    }
  }

  var result = {};
  var nSheet = ss.getSheetByName('\u05d4\u05e2\u05e8\u05d5\u05ea');
  if (nSheet) {
    var nRows = nSheet.getDataRange().getValues();
    for (var i = 1; i < nRows.length; i++) {
      var nk = String(nRows[i][0]) + '_' + String(nRows[i][1]);
      if (!result[nk]) result[nk] = {notes:[], fields:{}};
      result[nk].notes.push({text: nRows[i][2]||'', user: nRows[i][3]||'', date: nRows[i][4]||''});
    }
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doAuth(type, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('\u05d4\u05e8\u05e9\u05d0\u05d5\u05ea');
  if (!sheet) {
    sheet = ss.insertSheet('\u05d4\u05e8\u05e9\u05d0\u05d5\u05ea');
    sheet.appendRow(['\u05de\u05d6\u05d4\u05d4','\u05e1\u05d5\u05d2','\u05e9\u05dd','\u05e4\u05e2\u05d9\u05dc']);
    sheet.appendRow(['ADMIN2026','code','\u05de\u05e0\u05d4\u05dc','\u05db\u05df']);
  }
  var rows = sheet.getDataRange().getValues();
  var cleanVal = String(value).replace(/[\s\-()]/g, '').toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][3]).trim() !== '\u05db\u05df') continue;
    var id = String(rows[i][0]).replace(/[\s\-()]/g, '').toLowerCase();
    var idType = String(rows[i][1]).trim();
    if (idType === type && id === cleanVal) {
      return ContentService.createTextOutput(JSON.stringify({authorized:true, name:rows[i][2]||'', permType:idType})).setMimeType(ContentService.MimeType.JSON);
    }
    // phone1 is a higher-level phone permission
    if (idType === 'phone1' && type === 'phone' && id === cleanVal) {
      return ContentService.createTextOutput(JSON.stringify({authorized:true, name:rows[i][2]||'', permType:'phone1'})).setMimeType(ContentService.MimeType.JSON);
    }
    if (idType === 'domain' && type === 'email' && id.charAt(0) === '@') {
      if (cleanVal.indexOf(id) === cleanVal.length - id.length) {
        return ContentService.createTextOutput(JSON.stringify({authorized:true, name:rows[i][2]||value, permType:idType})).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({authorized:false})).setMimeType(ContentService.MimeType.JSON);
}

function isMessageReader(user) {
  return MESSAGE_READERS.indexOf(String(user || '').trim()) !== -1;
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}
