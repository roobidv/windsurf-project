var GEMINI_KEY = 'AIzaSyCF24GDp93Kn8f0--hmKB-Qvj5h9gHqeRY';

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
  var fieldColMap = {name_sec:12, phone_sec:13, name_mgr:14, phone_mgr:15, it_name:16, it_phone:17, email:18, address:19};
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

    if (e.parameter.action === 'generateProposal') {
     try {
      var templateId = '1g0qJvwZJ1WaZ1U6rLTUBIBaIUoju2cdIcb8GjR-QB9U';
      var templateFile = DriveApp.getFileById(templateId);
      var p = e.parameter;
      var schoolName = decodeURIComponent(p.schoolName || '');
      var code = decodeURIComponent(p.code || '');
      var authority = decodeURIComponent(p.authority || '');
      var address = decodeURIComponent(p.address || '');
      var managerName = decodeURIComponent(p.managerName || '');
      var managerPhone = decodeURIComponent(p.managerPhone || '');
      var budgetStr = decodeURIComponent(p.budget || '0');

      var copyName = '\u05d4\u05e6\u05e2\u05ea \u05e9\u05d9\u05e8\u05d5\u05ea - ' + schoolName + ' (' + code + ')';
      var copy = templateFile.makeCopy(copyName);
      var doc = DocumentApp.openById(copy.getId());
      var body = doc.getBody();

      body.replaceText('__________________', schoolName || '________');
      body.replaceText('\u05e9\u05dd \u05d4\u05de\u05d5\u05e1\u05d3: ________', '\u05e9\u05dd \u05d4\u05de\u05d5\u05e1\u05d3: ' + schoolName);
      body.replaceText('\u05e1\u05de\u05dc \u05de\u05d5\u05e1\u05d3: __________', '\u05e1\u05de\u05dc \u05de\u05d5\u05e1\u05d3: ' + code);
      body.replaceText('\u05e8\u05e9\u05d5\u05ea: __________', '\u05e8\u05e9\u05d5\u05ea: ' + authority);
      body.replaceText('\u05db\u05ea\u05d5\u05d1\u05ea \u05de\u05d5\u05e1\u05d3: ____________', '\u05db\u05ea\u05d5\u05d1\u05ea \u05de\u05d5\u05e1\u05d3: ' + address);
      body.replaceText('\u05e9\u05dd \u05de\u05e0\u05d4\u05dc/\u05ea \u05d4\u05de\u05d5\u05e1\u05d3: ___________', '\u05e9\u05dd \u05de\u05e0\u05d4\u05dc/\u05ea \u05d4\u05de\u05d5\u05e1\u05d3: ' + managerName);
      body.replaceText("\u05de\u05e1' \u05d8\u05dc' \u05de\u05e0\u05d4\u05dc/\u05ea \u05de\u05d5\u05e1\u05d3: _____________", "\u05de\u05e1' \u05d8\u05dc' \u05de\u05e0\u05d4\u05dc/\u05ea \u05de\u05d5\u05e1\u05d3: " + managerPhone);

      var budget = parseFloat(budgetStr.replace(/[^\d.]/g, '')) || 0;
      var monthly = Math.round(budget / 12);
      var totalVat = Math.round(budget * 1.17);
      body.replaceText('\u05e1\u05da \u05e9\u05dc _______', '\u05e1\u05da \u05e9\u05dc ' + monthly.toLocaleString() + ' \u20aa');
      body.replaceText('\u05d5\u05d1\u05e1\u05d4"\u05db _____\u05db\u05d5\u05dc\u05dc', '\u05d5\u05d1\u05e1\u05d4"\u05db ' + totalVat.toLocaleString() + ' \u20aa \u05db\u05d5\u05dc\u05dc');

      // Lookup user in הרשאות by column C (name)
      // Col A(0)=מזהה(phone), B(1)=סוג, C(2)=שם, D(3)=פעיל, E(4)=מייל
      var userId = decodeURIComponent(p.user || '');
      var authSheet2 = ss.getSheetByName('\u05d4\u05e8\u05e9\u05d0\u05d5\u05ea');
      var sigName = '', sigPhone = '', sigEmail = '';
      if (authSheet2 && userId) {
        var authRows2 = authSheet2.getDataRange().getValues();
        for (var ai = 1; ai < authRows2.length; ai++) {
          if (String(authRows2[ai][2]).trim() === userId) {
            sigName = String(authRows2[ai][2] || '').trim();
            sigPhone = String(authRows2[ai][0] || '').trim();
            sigEmail = String(authRows2[ai][4] || '').trim();
            break;
          }
        }
      }

      // Add signature to document
      if (sigName) {
        body.appendParagraph('\u05d1\u05d1\u05e8\u05db\u05d4,');
        body.appendParagraph(sigName);
        body.appendParagraph(sigPhone);
        body.appendParagraph('\u05e0\u05e1 \u05de\u05d8\u05d7');
        body.appendParagraph(sigEmail);
      }

      doc.saveAndClose();
      copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      // Send email with PDF attachment
      var toEmail = decodeURIComponent(p.toEmail || '');
      var pdfBlob = DocumentApp.openById(copy.getId()).getAs('application/pdf');
      pdfBlob.setName(copyName + '.pdf');

      var emailSubject = '\u05e0\u05e1 \u05de\u05d8\u05d7 \u2014 \u05d4\u05e6\u05e2\u05ea \u05e9\u05d9\u05e8\u05d5\u05ea \u2014 ' + schoolName;
      var emailBody = '\u05e9\u05dc\u05d5\u05dd \u05e8\u05d1,\n\n';
      emailBody += '\u05de\u05e6\u05d5\u05e8\u05e3 \u05d4\u05e6\u05e2\u05ea \u05e9\u05d9\u05e8\u05d5\u05ea \u05e2\u05d1\u05d5\u05e8 ' + schoolName + ' (\u05e1\u05de\u05dc ' + code + ').\n\n';
      if (sigName) {
        emailBody += '\u05d1\u05d1\u05e8\u05db\u05d4,\n' + sigName + '\n' + sigPhone + '\n\u05e0\u05e1 \u05de\u05d8\u05d7\n' + sigEmail;
      }

      var mailOptions = {
        attachments: [pdfBlob],
        name: '\u05e0\u05e1 \u05de\u05d8\u05d7'
      };
      if (sigEmail) mailOptions.bcc = sigEmail;
      if (sigEmail) mailOptions.replyTo = sigEmail;

      GmailApp.sendEmail(toEmail, emailSubject, emailBody, mailOptions);

      var docUrl = 'https://docs.google.com/document/d/' + copy.getId() + '/edit?usp=sharing';
      return ContentService.createTextOutput(JSON.stringify({status:'ok', docUrl: docUrl, sent: true})).setMimeType(ContentService.MimeType.JSON);
     } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({status:'error', message: err.message})).setMimeType(ContentService.MimeType.JSON);
     }
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
            return ContentService.createTextOutput(JSON.stringify({authorized:true, name:dRows[i][1]||''})).setMimeType(ContentService.MimeType.JSON);
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
      return ContentService.createTextOutput(JSON.stringify({authorized:true, name:rows[i][2]||''})).setMimeType(ContentService.MimeType.JSON);
    }
    if (idType === 'domain' && type === 'email' && id.charAt(0) === '@') {
      if (cleanVal.indexOf(id) === cleanVal.length - id.length) {
        return ContentService.createTextOutput(JSON.stringify({authorized:true, name:rows[i][2]||value})).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }
  return ContentService.createTextOutput(JSON.stringify({authorized:false})).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}
