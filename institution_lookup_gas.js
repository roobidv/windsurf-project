/**
 * Google Apps Script - Institution Lookup Proxy
 * ==============================================
 * סוכן חכם לחיפוש נתוני מוסדות חינוך בישראל
 *
 * התקנה:
 * 1. היכנס ל-https://script.google.com ולחץ "פרויקט חדש"
 * 2. העתק את כל התוכן של קובץ זה לעורך
 * 3. לחץ Deploy → New deployment → Web app
 * 4. הגדר: Execute as = Me, Who has access = Anyone
 * 5. לחץ Deploy והעתק את כתובת ה-URL
 * 6. הדבק את הכתובת בדף ה-HTML
 *
 * שימוש ישיר (ללא HTML):
 *   GET https://script.google.com/macros/s/YOUR_ID/exec?semel=310011
 *
 * תשובה (JSON):
 *   {
 *     "semel": "310011",
 *     "name": "גימס רוטשילד",
 *     "address": "שד ירושלים, אור עקיבא",
 *     "phone": "04-6361277",
 *     "principal": "רינת מראד",
 *     "email": "school@example.co.il",
 *     "source": "allschool.co.il"
 *   }
 */

function doGet(e) {
  var semel = (e && e.parameter && e.parameter.semel) || '';
  semel = semel.replace(/\D/g, '');

  if (!semel || semel.length !== 6) {
    return jsonResponse({ error: 'יש לספק סמל מוסד בן 6 ספרות', param: 'semel' });
  }

  try {
    var result = lookupInstitution(semel);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message, semel: semel });
  }
}

/**
 * Main lookup function - scrapes allschool.co.il
 */
function lookupInstitution(semel) {
  // allschool.co.il redirects /מוסדות/{semel}/any-slug → correct URL
  var url = 'https://www.allschool.co.il/%D7%9E%D7%95%D7%A1%D7%93%D7%95%D7%AA/' + semel + '/lookup';

  var options = {
    followRedirects: true,
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();

  if (statusCode === 404) {
    throw new Error('מוסד עם סמל ' + semel + ' לא נמצא');
  }

  if (statusCode !== 200) {
    throw new Error('שגיאה בגישה למקור המידע (HTTP ' + statusCode + ')');
  }

  var html = response.getContentText('UTF-8');
  return parseAllSchoolPage(semel, html);
}

/**
 * Parse the allschool.co.il HTML page and extract key fields
 */
function parseAllSchoolPage(semel, html) {
  // Remove HTML tags to get plain text lines
  var text = html.replace(/<[^>]+>/g, '\n');
  var lines = text.split('\n');
  var cleanLines = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/^\s+|\s+$/g, '');
    if (line) cleanLines.push(line);
  }

  var result = {
    semel: semel,
    name: '',
    address: '',
    phone: '',
    principal: '',
    email: '',
    source: 'allschool.co.il'
  };

  // Extract institution name from page title pattern: "בית ספר {name}"
  var titleMatch = html.match(/<title[^>]*>([^<]+)<\/title>/i);
  if (titleMatch) {
    var title = titleMatch[1].replace(/^\s+|\s+$/g, '');
    // Remove common prefixes/suffixes
    title = title.replace(/\s*[-–|,]\s*אולסקול.*$/i, '');
    title = title.replace(/^בית ספר\s+/i, '');
    result.name = title;
  }

  // Find key fields by looking for label → value pattern in consecutive lines
  for (var j = 0; j < cleanLines.length; j++) {
    var current = cleanLines[j];
    var next = (j + 1 < cleanLines.length) ? cleanLines[j + 1] : '';

    if (current === 'כתובת' && next && next !== 'טלפון') {
      result.address = next;
    }
    else if (current === 'טלפון' && next && /^\d{2,3}[-]?\d{7}$/.test(next)) {
      result.phone = next;
    }
    else if (current === 'מנהל/ת' && next && !/^[\d\s]+$/.test(next)) {
      result.principal = next;
    }
    else if ((current === 'דוא"ל' || current === 'אימייל' || current === 'email' || current === 'דוא\"ל') && next && next.indexOf('@') !== -1) {
      result.email = next;
    }
  }

  // Fallback: try regex patterns if consecutive-line parsing missed something
  if (!result.phone) {
    var phoneMatch = text.match(/0[2-9]-?\d{7}/);
    if (phoneMatch) result.phone = phoneMatch[0];
  }

  // Fallback: try to find email address in page
  if (!result.email) {
    var emailMatch = text.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
    if (emailMatch) result.email = emailMatch[0];
  }

  return result;
}

/**
 * Return a JSON response with proper CORS headers
 */
function jsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Test function - run this in the Apps Script editor to verify
 */
function testLookup() {
  var result = lookupInstitution('310011');
  Logger.log(JSON.stringify(result, null, 2));

  // Expected output:
  // {
  //   "semel": "310011",
  //   "name": "גימס רוטשילד, אור עקיבא",
  //   "address": "שד ירושלים, אור עקיבא",
  //   "phone": "04-6361277",
  //   "principal": "רינת מראד",
  //   "email": "school@example.co.il",
  //   "source": "allschool.co.il"
  // }
}
