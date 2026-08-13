---
description: how to create a mobile viewer app for Google Sheets data
---

# Create Mobile Viewer for Google Sheets

## Prerequisites
- A Google Sheet with data (published to web)
- The Google Sheet must be shared: **Anyone with the link → Viewer**

## Step 1: Get CSV Data URL
1. Open Google Sheet → File → Share → Publish to web
2. Select sheet tab, format: CSV → Publish
3. The CSV URL format: `https://docs.google.com/spreadsheets/d/e/{PUBLISHED_ID}/pub?gid=0&single=true&output=csv`

## Step 2: Download CSV and Create Self-Contained HTML
1. Download CSV via PowerShell:
```powershell
$url = "CSV_URL_HERE"
$wc = New-Object System.Net.WebClient
$respBytes = $wc.DownloadData($url)
$b64 = [Convert]::ToBase64String($respBytes)
```
2. Use the template from `C:\Users\USER\Documents\unbound\logisty-viewer\logisty2026_offline.html` as reference
3. Embed CSV as base64 in the HTML using `atob()` + `TextDecoder('utf-8')`
4. Adjust column mapping in `loadData()` to match the new sheet's columns
5. Adjust card template in `renderCards()` to display relevant fields

## Step 3: Create Google Apps Script for Cloud Notes
1. Open the Google Sheet → Extensions → Apps Script
2. Create `doPost(e)` and `doGet(e)` functions:
   - `doPost`: receives `{code, level, field, value}` and writes to "הערות" sheet tab
   - `doGet`: returns all saved notes as JSON
3. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone
4. First time: approve permissions (Advanced → Go to project → Allow)
5. Copy the deployment URL

## Step 4: Configure HTML with Script URL
1. Set `SCRIPT_URL` constant in the HTML to the Apps Script URL
2. Fields are saved to cloud automatically (debounced 1.5s) via `fetch(SCRIPT_URL, {method:'POST', mode:'no-cors', ...})`
3. On load, `loadCloudNotes()` fetches all notes via GET and merges with localStorage

## Step 5: Deploy to Mobile
- Since there's no Node.js, deploy via **Dropbox → WhatsApp/Email → Open in Chrome on Android**
- The file is self-contained (data embedded as base64), no network needed for data display
- Network only needed for cloud notes sync

## Key Architecture Decisions
- **Self-contained HTML**: CSV data embedded as base64 to avoid CORS issues
- **localStorage**: Immediate local save for responsiveness
- **Google Apps Script**: Free cloud backend, no API keys needed
- **mode: 'no-cors'**: POST works but response is opaque (save indicator based on .then() not response)
- **Separate "הערות" sheet tab**: Keeps original data untouched
- **JSONP not needed**: Apps Script supports CORS for GET requests

## Features Template
- RTL Hebrew interface
- Search bar (text search across all fields)
- Authority dropdown filter
- Education level chip filters
- Card-based UI with expandable details
- Contact fields (name + phone) with tap-to-call (tel:) and copy (clipboard)
- Email field with mailto: and copy
- Multi-line notes textarea
- Auto-save to cloud + localStorage fallback
- Conditional card border color based on field value
- Android back button closes expanded card
- Select-all on input focus
- Summary bar with totals

## File Locations
- Template source: `C:\Users\USER\Documents\unbound\logisty-viewer\logisty2026_offline.html`
- Dropbox copy: `C:\Users\USER\Dropbox\VB6\VBA\CascadeProjects\windsurf-project\logisty2026.html`
- Backup: `C:\Users\USER\Documents\unbound\logisty-viewer\Backup_20260516_0939\`

## Apps Script Template
```javascript
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("הערות");
  
  if (!sheet) {
    sheet = ss.insertSheet("הערות");
    sheet.appendRow(["מפתח1","מפתח2","הערה","שדה4","שדה5","...","עדכון"]);
  }
  
  var colMap = { "note":3, "field1":4, "field2":5 }; // adjust per project
  var key = data.code + "_" + data.level;
  var rows = sheet.getDataRange().getValues();
  var found = false;
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] + "_" + rows[i][1] === key) {
      var col = colMap[data.field];
      if (col) sheet.getRange(i+1, col).setValue(data.value);
      sheet.getRange(i+1, LAST_COL).setValue(new Date());
      found = true;
      break;
    }
  }
  
  if (!found && data.value.trim()) {
    var row = [data.code, data.level, "", "", "", new Date()];
    var col = colMap[data.field];
    if (col) row[col-1] = data.value;
    sheet.appendRow(row);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status:"ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("הערות");
  if (!sheet) return ContentService.createTextOutput("{}").setMimeType(ContentService.MimeType.JSON);
  
  var rows = sheet.getDataRange().getValues();
  var notes = {};
  for (var i = 1; i < rows.length; i++) {
    var key = rows[i][0] + "_" + rows[i][1];
    notes[key] = { note: rows[i][2]||"", field1: rows[i][3]||"" }; // adjust per project
  }
  return ContentService.createTextOutput(JSON.stringify(notes))
    .setMimeType(ContentService.MimeType.JSON);
}
```
