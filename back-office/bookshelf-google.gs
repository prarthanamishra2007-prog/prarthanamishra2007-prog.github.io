// ============================================================
// Prarthana's corner — "your stories" bookshelf engine
// (Google Apps Script — the tiny back office. Stores notes in a
//  Google Sheet, and lets the device that pinned a note remove it.)
//
// SETUP (needed only once):
//   1. Create a Google Sheet (any name).
//   2. Extensions → Apps Script → delete sample → paste this whole
//      file → Save.
//   3. Deploy → New deployment → type: Web app
//        Execute as : Me   ·  Who has access : Anyone
//   4. Copy the web-app URL (ends /exec) into stories.html.
//   After changing code: Deploy → Manage deployments → pencil →
//   Version: New version → Deploy.
// ============================================================

function doGet(e) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName('Shelf') || ss.insertSheet('Shelf');
  var data = sh.getDataRange().getValues();

  // if this is the old 3-column shelf, upgrade it to ids (keeps the notes)
  if (data.length && data[0][1] === 'name' && data[0][0] !== 'id') {
    var old = data;
    sh.clear();
    sh.appendRow(['id', 'name', 'message', 'time']);
    for (var r = 1; r < old.length; r++) {
      sh.appendRow(['legacy-' + old[r][2], old[r][0], old[r][1], old[r][2]]);
    }
    data = sh.getDataRange().getValues();
  }

  if (sh.getLastRow() === 0) sh.appendRow(['id', 'name', 'message', 'time']);

  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][2] || data[i][2] === 'message') continue;   // skip header + blanks
    out.push({
      id:      data[i][0] || '',
      name:    data[i][1] || 'someone',
      message: data[i][2] || '',
      time:    data[i][3] || ''
    });
  }
  out.reverse();   // newest first

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName('Shelf') || ss.insertSheet('Shelf');
  if (sh.getLastRow() === 0) sh.appendRow(['id', 'name', 'message', 'time']);

  var act = String(e.parameter.action || 'add');

  if (act === 'delete') {
    // find the row whose first cell matches this note's id, and remove it
    var data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(e.parameter.id)) {
        sh.deleteRow(i + 1);
        break;
      }
    }
  } else {
    // add a note (the id is made on the visitor's device)
    sh.appendRow([
      String(e.parameter.id      || ('legacy-' + new Date().getTime())),
      String(e.parameter.name    || 'someone'),
      String(e.parameter.message || ''),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    ]);
  }

  return ContentService
    .createTextOutput('{"ok":true}')
    .setMimeType(ContentService.MimeType.JSON);
}