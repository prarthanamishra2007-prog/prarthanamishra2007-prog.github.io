// ============================================================
// Prarthana's corner — back office engine
// (Google Apps Script)
//
// Two little jobs:
//   1. The bookshelf of notes (public — anyone can pin)
//   2. YOUR writing (posts for thoughts/books/voices —
//      private: only the secret key can add or remove)
//
// SETUP (needed only once):
//   1. Create a Google Sheet (any name).
//   2. Extensions → Apps Script → delete sample → paste this
//      whole file → Save.
//   3. Deploy → New deployment → type: Web app
//        Execute as : Me   ·   Who has access : Anyone
//   4. Copy the web-app URL (ends /exec) and paste it into:
//        - stories.html   (const SHELF_URL)
//        - posts.js       (const POSTS_URL)  ← for your writing
//   5. The secret key below MUST match the hidden box in
//      writer.html. Pick your own phrase and use it in both spots.
//   After changing code: Deploy → Manage deployments → pencil →
//   Version: New version → Deploy.
// ============================================================

// ---- your private key for publishing posts (change to your own phrase) ----
var BLOG_KEY = 'prarthana-pours-the-tea';

// ---- read: routes requests ----
function doGet(e) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var act = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : 'shelf';
  if (act === 'posts') return getPosts(ss, (e && e.parameter) ? e.parameter : {});
  return getShelf(ss);
}

// ---- the bookshelf notes (unchanged behaviour) ----
function getShelf(ss) {
  var sh   = ss.getSheetByName('Shelf') || ss.insertSheet('Shelf');
  var data = sh.getDataRange().getValues();

  // upgrade an old 3-column shelf to the new one with ids (keeps the notes)
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
    if (!data[i][2] || data[i][2] === 'message') continue;
    out.push({
      id:      data[i][0] || '',
      name:    data[i][1] || 'someone',
      message: data[i][2] || '',
      time:    data[i][3] || ''
    });
  }
  out.reverse();
  return textJson(out);
}

// ---- your posts, newest first (optional room filter) ----
function getPosts(ss, p) {
  var sh   = ensurePosts(ss);
  var data = sh.getDataRange().getValues();
  var room = p.room ? String(p.room) : '';
  var out  = [];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][2]) continue;                       // no title → skip
    if (room && String(data[i][1]) !== room) continue;
    out.push({
      id:      data[i][0] || '',
      room:    data[i][1] || 'thoughts',
      title:   data[i][2] || '',
      desc:    data[i][3] || '',
      body:    data[i][4] || '',
      date:    data[i][5] || '',
      order:   Number(data[i][6]) || 0,
      style:   data[i][7] || 'standard',
      image:   data[i][8] || '',
      caption: data[i][9] || ''
    });
  }
  out.sort(function (a, b) { return b.order - a.order; });   // newest first
  return textJson(out);
}

// ---- write: routes add/delete for the shelf and for posts ----
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var p  = e.parameter;
  var act = String(p.action || 'add');

  // ---- YOUR writing: requires the secret key ----
  if (act === 'addpost' || act === 'deletepost') {
    if (String(p.key || '') !== BLOG_KEY) return textJson({ ok:false, err:'key' });

    if (act === 'addpost') {
      var sh = ensurePosts(ss);
      sh.appendRow([
        String(p.id      || ''),
        String(p.room    || 'thoughts'),
        String(p.title   || ''),
        String(p.desc    || ''),
        String(p.body    || ''),
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        new Date().getTime(),
        String(p.style   || 'standard'),
        String(p.image   || ''),
        String(p.caption || '')
      ]);
    } else {
      var sh = ensurePosts(ss);
      var rows = sh.getDataRange().getValues();
      for (var i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][0]).trim() === String(p.id).trim()) { sh.deleteRow(i + 1); break; }
      }
    }
    return textJson({ ok:true });
  }

  // ---- the public bookshelf ----
  var sh2 = ss.getSheetByName('Shelf') || ss.insertSheet('Shelf');
  if (sh2.getLastRow() === 0) sh2.appendRow(['id', 'name', 'message', 'time']);

  if (act === 'delete') {
    var data = sh2.getDataRange().getValues();
    for (var j = data.length - 1; j >= 1; j--) {
      if (String(data[j][0]) === String(p.id)) { sh2.deleteRow(j + 1); break; }
    }
  } else {
    sh2.appendRow([
      String(p.id      || ('legacy-' + new Date().getTime())),
      String(p.name    || 'someone'),
      String(p.message || ''),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    ]);
  }
  return textJson({ ok:true });
}

// ---- helpers ----
function ensurePosts(ss) {
  var sh = ss.getSheetByName('Posts') || ss.insertSheet('Posts');

  // make sure the sheet has ALL the columns a post can use
  // (adds the newer ones — style, image, caption — if the sheet
  //  was created before they existed, so old posts are kept safe)
  var need = ['id', 'room', 'title', 'desc', 'body', 'date', 'order', 'style', 'image', 'caption'];
  if (sh.getLastRow() === 0) {
    sh.appendRow(need);
  } else {
    var have = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var c = 0; c < need.length; c++) {
      if (!have[c] || have[c] === '') {
        sh.getRange(1, c + 1).setValue(need[c]);
      }
    }
  }
  return sh;
}

function textJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}