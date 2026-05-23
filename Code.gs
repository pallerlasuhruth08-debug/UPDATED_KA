/****************************************************************************
 * Isha Karnataka — Google Apps Script backend
 *
 * This turns a Google Sheet into the database + login server for the
 * Isha Karnataka web app. Paste this whole file into the Apps Script editor
 * of a Sheet (Extensions -> Apps Script), then deploy as a Web app.
 * See the setup guide for click-by-click steps.
 *
 * It creates three tabs automatically on first use:
 *   - Data      : a single cell (A1) holding the whole centre tree as JSON
 *   - Users     : username | password | name | active   (your logins)
 *   - Sessions  : token | username | created            (managed for you)
 *
 * A default admin login is created the first time:
 *      username: admin     password: isha@2026
 * Change that password in the Users tab right after your first login.
 ****************************************************************************/

var T_DATA = 'Data';
var T_USERS = 'Users';
var T_SESS = 'Sessions';

/* ---- entry points ------------------------------------------------------ */

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (req.action) {
      case 'ping':   out = { ok: true, ts: Date.now() }; break;
      case 'login':  out = login_(req); break;      case 'signup': out = signup_(req); break;
      case 'logout': out = logout_(req); break;
      case 'load':   out = load_(req); break;
      case 'save':   out = save_(req); break;
      default:       out = { ok: false, error: 'Unknown action: ' + req.action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json_(out);
}

function doGet() {
  // a friendly health check if someone opens the URL in a browser
  return json_({ ok: true, msg: 'Isha Karnataka backend is live.' });
}

/* ---- helpers ----------------------------------------------------------- */

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) sh.appendRow(headers);
  }
  return sh;
}

function ensure_() {
  sheet_(T_DATA);
  var u = sheet_(T_USERS, ['username', 'password', 'name', 'active']);
  if (u.getLastRow() < 2) {
    u.appendRow(['admin', 'isha@2026', 'Administrator', 'yes']);
  }
  sheet_(T_SESS, ['token', 'username', 'created']);
}

/* ---- auth -------------------------------------------------------------- */

function login_(req) {
  ensure_();
  var rows = sheet_(T_USERS).getDataRange().getValues();
  var user = String(req.username || '').trim();
  var pass = String(req.password || '');
  for (var i = 1; i < rows.length; i++) {
    var active = String(rows[i][3]).toLowerCase();
    if (String(rows[i][0]).trim() === user &&
        String(rows[i][1]) === pass &&
        active !== 'no' && active !== 'false') {
      var token = Utilities.getUuid();
      sheet_(T_SESS).appendRow([token, rows[i][0], Date.now()]);
      return { ok: true, token: token, name: rows[i][2] || rows[i][0] };
    }
  }
  return { ok: false, error: 'Invalid username or password.' };
}

function signup_(req) {
  ensure_();
  var user = String(req.username || '').trim();
  var pass = String(req.password || '');
  var name = String(req.name || '').trim();
  if (!user || !pass) return { ok: false, error: 'Username and password required.' };
  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === user.toLowerCase()) {
      return { ok: false, error: 'Username already exists.' };
    }
  }
  sh.appendRow([user, pass, name || user, 'yes']);
  var token = Utilities.getUuid();
  sheet_(T_SESS).appendRow([token, user, Date.now()]);
  return { ok: true, token: token, name: name || user };
}

function userForToken_(token) {
  if (!token) return null;
  var rows = sheet_(T_SESS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(token)) return rows[i][1];
  }
  return null;
}

function logout_(req) {
  var s = sheet_(T_SESS);
  var rows = s.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(req.token)) s.deleteRow(i + 1);
  }
  return { ok: true };
}

/* ---- data load / save -------------------------------------------------- */

function load_(req) {
  if (!userForToken_(req.token)) return { ok: false, error: 'Not authenticated.' };
  var v = sheet_(T_DATA).getRange('A1').getValue();
  if (!v) return { ok: true, data: null };
  try { return { ok: true, data: JSON.parse(v) }; }
  catch (e) { return { ok: true, data: null }; }
}

function save_(req) {
  if (!userForToken_(req.token)) return { ok: false, error: 'Not authenticated.' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    sheet_(T_DATA).getRange('A1').setValue(JSON.stringify(req.data || {}));
  } catch (e) {
    return { ok: false, error: 'Busy, please retry.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return { ok: true };
}
