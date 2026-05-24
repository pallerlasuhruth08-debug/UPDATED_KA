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

// Bump this whenever you redeploy so the frontend can confirm
// which version is live. Visit the /exec URL in a browser to see it.
var BACKEND_VERSION = '2026-05-25-signup-email-centre';

/* ---- entry points ------------------------------------------------------ */

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (req.action) {
      case 'ping':         out = { ok: true, ts: Date.now(), version: BACKEND_VERSION }; break;
      case 'login':        out = login_(req); break;
      case 'signup':       out = signup_(req); break;
      case 'logout':       out = logout_(req); break;
      case 'load':         out = load_(req); break;
      case 'save':         out = save_(req); break;
      // admin-only endpoints
      case 'list_users':   out = list_users_(req); break;
      case 'approve_user': out = approve_user_(req); break;
      case 'reject_user':  out = reject_user_(req); break;
      default:             out = { ok: false, error: 'Unknown action: ' + req.action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json_(out);
}

function doGet() {
  // Friendly health check + version, so you can verify the deployment is current.
  return json_({
    ok: true,
    msg: 'Isha Karnataka backend is live.',
    version: BACKEND_VERSION,
    actions: ['ping','login','signup','logout','load','save','list_users','approve_user','reject_user']
  });
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
  // Users schema: username | password | name | active | email | centre
  // (email + centre were added later; existing rows just have blank values.)
  var u = sheet_(T_USERS, ['username', 'password', 'name', 'active', 'email', 'centre']);
  if (u.getLastRow() < 2) {
    u.appendRow(['admin', 'isha@2026', 'Administrator', 'yes', '', '']);
  } else {
    // Back-fill new column headers if an older sheet is missing them
    var headerRow = u.getRange(1, 1, 1, Math.max(u.getLastColumn(), 6)).getValues()[0];
    if (!headerRow[4]) u.getRange(1, 5).setValue('email');
    if (!headerRow[5]) u.getRange(1, 6).setValue('centre');
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
    if (String(rows[i][0]).trim() === user && String(rows[i][1]) === pass) {
      var active = String(rows[i][3]).toLowerCase().trim();
      if (active === 'pending') {
        return { ok: false, code: 'pending',
          error: 'Your account is pending admin approval. You will be able to log in once an admin has approved your access.' };
      }
      if (active === 'no' || active === 'false' || active === 'disabled') {
        return { ok: false, code: 'disabled',
          error: 'Your account has been disabled. Please contact an admin.' };
      }
      if (active !== 'yes' && active !== 'true' && active !== '') {
        return { ok: false, code: 'pending',
          error: 'Your account is pending admin approval.' };
      }
      var token = Utilities.getUuid();
      sheet_(T_SESS).appendRow([token, rows[i][0], Date.now()]);
      return { ok: true, token: token, name: rows[i][2] || rows[i][0],
               username: rows[i][0], admin: isAdminUsername_(rows[i][0]) };
    }
  }
  return { ok: false, error: 'Invalid username or password.' };
}

function isAdminUsername_(u) {
  return String(u || '').trim().toLowerCase() === 'admin';
}

function requireAdmin_(token) {
  var u = userForToken_(token);
  if (!u) return { ok: false, error: 'Not authenticated.' };
  if (!isAdminUsername_(u)) return { ok: false, error: 'Admin access required.' };
  return null; // ok
}

function signup_(req) {
  ensure_();
  // Email is the primary identifier for new signups; we store it in the
  // username column too so login_ (which keys off username) keeps working.
  var email  = String(req.email || req.username || '').trim();
  var user   = email;
  var pass   = String(req.password || '');
  var name   = String(req.name || '').trim();
  var centre = String(req.centre || '').trim();
  if (!user || !pass) return { ok: false, error: 'Email and password are required.' };
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, error: 'Please enter a valid email address.' };
  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === user.toLowerCase()) {
      return { ok: false, error: 'An account with that email already exists.' };
    }
  }
  // Account is created in PENDING state. An admin must approve (set active='yes')
  // — either via the Admin page in the app or by editing the sheet directly.
  sh.appendRow([user, pass, name || user, 'pending', email, centre]);
  return { ok: true, pending: true };
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

/* ---- admin: list / approve / reject users ------------------------------ */

function list_users_(req) {
  var gate = requireAdmin_(req.token); if (gate) return gate;
  var rows = sheet_(T_USERS).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    out.push({
      username: String(rows[i][0] || ''),
      name:     String(rows[i][2] || ''),
      active:   String(rows[i][3] || '').toLowerCase().trim(),
      email:    String(rows[i][4] || ''),
      centre:   String(rows[i][5] || '')
    });
  }
  return { ok: true, users: out };
}

function approve_user_(req) {
  var gate = requireAdmin_(req.token); if (gate) return gate;
  return setUserActive_(req.username, 'yes');
}

function reject_user_(req) {
  var gate = requireAdmin_(req.token); if (gate) return gate;
  // "Reject" disables the account rather than deleting it, so a username
  // can't immediately be re-registered by someone else (audit trail).
  return setUserActive_(req.username, 'no');
}

function setUserActive_(username, value) {
  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  var target = String(username || '').trim().toLowerCase();
  if (!target) return { ok: false, error: 'Username required.' };
  if (target === 'admin') return { ok: false, error: 'Cannot modify the admin account.' };
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === target) {
      sh.getRange(i + 1, 4).setValue(value); // active is column 4 (1-indexed)
      return { ok: true, username: rows[i][0], active: value };
    }
  }
  return { ok: false, error: 'User not found.' };
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
