/****************************************************************************
 * Isha Karnataka — Google Apps Script backend (security-hardened)
 *
 * This turns a Google Sheet into the database + login server for the
 * Isha Karnataka web app. Paste this whole file into the Apps Script editor
 * of a Sheet (Extensions -> Apps Script), then deploy as a Web app.
 *
 * Tabs created automatically on first use:
 *   - Data      : single cell (A1) holding the centre tree as JSON
 *   - Users     : username | password | name | active | email | centre |
 *                 failed_attempts | locked_until
 *                 (password column stores hashed values in the format
 *                  v1:<salt>:<iter>:<hex-hash> — see hashPassword_)
 *   - Sessions  : token | username | created
 *   - Audit     : timestamp | actor | action | target | details
 *
 * Default admin login: username "admin", password "isha@2026".
 * The default password MUST be changed on first login — the backend will
 * refuse to issue a session for the admin while the default password is in
 * use. Call action "change_password" with old + new passwords to set it.
 ****************************************************************************/

/* ---- sheet names ------------------------------------------------------- */
var T_DATA  = 'Data';
var T_USERS = 'Users';
var T_SESS  = 'Sessions';
var T_AUDIT = 'Audit';

var BACKEND_VERSION = '2026-05-26-role-field';

/* ---- security config --------------------------------------------------- */
var PW_HASH_ITER          = 1000;                  // SHA-256 stretch iterations
var PW_HASH_PREFIX        = 'v1';                  // versioned format
var DEFAULT_ADMIN_PW      = 'isha@2026';           // bootstrap only — must be rotated
var SESSION_MAX_AGE_MS    = 7 * 24 * 60 * 60 * 1000;     // 7 days
var MAX_FAILED_LOGINS     = 5;                     // attempts within window before lockout
var FAILED_LOGIN_WINDOW_MS= 15 * 60 * 1000;        // reset counter after 15 min of no attempts
var LOCKOUT_MS            = 30 * 60 * 1000;        // 30 min lockout
var MAX_LEN_TEXT          = 200;                   // generic short text field limit
var MAX_LEN_LONG          = 2000;                  // long text field limit
var GENERIC_LOGIN_ERR     = 'Invalid email or password.'; // single error to prevent user enumeration

/* ---- entry points ------------------------------------------------------ */

function doPost(e) {
  var out;
  try {
    var raw = (e && e.postData && e.postData.contents) || '{}';
    if (raw.length > 200000) return json_({ ok: false, error: 'Payload too large.' });
    var req = JSON.parse(raw);
    switch (req.action) {
      case 'ping':             out = { ok: true, ts: Date.now(), version: BACKEND_VERSION }; break;
      case 'login':            out = login_(req); break;
      case 'signup':           out = signup_(req); break;
      case 'logout':           out = logout_(req); break;
      case 'change_password':  out = changePassword_(req); break;
      case 'load':             out = load_(req); break;
      case 'save':             out = save_(req); break;
      // admin-only endpoints
      case 'list_users':       out = list_users_(req); break;
      case 'approve_user':     out = approve_user_(req); break;
      case 'reject_user':      out = reject_user_(req); break;
      case 'list_audit':       out = list_audit_(req); break;
      default:                 out = { ok: false, error: 'Unknown action: ' + req.action };
    }
  } catch (err) {
    // Never echo internal error details to the client
    out = { ok: false, error: 'Internal error.' };
  }
  return json_(out);
}

function doGet() {
  return json_({
    ok: true,
    msg: 'Isha Karnataka backend is live.',
    version: BACKEND_VERSION,
    actions: ['ping','login','signup','logout','change_password','load','save',
              'list_users','approve_user','reject_user','list_audit']
  });
}

/* ---- helpers: response, sheet, schema ---------------------------------- */

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

// Utility: run this from the Apps Script editor (Run -> showSheetUrl), then
// open View -> Logs to see the URL of the backend Sheet this script is bound to.
function showSheetUrl() {
  var url = ss_().getUrl();
  Logger.log(url);
  return url;
}

/* ===========================================================================
 * EDITOR-ONLY RECOVERY UTILITIES
 * Run these from the Apps Script editor (pick the function, click Run, then
 * open the Execution log / View -> Logs). They are NOT exposed over the web
 * app (doPost) — they only work when run by you from the editor.
 * ===========================================================================*/

// Show the state of every account (active status, whether the password is
// stored hashed or as legacy plaintext, lockout). Does NOT print the password
// or hash itself. Use this to see WHY a login is being rejected.
function diagnoseLogins() {
  ensure_();
  var rows = sheet_(T_USERS).getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var stored = rows[i][1];
    Logger.log(
      'user="%s"  active="%s"  pwFormat=%s  locked=%s  email="%s"',
      rows[i][0],
      rows[i][3],
      isHashed_(stored) ? 'hashed' : (stored ? 'PLAINTEXT/other' : 'EMPTY'),
      isLockedOut_(rows[i]),
      rows[i][4]
    );
  }
}

// Reset ONE account's password to a known value and make it active + unlocked.
// 1. Edit EMAIL and NEW_PASSWORD below.
// 2. Run resetUserPassword from the editor.
// 3. Log in with that email + password.
// 4. (Recommended) blank these back out afterwards.
function resetUserPassword() {
  var EMAIL        = 'admin';          // <-- the username/email to fix (e.g. 'admin' or 'someone@email.com')
  var NEW_PASSWORD = 'Isha@2026!';     // <-- the new password (>=8 chars, a letter + a digit)

  ensure_();
  var target = String(EMAIL).trim().toLowerCase();
  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === target) {
      sh.getRange(i + 1, 2).setValue(hashPassword_(NEW_PASSWORD)); // password (hashed)
      sh.getRange(i + 1, 4).setValue('yes');                       // active
      clearFailedLogins_(i + 1);                                   // clear lockout
      // Drop any existing sessions so old tokens can't linger
      var ss = sheet_(T_SESS); var srows = ss.getDataRange().getValues();
      for (var j = srows.length - 1; j >= 1; j--) {
        if (String(srows[j][1]).toLowerCase() === target) ss.deleteRow(j + 1);
      }
      Logger.log('OK: "%s" reset to the new password, set active, unlocked.', rows[i][0]);
      return;
    }
  }
  Logger.log('No account found for "%s". Run diagnoseLogins to see existing accounts.', EMAIL);
}

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
  // Users schema (9 cols). Older sheets get the new columns back-filled.
  var headers = ['username','password','name','active','email','centre','failed_attempts','locked_until','role'];
  var u = sheet_(T_USERS, headers);
  if (u.getLastRow() < 2) {
    // Bootstrap admin with a HASHED default password
    u.appendRow(['admin', hashPassword_(DEFAULT_ADMIN_PW), 'Administrator', 'yes', '', '', 0, '', 'admin']);
  } else {
    var lastCol = u.getLastColumn();
    var headerRow = u.getRange(1, 1, 1, Math.max(lastCol, headers.length)).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (!headerRow[i]) u.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  sheet_(T_SESS, ['token','username','created']);
  sheet_(T_AUDIT, ['timestamp','actor','action','target','details']);
}

/* ---- helpers: security primitives -------------------------------------- */

// Constant-time string comparison (timing-attack resistant)
function constantTimeEq_(a, b) {
  a = String(a == null ? '' : a);
  b = String(b == null ? '' : b);
  if (a.length !== b.length) {
    // Still consume time on length mismatch
    var x = 0;
    for (var i = 0; i < a.length; i++) x |= a.charCodeAt(i);
    return false;
  }
  var r = 0;
  for (var j = 0; j < a.length; j++) r |= (a.charCodeAt(j) ^ b.charCodeAt(j));
  return r === 0;
}

function bytesToHex_(bytes) {
  var h = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] & 0xff;
    h += (b < 16 ? '0' : '') + b.toString(16);
  }
  return h;
}

// PBKDF2-ish stretch using SHA-256. Apps Script has no native bcrypt/scrypt;
// this iterates SHA-256 over (password + salt) PW_HASH_ITER times. Output
// format: v1:<salt>:<iterations>:<hex-hash>
function hashPassword_(password) {
  var salt = Utilities.getUuid();
  return hashWithSalt_(password, salt, PW_HASH_ITER);
}

function hashWithSalt_(password, salt, iter) {
  var s = String(password) + salt;
  for (var i = 0; i < iter; i++) {
    s = bytesToHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s));
  }
  return PW_HASH_PREFIX + ':' + salt + ':' + iter + ':' + s;
}

function verifyPassword_(password, stored) {
  if (!stored) return false;
  var s = String(stored);
  // Legacy plaintext (no version prefix) — supported for one-time migration.
  if (s.indexOf(PW_HASH_PREFIX + ':') !== 0) {
    return constantTimeEq_(String(password), s);
  }
  var parts = s.split(':');
  if (parts.length !== 4) return false;
  var salt = parts[1];
  var iter = parseInt(parts[2], 10);
  if (!salt || !iter || iter > 50000) return false; // guard against tampered values
  var expected = parts[3];
  var actual = hashWithSalt_(password, salt, iter).split(':')[3];
  return constantTimeEq_(actual, expected);
}

function isHashed_(stored) {
  return String(stored || '').indexOf(PW_HASH_PREFIX + ':') === 0;
}

// Password complexity policy
function validatePasswordStrength_(pw) {
  if (typeof pw !== 'string') return 'Password is required.';
  if (pw.length < 8)         return 'Password must be at least 8 characters.';
  if (pw.length > 128)       return 'Password is too long.';
  if (!/[A-Za-z]/.test(pw))  return 'Password must include at least one letter.';
  if (!/\d/.test(pw))        return 'Password must include at least one digit.';
  if (/\s/.test(pw))         return 'Password cannot contain spaces.';
  return null;
}

function validateEmail_(email) {
  if (typeof email !== 'string') return false;
  if (email.length > MAX_LEN_TEXT) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function clampStr_(s, max) {
  s = String(s == null ? '' : s);
  return s.length > max ? s.substring(0, max) : s;
}

/* ---- helpers: rate limiting ------------------------------------------- */

function recordFailedLogin_(rowIndex) {
  var sh = sheet_(T_USERS);
  var attempts = parseInt(sh.getRange(rowIndex, 7).getValue(), 10) || 0;
  attempts++;
  sh.getRange(rowIndex, 7).setValue(attempts);
  if (attempts >= MAX_FAILED_LOGINS) {
    sh.getRange(rowIndex, 8).setValue(Date.now() + LOCKOUT_MS);
  }
}

function clearFailedLogins_(rowIndex) {
  var sh = sheet_(T_USERS);
  sh.getRange(rowIndex, 7).setValue(0);
  sh.getRange(rowIndex, 8).setValue('');
}

function isLockedOut_(row) {
  var lockedUntil = row[7];
  if (!lockedUntil) return false;
  var until = parseInt(lockedUntil, 10);
  return until && Date.now() < until;
}

/* ---- helpers: audit log ----------------------------------------------- */

function audit_(actor, action, target, details) {
  try {
    var sh = sheet_(T_AUDIT, ['timestamp','actor','action','target','details']);
    sh.appendRow([
      new Date(),
      String(actor  || ''),
      String(action || ''),
      String(target || ''),
      details ? clampStr_(JSON.stringify(details), MAX_LEN_LONG) : ''
    ]);
  } catch (e) { /* never let audit failures break the request */ }
}

/* ---- auth: signup ----------------------------------------------------- */

function signup_(req) {
  ensure_();
  var email = clampStr_(String(req.email || req.username || '').trim(), MAX_LEN_TEXT);
  var pass  = String(req.password || '');
  var name  = clampStr_(String(req.name || '').trim(),   MAX_LEN_TEXT);
  var centre= clampStr_(String(req.centre || '').trim(), MAX_LEN_TEXT);
  var role  = clampStr_(String(req.role   || '').trim(), MAX_LEN_TEXT);
  if (!email || !pass) return { ok: false, error: 'Email and password are required.' };
  if (!validateEmail_(email)) return { ok: false, error: 'Please enter a valid email address.' };
  var pwErr = validatePasswordStrength_(pass);
  if (pwErr) return { ok: false, error: pwErr };

  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  var emailLower = email.toLowerCase();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === emailLower) {
      // Generic message to avoid email-enumeration via signup
      return { ok: false, error: 'If that email is not in use, the account will be created.' };
    }
  }
  // Use email as username (lowercased for consistency). Note: the "role" field
  // is the self-declared role at registration. It does NOT grant any
  // privilege — admin privileges remain tied to the literal username 'admin'.
  sh.appendRow([emailLower, hashPassword_(pass), name || emailLower, 'pending', email, centre, 0, '', role]);
  audit_(emailLower, 'signup', emailLower, { centre: centre, role: role });
  return { ok: true, pending: true };
}

/* ---- auth: login ------------------------------------------------------ */

function login_(req) {
  ensure_();
  var user = clampStr_(String(req.username || req.email || '').trim().toLowerCase(), MAX_LEN_TEXT);
  var pass = String(req.password || '');
  if (!user || !pass) return { ok: false, error: GENERIC_LOGIN_ERR };

  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() !== user) continue;

    // Lockout check (always first — blocks attempts even with correct password)
    if (isLockedOut_(rows[i])) {
      audit_(user, 'login_blocked_lockout', user, null);
      return { ok: false, error: 'Account temporarily locked due to repeated failed attempts. Try again later.', code: 'locked' };
    }

    // Password check (always run hash to keep timing comparable)
    var passOk = verifyPassword_(pass, rows[i][1]);

    if (!passOk) {
      recordFailedLogin_(i + 1);
      audit_(user, 'login_failed', user, null);
      return { ok: false, error: GENERIC_LOGIN_ERR };
    }

    // Active check
    var active = String(rows[i][3]).toLowerCase().trim();
    if (active === 'pending') {
      audit_(user, 'login_blocked_pending', user, null);
      return { ok: false, code: 'pending',
        error: 'Your account is pending admin approval. You will be able to log in once an admin has approved your access.' };
    }
    if (active === 'no' || active === 'false' || active === 'disabled') {
      audit_(user, 'login_blocked_disabled', user, null);
      return { ok: false, code: 'disabled',
        error: 'Your account has been disabled. Please contact an admin.' };
    }
    if (active !== 'yes' && active !== 'true' && active !== '') {
      audit_(user, 'login_blocked_unknown_state', user, null);
      return { ok: false, code: 'pending',
        error: 'Your account is pending admin approval.' };
    }

    // Admin still using default bootstrap password? Force a password change.
    if (isAdminUsername_(user) && verifyPassword_(DEFAULT_ADMIN_PW, rows[i][1])) {
      audit_(user, 'login_blocked_default_admin_pw', user, null);
      return { ok: false, code: 'must_change_password',
        error: 'The default admin password must be changed before you can log in. Use action change_password.' };
    }

    // Success! Reset attempt counter, migrate legacy plaintext to hashed.
    clearFailedLogins_(i + 1);
    if (!isHashed_(rows[i][1])) {
      sh.getRange(i + 1, 2).setValue(hashPassword_(pass));
      audit_(user, 'password_migrated', user, null);
    }

    // Issue session
    var token = Utilities.getUuid();
    sheet_(T_SESS).appendRow([token, rows[i][0], Date.now()]);
    audit_(user, 'login_success', user, null);
    return {
      ok: true,
      token: token,
      name: rows[i][2] || rows[i][0],
      username: rows[i][0],
      admin: isAdminUsername_(rows[i][0])
    };
  }

  // No matching user — still spend time on a fake hash to keep timing similar
  verifyPassword_(pass, hashPassword_('dummy-no-such-user'));
  audit_(user, 'login_failed_unknown_user', user, null);
  return { ok: false, error: GENERIC_LOGIN_ERR };
}

/* ---- auth: change password ------------------------------------------- */

function changePassword_(req) {
  ensure_();
  var user    = clampStr_(String(req.username || req.email || '').trim().toLowerCase(), MAX_LEN_TEXT);
  var oldPass = String(req.old_password || req.oldPassword || '');
  var newPass = String(req.new_password || req.newPassword || '');
  if (!user || !oldPass || !newPass) return { ok: false, error: 'Username, old password and new password are required.' };
  var pwErr = validatePasswordStrength_(newPass);
  if (pwErr) return { ok: false, error: pwErr };
  if (oldPass === newPass) return { ok: false, error: 'New password must be different from the current one.' };

  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() !== user) continue;
    if (isLockedOut_(rows[i])) return { ok: false, error: 'Account temporarily locked. Try again later.', code: 'locked' };
    if (!verifyPassword_(oldPass, rows[i][1])) {
      recordFailedLogin_(i + 1);
      audit_(user, 'change_password_failed', user, null);
      return { ok: false, error: 'Old password is incorrect.' };
    }
    sh.getRange(i + 1, 2).setValue(hashPassword_(newPass));
    clearFailedLogins_(i + 1);
    // Invalidate all existing sessions for this user (force re-login everywhere)
    var ss = sheet_(T_SESS); var srows = ss.getDataRange().getValues();
    for (var j = srows.length - 1; j >= 1; j--) {
      if (String(srows[j][1]).toLowerCase() === user) ss.deleteRow(j + 1);
    }
    audit_(user, 'change_password_success', user, null);
    return { ok: true };
  }
  return { ok: false, error: 'User not found.' };
}

/* ---- helpers: sessions ----------------------------------------------- */

function isAdminUsername_(u) {
  return String(u || '').trim().toLowerCase() === 'admin';
}

function userForToken_(token) {
  if (!token) return null;
  var tok = String(token);
  if (tok.length < 8 || tok.length > 64) return null;
  var sh = sheet_(T_SESS);
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (constantTimeEq_(String(rows[i][0]), tok)) {
      var created = parseInt(rows[i][2], 10) || 0;
      if (Date.now() - created > SESSION_MAX_AGE_MS) {
        // Expired — drop the row and reject
        sh.deleteRow(i + 1);
        return null;
      }
      return rows[i][1];
    }
  }
  return null;
}

function requireAdmin_(token) {
  var u = userForToken_(token);
  if (!u) return { ok: false, error: 'Not authenticated.', code: 'auth_required' };
  if (!isAdminUsername_(u)) return { ok: false, error: 'Admin access required.', code: 'forbidden' };
  return null;
}

function logout_(req) {
  var s = sheet_(T_SESS);
  var rows = s.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (constantTimeEq_(String(rows[i][0]), String(req.token))) s.deleteRow(i + 1);
  }
  return { ok: true };
}

/* ---- admin: list / approve / reject users + audit -------------------- */

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
      centre:   String(rows[i][5] || ''),
      failed_attempts: parseInt(rows[i][6], 10) || 0,
      locked:   isLockedOut_(rows[i]),
      role:     String(rows[i][8] || '')
    });
  }
  return { ok: true, users: out };
}

function approve_user_(req) {
  var gate = requireAdmin_(req.token); if (gate) return gate;
  var actor = userForToken_(req.token);
  var result = setUserActive_(req.username, 'yes');
  if (result.ok) audit_(actor, 'admin_approve_user', req.username, null);
  return result;
}

function reject_user_(req) {
  var gate = requireAdmin_(req.token); if (gate) return gate;
  var actor = userForToken_(req.token);
  var result = setUserActive_(req.username, 'no');
  if (result.ok) audit_(actor, 'admin_reject_user', req.username, null);
  return result;
}

function setUserActive_(username, value) {
  var sh = sheet_(T_USERS);
  var rows = sh.getDataRange().getValues();
  var target = clampStr_(String(username || '').trim().toLowerCase(), MAX_LEN_TEXT);
  if (!target) return { ok: false, error: 'Username required.' };
  if (target === 'admin') return { ok: false, error: 'Cannot modify the admin account.' };
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === target) {
      sh.getRange(i + 1, 4).setValue(value);
      // If approving, clear any lockout
      if (value === 'yes') clearFailedLogins_(i + 1);
      // If disabling, invalidate all sessions
      if (value !== 'yes') {
        var ss = sheet_(T_SESS); var srows = ss.getDataRange().getValues();
        for (var j = srows.length - 1; j >= 1; j--) {
          if (String(srows[j][1]).toLowerCase() === target) ss.deleteRow(j + 1);
        }
      }
      return { ok: true, username: rows[i][0], active: value };
    }
  }
  return { ok: false, error: 'User not found.' };
}

function list_audit_(req) {
  var gate = requireAdmin_(req.token); if (gate) return gate;
  var sh = sheet_(T_AUDIT, ['timestamp','actor','action','target','details']);
  var rows = sh.getDataRange().getValues();
  var out = [];
  // Return most recent 200 events, newest first
  var start = Math.max(1, rows.length - 200);
  for (var i = rows.length - 1; i >= start; i--) {
    out.push({
      timestamp: rows[i][0] ? new Date(rows[i][0]).getTime() : null,
      actor:     String(rows[i][1] || ''),
      action:    String(rows[i][2] || ''),
      target:    String(rows[i][3] || ''),
      details:   String(rows[i][4] || '')
    });
  }
  return { ok: true, events: out };
}

/* ---- data load / save (auth-gated) ----------------------------------- */

function load_(req) {
  if (!userForToken_(req.token)) return { ok: false, error: 'Not authenticated.', code: 'auth_required' };
  var v = sheet_(T_DATA).getRange('A1').getValue();
  if (!v) return { ok: true, data: null };
  try { return { ok: true, data: JSON.parse(v) }; }
  catch (e) { return { ok: true, data: null }; }
}

function save_(req) {
  if (!userForToken_(req.token)) return { ok: false, error: 'Not authenticated.', code: 'auth_required' };
  var payload;
  try { payload = JSON.stringify(req.data || {}); }
  catch (e) { return { ok: false, error: 'Invalid payload.' }; }
  if (payload.length > 1000000) return { ok: false, error: 'Payload too large.' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    sheet_(T_DATA).getRange('A1').setValue(payload);
  } catch (e) {
    return { ok: false, error: 'Busy, please retry.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return { ok: true };
}
