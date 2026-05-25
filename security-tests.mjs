#!/usr/bin/env node
/**
 * Security test suite for the Isha Karnataka web app.
 *
 * Runs two kinds of checks:
 *
 *   STATIC  — greps the frontend file for known-bad patterns
 *             (unescaped user data in templates, missing CSP, etc.)
 *
 *   LIVE    — POSTs the deployed Apps Script backend with a battery of
 *             attack-shaped payloads (XSS, injection, weak passwords,
 *             missing auth, rate-limit triggers, etc.) and asserts the
 *             backend reacts safely.
 *
 * Live tests will create at most a handful of "smoke-test-*@example.com"
 * accounts in the Users sheet (left in pending state — never logged in).
 *
 * Usage:
 *   node security-tests.mjs
 *   node security-tests.mjs --skip-live      # only static checks
 *   node security-tests.mjs --url=<exec-url> # override backend URL
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FRONTEND = path.join(__dirname, 'centre-os-prototype.html');
const BACKEND  = path.join(__dirname, 'Code.gs');

const args = Object.fromEntries(
  process.argv.slice(2).map(a => a.startsWith('--') ? a.replace(/^--/,'').split('=').concat([true]).slice(0,2) : [a, true])
);
const SKIP_LIVE = !!args['skip-live'];
const API_URL = args.url || extractApiUrl();

let pass = 0, fail = 0;
const failures = [];

function it(name, fn){
  try {
    const r = fn();
    if (r && typeof r.then === 'function') {
      return r.then(() => mark(name, true)).catch(e => mark(name, false, e));
    }
    return mark(name, true);
  } catch (e) {
    return mark(name, false, e);
  }
}
function mark(name, ok, e){
  if (ok) { pass++; console.log('  \x1b[32m✓\x1b[0m', name); }
  else { fail++; failures.push([name, e]); console.log('  \x1b[31m✗\x1b[0m', name, e?.message ? '— '+e.message : ''); }
}
function assert(cond, msg){ if(!cond) throw new Error(msg||'assertion failed'); }
function assertEq(a, b, msg){ if(a!==b) throw new Error((msg||'expected eq')+' — got '+JSON.stringify(a)+', expected '+JSON.stringify(b)); }
function assertMatch(s, re, msg){ if(!re.test(String(s))) throw new Error((msg||'expected match '+re)+' — got '+JSON.stringify(s)); }
function assertNoMatch(s, re, msg){ if(re.test(String(s))) throw new Error((msg||'unexpected match '+re)+' in '+JSON.stringify(s).slice(0,120)); }

function section(title){ console.log('\n\x1b[1m'+title+'\x1b[0m'); }

function extractApiUrl(){
  const html = fs.readFileSync(FRONTEND, 'utf8');
  const m = html.match(/DEFAULT_API_URL='([^']+)'/);
  if(!m) throw new Error('Could not find DEFAULT_API_URL in '+FRONTEND);
  return m[1];
}

/* =========================================================================
 *  STATIC TESTS — read source code, ensure security-relevant patterns hold
 * ========================================================================= */
function readFiles(){
  return {
    html: fs.readFileSync(FRONTEND, 'utf8'),
    gs:   fs.readFileSync(BACKEND, 'utf8'),
  };
}

function staticTests(){
  section('Static frontend audit');
  const { html, gs } = readFiles();

  it('frontend declares Content-Security-Policy', () => {
    assertMatch(html, /<meta http-equiv="Content-Security-Policy"/i);
    assertMatch(html, /frame-ancestors 'none'/i);
    assertMatch(html, /object-src 'none'/i);
  });

  it('frontend sets X-Content-Type-Options=nosniff', () => {
    assertMatch(html, /<meta http-equiv="X-Content-Type-Options" content="nosniff"/i);
  });

  it('frontend has frame-busting script', () => {
    assertMatch(html, /window\.top\s*!==\s*window\.self/);
  });

  it('esc() helper escapes all 5 HTML-significant chars', () => {
    const fn = html.match(/function esc\(v\)\{[\s\S]+?\n\}/);
    assert(fn, 'esc function not found');
    const body = fn[0];
    for (const seq of ['&amp;','&lt;','&gt;','&quot;','&#39;']) {
      assertMatch(body, new RegExp(seq.replace(/[&]/g,'&').replace(/[<>]/g,c=>'\\'+c)), 'esc must produce '+seq);
    }
  });

  it('all target="_blank" links use rel="noopener"', () => {
    const links = html.match(/target="_blank"[^>]*>/g) || [];
    const bad = links.filter(l => !/rel\s*=\s*"[^"]*noopener/.test(l));
    assert(bad.length === 0, 'links missing rel=noopener: '+bad.slice(0,3).join(' | '));
  });

  it('no unescaped person-field interpolations in HTML templates', () => {
    // Allow ${esc(personX(...))} or personX(...) used as a variable check (not in template)
    // but reject patterns like ${personName(p)} directly in backtick template strings.
    const lines = html.split('\n');
    const bad = [];
    lines.forEach((l, i) => {
      // Skip the personReadView assignment site that's inside esc(...)
      if (/\$\{(person(Name|Mobile|Email|Profession|Occupation|Work|Education|Programs|Volunteering|Social))\(/.test(l)
          && !/esc\(person/.test(l)
          && !/personHas\(/.test(l)) {
        bad.push((i+1)+': '+l.trim().slice(0,140));
      }
    });
    assert(bad.length === 0, 'unescaped person fields:\n  '+bad.join('\n  '));
  });

  it('no unescaped centre/sector/place name interpolations', () => {
    const lines = html.split('\n');
    const bad = [];
    lines.forEach((l, i) => {
      // Look for ${X.name} or ${X.name||...} in HTML template contexts
      // but not for hardcoded role-meta (r.name, meta.name, sub.name, parent.name)
      const m = l.match(/\$\{[ns]\.name\b/g) || l.match(/\$\{[cp]\.name\b/g);
      if (!m) return;
      if (/esc\([ns]\.name\)/.test(l) || /esc\([cp]\.name\)/.test(l)) return;
      // Allowed: titles fed to showModal() or modalForm(), since both escape the title themselves
      if (/^\s*(showModal|modalForm)\s*\(/.test(l)) return;
      bad.push((i+1)+': '+l.trim().slice(0,140));
    });
    assert(bad.length === 0, 'unescaped names:\n  '+bad.join('\n  '));
  });

  it('no innerHTML direct assignment of UNAME without esc', () => {
    assertNoMatch(html, /innerHTML\s*=\s*[`'"][^`'"]*\$\{UNAME\b/);
  });

  it('localStorage isha_api purged when it differs from canonical URL', () => {
    assertMatch(html, /purging stale isha_api/);
    assertMatch(html, /localStorage\.removeItem\(['"]isha_api['"]\)/);
  });

  it('no auth-bypass _test-bypass.html in repo', () => {
    assert(!fs.existsSync(path.join(__dirname,'_test-bypass.html')), '_test-bypass.html must not exist');
  });

  section('Static backend audit');

  it('passwords are hashed, not stored plaintext', () => {
    assertMatch(gs, /function hashPassword_\b/);
    assertMatch(gs, /Utilities\.computeDigest/);
    assertMatch(gs, /PW_HASH_ITER\s*=\s*\d+/);
  });

  it('rate limiting: failed_attempts + lockout', () => {
    assertMatch(gs, /MAX_FAILED_LOGINS/);
    assertMatch(gs, /LOCKOUT_MS/);
    assertMatch(gs, /isLockedOut_/);
  });

  it('session expiry: SESSION_MAX_AGE_MS enforced in userForToken_', () => {
    assertMatch(gs, /SESSION_MAX_AGE_MS\s*=\s*\d+/);
    const fn = gs.match(/function userForToken_[\s\S]+?\n\}/);
    assert(fn, 'userForToken_ not found');
    assertMatch(fn[0], /SESSION_MAX_AGE_MS/);
  });

  it('constant-time comparison helper exists and is used in verifyPassword_', () => {
    assertMatch(gs, /function constantTimeEq_\b/);
    const fn = gs.match(/function verifyPassword_[\s\S]+?\n\}/);
    assert(fn && /constantTimeEq_/.test(fn[0]), 'verifyPassword_ must use constantTimeEq_');
  });

  it('password complexity policy enforced at signup', () => {
    assertMatch(gs, /function validatePasswordStrength_\b/);
    assertMatch(gs, /Password must be at least 8 characters/);
  });

  it('default admin password rotation enforced on login', () => {
    assertMatch(gs, /code: 'must_change_password'/);
    assertMatch(gs, /DEFAULT_ADMIN_PW/);
  });

  it('login uses generic error to prevent user-enumeration', () => {
    assertMatch(gs, /GENERIC_LOGIN_ERR\s*=\s*'Invalid email or password\.'/);
  });

  it('audit log written for sensitive admin actions', () => {
    assertMatch(gs, /function audit_\b/);
    assertMatch(gs, /admin_approve_user/);
    assertMatch(gs, /admin_reject_user/);
  });

  it('payload size limits on doPost and save_', () => {
    assertMatch(gs, /Payload too large/);
    assertMatch(gs, /raw\.length\s*>\s*\d+/);
  });

  it('internal error messages are not leaked to clients', () => {
    const fn = gs.match(/function doPost[\s\S]+?\n\}/);
    assert(fn && /Internal error/.test(fn[0]), 'doPost must wrap exceptions with a generic message');
    assertNoMatch(fn[0], /error:\s*String\(err\)/);
  });
}

/* =========================================================================
 *  LIVE TESTS — call deployed backend with attack-shaped payloads
 * ========================================================================= */
async function call(action, payload){
  const body = JSON.stringify(Object.assign({action}, payload||{}));
  const res = await fetch(API_URL, { method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body, redirect:'follow' });
  const text = await res.text();
  try { return JSON.parse(text); }
  catch { return { _nonjson: true, status: res.status, snippet: text.slice(0,150) }; }
}

async function liveTests(){
  section('Live backend probes  (against '+API_URL.slice(0,80)+'…)');
  // Sanity: backend must respond and advertise the hardened version
  let version = null;
  await it('ping returns ok + version', async () => {
    const r = await call('ping');
    assert(r.ok, 'ping not ok: '+JSON.stringify(r));
    version = r.version;
    console.log('       backend version =', version);
  });

  const HARDENED = version === '2026-05-25-hardened';
  if (!HARDENED) {
    console.log('  \x1b[33m! backend not yet redeployed with hardened Code.gs — these live tests will check current behaviour;\n    rerun after redeploying for full coverage.\x1b[0m');
  }

  await it('XSS payload in name field is accepted but not echoed in error', async () => {
    const xss = '<img src=x onerror=alert(1)>';
    const r = await call('signup', {
      email: 'smoke-test-xss-'+Date.now()+'@example.com',
      password: 'GoodPass123',
      name: xss,
      centre: '<script>alert(1)</script>'
    });
    // Either success (pending) or duplicate-email — never an HTML-injected error
    assert(typeof r === 'object', 'expected JSON');
    assert(!('snippet' in r), 'response should be JSON, got non-JSON');
    if (r.error) assertNoMatch(r.error, /<img|<script/i, 'error must not echo HTML tags');
  });

  await it('signup rejects weak password (too short)', async () => {
    const r = await call('signup', {email:'smoke-test-weak-'+Date.now()+'@example.com', password:'abc', name:'x', centre:''});
    if (HARDENED) assert(r.ok===false && /at least 8/.test(r.error||''), 'expected weak-password rejection; got '+JSON.stringify(r));
    else console.log('       (skipped — old backend allows weak passwords)');
  });

  await it('signup rejects weak password (no digit)', async () => {
    const r = await call('signup', {email:'smoke-test-nodigit-'+Date.now()+'@example.com', password:'onlyletters', name:'x', centre:''});
    if (HARDENED) assert(r.ok===false && /digit/.test(r.error||''), 'expected no-digit rejection; got '+JSON.stringify(r));
    else console.log('       (skipped — old backend allows weak passwords)');
  });

  await it('signup rejects invalid email', async () => {
    const r = await call('signup', {email:'not-an-email', password:'GoodPass123', name:'x', centre:''});
    if (HARDENED) assert(r.ok===false && /valid email/.test(r.error||''), 'expected email rejection; got '+JSON.stringify(r));
    else console.log('       (skipped — old backend may accept this)');
  });

  await it('login fails with generic message for unknown user', async () => {
    const r = await call('login', {username:'nope-'+Date.now()+'@example.com', password:'whatever'});
    assert(r.ok === false, 'expected failure');
    if (HARDENED) assert(/Invalid email or password/.test(r.error||''), 'expected generic message; got '+JSON.stringify(r));
  });

  await it('admin endpoint blocked without token', async () => {
    const r = await call('list_users', { token: '' });
    assert(r.ok === false, 'expected forbidden');
    if (HARDENED) assertMatch(r.error||'', /Not authenticated|Admin access required/);
  });

  await it('admin endpoint blocked with invalid token', async () => {
    const r = await call('list_users', { token: 'definitely-not-a-real-token-aaaaaaaaaa' });
    assert(r.ok === false, 'expected rejection');
    if (HARDENED) assertMatch(r.error||'', /Not authenticated|Admin access required/);
  });

  await it('save_ blocked without token', async () => {
    const r = await call('save', { token: '', data: { state:{}, uid:1 } });
    assert(r.ok === false, 'expected rejection');
  });

  await it('huge payload rejected', async () => {
    const big = 'x'.repeat(300_000);
    const r = await call('signup', { email: big+'@example.com', password:'GoodPass123', name:big, centre:big });
    if (HARDENED) assert(r.ok === false, 'expected rejection of huge payload');
  });

  await it('unknown action returns clean error (no HTML)', async () => {
    const r = await call('definitely_not_a_real_action');
    assert(r.ok === false && /Unknown action/.test(r.error||''), 'expected Unknown action error');
  });

  await it('SQL-injection-shaped username is treated as literal text', async () => {
    const r = await call('login', {username:"' OR 1=1 --", password:'x'});
    // Backend has no SQL — just sheet rows — so this should fail like any unknown user
    assert(r.ok === false);
  });

  // Rate-limit smoke test: 7 wrong-password attempts on a known username
  await it('rate-limit: repeated failures lock account', async () => {
    // Use a unique email for each run so we exercise lockout against a fresh account
    if (!HARDENED) { console.log('       (skipped — needs hardened backend)'); return; }
    const email = 'smoke-rl-'+Date.now()+'@example.com';
    const sup = await call('signup', { email, password:'GoodPass123', name:'rl', centre:'' });
    assert(sup.ok, 'precondition: signup must succeed; got '+JSON.stringify(sup));
    // Even though the account is "pending" approval, login_ checks password FIRST
    // and counts failures BEFORE the pending check.
    let lockedSeen = false;
    for (let i = 0; i < 7; i++) {
      const r = await call('login', { username: email, password: 'WRONG-pw-'+i });
      if (r.code === 'locked') { lockedSeen = true; break; }
    }
    assert(lockedSeen, 'expected lockout response after repeated bad logins');
  });
}

/* =========================================================================
 *  RUN
 * ========================================================================= */
(async () => {
  console.log('Isha Karnataka — security test suite');
  staticTests();
  if (!SKIP_LIVE) await liveTests();
  else section('(live tests skipped via --skip-live)');

  console.log('\n\x1b[1m'+pass+' passed, '+fail+' failed\x1b[0m');
  if (fail) {
    console.log('\nFailures:');
    failures.forEach(([n,e]) => console.log('  - '+n+(e?.message ? ': '+e.message : '')));
    process.exit(1);
  }
})();
