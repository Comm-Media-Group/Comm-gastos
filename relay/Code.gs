/**
 * COMM Gastos — mail relay (Google Apps Script web app)
 *
 * Delivery rules: reports may go to addresses on ALLOWED_RECIPIENT_DOMAINS, or to any address its owner has VERIFIED once with a
 * 6-digit code (so the relay can never be used to e-mail strangers).
 *
 * The app POSTs the expense report (Excel + receipts ZIP, base64) here and this script e-mails it from the Google account
 * that owns the script (use a company mailbox such as gastos@commmediagroup.com). The employee is set as Reply-To.
 *
 * SETUP: see README.md → "One-tap sending through a company mailbox".
 * Paste this whole file into script.google.com, change SECRET, then Deploy → New deployment → Web app
 * (Execute as: Me · Who has access: Anyone).
 */

// ── EDIT THESE ──────────────────────────────────────────────────────────────────────────────────────────────────────
const SECRET = 'CHANGE-ME-to-a-long-random-string';        // must match relayKey in the app's config.js
const ALLOWED_RECIPIENT_DOMAINS = ['commmediagroup.com'];   // To/CC may ONLY go to these domains (stops misuse as a spam relay)
const SENDER_NAME = 'COMM Gastos';
const MAX_EMAILS_PER_HOUR = 40;                             // whole company; raise if you have many employees
const MAX_EMAILS_PER_PERSON_PER_HOUR = 8;                   // one person can't use up everyone's allowance
const MAX_ATTACHMENT_BYTES = 20 * 1024 * 1024;              // Gmail's limit is 25 MB per message
// ────────────────────────────────────────────────────────────────────────────────────────────────────────────────────

function doGet() { return out_({ ok: true, service: 'comm-gastos-relay' }); }

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!d.key || d.key !== SECRET) return out_({ ok: false, error: 'unauthorized' });

    if (d.action === 'verify-start') return out_(verifyStart_(d));
    if (d.action === 'verify-check') return out_(verifyCheck_(d));

    // idempotency: a retry of the same report must not send twice
    const cache = CacheService.getScriptCache();
    if (d.id && cache.get('sent:' + d.id)) return out_({ ok: true, duplicate: true });

    const to = cleanList_(d.to), cc = cleanList_(d.cc);
    if (!to.length) return out_({ ok: false, error: 'no recipient' });
    const bad = to.concat(cc).filter(a => !isAllowed_(a) && !isVerified_(a));
    if (bad.length) return out_({ ok: false, error: 'recipient not allowed: ' + bad.join(', ') });
    if (to.length + cc.length > 5) return out_({ ok: false, error: 'too many recipients' });

    const replyTo = isEmail_(d.replyTo) ? String(d.replyTo).trim() : '';
    const subject = String(d.subject || '').replace(/[\r\n]+/g, ' ').slice(0, 250);
    const body = String(d.body || '').slice(0, 20000);
    if (!subject) return out_({ ok: false, error: 'no subject' });

    let total = 0;
    const attachments = (d.attachments || []).slice(0, 10).map(a => {
      const bytes = Utilities.base64Decode(a.b64 || '');
      total += bytes.length;
      return Utilities.newBlob(bytes, String(a.type || 'application/octet-stream'), String(a.name || 'file').slice(0, 120));
    });
    if (total > MAX_ATTACHMENT_BYTES) return out_({ ok: false, error: 'attachments too large' });

    if (!takeSlot_(String(d.replyTo || d.employee || 'anon').toLowerCase())) return out_({ ok: false, error: 'rate limit, try again later' });

    const msg = { to: to.join(','), subject: subject, body: body, name: SENDER_NAME, attachments: attachments };
    if (cc.length) msg.cc = cc.join(',');
    if (replyTo) msg.replyTo = replyTo;
    MailApp.sendEmail(msg);

    if (d.id) cache.put('sent:' + d.id, '1', 6 * 3600);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

/** Run this once from the editor (▶) to grant the script permission to send mail. */
function authorize() { MailApp.getRemainingDailyQuota(); Logger.log('Authorized. Remaining daily quota: ' + MailApp.getRemainingDailyQuota()); }

/** Optional: run from the editor to send yourself a test message. */
function testEmail() {
  MailApp.sendEmail({ to: Session.getActiveUser().getEmail(), subject: 'COMM Gastos relay test', body: 'The relay can send mail.', name: SENDER_NAME });
}

// ── address verification (one-time, lets employees receive reports on a personal address) ───────────────────────────
function verifyStart_(d) {
  const a = String(d.address || '').trim().toLowerCase();
  if (!isEmail_(a)) return { ok: false, error: 'invalid address' };
  if (isAllowed_(a) || isVerified_(a)) return { ok: true, alreadyAllowed: true };
  const cache = CacheService.getScriptCache(), hour = Math.floor(Date.now() / 3600000), kTry = 'vs:' + hour + ':' + hash_(a);
  const tries = Number(cache.get(kTry) || 0);
  if (tries >= 3) return { ok: false, error: 'too many codes requested, try again later' };
  if (!takeSlot_('verify:' + a)) return { ok: false, error: 'rate limit, try again later' };
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('vc:' + hash_(a), JSON.stringify({ code: code, tries: 0 }), 900);
  cache.put(kTry, String(tries + 1), 3700);
  MailApp.sendEmail({ to: a, name: SENDER_NAME, subject: 'Tu código de verificación de COMM Gastos',
    body: 'Tu código es: ' + code + '\n\nCaduca en 15 minutos. Si no lo has pedido tú, ignora este mensaje.' });
  return { ok: true };
}
function verifyCheck_(d) {
  const a = String(d.address || '').trim().toLowerCase(), cache = CacheService.getScriptCache(), k = 'vc:' + hash_(a);
  const raw = cache.get(k);
  if (!raw) return { ok: false, error: 'code expired' };
  const o = JSON.parse(raw);
  if (o.tries >= 5) { cache.remove(k); return { ok: false, error: 'too many attempts' }; }
  if (String(d.code || '').trim() !== o.code) { o.tries++; cache.put(k, JSON.stringify(o), 900); return { ok: false, error: 'wrong code' }; }
  PropertiesService.getScriptProperties().setProperty('ver:' + a, String(Date.now()));
  cache.remove(k);
  return { ok: true };
}
function isVerified_(addr) { return !!PropertiesService.getScriptProperties().getProperty('ver:' + String(addr).trim().toLowerCase()); }
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s))).replace(/[^A-Za-z0-9]/g, '').slice(0, 16); }

// ── helpers ──────────────────────────────────────────────────────────────────────────────────────────────────────────
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function isEmail_(s) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]{2,}$/.test(String(s || '').trim()); }
function cleanList_(v) { return (Array.isArray(v) ? v : String(v || '').split(/[,;]/)).map(s => String(s).trim()).filter(Boolean).filter(isEmail_); }
function isAllowed_(addr) {
  const dom = String(addr).split('@').pop().toLowerCase();
  return ALLOWED_RECIPIENT_DOMAINS.some(d => dom === d.toLowerCase());
}
function takeSlot_(who) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache(), hour = Math.floor(Date.now() / 3600000);
    const kAll = 'hour:' + hour, kMe = 'who:' + hour + ':' + hash_(who);
    const nAll = Number(cache.get(kAll) || 0), nMe = Number(cache.get(kMe) || 0);
    if (nAll >= MAX_EMAILS_PER_HOUR || nMe >= MAX_EMAILS_PER_PERSON_PER_HOUR) return false;
    cache.put(kAll, String(nAll + 1), 3700); cache.put(kMe, String(nMe + 1), 3700);
    return true;
  } finally { lock.releaseLock(); }
}
