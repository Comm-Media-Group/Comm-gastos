/* Sending: (1) Gmail API straight from the employee's own Gmail account, (2) the phone's share sheet,
   (3) download + mailto fallback. Everything runs in the browser; nothing is stored on a server. */
(function (root) {
  const GIS_SRC = 'https://accounts.google.com/gsi/client';
  const SCOPE = 'https://www.googleapis.com/auth/gmail.send';
  let gisPromise = null, tokenCache = null;

  const utf8b64 = s => {
    const b = new TextEncoder().encode(s); let bin = '';
    for (let i = 0; i < b.length; i += 0x8000) bin += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(bin);
  };
  // RFC 2047: encoded-words must stay under 75 chars, so long non-ASCII subjects are split and folded
const encHeader = s => {
  if (/^[\x20-\x7e]*$/.test(s)) return s;
  const chars = Array.from(s), out = [];
  for (let i = 0; i < chars.length; i += 12) out.push('=?UTF-8?B?' + utf8b64(chars.slice(i, i + 12).join('')) + '?=');
  return out.join('\r\n ');
};
  const asciiName = s => String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9._-]+/g, '_');
  const wrap76 = s => s.replace(/.{1,76}/g, '$&\r\n');
  const blobToB64 = blob => new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(',')[1] || '');
    r.onerror = () => rej(r.error);
    r.readAsDataURL(blob);
  });

  /** Builds a multipart/mixed RFC 822 message as a Blob. */
  async function buildMime({ to, cc, subject, text, attachments }) {
    const boundary = '=_cmg_' + Math.random().toString(36).slice(2) + Date.now().toString(36);
    const head = [
      'MIME-Version: 1.0',
      'To: ' + to,
      cc ? 'Cc: ' + cc : null,
      'Subject: ' + encHeader(subject),
      'Content-Type: multipart/mixed; boundary="' + boundary + '"'
    ].filter(Boolean).join('\r\n') + '\r\n\r\n';
    const parts = [head,
      '--' + boundary + '\r\nContent-Type: text/plain; charset="UTF-8"\r\nContent-Transfer-Encoding: base64\r\n\r\n' + wrap76(utf8b64(text)) + '\r\n'];
    for (const a of attachments) {
      const name = asciiName(a.name);
      parts.push('--' + boundary + '\r\nContent-Type: ' + a.type + '; name="' + name + '"\r\n' +
        'Content-Disposition: attachment; filename="' + name + '"\r\nContent-Transfer-Encoding: base64\r\n\r\n' +
        wrap76(await blobToB64(a.blob)) + '\r\n');
    }
    parts.push('--' + boundary + '--\r\n');
    return new Blob(parts, { type: 'message/rfc822' });
  }

  function loadGis() {
    if (root.google && root.google.accounts && root.google.accounts.oauth2) return Promise.resolve();
    if (!gisPromise) gisPromise = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = GIS_SRC; s.async = true;
      s.onload = () => res(); s.onerror = () => { gisPromise = null; rej(new Error('GIS_LOAD')); };
      document.head.appendChild(s);
    });
    return gisPromise;
  }

  /** Must be called from a user gesture (tap) so the Google consent popup is allowed. */
  async function getToken(clientId, loginHint) {
    if (tokenCache && tokenCache.exp > Date.now() + 60000 && tokenCache.hint === (loginHint || '')) return tokenCache.token;
    await loadGis();
    return new Promise((resolve, reject) => {
      const client = root.google.accounts.oauth2.initTokenClient({
        client_id: clientId, scope: SCOPE, login_hint: loginHint || undefined,
        callback: r => {
          if (r.error) return reject(new Error(r.error_description || r.error));
          tokenCache = { token: r.access_token, exp: Date.now() + (Number(r.expires_in) || 3000) * 1000, hint: loginHint || '' };
          resolve(r.access_token);
        },
        error_callback: e => reject(new Error((e && (e.message || e.type)) || 'oauth_cancelled'))
      });
      client.requestAccessToken({ prompt: tokenCache ? '' : undefined });
    });
  }

  async function gmailSend(token, mimeBlob) {
    const r = await fetch('https://gmail.googleapis.com/upload/gmail/v1/users/me/messages/send?uploadType=media', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'message/rfc822' },
      body: mimeBlob
    });
    if (!r.ok) {
      let msg = r.status + '';
      try { const j = await r.json(); msg = (j.error && j.error.message) || msg; } catch (e) {}
      const err = new Error(msg); err.status = r.status; throw err;
    }
    return r.json();
  }

  /** Share sheet with files (iOS/Android). Resolves true if the user completed it, false if cancelled. */
  async function shareFiles({ files, title, text }) {
    if (!navigator.canShare || !navigator.canShare({ files })) throw new Error('SHARE_UNSUPPORTED');
    try { await navigator.share({ files, title, text }); return true; }
    catch (e) { if (e && e.name === 'AbortError') return false; throw e; }
  }

  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }

  /** Posts the report to the company relay (Apps Script web app). Plain-text body = no CORS preflight. */
  async function relaySend({ url, key, payload }) {
    const ctrl = new AbortController(), timer = setTimeout(() => ctrl.abort(), 150000);
    try {
      const r = await fetch(url, { method: 'POST', body: JSON.stringify(Object.assign({}, payload, { key })), signal: ctrl.signal, redirect: 'follow' });
      const txt = await r.text();
      let j; try { j = JSON.parse(txt); } catch (e) { throw Object.assign(new Error('RELAY_BAD_RESPONSE'), { relay: true }); }
      if (!j.ok) throw Object.assign(new Error(j.error || 'relay error'), { relay: true });
      return j;
    } finally { clearTimeout(timer); }
  }

  root.CmgMail = { relaySend, buildMime, getToken, gmailSend, shareFiles, download, asciiName, blobToB64 };
})(window);
