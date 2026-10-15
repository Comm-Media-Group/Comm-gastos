# COMM Gastos

Mobile expense app for Comm Media Group employees. Photograph a receipt, pick the category, and at month end send
the **Excel "Hoja de gastos"** plus a **ZIP with every receipt photo** to the address each employee has set.
No server, no accounts: everything is stored on the phone and the app works offline.

## What each employee does
1. **Ajustes** (first launch): name, default company, cost approver (name + optional email), the address the reports go to
   (their work email or admin), and their personal email.
2. **+** → take a photo of the receipt → category, amount, date, concept / client / company → save.
3. **Enviar** → choose the month(s) → send. Pending expenses become "Enviado" and are not included again.

The Excel keeps your template's layout (FECHA · 8 category columns · PROYECTO · CLIENTE · EMPRESA, TOTAL row with `SUM` formulas),
one sheet per month, plus a **JUSTIFICANTE** column with the matching file name inside the ZIP.

## Deploy (5 minutes)
Upload this whole folder to any static HTTPS host (Netlify Drop, Cloudflare Pages, GitHub Pages, or the company web server)
and send employees the link. On the phone: iPhone → Safari → Share → *Add to Home Screen*; Android → Chrome → *Install app*.
Every time you change a file, bump `VERSION` in `sw.js` so phones pick up the update.

## How sending works
| Situation | What happens |
|---|---|
| `googleClientId` set **and** personal email is Gmail | **One tap**: sent from the employee's own Gmail account with both attachments, to the preset address (CC to the approver if enabled). |
| Otherwise, phone supports file sharing (iOS / Android) | The share sheet opens with the Excel + ZIP already attached; the employee picks Mail / Outlook / Gmail. The recipient is copied to the clipboard (share sheets cannot pre-fill it). |
| Desktop / fallback | Both files download and the mail app opens with subject and text filled in. |

### Enable one-tap sending from Gmail (one-time, ~10 min)
1. [Google Cloud Console](https://console.cloud.google.com) → new project → **APIs & Services → Library → Gmail API → Enable**.
2. **OAuth consent screen**: User type *External*; add scope `https://www.googleapis.com/auth/gmail.send`.
   While the app is in *Testing*, add each employee's Gmail under **Test users** (max 100). For more people, or to remove the
   "unverified app" warning, submit the app for Google verification (`gmail.send` is a sensitive scope).
3. **Credentials → Create credentials → OAuth client ID → Web application**. Under *Authorized JavaScript origins* add the exact
   URL where the app is hosted (e.g. `https://gastos.commmediagroup.com`).
4. Paste the client ID into `config.js` → `googleClientId`, redeploy, bump `VERSION` in `sw.js`.

The app only requests permission to *send* email (it cannot read the mailbox). Each employee approves once per session.

## Files
`index.html` · `styles.css` · `app.js` (UI + storage) · `xlsx-builder.js` (Excel) · `mail.js` (Gmail / share) · `config.js` (edit me) ·
`sw.js` + `manifest.json` (offline / installable) · `xlsx-assets.js` (styles of your original sheet) · `ejemplo/` (sample output).

## Notes
- Receipts are resized to ≤1700 px JPEG (~50–150 KB). PDFs are kept as they are.
- A single email must stay under 25 MB; the app warns if it would not fit (send month by month).
- Data lives in the phone's browser storage; clearing site data deletes unsent expenses. "Borrar gastos ya enviados" frees space.
