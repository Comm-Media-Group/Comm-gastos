# COMM Gastos

Mobile expense app for Comm Media Group employees. Photograph a receipt, pick the category, and at month end send
the **Excel "Hoja de gastos"** plus a **ZIP with every receipt photo** to the address each employee has set.
No server, no accounts: everything is stored on the phone and the app works offline.

## Rolling it out (you configure once, employees configure nothing)
**You, once:** set up the relay (below) and edit `config.js`: `relayUrl` + `relayKey`, `defaultSendTo` (e.g. the admin mailbox),
optionally `lockSendTo: true` (everyone's reports go to that address) and `approvers` (a pick-list of cost approvers, so people choose
instead of typing and the approver is copied automatically). Upload, then send the link.

**Message to employees (copy/paste):**
> Hola 👋 Ya tenemos la app de gastos: https://comm-media-group.github.io/Comm-gastos/
> 1) Ábrela en el móvil y añádela a la pantalla de inicio (iPhone: Safari → Compartir → «Añadir a pantalla de inicio»; Android: Chrome → «Instalar app»).
> 2) En **Ajustes** pon tu nombre, tu responsable y a dónde quieres que se envíe (administración o tu correo; si es personal te pedirá un código una sola vez). Nada más.
> 3) Haz fotos de tus tickets con **Foto rápida** (o el botón +) y completa el importe cuando puedas.
> 4) A fin de mes: **Informes → tu informe → Enviar ahora**. El Excel y los justificantes se envían solos a administración.
> 5) Cuando te hayan pagado el informe, pulsa **Reembolsado** y se borra del móvil.
> Tip: haz una **copia de seguridad** (Ajustes) de vez en cuando por si cambias de móvil.

## What each employee does
1. **Ajustes** (first launch): name, default company, cost approver (name + optional email), the address the reports go to
   (their work email or admin), and their personal email.
2. **+** → add one or more photos / PDFs of the receipt → category, amount, date, concept / client / company → save.
   - **Photo first, details later:** on Home, **Foto rápida** (camera) or **Importar** (gallery / Files, several at once) saves each
     picture straight into **Por completar** (shown on Home with a badge on the Inicio tab). Tap one to add the details.
     Unfinished items are never sent.
   - **Invoice arrived by email?** Android: open it in Gmail → *Share* → **COMM Gastos** (it lands in *Por completar*).
     iPhone (Apple does not let web apps appear in the share sheet): in Mail/Gmail tap the attachment → *Save to Files*, then
     in the app use **Importar** / **Galería / archivos**; or take a screenshot of the invoice and import that.
3. **Informes**: a permanent **General** report plus any report you create ("Noviembre 2026", "Viaje Valladolid"…).
   Every expense is filed into one report (dropdown at the top of the expense form). Filed it in General by mistake or later
   want it in the monthly report? Open **Gastos → Seleccionar**, tick the expenses, and **Mover** them to a report
   (or create a new one with the selection).
4. Open a report → **Enviar**. Each report goes out as its own Excel + ZIP; sent expenses become "Enviado" and are not included again.
5. Once finance has paid the report, open it and tap **Reembolsado**: the *sent* expenses and their receipt photos are deleted from the
   phone (unsent expenses are never touched) and trip reports disappear once empty.

The Excel keeps your template's layout (FECHA · 8 category columns · PROYECTO · CLIENTE · EMPRESA, TOTAL row with `SUM` formulas),
one sheet per month, plus a **JUSTIFICANTE** column with the matching file name inside the ZIP.

## Forcing an update
Normally phones pick up changes on their own (after you bump `VERSION` in `sw.js`). To force it: **Ajustes → Buscar actualizaciones**, or open
`https://<your-site>/?refresh`. Both clear only the app-file cache; expenses, photos and settings are untouched.
The version number is shown at the bottom of Ajustes.

## Backup & restore (new phone, lost phone)
**Ajustes → Copia de seguridad → Crear copia** makes one ZIP with every expense, report, receipt and your settings (it opens the
share sheet: save it to iCloud Drive / Google Drive / email it to yourself). On the new phone open the app and tap **Restaurar**
and pick that ZIP. Home reminds people when the last backup is older than 30 days.

## Deploy (5 minutes)
Upload this whole folder to any static HTTPS host (Netlify Drop, Cloudflare Pages, GitHub Pages, or the company web server)
and send employees the link. On the phone: iPhone → Safari → Share → *Add to Home Screen*; Android → Chrome → *Install app*.
Every time you change a file, bump `VERSION` in `sw.js` so phones pick up the update.

## How sending works
| Situation | What happens |
|---|---|
| `relayUrl` set **and** the employee's "Enviar a" address is on `relayDomains` (**recommended**) | **One tap, nothing to attach**: the Excel + ZIP are sent automatically from the company mailbox to the preset address (CC approver if enabled), with the employee as Reply-To. Works for any employee regardless of their email provider. |
| `googleClientId` set **and** personal email is Gmail | **One tap**: sent from the employee's own Gmail account with both attachments, to the preset address (CC to the approver if enabled). |
| Otherwise, phone supports file sharing (iOS / Android) | The share sheet opens with the Excel + ZIP already attached; the employee picks Mail / Outlook / Gmail. The recipient is copied to the clipboard (share sheets cannot pre-fill it). |
| Desktop / fallback | Both files download and the mail app opens with subject and text filled in. |

### One-tap sending through a company mailbox (recommended, ~10 min, once)
1. Sign in to Google **as the account that should appear as sender** (e.g. a shared mailbox like `gastos@commmediagroup.com`).
2. Go to <https://script.google.com> → **New project** → delete the sample code → paste the whole of `relay/Code.gs`.
3. Change `SECRET` to a long random string and check `ALLOWED_RECIPIENT_DOMAINS` (recipients outside these domains are refused).
4. Select the function `authorize` in the toolbar → **Run** → accept the permission prompts (only needed once).
5. **Deploy → New deployment → type: Web app** → *Execute as*: **Me**, *Who has access*: **Anyone** → Deploy → copy the **Web app URL**.
   (If *Anyone* is greyed out, your Google Workspace admin must allow it.)
6. In this repo's `config.js` set `relayUrl` (the URL) and `relayKey` (the same SECRET), commit, and bump `VERSION` in `sw.js`.
7. Open the app, complete a report and tap **Enviar ahora**. The email arrives at the preset address; replying goes to the employee.

Employees can send **wherever they want**, always with one tap once set up:
- a **company address** (admin, their own work email) works immediately;
- a **personal address** (their own Gmail…) is verified **once** in Ajustes: the app asks the relay to e-mail a 6-digit code to that address,
  the employee types it in, and from then on that address is one tap too. The relay only delivers to company addresses or addresses whose
  owner has proven they own them, so it can never be used to e-mail strangers (and sends at most 3 codes/hour per address).
- until an address is verified (or if the relay is down) the app falls back to the phone's share sheet with the files attached.
The "Enviar a" address is entered once in Ajustes and stays saved; each report shows it with a **Cambiar** link. Set `relayDomains` in
`config.js` to the same domains as `ALLOWED_RECIPIENT_DOMAINS` in `Code.gs`. The approver is typed by each employee (the optional
`approvers` pick-list in `config.js` is empty by default).

**Updating the relay after changing `Code.gs`:** Deploy → Manage deployments → ✏️ → Version: *New version* → Deploy (the URL stays the same).

Good to know: the key sits in the public `config.js`, so anyone who reads it could call the relay. That is why the script only sends to your
own domain, caps sends per hour, ignores duplicates and limits size; if you ever see misuse, change `SECRET` (script + `config.js`) and redeploy.
Quota: about 1,500 recipients/day with Google Workspace (100/day with a free Gmail account). After editing `Code.gs` use
*Deploy → Manage deployments → ✏️ → Version: New version → Deploy* so the live copy updates. If the relay is unreachable the app
still offers **Compartir** / **Descargar** as fallbacks.

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
