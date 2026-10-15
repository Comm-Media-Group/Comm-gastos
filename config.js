// ─────────────────────────────────────────────────────────────────────────────
//  COMM Gastos — company settings. Edit this file, then redeploy (and bump VERSION in sw.js).
// ─────────────────────────────────────────────────────────────────────────────
window.CMG_CONFIG = {
  // Google OAuth "Web application" client ID. When set, employees with a Gmail account can send the
  // report DIRECTLY from their own Gmail (no copy/paste, no share sheet). Leave empty to use the
  // phone's share sheet instead. See README.md → "Enable one-tap sending from Gmail".
  googleClientId: '',

  // ONE-TAP SENDING through a company mailbox (recommended). Deploy relay/Code.gs as an Apps Script web app (see README.md)
  // and paste its URL + the same secret here. Employees then just tap "Enviar ahora": Excel + ZIP are attached and sent
  // automatically from the company mailbox, with the employee as Reply-To. Leave '' to use the other sending methods.
  relayUrl: 'https://script.google.com/macros/s/AKfycbzlXNVP6TmoazBqX3fcsUG5yjs2dPFDpvatCp18gJqm3zpmskc66wDjOhFOIlG0ljof/exec',
  relayKey: '1b805001c96a5a1d0acffa9e21e859ce77be21d05687d18f',
  // Domains the relay may deliver to (must match ALLOWED_RECIPIENT_DOMAINS in relay/Code.gs). If an employee's "Send to" address is
  // on one of these, the app sends with one tap; for any other address (their own Gmail, a client…) it opens the phone's share
  // sheet with the files attached instead, so everybody can still send wherever they want.
  relayDomains: ['commmediagroup.com'],

  // Cost approvers employees can PICK from a list (saves typing, and fills the CC address). Leave [] to let them type a name.
  // e.g. approvers: [{ name: 'Pablo Campos', email: 'pablo@commmediagroup.com' }, { name: 'Laura Gómez', email: 'laura@commmediagroup.com' }],
  approvers: [],

  // true = employees cannot change the destination: everyone's reports go to defaultSendTo (below).
  lockSendTo: false,

  // Pre-fills the "Send reports to" field for new users (they can still change it).
  // e.g. 'administracion@commmediagroup.com'. Leave '' to make each employee enter it.
  defaultSendTo: '',

  // Suggestions for the EMPRESA column (employees can type anything).
  companies: ['Comm Media Group', 'Luxurycomm', 'Brandscomm', 'Comm Creative', 'Select Up']
};
