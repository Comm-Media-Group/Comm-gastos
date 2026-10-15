// ─────────────────────────────────────────────────────────────────────────────
//  COMM Gastos — company settings. Edit this file, then redeploy (and bump VERSION in sw.js).
// ─────────────────────────────────────────────────────────────────────────────
window.CMG_CONFIG = {
  // Google OAuth "Web application" client ID. When set, employees with a Gmail account can send the
  // report DIRECTLY from their own Gmail (no copy/paste, no share sheet). Leave empty to use the
  // phone's share sheet instead. See README.md → "Enable one-tap sending from Gmail".
  googleClientId: '',

  // Pre-fills the "Send reports to" field for new users (they can still change it).
  // e.g. 'administracion@commmediagroup.com'. Leave '' to make each employee enter it.
  defaultSendTo: '',

  // Suggestions for the EMPRESA column (employees can type anything).
  companies: ['Comm Media Group', 'Luxurycomm', 'Brandscomm', 'Comm Creative', 'Select Up']
};
