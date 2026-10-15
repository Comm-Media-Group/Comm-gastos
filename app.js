/* COMM Gastos — expense capture + monthly "Hoja de gastos" delivery. Vanilla JS, offline-first PWA. */
(() => {
'use strict';
const CFG = window.CMG_CONFIG || {};
const VERSION = '1.0.0';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const p2 = n => String(n).padStart(2, '0');
const today = () => { const d = new Date(); return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()); };
const isEmail = s => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(s || '').trim());
const isGmail = s => /@(gmail|googlemail)\.com$/i.test(String(s || '').trim());
const r2 = n => Math.round((Number(n) || 0) * 100) / 100;

/* ───────────── i18n ───────────── */
const TXT = {
  es: {
    home: 'Inicio', expenses: 'Gastos', send: 'Enviar', settings: 'Ajustes',
    pending: 'Pendiente de reembolso', nExp: 'Gastos', noReceipt: 'Sin justificante', thisMonth: 'Este mes',
    quick: 'Añadir rápido', recent: 'Últimos gastos', emptyT: 'Aún no hay gastos', emptyS: 'Pulsa + y haz una foto del justificante.',
    stPending: 'Pendientes', stSent: 'Enviados', stAll: 'Todos', all: 'Todas',
    sent: 'Enviado', noRec: 'Sin justificante',
    addExp: 'Nuevo gasto', editExp: 'Editar gasto', receipt: 'Justificante', takePhoto: 'Hacer foto', pickFile: 'Elegir archivo',
    removeRec: 'Quitar justificante', category: 'Categoría', amount: 'Importe (€)', date: 'Fecha', concept: 'Concepto / proyecto',
    conceptPh: 'p. ej. Taxi Hermes', client: 'Cliente', company: 'Empresa', save: 'Guardar gasto', del: 'Eliminar',
    delAsk: '¿Eliminar este gasto?', saved: 'Gasto guardado', deleted: 'Gasto eliminado', needAmount: 'Introduce un importe', needDate: 'Elige una fecha',
    wasSent: 'Enviado el', backPending: 'Volver a pendiente',
    sec_profile: 'Perfil', sec_approval: 'Aprobación', sec_delivery: 'Envío', sec_app: 'Aplicación',
    fullName: 'Nombre y apellidos', defCompany: 'Empresa por defecto', approver: 'Responsable de aprobación', approverEmail: 'Email del responsable (opcional)',
    ccApprover: 'Enviar copia (CC) al responsable', sendTo: 'Enviar las hojas a', sendToHint: 'Tu correo de trabajo o el de administración. Aquí llegará el Excel con el ZIP de justificantes.',
    personal: 'Mi correo personal (desde el que envío)', personalHint: 'Con Gmail el envío sale directamente de tu cuenta. Con otro correo se usa la hoja de compartir del móvil.',
    language: 'Idioma', saveSettings: 'Guardar ajustes', settingsSaved: 'Ajustes guardados', fixFields: 'Revisa los campos marcados',
    welcomeT: 'Bienvenida/o', welcomeS: 'Antes de empezar, rellena tus datos. Solo se guardan en este teléfono.',
    cleanSent: 'Borrar gastos ya enviados', cleanAsk: '¿Borrar los gastos enviados y sus fotos de este teléfono?', cleaned: 'Gastos enviados borrados',
    storageFull: '⚠ Almacenamiento del teléfono lleno', version: 'Versión',
    installIos: 'Instala la app: pulsa Compartir y después «Añadir a pantalla de inicio».', dismiss: 'Cerrar',
    sec_summary: 'Resumen', sec_period: 'Periodo', sec_content: 'Contenido', from: 'De', to: 'Para', cc: 'CC', approverL: 'Responsable', requester: 'Solicitante',
    notSet: 'Sin configurar', setupFirst: 'Completa tu perfil (nombre, responsable y destinatario) para poder enviar.', goSettings: 'Ir a ajustes',
    noPendT: 'Todo enviado', noPendS: 'No hay gastos pendientes de enviar.',
    dateC: 'Fecha', conceptC: 'Concepto', total: 'Total',
    withoutRec: n => n + (n === 1 ? ' gasto sin justificante' : ' gastos sin justificante'), attach: (n, mb) => n + ' justificantes · ' + mb + ' MB',
    sendNow: 'Enviar ahora', sendGmail: 'Enviar desde mi Gmail', shareMail: 'Compartir por correo', dlMail: 'Descargar y abrir correo',
    orUse: 'O también', viaGmail: 'Gmail', viaShare: 'Compartir…', viaDl: 'Descargar',
    building: 'Preparando archivos…', sending: 'Enviando…', sentOk: 'Enviado correctamente ✓', recCopied: 'Destinatario copiado',
    sentAsk: '¿Has enviado el correo? Márcalo para no volver a incluir estos gastos.', markedSent: 'Marcados como enviados',
    tooBig: 'El correo supera 25 MB. Envía por meses o con menos fotos.', gErr: 'No se pudo enviar con Gmail: ', gDenied: 'Permiso de Gmail cancelado',
    downloaded: 'Archivos descargados — adjúntalos al correo', history: 'Envíos anteriores', hist: (d, n, t) => d + ' · ' + n + ' gastos · ' + t,
    months: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
    cats: { restauracion: 'Restauración', taxi: 'Taxi', vtc: 'VTC', hotel: 'Hotel', alquiler: 'Alquiler coche', vuelos: 'Vuelos', atrezzo: 'Atrezzo', otros: 'Otros' },
    mailSubject: (per, name) => 'Hoja de gastos ' + per + ' – ' + name,
    mailBody: (o) => 'Hola,\n\nAdjunto la hoja de gastos de ' + o.name + ' correspondiente a ' + o.per + ', junto con el ZIP de justificantes.\n\n' +
      'Resumen\n- Gastos: ' + o.n + '\n- Total a reembolsar: ' + o.total + '\n- Justificantes adjuntos: ' + o.rec + ' de ' + o.n + '\n- Responsable de aprobación: ' + o.approver + '\n\nUn saludo,\n' + o.name + '\n\n— Enviado desde COMM Gastos'
  },
  en: {
    home: 'Home', expenses: 'Expenses', send: 'Send', settings: 'Settings',
    pending: 'Pending reimbursement', nExp: 'Expenses', noReceipt: 'No receipt', thisMonth: 'This month',
    quick: 'Quick add', recent: 'Latest expenses', emptyT: 'No expenses yet', emptyS: 'Tap + and take a photo of the receipt.',
    stPending: 'Pending', stSent: 'Sent', stAll: 'All', all: 'All',
    sent: 'Sent', noRec: 'No receipt',
    addExp: 'New expense', editExp: 'Edit expense', receipt: 'Receipt', takePhoto: 'Take photo', pickFile: 'Choose file',
    removeRec: 'Remove receipt', category: 'Category', amount: 'Amount (€)', date: 'Date', concept: 'Description / project',
    conceptPh: 'e.g. Taxi to Hermes', client: 'Client', company: 'Company', save: 'Save expense', del: 'Delete',
    delAsk: 'Delete this expense?', saved: 'Expense saved', deleted: 'Expense deleted', needAmount: 'Enter an amount', needDate: 'Pick a date',
    wasSent: 'Sent on', backPending: 'Mark as pending again',
    sec_profile: 'Profile', sec_approval: 'Approval', sec_delivery: 'Delivery', sec_app: 'App',
    fullName: 'Full name', defCompany: 'Default company', approver: 'Cost approver', approverEmail: 'Approver email (optional)',
    ccApprover: 'CC the approver on every report', sendTo: 'Send reports to', sendToHint: 'Your work email or the admin team’s. The Excel and the receipts ZIP arrive here.',
    personal: 'My personal email (sent from)', personalHint: 'With Gmail the message goes straight from your account. Other providers use your phone’s share sheet.',
    language: 'Language', saveSettings: 'Save settings', settingsSaved: 'Settings saved', fixFields: 'Please check the highlighted fields',
    welcomeT: 'Welcome', welcomeS: 'Fill in your details before you start. They are only stored on this phone.',
    cleanSent: 'Delete already-sent expenses', cleanAsk: 'Delete sent expenses and their photos from this phone?', cleaned: 'Sent expenses deleted',
    storageFull: '⚠ Phone storage is full', version: 'Version',
    installIos: 'Install the app: tap Share, then “Add to Home Screen”.', dismiss: 'Dismiss',
    sec_summary: 'Summary', sec_period: 'Period', sec_content: 'Contents', from: 'From', to: 'To', cc: 'CC', approverL: 'Approver', requester: 'Requester',
    notSet: 'Not set', setupFirst: 'Complete your profile (name, approver and recipient) to send reports.', goSettings: 'Go to settings',
    noPendT: 'All sent', noPendS: 'There are no pending expenses to send.',
    dateC: 'Date', conceptC: 'Description', total: 'Total',
    withoutRec: n => n + (n === 1 ? ' expense without receipt' : ' expenses without receipt'), attach: (n, mb) => n + ' receipts · ' + mb + ' MB',
    sendNow: 'Send now', sendGmail: 'Send from my Gmail', shareMail: 'Share by email', dlMail: 'Download & open email',
    orUse: 'Or use', viaGmail: 'Gmail', viaShare: 'Share…', viaDl: 'Download',
    building: 'Preparing files…', sending: 'Sending…', sentOk: 'Sent successfully ✓', recCopied: 'Recipient copied',
    sentAsk: 'Did you send the email? Mark these expenses as sent so they are not included again.', markedSent: 'Marked as sent',
    tooBig: 'The email exceeds 25 MB. Send month by month or with fewer photos.', gErr: 'Could not send with Gmail: ', gDenied: 'Gmail permission cancelled',
    downloaded: 'Files downloaded — attach them to the email', history: 'Previous reports', hist: (d, n, t) => d + ' · ' + n + ' expenses · ' + t,
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    cats: { restauracion: 'Meals', taxi: 'Taxi', vtc: 'Ride-hailing', hotel: 'Hotel', alquiler: 'Car rental', vuelos: 'Flights', atrezzo: 'Props', otros: 'Other' },
    mailSubject: null, mailBody: null // emails always go out in Spanish
  }
};

/* ───────────── state ───────────── */
const LS_STATE = 'cmg.state.v1', LS_SET = 'cmg.settings.v1', LS_FLAGS = 'cmg.flags.v1';
const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };
let settings = Object.assign({ name: '', company: '', approver: '', approverEmail: '', ccApprover: false,
  sendTo: CFG.defaultSendTo || '', personalEmail: '', lang: 'es' }, lsGet(LS_SET, {}));
let st = Object.assign({ expenses: [], batches: [], nextId: 1 }, lsGet(LS_STATE, {}));
let flags = lsGet(LS_FLAGS, {});
let rev = 0;
const ui = { page: 'home', status: 'pending', cat: 'all', months: null, draft: null, pkg: null, busy: false };
const L = () => TXT[settings.lang] || TXT.es;
const t = (k, ...a) => { const v = L()[k]; return typeof v === 'function' ? v(...a) : (v == null ? k : v); };
const catLabel = (k, lang) => (TXT[lang || settings.lang].cats[k] || k);

function persist() {
  rev++;
  if (!lsSet(LS_STATE, st)) toast(t('storageFull'), true);
  schedulePrepare();
}
function saveSettingsLS() { rev++; lsSet(LS_SET, settings); }

/* ───────────── receipts in IndexedDB (photos stay out of localStorage's 5 MB cap) ───────────── */
const idb = (() => {
  let dbp = null;
  const open = () => dbp || (dbp = new Promise((res, rej) => {
    const rq = indexedDB.open('cmg-gastos', 1);
    rq.onupgradeneeded = () => rq.result.createObjectStore('receipts');
    rq.onsuccess = () => res(rq.result); rq.onerror = () => rej(rq.error);
  }));
  const tx = async (mode, fn) => { const db = await open(); return new Promise((res, rej) => { const r = fn(db.transaction('receipts', mode).objectStore('receipts')); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); };
  return { put: (id, blob) => tx('readwrite', s => s.put(blob, id)), get: id => tx('readonly', s => s.get(id)), del: id => tx('readwrite', s => s.delete(id)) };
})();
const urls = new Map();
async function receiptUrl(id) {
  if (urls.has(id)) return urls.get(id);
  const b = await idb.get(id).catch(() => null);
  if (!b) return null;
  const u = URL.createObjectURL(b); urls.set(id, u); return u;
}
function dropUrl(id) { if (urls.has(id)) { URL.revokeObjectURL(urls.get(id)); urls.delete(id); } }

async function compress(file) {
  if (file.type === 'application/pdf') return { blob: file, type: 'application/pdf' };
  for (const opts of [{ imageOrientation: 'from-image' }, undefined]) {
    try {
      const bmp = await (opts ? createImageBitmap(file, opts) : createImageBitmap(file));
      const k = Math.min(1, 1700 / Math.max(bmp.width, bmp.height));
      const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(bmp.width * k)); c.height = Math.max(1, Math.round(bmp.height * k));
      const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(bmp, 0, 0, c.width, c.height);
      if (bmp.close) bmp.close();
      const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.8));
      if (blob) return { blob, type: 'image/jpeg' };
    } catch (e) { /* try next */ }
  }
  return { blob: file, type: file.type || 'image/jpeg' };
}
const extFor = type => ({ 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/heic': 'heic', 'image/heif': 'heic', 'application/pdf': 'pdf' }[type] || 'jpg');

/* ───────────── icons ───────────── */
const ICON = {
  restauracion: '<path d="M7 3v7a2 2 0 004 0V3M9 12v9M17 21V3c-2.6 1.6-3.6 4.2-3.6 7.6V14H17"/>',
  taxi: '<path d="M9.5 6V4h5v2M4 17l1.6-5.2A2 2 0 017.5 10.4h9a2 2 0 011.9 1.4L20 17M3.5 17h17v3h-17zM7 20v1.5M17 20v1.5"/><circle cx="7.5" cy="14.3" r=".4"/><circle cx="16.5" cy="14.3" r=".4"/>',
  vtc: '<rect x="7" y="2.5" width="10" height="19" rx="2.2"/><path d="M10.5 18.5h3M10 6h4"/>',
  hotel: '<path d="M3 19V5M3 15h18v4M21 15v-3a3 3 0 00-3-3h-7v6"/><circle cx="7" cy="11" r="1.7"/>',
  alquiler: '<circle cx="8" cy="15.5" r="4"/><path d="M11 12.5L20 3.5M16.5 7l3 3M14 9.5l2 2"/>',
  vuelos: '<path d="M17.8 19.2L16 11l3.5-3.5C21 6 21.5 4 21 3c-1-.5-3 0-4.5 1.5L13 8 4.8 6.2c-.5-.1-.9.1-1.1.5l-.3.5c-.2.5-.1 1 .3 1.3L9 12l-2 3H4l-1 1 3 2 2 3 1-1v-3l3-2 3.5 5.3c.3.4.8.5 1.3.3l.5-.2c.4-.3.6-.7.5-1.2z"/>',
  atrezzo: '<path d="M4 9h16v11H4zM4 9l1.6-4.6 3.4.9L7.6 9.6M10 5.6l3.4.9L12 10M15.2 6.9l3.4.9L17.2 11"/>',
  otros: '<path d="M21 8l-9-5-9 5v8l9 5 9-5zM3 8l9 5 9-5M12 13v8"/>'
};
const UI_ICON = {
  home: '<path d="M3 11.5L12 4l9 7.5M5.5 10v9.5h13V10"/>',
  expenses: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h4"/>',
  send: '<path d="M21 3L10 14M21 3l-7 18-4-7-7-4z"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 00.3 1.8l.1.1a2 2 0 11-2.8 2.8l-.1-.1a1.7 1.7 0 00-1.8-.3 1.7 1.7 0 00-1 1.5V21a2 2 0 01-4 0v-.1a1.7 1.7 0 00-1.1-1.5 1.7 1.7 0 00-1.8.3l-.1.1a2 2 0 11-2.8-2.8l.1-.1a1.7 1.7 0 00.3-1.8 1.7 1.7 0 00-1.5-1H3a2 2 0 010-4h.1a1.7 1.7 0 001.5-1.1 1.7 1.7 0 00-.3-1.8l-.1-.1a2 2 0 112.8-2.8l.1.1a1.7 1.7 0 001.8.3H9a1.7 1.7 0 001-1.5V3a2 2 0 014 0v.1a1.7 1.7 0 001 1.5 1.7 1.7 0 001.8-.3l.1-.1a2 2 0 112.8 2.8l-.1.1a1.7 1.7 0 00-.3 1.8V9a1.7 1.7 0 001.5 1H21a2 2 0 010 4h-.1a1.7 1.7 0 00-1.5 1z"/>',
  clip: '<path d="M21 11.5l-8.6 8.6a5 5 0 01-7.1-7.1l8.6-8.6a3.3 3.3 0 014.7 4.7l-8.6 8.6a1.7 1.7 0 01-2.4-2.4l7.9-7.9"/>',
  camera: '<path d="M4 8h3l1.5-2.5h7L17 8h3v11H4z"/><circle cx="12" cy="13.2" r="3.4"/>',
  file: '<path d="M14 3H6v18h12V7z"/><path d="M14 3v4h4"/>',
  mail: '<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/>',
  share: '<path d="M12 3v12M8 7l4-4 4 4M5 12v8h14v-8"/>',
  down: '<path d="M12 4v11M8 11l4 4 4-4M5 20h14"/>'
};
const svg = (p, cls = 'i') => `<svg class="${cls}" viewBox="0 0 24 24" aria-hidden="true">${p}</svg>`;
const CAT_KEYS = ['restauracion', 'taxi', 'vtc', 'hotel', 'alquiler', 'vuelos', 'atrezzo', 'otros'];

/* ───────────── formatting ───────────── */
const eur = n => new Intl.NumberFormat(settings.lang === 'en' ? 'en-IE' : 'es-ES', { style: 'currency', currency: 'EUR' }).format(Number(n) || 0);
const ymd = s => { const [y, m, d] = s.split('-').map(Number); return { y, m, d }; };
const dShort = s => { const { y, m, d } = ymd(s); return new Date(y, m - 1, d).toLocaleDateString(settings.lang === 'en' ? 'en-GB' : 'es-ES', { day: 'numeric', month: 'short' }); };
const monthName = (y, m) => t('months')[m - 1] + ' ' + y;
const monthNameEs = (y, m) => TXT.es.months[m - 1] + ' ' + y;
const parseAmt = v => { const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.')); return isFinite(n) ? n : NaN; };
const mb = bytes => (bytes / 1048576).toFixed(1);

let toastT;
function toast(msg, err) { const el = $('#toast'); el.textContent = msg; el.classList.toggle('err', !!err); el.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => el.classList.remove('show'), err ? 4200 : 2200); }

/* ───────────── profile ───────────── */
const profileOK = () => settings.name.trim() && settings.approver.trim() && isEmail(settings.sendTo);
const exp = id => st.expenses.find(e => e.id === id);
const sortDesc = (a, b) => b.date.localeCompare(a.date) || b.id - a.id;

/* ───────────── pages ───────────── */
function expRow(e) {
  const label = catLabel(e.cat);
  const sub = [label, e.client].filter(Boolean).join(' · ');
  const flag = e.status === 'sent' ? `<span class="tag ok">${esc(t('sent'))}</span>` : (!e.receipt ? `<span class="tag warn">${esc(t('noRec'))}</span>` : '');
  return `<button class="row" data-act="edit" data-id="${e.id}">
    <span class="ico">${svg(ICON[e.cat] || ICON.otros)}</span>
    <span class="row-main"><span class="row-t" style="display:block">${esc(e.concept || label)}</span><span class="row-s" style="display:block">${esc(sub)} ${e.receipt ? svg(UI_ICON.clip, 'i" style="width:12px;height:12px;vertical-align:-1px') : ''}</span></span>
    <span class="row-r"><span class="row-amt num" style="display:block">${eur(e.amount)}</span><span class="row-d" style="display:block">${flag || esc(dShort(e.date))}</span></span>
  </button>`;
}

function pageHome() {
  const pend = st.expenses.filter(e => e.status === 'pending');
  const total = pend.reduce((s, e) => s + e.amount, 0);
  const cm = today().slice(0, 7);
  const month = st.expenses.filter(e => e.date.startsWith(cm)).reduce((s, e) => s + e.amount, 0);
  const recent = [...st.expenses].sort((a, b) => b.id - a.id).slice(0, 5);
  const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) && !navigator.standalone && !flags.iosHint;
  return `<div class="page">
    ${ios ? `<div class="notice">${esc(t('installIos'))} <button data-act="hide-ios" class="muted" style="text-decoration:underline;margin-left:6px">${esc(t('dismiss'))}</button></div>` : ''}
    <div class="hero">
      <div class="hero-label">${esc(t('pending'))}</div>
      <div class="h-xl hero-amount num">${eur(total)}</div>
      <div class="hero-row">
        <div class="hero-stat"><b>${pend.length}</b><span>${esc(t('nExp'))}</span></div>
        <div class="hero-stat"><b>${pend.filter(e => !e.receipt).length}</b><span>${esc(t('noReceipt'))}</span></div>
        <div class="hero-stat"><b>${Math.round(month)}€</b><span>${esc(t('thisMonth'))}</span></div>
      </div>
    </div>
    <div class="sec"><b>01</b> ${esc(t('quick'))}</div>
    <div class="tiles">${CAT_KEYS.map(k => `<button class="tile" data-act="add" data-cat="${k}">${svg(ICON[k])}<span>${esc(catLabel(k))}</span></button>`).join('')}</div>
    <div class="sec"><b>02</b> ${esc(t('recent'))}</div>
    ${recent.length ? `<div class="card">${recent.map(expRow).join('')}</div>` : `<div class="card empty"><div class="h-lg">${esc(t('emptyT'))}</div>${esc(t('emptyS'))}</div>`}
  </div>`;
}

function pageExpenses() {
  let list = st.expenses.filter(e => ui.status === 'all' || e.status === ui.status).filter(e => ui.cat === 'all' || e.cat === ui.cat).sort(sortDesc);
  const groups = [];
  list.forEach(e => { const k = e.date.slice(0, 7); let g = groups[groups.length - 1]; if (!g || g.k !== k) groups.push(g = { k, items: [] }); g.items.push(e); });
  const seg = [['pending', 'stPending'], ['sent', 'stSent'], ['all', 'stAll']].map(([k, l]) => `<button class="${ui.status === k ? 'on' : ''}" data-act="status" data-v="${k}">${esc(t(l))}</button>`).join('');
  const chips = [['all', t('all')], ...CAT_KEYS.map(k => [k, catLabel(k)])].map(([k, l]) => `<button class="chip ${ui.cat === k ? 'on' : ''}" data-act="cat" data-v="${k}">${esc(l)}</button>`).join('');
  return `<div class="page"><div class="seg">${seg}</div><div class="chips">${chips}</div>
    ${groups.length ? groups.map(g => { const { y, m } = ymd(g.k + '-01'); return `<div class="month-h"><div class="h-lg">${esc(monthName(y, m))}</div><div class="muted num">${eur(g.items.reduce((s, e) => s + e.amount, 0))}</div></div><div class="card">${g.items.map(expRow).join('')}</div>`; }).join('')
      : `<div class="card empty"><div class="h-lg">${esc(t('emptyT'))}</div>${esc(t('emptyS'))}</div>`}
  </div>`;
}

const pendingMonths = () => {
  const m = new Map();
  st.expenses.filter(e => e.status === 'pending').forEach(e => { const k = e.date.slice(0, 7); const o = m.get(k) || { k, n: 0, total: 0 }; o.n++; o.total += e.amount; m.set(k, o); });
  return [...m.values()].sort((a, b) => a.k.localeCompare(b.k));
};
const selMonths = () => { const all = pendingMonths().map(x => x.k); if (!ui.months) return new Set(all); return new Set(all.filter(k => ui.months.has(k))); };
const selected = () => { const s = selMonths(); return st.expenses.filter(e => e.status === 'pending' && s.has(e.date.slice(0, 7))).sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id); };
const hasShare = () => { try { return !!(navigator.canShare && navigator.canShare({ files: [new File(['x'], 'x.txt', { type: 'text/plain' })] })); } catch (e) { return false; } };
const methods = () => {
  const m = []; if (CFG.googleClientId) m.push('gmail'); if (hasShare()) m.push('share'); m.push('download');
  const pref = CFG.googleClientId && isGmail(settings.personalEmail) ? 'gmail' : (hasShare() ? 'share' : (CFG.googleClientId ? 'gmail' : 'download'));
  return { primary: pref, others: m.filter(x => x !== pref) };
};

function pageSend() {
  const months = pendingMonths();
  if (!months.length) {
    return `<div class="page"><div class="card empty"><div class="h-lg">${esc(t('noPendT'))}</div>${esc(t('noPendS'))}</div>${historyHtml()}</div>`;
  }
  const sel = selected(), sm = selMonths();
  const total = sel.reduce((s, e) => s + e.amount, 0), recs = sel.filter(e => e.receipt), size = recs.reduce((s, e) => s + (e.receipt.size || 0), 0);
  const noRec = sel.length - recs.length;
  const mt = methods(), ok = profileOK() && sel.length;
  const label = { gmail: t('sendGmail'), share: t('shareMail'), download: t('dlMail') };
  const ic = { gmail: UI_ICON.mail, share: UI_ICON.share, download: UI_ICON.down };
  const alt = { gmail: t('viaGmail'), share: t('viaShare'), download: t('viaDl') };
  return `<div class="page">
    ${profileOK() ? '' : `<div class="notice warn">${esc(t('setupFirst'))}<br><button class="muted" style="text-decoration:underline;margin-top:6px" data-nav="settings">${esc(t('goSettings'))}</button></div>`}
    <div class="sec"><b>01</b> ${esc(t('sec_summary'))}</div>
    <div class="card"><dl class="kv">
      <dt>${esc(t('requester'))}</dt><dd>${esc(settings.name || '—')}</dd>
      <dt>${esc(t('from'))}</dt><dd>${esc(settings.personalEmail || '—')}</dd>
      <dt>${esc(t('to'))}</dt><dd>${esc(settings.sendTo || t('notSet'))}</dd>
      ${settings.ccApprover && isEmail(settings.approverEmail) ? `<dt>${esc(t('cc'))}</dt><dd>${esc(settings.approverEmail)}</dd>` : ''}
      <dt>${esc(t('approverL'))}</dt><dd>${esc(settings.approver || t('notSet'))}</dd>
    </dl></div>
    <div class="sec"><b>02</b> ${esc(t('sec_period'))}</div>
    <div class="chips">${months.map(x => { const { y, m } = ymd(x.k + '-01'); return `<button class="chip ${sm.has(x.k) ? 'on' : ''}" data-act="month" data-v="${x.k}">${esc(monthName(y, m))} · ${x.n}</button>`; }).join('')}</div>
    <div class="sec"><b>03</b> ${esc(t('sec_content'))}</div>
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>${esc(t('dateC'))}</th><th>${esc(t('conceptC'))}</th><th class="r">€</th></tr></thead><tbody>
      ${sel.map(e => `<tr><td class="num">${esc(dShort(e.date))}</td><td>${esc(e.concept || catLabel(e.cat))}<div class="muted small">${esc(catLabel(e.cat))}${e.receipt ? '' : ' · ' + esc(t('noRec'))}</div></td><td class="r num">${eur(e.amount)}</td></tr>`).join('')}
    </tbody><tfoot><tr><td colspan="2">${esc(t('total'))}</td><td class="r num">${eur(total)}</td></tr></tfoot></table></div></div>
    <p class="hint" style="margin:10px 2px 0">${esc(t('attach', recs.length, mb(size)))}${noRec ? ' · <span style="color:var(--warn)">' + esc(t('withoutRec', noRec)) + '</span>' : ''}</p>
    <div style="margin-top:22px"><button class="btn" id="send-main" data-act="send" data-m="${mt.primary}" ${ok ? '' : 'disabled'}>${svg(ic[mt.primary])}<span>${esc(label[mt.primary])}</span></button>
    <div class="btn-row" style="margin-top:10px">${mt.others.map(m => `<button class="btn ghost sm" data-act="send" data-m="${m}" ${ok ? '' : 'disabled'}>${svg(ic[m])}${esc(alt[m])}</button>`).join('')}</div></div>
    ${historyHtml()}
  </div>`;
}
function historyHtml() {
  if (!st.batches.length) return '';
  return `<div class="sec"><b>04</b> ${esc(t('history'))}</div><div class="card">${[...st.batches].reverse().slice(0, 8).map(b =>
    `<div class="row"><span class="ico">${svg(UI_ICON.mail)}</span><span class="row-main"><span class="row-t" style="display:block">${esc(b.period)}</span><span class="row-s" style="display:block">${esc(t('hist', new Date(b.at).toLocaleDateString(settings.lang === 'en' ? 'en-GB' : 'es-ES'), b.count, eur(b.total)))}</span></span></div>`).join('')}</div>`;
}

function pageSettings() {
  const first = !profileOK();
  const companies = [...new Set([...(CFG.companies || []), ...st.expenses.map(e => e.company).filter(Boolean)])];
  const f = (id, label, val, extra = '') => `<div class="fg"><label class="fl" for="s-${id}">${esc(label)}</label><input class="fi" id="s-${id}" value="${esc(val)}" ${extra}></div>`;
  return `<div class="page">
    ${first ? `<div class="h-xl" style="margin:6px 0 8px">${esc(t('welcomeT'))}</div><p class="muted" style="margin-bottom:6px">${esc(t('welcomeS'))}</p>` : ''}
    <div class="sec"><b>01</b> ${esc(t('sec_profile'))}</div>
    ${f('name', t('fullName'), settings.name, 'autocomplete="name"')}
    <div class="fg"><label class="fl" for="s-company">${esc(t('defCompany'))}</label><input class="fi" id="s-company" list="dl-co" value="${esc(settings.company)}"><datalist id="dl-co">${companies.map(c => `<option value="${esc(c)}">`).join('')}</datalist></div>
    <div class="sec"><b>02</b> ${esc(t('sec_approval'))}</div>
    ${f('approver', t('approver'), settings.approver)}
    ${f('approverEmail', t('approverEmail'), settings.approverEmail, 'type="email" inputmode="email" autocapitalize="off"')}
    <label class="check"><input type="checkbox" id="s-ccApprover" ${settings.ccApprover ? 'checked' : ''}><span>${esc(t('ccApprover'))}</span></label>
    <div class="sec"><b>03</b> ${esc(t('sec_delivery'))}</div>
    <div class="fg"><label class="fl" for="s-sendTo">${esc(t('sendTo'))}</label><input class="fi" id="s-sendTo" type="email" inputmode="email" autocapitalize="off" value="${esc(settings.sendTo)}"><div class="hint">${esc(t('sendToHint'))}</div></div>
    <div class="fg"><label class="fl" for="s-personalEmail">${esc(t('personal'))}</label><input class="fi" id="s-personalEmail" type="email" inputmode="email" autocapitalize="off" autocomplete="email" value="${esc(settings.personalEmail)}"><div class="hint">${esc(t('personalHint'))}</div></div>
    <button class="btn" data-act="save-settings">${esc(t('saveSettings'))}</button>
    <div class="sec"><b>04</b> ${esc(t('sec_app'))}</div>
    <div class="fl">${esc(t('language'))}</div>
    <div class="seg"><button class="${settings.lang === 'es' ? 'on' : ''}" data-act="lang" data-v="es">Español</button><button class="${settings.lang === 'en' ? 'on' : ''}" data-act="lang" data-v="en">English</button></div>
    <button class="btn danger sm" data-act="clean-sent">${esc(t('cleanSent'))}</button>
    <p class="hint" style="text-align:center;margin-top:18px">COMM Gastos ${VERSION}</p>
  </div>`;
}

function render() {
  document.documentElement.lang = settings.lang;
  $('#top-title').textContent = t(ui.page);
  document.querySelectorAll('[data-nav]').forEach(b => {
    if (!b.classList.contains('tab')) return;
    const k = b.dataset.nav; b.classList.toggle('on', k === ui.page);
    b.innerHTML = svg(UI_ICON[k]) + `<span>${esc(t(k))}</span>`;
  });
  const m = $('#main'), keep = m.scrollTop;
  m.innerHTML = { home: pageHome, expenses: pageExpenses, send: pageSend, settings: pageSettings }[ui.page]();
  if (ui.page === 'send') schedulePrepare(); else ui.pkg = null;
  m.scrollTop = keep;
}
function go(p) { ui.page = p; $('#main').scrollTop = 0; render(); }

/* ───────────── add / edit sheet ───────────── */
function openAdd(cat, id) {
  const e = id ? exp(id) : null;
  ui.draft = e ? { id: e.id, cat: e.cat, date: e.date, amount: String(e.amount).replace('.', ','), concept: e.concept || '', client: e.client || '', company: e.company || '', blob: null, type: e.receipt && e.receipt.type, hadReceipt: !!e.receipt, removeReceipt: false, status: e.status, sentAt: e.sentAt }
    : { id: null, cat: cat || 'restauracion', date: today(), amount: '', concept: '', client: '', company: settings.company || (st.expenses.length ? [...st.expenses].sort((a, b) => b.id - a.id)[0].company : '') || '', blob: null, type: null, hadReceipt: false, removeReceipt: false, status: 'pending' };
  const d = ui.draft;
  const uniq = k => [...new Set(st.expenses.map(x => x[k]).filter(Boolean))].slice(0, 60);
  const cos = [...new Set([...(CFG.companies || []), ...uniq('company')])];
  $('#add-title').textContent = e ? t('editExp') : t('addExp');
  $('#add-body').innerHTML = `
    ${d.status === 'sent' ? `<div class="notice"><b>${esc(t('wasSent'))}</b> ${esc(new Date(d.sentAt).toLocaleDateString(settings.lang === 'en' ? 'en-GB' : 'es-ES'))} <button class="muted" style="text-decoration:underline;margin-left:8px" data-act="unsend">${esc(t('backPending'))}</button></div>` : ''}
    <div class="fg"><label class="fl">${esc(t('receipt'))}</label><div class="dz" id="dz"></div></div>
    <div class="fg"><label class="fl">${esc(t('category'))}</label><div class="cats" id="cats">${CAT_KEYS.map(k => `<button type="button" class="cat ${d.cat === k ? 'on' : ''}" data-act="pick-cat" data-v="${k}">${svg(ICON[k])}<span>${esc(catLabel(k))}</span></button>`).join('')}</div></div>
    <div class="fr"><div class="fg"><label class="fl" for="f-amount">${esc(t('amount'))}</label><input class="fi num" id="f-amount" data-f="amount" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(d.amount)}"></div>
    <div class="fg"><label class="fl" for="f-date">${esc(t('date'))}</label><input class="fi" id="f-date" data-f="date" type="date" value="${esc(d.date)}"></div></div>
    <div class="fg"><label class="fl" for="f-concept">${esc(t('concept'))}</label><input class="fi" id="f-concept" data-f="concept" list="dl-concept" autocomplete="off" placeholder="${esc(t('conceptPh'))}" value="${esc(d.concept)}"></div>
    <div class="fr"><div class="fg"><label class="fl" for="f-client">${esc(t('client'))}</label><input class="fi" id="f-client" data-f="client" list="dl-client" autocomplete="off" value="${esc(d.client)}"></div>
    <div class="fg"><label class="fl" for="f-company">${esc(t('company'))}</label><input class="fi" id="f-company" data-f="company" list="dl-company" autocomplete="off" value="${esc(d.company)}"></div></div>
    <datalist id="dl-concept">${uniq('concept').map(v => `<option value="${esc(v)}">`).join('')}</datalist>
    <datalist id="dl-client">${uniq('client').map(v => `<option value="${esc(v)}">`).join('')}</datalist>
    <datalist id="dl-company">${cos.map(v => `<option value="${esc(v)}">`).join('')}</datalist>
    <input type="file" id="in-cam" accept="image/*" capture="environment" hidden>
    <input type="file" id="in-file" accept="image/*,application/pdf" hidden>`;
  $('#add-foot').innerHTML = (e ? `<button class="btn danger" data-act="del" aria-label="${esc(t('del'))}">${esc(t('del'))}</button>` : '') + `<button class="btn" data-act="save">${esc(t('save'))}</button>`;
  renderDz();
  $('#ov-add').classList.add('open'); $('#sh-add').classList.add('open');
}
function closeAdd() { $('#ov-add').classList.remove('open'); $('#sh-add').classList.remove('open'); if (ui.draft && ui.draft.preview) URL.revokeObjectURL(ui.draft.preview); ui.draft = null; if (document.activeElement) document.activeElement.blur(); }

async function renderDz() {
  const d = ui.draft, dz = $('#dz'); if (!d || !dz) return;
  let inner = '';
  const has = d.blob || (d.hadReceipt && !d.removeReceipt);
  if (has) {
    let src = d.preview;
    if (!src && d.id && !d.blob) src = await receiptUrl(d.id);
    const isPdf = d.type === 'application/pdf';
    inner = (isPdf || !src ? `<div class="pdf">${svg(UI_ICON.file)}<span>${isPdf ? 'PDF' : esc(t('receipt'))}</span></div>` : `<img src="${src}" alt="">`) +
      `<div class="btn-row" style="margin-top:12px"><button class="btn ghost sm" data-act="cam">${svg(UI_ICON.camera)}${esc(t('takePhoto'))}</button><button class="btn danger sm" data-act="rm-rec">${esc(t('removeRec'))}</button></div>`;
  } else {
    inner = `<div class="dz-btns"><button class="btn sm" data-act="cam">${svg(UI_ICON.camera)}${esc(t('takePhoto'))}</button><button class="btn ghost sm" data-act="file">${svg(UI_ICON.file)}${esc(t('pickFile'))}</button></div>`;
  }
  if (ui.draft === d) dz.innerHTML = inner;
}
async function onPicked(file) {
  const d = ui.draft; if (!file || !d) return;
  const r = await compress(file);
  if (ui.draft !== d) return;
  if (d.preview) URL.revokeObjectURL(d.preview);
  d.blob = r.blob; d.type = r.type; d.removeReceipt = false; d.preview = r.type === 'application/pdf' ? null : URL.createObjectURL(r.blob);
  renderDz();
}

async function saveExpense() {
  const d = ui.draft; if (!d) return;
  const amount = parseAmt(d.amount);
  if (!(amount > 0)) { $('#f-amount').classList.add('bad'); $('#f-amount').focus(); return toast(t('needAmount'), true); }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) { $('#f-date').classList.add('bad'); return toast(t('needDate'), true); }
  const id = d.id || st.nextId++;
  const keepOld = d.id ? exp(d.id) : null;
  let receipt = keepOld && keepOld.receipt && !d.removeReceipt ? keepOld.receipt : null;
  try {
    if (d.blob) { await idb.put(id, d.blob); receipt = { type: d.type, size: d.blob.size }; dropUrl(id); }
    else if (d.removeReceipt && keepOld) { await idb.del(id); dropUrl(id); receipt = null; }
  } catch (e) { toast('⚠ ' + e.message, true); return; }
  const rec = { id, date: d.date, cat: d.cat, amount: r2(amount), concept: d.concept.trim(), client: d.client.trim(), company: d.company.trim(),
    receipt, status: d.status || 'pending', sentAt: d.sentAt || null, batch: keepOld ? keepOld.batch : null, createdAt: keepOld ? keepOld.createdAt : Date.now() };
  if (keepOld) st.expenses[st.expenses.indexOf(keepOld)] = rec; else st.expenses.push(rec);
  persist(); closeAdd(); render(); toast(t('saved'));
}

/* ───────────── building the package ───────────── */
const slug = s => CmgMail.asciiName(s).replace(/^_+|_+$/g, '').slice(0, 32) || 'gasto';
const dmy = () => { const d = new Date(); return p2(d.getDate()) + '/' + p2(d.getMonth() + 1) + '/' + d.getFullYear(); };

function periodLabel(sel, fn) {
  const ks = [...new Set(sel.map(e => e.date.slice(0, 7)))].sort();
  const lab = k => { const { y, m } = ymd(k + '-01'); return fn(y, m); };
  if (ks.length === 1) return lab(ks[0]);
  const a = ymd(ks[0] + '-01'), b = ymd(ks[ks.length - 1] + '-01');
  return a.y === b.y ? TXT.es.months[a.m - 1] + ' – ' + fn(b.y, b.m) : lab(ks[0]) + ' – ' + lab(ks[ks.length - 1]);
}

async function buildPackage(sel) {
  const nameEs = settings.name.trim();
  let n = 0;
  const items = [];
  for (const e of sel) {
    let file = '', blob = null;
    if (e.receipt) { blob = await idb.get(e.id).catch(() => null); }
    if (blob) { n++; file = p2(n) + '_' + e.date + '_' + slug(e.concept || TXT.es.cats[e.cat]) + '.' + extFor(e.receipt.type); }
    items.push({ ...e, receiptFile: file, _blob: blob });
  }
  const bytes = await XlsxBuilder.build({ employee: nameEs, approver: settings.approver.trim(), sentDate: dmy() }, items, window.JSZip);
  const period = periodLabel(sel, monthNameEs), fileTag = slug(periodLabel(sel, (y, m) => TXT.es.months[m - 1] + '_' + y).replace(/ – /g, '_'));
  const who = slug(nameEs);
  const xlsxBlob = new Blob([bytes], { type: XLSX_MIME });
  let zipBlob = null;
  if (n) {
    const z = new JSZip();
    items.forEach(i => { if (i.receiptFile) z.file(i.receiptFile, i._blob); });
    zipBlob = await z.generateAsync({ type: 'blob', compression: 'STORE' });
  }
  const total = r2(sel.reduce((s, e) => s + e.amount, 0));
  return {
    xlsxBlob, zipBlob, xlsxName: `Hoja_de_gastos_${who}_${fileTag}.xlsx`, zipName: `Justificantes_${who}_${fileTag}.zip`,
    period, total, count: sel.length, receipts: n, ids: sel.map(e => e.id),
    subject: TXT.es.mailSubject(period, nameEs),
    body: TXT.es.mailBody({ name: nameEs, per: period, n: sel.length, total: new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' }).format(total), rec: n, approver: settings.approver.trim() })
  };
}

let prepT;
function schedulePrepare() {
  clearTimeout(prepT); ui.pkg = null;
  if (ui.page !== 'send') return;
  prepT = setTimeout(() => {
    const sel = selected(); if (!sel.length || !profileOK()) return;
    const key = rev + '|' + [...selMonths()].sort().join(',');
    const p = buildPackage(sel).then(v => { if (ui.pkg && ui.pkg.key === key) ui.pkg.value = v; return v; });
    ui.pkg = { key, promise: p, value: null }; p.catch(() => { });
  }, 250);
}

function markSent(pkg, method) {
  const at = Date.now(), bid = 'b' + at;
  st.expenses.forEach(e => { if (pkg.ids.includes(e.id)) { e.status = 'sent'; e.sentAt = at; e.batch = bid; } });
  st.batches.push({ id: bid, at, period: pkg.period, count: pkg.count, total: pkg.total, method, to: settings.sendTo });
  persist(); render();
}

async function doSend(method) {
  if (ui.busy) return;
  const sel = selected();
  if (!sel.length || !profileOK()) return;
  ui.busy = true;
  const btns = [...document.querySelectorAll('[data-act="send"]')]; btns.forEach(b => b.disabled = true);
  const main = $('#send-main'), orig = main && main.innerHTML;
  const busy = txt => { if (main) main.innerHTML = `<span class="spin"></span><span>${esc(txt)}</span>`; };
  try {
    const key = rev + '|' + [...selMonths()].sort().join(',');
    let pkg = ui.pkg && ui.pkg.key === key ? (ui.pkg.value || null) : null;
    if (!pkg) { busy(t('building')); pkg = ui.pkg && ui.pkg.key === key ? await ui.pkg.promise : await buildPackage(sel); }
    const cc = settings.ccApprover && isEmail(settings.approverEmail) ? settings.approverEmail.trim() : '';
    const atts = [{ name: pkg.xlsxName, type: XLSX_MIME, blob: pkg.xlsxBlob }].concat(pkg.zipBlob ? [{ name: pkg.zipName, type: 'application/zip', blob: pkg.zipBlob }] : []);

    if (method === 'gmail') {
      busy(t('sending'));
      let token;
      try { token = await CmgMail.getToken(CFG.googleClientId, settings.personalEmail.trim()); }
      catch (e) { throw Object.assign(new Error(t('gDenied')), { soft: true }); }
      const mime = await CmgMail.buildMime({ to: settings.sendTo.trim(), cc, subject: pkg.subject, text: pkg.body, attachments: atts });
      if (mime.size > 25 * 1024 * 1024) throw Object.assign(new Error(t('tooBig')), { soft: true });
      await CmgMail.gmailSend(token, mime);
      markSent(pkg, 'gmail'); toast(t('sentOk'));
    } else if (method === 'share') {
      try { await navigator.clipboard.writeText(settings.sendTo.trim()); toast(t('recCopied')); } catch (e) { }
      const files = atts.map(a => new File([a.blob], a.name, { type: a.type }));
      const done = await CmgMail.shareFiles({ files, title: pkg.subject, text: pkg.body });
      if (done && confirm(t('sentAsk'))) { markSent(pkg, 'share'); toast(t('markedSent')); }
    } else {
      atts.forEach((a, i) => setTimeout(() => CmgMail.download(a.blob, a.name), i * 400));
      toast(t('downloaded'));
      setTimeout(() => {
        location.href = 'mailto:' + encodeURIComponent(settings.sendTo.trim()).replace('%40', '@') + '?subject=' + encodeURIComponent(pkg.subject) + (cc ? '&cc=' + encodeURIComponent(cc) : '') + '&body=' + encodeURIComponent(pkg.body);
      }, 1200);
      setTimeout(() => { if (confirm(t('sentAsk'))) { markSent(pkg, 'download'); toast(t('markedSent')); } }, 6000);
    }
  } catch (e) {
    console.error(e);
    if (e && e.status === 401) { /* token expired/revoked → next tap asks again */ }
    toast(e && e.soft ? e.message : (method === 'gmail' ? t('gErr') : '') + (e && e.message || e), true);
  } finally {
    ui.busy = false;
    if (ui.page === 'send') { if (main && main.isConnected) main.innerHTML = orig; btns.forEach(b => b.disabled = false); }
  }
}

/* ───────────── settings ───────────── */
function saveSettings() {
  const v = id => $('#s-' + id).value.trim();
  const next = { name: v('name'), company: v('company'), approver: v('approver'), approverEmail: v('approverEmail'), ccApprover: $('#s-ccApprover').checked, sendTo: v('sendTo'), personalEmail: v('personalEmail') };
  const bad = { name: !next.name, approver: !next.approver, sendTo: !isEmail(next.sendTo), approverEmail: !!next.approverEmail && !isEmail(next.approverEmail), personalEmail: !!next.personalEmail && !isEmail(next.personalEmail) };
  Object.entries(bad).forEach(([k, b]) => $('#s-' + k).classList.toggle('bad', b));
  if (Object.values(bad).some(Boolean)) return toast(t('fixFields'), true);
  const was = profileOK();
  Object.assign(settings, next); saveSettingsLS(); toast(t('settingsSaved'));
  if (!was) go('home');
}

/* ───────────── events ───────────── */
document.addEventListener('click', async ev => {
  const nav = ev.target.closest('[data-nav]'); if (nav) { return go(nav.dataset.nav); }
  const el = ev.target.closest('[data-act]'); if (!el) return;
  const a = el.dataset.act, d = ui.draft;
  switch (a) {
    case 'add': return openAdd(el.dataset.cat);
    case 'edit': return openAdd(null, Number(el.dataset.id));
    case 'close-add': return closeAdd();
    case 'status': ui.status = el.dataset.v; return render();
    case 'cat': ui.cat = el.dataset.v; return render();
    case 'month': { const all = pendingMonths().map(x => x.k); const cur = selMonths(); const k = el.dataset.v; cur.has(k) ? cur.delete(k) : cur.add(k); ui.months = new Set(all.filter(x => cur.has(x))); return render(); }
    case 'send': return doSend(el.dataset.m);
    case 'hide-ios': flags.iosHint = 1; lsSet(LS_FLAGS, flags); return render();
    case 'lang': settings.lang = el.dataset.v; saveSettingsLS(); return render();
    case 'save-settings': return saveSettings();
    case 'clean-sent': {
      if (!confirm(t('cleanAsk'))) return;
      for (const e of st.expenses.filter(x => x.status === 'sent')) { await idb.del(e.id).catch(() => { }); dropUrl(e.id); }
      st.expenses = st.expenses.filter(x => x.status !== 'sent'); persist(); render(); return toast(t('cleaned'));
    }
    case 'pick-cat': if (d) { d.cat = el.dataset.v; document.querySelectorAll('#cats .cat').forEach(b => b.classList.toggle('on', b.dataset.v === d.cat)); } return;
    case 'cam': return $('#in-cam').click();
    case 'file': return $('#in-file').click();
    case 'rm-rec': if (d) { if (d.preview) URL.revokeObjectURL(d.preview); d.blob = null; d.preview = null; d.removeReceipt = true; d.type = null; renderDz(); } return;
    case 'save': return saveExpense();
    case 'unsend': if (d && d.id) { const e = exp(d.id); e.status = 'pending'; e.sentAt = null; e.batch = null; persist(); closeAdd(); render(); } return;
    case 'del': {
      if (!d || !d.id || !confirm(t('delAsk'))) return;
      await idb.del(d.id).catch(() => { }); dropUrl(d.id);
      st.expenses = st.expenses.filter(e => e.id !== d.id); persist(); closeAdd(); render(); return toast(t('deleted'));
    }
  }
});
document.addEventListener('input', ev => {
  const f = ev.target.dataset && ev.target.dataset.f; if (!f || !ui.draft) return;
  ui.draft[f] = ev.target.value; ev.target.classList.remove('bad');
});
document.addEventListener('change', ev => {
  if (ev.target.id === 'in-cam' || ev.target.id === 'in-file') { const f = ev.target.files && ev.target.files[0]; ev.target.value = ''; onPicked(f); }
});
document.addEventListener('keydown', ev => { if (ev.key === 'Escape' && ui.draft) closeAdd(); });

/* ───────────── boot ───────────── */
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => { });
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => { });
  let reloaded = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (reloaded || !navigator.serviceWorker.controller) return; reloaded = true; if (!ui.draft && !ui.busy) location.reload(); });
}
ui.page = profileOK() ? 'home' : 'settings';
render();
if (/[?&]add=true/.test(location.search) && profileOK()) setTimeout(() => openAdd(), 250);
window.__cmg = { st: () => st, settings: () => settings, buildPackage, idb };
})();
