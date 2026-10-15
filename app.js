/* COMM Gastos — expense capture + monthly "Hoja de gastos" delivery. Vanilla JS, offline-first PWA. */
(() => {
'use strict';
const CFG = window.CMG_CONFIG || {};
const VERSION = '1.3.0';
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
    home: 'Inicio', expenses: 'Gastos', send: 'Enviar', reports: 'Informes', settings: 'Ajustes',
    saveNext: n => 'Guardar y siguiente (' + n + ')', viewFile: 'Ver justificante', viewHint: 'Toca una foto para verla en grande', zoomHint: 'Toca la imagen para ampliar', noFile: 'No se encuentra el archivo',
    draftsT: 'Por completar', draftsHint: 'Fotos guardadas sin datos. Ábrelas para añadir importe y categoría; no se envían hasta completarlas.',
    quickPhoto: 'Foto rápida', importFiles: 'Importar', draftTag: 'Por completar', draftTitle: 'Foto sin datos', saveDraft: 'Completar después',
    draftsSaved: n => n === 1 ? 'Guardado en «Por completar»' : n + ' guardados en «Por completar»', draftSaved: 'Guardado en «Por completar»',
    needFile: 'Añade una foto o un importe', nFilesN: n => n + ' archivos', draftsN: n => n + ' por completar',
    draftsWarn: n => n + (n === 1 ? ' gasto sin completar no se incluirá.' : ' gastos sin completar no se incluirán.'), complete: 'Completar',
    sharedIn: n => n + (n === 1 ? ' archivo recibido → «Por completar»' : ' archivos recibidos → «Por completar»'),
    sec_backup: 'Copia de seguridad', backupHint: 'Guarda un ZIP con todos tus gastos, fotos y ajustes. Si cambias de teléfono, restáuralo en el nuevo. Guárdalo en iCloud Drive / Google Drive o envíatelo por email.',
    backupNow: 'Crear copia', restore: 'Restaurar', lastBackup: s => 'Última copia: ' + s, never: 'nunca', backupDone: 'Copia creada ✓',
    restoreAsk: (n, r, cur) => 'Se va a restaurar la copia (' + n + ' gastos, ' + r + ' archivos).' + (cur ? ' Esto REEMPLAZA los ' + cur + ' gastos que hay ahora en este teléfono.' : '') + ' ¿Continuar?',
    restoreDone: 'Copia restaurada ✓', restoreBad: 'Este archivo no es una copia válida de COMM Gastos',
    checkUpdates: 'Buscar actualizaciones', checking: 'Actualizando…',
    bkNudgeNever: 'Aún no has hecho ninguna copia de seguridad.', bkNudgeOld: d => 'Tu última copia de seguridad es de hace ' + d + ' días.', makeBackup: 'Hacer copia', later: 'Más tarde',
    newReport: 'Nuevo informe', editReport: 'Editar informe', repName: 'Nombre del informe', repNamePh: 'p. ej. Noviembre 2026 o Viaje Valladolid',
    repDest: 'Destino (opcional)', repStart: 'Desde', repEnd: 'Hasta', general: 'General', generalS: 'Gastos sin informe asignado',
    report: 'Informe', delReport: 'Eliminar informe', delReportAsk: '¿Eliminar este informe vacío?', needRepName: 'Pon un nombre al informe',
    repSaved: 'Informe guardado', repDeleted: 'Informe eliminado', repEmpty: 'Sin gastos', edit: 'Editar',
    repPend: n => n + (n === 1 ? ' pendiente' : ' pendientes'), repSent: n => n + (n === 1 ? ' enviado' : ' enviados'),
    sentSec: 'Enviados · pendientes de reembolso', reimbursed: 'Reembolsado',
    expensed: 'Reembolsado: borrar gastos y fotos', expensedHint: 'Cuando ya te hayan pagado este informe, márcalo: se borran de este teléfono los gastos enviados y sus justificantes. Los gastos aún sin enviar no se tocan.',
    expensedAsk: (n, r, name) => '¿Borrar ' + n + (n === 1 ? ' gasto enviado' : ' gastos enviados') + ' y ' + r + (r === 1 ? ' justificante' : ' justificantes') + ' de «' + name + '»? No se puede deshacer.',
    expensedDone: 'Informe cerrado y gastos borrados',
    select: 'Seleccionar', cancel: 'Cancelar', pickAll: 'Marcar todos', moveN: n => 'Mover (' + n + ')', moveNew: '+ Nuevo informe con la selección',
    moved: (n, name) => n + (n === 1 ? ' gasto movido a «' : ' gastos movidos a «') + name + '»',
    pending: 'Pendiente de reembolso', nExp: 'Gastos', noReceipt: 'Sin justificante', thisMonth: 'Este mes',
    quick: 'Añadir rápido', recent: 'Últimos gastos', emptyT: 'Aún no hay gastos', emptyS: 'Pulsa + y haz una foto del justificante.',
    stPending: 'Pendientes', stSent: 'Enviados', stAll: 'Todos', all: 'Todas',
    sent: 'Enviado', noRec: 'Sin justificante',
    addExp: 'Nuevo gasto', editExp: 'Editar gasto', receipt: 'Justificante', takePhoto: 'Hacer foto', pickFile: 'Galería / archivos',
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
    mailSubject: (per, name, rep) => rep ? 'Hoja de gastos ' + rep + (per ? ' (' + per + ')' : '') + ' – ' + name : 'Hoja de gastos ' + per + ' – ' + name,
    mailBody: (o) => 'Hola,\n\nAdjunto la hoja de gastos de ' + o.name + ' correspondiente a ' + o.per + ', junto con el ZIP de justificantes.\n\n' +
      'Resumen\n' + (o.report ? '- Informe: ' + o.report + '\n' : '') + '- Gastos: ' + o.n + '\n- Total a reembolsar: ' + o.total + '\n- Justificantes adjuntos: ' + o.rec + ' de ' + o.n + '\n- Responsable de aprobación: ' + o.approver + '\n\nUn saludo,\n' + o.name + '\n\n— Enviado desde COMM Gastos'
  },
  en: {
    home: 'Home', expenses: 'Expenses', send: 'Send', reports: 'Reports', settings: 'Settings',
    saveNext: n => 'Save & next (' + n + ')', viewFile: 'View receipt', viewHint: 'Tap a photo to view it full screen', zoomHint: 'Tap the image to zoom', noFile: 'File not found',
    draftsT: 'To complete', draftsHint: 'Photos saved without details. Open them to add amount and category; they are not sent until completed.',
    quickPhoto: 'Quick photo', importFiles: 'Import', draftTag: 'To complete', draftTitle: 'Photo, no details', saveDraft: 'Complete later',
    draftsSaved: n => n === 1 ? 'Saved to “To complete”' : n + ' saved to “To complete”', draftSaved: 'Saved to “To complete”',
    needFile: 'Add a photo or an amount', nFilesN: n => n + ' files', draftsN: n => n + ' to complete',
    draftsWarn: n => n + (n === 1 ? ' unfinished expense will not be included.' : ' unfinished expenses will not be included.'), complete: 'Complete',
    sharedIn: n => n + (n === 1 ? ' file received → “To complete”' : ' files received → “To complete”'),
    sec_backup: 'Backup', backupHint: 'Saves a ZIP with all your expenses, photos and settings. If you change phone, restore it on the new one. Keep it in iCloud Drive / Google Drive or email it to yourself.',
    backupNow: 'Create backup', restore: 'Restore', lastBackup: s => 'Last backup: ' + s, never: 'never', backupDone: 'Backup created ✓',
    restoreAsk: (n, r, cur) => 'About to restore the backup (' + n + ' expenses, ' + r + ' files).' + (cur ? ' This REPLACES the ' + cur + ' expenses currently on this phone.' : '') + ' Continue?',
    restoreDone: 'Backup restored ✓', restoreBad: 'This file is not a valid COMM Gastos backup',
    checkUpdates: 'Check for updates', checking: 'Updating…',
    bkNudgeNever: 'You have not made a backup yet.', bkNudgeOld: d => 'Your last backup is ' + d + ' days old.', makeBackup: 'Back up now', later: 'Later',
    newReport: 'New report', editReport: 'Edit report', repName: 'Report name', repNamePh: 'e.g. November 2026 or Valladolid trip',
    repDest: 'Destination (optional)', repStart: 'From', repEnd: 'To', general: 'General', generalS: 'Expenses without an assigned report',
    report: 'Report', delReport: 'Delete report', delReportAsk: 'Delete this empty report?', needRepName: 'Give the report a name',
    repSaved: 'Report saved', repDeleted: 'Report deleted', repEmpty: 'No expenses', edit: 'Edit',
    repPend: n => n + ' pending', repSent: n => n + ' sent',
    sentSec: 'Sent · awaiting reimbursement', reimbursed: 'Reimbursed',
    expensed: 'Reimbursed: delete expenses & photos', expensedHint: 'Once this report has been paid out, tick it: the sent expenses and their receipts are deleted from this phone. Expenses not yet sent are left alone.',
    expensedAsk: (n, r, name) => 'Delete ' + n + ' sent expense' + (n === 1 ? '' : 's') + ' and ' + r + ' receipt' + (r === 1 ? '' : 's') + ' from “' + name + '”? This cannot be undone.',
    expensedDone: 'Report closed and expenses deleted',
    select: 'Select', cancel: 'Cancel', pickAll: 'Select all', moveN: n => 'Move (' + n + ')', moveNew: '+ New report with selection',
    moved: (n, name) => n + ' expense' + (n === 1 ? '' : 's') + ' moved to “' + name + '”',
    pending: 'Pending reimbursement', nExp: 'Expenses', noReceipt: 'No receipt', thisMonth: 'This month',
    quick: 'Quick add', recent: 'Latest expenses', emptyT: 'No expenses yet', emptyS: 'Tap + and take a photo of the receipt.',
    stPending: 'Pending', stSent: 'Sent', stAll: 'All', all: 'All',
    sent: 'Sent', noRec: 'No receipt',
    addExp: 'New expense', editExp: 'Edit expense', receipt: 'Receipt', takePhoto: 'Take photo', pickFile: 'Gallery / files',
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
let st = Object.assign({ expenses: [], batches: [], nextId: 1, reports: [], nextRep: 1 }, lsGet(LS_STATE, {}));
function normalizeState(s) {
  if (!Array.isArray(s.reports)) s.reports = [];
  if (!s.reports.some(r => r.general)) s.reports.unshift({ id: 'gen', general: true, name: 'General' });
  s.expenses.forEach(e => {
    if (!e.report || !s.reports.some(r => r.id === e.report)) e.report = 'gen';
    if (!e.files) e.files = e.receipt ? [{ key: e.id, type: e.receipt.type, size: e.receipt.size }] : []; // v1.0 stored one receipt under the expense id
    delete e.receipt;
  });
}
normalizeState(st);
let flags = lsGet(LS_FLAGS, {});
let rev = 0;
const ui = { page: 'home', status: 'pending', rep: 'all', report: null, rdraft: null, selMode: false, picked: new Set(), moveTo: null, moveAfterCreate: false, months: null, draft: null, pkg: null, busy: false };
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
  return { put: (id, blob) => tx('readwrite', s => s.put(blob, id)), get: id => tx('readonly', s => s.get(id)), del: id => tx('readwrite', s => s.delete(id)), clear: () => tx('readwrite', s => s.clear()) };
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
  reports: '<path d="M4 4h11l5 5v11H4z"/><path d="M15 4v5h5M8 13h8M8 17h5"/>',
  folder: '<path d="M3 6.5A1.5 1.5 0 014.5 5H9l2 2.5h8.5A1.5 1.5 0 0121 9v9.5a1.5 1.5 0 01-1.5 1.5h-15A1.5 1.5 0 013 18.5z"/>',
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
const rep = id => st.reports.find(r => r.id === id) || st.reports[0];
const repName = r => (r.general ? t('general') : r.name);
const repMeta = r => r.general ? t('generalS') : [r.dest, r.start && r.end ? dShort(r.start) + ' – ' + dShort(r.end) : (r.start ? dShort(r.start) : '')].filter(Boolean).join(' · ');
const visibleExpenses = () => st.expenses.filter(e => ui.status === 'all' || e.status === ui.status).filter(e => ui.rep === 'all' || e.report === ui.rep).sort(sortDesc);
function moveExpenses(ids, to) {
  st.expenses.forEach(e => { if (ids.includes(e.id)) e.report = to; });
  ui.selMode = false; ui.picked = new Set(); persist(); render(); toast(t('moved', ids.length, repName(rep(to))));
}

/* ───────────── pages ───────────── */
const filesOf = e => e.files || [];
const nFiles = e => filesOf(e).length;
const keyOf = k => (/^\d+$/.test(String(k)) ? Number(k) : k);
const newKey = () => 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const draftsList = () => st.expenses.filter(e => e.draft).sort((a, b) => b.id - a.id);
async function dropFiles(e) { for (const f of filesOf(e)) { await idb.del(f.key).catch(() => { }); dropUrl(f.key); } }
function hydrateThumbs() {
  document.querySelectorAll('[data-thumb]').forEach(async el => {
    const u = await receiptUrl(keyOf(el.dataset.thumb));
    if (u && el.isConnected) el.style.backgroundImage = `url("${u}")`;
  });
}
const bkDays = () => (flags.lastBackup ? Math.floor((Date.now() - flags.lastBackup) / 864e5) : null);

function expRow(e) {
  const label = catLabel(e.cat), r = rep(e.report), fs = filesOf(e), first = fs[0];
  const pm = ui.selMode && ui.page === 'expenses';
  const isImg = first && first.type !== 'application/pdf';
  const ico = e.draft
    ? `<span class="ico th" ${isImg ? `data-thumb="${esc(first.key)}"` : ''}>${isImg ? '' : svg(UI_ICON.file)}</span>`
    : `<span class="ico">${svg(ICON[e.cat] || ICON.otros)}</span>`;
  const title = e.draft ? t('draftTitle') : (e.concept || label);
  const sub = e.draft ? [dShort(e.date), fs.length > 1 ? t('nFilesN', fs.length) : '', r.general ? '' : r.name].filter(Boolean).join(' · ')
    : [label, e.client, r.general ? '' : r.name].filter(Boolean).join(' · ');
  const clip = !e.draft && fs.length ? ` ${svg(UI_ICON.clip, 'i" style="width:12px;height:12px;vertical-align:-1px')}${fs.length > 1 ? ' ' + fs.length : ''}` : '';
  const flag = e.draft ? `<span class="tag warn">${esc(t('draftTag'))}</span>` : (e.status === 'sent' ? `<span class="tag ok">${esc(t('sent'))}</span>` : (!fs.length ? `<span class="tag warn">${esc(t('noRec'))}</span>` : ''));
  return `<button class="row" data-act="${pm ? 'pick' : 'edit'}" data-id="${e.id}">
    ${pm ? `<span class="pk ${ui.picked.has(e.id) ? 'on' : ''}"></span>` : ''}${ico}
    <span class="row-main"><span class="row-t" style="display:block">${esc(title)}</span><span class="row-s" style="display:block">${esc(sub)}${clip}</span></span>
    <span class="row-r"><span class="row-amt num" style="display:block">${e.draft ? '—' : eur(e.amount)}</span><span class="row-d" style="display:block">${flag || esc(dShort(e.date))}</span></span>
  </button>`;
}

function pageHome() {
  const real = st.expenses.filter(e => !e.draft), drafts = draftsList();
  const pend = real.filter(e => e.status === 'pending');
  const total = pend.reduce((s, e) => s + e.amount, 0);
  const cm = today().slice(0, 7);
  const month = real.filter(e => e.date.startsWith(cm)).reduce((s, e) => s + e.amount, 0);
  const recent = [...real].sort((a, b) => b.id - a.id).slice(0, 5);
  const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) && !navigator.standalone && !flags.iosHint;
  const d = bkDays(), nudge = st.expenses.length && (d === null || d >= 30) && !(flags.bkSnooze > Date.now());
  let n = 0; const sec = label => `<div class="sec"><b>${p2(++n)}</b> ${esc(label)}</div>`;
  return `<div class="page">
    ${ios ? `<div class="notice">${esc(t('installIos'))} <button data-act="hide-ios" class="muted" style="text-decoration:underline;margin-left:6px">${esc(t('dismiss'))}</button></div>` : ''}
    ${nudge ? `<div class="notice">${esc(d === null ? t('bkNudgeNever') : t('bkNudgeOld', d))}<div class="btn-row" style="margin-top:10px"><button class="btn sm" data-act="backup">${esc(t('makeBackup'))}</button><button class="btn ghost sm" data-act="bk-later">${esc(t('later'))}</button></div></div>` : ''}
    <div class="hero">
      <div class="hero-label">${esc(t('pending'))}</div>
      <div class="h-xl hero-amount num">${eur(total)}</div>
      <div class="hero-row">
        <div class="hero-stat"><b>${pend.length}</b><span>${esc(t('nExp'))}</span></div>
        <div class="hero-stat"><b>${pend.filter(e => !nFiles(e)).length}</b><span>${esc(t('noReceipt'))}</span></div>
        <div class="hero-stat"><b>${Math.round(month)}€</b><span>${esc(t('thisMonth'))}</span></div>
      </div>
    </div>
    <div class="btn-row" style="margin-top:14px">
      <button class="btn" data-act="quick-cam">${svg(UI_ICON.camera)}<span>${esc(t('quickPhoto'))}</span></button>
      <button class="btn ghost" data-act="quick-file">${svg(UI_ICON.file)}<span>${esc(t('importFiles'))}</span></button>
    </div>
    ${drafts.length ? `${sec(t('draftsT') + ' · ' + drafts.length)}<div class="card">${drafts.map(expRow).join('')}</div><p class="hint" style="margin:10px 2px 0">${esc(t('draftsHint'))}</p>` : ''}
    ${sec(t('quick'))}
    <div class="tiles">${CAT_KEYS.map(k => `<button class="tile" data-act="add" data-cat="${k}">${svg(ICON[k])}<span>${esc(catLabel(k))}</span></button>`).join('')}</div>
    ${sec(t('recent'))}
    ${recent.length ? `<div class="card">${recent.map(expRow).join('')}</div>` : `<div class="card empty"><div class="h-lg">${esc(t('emptyT'))}</div>${esc(t('emptyS'))}</div>`}
  </div>`;
}

/* quick capture: every photo/file becomes its own "Por completar" draft */
const defaultReport = () => {
  const ok = rid => rid && st.reports.some(r => r.id === rid);
  return ok(ui.report) ? ui.report : (ok(st.lastReport) ? st.lastReport : 'gen');
};
async function quickDrafts(list, silent) {
  const files = [...list].filter(f => f && f.size);
  if (!files.length) return 0;
  let n = 0;
  for (const f of files) {
    try {
      const r = await compress(f), key = newKey();
      await idb.put(key, r.blob);
      st.expenses.push({ id: st.nextId++, report: defaultReport(), date: today(), cat: 'otros', amount: 0, concept: '', client: '', company: settings.company || '',
        files: [{ key, type: r.type, size: r.blob.size }], draft: true, status: 'pending', sentAt: null, batch: null, createdAt: Date.now() });
      n++;
    } catch (e) { console.error(e); }
  }
  if (n) { persist(); render(); if (!silent) toast(t('draftsSaved', n)); }
  return n;
}

function pickbar() {
  if (!st.reports.some(r => r.id === ui.moveTo)) ui.moveTo = (st.reports.find(r => !r.general) || st.reports[0]).id;
  const n = ui.picked.size;
  return `<div class="pickbar"><div class="pickbar-in"><select class="fi" data-pick="to">${st.reports.map(r => `<option value="${esc(r.id)}" ${ui.moveTo === r.id ? 'selected' : ''}>${esc(repName(r))}</option>`).join('')}</select><button class="btn sm" style="width:auto;flex-shrink:0" data-act="move" ${n ? '' : 'disabled'}>${esc(t('moveN', n))}</button></div><button class="muted small" style="display:block;margin:10px auto 0;text-decoration:underline" data-act="move-new" ${n ? '' : 'disabled'}>${esc(t('moveNew'))}</button></div>`;
}

function pageExpenses() {
  if (ui.rep !== 'all' && !st.reports.some(r => r.id === ui.rep)) ui.rep = 'all';
  const list = visibleExpenses();
  const groups = [];
  list.forEach(e => { const k = e.date.slice(0, 7); let g = groups[groups.length - 1]; if (!g || g.k !== k) groups.push(g = { k, items: [] }); g.items.push(e); });
  const seg = [['pending', 'stPending'], ['sent', 'stSent'], ['all', 'stAll']].map(([k, l]) => `<button class="${ui.status === k ? 'on' : ''}" data-act="status" data-v="${k}">${esc(t(l))}</button>`).join('');
  const chips = [['all', t('all')], ...st.reports.map(r => [r.id, repName(r)])].map(([k, l]) => `<button class="chip ${ui.rep === k ? 'on' : ''}" data-act="rep" data-v="${esc(k)}">${esc(l)}</button>`).join('');
  return `<div class="page"><div style="display:flex;justify-content:flex-end;gap:8px;margin-bottom:12px">${ui.selMode ? `<button class="chip" data-act="pick-all">${esc(t('pickAll'))}</button>` : ''}<button class="chip ${ui.selMode ? 'on' : ''}" data-act="sel-mode">${esc(ui.selMode ? t('cancel') : t('select'))}</button></div><div class="seg">${seg}</div><div class="chips">${chips}</div>
    ${groups.length ? groups.map(g => { const { y, m } = ymd(g.k + '-01'); return `<div class="month-h"><div class="h-lg">${esc(monthName(y, m))}</div><div class="muted num">${eur(g.items.reduce((s, e) => s + e.amount, 0))}</div></div><div class="card">${g.items.map(expRow).join('')}</div>`; }).join('')
      : `<div class="card empty"><div class="h-lg">${esc(t('emptyT'))}</div>${esc(t('emptyS'))}</div>`}
    ${ui.selMode ? pickbar() + '<div style="height:96px"></div>' : ''}
  </div>`;
}

const pendingOf = rid => st.expenses.filter(e => e.status === 'pending' && !e.draft && e.report === rid);
const pendingMonths = rid => {
  const m = new Map();
  pendingOf(rid).forEach(e => { const k = e.date.slice(0, 7); const o = m.get(k) || { k, n: 0, total: 0 }; o.n++; o.total += e.amount; m.set(k, o); });
  return [...m.values()].sort((a, b) => a.k.localeCompare(b.k));
};
const selMonths = rid => { const all = pendingMonths(rid).map(x => x.k); if (!ui.months) return new Set(all); return new Set(all.filter(k => ui.months.has(k))); };
const selected = rid => { const s = selMonths(rid); return pendingOf(rid).filter(e => s.has(e.date.slice(0, 7))).sort((a, b) => a.date.localeCompare(b.date) || a.id - b.id); };
const hasShare = () => { try { return !!(navigator.canShare && navigator.canShare({ files: [new File(['x'], 'x.txt', { type: 'text/plain' })] })); } catch (e) { return false; } };
const methods = () => {
  const m = []; if (CFG.googleClientId) m.push('gmail'); if (hasShare()) m.push('share'); m.push('download');
  const pref = CFG.googleClientId && isGmail(settings.personalEmail) ? 'gmail' : (hasShare() ? 'share' : (CFG.googleClientId ? 'gmail' : 'download'));
  return { primary: pref, others: m.filter(x => x !== pref) };
};

function pageReports() {
  if (ui.report && !st.reports.some(r => r.id === ui.report)) ui.report = null;
  if (ui.report) return pageReport(ui.report);
  const rows = st.reports.map(r => {
    const ex = st.expenses.filter(e => e.report === r.id), pend = ex.filter(e => e.status === 'pending' && !e.draft), dr = ex.filter(e => e.draft).length, sent = ex.filter(e => e.status === 'sent').length;
    const total = pend.reduce((s, e) => s + e.amount, 0);
    const tags = (dr ? `<span class="tag warn">${esc(t('draftsN', dr))}</span> ` : '') + (pend.length ? `<span class="tag">${esc(t('repPend', pend.length))}</span> ` : '') + (sent ? `<span class="tag ok">${esc(t('repSent', sent))}</span>` : '');
    return `<button class="row" data-act="open-report" data-id="${esc(r.id)}"><span class="ico">${svg(r.general ? UI_ICON.expenses : UI_ICON.folder)}</span>
      <span class="row-main"><span class="row-t" style="display:block">${esc(repName(r))}</span><span class="row-s" style="display:block">${esc(repMeta(r))}</span><span style="display:block;margin-top:7px">${tags || `<span class="tag">${esc(t('repEmpty'))}</span>`}</span></span>
      <span class="row-r"><span class="row-amt num" style="display:block">${eur(total)}</span></span></button>`;
  }).join('');
  return `<div class="page"><button class="btn" data-act="new-report">${svg('<path d="M12 5v14M5 12h14"/>')}<span>${esc(t('newReport'))}</span></button>
    <div class="sec"><b>01</b> ${esc(t('reports'))}</div><div class="card">${rows}</div>${historyHtml()}</div>`;
}

function pageReport(id) {
  const r = rep(id), ex = st.expenses.filter(e => e.report === id);
  const pend = ex.filter(e => e.status === 'pending' && !e.draft), drafts = ex.filter(e => e.draft), sentL = ex.filter(e => e.status === 'sent').sort(sortDesc);
  const head = `<button class="muted small" data-act="back-reports" style="margin:2px 0 12px;letter-spacing:.08em;padding:6px 0">‹ ${esc(t('reports'))}</button>
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:4px">
      <div style="min-width:0"><div class="h-lg" style="font-size:36px;overflow-wrap:anywhere">${esc(repName(r))}</div><div class="muted small" style="margin-top:4px">${esc(repMeta(r))}</div></div>
      ${r.general ? '' : `<button class="btn ghost sm" style="width:auto;flex-shrink:0" data-act="edit-report">${esc(t('edit'))}</button>`}
    </div>`;
  let n = 0; const sec = label => `<div class="sec"><b>${p2(++n)}</b> ${esc(label)}</div>`;
  let body = '';
  if (!pend.length) {
    body = `<div class="card empty" style="margin-top:20px"><div class="h-lg">${esc(t('noPendT'))}</div>${esc(t('noPendS'))}</div>`;
  } else {
    const months = pendingMonths(id), sel = selected(id), sm = selMonths(id);
    const total = sel.reduce((s, e) => s + e.amount, 0), nRec = sel.reduce((s, e) => s + nFiles(e), 0), size = sel.reduce((s, e) => s + filesOf(e).reduce((a, f) => a + (f.size || 0), 0), 0);
    const noRec = sel.filter(e => !nFiles(e)).length, mt = methods(), ok = profileOK() && sel.length;
    const label = { gmail: t('sendGmail'), share: t('shareMail'), download: t('dlMail') };
    const ic = { gmail: UI_ICON.mail, share: UI_ICON.share, download: UI_ICON.down };
    const alt = { gmail: t('viaGmail'), share: t('viaShare'), download: t('viaDl') };
    body = `${profileOK() ? '' : `<div class="notice warn" style="margin-top:16px">${esc(t('setupFirst'))}<br><button class="muted" style="text-decoration:underline;margin-top:6px" data-nav="settings">${esc(t('goSettings'))}</button></div>`}
    ${sec(t('sec_summary'))}
    <div class="card"><dl class="kv">
      <dt>${esc(t('report'))}</dt><dd>${esc(repName(r))}</dd>
      <dt>${esc(t('requester'))}</dt><dd>${esc(settings.name || '—')}</dd>
      <dt>${esc(t('from'))}</dt><dd>${esc(settings.personalEmail || '—')}</dd>
      <dt>${esc(t('to'))}</dt><dd>${esc(settings.sendTo || t('notSet'))}</dd>
      ${settings.ccApprover && isEmail(settings.approverEmail) ? `<dt>${esc(t('cc'))}</dt><dd>${esc(settings.approverEmail)}</dd>` : ''}
      <dt>${esc(t('approverL'))}</dt><dd>${esc(settings.approver || t('notSet'))}</dd>
    </dl></div>
    ${sec(t('sec_period'))}
    <div class="chips">${months.map(x => { const { y, m } = ymd(x.k + '-01'); return `<button class="chip ${sm.has(x.k) ? 'on' : ''}" data-act="month" data-v="${x.k}">${esc(monthName(y, m))} · ${x.n}</button>`; }).join('')}</div>
    ${sec(t('sec_content'))}
    <div class="card"><div class="tbl-wrap"><table class="tbl"><thead><tr><th>${esc(t('dateC'))}</th><th>${esc(t('conceptC'))}</th><th class="r">€</th></tr></thead><tbody>
      ${sel.map(e => `<tr><td class="num">${esc(dShort(e.date))}</td><td>${esc(e.concept || catLabel(e.cat))}<div class="muted small">${esc(catLabel(e.cat))}${nFiles(e) ? '' : ' · ' + esc(t('noRec'))}</div></td><td class="r num">${eur(e.amount)}</td></tr>`).join('')}
    </tbody><tfoot><tr><td colspan="2">${esc(t('total'))}</td><td class="r num">${eur(total)}</td></tr></tfoot></table></div></div>
    <p class="hint" style="margin:10px 2px 0">${esc(t('attach', nRec, mb(size)))}${noRec ? ' · <span style="color:var(--warn)">' + esc(t('withoutRec', noRec)) + '</span>' : ''}</p>
    <div style="margin-top:22px"><button class="btn" id="send-main" data-act="send" data-m="${mt.primary}" ${ok ? '' : 'disabled'}>${svg(ic[mt.primary])}<span>${esc(label[mt.primary])}</span></button>
    <div class="btn-row" style="margin-top:10px">${mt.others.map(m => `<button class="btn ghost sm" data-act="send" data-m="${m}" ${ok ? '' : 'disabled'}>${svg(ic[m])}${esc(alt[m])}</button>`).join('')}</div></div>`;
  }
  const done = sentL.length ? `${sec(t('sentSec'))}<div class="card">${sentL.map(expRow).join('')}</div>
    <p class="hint" style="margin:10px 2px 14px">${esc(t('expensedHint'))}</p>
    <button class="btn ghost" data-act="expensed">${svg('<path d="M5 12.5l4.5 4.5L19 7.5"/>')}<span>${esc(t('expensed'))}</span></button>` : '';
  const draftWarn = drafts.length ? `<div class="notice warn" style="margin-top:16px">${esc(t('draftsWarn', drafts.length))} <button class="muted" style="text-decoration:underline;margin-left:6px" data-act="edit" data-id="${drafts[0].id}">${esc(t('complete'))}</button></div>` : '';
  return `<div class="page">${head}${draftWarn}${body}${done}</div>`;
}

function historyHtml() {
  if (!st.batches.length) return '';
  return `<div class="sec"><b>02</b> ${esc(t('history'))}</div><div class="card">${[...st.batches].reverse().slice(0, 8).map(b =>
    `<div class="row"><span class="ico">${svg(UI_ICON.mail)}</span><span class="row-main"><span class="row-t" style="display:block">${esc(b.reportName && !b.reportName.toLowerCase().includes(b.period.toLowerCase()) ? b.reportName + ' · ' + b.period : (b.reportName || b.period))}</span><span class="row-s" style="display:block">${esc(t('hist', new Date(b.at).toLocaleDateString(settings.lang === 'en' ? 'en-GB' : 'es-ES'), b.count, eur(b.total)))}</span></span>${b.expensedAt ? `<span class="tag ok">${esc(t('reimbursed'))}</span>` : ''}</div>`).join('')}</div>`;
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
    <div class="sec"><b>05</b> ${esc(t('sec_backup'))}</div>
    <p class="muted small" style="margin-bottom:8px">${esc(t('backupHint'))}</p>
    <p class="small" style="margin-bottom:14px">${esc(t('lastBackup', flags.lastBackup ? new Date(flags.lastBackup).toLocaleDateString(settings.lang === 'en' ? 'en-GB' : 'es-ES') : t('never')))}</p>
    <div class="btn-row"><button class="btn sm" data-act="backup">${svg(UI_ICON.share)}${esc(t('backupNow'))}</button><button class="btn ghost sm" data-act="restore">${svg(UI_ICON.down)}${esc(t('restore'))}</button></div>
    <p class="hint" style="text-align:center;margin-top:22px">COMM Gastos ${VERSION}</p>
    <button class="btn ghost sm" style="margin-top:6px" data-act="refresh">${esc(t('checkUpdates'))}</button>
  </div>`;
}

function render() {
  document.documentElement.lang = settings.lang;
  const inDetail = ui.page === 'reports' && ui.report && st.reports.some(r => r.id === ui.report);
  $('#top-title').textContent = inDetail ? '' : t(ui.page);
  document.querySelectorAll('[data-nav]').forEach(b => {
    if (!b.classList.contains('tab')) return;
    const k = b.dataset.nav; b.classList.toggle('on', k === ui.page);
    b.innerHTML = svg(UI_ICON[k]) + `<span>${esc(t(k))}</span>` + (k === 'home' && draftsList().length ? `<i class="badge">${draftsList().length}</i>` : '');
  });
  const m = $('#main'), keep = m.scrollTop;
  m.innerHTML = { home: pageHome, expenses: pageExpenses, reports: pageReports, settings: pageSettings }[ui.page]();
  if (ui.page === 'reports' && ui.report) schedulePrepare(); else ui.pkg = null;
  m.scrollTop = keep;
  hydrateThumbs();
}
function go(p) { ui.page = p; ui.report = null; ui.months = null; $('#main').scrollTop = 0; render(); }
function openReportPage(id) { ui.page = 'reports'; ui.report = id; ui.months = null; $('#main').scrollTop = 0; render(); }

/* ───────────── add / edit sheet ───────────── */
function openAdd(cat, id) {
  const e = id ? exp(id) : null;
  ui.draft = e ? { id: e.id, cat: e.cat, date: e.date, amount: e.amount ? String(e.amount).replace('.', ',') : '', concept: e.concept || '', client: e.client || '', company: e.company || '', report: e.report,
      files: filesOf(e).map(f => ({ ...f })), removed: [], status: e.status, sentAt: e.sentAt, sticky: !!e.draft, wasDraft: !!e.draft }
    : { id: null, cat: cat || 'restauracion', date: today(), amount: '', concept: '', client: '', company: settings.company || (st.expenses.length ? [...st.expenses].sort((a, b) => b.id - a.id)[0].company : '') || '', report: defaultReport(),
      files: [], removed: [], status: 'pending' };
  const d = ui.draft;
  const uniq = k => [...new Set(st.expenses.map(x => x[k]).filter(Boolean))].slice(0, 60);
  const cos = [...new Set([...(CFG.companies || []), ...uniq('company')])];
  $('#add-title').textContent = e ? (e.draft ? t('draftsT') : t('editExp')) : t('addExp');
  $('#add-body').innerHTML = `
    ${d.status === 'sent' ? `<div class="notice"><b>${esc(t('wasSent'))}</b> ${esc(new Date(d.sentAt).toLocaleDateString(settings.lang === 'en' ? 'en-GB' : 'es-ES'))} <button class="muted" style="text-decoration:underline;margin-left:8px" data-act="unsend">${esc(t('backPending'))}</button></div>` : ''}
    <div class="dz-wrap" id="dzw"><label class="fl">${esc(t('receipt'))}</label><div class="dz" id="dz"></div></div>
    <div class="fg"><label class="fl" for="f-report">${esc(t('report'))}</label><select class="fi" id="f-report" data-f="report">${st.reports.map(r => `<option value="${esc(r.id)}" ${d.report === r.id ? 'selected' : ''}>${esc(repName(r))}</option>`).join('')}</select></div>
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
    <input type="file" id="in-file" accept="image/*,application/pdf" multiple hidden>`;
  const others = e && e.draft ? draftsList().filter(x => x.id !== e.id).length : 0;
  $('#add-foot').innerHTML = e && e.draft
    ? `<button class="btn ghost" data-act="save">${esc(t('save'))}</button>` + (others ? `<button class="btn" data-act="save-next">${esc(t('saveNext', others))}</button>` : '')
    : (e ? `<button class="btn danger" data-act="del" aria-label="${esc(t('del'))}">${esc(t('del'))}</button>` : `<button class="btn ghost" data-act="save-draft">${esc(t('saveDraft'))}</button>`) + `<button class="btn" data-act="save">${esc(t('save'))}</button>`;
  if (e && e.draft) $('#add-body').insertAdjacentHTML('beforeend', `<button class="btn danger sm" style="margin-top:6px" data-act="del">${esc(t('del'))}</button>`);
  renderDz();
  $('#ov-add').classList.add('open'); $('#sh-add').classList.add('open');
}
function closeAdd() {
  $('#ov-add').classList.remove('open'); $('#sh-add').classList.remove('open');
  if (ui.draft) ui.draft.files.forEach(f => { if (f.preview) URL.revokeObjectURL(f.preview); });
  ui.draft = null; if (document.activeElement) document.activeElement.blur();
}

function renderDz() {
  const d = ui.draft, dz = $('#dz'); if (!d || !dz) return;
  const strip = !!d.sticky && d.files.length > 0;
  $('#dzw').classList.toggle('sticky', strip); dz.classList.toggle('strip', strip);
  const thumbs = d.files.map((f, i) => `<div class="thx"><button type="button" class="thx-v" data-act="view-file" data-i="${i}" aria-label="${esc(t('viewFile'))}">${f.type === 'application/pdf'
    ? `<div class="pdf">${svg(UI_ICON.file)}<span>PDF</span></div>`
    : `<img alt="" ${f.preview ? `src="${f.preview}"` : `data-key="${esc(f.key)}"`}>`}</button><button type="button" class="thx-x" data-act="rm-file" data-i="${i}" aria-label="${esc(t('removeRec'))}">×</button></div>`).join('');
  dz.innerHTML = (thumbs ? `<div class="thumbs">${thumbs}${strip ? `<button type="button" class="thx add" data-act="file" aria-label="${esc(t('pickFile'))}">+</button>` : ''}</div>` : '') +
    (thumbs ? `<div class="hint" style="margin:-2px 0 ${strip ? 0 : 12}px">${esc(t('viewHint'))}</div>` : '') +
    (strip ? '' : `<div class="dz-btns"><button type="button" class="btn sm ${d.files.length ? 'ghost' : ''}" data-act="cam">${svg(UI_ICON.camera)}${esc(t('takePhoto'))}</button><button type="button" class="btn ghost sm" data-act="file">${svg(UI_ICON.file)}${esc(t('pickFile'))}</button></div>`);
  dz.querySelectorAll('img[data-key]').forEach(async im => { const u = await receiptUrl(keyOf(im.dataset.key)); if (u && im.isConnected) im.src = u; });
}

/* full-screen receipt viewer (tap image to zoom / un-zoom) */
async function openViewer(i) {
  const f = ui.draft && ui.draft.files[i]; if (!f) return;
  const url = f.preview || (f.blob ? URL.createObjectURL(f.blob) : await receiptUrl(keyOf(f.key)));
  if (!url) return toast(t('noFile'), true);
  if (f.type === 'application/pdf') { window.open(url, '_blank'); return; }
  const v = $('#viewer'), im = $('#viewer-img');
  im.classList.remove('zoom'); im.src = url; v.hidden = false; v.scrollTop = 0; $('#viewer-hint').textContent = t('zoomHint');
  $('#viewer-sc').scrollTo(0, 0);
}
function closeViewer() { const v = $('#viewer'); v.hidden = true; $('#viewer-img').removeAttribute('src'); }
function onPicked(list) {
  const d = ui.draft, files = [...(list || [])].filter(f => f && f.size);
  if (!d || !files.length) return;
  // processed one after another; saveExpense waits for d.pickP so a fast "Guardar" can't lose a photo
  d.pickP = (d.pickP || Promise.resolve()).then(async () => {
    for (const f of files) {
      const r = await compress(f);
      if (ui.draft !== d) return;
      d.files.push({ blob: r.blob, type: r.type, size: r.blob.size, preview: r.type === 'application/pdf' ? null : URL.createObjectURL(r.blob) });
    }
    renderDz();
  });
}

async function saveExpense(asDraft, goNext) {
  const d = ui.draft; if (!d) return;
  if (d.pickP) await d.pickP;
  let amount = parseAmt(d.amount);
  if (asDraft) {
    if (!d.files.length && !(amount > 0)) return toast(t('needFile'), true);
    if (!(amount > 0)) amount = 0;
  } else {
    if (!(amount > 0)) { $('#f-amount').classList.add('bad'); $('#f-amount').focus(); return toast(t('needAmount'), true); }
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.date)) { $('#f-date').classList.add('bad'); return toast(t('needDate'), true); }
  const id = d.id || st.nextId++;
  const keepOld = d.id ? exp(d.id) : null;
  const files = [];
  try {
    for (const f of d.files) {
      if (f.blob) { const key = newKey(); await idb.put(key, f.blob); files.push({ key, type: f.type, size: f.size }); }
      else files.push({ key: f.key, type: f.type, size: f.size });
    }
    for (const k of d.removed) { await idb.del(k).catch(() => { }); dropUrl(k); }
  } catch (e) { toast('⚠ ' + e.message, true); return; }
  const rec = { id, report: d.report || 'gen', date: d.date, cat: d.cat, amount: r2(amount), concept: d.concept.trim(), client: d.client.trim(), company: d.company.trim(),
    files, draft: !!asDraft, status: d.status || 'pending', sentAt: d.sentAt || null, batch: keepOld ? keepOld.batch : null, createdAt: keepOld ? keepOld.createdAt : Date.now() };
  if (keepOld) st.expenses[st.expenses.indexOf(keepOld)] = rec; else st.expenses.push(rec);
  st.lastReport = rec.report; persist(); closeAdd(); render(); toast(t(asDraft ? 'draftSaved' : 'saved'));
  if (goNext) { const nx = draftsList()[0]; if (nx) setTimeout(() => openAdd(null, nx.id), 380); }
}

/* ───────────── report (trip / project) sheet ───────────── */
function monthSuggestions() {
  // months of the expenses being moved (if any), else this month and last month
  const fmt = (y, m) => TXT.es.months[m] + ' ' + y;
  if (ui.moveAfterCreate && ui.picked.size) {
    const ks = [...new Set(st.expenses.filter(e => ui.picked.has(e.id)).map(e => e.date.slice(0, 7)))].sort();
    return ks.slice(0, 3).map(k => { const { y, m } = ymd(k + '-01'); return fmt(y, m - 1); });
  }
  const n = new Date();
  return [0, -1].map(k => { const d = new Date(n.getFullYear(), n.getMonth() + k, 1); return fmt(d.getFullYear(), d.getMonth()); });
}
function openReport(id) {
  const r = id ? rep(id) : null;
  ui.rdraft = { id: r ? r.id : null, name: r ? r.name : (ui.moveAfterCreate && monthSuggestions().length === 1 ? monthSuggestions()[0] : ''), dest: r ? r.dest || '' : '', start: r ? r.start || '' : '', end: r ? r.end || '' : '' };
  const d = ui.rdraft;
  $('#rep-title').textContent = r ? t('editReport') : t('newReport');
  $('#rep-body').innerHTML = `
    <div class="fg"><label class="fl" for="r-name">${esc(t('repName'))}</label><input class="fi" id="r-name" data-r="name" autocomplete="off" placeholder="${esc(t('repNamePh'))}" value="${esc(d.name)}"></div>
    ${r ? '' : `<div class="chips" style="margin-top:-6px">${monthSuggestions().map(nm => `<button class="chip" data-act="sug" data-v="${nm}">${nm}</button>`).join('')}</div>`}
    <div class="fg"><label class="fl" for="r-dest">${esc(t('repDest'))}</label><input class="fi" id="r-dest" data-r="dest" autocomplete="off" value="${esc(d.dest)}"></div>
    <div class="fr"><div class="fg"><label class="fl" for="r-start">${esc(t('repStart'))}</label><input class="fi" id="r-start" data-r="start" type="date" value="${esc(d.start)}"></div>
    <div class="fg"><label class="fl" for="r-end">${esc(t('repEnd'))}</label><input class="fi" id="r-end" data-r="end" type="date" value="${esc(d.end)}"></div></div>`;
  const canDel = r && !st.expenses.some(e => e.report === r.id);
  $('#rep-foot').innerHTML = (canDel ? `<button class="btn danger" data-act="del-report">${esc(t('del'))}</button>` : '') + `<button class="btn" data-act="save-report">${esc(t('save'))}</button>`;
  $('#ov-rep').classList.add('open'); $('#sh-rep').classList.add('open');
}
function closeRep() { $('#ov-rep').classList.remove('open'); $('#sh-rep').classList.remove('open'); ui.rdraft = null; if (document.activeElement) document.activeElement.blur(); }
function saveReport() {
  const d = ui.rdraft; if (!d) return;
  if (!d.name.trim()) { $('#r-name').classList.add('bad'); $('#r-name').focus(); return toast(t('needRepName'), true); }
  const data = { name: d.name.trim(), dest: d.dest.trim(), start: d.start, end: d.end };
  if (d.id) Object.assign(rep(d.id), data);
  else { const r = Object.assign({ id: 'r' + (st.nextRep++), createdAt: Date.now() }, data); st.reports.push(r);
    if (ui.moveAfterCreate && ui.picked.size) { st.expenses.forEach(e => { if (ui.picked.has(e.id)) e.report = r.id; }); ui.selMode = false; ui.picked = new Set(); }
    ui.moveAfterCreate = false; ui.page = 'reports'; ui.report = r.id; ui.months = null; }
  persist(); closeRep(); render(); toast(t('repSaved'));
}
async function markExpensed() {
  const r = rep(ui.report); if (!r) return;
  const sent = st.expenses.filter(e => e.report === r.id && e.status === 'sent');
  if (!sent.length) return;
  const recs = sent.reduce((s, e) => s + nFiles(e), 0);
  if (!confirm(t('expensedAsk', sent.length, recs, repName(r)))) return;
  for (const e of sent) await dropFiles(e);
  st.expenses = st.expenses.filter(e => !(e.report === r.id && e.status === 'sent'));
  st.batches.forEach(b => { if (b.report === r.id && !b.expensedAt) b.expensedAt = Date.now(); });
  if (!r.general && !st.expenses.some(e => e.report === r.id)) { st.reports = st.reports.filter(x => x.id !== r.id); if (st.lastReport === r.id) st.lastReport = null; ui.report = null; }
  persist(); render(); toast(t('expensedDone'));
}

/* ───────────── building the package ───────────── */
const slug = s => CmgMail.asciiName(s).replace(/^_+|_+$/g, '').slice(0, 32) || 'gasto';
const dmyIso = s => { const { y, m, d } = ymd(s); return p2(d) + '/' + p2(m) + '/' + y; };
const dmy = () => { const d = new Date(); return p2(d.getDate()) + '/' + p2(d.getMonth() + 1) + '/' + d.getFullYear(); };

function periodLabel(sel, fn) {
  const ks = [...new Set(sel.map(e => e.date.slice(0, 7)))].sort();
  const lab = k => { const { y, m } = ymd(k + '-01'); return fn(y, m); };
  if (ks.length === 1) return lab(ks[0]);
  const a = ymd(ks[0] + '-01'), b = ymd(ks[ks.length - 1] + '-01');
  return a.y === b.y ? TXT.es.months[a.m - 1] + ' – ' + fn(b.y, b.m) : lab(ks[0]) + ' – ' + lab(ks[ks.length - 1]);
}

async function buildPackage(sel, r) {
  const nameEs = settings.name.trim();
  const repEs = r.general ? '' : r.name;
  let n = 0;
  const items = [];
  for (const e of sel) {
    const got = [];
    for (const f of filesOf(e)) { const b = await idb.get(f.key).catch(() => null); if (b) got.push({ b, type: f.type }); }
    const names = [];
    if (got.length) {
      n++;
      got.forEach((g, i) => names.push(p2(n) + '_' + e.date + '_' + slug(e.concept || TXT.es.cats[e.cat]) + (i ? '_' + (i + 1) : '') + '.' + extFor(g.type)));
    }
    items.push({ ...e, receiptFile: names.join(', '), _files: got.map((g, i) => ({ name: names[i], blob: g.b })) });
  }
  const range = r.start && r.end ? dmyIso(r.start) + ' – ' + dmyIso(r.end) : (r.start ? dmyIso(r.start) : '');
  const bytes = await XlsxBuilder.build({ employee: nameEs, approver: settings.approver.trim(), sentDate: dmy(), report: repEs,
    reportLine: repEs ? [repEs, r.dest, range].filter(Boolean).join(' · ') : '' }, items, window.JSZip);
  const period = periodLabel(sel, monthNameEs);
  const dup = !!repEs && repEs.toLowerCase().includes(period.toLowerCase()); // e.g. report "Noviembre 2026" for November: don't repeat it
  const fileTag = slug(periodLabel(sel, (y, m) => TXT.es.months[m - 1] + '_' + y).replace(/ – /g, '_'));
  const who = slug(nameEs), repTag = repEs ? slug(repEs) : '';
  const xlsxBlob = new Blob([bytes], { type: XLSX_MIME });
  let zipBlob = null;
  if (n) {
    const z = new JSZip();
    items.forEach(i => i._files.forEach(f => z.file(f.name, f.blob)));
    zipBlob = await z.generateAsync({ type: 'blob', compression: 'STORE' });
  }
  const total = r2(sel.reduce((s, e) => s + e.amount, 0));
  return {
    xlsxBlob, zipBlob, xlsxName: `Hoja_de_gastos_${who}_${[repTag, dup ? '' : fileTag].filter(Boolean).join('_')}.xlsx`, zipName: `Justificantes_${who}_${[repTag, dup ? '' : fileTag].filter(Boolean).join('_')}.zip`,
    period, total, count: sel.length, receipts: n, ids: sel.map(e => e.id), report: r.id, reportName: repEs,
    subject: TXT.es.mailSubject(dup ? '' : period, nameEs, repEs),
    body: TXT.es.mailBody({ name: nameEs, per: period, report: repEs, n: sel.length, total: new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' }).format(total), rec: n, approver: settings.approver.trim() })
  };
}

const pkgKey = () => rev + '|' + ui.report + '|' + [...selMonths(ui.report)].sort().join(',');
let prepT;
function schedulePrepare() {
  clearTimeout(prepT); ui.pkg = null;
  if (ui.page !== 'reports' || !ui.report) return;
  prepT = setTimeout(() => {
    if (!ui.report) return;
    const sel = selected(ui.report); if (!sel.length || !profileOK()) return;
    const key = pkgKey(), r = rep(ui.report);
    const p = buildPackage(sel, r).then(v => { if (ui.pkg && ui.pkg.key === key) ui.pkg.value = v; return v; });
    ui.pkg = { key, promise: p, value: null }; p.catch(() => { });
  }, 250);
}

function markSent(pkg, method) {
  const at = Date.now(), bid = 'b' + at;
  st.expenses.forEach(e => { if (pkg.ids.includes(e.id)) { e.status = 'sent'; e.sentAt = at; e.batch = bid; } });
  st.batches.push({ id: bid, at, report: pkg.report, reportName: pkg.reportName, period: pkg.period, count: pkg.count, total: pkg.total, method, to: settings.sendTo });
  persist(); render();
}

async function doSend(method) {
  if (ui.busy || !ui.report) return;
  const r = rep(ui.report), sel = selected(r.id);
  if (!sel.length || !profileOK()) return;
  ui.busy = true;
  const btns = [...document.querySelectorAll('[data-act="send"]')]; btns.forEach(b => b.disabled = true);
  const main = $('#send-main'), orig = main && main.innerHTML;
  const busy = txt => { if (main) main.innerHTML = `<span class="spin"></span><span>${esc(txt)}</span>`; };
  try {
    const key = pkgKey();
    let pkg = ui.pkg && ui.pkg.key === key ? (ui.pkg.value || null) : null;
    if (!pkg) { busy(t('building')); pkg = ui.pkg && ui.pkg.key === key ? await ui.pkg.promise : await buildPackage(sel, r); }
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
    toast(e && e.soft ? e.message : (method === 'gmail' ? t('gErr') : '') + (e && e.message || e), true);
  } finally {
    ui.busy = false;
    if (ui.page === 'reports') { if (main && main.isConnected) main.innerHTML = orig; btns.forEach(b => b.disabled = false); }
  }
}

/* ───────────── backup & restore (one ZIP: data.json + every receipt file) ───────────── */
async function makeBackup() {
  try {
    toast(t('building'));
    const z = new JSZip();
    z.file('data.json', JSON.stringify({ app: 'comm-gastos', format: 1, exportedAt: new Date().toISOString(), appVersion: VERSION, settings, state: st }));
    for (const e of st.expenses) for (const f of filesOf(e)) {
      const b = await idb.get(f.key).catch(() => null);
      if (b) z.file('receipts/' + f.key + '.' + extFor(f.type), b);
    }
    const blob = await z.generateAsync({ type: 'blob', compression: 'STORE' });
    const name = 'COMM_Gastos_copia_' + today() + '.zip';
    let done = true;
    try { done = await CmgMail.shareFiles({ files: [new File([blob], name, { type: 'application/zip' })], title: 'COMM Gastos — ' + t('sec_backup') }); }
    catch (e) { CmgMail.download(blob, name); }
    if (done) { flags.lastBackup = Date.now(); lsSet(LS_FLAGS, flags); render(); toast(t('backupDone')); }
  } catch (e) { console.error(e); toast(e.message || String(e), true); }
}

async function restoreBackup(file) {
  if (!file) return;
  try {
    const z = await JSZip.loadAsync(file);
    const j = z.file('data.json'); if (!j) throw new Error('bad');
    const data = JSON.parse(await j.async('string'));
    if (!data || data.app !== 'comm-gastos' || !data.state || !Array.isArray(data.state.expenses)) throw new Error('bad');
    const nExp = data.state.expenses.length, nRec = data.state.expenses.reduce((s, e) => s + ((e.files || (e.receipt ? [1] : [])).length), 0);
    if (st.expenses.length && !confirm(t('restoreAsk', nExp, nRec, st.expenses.length))) return;
    toast(t('building'));
    await idb.clear();
    urls.forEach(u => URL.revokeObjectURL(u)); urls.clear();
    const ns = Object.assign({ expenses: [], batches: [], nextId: 1, reports: [], nextRep: 1 }, data.state);
    for (const e of ns.expenses) {
      if (!e.files) e.files = e.receipt ? [{ key: e.id, type: e.receipt.type, size: e.receipt.size }] : [];
      delete e.receipt;
      for (const f of e.files) {
        const entry = z.file('receipts/' + f.key + '.' + extFor(f.type));
        if (entry) { const raw = await entry.async('blob'); await idb.put(f.key, new Blob([raw], { type: f.type })); }
      }
    }
    normalizeState(ns);
    st = ns; settings = Object.assign(settings, data.settings || {});
    lsSet(LS_SET, settings); persist();
    ui.page = profileOK() ? 'home' : 'settings'; ui.report = null; ui.rep = 'all'; ui.selMode = false; ui.picked = new Set();
    render(); toast(t('restoreDone'));
  } catch (e) { console.error(e); toast(t('restoreBad'), true); }
}

async function consumeShared() {
  // files sent to the app from another app's share sheet (Android): the service worker parks them in a cache
  try {
    if (!('caches' in window) || !(await caches.has('cmg-share'))) return;
    const c = await caches.open('cmg-share'), reqs = await c.keys(); if (!reqs.length) return;
    const files = [];
    for (const rq of reqs) { const res = await c.match(rq); const b = await res.blob(); files.push(new File([b], decodeURIComponent(res.headers.get('X-Name') || 'shared'), { type: res.headers.get('Content-Type') || b.type })); await c.delete(rq); }
    const n = await quickDrafts(files, true);
    if (n) { go('home'); toast(t('sharedIn', n)); }
  } catch (e) { console.error(e); }
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
    case 'rep': ui.rep = el.dataset.v; return render();
    case 'month': { const all = pendingMonths(ui.report).map(x => x.k); const cur = selMonths(ui.report); const k = el.dataset.v; cur.has(k) ? cur.delete(k) : cur.add(k); ui.months = new Set(all.filter(x => cur.has(x))); return render(); }
    case 'send': return doSend(el.dataset.m);
    case 'hide-ios': flags.iosHint = 1; lsSet(LS_FLAGS, flags); return render();
    case 'lang': settings.lang = el.dataset.v; saveSettingsLS(); return render();
    case 'save-settings': return saveSettings();
    case 'clean-sent': {
      if (!confirm(t('cleanAsk'))) return;
      for (const e of st.expenses.filter(x => x.status === 'sent')) await dropFiles(e);
      st.expenses = st.expenses.filter(x => x.status !== 'sent'); persist(); render(); return toast(t('cleaned'));
    }
    /* reports */
    case 'open-report': return openReportPage(el.dataset.id);
    case 'back-reports': return go('reports');
    case 'new-report': ui.moveAfterCreate = false; return openReport();
    case 'edit-report': ui.moveAfterCreate = false; return openReport(ui.report);
    case 'close-rep': return closeRep();
    case 'save-report': return saveReport();
    case 'sug': if (ui.rdraft) { ui.rdraft.name = el.dataset.v; const i = $('#r-name'); i.value = el.dataset.v; i.classList.remove('bad'); } return;
    case 'del-report': {
      const id = ui.rdraft && ui.rdraft.id; if (!id || !confirm(t('delReportAsk'))) return;
      st.reports = st.reports.filter(x => x.id !== id); if (ui.report === id) ui.report = null; if (st.lastReport === id) st.lastReport = null;
      persist(); closeRep(); render(); return toast(t('repDeleted'));
    }
    case 'expensed': return markExpensed();
    /* multi-select + move */
    case 'sel-mode': ui.selMode = !ui.selMode; ui.picked = new Set(); return render();
    case 'pick': { const id = Number(el.dataset.id); ui.picked.has(id) ? ui.picked.delete(id) : ui.picked.add(id); return render(); }
    case 'pick-all': { const v = visibleExpenses(); const all = v.every(e => ui.picked.has(e.id)); v.forEach(e => all ? ui.picked.delete(e.id) : ui.picked.add(e.id)); return render(); }
    case 'move': {
      if (!ui.picked.size || !st.reports.some(r => r.id === ui.moveTo)) return;
      moveExpenses([...ui.picked], ui.moveTo); return;
    }
    case 'move-new': if (!ui.picked.size) return; ui.moveAfterCreate = true; return openReport();
    /* expense sheet */
    case 'pick-cat': if (d) { d.cat = el.dataset.v; document.querySelectorAll('#cats .cat').forEach(b => b.classList.toggle('on', b.dataset.v === d.cat)); } return;
    case 'cam': return $('#in-cam').click();
    case 'file': return $('#in-file').click();
    case 'rm-file': if (d) { const i = Number(el.dataset.i), f = d.files[i]; if (f) { if (f.preview) URL.revokeObjectURL(f.preview); if (!f.blob) d.removed.push(f.key); d.files.splice(i, 1); renderDz(); } } return;
    case 'save': return saveExpense(false);
    case 'save-next': return saveExpense(false, true);
    case 'view-file': return openViewer(Number(el.dataset.i));
    case 'close-view': return closeViewer();
    case 'zoom-view': return $('#viewer-img').classList.toggle('zoom');
    case 'save-draft': return saveExpense(true);
    case 'quick-cam': return $('#in-qcam').click();
    case 'quick-file': return $('#in-qfile').click();
    case 'backup': return makeBackup();
    case 'restore': return $('#in-restore').click();
    case 'refresh': toast(t('checking')); return forceRefresh();
    case 'bk-later': flags.bkSnooze = Date.now() + 7 * 864e5; lsSet(LS_FLAGS, flags); return render();
    case 'unsend': if (d && d.id) { const e = exp(d.id); e.status = 'pending'; e.sentAt = null; e.batch = null; persist(); closeAdd(); render(); } return;
    case 'del': {
      if (!d || !d.id || !confirm(t('delAsk'))) return;
      const old = exp(d.id); if (old) await dropFiles(old);
      st.expenses = st.expenses.filter(e => e.id !== d.id); persist(); closeAdd(); render(); return toast(t('deleted'));
    }
  }
});
document.addEventListener('input', ev => {
  const f = ev.target.dataset && ev.target.dataset.f;
  if (f && ui.draft) { ui.draft[f] = ev.target.value; ev.target.classList.remove('bad'); }
  const r = ev.target.dataset && ev.target.dataset.r;
  if (r && ui.rdraft) { ui.rdraft[r] = ev.target.value; ev.target.classList.remove('bad'); }
});
document.addEventListener('change', ev => {
  const fid = ev.target.id;
  if (fid === 'in-cam' || fid === 'in-file') { const fl = [...ev.target.files]; ev.target.value = ''; onPicked(fl); }
  if (fid === 'in-qcam' || fid === 'in-qfile') { const fl = [...ev.target.files]; ev.target.value = ''; quickDrafts(fl); }
  if (fid === 'in-restore') { const f = ev.target.files[0]; ev.target.value = ''; restoreBackup(f); }
  if (ev.target.dataset && ev.target.dataset.pick === 'to') ui.moveTo = ev.target.value;
  if (ev.target.dataset && ev.target.dataset.f === 'report' && ui.draft) ui.draft.report = ev.target.value;
});
document.addEventListener('keydown', ev => { if (ev.key === 'Escape') { if (!$('#viewer').hidden) closeViewer(); else if (ui.rdraft) closeRep(); else if (ui.draft) closeAdd(); } });

/* ───────────── force update (keeps all expenses/photos; only clears the app-file cache) ───────────── */
async function forceRefresh() {
  try {
    (await navigator.serviceWorker.getRegistrations()).forEach(r => r.unregister());
    for (const k of await caches.keys()) if (k !== 'cmg-share') await caches.delete(k);
  } catch (e) { /* ignore */ }
  location.replace(location.pathname + '?v=' + Date.now());
}

/* ───────────── boot ───────────── */
if (/[?&]refresh(=|&|$)/.test(location.search)) forceRefresh();            // open  …/?refresh  to force the newest version
else if (/[?&]v=\d+/.test(location.search)) history.replaceState(null, '', location.pathname);
if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => { });
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => { });
  let reloaded = false;
  navigator.serviceWorker.addEventListener('controllerchange', () => { if (reloaded || !navigator.serviceWorker.controller) return; reloaded = true; if (!ui.draft && !ui.busy) location.reload(); });
}
ui.page = profileOK() ? 'home' : 'settings';
render();
consumeShared(); if (/[?&]shared=1/.test(location.search)) history.replaceState(null, '', location.pathname);
if (/[?&]add=true/.test(location.search) && profileOK()) setTimeout(() => openAdd(), 250);
window.__cmg = { st: () => st, settings: () => settings, buildPackage, idb };
})();
