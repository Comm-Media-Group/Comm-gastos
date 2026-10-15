/* Builds the "HOJA DE GASTOS" workbook: one sheet per month, same layout and styles as the
   original template (FECHA | CONCEPTOS x8 categories | PROYECTO | CLIENTE | EMPRESA),
   plus a JUSTIFICANTE column that points at the file name inside the receipts ZIP.
   Works in the browser (needs window.JSZip + window.XLSX_ASSETS) and in Node for tests. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsxBuilder = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  // category key -> column letter + header label (Spanish, as in the original sheet)
  const CATS = [
    { key: 'restauracion', col: 'C', label: 'RESTAURACION' },
    { key: 'taxi',         col: 'D', label: 'TAXI' },
    { key: 'vtc',          col: 'E', label: 'VTC' },
    { key: 'hotel',        col: 'F', label: 'HOTEL' },
    { key: 'alquiler',     col: 'G', label: 'ALQUILER COCHE' },
    { key: 'vuelos',       col: 'H', label: 'VUELOS' },
    { key: 'atrezzo',      col: 'I', label: 'ATREZZO' },
    { key: 'otros',        col: 'J', label: 'OTROS' }
  ];
  const MONTHS = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
  const MIN_ROWS = 25; // rows 6..30 exist in the template even when empty
  const FIRST = 6;

  // template style ids (cellXfs in the original styles.xml)
  const S = { plain: 1, hdr: 26, title: 27, titleFill: 28, concHdrL: 29, concHdrM: 30, concHdrR: 31,
              cat: 2, total: 3, totalNum: 4, date: 40, amount: 43, blank: 13 };

  // Excel draws shared edges from both cells, so a thin top border on row 6 would thin out the medium bottom border of the
  // header (FECHA / category boxes). Extra styles: first data row without top border, and a spacer row with a medium bottom.
  const X = { eurFirst: 44, numFirst: 45, textFirst: 46, spacer: 47, hdrBot: 48 };
  // Header boxes spanning two rows (FECHA, PROYECTO…) are NOT merged vertically: viewers disagree on how to draw borders inside
  // merged cells. Top half = border without bottom, bottom half (holds the text) = border without top.
  const HDR_TOP = 32; // existing xf: bold, centred, medium border on left/right/TOP only
  function extendStyles(xml) {
    xml = xml.replace('<borders count="22">', '<borders count="23">').replace('</borders>',
      '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top/><bottom style="medium"><color indexed="64"/></bottom><diagonal/></border></borders>');
    xml = xml.replace('<cellXfs count="44">', '<cellXfs count="49">').replace('</cellXfs>',
      '<xf numFmtId="164" fontId="0" fillId="5" borderId="15" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
      '<xf numFmtId="4" fontId="0" fillId="3" borderId="15" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="15" xfId="0" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="22" xfId="0" applyBorder="1"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="21" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs>');
    if (!/borders count="23"/.test(xml) || !/cellXfs count="49"/.test(xml)) throw new Error('styles.xml layout changed');
    return xml;
  }
  const esc = s => String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const r2 = n => Math.round((Number(n) || 0) * 100) / 100;

  function serial(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  const str = (ref, s, text) => `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(text)}</t></is></c>`;
  const num = (ref, s, v) => `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
  const empty = (ref, s) => `<c r="${ref}" s="${s}"/>`;

  function sheetXml(opts, items, selected) {
    const n = items.length;
    const dataEnd = Math.max(FIRST + MIN_ROWS - 1, FIRST + n - 1);
    const T = dataEnd + 2; // total row (one blank row above, like the template)
    const rows = [];

    // title (rows 1-2), spacer (3), headers (4-5)
    const titleCells = ['C','D','E','F','G','H','I','J','K','L','M','N','O'].map(c => empty(c + '1', S.titleFill)).join('');
    rows.push(`<row r="1" ht="15.75" customHeight="1">${str('B1', S.title, 'HOJA DE GASTOS ' + (opts.employee || '').toUpperCase() + (opts.report ? ' — ' + opts.report.toUpperCase() : ''))}${titleCells}</row>`);
    rows.push(`<row r="2" ht="15.75" customHeight="1">${empty('B2', S.title)}${titleCells.replace(/1"/g, '2"')}</row>`);
    rows.push('<row r="3" ht="6.75" customHeight="1"/>');
    rows.push(`<row r="4" ht="16" customHeight="1">${empty('B4', HDR_TOP)}${str('C4', S.concHdrL, 'CONCEPTOS')}` +
      ['D','E','F','G','H','I'].map(c => empty(c + '4', S.concHdrM)).join('') + empty('J4', S.concHdrR) +
      `${empty('L4', HDR_TOP)}${empty('M4', HDR_TOP)}${empty('N4', HDR_TOP)}${empty('O4', HDR_TOP)}</row>`);
    rows.push(`<row r="5" ht="16" customHeight="1">${str('B5', X.hdrBot, 'FECHA')}` +
      CATS.map(c => str(c.col + '5', S.cat, c.label)).join('') +
      `${str('L5', X.hdrBot, 'PROYECTO')}${str('M5', X.hdrBot, 'CLIENTE')}${str('N5', X.hdrBot, 'EMPRESA')}${str('O5', X.hdrBot, 'JUSTIFICANTE')}</row>`);

    // data rows
    const totals = Object.fromEntries(CATS.map(c => [c.key, 0]));
    for (let i = 0; i < dataEnd - FIRST + 1; i++) {
      const r = FIRST + i, e = items[i], first = i === 0;
      const sDate = first ? 39 : S.date, sAmt = first ? X.eurFirst : S.amount, sBlank = first ? X.numFirst : S.blank, sTxt = first ? X.textFirst : S.plain;
      let cells = '';
      if (e) {
        cells += num('B' + r, sDate, serial(e.date));
        CATS.forEach(c => {
          if (c.key === e.cat) { cells += num(c.col + r, sAmt, r2(e.amount)); totals[c.key] += r2(e.amount); }
          else cells += empty(c.col + r, sBlank);
        });
        cells += str('L' + r, sTxt, e.concept || '') + str('M' + r, sTxt, e.client || '') +
                 str('N' + r, sTxt, e.company || '') + str('O' + r, sTxt, e.receiptFile || 'Sin justificante');
      } else {
        cells += empty('B' + r, sTxt) + CATS.map(c => empty(c.col + r, sBlank)).join('') +
                 ['L','M','N','O'].map(c => empty(c + r, sTxt)).join('');
      }
      rows.push(`<row r="${r}">${cells}</row>`);
    }

    // spacer row (as in the template): medium bottom border closes the table, TOTAL row sits right below
    rows.push(`<row r="${T - 1}" ht="16" customHeight="1">${['B','C','D','E','F','G','H','I','J'].map(c => empty(c + (T - 1), X.spacer)).join('')}` +
      `${['L','M','N','O'].map(c => empty(c + (T - 1), S.plain)).join('')}</row>`);
    // totals row (SUM formulas with cached values so previews show numbers without recalculating)
    const grand = r2(CATS.reduce((s, c) => s + totals[c.key], 0));
    rows.push(`<row r="${T}" ht="16" customHeight="1">${str('B' + T, S.total, 'TOTAL')}` +
      CATS.map(c => `<c r="${c.col}${T}" s="${S.totalNum}"><f>SUM(${c.col}${FIRST}:${c.col}${T - 1})</f><v>${r2(totals[c.key])}</v></c>`).join('') +
      `${str('L' + T, S.total, 'TOTAL A REEMBOLSAR')}<c r="M${T}" s="${S.totalNum}"><f>SUM(C${T}:J${T})</f><v>${grand}</v></c></row>`);

    // sign-off block
    const meta = [
      ...(opts.reportLine ? ['Informe: ' + opts.reportLine] : []),
      'Solicitante: ' + (opts.employee || ''),
      'Responsable de aprobación: ' + (opts.approver || ''),
      'Fecha de envío: ' + (opts.sentDate || ''),
      'Visto bueno del responsable: ______________________'
    ];
    meta.forEach((m, i) => rows.push(`<row r="${T + 2 + i}">${str('B' + (T + 2 + i), 0, m)}</row>`));
    const last = T + 2 + meta.length - 1;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:O${last}"/>` +
      `<sheetViews><sheetView showGridLines="0"${selected ? ' tabSelected="1"' : ''} workbookViewId="0"/></sheetViews>` +
      `<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>` +
      `<cols><col min="2" max="2" width="12" customWidth="1"/><col min="3" max="3" width="14.6640625" customWidth="1"/>` +
      `<col min="7" max="7" width="15.83203125" customWidth="1"/><col min="11" max="11" width="3.33203125" customWidth="1"/>` +
      `<col min="12" max="12" width="28.5" customWidth="1"/><col min="13" max="13" width="19" customWidth="1"/>` +
      `<col min="14" max="14" width="22.83203125" customWidth="1"/><col min="15" max="15" width="34" customWidth="1"/></cols>` +
      `<sheetData>${rows.join('')}</sheetData>` +
      `<mergeCells count="2"><mergeCell ref="B1:O2"/><mergeCell ref="C4:J4"/></mergeCells>` +
      `<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` +
      `<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>`;
  }

  function sheetNames(groups) {
    const years = new Set(groups.map(g => g.year));
    const used = new Set();
    return groups.map(g => {
      let name = MONTHS[g.month - 1] + (years.size > 1 ? ' ' + g.year : '');
      let k = 2; while (used.has(name)) name = MONTHS[g.month - 1] + ' ' + g.year + ' (' + k++ + ')';
      used.add(name); return name;
    });
  }

  /** expenses: [{date:'YYYY-MM-DD', cat, amount, concept, client, company, receiptFile}] */
  async function build(opts, expenses, JSZipLib, assets) {
    const Z = JSZipLib || (typeof JSZip !== 'undefined' ? JSZip : require('jszip'));
    const A = assets || globalThis.XLSX_ASSETS || require('./xlsx-assets.js');
    const byMonth = new Map();
    [...expenses].sort((a, b) => a.date.localeCompare(b.date) || (a.id || 0) - (b.id || 0)).forEach(e => {
      const [y, m] = e.date.split('-').map(Number);
      const k = y * 100 + m;
      if (!byMonth.has(k)) byMonth.set(k, { year: y, month: m, items: [] });
      byMonth.get(k).items.push(e);
    });
    const groups = [...byMonth.keys()].sort((a, b) => a - b).map(k => byMonth.get(k));
    if (!groups.length) throw new Error('No expenses to export');
    const names = sheetNames(groups);

    const zip = new Z();
    zip.file('[Content_Types].xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      groups.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`);
    zip.file('_rels/.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('xl/workbook.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<bookViews><workbookView activeTab="0"/></bookViews><sheets>` +
      names.map((n, i) => `<sheet name="${esc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
      `</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      groups.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
      `<Relationship Id="rId${groups.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId${groups.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>`);
    zip.file('xl/styles.xml', extendStyles(A.styles));
    zip.file('xl/theme/theme1.xml', A.theme);
    groups.forEach((g, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(opts, g.items, i === 0)));

    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  return { build, CATS, MONTHS };
});
