/*
 * Inlezen en terugschrijven van het actieregister (.xlsx).
 * Werkt in de browser (window.Register) en in Node (module.exports) voor de tests.
 * Kolommen worden op koptekst herkend, dus volgorde en extra kolommen in de Excel
 * mogen veranderen zonder dat de app breekt.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Register = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // veldnaam -> herkende kopteksten (genormaliseerd, begin van de tekst telt)
  const FIELDS = {
    nr: ['nr', 'nummer', '#'],
    actie: ['actie', 'wat moet er gebeuren', 'omschrijving', 'taak'],
    domein: ['domein', 'categorie'],
    eigenaar: ['eigenaar', 'wie', 'verantwoordelijke'],
    prioriteit: ['prioriteit', 'prio'],
    deadline: ['deadline', 'wanneer', 'datum gereed', 'einddatum'],
    status: ['status'],
    achterstallig: ['achterstallig'],
    dagenOpen: ['dagen open'],
    rang: ['rang'],
    waarom: ['waarom nu', 'waarom'],
    aanpak: ['aanpak'],
    bron: ['bron'],
    besluitdatum: ['besluitdatum', 'datum besluit'],
    toelichting: ['toelichting', 'opmerking', 'notities'],
  };
  // kolommen die de Excel zelf uitrekent: nooit overschrijven
  const COMPUTED = new Set(['achterstallig', 'dagenOpen']);
  const DATE_FIELDS = new Set(['deadline', 'besluitdatum']);
  const NUMBER_FIELDS = new Set(['nr', 'rang']);

  const DEFAULT_STATUSES = ['Idee', 'Plan', 'Open', 'Loopt', 'Wacht op', 'Gepland', 'Te verifiëren', 'Klaar'];
  const DEFAULT_PRIORITIES = ['P1', 'P2', 'P3'];

  const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[—–-]/g, ' ').replace(/\s+/g, ' ').trim();

  function matchField(header) {
    const h = norm(header);
    if (!h) return null;
    for (const [field, names] of Object.entries(FIELDS)) {
      if (names.some((n) => h === n || h.startsWith(n + ' '))) return field;
    }
    return null;
  }

  // ---------- celwaarden ----------

  function plain(v) {
    if (v == null) return null;
    if (v instanceof Date) return v;
    if (typeof v === 'object') {
      if ('result' in v || 'formula' in v || 'sharedFormula' in v) return plain(v.result);
      if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
      if ('text' in v) return v.text; // hyperlink
      if ('error' in v) return null;
    }
    return v;
  }

  const pad = (n) => String(n).padStart(2, '0');

  // Excel-datum (UTC-middernacht in ExcelJS) of 'dd-mm-jjjj' -> 'jjjj-mm-dd'
  function toISODate(v) {
    v = plain(v);
    if (v == null || v === '') return '';
    if (v instanceof Date && !isNaN(v)) {
      return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())}`;
    }
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      return toISODate(new Date(Math.round((v - 25569) * 86400000)));
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
    if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    return s; // vrije tekst ("Q4", "na de zomer") laten we staan
  }

  const isISODate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s);
  const isoToExcel = (s) => {
    const [y, m, d] = s.split('-').map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  };

  function toText(v) {
    v = plain(v);
    if (v == null) return '';
    if (v instanceof Date) return toISODate(v);
    return String(v).trim();
  }

  // ---------- inlezen ----------

  function findHeader(ws) {
    const max = Math.min(ws.rowCount || 0, 30);
    for (let r = 1; r <= max; r++) {
      const cols = {};
      ws.getRow(r).eachCell({ includeEmpty: false }, (cell, c) => {
        const f = matchField(plain(cell.value));
        if (f && !(f in cols)) cols[f] = c;
      });
      if ('actie' in cols && ('status' in cols || 'eigenaar' in cols)) return { row: r, cols };
    }
    return null;
  }

  function pickSheet(wb) {
    const named = wb.worksheets.find((w) => norm(w.name).includes('actie'));
    const order = named ? [named, ...wb.worksheets.filter((w) => w !== named)] : wb.worksheets;
    for (const ws of order) {
      const header = findHeader(ws);
      if (header) return { ws, header };
    }
    return null;
  }

  function listFromValidation(ws, col) {
    // dropdown-waarden uit gegevensvalidatie halen (bv. Status, Prioriteit, Domein)
    const dv = ws.dataValidations && ws.dataValidations.model;
    if (!dv) return null;
    const letter = colLetter(col);
    for (const [range, rule] of Object.entries(dv)) {
      if (!rule || rule.type !== 'list' || !rule.formulae || !rule.formulae[0]) continue;
      const parts = range.split(/\s+/);
      const hit = parts.some((p) => p.replace(/\$/g, '').replace(/\d+/g, '').split(':')[0] === letter);
      if (!hit) continue;
      const f = String(rule.formulae[0]);
      if (f.startsWith('"')) return f.replace(/^"|"$/g, '').split(',').map((s) => s.trim()).filter(Boolean);
    }
    return null;
  }

  function colLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function parseWorkbook(wb) {
    const found = pickSheet(wb);
    if (!found) throw new Error('Geen actieregister gevonden: er is geen rij met kolomkoppen als "Actie" en "Status".');
    const { ws, header } = found;

    const headers = [];
    ws.getRow(header.row).eachCell({ includeEmpty: false }, (cell, c) => {
      headers.push({ col: c, title: toText(cell.value), field: matchField(plain(cell.value)) });
    });

    const titleLines = [];
    for (let r = 1; r < header.row; r++) {
      const t = toText(ws.getRow(r).getCell(1).value);
      if (t) titleLines.push(t);
    }

    const actions = [];
    const last = ws.rowCount;
    for (let r = header.row + 1; r <= last; r++) {
      const row = ws.getRow(r);
      const a = { row: r, extra: {} };
      let hasContent = false;
      for (const h of headers) {
        const raw = row.getCell(h.col).value;
        if (h.field) {
          let val;
          if (DATE_FIELDS.has(h.field)) val = toISODate(raw);
          else if (NUMBER_FIELDS.has(h.field)) { const p = plain(raw); val = p == null || p === '' ? '' : (typeof p === 'number' ? p : toText(p)); }
          else val = toText(raw);
          a[h.field] = val;
          if (h.field === 'actie' && val) hasContent = true;
        } else {
          const t = toText(raw);
          if (t) a.extra[h.title] = t;
        }
      }
      if (hasContent) actions.push(a);
    }

    const distinct = (f) => [...new Set(actions.map((a) => a[f]).filter(Boolean))].sort((x, y) => x.localeCompare(y, 'nl'));
    const lists = {
      status: listFromValidation(ws, header.cols.status) || mergeKnown(DEFAULT_STATUSES, distinct('status')),
      prioriteit: listFromValidation(ws, header.cols.prioriteit) || mergeKnown(DEFAULT_PRIORITIES, distinct('prioriteit')),
      domein: listFromValidation(ws, header.cols.domein) || distinct('domein'),
    };

    return {
      sheetName: ws.name,
      headerRow: header.row,
      columns: header.cols,
      headers,
      title: titleLines[0] || 'Actieregister',
      subtitle: titleLines.slice(1).join(' '),
      lists,
      actions,
    };
  }

  function mergeKnown(known, seen) {
    return [...known, ...seen.filter((s) => !known.includes(s))];
  }

  // ---------- terugschrijven ----------

  function writeValue(field, value) {
    if (value == null || value === '') return null;
    if (DATE_FIELDS.has(field) && isISODate(value)) return isoToExcel(value);
    if (NUMBER_FIELDS.has(field) && value !== '' && !isNaN(Number(value))) return Number(value);
    return value;
  }

  function shiftFormula(formula, fromRow, toRow) {
    // relatieve rijverwijzingen (F5, G5) meeschuiven naar de nieuwe rij
    return formula.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g, (m, d1, c, d2, r) =>
      d2 ? m : `${d1}${c}${d2}${Number(r) === fromRow ? toRow : Number(r) + (toRow - fromRow)}`);
  }

  function formulaOf(cell) {
    const v = cell.value;
    if (v && typeof v === 'object') {
      if (v.formula) return v.formula;
      if (v.sharedFormula) {
        // ExcelJS kan shared formulas teruggeven; vertaal via het bronadres
        const master = cell.worksheet.getCell(v.sharedFormula).value;
        if (master && master.formula) {
          return shiftFormula(master.formula, Number(v.sharedFormula.replace(/\D/g, '')), Number(cell.row));
        }
      }
    }
    return null;
  }

  /**
   * Schrijft de acties terug in het originele werkboek. Alleen gewijzigde en nieuwe
   * acties worden aangeraakt; opmaak, formules, tabellen en andere tabbladen blijven staan.
   */
  function applyActions(wb, parsed, actions) {
    const ws = wb.getWorksheet(parsed.sheetName);
    const cols = parsed.columns;
    let lastRow = parsed.actions.reduce((m, a) => Math.max(m, a.row), parsed.headerRow);

    const writable = Object.keys(cols).filter((f) => !COMPUTED.has(f));
    const changed = actions.filter((a) => a._changed && !a._new && a.row);
    for (const a of changed) {
      const row = ws.getRow(a.row);
      // alleen de velden die in de app zijn aangepast, zodat opmaak in andere cellen intact blijft
      const fields = Array.isArray(a._fields) ? writable.filter((f) => a._fields.includes(f)) : writable;
      for (const f of fields) {
        if (!(f in a)) continue;
        const cell = row.getCell(cols[f]);
        if (formulaOf(cell)) continue;
        cell.value = writeValue(f, a[f]);
      }
      row.commit && row.commit();
    }

    const added = actions.filter((a) => a._new);
    const template = lastRow > parsed.headerRow ? ws.getRow(lastRow) : null;
    for (const a of added) {
      const r = ++lastRow;
      const row = ws.getRow(r);
      if (template) {
        row.height = template.height;
        template.eachCell({ includeEmpty: true }, (tc, c) => {
          const cell = row.getCell(c);
          cell.style = JSON.parse(JSON.stringify(tc.style || {}));
          const f = formulaOf(tc);
          if (f) cell.value = { formula: shiftFormula(f, template.number, r) };
          if (tc.dataValidation) cell.dataValidation = JSON.parse(JSON.stringify(tc.dataValidation));
        });
      }
      for (const f of writable) {
        if (f in a) row.getCell(cols[f]).value = writeValue(f, a[f]);
      }
      a.row = r;
      row.commit && row.commit();
    }

    if (added.length) growTables(ws, parsed.headerRow, lastRow);
    return { changed: changed.length, added: added.length };
  }

  function growTables(ws, headerRow, lastRow) {
    // Excel-tabel laten meegroeien met de nieuwe rijen (dropdowns worden per cel gekopieerd)
    const tables = ws.tables ? Object.values(ws.tables) : [];
    for (const t of tables) {
      const model = t.table || t.model || t;
      const ref = model.tableRef || model.ref;
      if (!ref) continue;
      const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (!m || Number(m[2]) !== headerRow) continue;
      const newRef = `${m[1]}${m[2]}:${m[3]}${lastRow}`;
      if (model.tableRef !== undefined) model.tableRef = newRef;
      if (model.ref !== undefined) model.ref = newRef;
    }
  }

  // ---------- afgeleide velden ----------

  const DONE = 'Klaar';
  const NOT_ACTIVE = new Set(['Klaar', 'Idee', 'Plan']);

  function todayISO(now = new Date()) {
    return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  }

  function isOverdue(a, today = todayISO()) {
    return !!a.deadline && isISODate(a.deadline) && a.deadline < today && a.status !== DONE;
  }

  function daysOpen(a, today = todayISO()) {
    if (!a.besluitdatum || !isISODate(a.besluitdatum) || a.status === DONE) return null;
    return Math.round((isoToExcel(today) - isoToExcel(a.besluitdatum)) / 86400000);
  }

  function daysUntil(iso, today = todayISO()) {
    if (!iso || !isISODate(iso)) return null;
    return Math.round((isoToExcel(iso) - isoToExcel(today)) / 86400000);
  }

  return {
    FIELDS, DEFAULT_STATUSES, DEFAULT_PRIORITIES, DONE, NOT_ACTIVE,
    parseWorkbook, applyActions, matchField, toISODate, shiftFormula,
    isOverdue, daysOpen, daysUntil, todayISO, isISODate,
  };
});
