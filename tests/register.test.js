// Test: inlezen, wijzigen, rij toevoegen en terugschrijven van tests/fixture.xlsx.
// Gebruik: npm test   (maakt eerst de fixture met python3 + openpyxl)
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
const ExcelJS = require('exceljs');
const Register = require('../js/register.js');

const FIXTURE = path.join(__dirname, 'fixture.xlsx');
const OUT = path.join(__dirname, 'out.xlsx');

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fs.readFileSync(FIXTURE));
  const parsed = Register.parseWorkbook(wb);

  assert.equal(parsed.sheetName, 'Actieregister');
  assert.equal(parsed.headerRow, 4);
  assert.equal(parsed.title, 'Actieregister — Testpraktijk');
  assert.equal(parsed.actions.length, 4);
  const [a1, a2, a3] = parsed.actions;
  assert.equal(a1.actie, 'Testactie klaar');
  assert.equal(a1.deadline, '2026-08-31');
  assert.equal(a1.besluitdatum, '2026-07-08');
  assert.equal(a1.rang, 3);
  assert.equal(a2.eigenaar, 'Anna / Bert');
  assert.equal(a3.deadline, '');
  assert.deepEqual(parsed.lists.prioriteit, ['P1', 'P2', 'P3']);
  assert.ok(parsed.lists.status.includes('Te verifiëren'));

  const today = '2026-10-13';
  assert.equal(Register.isOverdue(a1, today), false); // klaar
  assert.equal(Register.isOverdue(a2, today), true);
  assert.equal(Register.daysOpen(a2, today), 62);
  assert.equal(Register.shiftFormula('IF(AND(F8<>"",F8<TODAY()),"JA","")', 8, 9), 'IF(AND(F9<>"",F9<TODAY()),"JA","")');
  assert.equal(Register.toISODate('5-1-2026'), '2026-01-05');

  // wijzigen + toevoegen
  const actions = parsed.actions.map((a) => ({ ...a }));
  Object.assign(actions[1], { status: 'Klaar', toelichting: 'Gedaan via app', actie: 'NIET SCHRIJVEN', _changed: true, _fields: ['status', 'toelichting'] });
  actions.push({ _new: true, nr: 5, actie: 'Nieuwe actie uit app', domein: 'B. Personeel en HRM', eigenaar: 'Anna',
    prioriteit: 'P2', deadline: '2026-11-15', status: 'Open', besluitdatum: today, toelichting: '' });
  const res = Register.applyActions(wb, parsed, actions);
  assert.deepEqual(res, { changed: 1, added: 1 });
  fs.writeFileSync(OUT, Buffer.from(await wb.xlsx.writeBuffer()));

  // opnieuw inlezen
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(fs.readFileSync(OUT));
  const p2 = Register.parseWorkbook(wb2);
  assert.equal(p2.actions.length, 5);
  assert.equal(p2.actions[1].status, 'Klaar');
  assert.equal(p2.actions[1].toelichting, 'Gedaan via app');
  assert.equal(p2.actions[1].actie, 'Testactie achterstallig'); // niet-gewijzigde velden blijven ongemoeid
  const added = p2.actions[4];
  assert.equal(added.row, 9);
  assert.equal(added.actie, 'Nieuwe actie uit app');
  assert.equal(added.deadline, '2026-11-15');
  const ws = wb2.getWorksheet('Actieregister');
  assert.match(ws.getCell('H9').value.formula, /F9<TODAY\(\)/);
  assert.match(ws.getCell('I9').value.formula, /TODAY\(\)-N9/);
  assert.equal(ws.getCell('G9').dataValidation.type, 'list');
  assert.match(ws.getCell('A9').fill.fgColor.argb, /DDEEFF$/);
  assert.equal(ws.tables.Actieregister.table.tableRef, 'A4:O9');
  assert.ok(wb2.getWorksheet('Overzicht'));
  assert.ok(wb2.getWorksheet('Toelichting'));
  console.log('OK: alle tests geslaagd →', path.relative(process.cwd(), OUT));
})().catch((e) => { console.error(e); process.exit(1); });
