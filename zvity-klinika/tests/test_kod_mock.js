// Перевірка Kod.gs на імітації Google Таблиць (без Google).
// Підготовка: python3 tests/export_wb.py Zvity_kliniki.xlsx /tmp/wb.json ; node tests/test_kod_mock.js /tmp/wb.json
const fs = require('fs'), path = require('path'), vm = require('vm'), assert = require('assert');
const wbJson = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));

function revive(v) { return v && typeof v === 'object' && v.$d ? new Date(v.$d) : v; }
function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function makeSheet(name, rows) {
  const data = rows.map(r => r.map(revive));
  function ensure(r, c) { while (data.length < r) data.push([]); for (let i = 0; i < r; i++) while (data[i].length < c) data[i].push(''); }
  function range(r, c, nr, nc) {
    return {
      getValues() { ensure(r + nr - 1, c + nc - 1); return data.slice(r - 1, r - 1 + nr).map(x => x.slice(c - 1, c - 1 + nc)); },
      getValue() { return this.getValues()[0][0]; },
      getDisplayValues() { return this.getValues().map(x => x.map(String)); },
      setValues(v) { ensure(r + nr - 1, c + nc - 1); assert(v.length === nr && v[0].length === nc, 'setValues size ' + name); v.forEach((row, i) => row.forEach((x, j) => { data[r - 1 + i][c - 1 + j] = x; })); },
      setValue(x) { ensure(r, c); data[r - 1][c - 1] = x; },
      clearContent() { ensure(r + nr - 1, c + nc - 1); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) data[r - 1 + i][c - 1 + j] = ''; }
    };
  }
  return {
    data, getName: () => name,
    getRange(a, b, c, d) {
      if (typeof a === 'string') {
        const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        const c1 = colNum(m[1]), r1 = +m[2], c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
        return range(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
      }
      return range(a, b, c || 1, d || 1);
    },
    insertRowAfter(r) { data.splice(r, 0, []); },
    getLastRow() { let l = 0; data.forEach((row, i) => { if (row.some(x => x !== '')) l = i + 1; }); return l; }
  };
}
const sheets = {}; Object.keys(wbJson).forEach(n => { sheets[n] = makeSheet(n, wbJson[n]); });
const ctx = {
  console, Date, Math, JSON, Object, String, Number, Array, RegExp, Error, isNaN, isFinite, parseFloat,
  SpreadsheetApp: { getActive: () => ({ getSheetByName: n => sheets[n] }), getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] }), getUi: () => { throw new Error('no ui'); } },
  LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Logger: { log() {} }, module: { exports: {} }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../apps-script/Import.gs'), 'utf8'), ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../apps-script/Kod.gs'), 'utf8'), ctx);
function blob(file) {
  const buf = fs.readFileSync(path.join(__dirname, 'samples', file));
  return { getName: () => file, getContentType: () => 'text/csv', getDataAsString: cs => new TextDecoder(cs === 'windows-1250' ? 'windows-1250' : 'utf-8').decode(buf) };
}
const V = sheets['Візити'].data, B = sheets['Банк'].data, F = sheets['Фактури'].data;
const lastRow = d => { let l = 0; d.forEach((r, i) => { if (i && r[0] !== '' && r[0] !== undefined) l = i; }); return l; };
let n = 0; const ok = (c, m) => { assert(c, m); n++; };

const v0 = lastRow(V), b0 = lastRow(B), f0 = lastRow(F);
console.log(ctx.importujBlob(blob('booksy_rezerwacje.csv'), 'AUTO', null));
console.log(ctx.importujBlob(blob('felgdent_wizyty.csv'), 'AUTO', null));
ok(lastRow(V) === v0 + 3, 'visits added ' + (lastRow(V) - v0));
const anna = V[v0 + 1];
ok(anna[2] === 'Kowalska Anna' && anna[6] === 'Endodoncja' && anna[8] === 1650 && anna[12] === 'Виконано' && anna[13] === 1.5, 'anna row ' + JSON.stringify(anna.slice(0, 14)));
ok(anna[3] === 'Так' && anna[4] === 'Booksy', 'anna new+source from Booksy: ' + anna[3] + '/' + anna[4]);
ok(V[v0 + 2][12] === 'Неявка', 'noshow');
ok(sheets['Послуги'].data.some(r => r[0] === 'Higienizacja'), 'new service added to Послуги');
ok(sheets['Довідники'].data.some(r => r[0] === 'Lekarz 9') === false, 'skipped planned visit doctor not added');
// повторний імпорт — без дублів
console.log(ctx.importujBlob(blob('felgdent_wizyty.csv'), 'AUTO', null));
ok(lastRow(V) === v0 + 3, 'no duplicates on re-import');

console.log(ctx.importujBlob(blob('mbank_wyciag.csv'), 'AUTO', null));
ok(lastRow(B) === b0 + 7, 'bank added ' + (lastRow(B) - b0));
const newB = B.slice(b0 + 1, b0 + 8);
ok(newB.find(r => String(r[4]).includes('GOOGLE'))[5] === 'Google Ads', 'bank autocat');
console.log(ctx.importujBlob(blob('ksef_faktury.csv'), 'AUTO', null));
ok(lastRow(F) === f0 + 2, 'invoices added');
const hs = F.slice(f0 + 1).find(r => r[0] === '2026/10/118');
ok(hs && hs[7] instanceof Date && hs[6] === 'Materiały i narzędzia (zakupy)', 'invoice paid via bank + category: ' + JSON.stringify(hs));
console.log(ctx.importujBlob(blob('google_ads_kampanie.csv'), 'AUTO', null));
ok(sheets['Google Ads'].data.filter(r => r[2] === 'Implanty').length >= 3, 'ads rows');
ok(sheets['Журнал імпорту'].data.length >= 7, 'journal rows');

// Felgdent з реальними статусами: «Brak informacji» → без джерела, напрям з префікса лікаря
const v1 = lastRow(V);
console.log(ctx.importujBlob(blob('felgdent_wizyty_pl.csv'), 'VISITS', null));
const endo = V.slice(v1 + 1).find(r => r[2] === 'Testowa Maria'), hig = V.slice(v1 + 1).find(r => r[2] === 'Próbna Ewa');
ok(endo[6] === 'Endodoncja' && endo[4] === 'ZnanyLekarz' && endo[12] === 'Виконано', 'endo row ' + JSON.stringify(endo.slice(0, 14)));
ok(hig[4] === '' && hig[6] === 'Profilaktyka' && hig[12] === 'Неявка', 'brak informacji -> empty source');
ok(lastRow(V) === v1 + 3, 'przesunieta/niepotwierdzona skipped');

// Звіти Felgdent (знеособлені копії реальних форматів), у т.ч. файл з кількома аркушами
const rep = JSON.parse(fs.readFileSync(path.join(__dirname, 'samples', 'felgdent_raporty.json'), 'utf8'));
const order = ['Srednia', 'Raport_pozyskania', 'Koszty_i_wplaty', 'Status_wizyt', 'Raport_przesuniec'];
const Ex = sheets['Витрати'].data, e0 = lastRow(Ex), v2 = lastRow(V);
order.forEach(pref => {
  const f = rep.find(x => x.file.startsWith(pref));
  const shs = f.sheets.map(sh => ({ name: sh.name, rows: sh.rows.map(r => r.map(revive)) }));
  console.log(ctx.importujArkusze_(f.file, shs, 'AUTO', new Date(2026, 9, 1)));
});
const docsNow = sheets['Довідники'].data.slice(3, 20).map(r => r[0]);
ok(docsNow.includes('Lekarz Testowy') && docsNow.includes('Nowy Lekarz'), 'FDUR adds doctors ' + docsNow);
ok(sheets['Довідники'].data[3 + docsNow.indexOf('Lekarz Testowy')][3] === 1.25, 'FDUR hours');
ok(lastRow(V) === v2 + 4, 'wplaty -> 4 visits');
const wv = V.slice(v2 + 1, v2 + 5);
ok(wv[1][6] === 'Protetyka' && wv[1][8] === 2000 && wv[1][11] === 'Karta' && wv[1][12] === 'Виконано', 'wplata korona ' + JSON.stringify(wv[1].slice(0, 14)));
ok(wv[2][5] === 'Lekarz Testowy' && wv[2][13] === 1.25 && wv[0][5] === '' && wv[0][13] === 0.7, 'doctor from addedBy + durations ' + wv.map(r => r[5] + '/' + r[13]));
ok(wv[1][4] === 'Z sieci' && wv[0][4] === '', 'source from pozyskanie; Brak informacji -> empty');
ok(lastRow(Ex) === e0 + 1 && Ex[e0 + 1][2] === 'Technik / laboratorium (koszt z Felgdent)' && Ex[e0 + 1][7] === 'Lekarz Testowy' && Ex[e0 + 1][5] === 300, 'koszty -> Витрати ' + JSON.stringify(Ex[e0 + 1]));
const FS = sheets['Felgdent статуси'].data;
ok(FS.filter(r => r[2] === 'Неявка' && r[3] === 12).length === 1 && FS[1][0] instanceof Date, 'status wizyt stored');
const Z = sheets['Записи Felgdent'].data;
ok(Z.filter(r => r[3] === 'Перенесено').length === 1 && Z.filter(r => r[3] === 'Неявка').length === 1, 'zapisy stored');
// повторний імпорт — без дублів
const f2 = rep.find(x => x.file.startsWith('Koszty'));
ctx.importujArkusze_(f2.file, f2.sheets.map(sh => ({ name: sh.name, rows: sh.rows.map(r => r.map(revive)) })), 'WPLATY', null);
ok(lastRow(V) === v2 + 4 && lastRow(Ex) === e0 + 1, 'no duplicates (forced WPLATY imports both sheets)');

console.log('OK — ' + n + ' перевірок Kod.gs на імітації пройдено');
