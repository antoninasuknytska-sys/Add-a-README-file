// Тести логіки імпорту: node tests/test_import.js
const fs = require('fs'), path = require('path'), assert = require('assert');
const vm = require('vm');
const I = { module: { exports: {} } };
vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../apps-script/Import.gs'), 'utf8'), I);
Object.assign(I, I.module.exports);
const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, 'defaults.json'), 'utf8'));
function read(name) {
  const buf = fs.readFileSync(path.join(__dirname, 'samples', name));
  let t = new TextDecoder('utf-8').decode(buf);
  if (t.includes('�')) t = new TextDecoder('windows-1250').decode(buf);   // як у czytajTabele()
  return I.parseCsvText(t);
}
function run(name, forced, opts) {
  const rows = read(name);
  let det = forced ? { type: forced, header: I.findHeader(rows, cfg.fields[forced]) } : I.detectType(rows, cfg.fields, cfg.hints, name);
  assert(!det.error, name + ': ' + det.error);
  const res = I.mapRecords(det.type, rows, det.header, Object.assign({ statusRules: cfg.statuses }, opts || {}));
  return { type: det.type, ...res };
}
let n = 0; const ok = (c, m) => { assert(c, m); n++; };

// числа й дати
ok(I.parseNumber('-9 000,00 PLN') === -9000, 'num pl');
ok(I.parseNumber('1.234,56') === 1234.56 && I.parseNumber('1,234.56') === 1234.56 && I.parseNumber('422,50') === 422.5, 'num formats');
ok(I.parseNumber('(250)') === -250 && I.parseNumber('+16290.00') === 16290 && I.parseNumber('abc') === null, 'num misc');
ok(+I.parseDate('05.10.2026') === +new Date(2026, 9, 5) && +I.parseDate('1 paź 2026') === +new Date(2026, 9, 1), 'dates');
ok(+I.parseDate('2026-10-05 09:00') === +new Date(2026, 9, 5) && +I.parseDate(46300) === +new Date(2026, 9, 5), 'date serial');
ok(Math.abs(I.parseDuration(null, '09:00', '10:30') - 1.5) < 1e-9 && I.parseDuration(90) === 1.5 && I.parseDuration('1:15') === 1.25, 'duration');

// mBank: Windows-1250, преамбула, «;», «-9 000,00 PLN», дві однакові операції
let r = run('mbank_wyciag.csv');
ok(r.type === 'BANK', 'mbank type ' + r.type);
ok(r.records.length === 7, 'mbank rows ' + r.records.length);
ok(r.records[0].amount === -9000 && r.records[2].amount === 24120.55, 'mbank amounts');
ok(r.records[0].description.includes('WYNAJMUJĄCY'), 'mbank polish letters');
ok(r.records[5].key !== r.records[6].key, 'duplicate same-day payments kept apart');
const cats = r.records.map(x => I.categorize(x.counterparty, x.description, x.amount, cfg.rules));
ok(cats[0] === 'Czynsz' && cats[1] === 'Google Ads' && cats[2] === 'Надходження: термінал', 'cats ' + cats);
ok(cats[3] === 'Materiały i narzędzia (zakupy)' && cats[4] === '' && cats[5] === 'Надходження: BLIK / перекази пацієнтів', 'cats2 ' + cats);

// PKO: контрагент і tytuł у колонках без назви
r = run('pko_historia.csv');
ok(r.type === 'BANK' && r.records.length === 2, 'pko');
ok(I.categorize(r.records[0].counterparty, r.records[0].description, r.records[0].amount, cfg.rules) === 'ZUS', 'pko zus: ' + r.records[0].description);
ok(I.categorize('', r.records[1].description, r.records[1].amount, cfg.rules) === 'Надходження: PayU raty', 'pko payu');

// Google Ads: 2 рядки заголовка, польські дати, рядок «Razem» пропускається
r = run('google_ads_kampanie.csv');
ok(r.type === 'ADS' && r.records.length === 2 && r.skipped >= 1, 'ads ' + r.records.length);
ok(r.records[0].cost === 422.5 && r.records[0].impressions === 1200 && +r.records[0].date === +new Date(2026, 9, 1), 'ads values');

// Фактури KSeF
r = run('ksef_faktury.csv');
ok(r.type === 'INVOICES' && r.records.length === 2 && r.records[0].gross === 5200 && r.records[0].nip === '5260001111', 'invoices');

// Звірка фактур: номер у tytule (1), сума + продавець (2)
const bank = [
  { idx: 0, date: new Date(2026, 9, 16), amount: -5200, counterparty: 'HENRY SCHEIN POLSKA', description: 'FV 2026/10/118' },
  { idx: 1, date: new Date(2026, 9, 18), amount: -6900, counterparty: 'PRACOWNIA PROTETYCZNA ABC', description: 'zaplata' },
  { idx: 2, date: new Date(2026, 9, 18), amount: -6900, counterparty: 'INNA FIRMA', description: 'cos' }];
const inv = r.records.map((x, i) => ({ idx: i, ...x }));
const mt = I.matchInvoices(inv, bank, []);
ok(mt.length === 2 && mt.find(m => m.invIdx === 0).bankIdx === 0 && mt.find(m => m.invIdx === 1).bankIdx === 1, 'match ' + JSON.stringify(mt));

// Felgdent: статуси, тривалість з годин, «Zaplanowana» пропускається, невідомий статус → попередження
r = run('felgdent_wizyty.csv');
ok(r.type === 'VISITS', 'felg type ' + r.type);
ok(r.records.length === 3 && r.skipped === 1, 'felg rows ' + r.records.length + '/' + r.skipped);
ok(r.records[0].status === 'Виконано' && r.records[0].duration === 1.5 && r.records[0].amount === 1650 && r.records[0].payment === 'Karta', 'felg rec0');
ok(r.records[1].status === 'Неявка', 'felg noshow');
ok(r.records[2].status === 'W trakcie' && r.warnings.some(w => w.includes('W trakcie')), 'felg unknown status');

// Booksy: визначення за назвою файлу, статуси
r = run('booksy_rezerwacje.csv');
ok(r.type === 'BOOKSY' && r.records[0].platform === 'Booksy' && r.records[0].status === 'Виконано' && r.records[1].status === 'Скасовано', 'booksy');


// Felgdent: реальні статуси й лікарі з префіксом спеціалізації
r = run('felgdent_wizyty_pl.csv', 'VISITS');
ok(r.records.length === 3 && r.skipped === 2, 'felg pl rows ' + r.records.length + '/' + r.skipped);
ok(r.records[0].status === 'Виконано' && Math.abs(r.records[0].duration - 56 / 60) < 0.01, 'odbyta');
ok(r.records[1].status === 'Неявка' && r.records[2].status === 'Скасовано', 'nie przyszlo / odwolana przez gabinet');
ok(r.records[0].source === 'ZnanyLekarz' && r.records[1].source === 'Brak informacji', 'source column');
ok(I.directionFromDoctor('ENDODONTA Jan Testowy') === 'Endodoncja' && I.directionFromDoctor('HIGIENA Anna Testowa') === 'Profilaktyka', 'prefix single');
ok(I.directionFromDoctor('Ortodonta Ola Testowa') === 'Ortodoncja' && I.directionFromDoctor('ZACH/PROT/CHIR Piotr Testowy') === '' && I.directionFromDoctor('Лікар 1') === '', 'prefix multi/none');

// Справжні формати звітів Felgdent (знеособлені копії): tests/samples/felgdent_raporty.json
const rep = JSON.parse(fs.readFileSync(path.join(__dirname, 'samples', 'felgdent_raporty.json'), 'utf8'));
const rv = v => v && v.$d ? new Date(v.$d) : v;
const got = {};
rep.forEach(f => f.sheets.forEach(sh => {
  const rows = sh.rows.map(r => r.map(rv));
  const det = I.detectType(rows, cfg.fields, cfg.hints, (f.sheets.length > 1 ? sh.name + ' | ' : '') + f.file);
  if (det.error) return;
  const sm = sh.name.match(/^(\d{1,2})(20\d{2})$/);
  got[det.type] = I.mapRecords(det.type, rows, det.header, { statusRules: cfg.statuses, month: sm ? new Date(+sm[2], +sm[1] - 1, 1) : new Date(2026, 9, 1) });
}));
ok(Object.keys(got).sort().join() === 'FDUR,FKOSZTY,FSOURCES,FSTATUS,FWIZYTY,FZAPISY,INVOICES,WPLATY', 'felgdent types ' + Object.keys(got));
const w = got.WPLATY.records;
ok(w.length === 4 && w[0].payment === 'Czek' && w[1].payment === 'Karta' && w[3].payment === 'Wpłata mieszana' && w[1].amount === 2000, 'wplaty');
ok(w[1].lines.length === 3 && w[1].lines[0] === 'korona cyrkonowa TEST', 'wplaty lines ' + w[1].lines);
ok(I.directionFromService(w[0].lines[0]) === 'Profilaktyka' && I.directionFromService(w[1].lines[0]) === 'Protetyka' &&
   I.directionFromService(w[2].lines[0]) === 'Zachowawcza' && I.directionFromService(w[3].lines[0]) === 'Zachowawcza', 'service directions');
ok(got.FKOSZTY.records[0].category === 'Koszty z Felgdent (kontrola)' && got.FKOSZTY.records[0].doctor === 'Lekarz Testowy', 'koszty');
ok(got.FSOURCES.records.length === 3 && got.FSOURCES.records[2].patient === 'ADAM INNY' && got.FSOURCES.records[0].source === 'Z sieci', 'sources');
ok(got.FSTATUS.records.find(r => r.raw === 'Pacjentów nie przyszło').status === 'Неявка' && got.FSTATUS.records.find(r => r.raw === 'Odbytych wizyt').status === 'Виконано' &&
   got.FSTATUS.records.find(r => r.raw === 'Odwołanych wizyt przez gabinet').status === 'Скасовано', 'fstatus');
ok(got.FDUR.records[0].hours === 1.25 && got.FDUR.records[1].hours === 0.75, 'fdur');
const z = got.FZAPISY.records;
ok(z.length === 5 && got.FZAPISY.skipped === 1, 'zapisy rows ' + z.length);
ok(z.map(x => x.status).join() === 'Перенесено,Виконано,Скасовано,Неявка,Видалено' && +z[1].date === +new Date(2026, 9, 8), 'zapisy statuses ' + z.map(x => x.status));
ok(+I.parseDate('09:35 07/10/2026') === +new Date(2026, 9, 7), 'felgdent datetime');

// Status wizyt oraz wpłat: одна послуга на рядок, тривалість лише в першому рядку візиту
const fw = got.FWIZYTY.records;
ok(fw.length === 5 && got.FWIZYTY.skipped === 4, 'fwizyty rows ' + fw.length + '/' + got.FWIZYTY.skipped);
ok(fw[0].service === 'Leczenie endodontyczne z wypełnieniem 3 kanałów' && fw[0].amount === 1900 && fw[0].paid === 1900 && fw[0].duration === 1.25 && fw[0].pos === '', 'fw first row');
ok(fw[1].duration === 0 && fw[1].pos === 'dodatkowa' && fw[1].paid === 0 && fw[1].payment === '', 'fw second row of visit');
ok(fw[2].status === 'Неявка' && fw[2].isNew === 'Так' && fw[2].service === '' && fw[3].isNew === 'Так' && fw[3].paid === 330, 'fw noshow / first visit');
ok(fw[4].status === 'Скасовано' && fw[4].typeDirection === 'Chirurgia', 'fw cancelled');
// Реєстр фактур: місяць з назви аркуша, оплата через ✓, лікар, ZUS без дати
const iv = got.INVOICES.records;
ok(iv.length === 5 && got.INVOICES.skipped === 2, 'register rows ' + iv.length + '/' + got.INVOICES.skipped);
ok(iv[0].paidDate && +iv[0].paidDate === +new Date(2026, 9, 8) && iv[1].paidDate && +iv[1].paidDate === +new Date(2026, 9, 17), 'paid via ✓ (1900 date ignored)');
ok(iv[2].seller === 'KRZYSZTOF' && iv[2].number === 'KRZYSZTOF 2026-10' && iv[2].paidDate === '', 'doctor invoice number fallback');
ok(iv[3].gross === 3994.14 && +iv[3].date === +new Date(2026, 9, 1), 'ZUS: amount text + month from sheet name');
ok(iv[4].due === null && I.categorize(iv[4].seller, '', -1, cfg.rules) === 'Laboratorium – płatność (звірка)', 'lab invoice');

console.log('OK — ' + n + ' перевірок пройдено');
