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

console.log('OK — ' + n + ' перевірок пройдено');
