/**
 * ЛОГІКА ІМПОРТУ (без звернень до Google): розбір CSV, пошук колонок,
 * чисел і дат, перетворення рядків файлу у записи, категорії банку,
 * звірка фактур з банком. Цей файл тестується окремо (tests/test_import.js).
 */

var TYPY = {
  BANK: 'Виписка з банку',
  INVOICES: 'Фактури (реєстр)',
  VISITS: 'Felgdent – візити',
  PATIENTS: 'Felgdent – пацієнти',
  ZL: 'ZnanyLekarz – записи',
  BOOKSY: 'Booksy – записи',
  ADS: 'Google Ads – звіт',
  WPLATY: 'Felgdent – wpłaty (оплати)',
  FWIZYTY: 'Felgdent – status wizyt oraz wpłat',
  FKOSZTY: 'Felgdent – koszty (технік)',
  FSOURCES: 'Felgdent – pozyskanie pacjenta',
  FSTATUS: 'Felgdent – status wizyt',
  FDUR: 'Felgdent – średnia długość wizyty',
  FZAPISY: 'Felgdent – przesunięcia wizyt'
};

/** Нормалізація тексту: малі літери, без польських діакритиків, без зайвих пробілів. */
function norm(s) {
  if (s === null || s === undefined) return '';
  return String(s).toLowerCase()
    .replace(/ł/g, 'l')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[ \s]+/g, ' ')
    .replace(/["'`*]/g, '')
    .replace(/:+$/, '')
    .trim();
}

/** CSV → масив рядків. Роздільник (; , табуляція) визначається автоматично. */
function parseCsvText(text) {
  text = String(text || '').replace(/^﻿/, '');
  var sample = text.split(/\r?\n/).slice(0, 40);
  var best = ';', bestCount = -1;
  [';', ',', '\t'].forEach(function (d) {
    var c = 0;
    sample.forEach(function (line) {
      var inQ = false;
      for (var i = 0; i < line.length; i++) {
        var ch = line[i];
        if (ch === '"') inQ = !inQ;
        else if (ch === d && !inQ) c++;
      }
    });
    if (c > bestCount) { best = d; bestCount = c; }
  });
  var rows = [], row = [], cell = '', inQ = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cell += '"'; i++; } else inQ = false;
      } else cell += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === best) { row.push(cell.trim()); cell = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cell.trim()); rows.push(row); row = []; cell = '';
    } else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell.trim()); rows.push(row); }
  return rows.filter(function (r) { return r.some(function (c) { return c !== ''; }); });
}

/** Число з тексту: "1 234,56 zł", "-1.234,56", "1,234.56", "(250)", "250-" */
function parseNumber(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  if (!s) return null;
  var neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[  \s]/g, '').replace(/(zł|zl|pln|eur|usd|€|\$)/gi, '');
  if (/-$/.test(s)) { neg = !neg; s = s.slice(0, -1); }
  var lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc >= 0 && ld >= 0) {
    if (lc > ld) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (lc >= 0) {
    var commas = (s.match(/,/g) || []).length;
    if (commas === 1 && !/,\d{3}$/.test(s)) s = s.replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if ((s.match(/\./g) || []).length > 1) {
    s = s.replace(/\./g, '');
  }
  if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s)) return null;
  var n = parseFloat(s);
  return neg ? -n : n;
}

var MIESIACE = { sty: 1, jan: 1, lut: 2, feb: 2, mar: 3, kwi: 4, apr: 4, maj: 5, may: 5, cze: 6, jun: 6,
  lip: 7, jul: 7, sie: 8, aug: 8, wrz: 9, sep: 9, paz: 10, oct: 10, lis: 11, nov: 11, gru: 12, dec: 12 };

function mkDate(y, m, d) {
  y = +y; m = +m; d = +d;
  if (y < 100) y += 2000;
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 2000 || y > 2100) return null;
  var dt = new Date(y, m - 1, d);
  return dt.getMonth() === m - 1 ? dt : null;
}

/** Дата з тексту / числа Excel / об'єкта Date. */
function parseDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) || v.getFullYear() < 2000 ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) { var b = new Date(1899, 11, 30); b.setDate(b.getDate() + Math.floor(v)); return b; }
    return null;
  }
  var s = norm(v), m;
  if ((m = s.match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/))) return mkDate(m[1], m[2], m[3]);
  if ((m = s.match(/^(\d{1,2})[-.\/](\d{1,2})[-.\/](\d{4})/))) return mkDate(m[3], m[2], m[1]);
  if ((m = s.match(/^(\d{1,2})[-.\/](\d{1,2})[-.\/](\d{2})(?!\d)/))) return mkDate(m[3], m[2], m[1]);
  if ((m = s.match(/^(\d{4})[-.\/](\d{1,2})$/))) return mkDate(m[1], m[2], 1);
  if ((m = s.match(/^(\d{1,2})\s+([a-z]+)\.?,?\s+(\d{4})/)) && MIESIACE[m[2].slice(0, 3)]) return mkDate(m[3], MIESIACE[m[2].slice(0, 3)], m[1]);
  if ((m = s.match(/^([a-z]+)\.?\s+(\d{1,2}),?\s+(\d{4})/)) && MIESIACE[m[1].slice(0, 3)]) return mkDate(m[3], MIESIACE[m[1].slice(0, 3)], m[2]);
  if ((m = s.match(/^([a-z]+)\s+(\d{4})$/)) && MIESIACE[m[1].slice(0, 3)]) return mkDate(m[2], MIESIACE[m[1].slice(0, 3)], 1);
  // «09:35 07/10/2026» (Felgdent): дата не на початку
  if ((m = s.match(/(?:^|\s)(\d{1,2})[-.\/](\d{1,2})[-.\/](\d{4})(?!\d)/))) return mkDate(m[3], m[2], m[1]);
  if ((m = s.match(/(?:^|\s)(\d{4})-(\d{1,2})-(\d{1,2})(?!\d)/))) return mkDate(m[1], m[2], m[3]);
  return null;
}

/** Час доби у годинах: "10:30" → 10.5; Date → години; 0.4375 (частка доби) → 10.5 */
function parseTime(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v.getHours() + v.getMinutes() / 60;
  if (typeof v === 'number') return v < 1 ? v * 24 : null;
  var m = String(v).match(/(\d{1,2})[:.](\d{2})/);
  return m ? +m[1] + +m[2] / 60 : null;
}

/** Тривалість візиту в годинах. */
function parseDuration(dur, start, end) {
  if (dur !== null && dur !== undefined && dur !== '') {
    if (typeof dur === 'number') return dur < 1 ? dur * 24 : (dur >= 5 ? dur / 60 : dur);
    if (Object.prototype.toString.call(dur) === '[object Date]') return dur.getHours() + dur.getMinutes() / 60;
    var s = String(dur), m;
    if ((m = s.match(/^\s*(\d{1,2}):(\d{2})/))) return +m[1] + +m[2] / 60;
    var n = parseNumber(s.replace(/min\.?|minut[ay]?|h|godz\.?/gi, ''));
    if (n !== null) return /min/i.test(s) || n >= 5 ? n / 60 : n;
  }
  var a = parseTime(start), b = parseTime(end);
  if (a !== null && b !== null && b > a) return Math.round((b - a) * 100) / 100;
  return null;
}

/** Призначення колонок файлу до полів: спочатку точний збіг назви, потім «містить». */
function assignColumns(headers, fields) {
  var map = {}, used = {};
  fields.forEach(function (f) {
    for (var a = 0; a < f.aliases.length && map[f.name] === undefined; a++) {
      var idx = headers.indexOf(f.aliases[a]);
      if (idx >= 0 && !used[idx]) { map[f.name] = idx; used[idx] = true; }
    }
  });
  fields.forEach(function (f) {
    if (map[f.name] !== undefined) return;
    for (var a = 0; a < f.aliases.length && map[f.name] === undefined; a++) {
      var al = f.aliases[a];
      if (al.length < 5) continue;
      for (var i = 0; i < headers.length; i++) {
        if (!used[i] && headers[i].length >= 3 && headers[i].indexOf(al) >= 0) { map[f.name] = i; used[i] = true; break; }
      }
    }
  });
  return map;
}

function requiredOk(map, fields) {
  var anyGroup = fields.filter(function (f) { return f.required === 'один із'; });
  var okReq = fields.every(function (f) { return f.required !== 'так' || map[f.name] !== undefined; });
  var okAny = !anyGroup.length || anyGroup.some(function (f) { return map[f.name] !== undefined; });
  return okReq && okAny;
}

/** Шукає рядок заголовків у перших 30 рядках файлу. */
function findHeader(rows, fields) {
  var best = null;
  for (var r = 0; r < Math.min(rows.length, 30); r++) {
    var headers = rows[r].map(norm);
    var map = assignColumns(headers, fields);
    if (!requiredOk(map, fields)) continue;
    var score = Object.keys(map).length;
    if (!best || score > best.score) best = { row: r, map: map, score: score };
  }
  return best;
}

/** Визначає тип файлу за назвою файлу та заголовками колонок. */
function detectType(rows, fieldsByType, hints, fileName) {
  var fn = norm(String(fileName || '').replace(/[_.\-]+/g, ' '));
  var cands = [];
  Object.keys(fieldsByType).forEach(function (t) {
    var h = findHeader(rows, fieldsByType[t]);
    if (h) cands.push({ type: t, header: h });
  });
  if (!cands.length) return { error: 'Не вдалося розпізнати тип файлу: не знайдено потрібних колонок. Виберіть тип вручну або додайте назви колонок на аркуші «Імпорт-налаштування».' };
  for (var i = 0; i < hints.length; i++) {
    if (hints[i].fragments.some(function (fr) { return fr && fn.indexOf(fr) >= 0; })) {
      var c = cands.filter(function (x) { return x.type === hints[i].type; })[0];
      if (c) return c;
    }
  }
  cands.sort(function (a, b) { return b.header.score - a.header.score; });
  if (cands.length > 1 && cands[0].header.score === cands[1].header.score) {
    var pair = [cands[0].type, cands[1].type].sort().join(',');
    if (pair === 'BOOKSY,ZL') return { error: 'Не зрозуміло, це ZnanyLekarz чи Booksy. Виберіть тип вручну або додайте в назву файлу «znany» чи «booksy».' };
  }
  return cands[0];
}

function mapStatus(raw, rules) {
  var s = norm(raw);
  if (!s) return null;
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].fragments.some(function (fr) { return fr && s.indexOf(fr) >= 0; })) return rules[i].status;
  }
  return null;
}

function mapPayment(raw) {
  var s = norm(raw);
  if (!s) return '';
  var t = [['payu', 'Raty PayU'], ['medirat', 'Mediraty'], ['medi rat', 'Mediraty'], ['rat', 'Raty PayU'], ['blik', 'BLIK'],
    ['czek', 'Czek'], ['mieszan', 'Wpłata mieszana'], ['kart', 'Karta'], ['terminal', 'Karta'], ['gotow', 'Gotówka'], ['przelew', 'Przelew'], ['bon', 'Bon podarunkowy'],
    ['voucher', 'Bon podarunkowy'], ['przedplat', 'Przedpłata (wcześniej)'], ['zaliczk', 'Przedpłata (wcześniej)']];
  for (var i = 0; i < t.length; i++) if (s.indexOf(t[i][0]) >= 0) return t[i][1];
  return String(raw).trim();
}

/**
 * Напрям з назви лікаря у Felgdent: «ENDODONTA Ім'я Прізвище» → Endodoncja.
 * Якщо в назві кілька спеціалізацій («ZACH/PROT/CHIR …»), повертає '' — тоді напрям береться з аркуша «Послуги».
 */
var PREFIKSY = [['endo', 'Endodoncja'], ['higien', 'Profilaktyka'], ['profil', 'Profilaktyka'], ['orto', 'Ortodoncja'],
  ['implant', 'Implantologia'], ['chir', 'Chirurgia'], ['prot', 'Protetyka'], ['pedo', 'Dziecięca'], ['dziec', 'Dziecięca'], ['zach', 'Zachowawcza']];
function directionFromDoctor(name) {
  var first = norm(name).split(' ')[0] || '';
  var parts = first.split('/').filter(String), found = [];
  parts.forEach(function (p) {
    for (var i = 0; i < PREFIKSY.length; i++) if (p.indexOf(PREFIKSY[i][0]) === 0) { if (found.indexOf(PREFIKSY[i][1]) < 0) found.push(PREFIKSY[i][1]); break; }
  });
  return found.length === 1 ? found[0] : '';
}

/** Напрям за назвою послуги Felgdent (ключові слова). Порядок = пріоритет, якщо в оплаті кілька послуг. */
var SLOWA_NAPRAMKU = [
  ['Implantologia', ['implant', 'augmentac', 'sinus lift', 'podniesienie dna zatoki', 'membran', 'kosc']],
  ['Ortodoncja', ['ortodon', 'aparat', 'zamek', 'luk ortod', 'retencj', 'nakladk', 'aligner']],
  ['Protetyka', ['koron', 'most', 'protez', 'licowk', 'wklad', 'onlay', 'inlay', 'overlay', 'tymczasow', 'wycisk', 'cyrkon', 'pmma']],
  ['Endodoncja', ['kanal', 'endo', 'dewital', 'ekstyrpac', 'mikroskop', 'rewizj', 'opracowanie kanal', 'wypelnienie kanal']],
  ['Chirurgia', ['ekstrakc', 'usuniecie zeba', 'usuniecie korzen', 'chirurg', 'resekc', 'nacieci', 'szwy', 'osemk', 'dluto']],
  ['Dziecięca', ['mleczn', 'dzieci', 'lakow', 'lapis', 'adaptac']],
  ['Profilaktyka', ['higieniz', 'skaling', 'scaling', 'piaskow', 'polerow', 'fluoryz', 'profilakty', 'kamien']],
  ['Zachowawcza', ['wypelnien', 'ubyt', 'odbudow', 'bol', 'doraz', 'konsultac', 'przeglad', 'leczenie zmian', 'laser', 'rtg', 'pantomogr', 'znieczul', 'wizyta']]
];
/** Однозначні слова, які перевіряються першими: «Wypełnienie ubytku korony zęba» — це пломба, а не коронка. */
var SLOWA_PEWNE = [['Endodoncja', ['kanal', 'endodont']], ['Zachowawcza', ['ubyt', 'wypelnienie ubytk']], ['Profilaktyka', ['higieniz', 'skaling']]];
function directionFromService(text) {
  var t = ' ' + norm(text) + ' ';
  for (var j = 0; j < SLOWA_PEWNE.length; j++) {
    if (SLOWA_PEWNE[j][1].some(function (w) { return t.indexOf(w) >= 0; })) return SLOWA_PEWNE[j][0];
  }
  for (var i = 0; i < SLOWA_NAPRAMKU.length; i++) {
    if (SLOWA_NAPRAMKU[i][1].some(function (w) { return t.indexOf(w) >= 0; })) return SLOWA_NAPRAMKU[i][0];
  }
  return '';
}
var PRIORYTET_NAPRAMKU = SLOWA_NAPRAMKU.map(function (x) { return x[0]; });

/** Напрям з «Typ wizyty» у Felgdent (якщо в рядку немає послуги). */
function directionFromType(typ) {
  var t = norm(typ);
  var map = [['endodon', 'Endodoncja'], ['chirurg', 'Chirurgia'], ['higien', 'Profilaktyka'], ['ortodon', 'Ortodoncja'], ['protet', 'Protetyka'],
    ['implant', 'Implantologia'], ['pedodon', 'Dziecięca'], ['periodon', 'Zachowawcza'], ['zachowaw', 'Zachowawcza'], ['wybiel', 'Profilaktyka'],
    ['estety', 'Zachowawcza'], ['awaryj', 'Zachowawcza'], ['kontrol', 'Zachowawcza'], ['konsult', 'Zachowawcza'], ['pierwsza', 'Zachowawcza']];
  for (var i = 0; i < map.length; i++) if (t.indexOf(map[i][0]) >= 0) return map[i][1];
  return '';
}

/** Рядки послуг з поля «Opis» Felgdent: «Wypełnienie … (25)\nKorona … (12)» → без номерів зубів. */
function splitServices(opis) {
  return String(opis || '').split(/\r?\n|;/).map(function (x) { return x.replace(/\s*\([\d,\s\-]+\)\s*$/, '').trim(); }).filter(String);
}

/** Статус з «Raport przesunięć wizyt» Felgdent → статус у звіті. */
function felgdentStatus(raw) {
  var s = norm(raw);
  if (!s) return '';
  if (s === '---' || s === '-') return 'Перенесено';
  if (s.indexOf('odbyla') >= 0 || s.indexOf('poczekalni') >= 0 || s.indexOf('odbyt') >= 0) return 'Виконано';
  if (s.indexOf('nie przysz') >= 0) return 'Неявка';
  if (s.indexOf('odwol') >= 0) return 'Скасовано';
  if (s.indexOf('usunieta') >= 0 || s.indexOf('usunie') >= 0) return 'Видалено';
  if (s.indexOf('niepotwierdz') >= 0 || s.indexOf('nie odpowiada') >= 0 || s.indexOf('potwierdz') >= 0) return 'Заплановано';
  return String(raw).trim();
}

function ymd(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}

function yesLike(v) {
  if (v === true) return true;
  var s = norm(v);
  return ['tak', 'yes', 'true', '1', 'x', 'nowy', 'nowa', 'так'].indexOf(s) >= 0;
}

var SUMY = ['razem', 'total', 'suma', 'lacznie', 'ogolem', 'wszystkie', 'итого', 'разом'];

/**
 * Рядки файлу → записи для запису в таблицю.
 * opts: {statusRules, month (Date для Google Ads без дат)}
 */
function mapRecords(type, rows, header, opts) {
  opts = opts || {};
  var recs = [], warnings = [], skipped = 0, seen = {};
  var m = header.map;
  function g(row, f) { return m[f] === undefined ? undefined : row[m[f]]; }
  function str(v) { return v === undefined || v === null ? '' : String(v).replace(/[ \u00a0]{2,}/g, ' ').trim(); }
  function uniq(key) { seen[key] = (seen[key] || 0) + 1; return key + '#' + seen[key]; }
  var unknownStatuses = {}, visitSeen = {};
  if (type === 'BANK') {
    var hdrRow = rows[header.row].map(norm), mapped = {};
    Object.keys(m).forEach(function (k) { mapped[m[k]] = true; });
    header.extra = [];
    hdrRow.forEach(function (h, ix) { if (!h && !mapped[ix]) header.extra.push(ix); });
  }

  for (var i = header.row + 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row || !row.some(function (c) { return c !== '' && c !== null; })) continue;
    var first = norm(row[0]);
    if (SUMY.some(function (w) { return first.indexOf(w) === 0; })) { skipped++; continue; }
    var rec = null;

    if (type === 'BANK') {
      var d = parseDate(g(row, 'date'));
      var amt;
      if (m.amount !== undefined) amt = parseNumber(g(row, 'amount'));
      else {
        var deb = parseNumber(g(row, 'debit')), cre = parseNumber(g(row, 'credit'));
        amt = (deb ? -Math.abs(deb) : 0) + (cre ? Math.abs(cre) : 0);
        if (!deb && !cre) amt = null;
      }
      if (!d || amt === null || amt === 0) { skipped++; continue; }
      var cp = str(g(row, 'counterparty')), ds = str(g(row, 'description'));
      // деякі банки (напр. PKO) кладуть контрагента й tytuł у колонки без назви
      (header.extra || []).forEach(function (ix) { if (str(row[ix])) ds += ' ' + str(row[ix]); });
      rec = { date: d, amount: Math.round(amt * 100) / 100, counterparty: cp, description: ds, account: str(g(row, 'account')) };
      rec.key = uniq(['B', ymd(d), rec.amount.toFixed(2), norm(cp).slice(0, 40), norm(ds).slice(0, 60)].join('|'));

    } else if (type === 'INVOICES') {
      var seller = str(g(row, 'seller')).replace(/\s+$/, '');
      var gross = parseNumber(g(row, 'gross'));
      var rawNum = g(row, 'number');
      var num = Object.prototype.toString.call(rawNum) === '[object Date]' ? '' : str(rawNum);
      var d2 = parseDate(g(row, 'date')) || opts.month || null;   // зарплати / ZUS без дати → місяць аркуша
      if (!seller || gross === null || gross === 0 || !d2) { skipped++; continue; }
      if (!num) num = seller + ' ' + ymd(d2).slice(0, 7);
      var nip = str(g(row, 'nip')).replace(/\D/g, '');
      var paid = parseDate(g(row, 'paidDate'));
      var flag = g(row, 'paidFlag');
      if (!paid && (flag === true || yesLike(flag) || norm(flag) === '✓')) paid = parseDate(g(row, 'due')) || d2;
      var notes = str(g(row, 'notes'));
      rec = { number: num, date: d2, seller: seller, nip: nip, gross: gross, due: parseDate(g(row, 'due')),
        category: str(g(row, 'category')), paidDate: paid || '', notes: notes };
      rec.key = uniq('F|' + norm(num).replace(/\s/g, '') + '|' + (nip || norm(seller).slice(0, 30)));

    } else if (type === 'VISITS') {
      var d3 = parseDate(g(row, 'date')), pat = str(g(row, 'patient'));
      if (!d3 || !pat) { skipped++; continue; }
      var st = 'Виконано';
      if (m.status !== undefined) {
        var raw = str(g(row, 'status'));
        st = mapStatus(raw, opts.statusRules || []);
        if (st === 'Пропустити') { skipped++; continue; }
        if (!st) { st = raw || 'Виконано'; if (raw) unknownStatuses[raw] = true; }
      }
      rec = { date: d3, patient: pat, doctor: str(g(row, 'doctor')), service: str(g(row, 'service')),
        amount: parseNumber(g(row, 'amount')) || 0, status: st,
        duration: parseDuration(g(row, 'duration'), g(row, 'start'), g(row, 'end')),
        lab: parseNumber(g(row, 'lab')), prep: parseNumber(g(row, 'prep')),
        payment: mapPayment(g(row, 'payment')), direction: str(g(row, 'direction')),
        isNew: m.isNew === undefined ? null : (yesLike(g(row, 'isNew')) ? 'Так' : 'Ні'),
        source: str(g(row, 'source')) };
      rec.key = uniq(['V', ymd(d3), norm(pat), norm(rec.doctor), norm(rec.service)].join('|'));

    } else if (type === 'PATIENTS') {
      var p = str(g(row, 'patient'));
      if (!p) { skipped++; continue; }
      rec = { patient: p, regDate: parseDate(g(row, 'regDate')), source: str(g(row, 'source')) };
      rec.key = 'P|' + norm(p);

    } else if (type === 'ZL' || type === 'BOOKSY') {
      var d4 = parseDate(g(row, 'date')), p2 = str(g(row, 'patient'));
      if (!d4 || !p2) { skipped++; continue; }
      var platform = type === 'ZL' ? 'ZnanyLekarz' : 'Booksy';
      var st2 = mapStatus(g(row, 'status'), opts.statusRules || []) || (str(g(row, 'status')) || 'Заплановано');
      if (st2 === 'Пропустити') st2 = 'Заплановано';
      rec = { date: d4, platform: platform, patient: p2, doctor: str(g(row, 'doctor')), service: str(g(row, 'service')), status: st2 };
      rec.key = uniq([platform, ymd(d4), norm(p2), norm(rec.service)].join('|'));

    } else if (type === 'ADS') {
      var camp = str(g(row, 'campaign'));
      var nc = norm(camp);
      if (SUMY.some(function (w) { return nc.indexOf(w) === 0; }) || nc === '--') { skipped++; continue; }
      var d5 = parseDate(g(row, 'date')) || opts.month || null;
      var cost = parseNumber(g(row, 'cost'));
      if (!d5 || cost === null) { skipped++; if (!d5) warnings.push('Google Ads: у файлі немає дат — вкажіть місяць звіту при завантаженні.'); continue; }
      if (!camp && m.date === undefined) { skipped++; continue; }
      rec = { date: d5, campaign: camp || '(усі кампанії)', impressions: parseNumber(g(row, 'impressions')) || 0,
        clicks: parseNumber(g(row, 'clicks')) || 0, cost: cost, conversions: parseNumber(g(row, 'conversions')) || 0 };
      rec.key = 'A|' + ymd(d5) + '|' + norm(rec.campaign);

    } else if (type === 'FWIZYTY') {
      var t12 = str(g(row, 'start'));
      if (!/^\d{1,2}:\d{2}$/.test(t12)) { skipped++; continue; }   // розділи «Wpisy…», «Wpis bez wizyty»
      var d12 = parseDate(g(row, 'date')), p12 = str(g(row, 'patient'));
      if (!d12 || !p12) { skipped++; continue; }
      var raw12 = str(g(row, 'status'));
      var st12 = mapStatus(raw12, opts.statusRules || []);
      if (st12 === 'Пропустити') { skipped++; continue; }
      if (!st12) { st12 = raw12; unknownStatuses[raw12] = true; }
      var el = str(g(row, 'elements'));
      var svc = /^wizyta bez|^brak/i.test(norm(el)) ? '' : el.replace(/^\d+(\.\d+)?\s+/, '');
      var doc12 = str(g(row, 'doctor'));
      var vkey = [ymd(d12), t12, norm(p12), norm(doc12)].join('|');
      var firstRow = !visitSeen[vkey]; visitSeen[vkey] = true;
      var endIdx = m.start + 1;
      var typ = str(g(row, 'type'));
      var a12 = parseNumber(g(row, 'amount')), paid12 = parseNumber(g(row, 'paid'));
      rec = { date: d12, patient: p12, doctor: doc12, service: svc, amount: a12 || 0, status: st12,
        duration: firstRow ? (parseDuration(null, t12, row[endIdx]) || 0) : 0,
        lab: null, prep: null, payment: mapPayment(g(row, 'payment') === '---' || g(row, 'payment') === '-' ? '' : g(row, 'payment')),
        direction: '', typeDirection: directionFromType(typ), isNew: /pierwsza/.test(norm(typ)) ? 'Так' : 'Ні', source: '',
        paid: paid12 === null ? '' : paid12, pos: firstRow ? '' : 'dodatkowa' };
      rec.key = uniq(['FW', ymd(d12), t12, norm(p12), norm(doc12), norm(svc)].join('|'));

    } else if (type === 'WPLATY') {
      var d6 = parseDate(g(row, 'date')), p6 = str(g(row, 'patient')), a6 = parseNumber(g(row, 'amount'));
      if (!d6 || a6 === null) { skipped++; continue; }
      var lines = splitServices(g(row, 'description'));
      rec = { date: d6, patient: p6 || '(без пацієнта)', doctor: '', addedBy: str(g(row, 'addedBy')), service: lines.join('; '), lines: lines,
        amount: a6, status: 'Виконано', duration: null, lab: null, prep: null, payment: mapPayment(g(row, 'payment')),
        direction: '', isNew: null, source: '' };
      rec.key = uniq(['W', ymd(d6), norm(p6), a6.toFixed(2), norm(rec.service).slice(0, 60)].join('|'));

    } else if (type === 'FKOSZTY') {
      var d7 = parseDate(g(row, 'date')), a7 = parseNumber(g(row, 'amount'));
      if (!d7 || a7 === null) { skipped++; continue; }
      var opis7 = str(g(row, 'description'));
      rec = { date: d7, doctor: str(g(row, 'addedBy')), description: opis7, amount: Math.abs(a7), payment: mapPayment(g(row, 'payment')),
        category: 'Koszty z Felgdent (kontrola)' };
      rec.key = uniq(['K', ymd(d7), norm(rec.doctor), rec.amount.toFixed(2), norm(opis7).slice(0, 40)].join('|'));

    } else if (type === 'FSOURCES') {
      var p8 = (str(g(row, 'firstName')) + ' ' + str(g(row, 'lastName'))).trim();
      if (!p8) { skipped++; continue; }
      rec = { patient: p8, source: str(g(row, 'source')), regDate: undefined };
      rec.key = 'P|' + norm(p8);

    } else if (type === 'FSTATUS') {
      var raw9 = str(g(row, 'status')), c9 = parseNumber(g(row, 'count'));
      if (!raw9 || c9 === null) { skipped++; continue; }
      if (!opts.month) { skipped++; warnings.push('Status wizyt: у файлі немає місяця — вкажіть «Місяць звіту» при завантаженні.'); continue; }
      var st9 = mapStatus(raw9, opts.statusRules || []);
      if (st9 === 'Пропустити') st9 = 'Заплановано';
      rec = { month: opts.month, raw: raw9, status: st9 || raw9, count: c9 };
      rec.key = 'S|' + ymd(opts.month).slice(0, 7) + '|' + norm(raw9);

    } else if (type === 'FDUR') {
      var doc10 = str(g(row, 'doctor')), h10 = parseDuration(g(row, 'avg'));
      if (!doc10 || h10 === null) { skipped++; continue; }
      rec = { doctor: doc10, hours: Math.round(h10 * 100) / 100 };

    } else if (type === 'FZAPISY') {
      var p11 = str(g(row, 'patient'));
      if (!p11) { skipped++; continue; }   // порожні рядки = блокування часу в графіку
      var orig = parseDate(g(row, 'orig')), nw = parseDate(g(row, 'newTerm'));
      var term = nw || orig;
      if (!term) { skipped++; continue; }
      var raw11 = str(g(row, 'status'));
      rec = { date: term, patient: p11, status: felgdentStatus(raw11), raw: raw11, moved: nw ? 'Так' : 'Ні', changedBy: str(g(row, 'changedBy')) };
      rec.key = uniq(['Z', str(g(row, 'added')), str(g(row, 'orig')), norm(p11)].join('|'));
    }
    if (rec) recs.push(rec);
  }
  var u = Object.keys(unknownStatuses);
  if (u.length) warnings.push('Невідомі статуси візитів (додайте на аркуш «Імпорт-налаштування»): ' + u.join(', '));
  return { records: recs, skipped: skipped, warnings: warnings.filter(function (w, i, a) { return a.indexOf(w) === i; }) };
}

/** Правило банку: {kw (нормалізоване), dir: 'out'|'in'|'any', category} */
function dirFromText(t) {
  var s = norm(t);
  if (s.indexOf('витр') === 0) return 'out';
  if (s.indexOf('надх') === 0) return 'in';
  return 'any';
}

function categorize(counterparty, description, amount, rules) {
  var text = ' ' + norm(counterparty + ' ' + description) + ' ';
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    if (!r.kw || text.indexOf(r.kw) < 0) continue;
    if (r.dir === 'out' && !(amount < 0)) continue;
    if (r.dir === 'in' && !(amount > 0)) continue;
    return r.category;
  }
  return amount > 0 ? 'Надходження: інше' : '';
}

/**
 * Звірка фактур з банком.
 * invoices: [{idx, number, gross, date, seller, nip}], bank: [{idx, date, amount, counterparty, description}]
 * Повертає [{invIdx, bankIdx, date}]
 */
function matchInvoices(invoices, bank, usedBank) {
  var used = {};
  (usedBank || []).forEach(function (i) { used[i] = true; });
  var out = [];
  var DAY = 86400000;
  var outflows = bank.filter(function (b) { return b.amount < 0; }).map(function (b) {
    return { idx: b.idx, date: b.date, amount: -b.amount,
      text: norm(b.counterparty + ' ' + b.description), flat: norm(b.counterparty + ' ' + b.description).replace(/[\s\/\-_.]/g, '') };
  });
  function inWindow(inv, b) {
    return b.date.getTime() >= inv.date.getTime() - 5 * DAY && b.date.getTime() <= inv.date.getTime() + 120 * DAY;
  }
  var pending = invoices.slice();
  // 1) номер фактури в тексті переказу
  pending = pending.filter(function (inv) {
    var num = norm(inv.number).replace(/[\s\/\-_.]/g, '');
    if (num.length < 4) return true;
    for (var i = 0; i < outflows.length; i++) {
      var b = outflows[i];
      if (!used[b.idx] && b.flat.indexOf(num) >= 0) { used[b.idx] = true; out.push({ invIdx: inv.idx, bankIdx: b.idx, date: b.date }); return false; }
    }
    return true;
  });
  // 2) та сама сума + продавець або NIP у тексті + вікно дат
  pending = pending.filter(function (inv) {
    var words = norm(inv.seller).split(/[\s,.]+/).filter(function (w) { return w.length >= 4 && ['spolka', 'polska', 'z o o', 'sp z o o'].indexOf(w) < 0; });
    for (var i = 0; i < outflows.length; i++) {
      var b = outflows[i];
      if (used[b.idx] || Math.abs(b.amount - inv.gross) > 0.01 || !inWindow(inv, b)) continue;
      var who = (inv.nip && b.flat.indexOf(inv.nip) >= 0) || words.some(function (w) { return b.text.indexOf(w) >= 0; });
      if (who) { used[b.idx] = true; out.push({ invIdx: inv.idx, bankIdx: b.idx, date: b.date }); return false; }
    }
    return true;
  });
  // 3) лише сума + дата, якщо кандидат єдиний і сума ≥ 100 zł
  pending.forEach(function (inv) {
    if (inv.gross < 100) return;
    var c = outflows.filter(function (b) { return !used[b.idx] && Math.abs(b.amount - inv.gross) <= 0.01 && inWindow(inv, b); });
    if (c.length === 1) { used[c[0].idx] = true; out.push({ invIdx: inv.idx, bankIdx: c[0].idx, date: c[0].date }); }
  });
  return out;
}

if (typeof module !== 'undefined') {
  module.exports = { norm: norm, parseCsvText: parseCsvText, parseNumber: parseNumber, parseDate: parseDate, parseDuration: parseDuration,
    findHeader: findHeader, detectType: detectType, mapRecords: mapRecords, categorize: categorize, dirFromText: dirFromText,
    matchInvoices: matchInvoices, mapStatus: mapStatus, mapPayment: mapPayment, directionFromDoctor: directionFromDoctor,
    directionFromService: directionFromService, directionFromType: directionFromType, splitServices: splitServices, felgdentStatus: felgdentStatus, PRIORYTET_NAPRAMKU: PRIORYTET_NAPRAMKU };
}
