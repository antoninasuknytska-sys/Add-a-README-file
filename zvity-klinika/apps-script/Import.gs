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
  ADS: 'Google Ads – звіт'
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
    return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
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
  var fn = norm(fileName);
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
    ['kart', 'Karta'], ['terminal', 'Karta'], ['gotow', 'Gotówka'], ['przelew', 'Przelew'], ['bon', 'Bon podarunkowy'],
    ['voucher', 'Bon podarunkowy'], ['przedplat', 'Przedpłata (wcześniej)'], ['zaliczk', 'Przedpłata (wcześniej)']];
  for (var i = 0; i < t.length; i++) if (s.indexOf(t[i][0]) >= 0) return t[i][1];
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
  function str(v) { return v === undefined || v === null ? '' : String(v).trim(); }
  function uniq(key) { seen[key] = (seen[key] || 0) + 1; return key + '#' + seen[key]; }
  var unknownStatuses = {};
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
      var num = str(g(row, 'number')), d2 = parseDate(g(row, 'date')), gross = parseNumber(g(row, 'gross'));
      if (!num || !d2 || gross === null) { skipped++; continue; }
      var nip = str(g(row, 'nip')).replace(/\D/g, '');
      rec = { number: num, date: d2, seller: str(g(row, 'seller')), nip: nip, gross: Math.abs(gross),
        due: parseDate(g(row, 'due')), category: str(g(row, 'category')) };
      rec.key = 'F|' + norm(num).replace(/\s/g, '') + '|' + (nip || norm(rec.seller).slice(0, 30));

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
    matchInvoices: matchInvoices, mapStatus: mapStatus, mapPayment: mapPayment };
}
