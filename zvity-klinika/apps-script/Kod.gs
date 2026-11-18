/**
 * МЕНЮ, ЗАВАНТАЖЕННЯ ФАЙЛІВ І ЗАПИС У ТАБЛИЦЮ.
 * Логіка розбору файлів — у Import.gs, розсилка звіту — у Raport.gs.
 */

var FOLDER_IMPORT = 'Звіти клініки – Імпорт';
var FOLDER_DONE = 'Оброблено';

var UKLAD = {
  VISITS:   { sheet: 'Візити', last: 3001, keyCol: 18,
              cols: { date: 1, patient: 3, isNew: 4, source: 5, doctor: 6, direction: 7, service: 8, amount: 9, lab: 10, prep: 11, payment: 12, status: 13, duration: 14 } },
  BANK:     { sheet: 'Банк', last: 3001, keyCol: 11,
              cols: { date: 1, amount: 3, counterparty: 4, description: 5, autoCat: 6, manualCat: 7, account: 10 } },
  INVOICES: { sheet: 'Фактури', last: 3001, keyCol: 11,
              cols: { number: 1, date: 2, seller: 3, nip: 4, gross: 5, due: 6, category: 7, paidDate: 8 } },
  PATIENTS: { sheet: 'Пацієнти', last: 3001, keyCol: 4, cols: { patient: 1, regDate: 2, source: 3 } },
  ONLINE:   { sheet: 'Онлайн-записи', last: 3001, keyCol: 8,
              cols: { date: 1, platform: 3, patient: 4, doctor: 5, service: 6, status: 7 } },
  ADS:      { sheet: 'Google Ads', last: 3001, keyCol: 8,
              cols: { date: 1, campaign: 3, impressions: 4, clicks: 5, cost: 6, conversions: 7 } },
  EXPENSES: { sheet: 'Витрати', last: 3001, keyCol: 7, cols: { date: 1, category: 3, description: 5, amount: 6, payment: 7 } },
  STAFF:    { sheet: 'Персонал', last: 500, keyCol: 6, cols: { date: 1, name: 3, role: 4, hours: 5, rate: 6 } }
};

// ------------------------------------------------------------------ меню
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🦷 Звіти клініки')
    .addItem('Завантажити файли…', 'pokazDialog')
    .addItem('Імпортувати з папки Google Drive', 'importZFolderuMenu')
    .addSeparator()
    .addItem('Перерахувати категорії банку', 'przeliczKategorieMenu')
    .addItem('Звірити фактури з банком', 'zestawFakturyMenu')
    .addItem('Оновити джерела й нових пацієнтів у візитах', 'uzupelnijZrodlaMenu')
    .addSeparator()
    .addItem('Надіслати звіт зараз', 'wyslijRaport')
    .addItem('Увімкнути автоматику (імпорт щоночі, звіт 15-го о 21:00)', 'ustawHarmonogram')
    .addSeparator()
    .addItem('Очистити демо-дані', 'wyczyscDemo')
    .addToUi();
}

function pokazDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog').setWidth(560).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Завантаження файлів');
}

function alert_(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); } }

// ------------------------------------------------------------------ імпорт
/** Викликається з вікна завантаження: {name, mime, b64, type, month: 'yyyy-mm'} */
function importujPlik(p) {
  var blob = Utilities.newBlob(Utilities.base64Decode(p.b64), p.mime || 'application/octet-stream', p.name);
  return importujBlob(blob, p.type, p.month ? parseDate(p.month) : null);
}

function importujBlob(blob, forcedType, month) {
  var name = blob.getName();
  var lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {
    var rows = czytajTabele(blob);
    var cfg = wczytajKonfig();
    var type, hdr;
    if (forcedType && forcedType !== 'AUTO') {
      type = forcedType;
      hdr = findHeader(rows, cfg.fields[type]);
      if (!hdr) throw new Error('У файлі не знайдено обов\'язкових колонок для типу «' + TYPY[type] + '». Очікувані назви: ' +
        cfg.fields[type].filter(function (f) { return f.required !== 'ні'; }).map(function (f) { return f.aliases.slice(0, 3).join(' / '); }).join('; '));
    } else {
      var det = detectType(rows, cfg.fields, cfg.hints, name);
      if (det.error) throw new Error(det.error);
      type = det.type; hdr = det.header;
    }
    var res = mapRecords(type, rows, hdr, { statusRules: cfg.statuses, month: month || null });
    var st = zapisz_(type, res.records, cfg);
    if (type === 'BANK') { przeliczKategorie_(cfg); zestawFaktury_(); }
    if (type === 'INVOICES') { kategorieFaktur_(cfg); zestawFaktury_(); }
    if (type === 'ZL' || type === 'BOOKSY' || type === 'PATIENTS') uzupelnijZrodla_();
    var warn = res.warnings.concat(st.warnings || []);
    dziennik_(name, TYPY[type], st.added, st.updated, res.skipped, warn.join(' | '));
    return '✅ ' + name + ' → ' + TYPY[type] + ': додано ' + st.added + ', оновлено ' + st.updated + ', пропущено ' + res.skipped +
      (warn.length ? '\n⚠ ' + warn.join('\n⚠ ') : '');
  } catch (e) {
    dziennik_(name, '—', 0, 0, 0, 'ПОМИЛКА: ' + e.message);
    throw new Error(name + ': ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/** Файл → масив рядків. CSV/TXT читаємо напряму (UTF-8 або Windows-1250), XLSX/XLS/ODS конвертуємо через Drive. */
function czytajTabele(blob) {
  var name = (blob.getName() || '').toLowerCase();
  var mime = blob.getContentType() || '';
  if (/\.pdf$/.test(name) || mime === 'application/pdf') throw new Error('PDF не підтримується. Вивантажте з банку / програми файл CSV або XLSX.');
  if (/\.(csv|txt)$/.test(name) || /text\//.test(mime)) {
    var t = blob.getDataAsString('UTF-8');
    if (t.indexOf('�') >= 0) t = blob.getDataAsString('windows-1250');
    return parseCsvText(t);
  }
  if (/\.(xlsx|xls|ods)$/.test(name) || /spreadsheet|excel/.test(mime)) {
    var f = Drive.Files.create({ name: 'tmp_import_' + name, mimeType: MimeType.GOOGLE_SHEETS }, blob);
    try {
      var ss = SpreadsheetApp.openById(f.id);
      var best = null;
      ss.getSheets().forEach(function (sh) {
        var v = sh.getDataRange().getValues();
        if (!best || v.length > best.length) best = v;
      });
      return best.map(function (r) { return r.map(function (c) { return typeof c === 'string' ? c.trim() : c; }); });
    } finally {
      DriveApp.getFileById(f.id).setTrashed(true);
    }
  }
  throw new Error('Невідомий формат файлу. Підтримуються CSV, TXT, XLSX, XLS, ODS.');
}

/** Налаштування з аркушів «Імпорт-налаштування» та «Правила банку». */
function wczytajKonfig() {
  var ss = SpreadsheetApp.getActive();
  var ic = ss.getSheetByName('Імпорт-налаштування');
  var fields = {};
  ic.getRange('A4:D200').getValues().forEach(function (r) {
    if (!r[0] || !r[1]) return;
    var t = String(r[0]).trim();
    (fields[t] = fields[t] || []).push({ name: String(r[1]).trim(), required: norm(r[2]) === 'так' ? 'так' : (norm(r[2]) === 'один із' ? 'один із' : 'ні'),
      aliases: String(r[3]).split(';').map(norm).filter(String) });
  });
  var statuses = ic.getRange('G4:H40').getValues().filter(function (r) { return r[0] && r[1]; }).map(function (r) {
    return { fragments: String(r[0]).split(';').map(norm).filter(String), status: String(r[1]).trim() };
  });
  var hints = ic.getRange('J4:K30').getValues().filter(function (r) { return r[0] && r[1]; }).map(function (r) {
    return { type: String(r[0]).trim(), fragments: String(r[1]).split(';').map(norm).filter(String) };
  });
  var rules = ss.getSheetByName('Правила банку').getRange('A4:C400').getValues().filter(function (r) { return r[0] && r[2]; }).map(function (r) {
    return { kw: norm(r[0]), dir: dirFromText(r[1]), category: String(r[2]).trim() };
  });
  return { fields: fields, statuses: statuses, hints: hints, rules: rules };
}

// ------------------------------------------------------------------ робота з аркушами
function Tabela_(kind) {
  var u = UKLAD[kind];
  var sh = SpreadsheetApp.getActive().getSheetByName(u.sheet);
  var n = u.last - 1;
  var data = sh.getRange(2, 1, n, u.keyCol).getValues();
  var keys = {}, next = 0;
  for (var i = 0; i < n; i++) {
    var k = data[i][u.keyCol - 1];
    if (k) keys[k] = i;
    if (data[i][0] !== '' && data[i][0] !== null) next = i + 1;
  }
  var dirty = {};
  return {
    data: data, u: u, keys: keys,
    get: function (i, f) { return data[i][u.cols[f] - 1]; },
    set: function (i, f, v) { data[i][u.cols[f] - 1] = (v === null || v === undefined) ? '' : v; dirty[u.cols[f]] = true; },
    rows: function () { return next; },
    upsert: function (key, rec, fieldsNew, fieldsUpdate) {
      var i = keys[key], isNew = i === undefined;
      if (isNew) {
        if (next >= n) throw new Error('Аркуш «' + u.sheet + '» заповнений. Зробіть копію файлу на новий рік або додайте рядки з формулами.');
        i = next++; keys[key] = i;
        data[i][u.keyCol - 1] = key; dirty[u.keyCol] = true;
      }
      (isNew ? fieldsNew : fieldsUpdate).forEach(function (f) {
        if (rec[f] === undefined || (!isNew && (rec[f] === null || rec[f] === ''))) return;
        data[i][u.cols[f] - 1] = rec[f] === null ? '' : rec[f]; dirty[u.cols[f]] = true;
      });
      return { index: i, isNew: isNew };
    },
    save: function () {
      var len = Math.max(next, 1);
      Object.keys(dirty).forEach(function (c) {
        c = +c;
        sh.getRange(2, c, len, 1).setValues(data.slice(0, len).map(function (r) { return [r[c - 1]]; }));
      });
    }
  };
}

function zapisz_(type, recs, cfg) {
  var added = 0, updated = 0, warnings = [];
  function count(r) { if (r.isNew) added++; else updated++; }

  if (type === 'BANK') {
    var t = Tabela_('BANK');
    recs.forEach(function (r) {
      r.autoCat = categorize(r.counterparty, r.description, r.amount, cfg.rules);
      count(t.upsert(r.key, r, ['date', 'amount', 'counterparty', 'description', 'autoCat', 'account'], []));
    });
    t.save();

  } else if (type === 'INVOICES') {
    var f = Tabela_('INVOICES');
    recs.forEach(function (r) {
      if (!r.category) r.category = categorize(r.seller, '', -1, cfg.rules);
      count(f.upsert(r.key, r, ['number', 'date', 'seller', 'nip', 'gross', 'due', 'category'], ['gross', 'due']));
    });
    f.save();

  } else if (type === 'PATIENTS') {
    var p = Tabela_('PATIENTS');
    recs.forEach(function (r) { count(p.upsert(r.key, r, ['patient', 'regDate', 'source'], ['regDate', 'source'])); });
    p.save();

  } else if (type === 'ZL' || type === 'BOOKSY') {
    var o = Tabela_('ONLINE');
    recs.forEach(function (r) { count(o.upsert(r.key, r, ['date', 'platform', 'patient', 'doctor', 'service', 'status'], ['status', 'doctor'])); });
    o.save();

  } else if (type === 'ADS') {
    var a = Tabela_('ADS');
    recs.forEach(function (r) { count(a.upsert(r.key, r, ['date', 'campaign', 'impressions', 'clicks', 'cost', 'conversions'], ['impressions', 'clicks', 'cost', 'conversions'])); });
    a.save();

  } else if (type === 'VISITS') {
    var res = zapiszWizyty_(recs);
    added = res.added; updated = res.updated; warnings = res.warnings;
  }
  return { added: added, updated: updated, warnings: warnings };
}

/** Візити: напрям з «Послуг», нові лікарі/послуги в довідники, джерело та «новий пацієнт». */
function zapiszWizyty_(recs) {
  var ss = SpreadsheetApp.getActive();
  var warnings = [];
  // послуги
  var shS = ss.getSheetByName('Послуги');
  var serv = shS.getRange(2, 1, 499, 3).getValues();
  var servMap = {}, servNext = 0;
  serv.forEach(function (r, i) { if (r[0]) { servMap[norm(r[0])] = { dir: r[1], dur: r[2] }; servNext = i + 1; } });
  var newServ = [];
  // лікарі
  var shD = ss.getSheetByName('Довідники');
  var docs = shD.getRange('A4:A20').getValues().map(function (r) { return r[0]; });
  var docSet = {}; docs.forEach(function (d) { if (d) docSet[norm(d)] = d; });
  var newDocs = [];
  // джерела
  var srcNames = shD.getRange('J4:J15').getValues().map(function (r) { return r[0]; }).filter(String);
  function mapSource(s) {
    var n = norm(s); if (!n) return '';
    for (var i = 0; i < srcNames.length; i++) if (norm(srcNames[i]) === n) return srcNames[i];
    for (i = 0; i < srcNames.length; i++) { var sn = norm(srcNames[i]).split(/[ /(]/)[0]; if (sn.length >= 4 && n.indexOf(sn) >= 0) return srcNames[i]; }
    if (/polec|znajom|rodzin/.test(n)) return srcNames.filter(function (x) { return /polecenie/i.test(x); })[0] || s;
    return s;
  }
  var ctx = kontekstPacjentow_();
  var t = Tabela_('VISITS');
  // перший візит кожного пацієнта (з існуючих + нових)
  var first = {};
  for (var i = 0; i < t.rows(); i++) {
    var p = norm(t.get(i, 'patient')), d = t.get(i, 'date');
    if (p && d instanceof Date && (!first[p] || d < first[p])) first[p] = d;
  }
  recs.forEach(function (r) { var p = norm(r.patient); if (!first[p] || r.date < first[p]) first[p] = r.date; });
  var usedFirstSeen = false, added = 0, updated = 0;

  recs.forEach(function (r) {
    var sk = norm(r.service);
    if (r.service && !servMap[sk]) { servMap[sk] = { dir: '', dur: '' }; newServ.push([r.service, r.direction || '', '']); }
    var sv = servMap[sk] || {};
    if (!r.direction) r.direction = sv.dir || '';
    if (r.duration === null && sv.dur) r.duration = sv.dur;
    if (r.doctor && !docSet[norm(r.doctor)]) { docSet[norm(r.doctor)] = r.doctor; newDocs.push(r.doctor); }
    else if (r.doctor) r.doctor = docSet[norm(r.doctor)];
    var pk = norm(r.patient);
    // новий пацієнт
    if (r.isNew === null) {
      var reg = ctx.reg[pk];
      if (reg) r.isNew = (reg.getFullYear() === r.date.getFullYear() && reg.getMonth() === r.date.getMonth()) ? 'Так' : 'Ні';
      else { r.isNew = first[pk] && first[pk].getTime() === r.date.getTime() ? 'Так' : 'Ні'; usedFirstSeen = true; }
    }
    // джерело
    r.source = mapSource(r.source) || ctx.src[pk] || ctx.online[pk + '|' + ymd(r.date)] || ctx.onlineAny[pk] || ctx.visitSrc[pk] || '';
    if (r.lab === null) delete r.lab;
    if (r.prep === null) delete r.prep;
    var res = t.upsert(r.key, r,
      ['date', 'patient', 'isNew', 'source', 'doctor', 'direction', 'service', 'amount', 'lab', 'prep', 'payment', 'status', 'duration'],
      ['amount', 'status', 'payment', 'duration', 'lab', 'prep', 'direction']);
    if (res.isNew) added++; else updated++;
    // джерело / «новий» не перезаписуємо, якщо вже вписано вручну
    if (!res.isNew && !t.get(res.index, 'source') && r.source) t.set(res.index, 'source', r.source);
  });
  t.save();
  if (newServ.length) {
    shS.getRange(servNext + 2, 1, newServ.length, 3).setValues(newServ);
    warnings.push('Нові послуги додано на аркуш «Послуги» — вкажіть для них напрям (червоні): ' + newServ.length);
  }
  if (newDocs.length) {
    var free = docs.map(function (d, i) { return d ? -1 : i; }).filter(function (i) { return i >= 0; });
    newDocs.slice(0, free.length).forEach(function (d, k) { shD.getRange(4 + free[k], 1).setValue(d); });
    warnings.push('Нові лікарі додано в «Довідники» — впишіть їхній %: ' + newDocs.join(', '));
  }
  if (usedFirstSeen) warnings.push('«Новий пацієнт» визначено за першим візитом у таблиці. Для точності завантажте список пацієнтів з Felgdent (з датою реєстрації).');
  return { added: added, updated: updated, warnings: warnings };
}

/** Довідкові дані про пацієнтів: дата реєстрації, джерело, онлайн-записи, джерело з попередніх візитів. */
function kontekstPacjentow_() {
  var ss = SpreadsheetApp.getActive();
  var reg = {}, src = {}, online = {}, onlineAny = {}, visitSrc = {};
  ss.getSheetByName('Пацієнти').getRange(2, 1, 3000, 3).getValues().forEach(function (r) {
    if (!r[0]) return;
    var k = norm(r[0]);
    if (r[1] instanceof Date) reg[k] = r[1];
    if (r[2]) src[k] = r[2];
  });
  ss.getSheetByName('Онлайн-записи').getRange(2, 1, 3000, 4).getValues().forEach(function (r) {
    if (!r[3] || !(r[0] instanceof Date)) return;
    var k = norm(r[3]);
    online[k + '|' + ymd(r[0])] = r[2];
    if (!onlineAny[k]) onlineAny[k] = r[2];
  });
  ss.getSheetByName('Візити').getRange(2, 3, 3000, 3).getValues().forEach(function (r) {
    if (r[0] && r[2] && !visitSrc[norm(r[0])]) visitSrc[norm(r[0])] = r[2];
  });
  return { reg: reg, src: src, online: online, onlineAny: onlineAny, visitSrc: visitSrc };
}

/** Після імпорту пацієнтів / онлайн-записів: заповнити порожні джерела та уточнити «новий пацієнт». */
function uzupelnijZrodla_() {
  var ctx = kontekstPacjentow_();
  var t = Tabela_('VISITS');
  for (var i = 0; i < t.rows(); i++) {
    var p = norm(t.get(i, 'patient')), d = t.get(i, 'date');
    if (!p || !(d instanceof Date)) continue;
    if (!t.get(i, 'source')) {
      var s = ctx.src[p] || ctx.online[p + '|' + ymd(d)] || ctx.onlineAny[p] || ctx.visitSrc[p];
      if (s) t.set(i, 'source', s);
    }
    var reg = ctx.reg[p];
    if (reg) t.set(i, 'isNew', reg.getFullYear() === d.getFullYear() && reg.getMonth() === d.getMonth() ? 'Так' : 'Ні');
  }
  t.save();
}

function przeliczKategorie_(cfg) {
  cfg = cfg || wczytajKonfig();
  var t = Tabela_('BANK');
  for (var i = 0; i < t.rows(); i++) {
    if (!(t.get(i, 'date') instanceof Date)) continue;
    t.set(i, 'autoCat', categorize(t.get(i, 'counterparty'), t.get(i, 'description'), Number(t.get(i, 'amount')), cfg.rules));
  }
  t.save();
}

function kategorieFaktur_(cfg) {
  var f = Tabela_('INVOICES');
  for (var i = 0; i < f.rows(); i++) {
    if (f.get(i, 'number') && !f.get(i, 'category')) {
      var c = categorize(f.get(i, 'seller'), '', -1, cfg.rules);
      if (c) f.set(i, 'category', c);
    }
  }
  f.save();
}

function zestawFaktury_() {
  var f = Tabela_('INVOICES'), b = Tabela_('BANK');
  var inv = [], bank = [];
  for (var i = 0; i < f.rows(); i++) {
    if (f.get(i, 'number') && !f.get(i, 'paidDate') && f.get(i, 'date') instanceof Date) {
      inv.push({ idx: i, number: String(f.get(i, 'number')), gross: Number(f.get(i, 'gross')), date: f.get(i, 'date'),
        seller: String(f.get(i, 'seller') || ''), nip: String(f.get(i, 'nip') || '').replace(/\D/g, '') });
    }
  }
  for (var j = 0; j < b.rows(); j++) {
    if (b.get(j, 'date') instanceof Date) bank.push({ idx: j, date: b.get(j, 'date'), amount: Number(b.get(j, 'amount')),
      counterparty: String(b.get(j, 'counterparty') || ''), description: String(b.get(j, 'description') || '') });
  }
  var matches = matchInvoices(inv, bank, []);
  matches.forEach(function (m) { f.set(m.invIdx, 'paidDate', m.date); });
  if (matches.length) f.save();
  return matches.length;
}

function dziennik_(file, type, added, updated, skipped, msg) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Журнал імпорту');
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 7).setValues([[new Date(), file, type, added, updated, skipped, msg || '']]);
}

// ------------------------------------------------------------------ команди меню
function przeliczKategorieMenu() { przeliczKategorie_(); alert_('Категорії банку перераховано.'); }
function zestawFakturyMenu() { alert_('Знайдено оплат фактур: ' + zestawFaktury_()); }
function uzupelnijZrodlaMenu() { uzupelnijZrodla_(); alert_('Джерела та «новий пацієнт» оновлено.'); }

/** Імпорт усіх файлів з папки Drive «Звіти клініки – Імпорт». Оброблені → підпапка «Оброблено». */
function importZFolderu() {
  var it = DriveApp.getFoldersByName(FOLDER_IMPORT);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_IMPORT);
  var di = folder.getFoldersByName(FOLDER_DONE);
  var done = di.hasNext() ? di.next() : folder.createFolder(FOLDER_DONE);
  var files = folder.getFiles(), out = [];
  var now = new Date(), prevMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS && file.getName().indexOf('tmp_import_') === 0) continue;
    try {
      var blob = file.getMimeType() === MimeType.GOOGLE_SHEETS
        ? file.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').setName(file.getName() + '.xlsx')
        : file.getBlob();
      out.push(importujBlob(blob, 'AUTO', prevMonth));
      file.moveTo(done);
    } catch (e) {
      out.push('❌ ' + e.message);
    }
  }
  return out;
}

function importZFolderuMenu() {
  var r = importZFolderu();
  alert_(r.length ? r.join('\n\n') : 'Папка «' + FOLDER_IMPORT + '» порожня. Скиньте туди файли й повторіть.');
}

/** Нічний імпорт + звіт 15-го о 21:00 (за Варшавою). */
function ustawHarmonogram() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['wyslijRaport', 'importZFolderu'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importZFolderu').timeBased().everyDays(1).atHour(5).inTimezone('Europe/Warsaw').create();
  ScriptApp.newTrigger('wyslijRaport').timeBased().onMonthDay(15).atHour(21).inTimezone('Europe/Warsaw').create();
  alert_('Готово: імпорт з папки «' + FOLDER_IMPORT + '» щоночі о 5:00, звіт на пошту 15-го числа між 21:00 і 22:00.');
}

/** Очищення демо-даних: лише колонки вводу, формули залишаються. */
function wyczyscDemo() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Видалити всі дані з аркушів Візити, Банк, Фактури, Витрати, Персонал, Онлайн-записи, Google Ads, Пацієнти, Журнал? Довідники й правила залишаться.',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  Object.keys(UKLAD).forEach(function (k) {
    var u = UKLAD[k];
    var sh = SpreadsheetApp.getActive().getSheetByName(u.sheet);
    var cols = Object.keys(u.cols).map(function (f) { return u.cols[f]; }).concat([u.keyCol]);
    cols.forEach(function (c) { sh.getRange(2, c, u.last - 1, 1).clearContent(); });
  });
  var j = SpreadsheetApp.getActive().getSheetByName('Журнал імпорту');
  if (j.getLastRow() > 1) j.getRange(2, 1, j.getLastRow() - 1, 7).clearContent();
  ui.alert('Демо-дані видалено. Тепер завантажте свої файли: меню 🦷 → «Завантажити файли…».');
}
