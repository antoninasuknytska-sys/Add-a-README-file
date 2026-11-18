/**
 * МІСЯЧНИЙ ЗВІТ НА ПОШТУ (польською): PDF аркуша «Raport» + кольоровий лист з головними цифрами.
 * Запускається автоматично 15-го числа о 21:00 (див. ustawHarmonogram у Kod.gs).
 */

var EMAIL = 'WPISZ_SWOJ_EMAIL@example.com'; // ← ваша пошта (кілька адрес — через кому)
var ZAPISZ_KOPIE_NA_DYSKU = true;           // копія PDF у папці «Raporty kliniki» на Google Drive
var STREFA = 'Europe/Warsaw';

var JEZYKI = ['pl'];                          // звіт лише польською

var KOLORY = {
  'добре':    { bg: '#9FD8A6', tx: '#0B4A17' }, 'dobrze': { bg: '#9FD8A6', tx: '#0B4A17' },
  'середньо': { bg: '#FFC56E', tx: '#6B3A00' }, 'średnio': { bg: '#FFC56E', tx: '#6B3A00' },
  'погано':   { bg: '#F28B82', tx: '#6E0B0B' }, 'źle': { bg: '#F28B82', tx: '#6E0B0B' }
};
var JEZYK_CFG = {
  pl: { arkusz: 'Raport', plik: 'Raport_kliniki_', temat: 'Raport kliniki', tytul: 'Raport kliniki za ', zla: 'źle',
        kol: ['Wskaźnik', 'Ten miesiąc', 'Do poprz. miesiąca', 'Ocena'], legenda: ['DOBRZE', 'ŚREDNIO', 'ŹLE'],
        uwaga: 'Zwrócić uwagę', braki: 'Dane niepełne', pdf: 'Pełny raport — w załączniku PDF.' }
};

function wyslijRaport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 1. Спершу забрати нові файли з папки імпорту, щоб звіт був повним
  try { importZFolderu(); } catch (e) { Logger.log('Імпорт перед звітом: ' + e.message); }

  // 2. Звітний місяць = попередній
  var teraz = new Date();
  var poprzedni = new Date(teraz.getFullYear(), teraz.getMonth() - 1, 1);
  ss.getSheetByName('Налаштування').getRange('B3').setValue(poprzedni);
  SpreadsheetApp.flush();
  Utilities.sleep(8000);

  // 3. PDF аркуша «Raport»
  var miesiac = Utilities.formatDate(poprzedni, STREFA, 'MM.yyyy');
  var ym = Utilities.formatDate(poprzedni, STREFA, 'yyyy-MM');
  var pdfs = [], sekcje = [];
  JEZYKI.forEach(function (jez) {
    var cfg = JEZYK_CFG[jez];
    var dash = ss.getSheetByName(cfg.arkusz);
    if (!dash) return;
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + dash.getSheetId() +
      '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER' +
      '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
    pdfs.push(UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
      .getBlob().setName(cfg.plik + ym + '.pdf'));
    sekcje.push(sekcjaHtml_(dash, cfg, miesiac));
  });

  var html = '<div style="font-family:Arial,sans-serif;font-size:14px">' + sekcje.join('<hr style="margin:28px 0">') +
    '<p><a href="' + ss.getUrl() + '">Otwórz arkusz</a></p></div>';
  MailApp.sendEmail({
    to: EMAIL,
    subject: JEZYKI.map(function (j) { return JEZYK_CFG[j].temat; }).join(' / ') + ' ' + miesiac,
    body: 'Raport kliniki ' + miesiac + ' (PDF).',
    htmlBody: html,
    attachments: pdfs
  });

  if (ZAPISZ_KOPIE_NA_DYSKU) {
    var it = DriveApp.getFoldersByName('Raporty kliniki');
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder('Raporty kliniki');
    pdfs.forEach(function (p) { folder.createFile(p); });
  }
}

/** Кольорова таблиця головних цифр для листа (рядки 6, 7, 9, 10 аркуша звіту). */
function sekcjaHtml_(dash, cfg, miesiac) {
  var nazwy = dash.getRange('B6:I6').getDisplayValues()[0];
  var wart = dash.getRange('B7:I7').getDisplayValues()[0];
  var zmiana = dash.getRange('B9:I9').getDisplayValues()[0];
  var ocena = dash.getRange('B10:I10').getValues()[0];
  var problemy = dash.getRange('A14:B21').getValues().filter(function (r) { return Number(r[1]) > 0; });
  var podpowiedzi = dash.getRange('A1:L200').getValues().filter(function (r) { return r[11] === cfg.zla && String(r[0]).length > 40; }).map(function (r) { return r[0]; });
  var td = 'padding:8px 12px;border:1px solid #ddd';
  var rows = nazwy.map(function (n, i) {
    var k = KOLORY[ocena[i]] || { bg: '#FFFFFF', tx: '#000000' };
    var c = ';background:' + k.bg + ';color:' + k.tx + ';font-weight:bold';
    return '<tr><td style="' + td + '">' + n + '</td><td style="' + td + c + ';font-size:16px">' + wart[i] + '</td>' +
      '<td style="' + td + ';color:#555">' + zmiana[i] + '</td><td style="' + td + c + '">' + (ocena[i] || '') + '</td></tr>';
  }).join('');
  return '<h2 style="color:#1F4E5A">' + cfg.tytul + miesiac + '</h2>' +
    '<p><span style="background:#9FD8A6;padding:3px 8px">' + cfg.legenda[0] + '</span> <span style="background:#FFC56E;padding:3px 8px">' +
    cfg.legenda[1] + '</span> <span style="background:#F28B82;padding:3px 8px">' + cfg.legenda[2] + '</span></p>' +
    '<table style="border-collapse:collapse"><tr style="background:#D9E7EA">' + cfg.kol.map(function (h) { return '<th style="padding:8px">' + h + '</th>'; }).join('') + '</tr>' + rows + '</table>' +
    (podpowiedzi.length ? '<h3 style="color:#6E0B0B">' + cfg.uwaga + '</h3><ul>' + podpowiedzi.map(function (p) { return '<li>' + p + '</li>'; }).join('') + '</ul>' : '') +
    (problemy.length ? '<h3 style="color:#6E0B0B">' + cfg.braki + '</h3><ul>' + problemy.map(function (p) { return '<li>' + p[0] + ': ' + p[1] + '</li>'; }).join('') + '</ul>' : '') +
    '<p>' + cfg.pdf + '</p>';
}
