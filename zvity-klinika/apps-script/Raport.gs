/**
 * МІСЯЧНИЙ ЗВІТ НА ПОШТУ: PDF Дашборду + кольоровий лист з головними цифрами.
 * Запускається автоматично 15-го числа о 21:00 (див. ustawHarmonogram у Kod.gs).
 */

var EMAIL = 'WPISZ_SWOJ_EMAIL@example.com'; // ← ваша пошта (кілька адрес — через кому)
var ZAPISZ_KOPIE_NA_DYSKU = true;           // копія PDF у папці «Raporty kliniki» на Google Drive
var STREFA = 'Europe/Warsaw';

var KOLORY = {
  'добре':    { bg: '#9FD8A6', tx: '#0B4A17' },
  'середньо': { bg: '#FFC56E', tx: '#6B3A00' },
  'погано':   { bg: '#F28B82', tx: '#6E0B0B' }
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

  // 3. PDF лише аркуша «Дашборд»
  var dash = ss.getSheetByName('Дашборд');
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + dash.getSheetId() +
    '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=CENTER' +
    '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
  var miesiac = Utilities.formatDate(poprzedni, STREFA, 'MM.yyyy');
  var pdf = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName('Zvit_kliniky_' + Utilities.formatDate(poprzedni, STREFA, 'yyyy-MM') + '.pdf');

  // 4. Кольоровий лист: показник / значення / зміна / оцінка (рядки 6, 7, 9, 10 Дашборду)
  var nazwy = dash.getRange('B6:I6').getDisplayValues()[0];
  var wart = dash.getRange('B7:I7').getDisplayValues()[0];
  var zmiana = dash.getRange('B9:I9').getDisplayValues()[0];
  var ocena = dash.getRange('B10:I10').getValues()[0];
  var problemy = dash.getRange('A14:B21').getValues().filter(function (r) { return Number(r[1]) > 0; });
  var podpowiedzi = dash.getRange('A1:L200').getValues().filter(function (r) { return r[11] === 'погано' && String(r[0]).length > 40; }).map(function (r) { return r[0]; });

  var rows = nazwy.map(function (n, i) {
    var k = KOLORY[ocena[i]] || { bg: '#FFFFFF', tx: '#000000' };
    return '<tr><td style="padding:8px 12px;border:1px solid #ddd">' + n + '</td>' +
      '<td style="padding:8px 12px;border:1px solid #ddd;font-size:16px;font-weight:bold;background:' + k.bg + ';color:' + k.tx + '">' + wart[i] + '</td>' +
      '<td style="padding:8px 12px;border:1px solid #ddd;color:#555">' + zmiana[i] + '</td>' +
      '<td style="padding:8px 12px;border:1px solid #ddd;font-weight:bold;background:' + k.bg + ';color:' + k.tx + '">' + (ocena[i] || '') + '</td></tr>';
  }).join('');
  var html = '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<h2 style="color:#1F4E5A">Звіт клініки за ' + miesiac + '</h2>' +
    '<p><span style="background:#9FD8A6;padding:3px 8px">ДОБРЕ</span> ' +
    '<span style="background:#FFC56E;padding:3px 8px">СЕРЕДНЬО</span> ' +
    '<span style="background:#F28B82;padding:3px 8px">ПОГАНО</span></p>' +
    '<table style="border-collapse:collapse"><tr style="background:#D9E7EA"><th style="padding:8px">Показник</th><th>Цей місяць</th><th>До мин. місяця</th><th>Оцінка</th></tr>' +
    rows + '</table>' +
    (podpowiedzi.length ? '<h3 style="color:#6E0B0B">Звернути увагу</h3><ul>' + podpowiedzi.map(function (p) { return '<li>' + p + '</li>'; }).join('') + '</ul>' : '') +
    (problemy.length ? '<h3 style="color:#6E0B0B">Дані неповні</h3><ul>' + problemy.map(function (p) { return '<li>' + p[0] + ': ' + p[1] + '</li>'; }).join('') + '</ul>' : '') +
    '<p>Повний звіт — у PDF у вкладенні. <a href="' + ss.getUrl() + '">Відкрити таблицю</a></p></div>';

  MailApp.sendEmail({
    to: EMAIL,
    subject: 'Звіт клініки за ' + miesiac,
    body: 'Звіт клініки за ' + miesiac + '. Відкрийте лист у форматі HTML або PDF у вкладенні.',
    htmlBody: html,
    attachments: [pdf]
  });

  if (ZAPISZ_KOPIE_NA_DYSKU) {
    var it = DriveApp.getFoldersByName('Raporty kliniki');
    var folder = it.hasNext() ? it.next() : DriveApp.createFolder('Raporty kliniki');
    folder.createFile(pdf);
  }
}
