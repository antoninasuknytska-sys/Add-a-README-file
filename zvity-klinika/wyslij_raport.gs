/**
 * Автоматичне надсилання місячного звіту клініки у PDF.
 *
 * Як встановити (один раз):
 * 1. Відкрийте файл звітів у Google Таблицях.
 * 2. Розширення → Apps Script → видаліть порожній код, вставте цей файл.
 * 3. Впишіть свою пошту в EMAIL нижче (кілька адрес через кому).
 * 4. Збережіть, виберіть функцію ustawHarmonogram і натисніть «Запустити» → дозвольте доступ.
 * Після цього щомісяця 15-го числа між 21:00 і 22:00 (за Варшавою) прийде PDF
 * з Дашбордом за ПОПЕРЕДНІЙ місяць. Для перевірки запустіть wyslijRaport вручну.
 */

const EMAIL = 'WPISZ_SWOJ_EMAIL@example.com'; // ← ваша пошта
const ZAPISZ_KOPIE_NA_DYSKU = true;          // зберігати копію PDF у папці «Raporty kliniki» на Google Drive

const ARKUSZ_DASHBOARD = 'Дашборд';
const ARKUSZ_USTAWIENIA = 'Налаштування';
const KOMORKA_MIESIAC = 'B3';
const STREFA = 'Europe/Warsaw';

function wyslijRaport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teraz = new Date();
  const poprzedni = new Date(teraz.getFullYear(), teraz.getMonth() - 1, 1);

  // 1. Ставимо звітний місяць = попередній місяць і чекаємо на перерахунок
  ss.getSheetByName(ARKUSZ_USTAWIENIA).getRange(KOMORKA_MIESIAC).setValue(poprzedni);
  SpreadsheetApp.flush();
  Utilities.sleep(5000);

  // 2. Експортуємо лише аркуш «Дашборд» у PDF (A4, по ширині сторінки)
  const dash = ss.getSheetByName(ARKUSZ_DASHBOARD);
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf' +
    '&gid=' + dash.getSheetId() +
    '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false' +
    '&pagenum=CENTER&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
  const nazwaMiesiaca = Utilities.formatDate(poprzedni, STREFA, 'MM.yyyy');
  const pdf = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName('Zvit_kliniky_' + Utilities.formatDate(poprzedni, STREFA, 'yyyy-MM') + '.pdf');

  // 3. Ключові цифри в тексті листа (рядок 6 Дашборду)
  const v = dash.getRange('B6:I6').getValues()[0];
  const zl = n => Math.round(n).toLocaleString('pl-PL') + ' zł';
  const pr = n => (n * 100).toFixed(1) + '%';
  const tresc =
    'Звіт клініки за ' + nazwaMiesiaca + '\n\n' +
    'Виручка: ' + zl(v[0]) + '\n' +
    'Чистий прибуток: ' + zl(v[1]) + ' (рентабельність ' + pr(v[2]) + ')\n' +
    'Візитів: ' + v[3] + ', нових пацієнтів: ' + v[4] + '\n' +
    'Середній чек: ' + zl(v[5]) + '\n' +
    'Завантаження крісел: ' + pr(v[6]) + ', неявки та скасування: ' + pr(v[7]) + '\n\n' +
    'Повний звіт у PDF у вкладенні. Таблиця: ' + ss.getUrl();

  MailApp.sendEmail({
    to: EMAIL,
    subject: 'Звіт клініки за ' + nazwaMiesiaca,
    body: tresc,
    attachments: [pdf]
  });

  if (ZAPISZ_KOPIE_NA_DYSKU) {
    const it = DriveApp.getFoldersByName('Raporty kliniki');
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Raporty kliniki');
    folder.createFile(pdf);
  }
}

/** Запустіть один раз: створює щомісячний запуск 15-го числа о 21:00. */
function ustawHarmonogram() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'wyslijRaport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('wyslijRaport')
    .timeBased()
    .onMonthDay(15)
    .atHour(21)
    .inTimezone(STREFA)
    .create();
}
