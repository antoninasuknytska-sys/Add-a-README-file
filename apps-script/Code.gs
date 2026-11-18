/**
 * „Tablica spraw recepcji” — przekazywanie zmiany bez gubienia pacjentów.
 * Google Apps Script podpięty do Arkusza Google.
 *
 * Arkusze tworzone przez setup():
 *   Sprawy      — wszystkie sprawy (jedna sprawa = jeden wiersz)
 *   Zmiany      — dziennik przekazania zmian
 *   Lekarze     — lekarze i ich numery WhatsApp
 *   Ustawienia  — recepcjonistki, e-mail właściciela, filtr maili z tomografią itd.
 */

const SHEET_TASKS = 'Sprawy';
const SHEET_SHIFTS = 'Zmiany';
const SHEET_DOCTORS = 'Lekarze';
const SHEET_SETTINGS = 'Ustawienia';
const SHEET_INBOX = 'Skrzynka';
const SHEET_DEBTS = 'Potrącenia';
const DEBT_COLS = ['id', 'date', 'person', 'category', 'patient', 'description', 'amount', 'status', 'addedBy', 'addedAt', 'settledBy', 'settledAt', 'taskId'];
const DEBT_HEADERS = ['ID', 'Data zdarzenia', 'Kogo dotyczy', 'Kategoria', 'Pacjent', 'Opis', 'Kwota (zł)', 'Status', 'Kto dodał', 'Dodano', 'Kto rozliczył', 'Rozliczono', 'ID sprawy'];
const KT_LABEL = 'CRM-KT-done';
const BOT = '🤖 Automat';

const TASK_COLS = [
  'id', 'createdAt', 'createdBy', 'type', 'patient', 'phone', 'channel', 'doctor',
  'purpose', 'stage', 'dueAt', 'status', 'updatedAt', 'updatedBy', 'closedAt',
  'source', 'history'
];
const TASK_HEADERS = [
  'ID', 'Utworzono', 'Kto utworzył', 'Typ', 'Pacjent', 'Telefon', 'Kanał', 'Lekarz',
  'O co chodzi', 'Etap', 'Termin', 'Status', 'Zaktualizowano', 'Kto zaktualizował', 'Zamknięto',
  'Źródło', 'Historia (JSON)'
];
const SHIFT_COLS = ['id', 'at', 'from', 'note', 'openCount', 'overdueCount', 'acceptedBy', 'acceptedAt'];
const INBOX_COLS = ['id', 'msgId', 'receivedAt', 'source', 'kind', 'name', 'phone', 'email', 'subject', 'text', 'link', 'mailbox', 'urgent', 'status', 'handledBy', 'handledAt', 'note', 'taskId', 'when', 'service', 'doctor'];
const INBOX_HEADERS = ['ID', 'ID maila', 'Otrzymano', 'Źródło', 'Rodzaj', 'Imię i nazwisko', 'Telefon', 'E-mail', 'Temat', 'Treść', 'Link', 'Skrzynka', 'Pilne', 'Status', 'Kto obsłużył', 'Kiedy obsłużono', 'Notatka', 'ID sprawy', 'Termin wizyty', 'Usługa', 'Do kogo'];
const SHIFT_HEADERS = ['ID', 'Przekazano', 'Kto przekazał', 'Notatka', 'Otwartych spraw', 'Zaległych', 'Kto przejął', 'Przejęto'];

/**
 * Typy spraw i ich etapy.
 *   due   — po ilu godzinach od przejścia na etap sprawa staje się zaległa;
 *   next  — „naturalny” następny etap (domyślnie kolejny na liście);
 *   retry — etap „nie odebrał”, na który przenosi przycisk 📵;
 *   nag   — etap czekania na odpowiedź lekarza (przycisk 🔔 „Przypomnij lekarzowi”);
 *   ask   — pytanie, na które trzeba odpowiedzieć przy przejściu na etap;
 *   done  — etap końcowy, sprawa się zamyka.
 */
const TYPES = {
  lead: {
    label: 'Nowe zapytanie (strona, mail, SMS, Instagram)', icon: '🙋',
    stages: [
      { n: 'Oddzwonić / odpisać', due: 1 },
      { n: 'Nie odebrał — ponowić', due: 3, retry: true },
      { n: 'Pacjent się zastanawia', due: 72 },
      { n: 'Umówiony', done: true }
    ]
  },
  fb: {
    label: 'Ogłoszenia FB', icon: '📣',
    stages: [
      { n: 'Oddzwonić / odpisać', due: 1 },
      { n: 'Nie odebrał — ponowić', due: 3, retry: true },
      { n: 'Pacjent się zastanawia', due: 72 },
      { n: 'Umówiony', done: true }
    ]
  },
  kt: {
    label: 'Tomografia', icon: '🦷',
    stages: [
      { n: 'Skierowanie na tomografię wysłane', due: 72 },
      { n: 'Tomografia gotowa — pokazać lekarzowi', due: 2 },
      { n: 'Czekamy na odpowiedź lekarza', due: 4, nag: true },
      { n: 'Lekarz odpowiedział — umówić pacjenta', due: 4, ask: 'Co powiedział lekarz: kiedy i na ile czasu umówić?' },
      { n: 'Pacjent umówiony', done: true }
    ]
  },
  rtg: {
    label: 'RTG', icon: '🩻',
    stages: [
      { n: 'Czekamy na RTG od pacjenta', due: 72 },
      { n: 'RTG otrzymane — pokazać lekarzowi', due: 2 },
      { n: 'Czekamy na odpowiedź lekarza', due: 4, nag: true },
      { n: 'Lekarz odpowiedział — umówić pacjenta', due: 4, ask: 'Co powiedział lekarz: kiedy i na ile czasu umówić?' },
      { n: 'Pacjent umówiony', done: true }
    ]
  },
  plan: {
    label: 'Plan leczenia', icon: '📋',
    stages: [
      { n: 'Lekarz przygotowuje plan / wycenę', due: 48, nag: true },
      { n: 'Plan gotowy — przekazać pacjentowi', due: 24, next: 3 },
      { n: 'Nie odebrał — ponowić', due: 24, retry: true, next: 3 },
      { n: 'Pacjent się zastanawia', due: 72, next: 4 },
      { n: 'Pacjent zaakceptował — umówiony', done: true }
    ]
  },
  call_after: {
    label: 'Telefon po zabiegu', icon: '📞',
    stages: [
      { n: 'Zadzwonić', due: 24, next: 3 },
      { n: 'Nie odebrał — oddzwonić', due: 3, retry: true, next: 3 },
      { n: 'Są dolegliwości — przekazano lekarzowi', due: 4, nag: true, next: 3, ask: 'Jakie dolegliwości? Co przekazano lekarzowi?' },
      { n: 'Rozmowa — wszystko dobrze', done: true }
    ]
  },
  lab: {
    label: 'Laboratorium (korony, licówki)', icon: '🧪',
    stages: [
      { n: 'Wysłano do laboratorium', due: 120 },
      { n: 'Praca gotowa — umówić na osadzenie', due: 24, next: 3 },
      { n: 'Nie odebrał — oddzwonić', due: 3, retry: true, next: 3 },
      { n: 'Umówiony na osadzenie', done: true }
    ]
  },
  confirm: {
    label: 'Potwierdzenie wizyty', icon: '✅',
    stages: [
      { n: 'Potwierdzić wizytę', due: 24, next: 2 },
      { n: 'Nie odebrał — ponowić', due: 3, retry: true, next: 2 },
      { n: 'Potwierdzone', done: true }
    ]
  },
  revisit: {
    label: 'Wizyta kontrolna', icon: '🔁',
    stages: [
      { n: 'Zaprosić na wizytę', due: 24, next: 2 },
      { n: 'Nie odebrał — ponowić', due: 24, retry: true, next: 2 },
      { n: 'Umówiony', done: true }
    ]
  },
  other: {
    label: 'Bieżące zadania', icon: '📌',
    stages: [
      { n: 'Do zrobienia', due: 8 },
      { n: 'W trakcie', due: 8 },
      { n: 'Zrobione', done: true }
    ]
  }
};

/** Sekcje raportu miesięcznego (listy pacjentów). */
const REPORT_GROUPS = [
  { title: '🦷 Tomografia i 🩻 RTG', types: ['kt', 'rtg'] },
  { title: '📋 Plan leczenia', types: ['plan'] },
  { title: '📞 Telefony po zabiegu', types: ['call_after'] },
  { title: '🧪 Laboratorium', types: ['lab'] },
  { title: '✅ Potwierdzenia wizyt', types: ['confirm'] },
  { title: '📌 Bieżące zadania', types: ['other'] },
  { title: '📣 Ogłoszenia FB', types: ['fb'] }
];

/** Lekarze i higienistki: [nazwa na tablicy, WhatsApp, inne nazwy w Booksy / ZnanyLekarz]. */
const DEFAULT_DOCTORS = [
  ['Krzysztof Dragun', '', 'Krzysztof, dr Dragun'],
  ['Andrzej Kulpa', '', 'Andrzej, dr Kulpa'],
  ['Stanisław Tsydzik', '', 'Stanisław, dr Tsydzik'],
  ['Ewa Winiarska', '', 'Ewa, dr Winiarska'],
  ['Piotr Wereski', '', 'Piotr, dr Wereski'],
  ['Vasyl Kuz', '', 'Vasyl, dr Kuz'],
  ['Kornelia Borska', '', 'Kornelia, dr Borska'],
  ['Artur', '', 'Artur'],
  ['Oliwia (higienistka)', '', 'Oliwia'],
  ['Antonina (higienistka)', '', 'Antonina']
];

const DEFAULT_SETTINGS = [
  ['clinicName', 'Gabinet stomatologiczny', 'Nazwa wyświetlana u góry'],
  ['admins', 'Nadia N, Nadia R, Tosia, Właściciel', 'Kto pracuje z tablicą: rejestratorki + właściciel (po przecinku)'],
  ['ownerEmail', 'dentonia.pl@gmail.com', 'E-mail właściciela (recepcja nie ma do niego dostępu): poranny raport, odpowiedzi na pytania, miesięczne podsumowanie potrąceń'],
  ['reportTime', '20:30', 'Godzina raportu dziennego (codziennie) i miesięcznego (ostatniego dnia miesiąca), np. 20:30. Po zmianie uruchom ponownie „1. Konfiguracja początkowa”'],
  ['staleHours', '48', 'Po ilu godzinach bez ruchu (po terminie) sprawa jest „zawieszona”'],
  ['mailboxes', 'polishdent.pl@gmail.com=PL', 'Skrzynka recepcji (adres=znaczek). Maile wysłane z niej nie trafiają do „Skrzynki”. Przy kilku skrzynkach (po przecinku) na wiadomościach pojawia się znaczek'],
  ['ownerName', 'Właściciel', 'Imię właściciela z listy admins (nie pojawia się na liście osób do potrąceń)'],
  ['staff', '', 'Pozostali pracownicy do zakładki „Potrącenia” (asystentki, higienistki itp.), po przecinku. Lekarze i recepcja są dodawani automatycznie'],
  ['notifyAnswers', 'nie', 'tak = e-mail do właściciela, gdy recepcja odpowie na jego pytanie (domyślnie nie — odpowiedzi są na tablicy)'],
  ['debtsSheetUrl', '', 'Link do arkusza „Potrącenia” na koncie właściciela (dentonia.pl), udostępnionego dla polishdent.pl@gmail.com jako Edytor. Puste = potrącenia w tym arkuszu'],
  ['ownerPin', '', 'Opcjonalny PIN do zakładki „Potrącenia”. Puste = recepcja też widzi listę (domyślnie)'],
  ['inboxIgnore', 'diagdent, portal lekarza, mailer-daemon, accounts.google.com', 'Pomijaj maile, których nadawca lub temat zawiera te słowa (po przecinku)'],
  ['inboxAlertMinutes', '30', 'Po ilu minutach nieobsłużona wiadomość w skrzynce robi się czerwona'],
  ['ktEmailQuery', '"Portal Lekarza" "Nowe badanie"', 'Zapytanie Gmail dla maili Diagdent o gotowej tomografii. Puste = automatyzacja wyłączona']
];

/* ===================== Instalacja ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🦷 Tablica spraw')
    .addItem('1. Konfiguracja początkowa', 'setup')
    .addItem('Sprawdź teraz pocztę (skrzynka + tomografie)', 'checkMail')
    .addItem('Wyślij teraz raport dzienny', 'dailyReport')
    .addItem('Wyślij teraz raport miesięczny (poprzedni miesiąc)', 'monthlyReport')
    .addToUi();
}

/** Uruchom raz: tworzy arkusze i automatyczne wyzwalacze. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  // Wyzwalacze czytają pocztę konta, które uruchomiło setup — musi to być skrzynka recepcji
  const box = (mailboxes_(getSettings_())[0] || {}).email;
  const runner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (box && runner && runner !== box) {
    throw new Error('Uruchom „setup” zalogowana na koncie ' + box + ' (teraz: ' + runner + '). Wyzwalacze czytają pocztę konta, które je utworzyło.');
  }
  ensureSheet_(ss, SHEET_TASKS, TASK_HEADERS);
  ensureSheet_(ss, SHEET_SHIFTS, SHIFT_HEADERS);
  ensureSheet_(ss, SHEET_INBOX, INBOX_HEADERS);
  const doctors = ensureSheet_(ss, SHEET_DOCTORS, ['Lekarz (jak ma być widoczny na tablicy)', 'WhatsApp (48XXXXXXXXX)', 'Inne nazwy: jak jest zapisany w Booksy i ZnanyLekarz (po przecinku)']);
  if (doctors.getLastRow() < 2) {
    doctors.getRange(2, 1, DEFAULT_DOCTORS.length, 3).setValues(DEFAULT_DOCTORS);
  }
  doctors.setColumnWidth(1, 260).setColumnWidth(3, 420);
  const settings = ensureSheet_(ss, SHEET_SETTINGS, ['Klucz', 'Wartość', 'Opis']);
  const existing = settings.getLastRow() > 1
    ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  DEFAULT_SETTINGS.forEach(row => { if (existing.indexOf(row[0]) < 0) settings.appendRow(row); });
  settings.setColumnWidth(2, 320).setColumnWidth(3, 520);

  ScriptApp.getProjectTriggers().forEach(t => {
    if (['checkKtEmails', 'checkMail', 'dailyReport'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkMail').timeBased().everyMinutes(5).create();
  const tm = String(getSettings_().reportTime || '20:30').match(/(\d{1,2})(?::(\d{2}))?/) || [0, 20, 30];
  ScriptApp.newTrigger('dailyReport').timeBased().everyDays(1)
    .atHour(Math.min(23, Number(tm[1]))).nearMinute(Number(tm[2] || 0)).create(); // strefa czasowa projektu: Europe/Warsaw
  if (getSettings_().debtsSheetUrl) debtSheet_(); // sprawdź dostęp do arkusza właściciela

  try {
    SpreadsheetApp.getUi().alert('Gotowe! Uzupełnij arkusze „Ustawienia” i „Lekarze”, a potem opublikuj aplikację internetową (patrz instrukcja).');
  } catch (e) { /* uruchomione nie z arkusza */ }
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ===================== Aplikacja internetowa ===================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Tablica spraw recepcji')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Wszystko, czego potrzebuje interfejs: ustawienia, otwarte sprawy, zamknięte z 30 dni, dziennik zmian. */
function getData() {
  const s = getSettings_();
  const monthAgo = Date.now() - 30 * 864e5;
  const tasks = readTasks_().filter(t => t.status === 'open' || (t.closedAt || 0) > monthAgo);
  const twoWeeks = Date.now() - 14 * 864e5;
  const inbox = readInbox_().filter(i => i.status === 'new' || (i.receivedAt || 0) > twoWeeks);
  return {
    now: Date.now(),
    config: {
      types: TYPES,
      clinicName: s.clinicName || 'Gabinet stomatologiczny',
      admins: String(s.admins || '').split(',').map(x => x.trim()).filter(Boolean),
      staleHours: Number(s.staleHours) || 48,
      doctors: readDoctors_(),
      ktAutomation: !!s.ktEmailQuery,
      multiBox: mailboxes_(s).length > 1,
      ownerName: String(s.ownerName || 'Właściciel').trim(),
      staff: String(s.staff || '').split(',').map(x => x.trim()).filter(Boolean),
      hasPin: !!String(s.ownerPin || '').trim(),
      viewer: activeEmail_(),
      viewerIsOwner: !!activeEmail_() && activeEmail_() === String(s.ownerEmail || '').toLowerCase().trim(),
      debtsUrl: String(s.debtsSheetUrl || '').trim(),
      inboxAlertMinutes: Number(s.inboxAlertMinutes) || 30
    },
    tasks: tasks,
    inbox: inbox,
    shifts: readShifts_().slice(-15)
  };
}

/** Utwórz lub zaktualizuj sprawę. Historię scalamy, żeby dwie osoby nie nadpisały sobie wpisów. */
function saveTask(task) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { upsertTask_(task); } finally { lock.releaseLock(); }
  return getData();
}

function upsertTask_(task) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  const rowIdx = findRow_(sh, task.id);
  if (rowIdx > 0) {
    const current = rowToTask_(sh.getRange(rowIdx, 1, 1, TASK_COLS.length).getValues()[0]);
    task.history = mergeHistory_(current.history, task.history);
    try { if (/^(tak|yes|1)$/i.test(String(getSettings_().notifyAnswers || '').trim())) notifyAnswers_(current.history, task); } catch (e) { /* e-mail nie może zablokować zapisu */ }
    sh.getRange(rowIdx, 1, 1, TASK_COLS.length).setValues([taskToRow_(task)]);
  } else {
    sh.appendRow(taskToRow_(task));
  }
}

/** Zapisz element skrzynki; opcjonalnie razem ze sprawą (utworzoną z niego lub uzupełnioną). payload = {item, task?} */
function saveInbox(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (payload.task) upsertTask_(payload.task);
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INBOX);
    const row = inboxToRow_(payload.item);
    const rowIdx = findRow_(sh, payload.item.id);
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/** Dodaj nadawcę do listy ignorowanych (inboxIgnore) i zamknij element. payload = {email, item} */
function ignoreSender(payload) {
  const email = String(payload.email || '').toLowerCase().trim();
  if (email) {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === 'inboxIgnore') {
        const cur = String(vals[i][1] || '');
        if (cur.toLowerCase().indexOf(email) < 0) sh.getRange(i + 2, 2).setValue(cur ? cur + ', ' + email : email);
      }
    }
  }
  return saveInbox({ item: payload.item });
}

/** Przekazanie zmiany (wpis bez acceptedBy) lub przejęcie (wpis z id i acceptedBy). */
function saveShift(entry) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHIFTS);
    const rowIdx = findRow_(sh, entry.id);
    const row = SHIFT_COLS.map(c => {
      const v = entry[c];
      if ((c === 'at' || c === 'acceptedAt') && v) return new Date(v);
      return v === undefined || v === null ? '' : v;
    });
    if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return getData();
}

/* ===================== Potrącenia (reklamacje, koszty do pokrycia) ===================== */

/** E-mail osoby, która otworzyła tablicę (gdy aplikacja działa jako „użytkownik uzyskujący dostęp”). */
function activeEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').toLowerCase().trim(); } catch (e) { return ''; }
}

function checkPin_(pin) {
  const p = String(getSettings_().ownerPin || '').trim();
  if (p && String(pin || '').trim() !== p) throw new Error('Nieprawidłowy PIN');
}

/** Lista potrąceń — tylko z PIN-em właściciela (jeśli ustawiony). */
function getDebts(pin) {
  checkPin_(pin);
  return readDebts_();
}

/** Nowy wpis może dodać każdy; zmiana istniejącego (rozliczenie, edycja) wymaga PIN-u. payload = {pin, debt} */
function saveDebt(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = debtSheet_();
    const rowIdx = findRow_(sh, payload.debt.id);
    const row = DEBT_COLS.map(c => {
      const v = payload.debt[c];
      if (['date', 'addedAt', 'settledAt'].indexOf(c) >= 0) return v ? new Date(v) : '';
      if (c === 'amount') return Math.round(Number(String(v).replace(',', '.')) * 100) / 100 || 0;
      return v === undefined || v === null ? '' : v;
    });
    if (rowIdx > 0) {
      checkPin_(payload.pin);
      const oldStatus = sh.getRange(rowIdx, DEBT_COLS.indexOf('status') + 1).getValue();
      const viewer = activeEmail_(), owner = String(getSettings_().ownerEmail || '').toLowerCase().trim();
      if (viewer && owner && viewer !== owner && String(payload.debt.status) !== String(oldStatus)) {
        throw new Error('Tylko właściciel może oznaczać potrącenia jako rozliczone i je usuwać.');
      }
      sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  } finally {
    lock.releaseLock();
  }
  try { checkPin_(payload.pin); return readDebts_(); } catch (e) { return null; }
}

/**
 * Arkusz potrąceń: na koncie właściciela (debtsSheetUrl), a jeśli link pusty — w tym arkuszu.
 * Plik właściciela musi być udostępniony dla polishdent.pl@gmail.com jako Edytor.
 */
function debtSheet_() {
  const url = String(getSettings_().debtsSheetUrl || '').trim();
  let ss;
  if (url) {
    try { ss = SpreadsheetApp.openByUrl(url); }
    catch (e) { throw new Error('Brak dostępu do arkusza potrąceń właściciela. Sprawdź, czy jest udostępniony dla konta recepcji jako Edytor.'); }
  } else {
    ss = SpreadsheetApp.getActive();
  }
  return ensureSheet_(ss, SHEET_DEBTS, DEBT_HEADERS);
}

function readDebts_() {
  const sh = debtSheet_();
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, DEBT_COLS.length).getValues()
    .filter(r => r[0])
    .map(r => {
      const o = {};
      DEBT_COLS.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].getTime() : r[i]; });
      o.amount = Number(o.amount) || 0;
      return o;
    });
}

/** Podsumowanie potrąceń za poprzedni miesiąc (do raportu z 1. dnia miesiąca). */
/** Okres raportu miesięcznego: ostatniego dnia miesiąca — bieżący miesiąc, w inne dni — poprzedni. */
function reportPeriod_() {
  const d = new Date();
  const last = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getDate() === 1;
  const m = last ? d.getMonth() : d.getMonth() - 1;
  return { from: new Date(d.getFullYear(), m, 1).getTime(), to: new Date(d.getFullYear(), m + 1, 1).getTime() };
}

function isLastDayOfMonth_() {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1).getDate() === 1;
}

function debtsMonthHtml_() {
  const per0 = reportPeriod_();
  const from = per0.from, to = per0.to;
  const list = readDebts_().filter(x => x.status !== 'cancelled' && ((x.date >= from && x.date < to) || (x.status === 'open' && x.date < from)));
  if (!list.length) return '<h3>💰 Potrącenia w tym okresie — brak</h3>';
  const per = {};
  list.forEach(x => { const p = per[x.person] = per[x.person] || { sum: 0, open: 0, n: 0 }; p.sum += x.amount; p.n++; if (x.status === 'open') p.open += x.amount; });
  const zl = n => n.toFixed(2).replace('.', ',') + ' zł';
  const fd = ts => Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  return '<h3>💰 Potrącenia: ' + Utilities.formatDate(new Date(from), Session.getScriptTimeZone(), 'MM.yyyy') + ' (+ nierozliczone wcześniejsze)</h3>' +
    '<table border="1" cellpadding="6" style="border-collapse:collapse;font-size:13px"><tr><th>Kto</th><th>Pozycji</th><th>Razem</th><th>Do potrącenia</th></tr>' +
    Object.keys(per).sort((a, b) => per[b].open - per[a].open).map(k => '<tr><td>' + esc_(k) + '</td><td>' + per[k].n + '</td><td>' + zl(per[k].sum) + '</td><td><b>' + zl(per[k].open) + '</b></td></tr>').join('') + '</table>' +
    '<p><b>Szczegóły:</b></p><table border="1" cellpadding="6" style="border-collapse:collapse;font-size:13px"><tr><th>Data</th><th>Kto</th><th>Kategoria</th><th>Pacjent</th><th>Opis</th><th>Kwota</th><th>Status</th></tr>' +
    list.sort((a, b) => (a.person > b.person ? 1 : a.person < b.person ? -1 : a.date - b.date)).map(x => '<tr><td>' + fd(x.date) + '</td><td>' + esc_(x.person) + '</td><td>' + esc_(x.category) + '</td><td>' + esc_(x.patient) +
      '</td><td>' + esc_(x.description) + '</td><td>' + zl(x.amount) + '</td><td>' + (x.status === 'settled' ? '✅ potrącono' : '<b>do potrącenia</b>') + '</td></tr>').join('') + '</table>';
}

/* ===================== Automatyzacja ===================== */

/** Wyzwalacz co 5 min: najpierw tomografie (Diagdent), potem skrzynka. */
function checkMail() {
  checkKtEmails();
  checkInbox();
}

/**
 * Skrzynka: zbiera nowe wiadomości od pacjentów i systemów rezerwacji do zakładki „📥 Skrzynka”.
 * Źródła: Booksy, ZnanyLekarz, formularz i wiadomości ze strony, maile pacjentów,
 * SMS-y i nieodebrane połączenia z telefonu (MacroDroid wysyła je mailem z tematem „CRM SMS …” / „CRM NIEODEBRANE …”).
 */
function checkInbox() {
  const s = getSettings_();
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const last = Number(props.getProperty('inboxLastCheck')) || (now - 24 * 36e5);
  const since = last - 15 * 6e4; // zapas na opóźnienia Gmaila
  const after = ' after:' + Math.floor(since / 1000);
  const queries = [
    'in:inbox category:primary' + after,
    'in:inbox' + after + ' (booksy OR znanylekarz OR docplanner OR "Twoje imię" OR "New message from" OR "Nowa wiadomość od" OR subject:CRM)'
  ];
  const boxes = mailboxes_(s);
  const doctors = readDoctors_();
  const ignore = list_(s.inboxIgnore);
  const seen = {};
  readInbox_().forEach(i => { seen[i.msgId] = 1; });
  const fresh = [];
  queries.forEach(q => GmailApp.search(q, 0, 50).forEach(th => th.getMessages().forEach(msg => {
    const id = msg.getId();
    if (seen[id]) return;
    seen[id] = 1;
    if (msg.getDate().getTime() < since) return;
    const item = classifyMail_(msg, boxes, ignore, doctors);
    if (!item) return;
    item.id = newId_();
    item.msgId = id;
    item.receivedAt = msg.getDate().getTime();
    item.link = mailLink_(th.getId());
    item.status = 'new';
    fresh.push(item);
  })));

  if (fresh.length) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INBOX);
      fresh.sort((a, b) => a.receivedAt - b.receivedAt).forEach(i => sh.appendRow(inboxToRow_(i)));
    } finally { lock.releaseLock(); }
  }
  props.setProperty('inboxLastCheck', String(now));
}

/** Rozpoznaj, skąd jest wiadomość, i wyciągnij pacjenta, telefon i najważniejsze informacje. null = pomiń. */
function classifyMail_(msg, boxes, ignore, doctors) {
  const it = classifyRaw_(msg, boxes, ignore);
  if (it && doctors && doctors.length && it.source !== 'sms' && it.source !== 'call') {
    it.doctor = resolveDoctor_(it.doctor, (msg.getSubject() || '') + ' ' + (msg.getPlainBody() || htmlToText_(msg.getBody())), doctors);
  }
  return it;
}

/** Normalizacja do porównań: małe litery, bez polskich znaków i interpunkcji. */
function normText_(s) {
  return ' ' + String(s || '').toLowerCase().replace(/ł/g, 'l').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
}

/**
 * Do którego lekarza jest wizyta. Szukamy nazw z arkusza „Lekarze” (kolumna A i „Inne nazwy”):
 * najpierw w polu znalezionym w mailu (np. „pracownik: Oliwia”), potem w całej treści maila.
 */
function resolveDoctor_(found, text, doctors) {
  const names = d => [d.name].concat(d.aliases || [])
    .map(a => normText_(a).replace(/ (dr|lek|dent|n|med|stom|hig|mgr) /g, ' ').replace(/\s+/g, ' '))
    .filter(a => a.trim().length >= 4);
  if (found) {
    const f = normText_(found);
    const d = doctors.filter(x => names(x).some(a => f === a || f.indexOf(a) >= 0 || a.indexOf(f) >= 0))[0];
    if (d) return d.name;
  }
  // W całej treści szukamy tylko pełnych nazw (min. 2 słowa), żeby „Piotr” pacjenta nie dał lekarza Piotra
  const t = normText_(text);
  const d = doctors.filter(x => names(x).some(a => a.trim().split(' ').length >= 2 && t.indexOf(a) >= 0))[0];
  return d ? d.name : (found || '');
}

function classifyRaw_(msg, boxes, ignore) {
  const clinic = boxes.map(b => b.email);
  const from = msg.getFrom() || '';
  const fromL = from.toLowerCase();
  const subject = msg.getSubject() || '';
  const subjL = subject.toLowerCase();
  const raw = String(msg.getPlainBody() || htmlToText_(msg.getBody()) || '').replace(/\r/g, '');
  const flat = raw.replace(/\s+/g, ' ').trim();
  const to = ((msg.getTo() || '') + ' ' + (msg.getCc() || '')).toLowerCase();
  const box = boxes.filter(b => to.indexOf(b.email) >= 0)[0];
  const it = {
    source: 'email', kind: '', name: '', phone: '', email: '', subject: subject, text: '', urgent: '',
    mailbox: box ? box.label : '', when: '', service: '', doctor: ''
  };

  // SMS i połączenia z telefonu (MacroDroid)
  let m = subject.match(/^CRM\s+(SMS|NIEODEBRANE)\s*(.*)$/i);
  if (m) {
    const rest = m[2].trim();
    const ph = rest.match(/\+?[\d\s-]{6,}/);
    it.phone = ph ? ph[0].trim() : '';
    it.name = (ph ? rest.replace(ph[0], '') : rest).replace(/[\[\]()]/g, '').trim();
    if (m[1].toUpperCase() === 'SMS') {
      if (digits_(it.phone).length < 9) return null; // SMS-y serwisowe (Play, kody, reklamy)
      it.source = 'sms'; it.kind = 'SMS od pacjenta'; it.text = flat.slice(0, 1000);
    } else {
      it.source = 'call'; it.kind = 'Nieodebrane połączenie — oddzwonić'; it.urgent = '1';
      if (!digits_(it.phone)) it.name = it.name || 'Numer zastrzeżony';
    }
    return it;
  }

  if (clinic.some(e => fromL.indexOf(e) >= 0)) return null; // nasze własne wysłane maile
  if (ignore.some(w => fromL.indexOf(w) >= 0 || subjL.indexOf(w) >= 0)) return null;

  // Booksy: „Anna Kowalska: nowa rezerwacja piątek, 27 listopada 2026 16:15”
  if (fromL.indexOf('booksy') >= 0 || /booksy/i.test(flat)) {
    it.source = 'booksy';
    it.name = subject.indexOf(':') > 0 ? subject.split(':')[0].trim() : displayName_(from);
    const when = subject.split(':').slice(1).join(':').replace(/^\s*(nowa\s+)?\S*\s*/i, '').trim();
    const pending = /oczekuje na potwierdzenie/i.test(flat);
    if (/odwoła|anulowa/i.test(subjL)) it.kind = 'Booksy: odwołanie wizyty';
    else if (/zmian|przełoż|przesun/i.test(subjL)) it.kind = 'Booksy: zmiana terminu';
    else if (/rezerwacj/i.test(subjL)) it.kind = pending ? 'Booksy: nowa rezerwacja — POTWIERDŹ' : 'Booksy: nowa rezerwacja';
    else it.kind = 'Booksy: ' + subject;
    if (pending) it.urgent = '1';
    it.phone = findPhone_(flat);
    it.email = findEmail_(flat, ['booksy']);
    const service = flat.match(/([A-ZĄĆĘŁŃÓŚŹŻ][^.!?]{3,90}?)\s+(\d+,\d{2}\s*zł)/);
    const worker = raw.match(/pracownik:\s*([^\n,]+)/i);
    const range = flat.match(/\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}/);
    it.when = (when || (subject.match(/\d{1,2}\s+\S+\s+\d{4}\s+\d{1,2}:\d{2}/) || [''])[0]) + (range ? ' (' + range[0].replace(/\s/g, '') + ')' : '');
    it.service = service ? service[1].trim() + ', ' + service[2] : '';
    it.doctor = worker ? worker[1].trim() : '';
    return it;
  }

  // ZnanyLekarz: „❌ Julia Kowalczyk odwołał(a) wizytę umówioną na sobota 14.11.2026 o 10:00”
  if (/znanylekarz|docplanner/i.test(fromL) || /znanylekarz/i.test(flat)) {
    it.source = 'znanylekarz';
    const clean = subject.replace(/^[^A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]+/, '');
    it.name = clean.split(/\s+(odwołał|umówił|zmienił|przełożył|przesunął|wysłał|napisał|zostawił|dodał|ma pytanie|zadał)/i)[0].trim();
    if (/odwoła/i.test(subjL)) it.kind = 'ZnanyLekarz: odwołanie wizyty';
    else if (/umówił|nowa wizyta|zarezerwowa/i.test(subjL)) it.kind = 'ZnanyLekarz: nowa wizyta';
    else if (/zmienił|przełożył|przesun/i.test(subjL)) it.kind = 'ZnanyLekarz: zmiana terminu';
    else if (/wiadomoś|pytani/i.test(subjL)) { it.kind = 'ZnanyLekarz: wiadomość od pacjenta'; it.urgent = '1'; }
    else it.kind = 'ZnanyLekarz: ' + subject.replace(/^[^A-Za-zĄĆĘŁŃÓŚŹŻ]+/, '');
    it.phone = findPhone_(flat);
    it.email = findEmail_(flat, ['znanylekarz', 'docplanner']);
    const when = subject.match(/\d{1,2}\.\d{1,2}\.\d{4}\s+o\s+\d{1,2}:\d{2}/) || flat.match(/\d{1,2}\s+\S+\s+\d{4}\s+o\s+\d{1,2}:\d{2}/);
    const service = raw.match(/([A-ZĄĆĘŁŃÓŚŹŻ][^()\n]{3,90}?\(\d+\s*min\))/);
    it.when = when ? when[0] : '';
    it.service = service ? service[1].trim() : '';
    it.doctor = findDoctor_(flat);
    if (/wiadomoś|pytani/i.test(it.kind)) it.text = flat.slice(0, 1000);
    return it;
  }

  // Formularz na stronie: „Twoje imię: Anna / Twój numer telefonu: 888…”
  m = flat.match(/Twoje imię:\s*(.+?)\s+Twój numer telefonu:\s*([+\d][\d\s-]{6,})/i);
  if (m) {
    it.source = 'www_form'; it.kind = 'Formularz na stronie — oddzwonić'; it.urgent = '1';
    it.name = m[1].trim(); it.phone = m[2].trim();
    formFields_(it, raw);
    it.text = flat.replace(m[0], '').trim().slice(0, 500);
    return it;
  }

  // Wiadomość ze strony: „New message from Anna (Nowa wiadomość od Anny)” / „-Anna ( 888… )”
  if (/new message from|nowa wiadomość od/i.test(subject)) {
    it.source = 'www_msg'; it.kind = 'Wiadomość ze strony — oddzwonić'; it.urgent = '1';
    const n = subject.match(/new message from\s+(.+?)(\s*\(|$)/i);
    it.name = n ? n[1].trim() : displayName_(from);
    it.phone = findPhone_(flat);
    it.email = findEmail_(flat, clinic);
    formFields_(it, raw);
    it.text = flat.slice(0, 1000);
    return it;
  }

  // WeTransfer: pacjent przesyła zdjęcia RTG / tomografię (link wygasa po kilku dniach!)
  if (/wetransfer/i.test(fromL) || /we\.tl\//i.test(flat)) {
    if (/pobran|downloaded|został[oa]? pobrane|wygas(a|ł) (jutro|dziś)|expire[sd]? (soon|today)/i.test(subjL)) return null; // potwierdzenia pobrania
    it.source = 'wetransfer'; it.urgent = '1';
    it.kind = 'WeTransfer: pacjent przesłał pliki (RTG / tomografia) — pobrać';
    const title = subject.replace(/\s+(sent successfully to|sent you|przesłał\(a\)|przesłał|przesłała|wysłał|wysłała).*$/i, '').replace(/^\W+/, '').trim();
    it.name = title.replace(/\s+(rtg|kt|ct|cbct|tomografi\S*|zdj[eę]ci\S*|pantomogram\S*|pliki?)$/i, '').trim() || displayName_(from);
    it.email = findEmail_(subject + ' ' + flat, ['wetransfer'].concat(clinic));
    const link = flat.match(/https?:\/\/we\.tl\/\S+/i);
    const exp = flat.match(/(Wygasa|Expires(?: on)?|Link wygaśnie)\s*:?\s*([^.·]{4,40}\d{4})/i);
    const size = flat.match(/\d+\s*(element|plik|item|file)\S*,?\s*(łącznie\s*)?[\d,.]+\s*[MGK]B/i);
    it.when = exp ? 'link ważny do: ' + exp[2].trim() : '';
    it.service = size ? size[0] : '';
    it.text = link ? 'Link: ' + link[0] : flat.slice(0, 300);
    return it;
  }

  // Zwykły mail (pacjent pisze bezpośrednio)
  it.kind = 'E-mail: ' + (subject || '(bez tematu)');
  it.name = displayName_(from);
  it.email = (from.match(/[^\s<>"]+@[^\s<>"]+/) || [''])[0].toLowerCase();
  it.phone = findPhone_(flat);
  it.text = raw.split(/\n\s*(On .{5,120}wrote:|W dniu .{5,160}pisze:|Dnia .{5,160}napisał)/)[0].replace(/\s+/g, ' ').trim().slice(0, 1000);
  return it;
}

/** Pola z formularza strony, jeśli formularz je ma (termin, usługa, lekarz). */
function formFields_(it, raw) {
  const f = re => { const m = String(raw).match(re); return m ? m[1].trim() : ''; };
  it.service = it.service || f(/(?:Usługa|Zabieg|Rodzaj wizyty|Temat)\s*:\s*(.+)/i);
  it.when = it.when || f(/(?:Termin|Data wizyty|Preferowany termin|Data)\s*:\s*(.+)/i);
  it.doctor = it.doctor || f(/(?:Lekarz|Specjalista|Do kogo)\s*:\s*(.+)/i);
}

/** „lek. dent. Anna Nowak”, „dr n. med. Jan Kowalski”, „Lekarz: …” */
function findDoctor_(text) {
  const m = String(text).match(/(?:lek\.\s*dent\.|lek\.|dr\.?\s*(?:n\.\s*med\.)?|lekarz:|specjalista:)\s*([A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+(?:[\s-]+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+)?)/);
  return m ? m[1].trim() : '';
}

/** „polishdent.pl@gmail.com=PL, dentonia.pl@gmail.com=WRO” → [{email, label}] */
function mailboxes_(s) {
  return list_(s.mailboxes || s.clinicEmails).map(x => {
    const p = x.split('=');
    return { email: p[0].trim(), label: (p[1] || p[0].split('@')[0]).trim().toUpperCase() };
  });
}

function displayName_(from) {
  const m = String(from).match(/^\s*"?([^"<]+?)"?\s*</);
  return m ? m[1].trim() : String(from).replace(/[<>]/g, '').trim();
}

function findPhone_(text) {
  const m = String(text).match(/(\+48[\s-]?)?\b\d{3}[\s-]?\d{3}[\s-]?\d{3}\b/);
  return m ? m[0].trim() : '';
}

function findEmail_(text, skip) {
  const all = String(text).match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g) || [];
  return (all.filter(e => !(skip || []).some(w => e.toLowerCase().indexOf(w) >= 0))[0] || '').toLowerCase();
}

function digits_(s) { return String(s || '').replace(/\D/g, ''); }

function list_(s) { return String(s || '').toLowerCase().split(',').map(x => x.trim()).filter(Boolean); }

function mailLink_(threadId) {
  let user = '';
  try { user = Session.getEffectiveUser().getEmail(); } catch (e) { /* brak */ }
  return 'https://mail.google.com/mail/' + (user ? '?authuser=' + encodeURIComponent(user) : 'u/0/') + '#all/' + threadId;
}

/**
 * Co 15 min: szuka nowych maili Diagdent „Nowe badanie w systemie Portal Lekarza”.
 * Z treści „…pojawiło się już badanie <BADANIE> dla pacjenta <IMIĘ NAZWISKO> (<DATA UR.>).”
 * wyciąga pacjenta. Jeśli jest otwarta sprawa tomografii tego pacjenta — przenosi ją na etap
 * „Tomografia gotowa — pokazać lekarzowi”. Jeśli nie — tworzy nową sprawę z imieniem z maila.
 */
function checkKtEmails() {
  const s = getSettings_();
  if (!s.ktEmailQuery) return;
  const label = GmailApp.getUserLabelByName(KT_LABEL) || GmailApp.createLabel(KT_LABEL);
  const threads = GmailApp.search('(' + s.ktEmailQuery + ') -label:' + KT_LABEL, 0, 30);

  threads.forEach(th => {
    th.getMessages().forEach(msg => {
      const subject = msg.getSubject() || '';
      const body = msg.getPlainBody() || '';
      const info = parseStudyEmail_(body) || parseStudyEmail_(htmlToText_(msg.getBody()));
      const link = mailLink_(th.getId());
      const now = Date.now();
      const stageIdx = 1; // „Tomografia gotowa — pokazać lekarzowi”
      const openKt = readTasks_().filter(t => t.status === 'open' && t.type === 'kt');
      const match = info ? matchPatient_(openKt, info.name) : matchPatientInText_(openKt, subject + ' ' + body);
      const what = info ? info.study + (info.birth ? ' (ur. ' + info.birth + ')' : '') : '„' + subject + '”';

      if (match) {
        match.stage = stageIdx;
        match.dueAt = now + TYPES.kt.stages[stageIdx].due * 36e5;
        match.updatedAt = now;
        match.updatedBy = BOT;
        match.history.push({ t: now, who: BOT, text: '📩 Przyszedł mail: badanie gotowe — ' + what + '. Pokaż lekarzowi. ' + link });
        saveTask(match);
      } else {
        saveTask({
          id: newId_(), createdAt: now, createdBy: BOT, type: 'kt',
          patient: info ? info.name : '❓ Ustalić pacjenta', phone: '', channel: 'E-mail', doctor: '',
          purpose: 'Gotowe badanie w Portalu Lekarza: ' + what + '. Pokazać lekarzowi i uzyskać odpowiedź: kiedy i na ile czasu umówić pacjenta.' +
            (info ? ' Uzupełnij telefon i lekarza.' : ' Nie udało się odczytać pacjenta z maila — sprawdź mail.'),
          stage: stageIdx, dueAt: now + 2 * 36e5, status: 'open', updatedAt: now, updatedBy: BOT,
          closedAt: '', source: 'email',
          history: [{ t: now, who: BOT, text: '📩 Utworzono z maila: ' + what + ' ' + link }]
        });
      }
    });
    th.addLabel(label);
  });
}

/** „…badanie Tomografia szczęki i żuchwy dla pacjenta Jan Kowalski (02.10.1969).” */
function parseStudyEmail_(text) {
  const t = String(text || '').replace(/[*_]/g, '').replace(/\s+/g, ' ');
  const m = t.match(/badanie\s+(.+?)\s+dla pacjent(?:a|ki)\s+(.+?)\s*\((\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{2,4})\)/i);
  if (!m) return null;
  return { study: m[1].trim(), name: m[2].trim(), birth: m[3] };
}

function htmlToText_(html) {
  return String(html || '').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
}

/** Normalizacja: małe litery, bez polskich znaków (Łukasz = lukasz). */
function tokens_(s) {
  return String(s || '').toLowerCase().replace(/ł/g, 'l').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .split(/[^a-zа-яіїєґ0-9]+/i).filter(w => w.length >= 3);
}

/** Dopasuj sprawę po imieniu i nazwisku z maila (kolejność słów dowolna). Tylko jednoznaczne trafienie. */
function matchPatient_(tasks, name) {
  const nt = tokens_(name);
  let c = tasks.filter(t => { const pt = tokens_(t.patient); return pt.length && pt.every(p => nt.indexOf(p) >= 0); });
  if (c.length === 1) return c[0];
  if (!c.length) {
    c = tasks.filter(t => tokens_(t.patient).some(p => nt.indexOf(p) >= 0));
    if (c.length === 1) return c[0];
  }
  return null;
}

/** Zapasowe dopasowanie, gdy format maila się zmieni: wszystkie słowa nazwiska pacjenta są w treści. */
function matchPatientInText_(tasks, text) {
  const tt = tokens_(text);
  const c = tasks.filter(t => { const pt = tokens_(t.patient); return pt.length && pt.every(p => tt.indexOf(p) >= 0); });
  return c.length === 1 ? c[0] : null;
}

/** Codziennie rano: raport dla właściciela o zaległych i „zawieszonych” sprawach. */
function dailyReport() {
  const s = getSettings_();
  if (!s.ownerEmail) return;
  const now = Date.now();
  const staleMs = (Number(s.staleHours) || 48) * 36e5;
  const all = readTasks_();
  const open = all.filter(t => t.status === 'open');
  const overdue = open.filter(t => t.dueAt && t.dueAt < now);
  const stale = open.filter(t => (!t.dueAt || t.dueAt < now) && now - (t.updatedAt || t.createdAt) > staleMs);
  const weekAgo = now - 7 * 864e5;
  const closedWeek = all.filter(t => t.status !== 'open' && (t.closedAt || 0) > weekAgo);

  const row = t => '<tr><td>' + esc_(TYPES[t.type] ? TYPES[t.type].icon + ' ' + TYPES[t.type].label : t.type) +
    '</td><td><b>' + esc_(t.patient) + '</b></td><td>' + esc_(stageName_(t)) + '</td><td>' + esc_(t.purpose) +
    '</td><td>' + Math.floor((now - (t.updatedAt || t.createdAt)) / 36e5) + ' godz. temu (' + esc_(t.updatedBy || t.createdBy) + ')</td></tr>';
  const table = (title, list) => list.length
    ? '<h3>' + title + ' — ' + list.length + '</h3><table border="1" cellpadding="6" style="border-collapse:collapse;font-size:13px">' +
      '<tr><th>Typ</th><th>Pacjent</th><th>Etap</th><th>O co chodzi</th><th>Ostatni ruch</th></tr>' + list.map(row).join('') + '</table>'
    : '<h3>' + title + ' — brak 👍</h3>';

  const html = '<div style="font-family:Arial,sans-serif">' +
    '<p>Otwartych spraw: <b>' + open.length + '</b>, zaległych: <b style="color:#c62828">' + overdue.length +
    '</b>, bez ruchu ponad ' + (s.staleHours || 48) + ' godz.: <b>' + stale.length + '</b>, zamknięto w 7 dni: <b>' + closedWeek.length + '</b>.</p>' +
    table('⏰ Zaległe', overdue) + table('🧊 Zawieszone bez ruchu', stale) +
    questionsHtml_(open) +
    '<p><a href="' + ScriptApp.getService().getUrl() + '">Otwórz tablicę</a></p></div>';

  MailApp.sendEmail({
    to: s.ownerEmail,
    subject: (s.clinicName || 'Gabinet') + ': ' + overdue.length + ' zaległych, ' + stale.length + ' zawieszonych spraw',
    htmlBody: html
  });
  if (isLastDayOfMonth_()) monthlyReport();
}

/**
 * Raport miesięczny (ostatniego dnia miesiąca o 20:30, za bieżący miesiąc; uruchomiony ręcznie w innym dniu — za poprzedni): sprawy, skrzynka, praca recepcji,
 * przypomnienia lekarzom, pytania i potrącenia na osobę.
 */
function monthlyReport() {
  const s = getSettings_();
  if (!s.ownerEmail) return;
  const per0 = reportPeriod_();
  const from = per0.from, to = per0.to;
  const inM = ts => ts >= from && ts < to;
  const MONTHS = ['styczeń', 'luty', 'marzec', 'kwiecień', 'maj', 'czerwiec', 'lipiec', 'sierpień', 'wrzesień', 'październik', 'listopad', 'grudzień'];
  const label = MONTHS[new Date(from).getMonth()] + ' ' + new Date(from).getFullYear();
  const tbl = (head, rows) => rows.length
    ? '<table border="1" cellpadding="6" style="border-collapse:collapse;font-size:13px"><tr>' + head.map(h => '<th>' + h + '</th>').join('') + '</tr>' +
      rows.map(r => '<tr>' + r.map(c => '<td>' + c + '</td>').join('') + '</tr>').join('') + '</table>'
    : '<p style="color:#666">brak danych</p>';
  const avg = (arr, unit) => arr.length ? (arr.reduce((a, b) => a + b, 0) / arr.length / unit).toFixed(1).replace('.', ',') : '—';

  const tasks = readTasks_();
  const created = tasks.filter(t => inM(t.createdAt));
  const done = tasks.filter(t => t.status === 'done' && inM(t.closedAt));
  const cancelled = tasks.filter(t => t.status === 'cancelled' && inM(t.closedAt));
  const open = tasks.filter(t => t.status === 'open');
  const now = Date.now();

  // Sprawy według typów
  const typeRows = Object.keys(TYPES).map(k => {
    const c = created.filter(t => t.type === k).length, dn = done.filter(t => t.type === k);
    const cn = cancelled.filter(t => t.type === k).length, op = open.filter(t => t.type === k).length;
    if (!c && !dn.length && !cn && !op) return null;
    return [TYPES[k].icon + ' ' + esc_(TYPES[k].label), c, dn.length, cn, op, avg(dn.map(t => t.closedAt - t.createdAt), 864e5)];
  }).filter(Boolean);

  // Praca recepcji (z historii spraw)
  const people = {};
  const P = n => (people[n] = people[n] || { actions: 0, created: 0, closed: 0, calls: 0, nags: 0, answers: 0 });
  const nagsByDoctor = {};
  const qAsked = [], qAnswered = [];
  tasks.forEach(t => (t.history || []).forEach(h => {
    if (h.q) {
      if (inM(h.t)) qAsked.push(h);
      if (h.a && inM(h.a.t)) { qAnswered.push(h.a.t - h.t); P(h.a.who).answers++; }
      return;
    }
    if (!inM(h.t) || !h.who || h.who === '🤖 Automat') return;
    const p = P(h.who);
    p.actions++;
    if (h.text.indexOf('🆕') === 0) p.created++;
    if (h.text.indexOf('📵') === 0) p.calls++;
    if (h.text.indexOf('🔔') === 0) { p.nags++; const doc = t.doctor || '(bez lekarza)'; nagsByDoctor[doc] = (nagsByDoctor[doc] || 0) + 1; }
  }));
  done.concat(cancelled).forEach(t => { if (t.updatedBy && t.updatedBy !== '🤖 Automat') P(t.updatedBy).closed++; });
  const owner = String(s.ownerName || 'Właściciel').trim();
  const peopleRows = Object.keys(people).filter(n => n !== owner).map(n => {
    const p = people[n]; return [esc_(n), p.actions, p.created, p.closed, p.calls, p.nags, p.answers];
  });

  // Skrzynka
  const SRC = { booksy: '📅 Booksy', znanylekarz: '🩺 ZnanyLekarz', www_form: '🌐 Formularz na stronie', www_msg: '💬 Wiadomość ze strony',
    wetransfer: '📦 WeTransfer', email: '✉️ E-mail', sms: '📱 SMS', call: '📞 Nieodebrane połączenia' };
  const inbox = readInbox_().filter(i => inM(i.receivedAt));
  const inboxRows = Object.keys(SRC).map(k => {
    const l = inbox.filter(i => i.source === k);
    if (!l.length) return null;
    const h = l.filter(i => i.status !== 'new' && i.handledAt);
    return [SRC[k], l.length, h.length, l.filter(i => i.taskId).length, (h.length ? avg(h.map(i => i.handledAt - i.receivedAt), 6e4) + ' min' : '—')];
  }).filter(Boolean);
  const handledBy = {};
  inbox.forEach(i => { if (i.handledBy) handledBy[i.handledBy] = (handledBy[i.handledBy] || 0) + 1; });

  const html = '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<h2>📅 Raport miesięczny: ' + label + '</h2>' +
    '<p>Nowych spraw: <b>' + created.length + '</b> · zrobionych: <b>' + done.length + '</b> · zamkniętych bez realizacji: <b>' + cancelled.length +
    '</b> · średnio dni do realizacji: <b>' + avg(done.map(t => t.closedAt - t.createdAt), 864e5) + '</b></p>' +
    '<p>Stan na dziś: otwartych <b>' + open.length + '</b>, zaległych <b style="color:#c62828">' + open.filter(t => t.dueAt && t.dueAt < now).length + '</b>.</p>' +
    '<h3>Sprawy według typów</h3>' + tbl(['Typ', 'Nowe', 'Zrobione', 'Bez realizacji', 'Otwarte teraz', 'Śr. dni'], typeRows) +
    REPORT_GROUPS.map(g => {
      const l = tasks.filter(t => g.types.indexOf(t.type) >= 0 && (inM(t.createdAt) || inM(t.closedAt || 0) || t.status === 'open'))
        .sort((a, b) => a.createdAt - b.createdAt);
      const st = t => t.status === 'done' ? '✅ ' + esc_(stageName_(t)) : t.status === 'cancelled' ? '✖️ bez realizacji' : '⏳ ' + esc_(stageName_(t)) + (t.dueAt && t.dueAt < now ? ' <b style="color:#c62828">(zaległe)</b>' : '');
      const head = ['Pacjent', 'Lekarz', 'O co chodzi', 'Status', 'Utworzono'].concat(g.types.length > 1 ? ['Typ'] : []);
      return '<h3>' + g.title + ' — ' + l.length + ' <span style="font-weight:normal;color:#666">(zrobione: ' + l.filter(t => t.status === 'done').length +
        ', otwarte: ' + l.filter(t => t.status === 'open').length + ')</span></h3>' +
        tbl(head, l.map(t => [
          '<b>' + esc_(t.patient) + '</b>', esc_(t.doctor || ''), esc_(String(t.purpose || '').slice(0, 140)), st(t),
          Utilities.formatDate(new Date(t.createdAt), Session.getScriptTimeZone(), 'dd.MM')].concat(g.types.length > 1 ? [TYPES[t.type].icon + ' ' + TYPES[t.type].label] : [])));
    }).join('') +
    '<h3>📥 Skrzynka: ' + inbox.length + ' wiadomości</h3>' + tbl(['Źródło', 'Wiadomości', 'Obsłużone', 'Zrobiono sprawę', 'Śr. czas reakcji'], inboxRows) +
    '<h3>👩‍💼 Praca recepcji</h3>' + tbl(['Kto', 'Działań', 'Utworzone sprawy', 'Zamknięte', 'Nieodebrane telefony', 'Przypomnienia lekarzom', 'Odpowiedzi na pytania'],
      peopleRows) +
    (Object.keys(handledBy).length ? '<p>Obsłużone wiadomości w skrzynce: ' + Object.keys(handledBy).map(n => esc_(n) + ' — ' + handledBy[n]).join(', ') + '</p>' : '') +
    '<h3>👨‍⚕️ Przypomnienia lekarzom (ile razy recepcja musiała przypominać)</h3>' +
    tbl(['Lekarz', 'Przypomnień'], Object.keys(nagsByDoctor).sort((a, b) => nagsByDoctor[b] - nagsByDoctor[a]).map(k => [esc_(k), nagsByDoctor[k]])) +
    '<h3>❓ Pytania</h3><p>Zadanych: <b>' + qAsked.length + '</b> · odpowiedzianych: <b>' + qAnswered.length + '</b> · średni czas odpowiedzi: <b>' +
    avg(qAnswered, 36e5) + ' godz.</b></p>' +
    debtsMonthHtml_() +
    '<p><a href="' + ScriptApp.getService().getUrl() + '">Otwórz tablicę</a></p></div>';

  MailApp.sendEmail({ to: s.ownerEmail, subject: (s.clinicName || 'Gabinet') + ': raport miesięczny — ' + label, htmlBody: html });
}

/* ===================== Pomocnicze ===================== */

function questionsHtml_(open) {
  const qs = [];
  open.forEach(t => (t.history || []).forEach(h => { if (h.q && !h.a) qs.push('<li><b>' + esc_(t.patient) + '</b>: ' + esc_(h.text.replace(/^❓\s*/, '')) + ' <i>(' + esc_(h.who) + (h.to ? ' → ' + esc_(h.to) : '') + ')</i></li>'); }));
  return '<h3>❓ Pytania bez odpowiedzi — ' + qs.length + '</h3>' + (qs.length ? '<ul>' + qs.join('') + '</ul>' : '');
}

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) out[r[0]] = String(r[1]).trim(); });
  return out;
}

function readDoctors_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DOCTORS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), phone: String(r[1]).replace(/\D/g, ''), aliases: list_(r[2]) }));
}

function readTasks_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, TASK_COLS.length).getValues()
    .filter(r => r[0]).map(rowToTask_);
}

function readShifts_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SHIFTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, SHIFT_COLS.length).getValues()
    .filter(r => r[0])
    .map(r => {
      const o = {};
      SHIFT_COLS.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].getTime() : r[i]; });
      return o;
    });
}

function readInbox_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INBOX);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, INBOX_COLS.length).getValues()
    .filter(r => r[0])
    .map(r => {
      const o = {};
      INBOX_COLS.forEach((c, i) => { o[c] = r[i] instanceof Date ? r[i].getTime() : r[i]; });
      o.msgId = String(o.msgId).replace(/^'/, ''); o.phone = String(o.phone).replace(/^'/, '');
      return o;
    });
}

function inboxToRow_(i) {
  return INBOX_COLS.map(c => {
    const v = i[c];
    if ((c === 'receivedAt' || c === 'handledAt') && v) return new Date(v);
    if (c === 'phone' || c === 'msgId') return v ? "'" + String(v).replace(/^'/, '') : '';
    return v === undefined || v === null ? '' : v;
  });
}

function rowToTask_(r) {
  const t = {};
  TASK_COLS.forEach((c, i) => {
    let v = r[i];
    if (v instanceof Date) v = v.getTime();
    t[c] = v;
  });
  t.stage = Number(t.stage) || 0;
  t.phone = String(t.phone || '').replace(/^'/, '');
  try { t.history = t.history ? JSON.parse(t.history) : []; } catch (e) { t.history = []; }
  return t;
}

function taskToRow_(t) {
  return TASK_COLS.map(c => {
    const v = t[c];
    if (c === 'history') return JSON.stringify(v || []);
    if (c === 'phone') return v ? "'" + String(v).replace(/^'/, '') : ''; // „+48…” jako tekst, nie liczba
    if (['createdAt', 'dueAt', 'updatedAt', 'closedAt'].indexOf(c) >= 0) return v ? new Date(v) : '';
    return v === undefined || v === null ? '' : v;
  });
}

function findRow_(sh, id) {
  if (!id || sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** Scal historię z arkusza (a) i z przeglądarki (b). Ten sam wpis: wygrywa nowsza wersja, ale odpowiedź na pytanie nigdy nie ginie. */
function mergeHistory_(a, b) {
  const map = {};
  const order = [];
  (a || []).concat(b || []).forEach(h => {
    const k = h.t + '|' + h.who + '|' + h.text;
    if (!map[k]) { order.push(k); map[k] = h; return; }
    if (map[k].a && !h.a) return;
    map[k] = h;
  });
  return order.map(k => map[k]).sort((x, y) => x.t - y.t);
}

/** Nowe odpowiedzi na pytania właściciela → e-mail do właściciela. */
function notifyAnswers_(before, after) {
  const s = getSettings_();
  if (!s.ownerEmail) return;
  const owner = String(s.ownerName || 'Właściciel').trim();
  const old = {};
  (before || []).forEach(h => { if (h.q) old[h.t + '|' + h.who] = h; });
  (after.history || []).forEach(h => {
    if (!h.q || !h.a || h.who !== owner) return;
    const o = old[h.t + '|' + h.who];
    if (o && o.a) return;
    MailApp.sendEmail({
      to: s.ownerEmail,
      subject: '✅ Odpowiedź: ' + after.patient + ' — ' + h.a.who,
      htmlBody: '<div style="font-family:Arial,sans-serif"><p><b>Pytanie:</b> ' + esc_(h.text.replace(/^❓\s*/, '')) + '</p><p><b>' + esc_(h.a.who) + ':</b> ' + esc_(h.a.text) +
        '</p><p style="color:#666">Sprawa: ' + esc_(after.patient) + ' — ' + esc_(after.purpose) + '</p><p><a href="' + ScriptApp.getService().getUrl() + '">Otwórz tablicę</a></p></div>'
    });
  });
}

function stageName_(t) {
  const type = TYPES[t.type];
  return type && type.stages[t.stage] ? type.stages[t.stage].n : '';
}

function newId_() {
  return 'T' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
