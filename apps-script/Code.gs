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
  kt: {
    label: 'Tomografia / plan leczenia', icon: '🦷',
    stages: [
      { n: 'Skierowanie na tomografię wysłane', due: 72 },
      { n: 'Tomografia gotowa — pokazać lekarzowi', due: 2 },
      { n: 'Czekamy na odpowiedź lekarza', due: 4, nag: true },
      { n: 'Lekarz odpowiedział — umówić pacjenta', due: 4, ask: 'Co powiedział lekarz: kiedy i na ile czasu umówić?' },
      { n: 'Pacjent umówiony', done: true }
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
    label: 'Inne', icon: '📌',
    stages: [
      { n: 'Do zrobienia', due: 8 },
      { n: 'W trakcie', due: 8 },
      { n: 'Zrobione', done: true }
    ]
  }
};

const DEFAULT_SETTINGS = [
  ['clinicName', 'Gabinet stomatologiczny', 'Nazwa wyświetlana u góry'],
  ['admins', 'Nadia N, Nadia R, Tosia, Właściciel', 'Kto pracuje z tablicą (po przecinku)'],
  ['ownerEmail', '', 'E-mail właściciela: codziennie rano przychodzi tu raport o zaległych sprawach'],
  ['reportHour', '8', 'O której godzinie wysyłać poranny raport (0–23)'],
  ['staleHours', '48', 'Po ilu godzinach bez ruchu (po terminie) sprawa jest „zawieszona”'],
  ['ktEmailQuery', '"Portal Lekarza" "Nowe badanie"', 'Zapytanie Gmail dla maili Diagdent o gotowej tomografii. Puste = automatyzacja wyłączona']
];

/* ===================== Instalacja ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🦷 Tablica spraw')
    .addItem('1. Konfiguracja początkowa', 'setup')
    .addItem('Sprawdź teraz maile z tomografią', 'checkKtEmails')
    .addItem('Wyślij teraz raport właścicielowi', 'dailyReport')
    .addToUi();
}

/** Uruchom raz: tworzy arkusze i automatyczne wyzwalacze. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_TASKS, TASK_HEADERS);
  ensureSheet_(ss, SHEET_SHIFTS, SHIFT_HEADERS);
  const doctors = ensureSheet_(ss, SHEET_DOCTORS, ['Lekarz', 'WhatsApp (48XXXXXXXXX)']);
  if (doctors.getLastRow() < 2) {
    doctors.getRange(2, 1, 2, 2).setValues([['Lekarz 1', ''], ['Lekarz 2', '']]);
  }
  const settings = ensureSheet_(ss, SHEET_SETTINGS, ['Klucz', 'Wartość', 'Opis']);
  const existing = settings.getLastRow() > 1
    ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  DEFAULT_SETTINGS.forEach(row => { if (existing.indexOf(row[0]) < 0) settings.appendRow(row); });
  settings.setColumnWidth(2, 320).setColumnWidth(3, 520);

  ScriptApp.getProjectTriggers().forEach(t => {
    if (['checkKtEmails', 'dailyReport'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkKtEmails').timeBased().everyMinutes(15).create();
  const hour = Math.min(23, Math.max(0, parseInt(getSettings_().reportHour, 10) || 8));
  ScriptApp.newTrigger('dailyReport').timeBased().everyDays(1).atHour(hour).create();

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
  return {
    now: Date.now(),
    config: {
      types: TYPES,
      clinicName: s.clinicName || 'Gabinet stomatologiczny',
      admins: String(s.admins || '').split(',').map(x => x.trim()).filter(Boolean),
      staleHours: Number(s.staleHours) || 48,
      doctors: readDoctors_(),
      ktAutomation: !!s.ktEmailQuery
    },
    tasks: tasks,
    shifts: readShifts_().slice(-15)
  };
}

/** Utwórz lub zaktualizuj sprawę. Historię scalamy, żeby dwie osoby nie nadpisały sobie wpisów. */
function saveTask(task) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_TASKS);
    const rowIdx = findRow_(sh, task.id);
    if (rowIdx > 0) {
      const current = rowToTask_(sh.getRange(rowIdx, 1, 1, TASK_COLS.length).getValues()[0]);
      task.history = mergeHistory_(current.history, task.history);
      sh.getRange(rowIdx, 1, 1, TASK_COLS.length).setValues([taskToRow_(task)]);
    } else {
      sh.appendRow(taskToRow_(task));
    }
  } finally {
    lock.releaseLock();
  }
  return getData();
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

/* ===================== Automatyzacja ===================== */

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
      const link = 'https://mail.google.com/mail/u/0/#all/' + th.getId();
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
    '<p><a href="' + ScriptApp.getService().getUrl() + '">Otwórz tablicę</a></p></div>';

  MailApp.sendEmail({
    to: s.ownerEmail,
    subject: (s.clinicName || 'Gabinet') + ': ' + overdue.length + ' zaległych, ' + stale.length + ' zawieszonych spraw',
    htmlBody: html
  });
}

/* ===================== Pomocnicze ===================== */

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
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(r => r[0])
    .map(r => ({ name: String(r[0]).trim(), phone: String(r[1]).replace(/\D/g, '') }));
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

function rowToTask_(r) {
  const t = {};
  TASK_COLS.forEach((c, i) => {
    let v = r[i];
    if (v instanceof Date) v = v.getTime();
    t[c] = v;
  });
  t.stage = Number(t.stage) || 0;
  try { t.history = t.history ? JSON.parse(t.history) : []; } catch (e) { t.history = []; }
  return t;
}

function taskToRow_(t) {
  return TASK_COLS.map(c => {
    const v = t[c];
    if (c === 'history') return JSON.stringify(v || []);
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

function mergeHistory_(a, b) {
  const seen = {};
  const out = [];
  (a || []).concat(b || []).forEach(h => {
    const k = h.t + '|' + h.who + '|' + h.text;
    if (!seen[k]) { seen[k] = 1; out.push(h); }
  });
  return out.sort((x, y) => x.t - y.t);
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
