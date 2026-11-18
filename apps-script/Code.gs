/**
 * «Передача зміни» — єдина дошка справ адміністраторів стоматології.
 * Google Apps Script, прив'язаний до Google Таблиці.
 *
 * Аркуші, які створює setup():
 *   Справи        — усі справи (одна справа = один рядок)
 *   Зміни         — журнал передачі змін
 *   Лікарі        — імена лікарів і їхні номери WhatsApp
 *   Налаштування  — адміністратори, пошта власника, фільтр листів про КТ тощо
 */

const SHEET_TASKS = 'Справи';
const SHEET_SHIFTS = 'Зміни';
const SHEET_DOCTORS = 'Лікарі';
const SHEET_SETTINGS = 'Налаштування';
const KT_LABEL = 'CRM-KT-done';

const TASK_COLS = [
  'id', 'createdAt', 'createdBy', 'type', 'patient', 'phone', 'channel', 'doctor',
  'purpose', 'stage', 'dueAt', 'status', 'updatedAt', 'updatedBy', 'closedAt',
  'source', 'history'
];
const TASK_HEADERS = [
  'ID', 'Створено', 'Хто створив', 'Тип', 'Пацієнт', 'Телефон', 'Канал', 'Лікар',
  'Суть / навіщо', 'Етап', 'Термін', 'Статус', 'Оновлено', 'Хто оновив', 'Закрито',
  'Джерело', 'Історія (JSON)'
];
const SHIFT_COLS = ['id', 'at', 'from', 'note', 'openCount', 'overdueCount', 'acceptedBy', 'acceptedAt'];
const SHIFT_HEADERS = ['ID', 'Коли здано', 'Хто здав', 'Нотатка', 'Відкритих справ', 'Прострочених', 'Хто прийняв', 'Коли прийнято'];

/**
 * Типи справ і їхні етапи.
 *   due   — через скільки годин після переходу на етап справа стає простроченою;
 *   next  — «природний» наступний етап (за замовчуванням наступний за списком);
 *   retry — етап «не відповів», на який переводить кнопка 📵;
 *   nag   — етап очікування відповіді лікаря (кнопка 🔔 «Нагадати лікарю»);
 *   ask   — питання, на яке треба відповісти при переході на етап;
 *   done  — фінальний етап, справа закривається.
 */
const TYPES = {
  kt: {
    label: 'КТ / план лікування', icon: '🦷', doctorRequired: true,
    stages: [
      { n: 'Направлення на КТ надіслано', due: 72 },
      { n: 'КТ готова — показати лікарю', due: 2 },
      { n: 'Чекаємо відповідь лікаря', due: 4, nag: true },
      { n: 'Лікар відповів — записати пацієнта', due: 4, ask: 'Що сказав лікар: коли і на скільки часу записати?' },
      { n: 'Пацієнта записано', done: true }
    ]
  },
  call_after: {
    label: 'Дзвінок після процедури', icon: '📞',
    stages: [
      { n: 'Зателефонувати', due: 24, next: 3 },
      { n: 'Не відповів — передзвонити', due: 3, retry: true, next: 3 },
      { n: 'Є скарги — передано лікарю', due: 4, nag: true, next: 3, ask: 'Які скарги? Що передали лікарю?' },
      { n: 'Поговорили — все добре', done: true }
    ]
  },
  lab: {
    label: 'Лабораторія (коронки, вініри)', icon: '🧪',
    stages: [
      { n: 'Відправлено в лабораторію', due: 120 },
      { n: 'Робота готова — записати на фіксацію', due: 24, next: 3 },
      { n: 'Не відповів — передзвонити', due: 3, retry: true, next: 3 },
      { n: 'Записано на фіксацію', done: true }
    ]
  },
  confirm: {
    label: 'Підтвердження запису', icon: '✅',
    stages: [
      { n: 'Підтвердити запис', due: 24, next: 2 },
      { n: 'Не відповів — повторити', due: 3, retry: true, next: 2 },
      { n: 'Підтверджено', done: true }
    ]
  },
  revisit: {
    label: 'Повторний візит', icon: '🔁',
    stages: [
      { n: 'Запросити на візит', due: 24, next: 2 },
      { n: 'Не відповів — повторити', due: 24, retry: true, next: 2 },
      { n: 'Записано', done: true }
    ]
  },
  other: {
    label: 'Інше', icon: '📌',
    stages: [
      { n: 'Зробити', due: 8 },
      { n: 'В процесі', due: 8 },
      { n: 'Виконано', done: true }
    ]
  }
};

const DEFAULT_SETTINGS = [
  ['clinicName', 'Стоматологія', 'Назва, що показується вгорі'],
  ['admins', 'Адміністратор 1, Адміністратор 2, Власник', 'Хто працює з дошкою (через кому)'],
  ['ownerEmail', '', 'Пошта власника: сюди щоранку приходить звіт про прострочені справи'],
  ['reportHour', '8', 'О котрій годині надсилати ранковий звіт (0–23)'],
  ['staleHours', '48', 'Після скількох годин без руху справа вважається «завислою»'],
  ['ktEmailQuery', '', 'Пошуковий запит Gmail для листів «томографія готова», напр.: from:(info@ct-center.ua) newer_than:7d. Порожньо = автоматизацію вимкнено']
];

/* ===================== Встановлення ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🦷 Дошка справ')
    .addItem('1. Початкове налаштування', 'setup')
    .addItem('Перевірити пошту КТ зараз', 'checkKtEmails')
    .addItem('Надіслати звіт власнику зараз', 'dailyReport')
    .addToUi();
}

/** Запустіть один раз: створює аркуші й автоматичні тригери. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_TASKS, TASK_HEADERS);
  ensureSheet_(ss, SHEET_SHIFTS, SHIFT_HEADERS);
  const doctors = ensureSheet_(ss, SHEET_DOCTORS, ['Лікар', 'WhatsApp (380XXXXXXXXX)']);
  if (doctors.getLastRow() < 2) {
    doctors.getRange(2, 1, 2, 2).setValues([['Лікар 1', ''], ['Лікар 2', '']]);
  }
  const settings = ensureSheet_(ss, SHEET_SETTINGS, ['Ключ', 'Значення', 'Опис']);
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
    SpreadsheetApp.getUi().alert('Готово! Заповніть аркуші «Налаштування» і «Лікарі», потім опублікуйте веб-застосунок (див. інструкцію).');
  } catch (e) { /* запуск не з таблиці */ }
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

/* ===================== Веб-застосунок ===================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Дошка справ адміністраторів')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Усе, що потрібно інтерфейсу: налаштування, відкриті справи, закриті за 30 днів, журнал змін. */
function getData() {
  const s = getSettings_();
  const monthAgo = Date.now() - 30 * 864e5;
  const tasks = readTasks_().filter(t => t.status === 'open' || (t.closedAt || 0) > monthAgo);
  return {
    now: Date.now(),
    config: {
      types: TYPES,
      clinicName: s.clinicName || 'Стоматологія',
      admins: String(s.admins || '').split(',').map(x => x.trim()).filter(Boolean),
      staleHours: Number(s.staleHours) || 48,
      doctors: readDoctors_(),
      ktAutomation: !!s.ktEmailQuery
    },
    tasks: tasks,
    shifts: readShifts_().slice(-15)
  };
}

/** Створити або оновити справу. Історію зливаємо, щоб дві адміністраторки не затерли записи одна одної. */
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

/** Здати зміну (entry без acceptedBy) або прийняти (entry з id і acceptedBy). */
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

/* ===================== Автоматизація ===================== */

/**
 * Кожні 15 хв: шукає нові листи «томографія готова».
 * Якщо в листі є прізвище пацієнта з відкритої справи КТ — переводить її на етап «КТ готова — показати лікарю».
 * Якщо не знайшли — створює нову справу, щоб адміністратор розібрався.
 */
function checkKtEmails() {
  const s = getSettings_();
  if (!s.ktEmailQuery) return;
  const label = GmailApp.getUserLabelByName(KT_LABEL) || GmailApp.createLabel(KT_LABEL);
  const threads = GmailApp.search('(' + s.ktEmailQuery + ') -label:' + KT_LABEL, 0, 30);
  if (!threads.length) return;

  threads.forEach(th => {
    const msg = th.getMessages()[th.getMessageCount() - 1];
    const subject = msg.getSubject() || '';
    const text = norm_(subject + ' ' + msg.getPlainBody());
    const openKt = readTasks_().filter(t => t.status === 'open' && t.type === 'kt');
    const match = openKt.find(t => nameInText_(t.patient, text));
    const now = Date.now();
    const link = 'https://mail.google.com/mail/u/0/#all/' + th.getId();
    const stageIdx = 1; // «КТ готова — показати лікарю»

    if (match) {
      match.stage = stageIdx;
      match.dueAt = now + TYPES.kt.stages[stageIdx].due * 36e5;
      match.updatedAt = now;
      match.updatedBy = '🤖 Автомат';
      match.history.push({ t: now, who: '🤖 Автомат', text: 'Прийшов лист: КТ готова («' + subject + '»). Покажіть лікарю. ' + link });
      saveTask(match);
    } else {
      saveTask({
        id: newId_(), createdAt: now, createdBy: '🤖 Автомат', type: 'kt',
        patient: '❓ Уточнити пацієнта', phone: '', channel: 'email', doctor: '',
        purpose: 'Прийшов лист про готову томографію: «' + subject + '». Знайдіть пацієнта, впишіть ім\'я та лікаря.',
        stage: stageIdx, dueAt: now + 2 * 36e5, status: 'open', updatedAt: now, updatedBy: '🤖 Автомат',
        closedAt: '', source: 'email',
        history: [{ t: now, who: '🤖 Автомат', text: 'Створено з листа «' + subject + '» ' + link }]
      });
    }
    th.addLabel(label);
  });
}

/** Щоранку: звіт власнику про прострочені й «завислі» справи. */
function dailyReport() {
  const s = getSettings_();
  if (!s.ownerEmail) return;
  const now = Date.now();
  const staleMs = (Number(s.staleHours) || 48) * 36e5;
  const open = readTasks_().filter(t => t.status === 'open');
  const overdue = open.filter(t => t.dueAt && t.dueAt < now);
  const stale = open.filter(t => (!t.dueAt || t.dueAt < now) && now - (t.updatedAt || t.createdAt) > staleMs);
  const weekAgo = now - 7 * 864e5;
  const closedWeek = readTasks_().filter(t => t.status !== 'open' && (t.closedAt || 0) > weekAgo);

  const row = t => '<tr><td>' + esc_(TYPES[t.type] ? TYPES[t.type].icon + ' ' + TYPES[t.type].label : t.type) +
    '</td><td><b>' + esc_(t.patient) + '</b></td><td>' + esc_(stageName_(t)) + '</td><td>' + esc_(t.purpose) +
    '</td><td>' + Math.floor((now - (t.updatedAt || t.createdAt)) / 36e5) + ' год тому (' + esc_(t.updatedBy || t.createdBy) + ')</td></tr>';
  const table = (title, list) => list.length
    ? '<h3>' + title + ' — ' + list.length + '</h3><table border="1" cellpadding="6" style="border-collapse:collapse;font-size:13px">' +
      '<tr><th>Тип</th><th>Пацієнт</th><th>Етап</th><th>Суть</th><th>Останній рух</th></tr>' + list.map(row).join('') + '</table>'
    : '<h3>' + title + ' — немає 👍</h3>';

  const html = '<div style="font-family:Arial,sans-serif">' +
    '<p>Відкритих справ: <b>' + open.length + '</b>, прострочених: <b style="color:#c62828">' + overdue.length +
    '</b>, без руху понад ' + (s.staleHours || 48) + ' год: <b>' + stale.length + '</b>, закрито за 7 днів: <b>' + closedWeek.length + '</b>.</p>' +
    table('⏰ Прострочені', overdue) + table('🧊 Зависли без руху', stale) +
    '<p><a href="' + ScriptApp.getService().getUrl() + '">Відкрити дошку</a></p></div>';

  MailApp.sendEmail({
    to: s.ownerEmail,
    subject: (s.clinicName || 'Стоматологія') + ': ' + overdue.length + ' прострочених, ' + stale.length + ' завислих справ',
    htmlBody: html
  });
}

/* ===================== Допоміжні ===================== */

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

function norm_(s) {
  return String(s || '').toLowerCase().replace(/[ʼ'’`]/g, '').replace(/\s+/g, ' ');
}

/** Пацієнт «знайдений» у листі, якщо там є його прізвище (перше слово) або всі слова імені. */
function nameInText_(patient, text) {
  const parts = norm_(patient).split(' ').filter(p => p.length >= 3);
  if (!parts.length) return false;
  return parts.every(p => text.indexOf(p) >= 0) || text.indexOf(parts[0]) >= 0;
}

function newId_() {
  return 'T' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
