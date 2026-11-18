// Локальний стенд для перевірки: запускає Code.gs у Node з імітацією Google Таблиці.
// node dev/server.js  →  http://localhost:8787
const fs = require('fs'), path = require('path'), http = require('http'), vm = require('vm');
const root = path.join(__dirname, '..');

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.rows[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; }
  setValues(v) { v.forEach((row, i) => { const R = this.sh.rows[this.r - 1 + i] = this.sh.rows[this.r - 1 + i] || []; row.forEach((x, j) => R[this.c - 1 + j] = x); }); return this; }
  setFontWeight() { return this; } setBackground() { return this; }
}
class Sheet {
  constructor(name) { this.name = name; this.rows = []; }
  getLastRow() { return this.rows.length; }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  appendRow(row) { this.rows.push(row.slice()); }
  setFrozenRows() {} setColumnWidth() { return this; }
}
const sheets = {};
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = new Sheet(n)) };
const sentMail = [];
const ctx = {
  console, Date, JSON, Math,
  SpreadsheetApp: { getActive: () => ss, getUi: () => ({ alert() {}, createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased() { return this; }, everyMinutes() { return this; }, everyDays() { return this; }, atHour() { return this; }, create() {} }), getService: () => ({ getUrl: () => 'http://localhost:8787' }) },
  MailApp: { sendEmail: m => sentMail.push(m) },
  GmailApp: null,
  HtmlService: {}
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'Code.gs'), 'utf8'), ctx);
ctx.setup();
// демо-налаштування
const set = sheets['Налаштування'];
set.rows.forEach(r => { if (r[0] === 'ownerEmail') r[1] = 'owner@example.com'; if (r[0] === 'admins') r[1] = 'Оля, Марта, Власник'; });
sheets['Лікарі'].rows = [sheets['Лікарі'].rows[0], ['Петренко', '380501112233'], ['Коваль', '380671234567']];

// Gmail-імітація для перевірки checkKtEmails
const inbox = [];
ctx.GmailApp = {
  getUserLabelByName: () => ({}), createLabel: () => ({}),
  search: () => inbox.filter(t => !t.done).map(t => ({ getId: () => t.id, getMessageCount: () => 1, getMessages: () => [{ getSubject: () => t.subject, getPlainBody: () => t.body }], addLabel() { t.done = true; } }))
};
ctx.__inbox = inbox; ctx.__sentMail = sentMail; ctx.__sheets = sheets;
module.exports = ctx;

if (require.main === module) {
  require('./seed')(ctx);
  const shim = `<script>window.google={script:{run:(function mk(ok,fail){return new Proxy({},{get(_,k){if(k==='withSuccessHandler')return f=>mk(f,fail);if(k==='withFailureHandler')return f=>mk(ok,f);return a=>fetch('/api/'+k,{method:'POST',body:JSON.stringify(a===undefined?null:a)}).then(r=>r.json()).then(j=>j.error?fail&&fail(new Error(j.error)):ok&&ok(j))}})})()}};</script>`;
  http.createServer((req, res) => {
    if (req.url.startsWith('/api/')) {
      let b = ''; req.on('data', d => b += d); req.on('end', () => {
        try { const fn = req.url.slice(5); const arg = JSON.parse(b || 'null'); const out = ctx[fn](arg === null ? undefined : arg); res.end(JSON.stringify(out === undefined ? {} : out)); }
        catch (e) { console.error(e); res.end(JSON.stringify({ error: e.message })); }
      }); return;
    }
    res.setHeader('content-type', 'text/html; charset=utf-8');
    res.end(fs.readFileSync(path.join(root, 'Index.html'), 'utf8').replace('<head>', '<head>' + shim));
  }).listen(8787, () => console.log('http://localhost:8787'));
}
