// Lokalne środowisko testowe: uruchamia Code.gs w Node z imitacją Arkusza Google.
// node dev/server.js  →  http://localhost:8787
const fs = require('fs'), path = require('path'), http = require('http'), vm = require('vm');
const root = path.join(__dirname, '..');

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) { const v = (this.sh.rows[this.r - 1 + i] || [])[this.c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; }
  setValues(v) { v.forEach((row, i) => { const R = this.sh.rows[this.r - 1 + i] = this.sh.rows[this.r - 1 + i] || []; row.forEach((x, j) => R[this.c - 1 + j] = x); }); return this; }
  getValue() { return this.getValues()[0][0]; }
  setValue(v) { return this.setValues([[v]]); }
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
const ownerSheets = {}; // arkusz potrąceń na koncie właściciela
const ownerSs = { getSheetByName: n => ownerSheets[n] || null, insertSheet: n => (ownerSheets[n] = new Sheet(n)) };
const sentMail = [];
const ctx = {
  console, Date, JSON, Math,
  SpreadsheetApp: { getActive: () => ss, openByUrl: u => { if (!/owner/.test(u)) throw new Error('no access'); return ownerSs; }, getUi: () => ({ alert() {}, createMenu: () => ({ addItem() { return this; }, addToUi() {} }) }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased() { return this; }, everyMinutes() { return this; }, everyDays() { return this; }, atHour() { return this; }, nearMinute() { return this; }, create() {} }), getService: () => ({ getUrl: () => 'http://localhost:8787' }) },
  MailApp: { sendEmail: m => sentMail.push(m) },
  Utilities: { formatDate: (d, tz, f) => f === 'MM.yyyy' ? (d.getMonth() + 1) + '.' + d.getFullYear() : f === 'HH:mm' ? String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') : d.toLocaleDateString('pl-PL') },
  GmailApp: null,
  HtmlService: {}
};
ctx.__viewer = '';
ctx.Session = { getEffectiveUser: () => ({ getEmail: () => 'polishdent.pl@gmail.com' }), getActiveUser: () => ({ getEmail: () => ctx.__viewer }), getScriptTimeZone: () => 'Europe/Warsaw' };
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'Code.gs'), 'utf8'), ctx);
ctx.setup();
// ustawienia demo
sheets['Ustawienia'].rows.forEach(r => { if (r[0] === 'ownerEmail') r[1] = 'owner@example.com'; });
sheets['Lekarze'].rows.push(['dr Nowak', '48501112233', 'lek. dent. Anna Nowak, Anna Nowak'], ['dr Wiśniewska', '48601234567', 'Karolina Wiśniewska']);
sheets['Ustawienia'].rows.forEach(r => { if (r[0] === 'debtsSheetUrl') r[1] = 'https://docs.google.com/spreadsheets/d/owner-debts'; });
ctx.__ownerSheets = ownerSheets;

// imitacja Gmaila do testu checkKtEmails
const inbox = [];
ctx.GmailApp = {
  getUserLabelByName: () => ({}), createLabel: () => ({}),
  // t = {id, subject, body, from, to, date, kt}; zapytania z "Portal Lekarza" zwracają tylko maile Diagdent
  search: q => inbox.filter(t => (/Portal Lekarza/.test(q) ? t.kt && !t.done : !t.kt)).map(t => ({
    getId: () => t.id, getMessageCount: () => 1,
    getMessages: () => [{
      getId: () => 'm' + t.id, getSubject: () => t.subject, getPlainBody: () => t.body, getBody: () => t.body,
      getFrom: () => t.from || '', getTo: () => t.to || 'polishdent.pl@gmail.com', getCc: () => '', getDate: () => new Date(t.date || Date.now()), getHeader: h => (t.headers || {})[h] || ''
    }],
    addLabel() { t.done = true; }
  }))
};
const props = {};
ctx.PropertiesService = { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) };
ctx.__inbox = inbox; ctx.__sentMail = sentMail; ctx.__sheets = sheets;
module.exports = ctx;

if (require.main === module) {
  require('./seed')(ctx);
  require('./sample-mails')(ctx);
  ctx.checkMail();
  const shim = `<script>window.google={script:{run:(function mk(ok,fail){return new Proxy({},{get(_,k){if(k==='withSuccessHandler')return f=>mk(f,fail);if(k==='withFailureHandler')return f=>mk(ok,f);return a=>fetch('/api/'+k,{method:'POST',body:JSON.stringify(a===undefined?null:a)}).then(r=>r.json()).then(j=>j&&j.error?fail&&fail(new Error(j.error)):ok&&ok(j))}})})()}};</script>`;
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
