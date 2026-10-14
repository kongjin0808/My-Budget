/**
 * Pocket Ledger - Google Apps Script backend
 * Paste this into the Apps Script editor of your Google Sheet
 * (Extensions > Apps Script), then deploy it as a Web app.
 *
 * Sheets created automatically:
 *   Expenses     one row per day      Date | Breakfast | Lunch | Dinner | Others | OthersNote | Total
 *   Investments  one row per aspect   Date | Aspect | Amount | Grow
 *   Todo         one row per item     Month | Aspect | Amount | Dateline | Paid
 */

// >>> CHANGE THIS to a long random string, and enter the same value in the app's Settings.
const TOKEN = 'CHANGE_ME_TO_A_LONG_RANDOM_STRING';

const SCHEMA = {
  Expenses:    { headers: ['Date', 'Breakfast', 'Lunch', 'Dinner', 'Others', 'OthersNote', 'Total'], textCols: [1] },
  Investments: { headers: ['Date', 'Aspect', 'Amount', 'Grow'], textCols: [1] },
  Todo:        { headers: ['Month', 'Aspect', 'Amount', 'Dateline', 'Paid'], textCols: [1, 4] }
};

/** Run this once from the editor (select "setup" > Run) to create the sheets and grant permissions. */
function setup() {
  Object.keys(SCHEMA).forEach(sheet_);
}

/* ------------------------------ entry points ------------------------------ */

function doGet(e) {
  return run_((e && e.parameter) || {});
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'Bad request body.' });
  }
  return run_(body);
}

function run_(req) {
  try {
    if (TOKEN === 'CHANGE_ME_TO_A_LONG_RANDOM_STRING') {
      return out_({ ok: false, error: 'Set your own TOKEN in Code.gs and redeploy.' });
    }
    if (req.token !== TOKEN) {
      return out_({ ok: false, error: 'Wrong token. It must match TOKEN in Code.gs.' });
    }
    switch (req.action) {
      case 'ping':            return out_({ ok: true, message: 'pong' });
      case 'getExpenses':     return out_(getExpenses_(req));
      case 'saveExpense':     return out_(saveExpense_(req));
      case 'getInvestments':  return out_(getInvestments_(req));
      case 'saveInvestments': return out_(saveInvestments_(req));
      case 'getTodo':         return out_(getTodo_(req));
      case 'saveTodo':        return out_(saveTodo_(req));
      default:                return out_({ ok: false, error: 'Unknown action: ' + req.action });
    }
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}

/* -------------------------------- Expenses -------------------------------- */

function getExpenses_(req) {
  const from = str_(req.from) || '0000-00-00';
  const to = str_(req.to) || '9999-99-99';
  const rows = readRows_(sheet_('Expenses'), 7).map(function (r) {
    return { date: dstr_(r[0]), breakfast: n_(r[1]), lunch: n_(r[2]), dinner: n_(r[3]),
             others: n_(r[4]), note: String(r[5] || ''), total: n_(r[6]) };
  }).filter(function (r) { return r.date >= from && r.date <= to; });
  return { ok: true, rows: rows };
}

function saveExpense_(req) {
  const date = str_(req.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Invalid date.');
  const b = n_(req.breakfast), l = n_(req.lunch), d = n_(req.dinner), o = n_(req.others);
  const row = [date, b, l, d, o, String(req.note || '').slice(0, 200), round2_(b + l + d + o)];

  return withLock_(function () {
    const sh = sheet_('Expenses');
    const last = sh.getLastRow();
    let target = 0;
    if (last >= 2) {
      const dates = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < dates.length; i++) {
        if (dstr_(dates[i][0]) === date) { target = i + 2; break; }
      }
    }
    if (!target) target = last + 1;
    sh.getRange(target, 1, 1, 7).setValues([row]);
    if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, 7).sort({ column: 1, ascending: true });
    return { ok: true };
  });
}

/* ------------------------------- Investments ------------------------------ */

function normInv_(r) { r[0] = dstr_(r[0]); return r; }

function getInvestments_(req) {
  const only = str_(req.date);
  const rows = readRows_(sheet_('Investments'), 4).map(normInv_)
    .filter(function (r) { return !only || r[0] === only; })
    .map(function (r) { return { date: r[0], aspect: String(r[1]), amount: n_(r[2]), grow: n_(r[3]) }; });
  return { ok: true, rows: rows };
}

function saveInvestments_(req) {
  const date = str_(req.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Invalid date.');
  const items = (req.rows || []).map(function (r) {
    return [date, String(r.aspect || '').slice(0, 100), n_(r.amount), n_(r.grow)];
  });
  return withLock_(function () {
    replaceRows_(sheet_('Investments'), 4, normInv_, function (r) { return r[0] !== date; }, items);
    return { ok: true };
  });
}

/* ---------------------------------- Todo ---------------------------------- */

function normTodo_(r) { r[0] = dstr_(r[0], 'yyyy-MM'); r[3] = dstr_(r[3]); return r; }

function getTodo_(req) {
  const month = str_(req.month);
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('Invalid month.');
  const rows = readRows_(sheet_('Todo'), 5).map(normTodo_)
    .filter(function (r) { return r[0] === month; })
    .map(function (r) {
      return { month: r[0], aspect: String(r[1]), amount: n_(r[2]), dateline: r[3],
               paid: r[4] === true || String(r[4]).toLowerCase() === 'true' };
    });
  return { ok: true, rows: rows };
}

function saveTodo_(req) {
  const month = str_(req.month);
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('Invalid month.');
  const items = (req.rows || []).map(function (r) {
    return [month, String(r.aspect || '').slice(0, 100), n_(r.amount), str_(r.dateline), r.paid === true];
  });
  return withLock_(function () {
    replaceRows_(sheet_('Todo'), 5, normTodo_, function (r) { return r[0] !== month; }, items);
    return { ok: true };
  });
}

/* --------------------------------- helpers -------------------------------- */

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const s = SCHEMA[name];
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, s.headers.length).setValues([s.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Store dates as plain text so Sheets never converts or shifts them.
    s.textCols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}

function readRows_(sh, ncols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ncols).getValues();
}

/** Replace all rows matching "not keepFn" with newRows (used to overwrite one date / month). */
function replaceRows_(sh, ncols, normFn, keepFn, newRows) {
  const kept = readRows_(sh, ncols).map(normFn).filter(keepFn);
  const all = kept.concat(newRows);
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, ncols).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, ncols).setValues(all);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function str_(v) { return v === undefined || v === null ? '' : String(v).trim(); }
function n_(v) { const n = Number(v); return isFinite(n) ? round2_(n) : 0; }
function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }
function dstr_(v, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), fmt || 'yyyy-MM-dd');
  return String(v === undefined || v === null ? '' : v).trim();
}
