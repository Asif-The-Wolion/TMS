/**
 * Teacher Planner — Google Sheets Backup endpoint
 * ================================================
 * Deploy this as a Web App (Deploy → New deployment → Web app → Execute as:
 * Me → Who has access: Anyone with the link). Paste the resulting URL into
 * the app's Settings → "Backup to Google Sheets".
 *
 * Access control is the Web App URL itself — without it, this endpoint is
 * unreachable, so there's no separate PIN check here. This script only ever
 * WRITES when the app pushes a backup (doPost), and only READS when asked
 * with ?action=restore (doGet). Nothing else in your Drive/Sheets is
 * touched — this script only acts on the ONE spreadsheet it's attached to.
 */

const BACKUP_SHEET_NAME = 'Backup (raw — do not edit)';
const PREVIOUS_SHEET_NAME = 'Backup previous (raw — do not edit)';
const CHUNK_SIZE = 40000; // Google Sheets caps a single cell around 50,000 chars

// OPTIONAL shared secret. Leave '' to keep working exactly as before (URL-only access).
// If you set one here, paste the SAME text into the app: Settings → Sync token. All requests
// without it are refused. Redeploy (Deploy → Manage deployments → Edit → New version) after changing.
const SYNC_TOKEN = '';

// Live sync keeps a revision number. A push is only accepted if it was built on the CURRENT
// revision; otherwise the app must pull + merge first. This is what stops a stale phone/PC
// from silently overwriting the other device's changes.
function getRev_() {
  var p = PropertiesService.getScriptProperties();
  var v = p.getProperty('REV');
  if (v === null) {
    // First time (or an existing pre-sync sheet): if a backup already exists it is revision 1,
    // so no device can "push over it" without pulling first.
    v = String(readBackupTab_() ? 1 : 0);
    p.setProperty('REV', v);
  }
  return Number(v) || 0;
}
function setRev_(n) { PropertiesService.getScriptProperties().setProperty('REV', String(n)); }
function tokenOk_(t) { return !SYNC_TOKEN || String(t || '') === SYNC_TOKEN; }

function doPost(e) {
  // One write at a time: two devices pushing together could otherwise interleave
  // clear + write and leave a corrupted (half old / half new) backup.
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(20000);
    locked = true;
    var body = JSON.parse(e.postData.contents);
    if (!tokenOk_(body.token)) return jsonOut_({ ok: false, error: 'Wrong or missing sync token.' });
    var data = body.data;
    var props = PropertiesService.getScriptProperties();
    var cur = getRev_();
    if (body.action === 'sync') {
      if (Number(body.baseRev) !== cur) return jsonOut_({ ok: false, conflict: true, rev: cur });
    } else if (props.getProperty('SYNC_MODE') === '1') {
      // An old build of the app (not yet reloaded) must not blind-overwrite a synced sheet.
      return jsonOut_({ ok: false, refused: true, error: 'Live sync is on for this sheet. Reload the app to get the new version, then try again.' });
    }
    var refusal = refuseUnsafeBackup_(data);
    if (refusal) return jsonOut_({ ok: false, refused: true, error: refusal });
    rotatePreviousBackup_();
    writeBackupTab_(JSON.stringify(data));
    writeReadableTabs_(data);
    var rev = cur + 1;
    setRev_(rev);
    if (body.action === 'sync') props.setProperty('SYNC_MODE', '1');
    return jsonOut_({ ok: true, rev: rev, timestamp: new Date().toISOString() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    if (locked) { try { lock.releaseLock(); } catch (ignore) {} }
  }
}

// Never let a bad push destroy a good backup.
function refuseUnsafeBackup_(data) {
  if (!data || typeof data !== 'object' || !Array.isArray(data.students)) {
    return 'The data sent did not look like a Teacher Planner backup, so the existing backup was kept.';
  }
  if (data.students.length === 0) {
    var existing = null;
    try { existing = JSON.parse(readBackupTab_() || 'null'); } catch (ignore) {}
    var had = existing && Array.isArray(existing.students) ? existing.students.length : 0;
    if (had > 0) {
      return 'Refused: this would replace a backup of ' + had + ' students with an EMPTY one (a new device or cleared browser?). Use Restore in the app first. If you really want to start over, clear the "' + BACKUP_SHEET_NAME + '" tab by hand.';
    }
  }
  return '';
}

// Keep one older copy of the backup (refreshed at most once a day) in its own tab,
// so a bad-but-not-empty push can still be recovered from.
function rotatePreviousBackup_() {
  var current = readBackupTab_();
  if (!current) return;
  var sheet = getSheet_(PREVIOUS_SHEET_NAME);
  var stamp = String(sheet.getRange(1, 1).getValue() || '');
  var m = /Snapshot taken: (\S+)/.exec(stamp);
  var last = m ? Date.parse(m[1]) : NaN;
  if (!isNaN(last) && (new Date().getTime() - last) < 24 * 3600 * 1000) return;
  sheet.clearContents();
  sheet.getRange(1, 1).setValue('Snapshot taken: ' + new Date().toISOString() + ' — older copy of the main backup, kept before it was overwritten. Do not edit.');
  var rows = [];
  for (var i = 0; i < current.length; i += CHUNK_SIZE) rows.push([current.slice(i, i + CHUNK_SIZE)]);
  if (rows.length) sheet.getRange(2, 1, rows.length, 1).setValues(rows);
}

function doGet(e) {
  try {
    var action = (e.parameter.action || '').toString();
    if (!tokenOk_(e.parameter.token)) return jsonOut_({ ok: false, error: 'Wrong or missing sync token.' });
    if (action === 'meta') {
      // Tiny and fast: just the revision, so devices can poll cheaply.
      return jsonOut_({ ok: true, rev: getRev_() });
    }
    if (action === 'restore') {
      // ?action=restore&which=previous  → the older copy kept by rotatePreviousBackup_
      // Read under the same lock as writes so a half-written backup is never returned.
      var lock = LockService.getScriptLock();
      var got = false;
      try { lock.waitLock(15000); got = true; } catch (ignore) {}
      try {
        var raw = readBackupTab_((e.parameter.which || '') === 'previous' ? PREVIOUS_SHEET_NAME : BACKUP_SHEET_NAME);
        if (!raw) return jsonOut_({ ok: false, error: 'No backup found yet — back up from the app at least once first.' });
        var data = JSON.parse(raw);
        return jsonOut_({ ok: true, data: data, rev: getRev_() });
      } finally {
        if (got) { try { lock.releaseLock(); } catch (ignore) {} }
      }
    }
    return jsonOut_({ ok: true, message: 'Teacher Planner backup endpoint is running.' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function writeBackupTab_(rawJsonString) {
  var sheet = getSheet_(BACKUP_SHEET_NAME);
  sheet.clearContents();
  sheet.getRange(1, 1).setValue('Last updated: ' + new Date().toISOString() + ' — raw data below, do not edit.');
  var chunks = [];
  for (var i = 0; i < rawJsonString.length; i += CHUNK_SIZE) {
    chunks.push(rawJsonString.slice(i, i + CHUNK_SIZE));
  }
  var rows = chunks.map(function (c) { return [c]; });
  if (rows.length) sheet.getRange(2, 1, rows.length, 1).setValues(rows);
}

function readBackupTab_(name) {
  var sheet = getSheet_(name || BACKUP_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var joined = values.map(function (r) { return r[0]; }).join('');
  return joined || null;
}

function writeReadableTabs_(data) {
  var students = data.students || [];
  writeStudentsTab_(students);
  writeFeesTab_(students, data.fees || []);
  writeAttendanceTab_(students, data.attendance || []);
  writeExamsTab_(students, data.exams || []);
  writeExpensesTab_(data.expenses || []);
  writeIncomeTab_(data.incomes || []);
  writeLoansTab_(data.loans || []);
  writeSavingsTab_(data.savingPots || [], data.savingTx || []);
}

// Money tabs: only created once there is something to show, but kept in sync
// (cleared) if they already exist and everything was deleted.
function moneySheet_(name, rowsCount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!rowsCount && !ss.getSheetByName(name)) return null;
  var sheet = getSheet_(name);
  sheet.clearContents();
  return sheet;
}
function putTable_(sheet, header, rows) {
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}
function byDateDesc_(a, b) { return String(b.date || '').localeCompare(String(a.date || '')); }

function writeExpensesTab_(expenses) {
  var sheet = moneySheet_('Expenses', expenses.length);
  if (!sheet) return;
  putTable_(sheet, ['Date', 'Type', 'Category', 'Amount', 'Note'], expenses.slice().sort(byDateDesc_).map(function (e) {
    return [e.date || '', e.kind === 'business' ? 'Teaching' : 'Personal', e.category || '', Number(e.amount) || 0, e.note || ''];
  }));
}
function writeIncomeTab_(incomes) {
  var sheet = moneySheet_('Income', incomes.length);
  if (!sheet) return;
  putTable_(sheet, ['Date', 'Source', 'Amount', 'Note'], incomes.slice().sort(byDateDesc_).map(function (i) {
    return [i.date || '', i.source || '', Number(i.amount) || 0, i.note || ''];
  }));
}
function writeLoansTab_(loans) {
  var sheet = moneySheet_('Loans', loans.length);
  if (!sheet) return;
  putTable_(sheet, ['Person', 'Direction', 'Amount', 'Date', 'Due', 'Repaid', 'Outstanding', 'Note'], loans.slice().sort(byDateDesc_).map(function (l) {
    var paid = (l.repayments || []).reduce(function (a, r) { return a + (Number(r.amount) || 0); }, 0);
    var amt = Number(l.amount) || 0;
    return [l.person || '', l.direction === 'borrowed' ? 'I borrowed' : 'I lent', amt, l.date || '', l.dueDate || '', paid, Math.max(0, amt - paid), l.note || ''];
  }));
}
function writeSavingsTab_(pots, tx) {
  var sheet = moneySheet_('Savings', tx.length + pots.length);
  if (!sheet) return;
  var names = {};
  pots.forEach(function (p) { names[p.id] = p.name; });
  putTable_(sheet, ['Pot', 'Date', 'Type', 'Amount', 'Note'], tx.slice().sort(byDateDesc_).map(function (t) {
    return [names[t.potId] || '(deleted pot)', t.date || '', t.opening ? 'Opening balance' : (t.type === 'out' ? 'Withdraw' : 'Deposit'), Number(t.amount) || 0, t.note || ''];
  }));
}

function studentName_(students, id) {
  for (var i = 0; i < students.length; i++) {
    if (students[i].id === id) return students[i].name;
  }
  return '(unknown)';
}

function writeStudentsTab_(students) {
  var sheet = getSheet_('Students');
  sheet.clearContents();
  var header = ['Name', 'Type', 'Class', 'School', 'Subjects', 'Monthly Fee', 'Guardian', 'Phone', 'Area', 'Active'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  var rows = students.map(function (s) {
    return [s.name, s.type, s.class || '', s.school || '', (s.subjects || []).join(', '), s.monthlyFee || 0, s.guardianName || '', s.guardianPhone || '', s.area || '', s.active !== false ? 'Yes' : 'No'];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function writeFeesTab_(students, fees) {
  var sheet = getSheet_('Fees');
  sheet.clearContents();
  var header = ['Student', 'Month', 'Monthly Fee', 'Paid', 'Payment Date', 'Method', 'Notes'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  var rows = fees.map(function (f) {
    return [studentName_(students, f.studentId), f.month, f.monthlyFee || 0, f.paid || 0, f.paymentDate || '', f.method || '', f.notes || ''];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function writeAttendanceTab_(students, attendance) {
  var sheet = getSheet_('Attendance');
  sheet.clearContents();
  var header = ['Student', 'Date', 'Status'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  var rows = attendance.map(function (a) {
    return [studentName_(students, a.studentId), a.date, a.status];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function writeExamsTab_(students, exams) {
  var sheet = getSheet_('Exams');
  sheet.clearContents();
  var header = ['Student', 'Exam', 'Subject', 'Date', 'Status', 'Obtained', 'Total', 'Percent'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  var rows = exams.map(function (e) {
    var pct = e.total > 0 ? Math.round((e.obtained / e.total) * 100) : '';
    return [studentName_(students, e.studentId), e.name, e.subject, e.date, e.status || 'Completed', e.obtained, e.total, pct];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}