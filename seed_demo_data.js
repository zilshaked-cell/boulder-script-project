// Deterministic demo data seeding (manual-only). Uses existing handlers and schemas.
// Scenarios distribution (pattern = (dayIndex*7 + empIndex) % 10):
// 0-4: Simple shift (≈50%) – no request, no fix (G/H cleared)
// 5: Pending request – request points to shift, shift row untouched
// 6: Approved, not applied – request approved, shift row untouched
// 7: Approved + applied – request approved and G/H on shift row updated
// 8: Rejected – request rejected, shift row untouched
// 9: Manual fix – no request, G/H filled directly (manager-style)
// All seeded rows are tagged with SEED_TAG in notes/shiftNote; no randomness used.
// Manual undo: SEED_clearAllDemoData clears only demo-tagged rows via clearContent() (no deletes, no formatting changes).

var SEED_TAG = '[SEED_DEMO_M2]';
var REQUEST_DATA_HEADER_CANDIDATES = [
  'סטטוס בקשה',
  'ID בקשה',
  'הערות למשמרת',
  'ID משמרת',
  'חותמת זמן',
  'ID עובד',
  'שם מלא',
  'כניסה / יציאה',
  'תיקון תאריך',
  'תיקון שעה',
  'ID סוג עבודה',
  'ID סוגי עבודה',
  'סוג עבודה',
  'סוגי עבודה',
  'סוג העבודה',
  'מחלקה',
  'כמות היחידות'
];

function SEED_generateDemoDataForLastTwoMonths() {
  var today = new Date();
  var end = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var start = new Date(end);
  start.setMonth(start.getMonth() - 2);
  SEED_generateDemoDataForDateRange(start, end);
}

function SEED_generateDemoDataForDateRange(startDateInput, endDateInput) {
  var startDate = normalizeDate_(startDateInput);
  var endDate = normalizeDate_(endDateInput);
  if (!startDate || !endDate || startDate > endDate) {
    throw new Error('Invalid date range for seeding.');
  }

  var shiftSheet = getShiftSheet_();
  var requestsSheet = getRequestsSheet_();

  var employees = loadEmployees_();
  if (employees.length === 0) throw new Error('No employees found to seed.');
  var jobs = loadJobs_();
  if (jobs.length === 0) throw new Error('No jobs found to seed.');

  var selectedEmployees = employees.slice(0, 10);

  // Abort if seed data already exists in date range.
  if (hasSeedTagInShifts_(shiftSheet, startDate, endDate)) {
    throw new Error('Seed data already exists in shift sheet for this date range (' + SEED_TAG + '); aborting.');
  }
  if (hasSeedTagInRequests_(requestsSheet, startDate, endDate)) {
    throw new Error('Seed data already exists in requests sheet for this date range (' + SEED_TAG + '); aborting.');
  }

  var dayIndex = 0;
  for (var d = new Date(startDate); d <= endDate; d = addDays_(d, 1)) {
    if (isWeekend_(d)) {
      dayIndex++;
      continue;
    }

    var dateStr = formatDate_(d);
    for (var e = 0; e < selectedEmployees.length; e++) {
      var emp = selectedEmployees[e];
      var pattern = (dayIndex * 7 + e) % 10;

      var jobPrimary = jobs[e % jobs.length];
      var jobSecondary = jobs[(e + 1) % jobs.length];

      applyScenario_(shiftSheet, requestsSheet, dateStr, dayIndex, e, emp, jobPrimary, jobSecondary, pattern);
    }
    dayIndex++;
  }
}

// --- Scenario application ---
function applyScenario_(shiftSheet, requestsSheet, dateStr, dayIndex, empIndex, emp, jobA, jobB, pattern) {
  var entryMinutesBase = 8 * 60 + (empIndex % 3) * 10; // 08:00, 08:10, 08:20
  var exitMinutesBase = 16 * 60 + (empIndex % 2) * 10; // 16:00 or 16:10

  var entryTime = formatTimeFromMinutes_(entryMinutesBase);
  var exitTime = formatTimeFromMinutes_(exitMinutesBase);

  var shiftUnitsFull = 8;

  // Scenario mapping (biased):
  // 0-4 simple, 5 pending, 6 approved-not-applied, 7 approved+applied, 8 rejected, 9 manual fix.
  if (pattern <= 4) {
    // Simple shift: clear G/H after insert, no request
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'כניסה',
      fixTime: entryTime,
      units: shiftUnitsFull,
      scenarioLabel: 'SIMPLE',
      clearFixAfterInsert: true
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'יציאה',
      fixTime: exitTime,
      units: '',
      scenarioLabel: 'SIMPLE',
      clearFixAfterInsert: true
    });
  } else if (pattern === 5) {
    // Pending request, shift untouched (no fixes on row)
    var entryP = SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'כניסה',
      fixTime: entryTime,
      units: shiftUnitsFull,
      scenarioLabel: 'PENDING_TIME',
      clearFixAfterInsert: true
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'יציאה',
      fixTime: exitTime,
      units: '',
      scenarioLabel: 'PENDING_TIME',
      clearFixAfterInsert: true
    });
    SEED_newRequest_({
      requestsSheet: requestsSheet,
      employee: emp,
      job: jobA,
      status: 'pending',
      shiftId: entryP.shiftId,
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase + 10),
      shiftNoteLabel: 'PENDING_TIME',
      units: shiftUnitsFull
    });
  } else if (pattern === 6) {
    // Approved but NOT applied to shift row
    var entryA = SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'כניסה',
      fixTime: entryTime,
      units: shiftUnitsFull,
      scenarioLabel: 'APPROVED_NOT_APPLIED',
      clearFixAfterInsert: true
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'יציאה',
      fixTime: exitTime,
      units: '',
      scenarioLabel: 'APPROVED_NOT_APPLIED',
      clearFixAfterInsert: true
    });
    SEED_newRequest_({
      requestsSheet: requestsSheet,
      employee: emp,
      job: jobA,
      status: 'approved',
      shiftId: entryA.shiftId,
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase - 5),
      shiftNoteLabel: 'APPROVED_NOT_APPLIED',
      units: shiftUnitsFull
    });
  } else if (pattern === 7) {
    // Approved AND applied to shift row (G/H updated)
    var entryAA = SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'כניסה',
      fixTime: entryTime,
      units: shiftUnitsFull,
      scenarioLabel: 'APPROVED_APPLIED',
      clearFixAfterInsert: true
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'יציאה',
      fixTime: exitTime,
      units: '',
      scenarioLabel: 'APPROVED_APPLIED',
      clearFixAfterInsert: true
    });
    var approvedFixTime = formatTimeFromMinutes_(entryMinutesBase - 5);
    SEED_newRequest_({
      requestsSheet: requestsSheet,
      employee: emp,
      job: jobA,
      status: 'approved',
      shiftId: entryAA.shiftId,
      fixDate: dateStr,
      fixTime: approvedFixTime,
      shiftNoteLabel: 'APPROVED_APPLIED',
      units: shiftUnitsFull
    });
    SEED_applyFixToShift_(shiftSheet, entryAA.rowIndex, dateStr, approvedFixTime, 'APPROVED_APPLIED');
  } else if (pattern === 8) {
    // Rejected request, shift untouched
    var entryR = SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'כניסה',
      fixTime: entryTime,
      units: shiftUnitsFull,
      scenarioLabel: 'REJECTED',
      clearFixAfterInsert: true
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'יציאה',
      fixTime: exitTime,
      units: '',
      scenarioLabel: 'REJECTED',
      clearFixAfterInsert: true
    });
    SEED_newRequest_({
      requestsSheet: requestsSheet,
      employee: emp,
      job: jobA,
      status: 'rejected',
      shiftId: entryR.shiftId,
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase + 15),
      shiftNoteLabel: 'REJECTED',
      units: shiftUnitsFull + 1
    });
  } else if (pattern === 9) {
    // Manual fix on shift row, no request
    var entryM = SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'כניסה',
      fixTime: entryTime,
      units: shiftUnitsFull,
      scenarioLabel: 'MANUAL_FIX',
      clearFixAfterInsert: true,
      applyFixOverride: {
        fixDate: dateStr,
        fixTime: formatTimeFromMinutes_(entryMinutesBase - 10)
      }
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      employee: emp,
      job: jobA,
      dateStr: dateStr,
      direction: 'יציאה',
      fixTime: exitTime,
      units: '',
      scenarioLabel: 'MANUAL_FIX',
      clearFixAfterInsert: true
    });
    // No request created here by design.
  }
}

// --- Data loaders and helpers ---
function loadEmployees_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('פרטי עובדים');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  function colByName(name) {
    var idx = headers.indexOf(name);
    return idx >= 0 ? idx + 1 : null;
  }
  var colId = colByName('ID עובד');
  var colName = colByName('שם מלא');
  var colDept = colByName('מחלקה');
  if (!colId || !colName) return [];
  var vals = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var id = (vals[i][colId - 1] || '').toString().trim();
    var name = (vals[i][colName - 1] || '').toString().trim();
    var dept = colDept ? (vals[i][colDept - 1] || '').toString().trim() : '';
    if (id && name) {
      out.push({ id: id, name: name, department: dept });
    }
  }
  out.sort(function (a, b) {
    return a.id.localeCompare(b.id) || a.name.localeCompare(b.name);
  });
  return out;
}

function loadJobs_() {
  if (typeof OPT !== 'undefined' && OPT.getAllJobs) {
    var jobs = OPT.getAllJobs(false) || [];
    var filtered = jobs
      .map(function (j) {
        return {
          id: (j.id || '').toString().trim(),
          name: (j.name || '').toString().trim(),
          department: (j.department || '').toString().trim()
        };
      })
      .filter(function (j) { return j.id && j.name; });
    filtered.sort(function (a, b) { return a.id.localeCompare(b.id) || a.name.localeCompare(b.name); });
    return filtered;
  }
  throw new Error('OPT.getAllJobs is not available; cannot seed without jobs.');
}

function SEED_newShift_(opts) {
  var notes = (opts.scenarioLabel ? (opts.scenarioLabel + ' ') : '') + SEED_TAG;
  var payload = {
    employeeId: opts.employee.id,
    employeeName: opts.employee.name,
    direction: opts.direction,
    fixDate: opts.dateStr,
    fixTime: opts.fixTime,
    jobId: opts.job.id,
    jobName: opts.job.name,
    department: opts.job.department,
    notes: notes,
    units: opts.units
  };

  var res = handleShiftPost({ postData: { contents: JSON.stringify(payload) } });
  var shiftId = '';
  try {
    var parsed = JSON.parse(res && res.getContent ? res.getContent() : '{}');
    if (parsed && parsed.shiftId) shiftId = parsed.shiftId;
  } catch (err) {}

  var rowIndex = findShiftRowById_(opts.shiftSheet, shiftId);
  if (!rowIndex) {
    throw new Error('Could not locate shift row for shiftId=' + shiftId);
  }

  if (opts.clearFixAfterInsert) {
    opts.shiftSheet.getRange(rowIndex, 7, 1, 2).clearContent(); // G,H
  }

  if (opts.applyFixOverride) {
    // Only touch row we just wrote and that carries SEED_TAG in notes.
    var noteVal = opts.shiftSheet.getRange(rowIndex, 1).getValue() || '';
    if (noteVal.toString().indexOf(SEED_TAG) >= 0) {
      opts.shiftSheet.getRange(rowIndex, 7, 1, 2).setValues([[opts.applyFixOverride.fixDate, opts.applyFixOverride.fixTime]]);
    }
  }

  return { shiftId: shiftId, rowIndex: rowIndex };
}

function SEED_applyFixToShift_(shiftSheet, rowIndex, fixDate, fixTime, label) {
  var noteVal = shiftSheet.getRange(rowIndex, 1).getValue() || '';
  if (noteVal.toString().indexOf(SEED_TAG) < 0) return; // safety
  shiftSheet.getRange(rowIndex, 7, 1, 2).setValues([[fixDate, fixTime]]);
  // Optionally reinforce note label
  var newNote = (label ? (label + ' ') : '') + SEED_TAG;
  shiftSheet.getRange(rowIndex, 1).setValue(newNote);
}

function SEED_newRequest_(opts) {
  var note = (opts.shiftNoteLabel ? (opts.shiftNoteLabel + ' ') : '') + SEED_TAG;
  var payload = {
    employeeId: opts.employee.id,
    employeeName: opts.employee.name,
    status: opts.status,
    shiftId: opts.shiftId || '',
    direction: opts.direction || '',
    fixDate: opts.fixDate,
    fixTime: opts.fixTime,
    jobId: opts.job.id,
    jobName: opts.job.name,
    department: opts.job.department,
    shiftNote: note,
    units: opts.units
  };

  var res = handleRequestPost({ postData: { contents: JSON.stringify(payload) } });
  var requestId = '';
  var rowIndex = null;
  try {
    var parsed = JSON.parse(res && res.getContent ? res.getContent() : '{}');
    requestId = parsed && parsed.requestId ? parsed.requestId : '';
    rowIndex = parsed && parsed.rowIndex ? parsed.rowIndex : null;
  } catch (err) {}
  return { requestId: requestId, rowIndex: rowIndex };
}

function findShiftRowById_(sheet, shiftId) {
  if (!shiftId) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var rng = sheet.getRange(2, 2, lastRow - 1, 1);
  var found = rng.createTextFinder(String(shiftId)).matchEntireCell(true).findNext();
  return found ? found.getRow() : null;
}

function hasSeedTagInShifts_(sheet, startDate, endDate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(8, lastCol)).getValues();
  for (var i = 0; i < data.length; i++) {
    var notes = (data[i][0] || '').toString();
    if (notes.indexOf(SEED_TAG) < 0) continue;
    var fixDate = asDate_(data[i][6]); // column G
    var tsDate = asDate_(data[i][2]);  // column C timestamp
    var dateToCheck = fixDate || tsDate;
    if (!dateToCheck) continue;
    if (dateToCheck >= startDate && dateToCheck <= endDate) {
      return true;
    }
  }
  return false;
}

function hasSeedTagInRequests_(sheet, startDate, endDate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  function colByName(name) {
    var idx = headers.indexOf(name);
    return idx >= 0 ? idx + 1 : null;
  }
  var colNote = colByName('הערות למשמרת');
  var colFixDate = colByName('תיקון תאריך');
  var colTs = colByName('חותמת זמן');
  if (!colNote || !colFixDate) return false;
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var note = (data[i][colNote - 1] || '').toString();
    if (note.indexOf(SEED_TAG) < 0) continue;
    var fixDate = asDate_(data[i][colFixDate - 1]);
    var tsDate = colTs ? asDate_(data[i][colTs - 1]) : null;
    var dateToCheck = fixDate || tsDate;
    if (!dateToCheck) continue;
    if (dateToCheck >= startDate && dateToCheck <= endDate) {
      return true;
    }
  }
  return false;
}

// --- Demo data undo (clears only demo-tagged rows; no deletes, no formatting changes) ---
function SEED_clearAllDemoData() {
  var clearedShifts = SEED_clearDemoShifts_();
  var clearedRequests = SEED_clearDemoRequests_();
  Logger.log('SEED_clearAllDemoData: cleared ' + clearedShifts + ' shift rows and ' + clearedRequests + ' request rows.');
  return {
    clearedShifts: clearedShifts,
    clearedRequests: clearedRequests
  };
}

function SEED_clearDemoShifts_() {
  if (typeof SHEET_NAME === 'undefined') return 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var cleared = 0;
  var dataCols = Math.min(12, lastCol);

  for (var i = 0; i < values.length; i++) {
    var notes = String(values[i][0] || ''); // column A
    if (notes.indexOf(SEED_TAG) === -1) continue;
    var rowIndex = i + 2;
    sheet.getRange(rowIndex, 1, 1, dataCols).clearContent();
    cleared++;
  }

  return cleared;
}

function SEED_clearDemoRequests_() {
  if (typeof SHEET_NAME_REQUESTS === 'undefined') return 0;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var noteCol = detectNoteCol_(headers);
  var dataLastCol = detectRequestDataLastCol_(headers, lastCol);

  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var cleared = 0;

  for (var i = 0; i < values.length; i++) {
    var isDemo = false;
    if (noteCol) {
      var noteVal = String(values[i][noteCol - 1] || '');
      isDemo = noteVal.indexOf(SEED_TAG) !== -1;
    } else {
      for (var c = 0; c < lastCol; c++) {
        var cellVal = String(values[i][c] || '');
        if (cellVal.indexOf(SEED_TAG) !== -1) { isDemo = true; break; }
      }
    }
    if (!isDemo) continue;

    var rowIndex = i + 2;
    sheet.getRange(rowIndex, 1, 1, dataLastCol).clearContent();
    cleared++;
  }

  return cleared;
}

function detectNoteCol_(headers) {
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    if (!h) continue;
    if (h.indexOf('הער') !== -1 || /note/i.test(h)) {
      return c + 1; // 1-based
    }
  }
  return null;
}

function detectRequestDataLastCol_(headers, lastCol) {
  var maxIdx = 0;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim();
    for (var j = 0; j < REQUEST_DATA_HEADER_CANDIDATES.length; j++) {
      if (h === REQUEST_DATA_HEADER_CANDIDATES[j]) {
        if (c + 1 > maxIdx) maxIdx = c + 1;
        break;
      }
    }
  }
  if (!maxIdx) maxIdx = lastCol;
  return maxIdx;
}

function getShiftSheet_() {
  if (typeof SHEET_NAME === 'undefined') {
    throw new Error('SHEET_NAME is not defined.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Shift sheet not found: ' + SHEET_NAME);
  return sh;
}

function getRequestsSheet_() {
  if (typeof SHEET_NAME_REQUESTS === 'undefined') {
    throw new Error('SHEET_NAME_REQUESTS is not defined.');
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME_REQUESTS);
  if (!sh) throw new Error('Requests sheet not found: ' + SHEET_NAME_REQUESTS);
  return sh;
}

// --- Generic helpers ---
function normalizeDate_(v) {
  if (v instanceof Date) return new Date(v.getTime());
  if (typeof v === 'string') {
    var d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function formatDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone && Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd');
}

function formatTimeFromMinutes_(mins) {
  var h = Math.floor(mins / 60);
  var m = mins % 60;
  return (h < 10 ? '0' + h : h) + ':' + (m < 10 ? '0' + m : m);
}

function addDays_(d, n) {
  return new Date(d.getTime() + n * 24 * 60 * 60 * 1000);
}

function asDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var d = new Date(v);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function isWeekend_(d) {
  var day = d.getDay();
  return day === 5 || day === 6; // Fri/Sat
}
