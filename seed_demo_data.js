/* global OPT, SHEET_NAME, SHEET_NAME_REQUESTS */
/* exported SEED_generateDemoDataForLastTwoMonths, SEED_clearAllDemoData */
// Deterministic demo data seeding (manual-only). Uses batch writes (setValues) to avoid execution timeouts.
// Scenarios distribution (pattern = (dayIndex*7 + empIndex) % 10):
// 0-4: Simple shift (≈50%) – no request, no fix (G/H empty)
// 5: Pending request – request points to shift, shift row untouched
// 6: Approved, not applied – request approved, shift row untouched
// 7: Approved + applied – request approved and G/H on shift row updated
// 8: Rejected – request rejected, shift row untouched
// 9: Manual fix – no request, G/H filled directly (manager-style)
// All seeded rows are tagged with SEED_TAG in notes/shiftNote; no randomness used.
// Manual undo: SEED_clearAllDemoData clears only demo-tagged rows via clearContent() (no deletes, no formatting changes).

// Shifts sheet columns (A–L) as written by handleShiftPost:
// A: הערות למשמרת (notes)
// B: ID משמרת (shiftId)
// C: חותמת זמן (timestamp)
// D: ID עובד (employeeId)
// E: שם מלא (employeeName)
// F: כניסה / יציאה (direction)
// G: תיקון תאריך (fixDate)
// H: תיקון שעה (fixTime)
// I: ID סוג עבודה (jobId)
// J: סוג עבודה (jobName)
// K: מחלקה (department)
// L: כמות היחידות (units)

// Requests sheet columns (resolved from headers in the sheet, mirroring handleRequestPost):
// Status, ID בקשה, הערות למשמרת, ID משמרת, חותמת זמן, ID עובד, שם מלא,
// כניסה / יציאה, תיקון תאריך, תיקון שעה, ID סוג עבודה/ID סוגי עבודה,
// סוג עבודה/סוגי עבודה/סוג העבודה, מחלקה, כמות היחידות (only columns present are filled).

var SEED_TAG = '[SEED_DEMO_M2]';
var SEED_PROGRESS_KEY = 'SEED_DEMO_LAST_TWO_MONTHS_PROGRESS';
var SEED_MAX_DAYS_PER_RUN = 7; // reduce if still close to timeout
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
  var props = PropertiesService.getScriptProperties();

  var today = new Date();
  var end = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  end.setHours(0, 0, 0, 0);

  var start = new Date(end.getTime());
  start.setMonth(start.getMonth() - 2);
  start.setHours(0, 0, 0, 0);

  var progressIso = props.getProperty(SEED_PROGRESS_KEY);
  var currentStart = progressIso ? new Date(progressIso) : new Date(start.getTime());
  if (isNaN(currentStart.getTime()) || currentStart < start) {
    currentStart = new Date(start.getTime());
  }

  if (currentStart > end) {
    props.deleteProperty(SEED_PROGRESS_KEY);
    Logger.log('SEED: last two months already fully seeded; nothing to do.');
    return;
  }

  var chunkEnd = new Date(currentStart.getTime());
  for (var i = 1; i < SEED_MAX_DAYS_PER_RUN; i++) {
    var d = new Date(chunkEnd.getTime());
    d.setDate(d.getDate() + 1);
    if (d > end) break;
    chunkEnd = d;
  }

  var rangeStart = new Date(currentStart.getFullYear(), currentStart.getMonth(), currentStart.getDate());
  var rangeEnd = new Date(chunkEnd.getFullYear(), chunkEnd.getMonth(), chunkEnd.getDate());

  var result = SEED_generateDemoDataForDateRange(rangeStart, rangeEnd);

  var nextStart = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth(), rangeEnd.getDate());
  nextStart.setDate(nextStart.getDate() + 1);
  nextStart.setHours(0, 0, 0, 0);

  if (nextStart > end) {
    props.deleteProperty(SEED_PROGRESS_KEY);
    Logger.log('SEED: completed last two months. Chunk ' + rangeStart.toDateString() + ' – ' + rangeEnd.toDateString() + '. Result: ' + JSON.stringify(result));
  } else {
    props.setProperty(SEED_PROGRESS_KEY, nextStart.toISOString());
    Logger.log('SEED: processed chunk ' + rangeStart.toDateString() + ' – ' + rangeEnd.toDateString() + '. Next start: ' + nextStart.toDateString() + '. Result: ' + JSON.stringify(result));
  }
}

function SEED_generateDemoDataForDateRange(startDateInput, endDateInput) {
  var startDate = normalizeDate_(startDateInput);
  var endDate = normalizeDate_(endDateInput);
  if (!startDate || !endDate || startDate > endDate) {
    throw new Error('Invalid date range for seeding.');
  }

  var shiftSheet = getShiftSheet_();
  var requestsSheet = getRequestsSheet_();
  var requestLayout = getRequestLayout_(requestsSheet);

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

  var ctx = {
    shiftRows: [],
    requestRows: [],
    shiftSeq: 1,
    requestSeq: 1,
    requestLayout: requestLayout
  };

  var dayIndex = 0;
  for (var d = new Date(startDate); d <= endDate; d = addDays_(d, 1)) {
    if (isWeekend_(d)) {
      dayIndex++;
      continue;
    }

    var dateStr = formatDate_(d);
    var workDate = new Date(d.getTime());
    for (var e = 0; e < selectedEmployees.length; e++) {
      var emp = selectedEmployees[e];
      var pattern = (dayIndex * 7 + e) % 10;

      var jobPrimary = jobs[e % jobs.length];

      applyScenario_(shiftSheet, requestsSheet, workDate, dateStr, dayIndex, e, emp, jobPrimary, pattern, ctx);
    }
    dayIndex++;
  }

  // Batch append shifts
  if (ctx.shiftRows.length > 0) {
    var startRowShifts = shiftSheet.getLastRow() + 1;
    shiftSheet.getRange(startRowShifts, 1, ctx.shiftRows.length, 12).setValues(ctx.shiftRows);
  }

  // Batch append requests (up to detected data columns to avoid touching formula columns to the right)
  if (ctx.requestRows.length > 0) {
    var dataCols = ctx.requestLayout.dataLastCol;
    var startRowReq = requestsSheet.getLastRow() + 1;
    requestsSheet.getRange(startRowReq, 1, ctx.requestRows.length, dataCols).setValues(ctx.requestRows);
  }

  return {
    createdShifts: ctx.shiftRows.length,
    createdRequests: ctx.requestRows.length
  };
}

// --- Scenario application ---
function applyScenario_(shiftSheet, requestsSheet, workDate, dateStr, dayIndex, empIndex, emp, jobA, pattern, ctx) {
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
      ctx: ctx,
      workDate: workDate,
      timestampTime: entryTime,
      employee: emp,
      job: jobA,
      direction: 'כניסה',
      fixDate: '',
      fixTime: '',
      units: shiftUnitsFull,
      scenarioLabel: 'SIMPLE'
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: exitTime,
      employee: emp,
      job: jobA,
      direction: 'יציאה',
      fixDate: '',
      fixTime: '',
      units: '',
      scenarioLabel: 'SIMPLE'
    });
  } else if (pattern === 5) {
    // Pending request, shift untouched (no fixes on row)
    var entryP = SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: entryTime,
      employee: emp,
      job: jobA,
      direction: 'כניסה',
      fixDate: '',
      fixTime: '',
      units: shiftUnitsFull,
      scenarioLabel: 'PENDING_TIME'
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: exitTime,
      employee: emp,
      job: jobA,
      direction: 'יציאה',
      fixDate: '',
      fixTime: '',
      units: '',
      scenarioLabel: 'PENDING_TIME'
    });
    SEED_newRequest_({
      ctx: ctx,
      employee: emp,
      job: jobA,
      status: 'pending',
      shiftId: entryP.shiftId,
      workDate: workDate,
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase + 10),
      shiftNoteLabel: 'PENDING_TIME',
      units: shiftUnitsFull,
      direction: 'כניסה',
      timestampTime: formatTimeFromMinutes_(entryMinutesBase + 10)
    });
  } else if (pattern === 6) {
    // Approved but NOT applied to shift row
    var entryA = SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: entryTime,
      employee: emp,
      job: jobA,
      direction: 'כניסה',
      fixDate: '',
      fixTime: '',
      units: shiftUnitsFull,
      scenarioLabel: 'APPROVED_NOT_APPLIED'
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: exitTime,
      employee: emp,
      job: jobA,
      direction: 'יציאה',
      fixDate: '',
      fixTime: '',
      units: '',
      scenarioLabel: 'APPROVED_NOT_APPLIED'
    });
    SEED_newRequest_({
      ctx: ctx,
      employee: emp,
      job: jobA,
      status: 'approved',
      shiftId: entryA.shiftId,
      workDate: workDate,
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase - 5),
      shiftNoteLabel: 'APPROVED_NOT_APPLIED',
      units: shiftUnitsFull,
      direction: 'כניסה',
      timestampTime: formatTimeFromMinutes_(entryMinutesBase - 5)
    });
  } else if (pattern === 7) {
    // Approved AND applied to shift row (G/H updated)
    var entryAA = SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: entryTime,
      employee: emp,
      job: jobA,
      direction: 'כניסה',
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase - 5),
      units: shiftUnitsFull,
      scenarioLabel: 'APPROVED_APPLIED'
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: exitTime,
      employee: emp,
      job: jobA,
      direction: 'יציאה',
      fixDate: '',
      fixTime: '',
      units: '',
      scenarioLabel: 'APPROVED_APPLIED'
    });
    var approvedFixTime = formatTimeFromMinutes_(entryMinutesBase - 5);
    SEED_newRequest_({
      ctx: ctx,
      employee: emp,
      job: jobA,
      status: 'approved',
      shiftId: entryAA.shiftId,
      workDate: workDate,
      fixDate: dateStr,
      fixTime: approvedFixTime,
      shiftNoteLabel: 'APPROVED_APPLIED',
      units: shiftUnitsFull,
      direction: 'כניסה',
      timestampTime: approvedFixTime
    });
  } else if (pattern === 8) {
    // Rejected request, shift untouched
    var entryR = SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: entryTime,
      employee: emp,
      job: jobA,
      direction: 'כניסה',
      fixDate: '',
      fixTime: '',
      units: shiftUnitsFull,
      scenarioLabel: 'REJECTED'
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: exitTime,
      employee: emp,
      job: jobA,
      direction: 'יציאה',
      fixDate: '',
      fixTime: '',
      units: '',
      scenarioLabel: 'REJECTED'
    });
    SEED_newRequest_({
      ctx: ctx,
      employee: emp,
      job: jobA,
      status: 'rejected',
      shiftId: entryR.shiftId,
      workDate: workDate,
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase + 15),
      shiftNoteLabel: 'REJECTED',
      units: shiftUnitsFull + 1,
      direction: 'כניסה',
      timestampTime: formatTimeFromMinutes_(entryMinutesBase + 15)
    });
  } else if (pattern === 9) {
    // Manual fix on shift row, no request
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: formatTimeFromMinutes_(entryMinutesBase - 10),
      employee: emp,
      job: jobA,
      direction: 'כניסה',
      fixDate: dateStr,
      fixTime: formatTimeFromMinutes_(entryMinutesBase - 10),
      units: shiftUnitsFull,
      scenarioLabel: 'MANUAL_FIX'
    });
    SEED_newShift_({
      shiftSheet: shiftSheet,
      ctx: ctx,
      workDate: workDate,
      timestampTime: exitTime,
      employee: emp,
      job: jobA,
      direction: 'יציאה',
      fixDate: '',
      fixTime: '',
      units: '',
      scenarioLabel: 'MANUAL_FIX'
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
  var shiftId = 'SEED_SHIFT_' + opts.ctx.shiftSeq++;

  var ts = buildTimestamp_(opts.workDate, opts.timestampTime);

  var unitsVal = '';
  if (opts.units !== null && opts.units !== undefined && String(opts.units).trim() !== '') {
    var n = Number(opts.units);
    unitsVal = isNaN(n) ? opts.units : n;
  }

  var row = [
    notes,
    shiftId,
    ts,
    opts.employee.id,
    opts.employee.name,
    opts.direction,
    opts.fixDate || '',
    opts.fixTime || '',
    opts.job.id,
    opts.job.name,
    opts.job.department,
    unitsVal
  ];

  opts.ctx.shiftRows.push(row);
  return { shiftId: shiftId };
}

function SEED_newRequest_(opts) {
  var layout = opts.ctx.requestLayout;
  var dataLastCol = layout.dataLastCol;
  var row = new Array(dataLastCol);
  for (var i = 0; i < dataLastCol; i++) row[i] = '';

  var requestId = 'SEED_REQ_' + opts.ctx.requestSeq++;
  var ts = buildTimestamp_(opts.workDate, opts.timestampTime || opts.fixTime || '12:00');
  var note = (opts.shiftNoteLabel ? (opts.shiftNoteLabel + ' ') : '') + SEED_TAG;
  var unitsVal = '';
  if (opts.units !== null && opts.units !== undefined && String(opts.units).trim() !== '') {
    var n = Number(opts.units);
    unitsVal = isNaN(n) ? opts.units : n;
  }

  if (layout.COL_STATUS) row[layout.COL_STATUS - 1] = opts.status;
  row[layout.COL_REQUEST_ID - 1] = requestId;
  if (layout.COL_SHIFT_NOTE) row[layout.COL_SHIFT_NOTE - 1] = note;
  if (layout.COL_SHIFT_ID) row[layout.COL_SHIFT_ID - 1] = opts.shiftId || '';
  if (layout.COL_TIMESTAMP) row[layout.COL_TIMESTAMP - 1] = ts;
  if (layout.COL_EMP_ID) row[layout.COL_EMP_ID - 1] = opts.employee.id;
  if (layout.COL_EMP_NAME) row[layout.COL_EMP_NAME - 1] = opts.employee.name;
  if (layout.COL_DIRECTION) row[layout.COL_DIRECTION - 1] = opts.direction || '';
  if (layout.COL_FIX_DATE) row[layout.COL_FIX_DATE - 1] = opts.fixDate || '';
  if (layout.COL_FIX_TIME) row[layout.COL_FIX_TIME - 1] = opts.fixTime || '';
  if (layout.COL_JOB_ID) row[layout.COL_JOB_ID - 1] = opts.job.id;
  if (layout.COL_JOB_NAME) row[layout.COL_JOB_NAME - 1] = opts.job.name;
  if (layout.COL_DEPT) row[layout.COL_DEPT - 1] = opts.job.department;
  if (layout.COL_UNITS) row[layout.COL_UNITS - 1] = unitsVal;

  opts.ctx.requestRows.push(row);
  return { requestId: requestId };
}

function hasSeedTagInShifts_(sheet, startDate, endDate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var notesRange = sheet.getRange(2, 1, lastRow - 1, 1);
  var finder = notesRange.createTextFinder(SEED_TAG);
  var match = finder ? finder.findNext() : null;
  while (match) {
    var r = match.getRow();
    var rowVals = sheet.getRange(r, 1, 1, 8).getValues()[0];
    var fixDate = asDate_(rowVals[6]); // G
    var tsDate = asDate_(rowVals[2]);  // C
    var dateToCheck = fixDate || tsDate;
    if (dateToCheck && dateToCheck >= startDate && dateToCheck <= endDate) {
      return true;
    }
    match = finder.findNext();
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

  var notesRange = sheet.getRange(2, colNote, lastRow - 1, 1);
  var finder = notesRange.createTextFinder(SEED_TAG);
  var match = finder ? finder.findNext() : null;
  while (match) {
    var r = match.getRow();
    var rowVals = sheet.getRange(r, 1, 1, Math.max(colFixDate, colTs || colFixDate)).getValues()[0];
    var fixDate = asDate_(rowVals[colFixDate - 1]);
    var tsDate = colTs ? asDate_(rowVals[colTs - 1]) : null;
    var dateToCheck = fixDate || tsDate;
    if (dateToCheck && dateToCheck >= startDate && dateToCheck <= endDate) {
      return true;
    }
    match = finder.findNext();
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

function getRequestLayout_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) lastCol = 1;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });

  function colByHeader_(name) {
    var idx = headers.indexOf(name);
    return idx >= 0 ? (idx + 1) : null; // 1-based
  }

  function colByAny_(names) {
    for (var i = 0; i < names.length; i++) {
      var c = colByHeader_(names[i]);
      if (c) return c;
    }
    return null;
  }

  var layout = {
    COL_STATUS:     colByAny_(['סטטוס בקשה']),
    COL_REQUEST_ID: colByAny_(['ID בקשה']),
    COL_SHIFT_NOTE: colByAny_(['הערות למשמרת']),
    COL_SHIFT_ID:   colByAny_(['ID משמרת']),
    COL_TIMESTAMP:  colByAny_(['חותמת זמן']),
    COL_EMP_ID:     colByAny_(['ID עובד']),
    COL_EMP_NAME:   colByAny_(['שם מלא']),
    COL_DIRECTION:  colByAny_(['כניסה / יציאה']),
    COL_FIX_DATE:   colByAny_(['תיקון תאריך']),
    COL_FIX_TIME:   colByAny_(['תיקון שעה']),
    COL_JOB_ID:     colByAny_(['ID סוג עבודה', 'ID סוגי עבודה']),
    COL_JOB_NAME:   colByAny_(['סוג עבודה', 'סוגי עבודה', 'סוג העבודה']),
    COL_DEPT:       colByAny_(['מחלקה']),
    COL_UNITS:      colByAny_(['כמות היחידות']),
    dataLastCol:    detectRequestDataLastCol_(headers, lastCol)
  };

  if (!layout.COL_REQUEST_ID || !layout.COL_TIMESTAMP) {
    throw new Error('Requests sheet is missing required headers (ID בקשה / חותמת זמן).');
  }

  return layout;
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

function buildTimestamp_(workDate, timeStr) {
  var parts = String(timeStr || '00:00').split(':');
  var hours = parts.length > 0 ? parseInt(parts[0], 10) || 0 : 0;
  var minutes = parts.length > 1 ? parseInt(parts[1], 10) || 0 : 0;
  return new Date(
    workDate.getFullYear(),
    workDate.getMonth(),
    workDate.getDate(),
    hours,
    minutes,
    0,
    0
  );
}
