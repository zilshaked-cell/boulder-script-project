// Manual-only test harness that calls existing web-app handlers without changing them.
// All TEST_* functions are run from the Apps Script editor (no triggers/endpoints).
// Undo only touches rows whose IDs match what was logged.

var TEST_WEBAPP_LOG_SHEET_NAME = 'TEST_WEBAPP_LOG';
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

function ensureTestLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEST_WEBAPP_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TEST_WEBAPP_LOG_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([[
      'batchId',
      'timestamp',
      'entity',
      'sheetName',
      'rowIndex',
      'idValue',
      'mode',
      'status'
    ]]);
  }

  return sheet;
}

function getNextBatchId_() {
  var sheet = ensureTestLogSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxId = 0;
  for (var i = 0; i < vals.length; i++) {
    var n = parseInt(vals[i][0], 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  return maxId + 1;
}

function TEST_insertSampleShiftsFromWebApp() {
  var batchId = getNextBatchId_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName(typeof SHEET_NAME !== 'undefined' ? SHEET_NAME : '');
  if (!shiftSheet) {
    throw new Error('Shift sheet not found; expected SHEET_NAME to point to the shift report sheet.');
  }

  var sampleShifts = [
    {
      employeeId: 'EMP_TEST_001',
      employeeName: 'Test Employee One',
      direction: 'כניסה',
      fixDate: '2024-01-01',
      fixTime: '08:00',
      jobId: 'JOB_TEST',
      jobName: 'Job Test',
      department: 'Dept Test',
      notes: 'TEST BATCH entry',
      units: 1
    },
    {
      employeeId: 'EMP_TEST_001',
      employeeName: 'Test Employee One',
      direction: 'יציאה',
      fixDate: '2024-01-01',
      fixTime: '16:00',
      jobId: 'JOB_TEST',
      jobName: 'Job Test',
      department: 'Dept Test',
      notes: 'TEST BATCH exit',
      units: 1
    },
    {
      employeeId: 'EMP_TEST_002',
      employeeName: 'Test Employee Two',
      direction: 'כניסה',
      fixDate: '2024-01-02',
      fixTime: '09:00',
      jobId: 'JOB_ALT',
      jobName: 'Alt Job',
      department: 'Alt Dept',
      notes: 'TEST BATCH alt',
      units: 2
    }
  ];

  var logSheet = ensureTestLogSheet_();
  var lastRowBefore = shiftSheet.getLastRow();

  for (var i = 0; i < sampleShifts.length; i++) {
    var payload = sampleShifts[i];
    var event = { postData: { contents: JSON.stringify(payload) } };

    try {
      handleShiftPost(event);
    } catch (err) {
      // If production handler throws, skip logging for this payload.
      continue;
    }

    var lastRowAfter = shiftSheet.getLastRow();
    if (lastRowAfter <= lastRowBefore) {
      continue;
    }

    var lastCol = shiftSheet.getLastColumn();
    var rowValues = shiftSheet.getRange(lastRowAfter, 1, 1, lastCol).getValues()[0];
    var idValue = rowValues[1]; // Column B (shiftId)

    logSheet.appendRow([
      batchId,
      new Date(),
      'shift',
      shiftSheet.getName(),
      lastRowAfter,
      idValue,
      'inserted',
      'added'
    ]);

    lastRowBefore = lastRowAfter;
  }
}

function TEST_insertSampleRequestsFromWebApp() {
  var batchId = getNextBatchId_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requestsSheet = ss.getSheetByName(typeof SHEET_NAME_REQUESTS !== 'undefined' ? SHEET_NAME_REQUESTS : '');
  if (!requestsSheet) {
    throw new Error('Requests sheet not found; expected SHEET_NAME_REQUESTS to be defined.');
  }

  var sampleRequests = [
    {
      employeeId: 'EMP_TEST_003',
      employeeName: 'Test Employee Three',
      direction: 'כניסה',
      fixDate: '2024-01-03',
      fixTime: '07:30',
      jobId: 'JOB_REQ',
      jobName: 'Req Job',
      department: 'Req Dept',
      status: 'pending',
      units: 1
    },
    {
      employeeId: 'EMP_TEST_004',
      employeeName: 'Test Employee Four',
      direction: 'יציאה',
      fixDate: '2024-01-03',
      fixTime: '17:00',
      jobId: 'JOB_REQ',
      jobName: 'Req Job',
      department: 'Req Dept',
      status: 'pending',
      units: 0.5
    }
  ];

  var logSheet = ensureTestLogSheet_();

  for (var i = 0; i < sampleRequests.length; i++) {
    var payload = sampleRequests[i];
    var event = { postData: { contents: JSON.stringify(payload) } };
    var response;

    try {
      response = handleRequestPost(event);
    } catch (err) {
      continue;
    }

    if (!response || typeof response.getContent !== 'function') {
      continue;
    }

    var content;
    try {
      content = response.getContent();
    } catch (err2) {
      continue;
    }

    var parsed;
    try {
      parsed = JSON.parse(content);
    } catch (err3) {
      continue;
    }

    if (!parsed || !parsed.requestId || !parsed.rowIndex) {
      continue;
    }

    logSheet.appendRow([
      batchId,
      new Date(),
      'request',
      requestsSheet.getName(),
      parsed.rowIndex,
      parsed.requestId,
      parsed.mode || 'inserted',
      'added'
    ]);
  }
}

function TEST_undoLastTestBatch() {
  var logSheet = ensureTestLogSheet_();
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) return; // no data

  var rows = logSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var addedRows = [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][7] || '').trim() === 'added') {
      addedRows.push({ idx: i, row: rows[i] });
    }
  }
  if (addedRows.length === 0) return;

  var latestBatchId = 0;
  for (var j = 0; j < addedRows.length; j++) {
    var b = parseInt(addedRows[j].row[0], 10);
    if (!isNaN(b) && b > latestBatchId) latestBatchId = b;
  }
  if (!latestBatchId) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var statuses = rows.map(function (r) { return r[7]; });

  function findRequestIdCol_(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return null;
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c] || '').trim() === 'ID בקשה') return c + 1; // 1-based
    }
    return null;
  }

  function getRequestDataLastCol_(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) return lastCol;
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var maxIdx = 0;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '').trim();
      for (var j = 0; j < REQUEST_DATA_HEADER_CANDIDATES.length; j++) {
        if (h === REQUEST_DATA_HEADER_CANDIDATES[j]) {
          if (i + 1 > maxIdx) maxIdx = i + 1; // 1-based
          break;
        }
      }
    }
    return maxIdx || lastCol;
  }

  var targetRows = addedRows.filter(function (ar) {
    var b = parseInt(ar.row[0], 10);
    return b === latestBatchId;
  });

  for (var k = 0; k < targetRows.length; k++) {
    var entry = targetRows[k];
    var entryRow = entry.row;
    var logIndex = entry.idx;

    var entity = String(entryRow[2] || '').trim();
    var sheetName = String(entryRow[3] || '').trim();
    var rowIndex = parseInt(entryRow[4], 10);
    var idValue = String(entryRow[5] || '').trim();

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      statuses[logIndex] = 'skipped_sheet_missing';
      continue;
    }

    if (!rowIndex || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      statuses[logIndex] = 'skipped_row_oob';
      continue;
    }

    var lastCol = sheet.getLastColumn();
    var rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

    var idCol;
    if (entity === 'shift') {
      idCol = 2; // column B
    } else if (entity === 'request') {
      idCol = findRequestIdCol_(sheet);
    }

    if (!idCol) {
      statuses[logIndex] = 'skipped_id_col_missing';
      continue;
    }

    var currentId = String(rowValues[idCol - 1] || '').trim();
    if (currentId !== idValue) {
      statuses[logIndex] = 'skipped_id_mismatch';
      continue;
    }

    var dataCols = entity === 'shift'
      ? Math.min(12, lastCol)
      : Math.min(getRequestDataLastCol_(sheet), lastCol);
    sheet.getRange(rowIndex, 1, 1, dataCols).clearContent();
    statuses[logIndex] = 'undone';
  }

  logSheet.getRange(2, 8, statuses.length, 1).setValues(statuses.map(function (s) { return [s]; }));
}
