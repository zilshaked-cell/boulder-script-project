function handleShiftPost(e) {
  var lock = LockService.getDocumentLock();
  lock.tryLock(30000);

  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    if (!raw) return jsonResponse({ success: false, error: 'Missing body' });

    var data;
    try {
      data = JSON.parse(raw);
    } catch (err) {
      return jsonResponse({ success: false, error: 'Invalid JSON: ' + err });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ success: false, error: 'Sheet not found: ' + SHEET_NAME });

    function norm_(v) {
      if (v === null || v === undefined) return '';
      return String(v).replace(/\s+/g, ' ').trim();
    }
    function normKey_(v) { return norm_(v).toLowerCase(); }

    // --- קלט מהווב-אפ (תומך גם בשמות ישנים) ---
    var employeeId   = norm_(data.employeeId);
    var employeeName = norm_(data.employeeName);

    var direction = norm_(data.direction); // "כניסה"/"יציאה"
    var fixDate   = norm_(data.fixDate || data.date);
    var fixTime   = norm_(data.fixTime || data.time);

    var jobId   = norm_(data.jobId);
    var jobName = norm_(data.jobName);

    var units = data.units;
    var notes = norm_(data.notes) || 'kind=shiftReport';

    // --- רזולב עובד לפי "פרטי עובדים" אם חסר ID/שם ---
    function resolveEmployee_(empId, empName) {
      var out = { id: norm_(empId), name: norm_(empName) };

      if (!out.id && !out.name) return out;

      var shEmp = ss.getSheetByName('פרטי עובדים');
      if (!shEmp) return out;

      // נוודא IDs לעובדים קיימים לפני חיפוש
      if (typeof EMP !== 'undefined' && EMP.ensureEmployeeIds) {
        try { EMP.ensureEmployeeIds(); } catch (e) {}
      }

      var headerRow = 1;
      var lastRow = shEmp.getLastRow();
      var lastCol = shEmp.getLastColumn();
      if (lastRow <= headerRow) return out;

      var headers = shEmp.getRange(headerRow, 1, 1, lastCol).getValues()[0];
      function colIndexByHeader_(h) {
        var idx = headers.indexOf(h);
        return idx >= 0 ? (idx + 1) : null;
      }

      var colEmpId = colIndexByHeader_('ID עובד') || 2;     // fallback B
      var colName  = colIndexByHeader_('שם מלא') || 3;      // fallback C

      var values = shEmp.getRange(headerRow + 1, 1, lastRow - headerRow, Math.max(colEmpId, colName)).getValues();

      var byName = {};
      var byId = {};
      for (var i = 0; i < values.length; i++) {
        var rid = norm_(values[i][colEmpId - 1]);
        var rnm = norm_(values[i][colName - 1]);
        if (rid && rnm) {
          if (!byName[normKey_(rnm)]) byName[normKey_(rnm)] = rid;
          if (!byId[rid]) byId[rid] = rnm;
        }
      }

      if (!out.id && out.name) out.id = byName[normKey_(out.name)] || '';
      if (!out.name && out.id) out.name = byId[out.id] || '';

      return out;
    }

    var empResolved = resolveEmployee_(employeeId, employeeName);
    employeeId = empResolved.id;
    employeeName = empResolved.name;

    // --- רזולב סוג עבודה מהאופציות (כדי למלא ID+מחלקה) ---
    var department = norm_(data.department);

    function resolveJob_(jid, jname) {
      var out = { id: norm_(jid), name: norm_(jname), department: '' };

      if (typeof OPT === 'undefined') return out;

      // אם יש שם – הכי טוב: resolve לפי שם
      if (out.name && OPT.getJobByName) {
        var jr = OPT.getJobByName(out.name);
        if (jr) {
          out.id = out.id || norm_(jr.id);
          out.name = norm_(jr.name) || out.name;
          out.department = norm_(jr.department);
          return out;
        }
      }

      // אם יש רק ID – ננסה לסרוק רשימת עבודות
      if (out.id && OPT.getAllJobs) {
        var list = OPT.getAllJobs(false) || [];
        for (var i = 0; i < list.length; i++) {
          if (norm_(list[i].id) === out.id) {
            out.name = out.name || norm_(list[i].name);
            out.department = norm_(list[i].department);
            return out;
          }
        }
      }

      return out;
    }

    var jobResolved = resolveJob_(jobId, jobName);
    jobId = jobResolved.id;
    jobName = jobResolved.name;
    department = department || jobResolved.department;

    // --- ולידציות מינימום ---
    if (!employeeName && !employeeId) {
      return jsonResponse({ success: false, error: 'Missing employeeId/employeeName' });
    }
    if (!direction) {
      return jsonResponse({ success: false, error: 'Missing direction' });
    }
    if (!fixDate || !fixTime) {
      return jsonResponse({ success: false, error: 'Missing fixDate/fixTime' });
    }
    if (!jobName && !jobId) {
      return jsonResponse({ success: false, error: 'Missing jobId/jobName' });
    }

    // --- סכימת "דיווח שעות עבודה" לפי ה-contract (A..L) ---
    // A הערות
    // B ID משמרת
    // C חותמת זמן
    // D ID עובד
    // E שם מלא
    // F כניסה / יציאה
    // G תיקון תאריך
    // H תיקון שעה
    // I ID סוג עבודה
    // J סוג עבודה
    // K מחלקה
    // L יחידות

    var shiftId = Utilities.getUuid();
    var now = new Date();

    var unitsVal = '';
    if (units !== null && units !== undefined && norm_(units) !== '') {
      var n = Number(units);
      unitsVal = isNaN(n) ? norm_(units) : n;
    }

    var row = [
      notes,
      shiftId,
      now,
      employeeId,
      employeeName,
      direction,
      fixDate,
      fixTime,
      jobId,
      jobName,
      department,
      unitsVal
    ];

    sheet.appendRow(row);

    return jsonResponse({ success: true, shiftId: shiftId });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
