function handleGetShifts(params) {
  var logger = null;
  var startMs = new Date().getTime();
  try {
    if (typeof ensureModuleLoggerDefined_ === "function") {
      logger = ensureModuleLoggerDefined_("SHIFT_GET");
    }
  } catch (_ignoredLogger) {}

  function finish_(res, errorCode, err) {
    if (typeof logDuration_ === "function") {
      logDuration_(
        logger,
        "shift.get",
        startMs,
        {
          success: !!(res && res.success),
          returned:
            res && res.shifts && Array.isArray(res.shifts)
              ? res.shifts.length
              : null,
          error: res && res.error,
        },
        res && res.success ? "info" : "warn",
        errorCode,
        err
      );
    }
    return res;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return finish_(
      jsonResponse({ success: false, error: "Sheet not found: " + SHEET_NAME }),
      "SHEET_NOT_FOUND"
    );
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return finish_(
      jsonResponse({
        success: true,
        shifts: [],
        items: [],
      })
    );
  }

  function norm_(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  // לפי סכימה A..L
  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  var employeeIdFilter = norm_(params && (params.employeeId || params.empId));
  var employeeNameFilter = norm_(
    params && (params.employeeName || params.employee)
  );

  var out = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];

    var rec = {
      notes: r[0], // A
      shiftId: r[1], // B
      timestamp: r[2], // C
      employeeId: r[3], // D
      employeeName: r[4], // E
      direction: r[5], // F
      fixDate: r[6], // G
      fixTime: r[7], // H
      jobId: r[8], // I
      jobName: r[9], // J
      department: r[10], // K
      units: r[11], // L
    };

    // סינון (תומך גם בישן: employeeName)
    if (employeeIdFilter) {
      if (norm_(rec.employeeId) !== employeeIdFilter) continue;
    } else if (employeeNameFilter) {
      if (norm_(rec.employeeName) !== employeeNameFilter) continue;
    }

    // דילוג על שורות ריקות לגמרי
    if (!rec.shiftId && !rec.employeeId && !rec.employeeName) continue;

    out.push(rec);
  }

  return finish_(
    jsonResponse({
      success: true,
      shifts: out,
      items: out,
    })
  );
}
