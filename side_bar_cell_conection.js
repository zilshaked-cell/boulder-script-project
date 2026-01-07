function EMP_getSelectedEmployeeId() {
  var logger = null;
  var startMs = new Date().getTime();
  try {
    if (typeof ensureModuleLoggerDefined_ === "function") {
      logger = ensureModuleLoggerDefined_("SIDEBAR_CELL_SELECT");
    }
  } catch (_ignored) {}

  // --- התאמה לטבלה שלך ---
  var SHEET_NAME = "פרטי עובדים"; // אם שם הגיליון אחר – עדכן כאן
  var HEADER_ROW = 1; // שורת הכותרת
  var COL_EMP_ID = 2; // עמודת ID עובד (B=2)

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== SHEET_NAME) {
    return { ok: true, employeeId: null };
  }

  var cell = sh.getActiveCell();
  if (!cell) {
    return { ok: true, employeeId: null };
  }

  var row = cell.getRow();
  if (row <= HEADER_ROW) {
    return { ok: true, employeeId: null };
  }

  var empId = sh.getRange(row, COL_EMP_ID).getValue();
  if (!empId) {
    return { ok: true, employeeId: null };
  }

  var res = {
    ok: true,
    employeeId: String(empId),
  };

  if (typeof logDuration_ === "function") {
    logDuration_(logger, "sidebar.cell.getEmployeeId", startMs, {
      employeeId: res.employeeId || null,
    });
  }

  return res;
}
