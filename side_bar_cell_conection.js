function EMP_getSelectedEmployeeId() {
  // --- התאמה לטבלה שלך ---
  var SHEET_NAME = 'פרטי עובדים'; // אם שם הגיליון אחר – עדכן כאן
  var HEADER_ROW = 1;             // שורת הכותרת
  var COL_EMP_ID = 2;             // עמודת ID עובד (B=2)

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

  return {
    ok: true,
    employeeId: String(empId)
  };
}
