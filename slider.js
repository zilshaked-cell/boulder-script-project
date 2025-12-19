/**
 * בודק אם יש לעובד שורות מוסתרות בגליון "פרטי עובדים"
 * לפי הערך בעמודה B ("ID עובד").
 *
 * מחזיר true אם לפחות שורה אחת של העובד מוסתרת ע"י פילטר או הסתרה ידנית.
 */
function EMP_hasHiddenRowsForEmployee(employeeId) {
  if (!employeeId) return false;

  // 1) אם מודול EMP קיים – זה המקור המועדף
  try {
    if (typeof EMP !== 'undefined' && EMP.hasHiddenRowsForEmployee) {
      return !!EMP.hasHiddenRowsForEmployee(employeeId);
    }
  } catch (err) {
    // נופלים לפתרון המקומי
    Logger.log('EMP.hasHiddenRowsForEmployee failed, fallback: ' + err);
  }

  // 2) Fallback מקומי – תומך גם במבנה עם employee_id וגם במבנה עם ID עובד בעמודה B
  var SHEET_NAME = 'פרטי עובדים';
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return false;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol < 1) return false;

  var target = String(employeeId).trim();

  // קוראים כותרות
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return String(v || '').trim();
  });

  // עדיפויות למציאת עמודת מזהה
  // 1) employee_id (מבנה טכני ישן)
  // 2) ID עובד (המבנה הנוכחי שלך)
  // 3) fallback לעמודה B
  var idCol = -1;

  var idxEmployeeId = headers.indexOf('employee_id');
  if (idxEmployeeId !== -1) {
    idCol = idxEmployeeId + 1; // 1-based
  } else {
    var idxHebId = headers.indexOf('ID עובד');
    if (idxHebId !== -1) idCol = idxHebId + 1;
  }

  if (idCol === -1) idCol = 2; // B

  // לקריאה יעילה – קוראים רק עמודה אחת (idCol)
  var numRows = lastRow - 1;
  var idValues = sheet.getRange(2, idCol, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var val = String(idValues[i][0] || '').trim();
    if (!val) continue;
    if (val !== target) continue;

    var rowIndex = 2 + i;

    var hiddenByFilter = false;
    var hiddenByUser = false;

    if (sheet.isRowHiddenByFilter) hiddenByFilter = sheet.isRowHiddenByFilter(rowIndex);
    if (sheet.isRowHiddenByUser) hiddenByUser = sheet.isRowHiddenByUser(rowIndex);

    if (hiddenByFilter || hiddenByUser) return true;
  }

  return false;
}


/**
 * מבטל פילטר רגיל על הכרטיסייה "פרטי עובדים" כדי שכל השורות יראו.
 * לא משנה שום ערך בטבלה, רק את מצב התצוגה.
 */
function EMP_revealAllRowsIfFiltered() {
  // 1) מודול EMP (מועדף)
  try {
    if (typeof EMP !== 'undefined' && EMP.revealAllRowsIfFiltered) {
      return EMP.revealAllRowsIfFiltered();
    }
  } catch (err) {
    Logger.log('EMP.revealAllRowsIfFiltered failed, fallback: ' + err);
  }

  // 2) fallback
  var SHEET_NAME = 'פרטי עובדים';
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { ok: false, error: 'sheet not found' };

  var filter = sheet.getFilter();
  if (filter) {
    filter.remove();
    return { ok: true, changed: true };
  }
  return { ok: true, changed: false };
}
