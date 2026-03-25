/* LEGACY SHADOW FILE
 * This file is kept for visibility and comparison only.
 * Canonical active counterpart: boulder-script-project/employees_api.js
 * Do not treat this file as the primary API implementation path.
 */
// employees_api.gs

/**
 * API לקריאת רשימת עובדים מהגיליון "פרטי עובדים".
 * לא נוגע במודול EMP ולא מחליף אותו – רק קורא מהטבלה.
 */

const EMPLOYEE_SHEET_NAME = "פרטי עובדים";
const EMPLOYEE_HEADER_ROW = 1;

function handleGetEmployees(params) {
  params = params || {};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET_NAME);
  if (!sheet) {
    return jsonResponse({
      success: false,
      error: "Sheet not found: " + EMPLOYEE_SHEET_NAME,
    });
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= EMPLOYEE_HEADER_ROW) {
    return jsonResponse({ success: true, employees: [] });
  }

  var lastCol = sheet.getLastColumn();
  var headerRow = sheet
    .getRange(EMPLOYEE_HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  function norm(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  function buildHeaderIndex(headers) {
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var key = norm(headers[i]);
      if (key) map[key] = i + 1; // 1-based index
    }
    return map;
  }

  var H = buildHeaderIndex(headerRow);

  var colStatus = H["סטטוס"] || 0;
  var colId = H["ID עובד"] || 0;
  var colName = H["שם מלא"] || 0;
  var colJobName = H["סוג עבודה"] || 0;
  var colJobId = H["ID סוג עבודה"] || 0;
  var colDept = H["מחלקה"] || 0;

  var employeeIdFilter =
    norm(params.employeeId || params.employee_id || params.id || "") || "";
  var employeeNameFilter =
    norm(params.employeeName || params.employee_name || params.name || "") ||
    "";

  var activeOnlyRaw = String(
    params.onlyActive || params.activeOnly || ""
  ).toLowerCase();
  var activeOnly =
    activeOnlyRaw === "true" ||
    activeOnlyRaw === "1" ||
    activeOnlyRaw === "yes";

  var dataRange = sheet.getRange(
    EMPLOYEE_HEADER_ROW + 1,
    1,
    lastRow - EMPLOYEE_HEADER_ROW,
    lastCol
  );
  var values = dataRange.getValues();

  // אוספים עובדים לפי ID (אם אין – לפי שם)
  var mapByKey = {};

  for (var r = 0; r < values.length; r++) {
    var row = values[r];

    var name = colName ? norm(row[colName - 1]) : "";
    if (!name) continue; // אין טעם בשורה בלי שם

    var id = colId ? norm(row[colId - 1]) : "";
    var statusVal = colStatus ? norm(row[colStatus - 1]) : "";
    var jobName = colJobName ? norm(row[colJobName - 1]) : "";
    var jobId = colJobId ? norm(row[colJobId - 1]) : "";
    var dept = colDept ? norm(row[colDept - 1]) : "";

    // פילטרים – קודם ID, אחר כך שם
    if (employeeIdFilter) {
      if (!id || norm(id) !== employeeIdFilter) continue;
    } else if (employeeNameFilter) {
      if (norm(name) !== employeeNameFilter) continue;
    }

    var rowActive = true;
    if (statusVal === "לא פעיל") rowActive = false;

    var key = id || name;
    if (!mapByKey[key]) {
      mapByKey[key] = {
        employeeId: id,
        employeeName: name,
        status: statusVal || "",
        active: rowActive,
        jobs: [],
        departments: [],
      };
    } else {
      // אם כבר יש עובד במפה – לעדכן "active" אם אחת השורות שלו פעילה
      if (rowActive) mapByKey[key].active = true;
      if (!mapByKey[key].status && statusVal) {
        mapByKey[key].status = statusVal;
      }
    }

    var emp = mapByKey[key];

    // הוספת עבודה (אם קיימת) – בלי כפילויות
    if (jobName || jobId) {
      var jobKey = jobId + "|" + jobName;
      var exists = false;
      for (var j = 0; j < emp.jobs.length; j++) {
        var jItem = emp.jobs[j];
        var jKey = (jItem.jobId || "") + "|" + (jItem.jobName || "");
        if (jKey === jobKey) {
          exists = true;
          break;
        }
      }
      if (!exists) {
        emp.jobs.push({
          jobId: jobId || "",
          jobName: jobName || "",
          department: dept || "",
        });
      }
    }

    // הוספת מחלקה (אם קיימת) – בלי כפילויות
    if (dept) {
      if (emp.departments.indexOf(dept) === -1) {
        emp.departments.push(dept);
      }
    }
  }

  var employees = Object.keys(mapByKey).map(function (key) {
    return mapByKey[key];
  });

  if (activeOnly) {
    employees = employees.filter(function (emp) {
      return emp.active;
    });
  }

  // מיון לפי שם
  employees.sort(function (a, b) {
    return String(a.employeeName || "").localeCompare(
      String(b.employeeName || "")
    );
  });

  return jsonResponse({
    success: true,
    employees: employees,
  });
}

/**
 * נקודת API לבדיקה האם מייל קיים בגיליון "פרטי עובדים" והאם הוא פעיל.
 * מחזירה shape מתאים ל-web app:
 *   { success: true, found: boolean, employee?: { id, name, email, status, active } }
 */
function handleEmployeeExistsByEmail(params) {
  params = params || {};

  function norm(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  var email = norm(params.email || params.mail || "").toLowerCase();
  if (!email) {
    return jsonResponse({ success: false, error: "missing email" });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EMPLOYEE_SHEET_NAME);
  if (!sheet) {
    return jsonResponse({ success: false, error: "sheet not found" });
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= EMPLOYEE_HEADER_ROW) {
    return jsonResponse({ success: true, found: false, employee: null });
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet
    .getRange(EMPLOYEE_HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  function findHeader(headersArr, candidates) {
    for (var i = 0; i < headersArr.length; i++) {
      var h = norm(headersArr[i]).toLowerCase();
      if (candidates.indexOf(h) !== -1) return i + 1; // 1-based
    }
    return 0;
  }

  var colEmail = findHeader(headers, [
    "מייל",
    "אימייל",
    "email",
    'דוא"ל',
    "דוא" + "ל",
  ]);
  var colStatus = findHeader(headers, ["סטטוס", "status"]);
  var colId = findHeader(headers, ["id עובד", "id", "employee id"]);
  var colName = findHeader(headers, ["שם מלא", "שם", "full name", "name"]);

  if (!colEmail) {
    return jsonResponse({ success: false, error: "email column not found" });
  }

  var data = sheet
    .getRange(
      EMPLOYEE_HEADER_ROW + 1,
      1,
      lastRow - EMPLOYEE_HEADER_ROW,
      lastCol
    )
    .getValues();

  var found = null;

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var rowEmailRaw = row[colEmail - 1];
    var rowEmail = norm(rowEmailRaw).toLowerCase();
    if (!rowEmail) continue;

    if (rowEmail === email) {
      var statusVal = colStatus ? norm(row[colStatus - 1]) : "";
      var active = true;
      var statusLower = statusVal.toLowerCase();
      if (statusLower === "לא פעיל" || statusLower === "inactive") {
        active = false;
      }

      found = {
        id: colId ? norm(row[colId - 1]) : "",
        name: colName ? norm(row[colName - 1]) : "",
        email: rowEmail,
        status: statusVal,
        active: active,
      };
      break;
    }
  }

  return jsonResponse({
    success: true,
    found: Boolean(found),
    employee: found,
  });
}
