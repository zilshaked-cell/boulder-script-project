// employees_api.gs

/**
 * API לקריאת רשימת עובדים מהגיליון "פרטי עובדים".
 * לא נוגע במודול EMP ולא מחליף אותו – רק קורא מהטבלה.
 */

const EMPLOYEE_SHEET_NAME = "פרטי עובדים";
const EMPLOYEE_HEADER_ROW = 1;

// eslint-disable-next-line no-unused-vars
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
    params.onlyActive || params.activeOnly || "",
  ).toLowerCase();
  var activeOnly =
    activeOnlyRaw === "true" ||
    activeOnlyRaw === "1" ||
    activeOnlyRaw === "yes";

  var dataRange = sheet.getRange(
    EMPLOYEE_HEADER_ROW + 1,
    1,
    lastRow - EMPLOYEE_HEADER_ROW,
    lastCol,
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
      String(b.employeeName || ""),
    );
  });

  return jsonResponse({
    success: true,
    employees: employees,
  });
}

/**
 * בדיקת קיום עובד לפי מייל (לשימוש הלוגין של הווב-אפ).
 * מחזיר success+found, ואם נמצא – מצרף active+employeeId+employeeName.
 */
// eslint-disable-next-line no-unused-vars
function employeeExistsByEmail(params) {
  params = params || {};

  function norm(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim().toLowerCase();
  }

  var email = norm(params.email || params.mail || "");
  if (!email) {
    return jsonResponse({ success: false, error: "missing email" });
  }

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
    return jsonResponse({ success: true, found: false });
  }

  var lastCol = sheet.getLastColumn();
  var headerRow = sheet
    .getRange(EMPLOYEE_HEADER_ROW, 1, 1, lastCol)
    .getValues()[0];

  function buildHeaderIndex(headers) {
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var key = norm(headers[i]);
      if (key) map[key] = i + 1; // 1-based index
    }
    return map;
  }

  var H = buildHeaderIndex(headerRow);

  var requiredHeaders = ["מייל", "סטטוס", "ID עובד", "שם מלא"];
  var missing = [];
  requiredHeaders.forEach(function (colName) {
    if (!H[norm(colName)]) missing.push(colName);
  });

  if (missing.length) {
    return jsonResponse({
      success: false,
      error: "schema_mismatch",
      message:
        "employeeExistsByEmail: missing required columns in '" +
        EMPLOYEE_SHEET_NAME +
        "': " +
        missing.join(", "),
    });
  }

  var colEmail = H[norm("מייל")];
  var colStatus = H[norm("סטטוס")];
  var colId = H[norm("ID עובד")];
  var colName = H[norm("שם מלא")];
  var colSystemRole = H[norm("System Role")] || 0; // optional

  var dataRange = sheet.getRange(
    EMPLOYEE_HEADER_ROW + 1,
    1,
    lastRow - EMPLOYEE_HEADER_ROW,
    lastCol,
  );
  var values = dataRange.getValues();

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var rowEmail = norm(row[colEmail - 1]);
    if (!rowEmail) continue;
    if (rowEmail !== email) continue;

    var statusVal = colStatus ? String(row[colStatus - 1] || "").trim() : "";

    var statusLower = statusVal.toLowerCase();
    var inactive =
      statusLower.indexOf("לא פעיל") !== -1 ||
      statusLower.indexOf("inactive") !== -1;
    var active = !inactive;
    var empId = colId ? String(row[colId - 1] || "").trim() : "";
    var empName = colName ? String(row[colName - 1] || "").trim() : "";

    var rawRole = colSystemRole
      ? String(row[colSystemRole - 1] || "")
          .trim()
          .toLowerCase()
      : "";
    var systemRole = rawRole === "admin" ? "ADMIN" : "EMPLOYEE";

    return jsonResponse({
      success: true,
      found: true,
      employee: {
        id: empId,
        name: empName,
        email: rowEmail,
        active: active,
        systemRole: systemRole,
      },
    });
  }

  return jsonResponse({ success: true, found: false });
}
