/* global SpreadsheetApp, Logger, OPT, REQ, SCH, OPT_onOpen, OPT_onEdit, REQ_onEdit, SCH_onChange, EMP_checkEmployeesHeader_, getEmployeeById_, saveEmployeePayload_ */
/* eslint-disable no-inner-declarations, no-prototype-builtins */
var EMP = EMP || {};

(function () {
  var CONFIG = {
    SHEET_NAME_EMPLOYEES: "פרטי עובדים",
    SHEET_NAME_LOG: "LOG_EMPLOYEES",
    HEADER_ROW: 1,
    COL: {
      ACTIVE: 1, // A – סטטוס
      ID: 2, // B – ID עובד
      FULL_NAME: 3, // C – שם מלא
      JOB_TYPE: 11, // K – סוג העבודה
      DEPARTMENT: 12, // L – מחלקה
      AMOUNT: 13, // M – סכום
      PAYMENT_MODE: 16, // P – אופן תשלום
      NOTES: 17, // Q – הערות
    },
    PROPS: {
      SCHEMA: "EMP_SCHEMA_SNAPSHOT",
    },
    MAIL: {
      TO: "ZIL.SHAKED@GMAIL.COM",
      SUBJECT_PREFIX: "שכר בולדר חיפה – שינוי מבנה עובדים",
    },
  };

  function debugLog_(msg) {
    try {
      Logger.log("[EMP_DEBUG] " + msg);
    } catch (_e) {
      // no-op in environments without Logger
    }
  }

  function colIndexByHeader_(headers, name) {
    if (!headers || !headers.length || !name) return null;

    var idx = headers.indexOf(name);
    if (idx >= 0) return idx + 1;

    var target = String(name).trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (h === name) continue;
      if (
        String(h || "")
          .trim()
          .toLowerCase() === target
      ) {
        return i + 1;
      }
    }
    return null;
  }

  /**
   * Returns column indices for core employee/job/pay fields or an error.
   * Prefers live headers; falls back to CONFIG.COL when a header is not found.
   * If any required column is missing, no indices are returned to avoid corrupt writes.
   * @return {{ ok: true, cols: { employeeIdCol: number, jobTypeIdCol: number, jobTypeNameCol: number, departmentCol: number, payTypeIdCol: number, payTypeNameCol: number } } | { ok: false, error: string }}
   * @private
   */
  function EMP_getEmployeeColumns_() {
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      return {
        ok: false,
        error: 'לא נמצאה כרטיסייה "' + CONFIG.SHEET_NAME_EMPLOYEES + '"',
      };
    }

    var headerRow = CONFIG.HEADER_ROW;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    function resolve_(headerName, fallbackIdx) {
      var idx = colIndexByHeader_(headers, headerName);
      if (idx) return idx;
      if (fallbackIdx && typeof fallbackIdx === "number") return fallbackIdx;
      return null;
    }

    var employeeIdCol = resolve_("ID עובד", CONFIG.COL.ID);
    var jobTypeIdCol = resolve_("ID סוגי עבודה", null);
    var jobTypeNameCol = resolve_("סוג העבודה", CONFIG.COL.JOB_TYPE);
    var departmentCol = resolve_("מחלקה", CONFIG.COL.DEPARTMENT);
    var payTypeIdCol = resolve_("ID אופן תשלום", null);
    var payTypeNameCol = resolve_("אופן תשלום", CONFIG.COL.PAYMENT_MODE);

    var missing = [];
    if (!employeeIdCol) missing.push("ID עובד");
    if (!jobTypeNameCol) missing.push("סוג העבודה");
    if (!departmentCol) missing.push("מחלקה");
    if (!payTypeNameCol) missing.push("אופן תשלום");
    // ID columns are required for safe writes; if absent we prefer to skip backfill rather than write to a wrong column.
    if (!jobTypeIdCol) missing.push("ID סוגי עבודה");
    if (!payTypeIdCol) missing.push("ID אופן תשלום");

    if (missing.length) {
      return {
        ok: false,
        error: "Missing required header(s): " + missing.join(", "),
      };
    }

    return {
      ok: true,
      cols: {
        employeeIdCol: employeeIdCol,
        jobTypeIdCol: jobTypeIdCol,
        jobTypeNameCol: jobTypeNameCol,
        departmentCol: departmentCol,
        payTypeIdCol: payTypeIdCol,
        payTypeNameCol: payTypeNameCol,
      },
    };
  }

  function getSpreadsheet_() {
    return SpreadsheetApp.getActive();
  }

  function getEmployeesSheet_() {
    var ss = getSpreadsheet_();
    return ss.getSheetByName(CONFIG.SHEET_NAME_EMPLOYEES);
  }

  function getLogSheet_() {
    var ss = getSpreadsheet_();
    var sh = ss.getSheetByName(CONFIG.SHEET_NAME_LOG);
    if (!sh) {
      sh = ss.insertSheet(CONFIG.SHEET_NAME_LOG);
      sh.appendRow([
        "timestamp",
        "tx_id",
        "employee_id",
        "sheet",
        "row",
        "col",
        "old_value",
        "new_value",
        "user",
      ]);
    }
    return sh;
  }

  function stringValue_(val) {
    if (val === null || val === undefined) return "";
    return String(val).trim();
  }

  function validationError_(msg) {
    var e = new Error(msg);
    e.code = "VALIDATION_ERROR";
    return e;
  }

  function parseJobTypeIdsCell_(cellVal) {
    var raw = stringValue_(cellVal);
    if (!raw) return [];
    // Split on comma/semicolon/whitespace; keep simple to match existing free-text cells.
    var parts = raw
      .split(/[,;\s]+/)
      .map(function (p) {
        return stringValue_(p);
      })
      .filter(function (p) {
        return !!p;
      });
    var out = [];
    for (var i = 0; i < parts.length; i++) {
      if (out.indexOf(parts[i]) === -1) out.push(parts[i]);
    }
    return out;
  }

  function serializeJobTypeIds_(ids) {
    if (!Array.isArray(ids) || !ids.length) return "";
    var uniq = [];
    ids.forEach(function (id) {
      var v = stringValue_(id);
      if (!v) return;
      if (uniq.indexOf(v) === -1) uniq.push(v);
    });
    return uniq.join(",");
  }

  function normalizeSystemRole_(val) {
    var v = stringValue_(val).toLowerCase();
    if (!v) return "";
    if (v === "admin" || v === "administrator") return "ADMIN";
    if (v === "shift_manager") return "SHIFT_MANAGER";
    if (v === "viewer") return "VIEWER";
    return "EMPLOYEE";
  }

  function getScriptProps_() {
    return PropertiesService.getDocumentProperties();
  }

  function generateUuid_() {
    return Utilities.getUuid();
  }

  /**
   * קובע אם שורה נחשבת "פעילה" לפי הערך בעמודת סטטוס.
   * תומך גם בטקסט ("פעיל"/"לא פעיל") וגם בבוליאן true/false.
   */
  function isRowActiveFlag_(val) {
    if (val === true) return true;
    if (val === false) return false;

    var s = String(val || "").trim();
    if (!s) return false;
    if (s.toUpperCase() === "TRUE") return true;
    if (s === "פעיל") return true;
    if (s === "לא פעיל") return false;

    // ברירת מחדל: אם יש ערך כלשהו בסטטוס והוא לא "לא פעיל" – נחשב פעיל
    return true;
  }

  function normalizeSpace_(val) {
    if (val === null || val === undefined) return "";
    return String(val).replace(/\s+/g, " ").trim();
  }

  function normalizeKey_(val) {
    var v = normalizeSpace_(val);
    return v ? v.toLowerCase() : "";
  }

  function getOptionsSheet_() {
    var ss = getSpreadsheet_();
    var candidates = ["אופציות בחירה ו ID'S"];
    for (var i = 0; i < candidates.length; i++) {
      var sh = ss.getSheetByName(candidates[i]);
      if (sh) return sh;
    }
    return null;
  }

  function resolveOptionJobColumns_() {
    var sh = getOptionsSheet_();
    if (!sh) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "אופציות בחירה ו ID\'S"' };
    }

    var headerRow = 1;
    var lastCol = sh.getLastColumn();
    var headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    function resolve_(names, fallbackIdx) {
      for (var n = 0; n < names.length; n++) {
        var idx = colIndexByHeader_(headers, names[n]);
        if (idx) return idx;
      }
      if (fallbackIdx && typeof fallbackIdx === "number") return fallbackIdx;
      return null;
    }

    var jobNameCol = resolve_(["סוגי עבודה", "סוג העבודה", "סוג עבודה"], null);
    var jobIdCol = resolve_(["ID סוגי עבודה", "ID סוג עבודה"], null);

    if (!jobNameCol || !jobIdCol) {
      return {
        ok: false,
        error: "Missing required option headers (job name / job ID)",
      };
    }

    return {
      ok: true,
      sheet: sh,
      headerRow: headerRow,
      jobNameCol: jobNameCol,
      jobIdCol: jobIdCol,
    };
  }

  function buildJobNameToIdMap_() {
    if (typeof OPT !== "undefined" && OPT.ensureCatalogIds) {
      try {
        OPT.ensureCatalogIds();
      } catch (_ignored) {}
    }

    var resolved = resolveOptionJobColumns_();
    if (!resolved.ok) return { ok: false, error: resolved.error };

    var sh = resolved.sheet;
    var headerRow = resolved.headerRow;
    var lastRow = sh.getLastRow();
    if (lastRow <= headerRow) return { ok: true, map: {} };

    var lastCol = sh.getLastColumn();
    var values = sh
      .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
      .getValues();

    var map = {};
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var name = normalizeKey_(row[resolved.jobNameCol - 1]);
      var id = normalizeSpace_(row[resolved.jobIdCol - 1]);
      if (!name || !id) continue;
      if (!map.hasOwnProperty(name)) {
        map[name] = id;
      }
    }

    return { ok: true, map: map };
  }

  function EMP_fillJobIdForRow_(sheet, rowIndex, cols, jobLookup) {
    var sh = sheet || getEmployeesSheet_();
    if (!sh) return false;

    var colsResult = cols;
    if (!colsResult) {
      var resolved = EMP_getEmployeeColumns_();
      if (!resolved || !resolved.ok) return false;
      colsResult = resolved.cols;
    }

    var lookup = jobLookup;
    if (!lookup) {
      var mapRes = buildJobNameToIdMap_();
      if (!mapRes.ok) return false;
      lookup = mapRes.map;
    }

    var jobName = sh.getRange(rowIndex, colsResult.jobTypeNameCol).getValue();
    var deptVal = null;
    if (colsResult.departmentCol) {
      deptVal = sh.getRange(rowIndex, colsResult.departmentCol).getValue();
    }
    var key = normalizeKey_(jobName);
    if (!key) return false;

    var targetId = lookup[key];
    if (!targetId && typeof OPT !== "undefined") {
      var jobRec = null;
      if (OPT.getJobByNameAndDepartment) {
        jobRec = OPT.getJobByNameAndDepartment(jobName, deptVal);
      }
      if (!jobRec && OPT.getJobByName) {
        jobRec = OPT.getJobByName(jobName);
      }
      if (jobRec && jobRec.id) {
        targetId = jobRec.id;
      }
    }

    if (!targetId) {
      try {
        Logger.log("[EMP_fillJobIdForRow_] missing job type: " + jobName);
      } catch (_ignored2) {}
      return false;
    }

    var currentId = normalizeSpace_(
      sh.getRange(rowIndex, colsResult.jobTypeIdCol).getValue()
    );
    if (currentId === targetId) return false;

    sh.getRange(rowIndex, colsResult.jobTypeIdCol).setValue(targetId);
    return true;
  }

  /**
   * מבטיח ID עובד יציב לפי שם מלא:
   * - אם קיימת שורה עם אותו שם מלא ו-ID קיים → משתמשים באותו ID בכל השורות
   * - אם אין ID קיים לשם → מייצרים UUID פעם אחת ומשתמשים בו לכל השורות העתידיות עם אותו שם
   * - אם נמצא אותו שם עם שני IDs שונים → קונפליקט: לא דורסים IDs קיימים ולא ממלאים חסרים לשם הזה, ונרשם ללוג פעם אחת
   */
  function ensureEmployeeIds_(sheet) {
    if (!sheet) sheet = getEmployeesSheet_();
    if (!sheet) return;

    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return;

    var idCol = CONFIG.COL.ID; // B
    var nameCol = CONFIG.COL.FULL_NAME; // C

    var range = sheet.getRange(
      CONFIG.HEADER_ROW + 1,
      1,
      lastRow - CONFIG.HEADER_ROW,
      Math.max(idCol, nameCol)
    );
    var values = range.getValues();
    var txId = Utilities.getUuid();

    function normalizeName_(val) {
      if (val === null || val === undefined) return "";
      return String(val).replace(/\s+/g, " ").trim().toLowerCase();
    }
    function normalizeId_(val) {
      if (val === null || val === undefined) return "";
      return String(val).trim();
    }

    var existingIds = {};
    var nameToId = {};
    var conflictNames = {};
    var changed = false;

    // כדי לא להציף לוגים בכל עריכה, נשמור קאש של שמות בקונפליקט
    var props = PropertiesService.getDocumentProperties();
    var prevConflictArr;
    try {
      prevConflictArr = JSON.parse(
        props.getProperty("EMP_CONFLICT_NAMES") || "[]"
      );
      if (!Array.isArray(prevConflictArr)) prevConflictArr = [];
    } catch (e) {
      prevConflictArr = [];
    }
    var prevConflictSet = {};
    for (var pc = 0; pc < prevConflictArr.length; pc++)
      prevConflictSet[prevConflictArr[pc]] = true;

    // PASS 1: בונים מפה name->id מתוך IDs קיימים ומזהים קונפליקטים
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var name = normalizeName_(row[nameCol - 1]);
      var id = normalizeId_(row[idCol - 1]);

      if (id) existingIds[id] = true;
      if (!name || !id) continue;

      if (!nameToId[name]) {
        nameToId[name] = id;
      } else if (nameToId[name] !== id) {
        conflictNames[name] = true;

        // לוג פעם אחת לכל שם שמתנגש
        if (!prevConflictSet[name]) {
          logEmployeeChange_(
            id,
            CONFIG.SHEET_NAME_EMPLOYEES,
            CONFIG.HEADER_ROW + 1 + i,
            idCol,
            nameToId[name],
            id,
            txId
          );
        }

        // מונעים שימוש במיפוי הזה למילוי אוטומטי
        nameToId[name] = null;
      }
    }

    props.setProperty(
      "EMP_CONFLICT_NAMES",
      JSON.stringify(Object.keys(conflictNames))
    );

    // PASS 2: ממלאים IDs חסרים לפי המפה (או מייצרים פעם אחת)
    for (var j = 0; j < values.length; j++) {
      var row2 = values[j];
      var name2 = normalizeName_(row2[nameCol - 1]);
      if (!name2) continue;

      var currId = normalizeId_(row2[idCol - 1]);
      if (currId) continue; // לא דורסים ID קיים

      if (conflictNames[name2]) {
        // בקונפליקט לא ממלאים אוטומטית
        continue;
      }

      if (nameToId[name2]) {
        row2[idCol - 1] = nameToId[name2];
        changed = true;
        continue;
      }

      var newId;
      do {
        newId = generateUuid_();
      } while (existingIds[newId]);

      existingIds[newId] = true;
      row2[idCol - 1] = newId;
      nameToId[name2] = newId;
      changed = true;
    }

    if (changed) {
      range.setValues(values);
    }
  }

  /** צילום סכימה (שם+אינדקס של כל עמודה) לקריאת שינויים */
  function snapshotSchema_(sheet) {
    if (!sheet) sheet = getEmployeesSheet_();
    if (!sheet) return null;

    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    var result = [];
    for (var i = 0; i < headers.length; i++) {
      result.push((headers[i] || "") + "#" + (i + 1));
    }
    return result.join("||");
  }

  /** השוואת סכימה ושליחת מייל אם יש שינוי */
  function checkSchemaAndMaybeEmail_(sheet, reason) {
    if (!sheet) sheet = getEmployeesSheet_();
    if (!sheet) return;

    var props = getScriptProps_();
    var prev = props.getProperty(CONFIG.PROPS.SCHEMA) || "";
    var current = snapshotSchema_(sheet);
    if (!current) return;

    if (!prev) {
      // הפעלה ראשונה – שומרים בסיס
      props.setProperty(CONFIG.PROPS.SCHEMA, current);
      return;
    }

    if (prev === current) return;

    // הסכימה השתנתה – שולחים מייל, אבל לא משנים כלום
    props.setProperty(CONFIG.PROPS.SCHEMA, current);
    sendSchemaChangeEmail_(prev, current, reason);
  }

  function sendSchemaChangeEmail_(prev, current, reason) {
    var user = Session.getActiveUser().getEmail() || "unknown";
    var ss = getSpreadsheet_();

    var body = "";
    body +=
      'זוהה שינוי במבנה כרטיסיית "פרטי עובדים" בקובץ: ' + ss.getName() + "\n\n";
    body += "סיבה (טריגר): " + (reason || "לא ידוע") + "\n";
    body += "משתמש מבצע: " + user + "\n\n";
    body += 'תכנון המערכת בשיחה: "תכנון מערכת ID\'S"\n';
    body += "שים לב לעדכן את הקוד בהתאם לשינוי בעמודות (שמות/מיקומים).\n\n";
    body += "Snapshot קודם:\n" + prev + "\n\n";
    body += "Snapshot נוכחי:\n" + current + "\n";

    MailApp.sendEmail({
      to: CONFIG.MAIL.TO,
      subject:
        CONFIG.MAIL.SUBJECT_PREFIX + " (" + (reason || "שינוי סכימה") + ")",
      body: body,
    });
  }

  /** Skeleton לוג */
  function logEmployeeChange_(
    employeeId,
    sheetName,
    row,
    col,
    oldVal,
    newVal,
    txId
  ) {
    var sh = getLogSheet_();
    sh.appendRow([
      new Date(),
      txId || Utilities.getUuid(),
      employeeId || "",
      sheetName || "",
      row || "",
      col || "",
      oldVal === undefined ? "" : oldVal,
      newVal === undefined ? "" : newVal,
      Session.getActiveUser().getEmail() || "",
    ]);
  }

  /** אוסף את רשימת העובדים והסוגים שלהם לסטריפ */
  function collectEmployees_() {
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "פרטי עובדים"' };
    }

    // מוודא שלכל עובד יש ID עובד
    ensureEmployeeIds_(sheet);

    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) {
      return { ok: true, employees: [], cols: CONFIG.COL };
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];
    var colGender = colIndexByHeader_(headers, "מין");
    var colIdNum = colIndexByHeader_(headers, "תז");
    var colPhone = colIndexByHeader_(headers, "טלפון");
    var colBirthdate = colIndexByHeader_(headers, "ת. לידה");
    var colEmail = colIndexByHeader_(headers, "מייל");
    var colShirt = colIndexByHeader_(headers, "מידת חולצה");
    var colTravel = colIndexByHeader_(headers, "עלות החזרי נסיעות יומי");
    var colSystemRole = colIndexByHeader_(headers, "System Role");

    function normalizeStr(val) {
      if (val === null || val === undefined) return "";
      return String(val).trim();
    }

    function normalizeSystemRole(val) {
      var v = normalizeStr(val).toLowerCase();
      if (v === "admin" || v === "administrator") return "admin";
      if (v === "employee" || v === "user") return "employee";
      return "";
    }

    var data = sheet
      .getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, lastCol)
      .getValues();

    var employeesById = {};

    for (var r = 0; r < data.length; r++) {
      var rowIndex = CONFIG.HEADER_ROW + 1 + r;
      var row = data[r];

      var name = String(row[CONFIG.COL.FULL_NAME - 1] || "").trim();
      var id = String(row[CONFIG.COL.ID - 1] || "").trim();
      if (!name || !id) continue; // לא שורת עובד

      var statusVal = row[CONFIG.COL.ACTIVE - 1];
      var rowActive = isRowActiveFlag_(statusVal);

      var jobType = row[CONFIG.COL.JOB_TYPE - 1] || "";
      var dept = row[CONFIG.COL.DEPARTMENT - 1] || "";
      var amount = row[CONFIG.COL.AMOUNT - 1] || "";
      var payment = row[CONFIG.COL.PAYMENT_MODE - 1] || "";
      var notes = row[CONFIG.COL.NOTES - 1] || "";

      var rawGender = colGender ? row[colGender - 1] : "";
      var rawIdNum = colIdNum ? row[colIdNum - 1] : "";
      var rawPhone = colPhone ? row[colPhone - 1] : "";
      var rawBirthdate = colBirthdate ? row[colBirthdate - 1] : "";
      var rawEmail = colEmail ? row[colEmail - 1] : "";
      var rawShirt = colShirt ? row[colShirt - 1] : "";
      var rawTravel = colTravel ? row[colTravel - 1] : "";
      var rawSystemRole = colSystemRole ? row[colSystemRole - 1] : "";

      var gender = normalizeStr(rawGender);
      var idNum = normalizeStr(rawIdNum);
      var phone = normalizeStr(rawPhone);
      var birthdate = normalizeStr(rawBirthdate);
      var email = normalizeStr(rawEmail);
      var shirtSize = normalizeStr(rawShirt);
      var travelCost = normalizeStr(rawTravel);
      var systemRole = normalizeSystemRole(rawSystemRole);

      var emp = employeesById[id];
      if (!emp) {
        emp = {
          id: id,
          name: name,
          rows: [],
          anyActive: false,
          gender: gender,
          idNumber: idNum,
          phone: phone,
          birthdate: birthdate,
          email: email,
          shirtSize: shirtSize,
          travelCost: travelCost,
          systemRole: systemRole,
        };
        employeesById[id] = emp;
      } else {
        if (!emp.gender && gender) emp.gender = gender;
        if (!emp.idNumber && idNum) emp.idNumber = idNum;
        if (!emp.phone && phone) emp.phone = phone;
        if (!emp.birthdate && birthdate) emp.birthdate = birthdate;
        if (!emp.email && email) emp.email = email;
        if (!emp.shirtSize && shirtSize) emp.shirtSize = shirtSize;
        if (!emp.travelCost && travelCost) emp.travelCost = travelCost;
        if (systemRole === "admin") {
          emp.systemRole = "admin";
        } else if (!emp.systemRole && systemRole) {
          emp.systemRole = systemRole;
        }
      }

      emp.rows.push({
        rowIndex: rowIndex,
        rowActive: rowActive,
        jobType: jobType,
        department: dept,
        amount: amount,
        paymentMode: payment,
        notes: notes,
        personal: {
          fullName: name,
          gender: rawGender,
          idNumber: rawIdNum,
          phone: rawPhone,
          birthdate: rawBirthdate,
          email: rawEmail,
          shirtSize: rawShirt,
          travelCost: rawTravel,
          systemRole: rawSystemRole,
        },
      });

      if (rowActive) emp.anyActive = true;
    }

    var list = [];
    Object.keys(employeesById).forEach(function (idKey) {
      list.push(employeesById[idKey]);
    });

    list.sort(function (a, b) {
      return a.name.localeCompare(b.name);
    });

    return {
      ok: true,
      employees: list,
      cols: CONFIG.COL,
    };
  }

  function validateBulkActionDefinition_(definition) {
    if (!definition || typeof definition !== "object") {
      throw validationError_("Missing bulk action definition");
    }

    var mode = definition.mode;
    if (mode !== "DRY_RUN" && mode !== "EXECUTE") {
      throw validationError_("Unsupported mode: " + mode);
    }

    var bulkType = definition.bulkType;
    if (
      bulkType !== "EMPLOYEE_JOBTYPE_ADD" &&
      bulkType !== "EMPLOYEE_JOBTYPE_REMOVE"
    ) {
      throw validationError_("Unsupported bulkType: " + bulkType);
    }

    var params = definition.params || {};
    var jobTypeId = stringValue_(params.jobTypeId);
    if (!jobTypeId) {
      throw validationError_("Missing jobTypeId in params");
    }

    var jobRecord = null;
    if (typeof OPT !== "undefined" && OPT.getJobById) {
      jobRecord = OPT.getJobById(jobTypeId) || null;
    }
    if (!jobRecord) {
      throw validationError_("Unknown jobTypeId: " + jobTypeId);
    }

    var filters = definition.filters || {};
    var statusFilter = stringValue_(filters.status || "ALL").toUpperCase();
    var systemRoleFilter = stringValue_(
      filters.systemRole || "ALL"
    ).toUpperCase();
    var branchFilterRaw =
      filters.branch === null ? null : stringValue_(filters.branch);
    var branchFilter = branchFilterRaw ? branchFilterRaw : null;
    var filterJobTypeId = filters.jobTypeId
      ? stringValue_(filters.jobTypeId)
      : null;

    var allowedStatus = { ALL: true, ACTIVE: true, INACTIVE: true };
    if (!allowedStatus[statusFilter]) {
      throw validationError_("Invalid status filter: " + statusFilter);
    }

    var allowedRoles = {
      ALL: true,
      ADMIN: true,
      SHIFT_MANAGER: true,
      EMPLOYEE: true,
      VIEWER: true,
    };
    if (!allowedRoles[systemRoleFilter]) {
      throw validationError_("Invalid systemRole filter: " + systemRoleFilter);
    }

    return {
      mode: mode,
      bulkType: bulkType,
      params: { jobTypeId: jobTypeId },
      jobRecord: jobRecord,
      filters: {
        status: statusFilter,
        systemRole: systemRoleFilter,
        branch: branchFilter,
        jobTypeId: filterJobTypeId,
      },
    };
  }

  function loadEmployeesForBulkAction_(filters, logger) {
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      throw new Error('לא נמצאה כרטיסייה "פרטי עובדים"');
    }

    ensureEmployeeIds_(sheet);

    var colsResult = EMP_getEmployeeColumns_();
    if (!colsResult || !colsResult.ok) {
      throw new Error(
        colsResult && colsResult.error
          ? colsResult.error
          : "Missing employee columns"
      );
    }
    var cols = colsResult.cols;

    var headerRow = CONFIG.HEADER_ROW;
    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return [];

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    var colEmail =
      colIndexByHeader_(headers, "מייל") || colIndexByHeader_(headers, "email");
    var colSystemRole =
      colIndexByHeader_(headers, "System Role") ||
      colIndexByHeader_(headers, "system role");
    var colBranch =
      colIndexByHeader_(headers, "branch") ||
      colIndexByHeader_(headers, "Branch") ||
      colIndexByHeader_(headers, "סניף");

    var data = sheet
      .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
      .getValues();

    var byId = {};

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var absRow = headerRow + 1 + i;

      var employeeId = stringValue_(row[cols.employeeIdCol - 1]);
      var employeeName = stringValue_(row[CONFIG.COL.FULL_NAME - 1]);
      if (!employeeId || !employeeName) continue;

      var email = colEmail ? stringValue_(row[colEmail - 1]) : "";
      var systemRole = colSystemRole
        ? normalizeSystemRole_(row[colSystemRole - 1])
        : "";
      var branchVal = colBranch ? stringValue_(row[colBranch - 1]) : "";

      var rowStatusVal = row[CONFIG.COL.ACTIVE - 1];
      var rowActive = isRowActiveFlag_(rowStatusVal);
      var jobTypeIdsCell = parseJobTypeIdsCell_(row[cols.jobTypeIdCol - 1]);
      var department = stringValue_(row[cols.departmentCol - 1]);

      var emp = byId[employeeId];
      if (!emp) {
        emp = {
          id: employeeId,
          name: employeeName,
          email: email,
          systemRole: systemRole,
          branch: branchVal,
          anyActive: false,
          jobTypeIds: jobTypeIdsCell.slice(),
          rows: [],
          rowIndexForWrite: absRow,
        };
        byId[employeeId] = emp;
      } else {
        if (email && !emp.email) emp.email = email;
        if (systemRole && !emp.systemRole) emp.systemRole = systemRole;
        if (branchVal && !emp.branch) emp.branch = branchVal;
        // Merge job type ids from additional rows if present.
        jobTypeIdsCell.forEach(function (jid) {
          if (emp.jobTypeIds.indexOf(jid) === -1) emp.jobTypeIds.push(jid);
        });
      }

      emp.rows.push({
        rowIndex: absRow,
        active: rowActive,
        jobTypeIds: jobTypeIdsCell,
        department: department,
      });

      if (rowActive) emp.anyActive = true;
    }

    var statusFilter = filters.status;
    var branchFilter = filters.branch;
    var systemRoleFilter = filters.systemRole;
    var jobTypeFilter = filters.jobTypeId;

    var employees = [];
    Object.keys(byId).forEach(function (k) {
      var emp = byId[k];

      if (statusFilter === "ACTIVE" && !emp.anyActive) return;
      if (statusFilter === "INACTIVE" && emp.anyActive) return;

      if (branchFilter) {
        var empBranch = emp.branch ? emp.branch.toLowerCase() : "";
        if (!empBranch || empBranch !== branchFilter.toLowerCase()) return;
      }

      if (systemRoleFilter && systemRoleFilter !== "ALL") {
        var empRole = emp.systemRole ? emp.systemRole.toUpperCase() : "";
        if (empRole !== systemRoleFilter) return;
      }

      if (jobTypeFilter && emp.jobTypeIds.indexOf(jobTypeFilter) === -1) {
        return;
      }

      employees.push(emp);
    });

    return employees;
  }

  function computeBulkJobTypeBulkAction_(
    employees,
    bulkType,
    params,
    jobRecord
  ) {
    var jobTypeId =
      params && params.jobTypeId ? stringValue_(params.jobTypeId) : "";

    var summary = {
      targetEmployeesCount: employees.length,
      affectedEmployeesCount: 0,
      noChangeCount: 0,
    };
    var applied = [];
    var skipped = [];
    var errors = [];
    var changes = [];

    employees.forEach(function (emp) {
      try {
        var beforeIds = emp.jobTypeIds || [];
        var afterIds = beforeIds.slice();
        var had = beforeIds.indexOf(jobTypeId) !== -1;

        if (bulkType === "EMPLOYEE_JOBTYPE_ADD") {
          if (had) {
            skipped.push({
              employeeId: emp.id,
              employeeName: emp.name,
              email: emp.email || "",
              reasonCode: "ALREADY_HAS_JOBTYPE",
              reasonMessage: "העובד כבר מחזיק את סוג העבודה הזה",
            });
            summary.noChangeCount++;
            return;
          }
          afterIds.push(jobTypeId);
        } else if (bulkType === "EMPLOYEE_JOBTYPE_REMOVE") {
          if (!had) {
            skipped.push({
              employeeId: emp.id,
              employeeName: emp.name,
              email: emp.email || "",
              reasonCode: "JOBTYPE_NOT_ASSIGNED",
              reasonMessage: "לעובד אין את סוג העבודה הזה",
            });
            summary.noChangeCount++;
            return;
          }
          afterIds = beforeIds.filter(function (id) {
            return id !== jobTypeId;
          });
        } else {
          throw new Error("Unknown bulkType: " + bulkType);
        }

        var jobName = jobRecord && jobRecord.name ? String(jobRecord.name) : "";
        var jobShort =
          jobRecord && jobRecord.shortCode ? String(jobRecord.shortCode) : "";
        var jobLabel = jobName
          ? "'" +
            jobName +
            (jobShort ? " (" + jobShort + ")" : "") +
            "' (" +
            jobTypeId +
            ")"
          : jobTypeId;

        applied.push({
          employeeId: emp.id,
          employeeName: emp.name,
          email: emp.email || "",
          changeDescription:
            bulkType === "EMPLOYEE_JOBTYPE_ADD"
              ? "נוסף סוג העבודה " + jobLabel + "."
              : "הוסר סוג העבודה " + jobLabel + ".",
        });

        summary.affectedEmployeesCount++;

        var targetRowIndex =
          emp.rowIndexForWrite ||
          (emp.rows && emp.rows.length ? emp.rows[0].rowIndex : null);

        changes.push({
          employeeId: emp.id,
          employeeName: emp.name,
          email: emp.email || "",
          action: bulkType,
          rowIndexForWrite: targetRowIndex,
          beforeJobTypeIds: beforeIds,
          afterJobTypeIds: afterIds,
        });
      } catch (e) {
        errors.push({
          employeeId: emp.id || null,
          employeeName: emp.name || null,
          email: emp.email || null,
          errorCode: "ROW_PROCESSING_ERROR",
          errorMessage:
            e && e.message ? String(e.message) : "Unknown row error",
        });
        summary.noChangeCount++;
      }
    });

    return {
      result: {
        summary: summary,
        applied: applied,
        skipped: skipped,
        errors: errors,
      },
      changes: changes,
    };
  }

  function applyBulkJobTypeChanges_(
    computation,
    bulkType,
    params,
    jobRecord,
    logger
  ) {
    var changes = computation && computation.changes ? computation.changes : [];
    if (!changes.length) return;

    if (changes.length > 100) {
      throw new Error("Too many rows to update in one run (max 100)");
    }

    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      throw new Error("לא ניתן לקבל נעילה לביצוע הפעולה");
    }

    try {
      var sheet = getEmployeesSheet_();
      if (!sheet) {
        throw new Error('לא נמצאה כרטיסייה "פרטי עובדים"');
      }

      var colsResult = EMP_getEmployeeColumns_();
      if (!colsResult || !colsResult.ok) {
        throw new Error(
          colsResult && colsResult.error
            ? colsResult.error
            : "Missing employee columns"
        );
      }
      var cols = colsResult.cols;
      if (!cols.jobTypeIdCol) {
        throw new Error("Missing jobTypeId column");
      }

      var headerRow = CONFIG.HEADER_ROW;

      var updates = [];
      for (var i = 0; i < changes.length; i++) {
        var ch = changes[i];
        var rowIndex = ch.rowIndexForWrite || ch.rowIndex || ch.baseRowIndex;
        if (!rowIndex) {
          continue;
        }

        updates.push({
          rowIndex: rowIndex,
          value: serializeJobTypeIds_(ch.afterJobTypeIds || []),
        });
      }

      if (!updates.length) return;

      for (var u = 0; u < updates.length; u++) {
        var upd = updates[u];
        sheet.getRange(upd.rowIndex, cols.jobTypeIdCol).setValue(upd.value);
      }

      try {
        sheet
          .getRange(
            headerRow + 1,
            cols.jobTypeIdCol,
            sheet.getLastRow() - headerRow,
            1
          )
          .setNumberFormat("@");
      } catch (_ignored) {}
    } finally {
      lock.releaseLock();
    }
  }

  function EmployeesModule_adminRunBulkAction_(definition, context) {
    var logger =
      (context && context.logger) ||
      ensureModuleLoggerDefined_("ADMIN_BULK_ACTION_RUN");

    var validated = validateBulkActionDefinition_(definition);

    var employees = loadEmployeesForBulkAction_(validated.filters, logger);
    var computation = computeBulkJobTypeBulkAction_(
      employees,
      validated.bulkType,
      validated.params,
      validated.jobRecord
    );

    try {
      appendSystemLog_({
        layer: "admin.bulk",
        operation: "admin.runBulkAction",
        step: validated.mode,
        severity: "info",
        actor: (logger && logger.ctx && logger.ctx.actor) || "",
        details: {
          bulkType: validated.bulkType,
          mode: validated.mode,
          target: computation.result.summary.targetEmployeesCount,
          affected: computation.result.summary.affectedEmployeesCount,
          noChange: computation.result.summary.noChangeCount,
          errors: computation.result.errors
            ? computation.result.errors.length
            : 0,
        },
      });
    } catch (_ignored) {}

    if (logger && logger.info) {
      try {
        logger.info("admin.runBulkAction.summary", {
          mode: validated.mode,
          bulkType: validated.bulkType,
          target: computation.result.summary.targetEmployeesCount,
          affected: computation.result.summary.affectedEmployeesCount,
          noChange: computation.result.summary.noChangeCount,
          errors: computation.result.errors
            ? computation.result.errors.length
            : 0,
        });
      } catch (_ignored) {}
    }

    if (validated.mode === "EXECUTE") {
      applyBulkJobTypeChanges_(
        computation,
        validated.bulkType,
        validated.params,
        validated.jobRecord,
        logger
      );
    }

    return {
      mode: validated.mode,
      bulkType: validated.bulkType,
      definition: definition,
      result: computation.result,
    };
  }

  /**
   * Handles admin.reportBulkActionIssue.
   *
   * @param {Object} payload BulkActionIssueReportPayload
   * @param {Object} context ActionContext
   * @return {Object} BulkActionIssueReportResponse
   */
  function EmployeesModule_adminReportBulkActionIssue_(payload, context) {
    if (!payload) {
      return {
        ok: false,
        errorCode: "VALIDATION_ERROR",
        errorMessage: "Missing payload for admin.reportBulkActionIssue",
      };
    }

    var mode = payload.mode;
    var bulkType = payload.bulkType;
    var definition = payload.definition || {};
    var summary = payload.summary || {};
    var issueText = payload.issueText || "";
    var requestedBy = payload.requestedBy || "UNKNOWN";

    var filters = definition.filters || {};
    var params = definition.params || {};

    var filterJobTypeId = filters.jobTypeId || "";
    var paramJobTypeId = params.jobTypeId || "";

    var subject =
      "[Boulder Admin] דיווח על בעיה בפעולה קבוצתית - " +
      String(bulkType) +
      " (" +
      String(mode) +
      ")";

    var lines = [];
    lines.push("דווחה בעיה בפעולה קבוצתית בממשק הניהול של בולדר.");
    lines.push("");
    lines.push("פרטי הדיווח:");
    lines.push("מדווח/ת: " + String(requestedBy));
    lines.push("מצב פעולה (mode): " + String(mode));
    lines.push("סוג פעולה (bulkType): " + String(bulkType));
    lines.push("");
    lines.push("פילטרים:");
    lines.push("סטטוס עובדים: " + String(filters.status || ""));
    lines.push("סניף/מחלקה: " + String(filters.branch || ""));
    lines.push("תפקיד מערכת: " + String(filters.systemRole || ""));
    lines.push("jobTypeId בפילטר: " + String(filterJobTypeId));
    lines.push("");
    lines.push("פרמטרים:");
    lines.push("jobTypeId בפרמטרים: " + String(paramJobTypeId));
    lines.push("");
    lines.push("סיכום פעולה (summary):");
    lines.push('סה"כ בקהל היעד: ' + String(summary.targetEmployeesCount || 0));
    lines.push("עודכנו בפועל: " + String(summary.affectedEmployeesCount || 0));
    lines.push("ללא שינוי: " + String(summary.noChangeCount || 0));
    lines.push("");
    lines.push("תיאור הבעיה:");
    lines.push(issueText);
    lines.push("");
    lines.push("זמן שרת: " + new Date().toISOString());

    var body = lines.join("\n");

    var recipients = [];
    if (
      typeof ACCESS_ISSUE_RECIPIENTS !== "undefined" &&
      ACCESS_ISSUE_RECIPIENTS &&
      ACCESS_ISSUE_RECIPIENTS.length
    ) {
      recipients = ACCESS_ISSUE_RECIPIENTS;
    } else if (CONFIG && CONFIG.MAIL && CONFIG.MAIL.TO) {
      recipients = [CONFIG.MAIL.TO];
    }

    if (!recipients.length) {
      return {
        ok: false,
        errorCode: "CONFIGURATION_MISSING",
        errorMessage:
          "No admin notification recipients configured for bulk action issue report",
      };
    }

    try {
      MailApp.sendEmail({
        to: recipients.join(","),
        subject: subject,
        body: body,
      });

      try {
        appendSystemLog_({
          layer: "admin.bulk",
          operation: "admin.reportBulkActionIssue",
          step: "reported",
          severity: "warn",
          actor:
            (context &&
              context.logger &&
              context.logger.ctx &&
              context.logger.ctx.actor) ||
            requestedBy,
          details: {
            bulkType: bulkType,
            mode: mode,
            target: summary.targetEmployeesCount || 0,
            affected: summary.affectedEmployeesCount || 0,
            noChange: summary.noChangeCount || 0,
          },
        });
      } catch (_ignored) {}
    } catch (err) {
      return {
        ok: false,
        errorCode: "MAIL_SEND_FAILED",
        errorMessage:
          err && err.message ? String(err.message) : "Mail send failed",
      };
    }

    return {
      ok: true,
    };
  }

  function getSidebarBootstrap_() {
    try {
      var base = collectEmployees_();
      if (!base) {
        debugLog_(
          "bootstrap failed: collectEmployees_ returned null/undefined"
        );
        return { ok: false, error: "collectEmployees returned undefined/null" };
      }
      if (!base.ok) return base;

      var jobs = [];
      var payments = [];

      if (typeof OPT !== "undefined") {
        try {
          if (OPT.getAllJobs) {
            jobs = OPT.getAllJobs(true) || []; // רק פעילים
          }
        } catch (e) {
          debugLog_(
            "bootstrap jobs failed: " + (e && e.message ? e.message : e)
          );
          jobs = [];
        }

        try {
          if (OPT.getAllPayments) {
            payments = OPT.getAllPayments(true) || []; // רק פעילים
          }
        } catch (e2) {
          debugLog_(
            "bootstrap payments failed: " + (e2 && e2.message ? e2.message : e2)
          );
          payments = [];
        }
      }

      debugLog_(
        "bootstrap ok: employees=" +
          (base.employees ? base.employees.length : 0) +
          " jobs=" +
          jobs.length +
          " payments=" +
          payments.length
      );
      base.jobs = jobs;
      base.payments = payments;
      return base;
    } catch (e) {
      var msg = e && e.message ? e.message : e;
      var stk = e && e.stack ? "\n" + e.stack : "";
      debugLog_("bootstrap exception: " + msg);
      return { ok: false, error: "EMP_getSidebarBootstrap: " + msg + stk };
    }
  }

  function getEmployeeById_(employeeId) {
    try {
      var bootstrap = collectEmployees_();
      if (!bootstrap.ok) return bootstrap;

      var emp = null;
      for (var i = 0; i < bootstrap.employees.length; i++) {
        if (bootstrap.employees[i].id === employeeId) {
          emp = bootstrap.employees[i];
          break;
        }
      }
      if (!emp) {
        return { ok: false, error: "לא נמצא עובד עם ID הזה" };
      }
      return { ok: true, employee: emp, cols: CONFIG.COL };
    } catch (e) {
      var msg2 = e && e.message ? e.message : e;
      var stk2 = e && e.stack ? "\n" + e.stack : "";
      return { ok: false, error: "EMP_getEmployeeById: " + msg2 + stk2 };
    }
  }

  /**
   * שמירת payload מהסטריפ לתוך "פרטי עובדים".
   * לא משנה מבנה, עובד רק על השורות הרלוונטיות.
   */
  function saveEmployeePayload_(payload) {
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "פרטי עובדים"' };
    }

    ensureEmployeeIds_(sheet);

    // שמירה על עיצוב קיים: אם הסייבר מוסיף שורה חדשה, מעתיקים לה את העיצוב מהשורה האחרונה בטווח הנתונים
    function ensureRowFormatting_(targetRow) {
      var currentLastRow = sheet.getLastRow();
      if (targetRow <= currentLastRow) return;

      var templateRow = Math.max(CONFIG.HEADER_ROW + 1, currentLastRow);
      if (templateRow <= CONFIG.HEADER_ROW) return; // אין שורת נתונים ממנה ניתן להעתיק עיצוב

      var lastColForCopy = sheet.getLastColumn();
      sheet
        .getRange(templateRow, 1, 1, lastColForCopy)
        .copyFormatToRange(sheet, 1, lastColForCopy, targetRow, targetRow);
    }

    var employeeId = (payload.id || "").trim();
    var baseName = (payload.name || "").trim();
    if (!employeeId) {
      return { ok: false, error: "חסר ID עובד (ID עובד בעמודה B)" };
    }
    if (!baseName) {
      return { ok: false, error: "שם העובד לא יכול להיות ריק" };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= CONFIG.HEADER_ROW) {
      lastRow = CONFIG.HEADER_ROW;
    }

    var data = sheet
      .getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, lastCol)
      .getValues();

    var idRel = CONFIG.COL.ID - 1;
    var foundRows = [];
    var rowsTouched = [];

    for (var r = 0; r < data.length; r++) {
      var rowIdx = CONFIG.HEADER_ROW + 1 + r;
      var row = data[r];
      if (String(row[idRel] || "").trim() === employeeId) {
        foundRows.push({ rel: r, abs: rowIdx, row: row });
      }
    }

    var rowsPayload = payload.rows || [];
    var txId = Utilities.getUuid();

    foundRows.forEach(function (fr) {
      var oldVal = fr.row[CONFIG.COL.ACTIVE - 1];
      var newVal = "לא פעיל";
      if (String(oldVal || "") !== String(newVal)) {
        logEmployeeChange_(
          employeeId,
          sheet.getName(),
          fr.abs,
          CONFIG.COL.ACTIVE,
          oldVal,
          newVal,
          txId
        );
      }
      sheet.getRange(fr.abs, CONFIG.COL.ACTIVE).setValue(newVal);
    });

    for (var i = 0; i < rowsPayload.length; i++) {
      var pr = rowsPayload[i];
      var targetRow = pr.rowIndex ? parseInt(pr.rowIndex, 10) : null;

      if (!targetRow || targetRow <= CONFIG.HEADER_ROW) {
        targetRow = sheet.getLastRow() + 1;
      }

      ensureRowFormatting_(targetRow);

      // נאסוף כל שורה שנעדכן כדי שנוכל להשלים לה ID סוג עבודה בהמשך
      rowsTouched.push(targetRow);

      var colsToUpdate = [];

      colsToUpdate.push({
        col: CONFIG.COL.ACTIVE,
        newVal: pr.rowActive ? "פעיל" : "לא פעיל",
      });

      colsToUpdate.push({ col: CONFIG.COL.ID, newVal: employeeId });
      colsToUpdate.push({ col: CONFIG.COL.FULL_NAME, newVal: baseName });

      if (CONFIG.COL.JOB_TYPE) {
        colsToUpdate.push({
          col: CONFIG.COL.JOB_TYPE,
          newVal: pr.jobType || "",
        });
      }
      if (CONFIG.COL.DEPARTMENT) {
        colsToUpdate.push({
          col: CONFIG.COL.DEPARTMENT,
          newVal: pr.department || "",
        });
      }
      if (CONFIG.COL.AMOUNT) {
        colsToUpdate.push({ col: CONFIG.COL.AMOUNT, newVal: pr.amount || "" });
      }
      if (CONFIG.COL.PAYMENT_MODE) {
        colsToUpdate.push({
          col: CONFIG.COL.PAYMENT_MODE,
          newVal: pr.paymentMode || "",
        });
      }
      if (CONFIG.COL.NOTES) {
        colsToUpdate.push({ col: CONFIG.COL.NOTES, newVal: pr.notes || "" });
      }

      for (var j = 0; j < colsToUpdate.length; j++) {
        var info = colsToUpdate[j];
        if (!info.col) continue;

        var oldCell = sheet.getRange(targetRow, info.col).getValue();
        var newCell = info.newVal;

        if (String(oldCell || "") !== String(newCell || "")) {
          logEmployeeChange_(
            employeeId,
            sheet.getName(),
            targetRow,
            info.col,
            oldCell,
            newCell,
            txId
          );
          sheet.getRange(targetRow, info.col).setValue(newCell);
        }
      }
    }

    // עדכון שדות אישיים שנשלחו מהסיידבר – מיישרים לכל השורות של העובד
    var personalFields = {
      name: baseName,
      gender: payload.gender,
      idNumber: payload.idNumber,
      phone: payload.phone,
      birthdate: payload.birthdate,
      email: payload.email,
      shirtSize: payload.shirtSize,
      travelCost: payload.travelCost,
      systemRole: payload.systemRole,
    };

    function hasNonEmptyValue_(val) {
      if (val === null || val === undefined) return false;
      if (typeof val === "number") return true;
      return String(val).trim() !== "";
    }

    Object.keys(personalFields).forEach(function (key) {
      var val = personalFields[key];
      if (!hasNonEmptyValue_(val)) return;
      applyPersonalFieldChoice_(employeeId, key, val);
    });

    // השלמת ID סוג עבודה עבור שורות שנגעו (קיים או חדש) לפי שם סוג העבודה
    if (rowsTouched.length) {
      var colsResult = EMP_getEmployeeColumns_();
      var mapRes = buildJobNameToIdMap_();
      if (colsResult && colsResult.ok && mapRes && mapRes.ok) {
        var seenRow = {};
        for (var rt = 0; rt < rowsTouched.length; rt++) {
          var rNum = rowsTouched[rt];
          if (rNum <= CONFIG.HEADER_ROW) continue;
          if (seenRow[rNum]) continue;
          seenRow[rNum] = true;
          EMP_fillJobIdForRow_(sheet, rNum, colsResult.cols, mapRes.map);
        }
      }
    }

    // אחרי שמירה: ממלאים שדות אישיים חסרים בשורות אחרות של אותו עובד ללא דריסה
    fillMissingPersonalFieldsForEmployee_(sheet, employeeId, txId);

    return { ok: true, id: employeeId };
  }

  /** משלים שדות אישיים חסרים לשורות עובד ללא דריסה של ערכים קיימים */
  function fillMissingPersonalFieldsForEmployee_(sheet, employeeId, txId) {
    if (!sheet || !employeeId) return;

    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return;

    var lastCol = sheet.getLastColumn();
    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];

    var fieldHeaderMap = {
      name: "שם מלא",
      gender: "מין",
      idNumber: "תז",
      phone: "טלפון",
      birthdate: "ת. לידה",
      email: "מייל",
      shirtSize: "מידת חולצה",
      travelCost: "עלות החזרי נסיעות יומי",
      systemRole: "System Role",
    };

    function normalizeSystemRole(val) {
      var v = (val || "").toString().trim().toLowerCase();
      if (v === "admin" || v === "administrator") return "admin";
      if (v === "employee" || v === "user") return "employee";
      return "";
    }

    var colEmpId = colIndexByHeader_(headers, "ID עובד") || CONFIG.COL.ID;
    if (!colEmpId) return;

    var fieldCols = {};
    Object.keys(fieldHeaderMap).forEach(function (k) {
      var col = colIndexByHeader_(headers, fieldHeaderMap[k]);
      if (col) fieldCols[k] = col;
    });

    var data = sheet
      .getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, lastCol)
      .getValues();

    var valuesByField = {};
    var rowsByField = {};

    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      var absRow = CONFIG.HEADER_ROW + 1 + r;
      var eid = String(row[colEmpId - 1] || "").trim();
      if (eid !== String(employeeId)) continue;

      Object.keys(fieldCols).forEach(function (k) {
        var col = fieldCols[k];
        var val = row[col - 1];
        if (val !== null && val !== undefined && String(val).trim()) {
          if (valuesByField[k] === undefined) valuesByField[k] = val;
        } else {
          if (!rowsByField[k]) rowsByField[k] = [];
          rowsByField[k].push({ absRow: absRow, col: col });
        }
      });
    }

    var updates = [];
    Object.keys(rowsByField).forEach(function (k) {
      if (valuesByField[k] === undefined) return;
      var rows = rowsByField[k];
      for (var i = 0; i < rows.length; i++) {
        var target = rows[i];
        updates.push({
          row: target.absRow,
          col: target.col,
          value: valuesByField[k],
          fieldKey: k,
        });
      }
    });

    if (!updates.length) return;

    updates.forEach(function (u) {
      var oldVal = sheet.getRange(u.row, u.col).getValue();
      if (oldVal !== null && oldVal !== undefined && String(oldVal).trim()) {
        return; // בטיחות כפולה – לא דורסים ערכים קיימים
      }
      logEmployeeChange_(
        employeeId,
        sheet.getName(),
        u.row,
        u.col,
        oldVal,
        u.value,
        txId
      );
      sheet.getRange(u.row, u.col).setValue(u.value);
    });
  }

  /** יישום בחירת ערך אחיד לשדה אישי בכל השורות של עובד */
  function applyPersonalFieldChoice_(employeeId, fieldKey, value) {
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "פרטי עובדים"' };
    }

    var eid = (employeeId || "").trim();
    if (!eid) {
      return { ok: false, error: "חסר ID עובד" };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) {
      return { ok: false, error: "אין שורות נתונים בגליון" };
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];

    var fieldHeaderMap = {
      name: "שם מלא",
      gender: "מין",
      idNumber: "תז",
      phone: "טלפון",
      birthdate: "ת. לידה",
      email: "מייל",
      shirtSize: "מידת חולצה",
      travelCost: "עלות החזרי נסיעות יומי",
      systemRole: "System Role",
    };

    function normalizeSystemRole(val) {
      var v = (val || "").toString().trim().toLowerCase();
      if (v === "admin" || v === "administrator") return "admin";
      if (v === "employee" || v === "user") return "employee";
      return "";
    }

    var targetHeader = fieldHeaderMap[fieldKey];
    if (!targetHeader) {
      return { ok: false, error: "שדה לא נתמך לעדכון: " + fieldKey };
    }

    var colEmpId = colIndexByHeader_(headers, "ID עובד") || CONFIG.COL.ID;
    var targetCol = colIndexByHeader_(headers, targetHeader);

    if (!colEmpId) {
      return { ok: false, error: 'לא נמצאה עמודת "ID עובד" לזיהוי שורות' };
    }
    if (!targetCol) {
      return { ok: false, error: 'לא נמצאה עמודה לשדה "' + targetHeader + '"' };
    }

    var dataRange = sheet.getRange(
      CONFIG.HEADER_ROW + 1,
      1,
      lastRow - CONFIG.HEADER_ROW,
      lastCol
    );
    var data = dataRange.getValues();
    var txId = Utilities.getUuid();
    var updated = 0;

    for (var i = 0; i < data.length; i++) {
      var rowNum = CONFIG.HEADER_ROW + 1 + i;
      var rowEmpId = String(data[i][colEmpId - 1] || "").trim();
      if (rowEmpId !== eid) continue;

      var oldVal = data[i][targetCol - 1];
      var newVal = value === undefined ? "" : value;
      if (fieldKey === "systemRole") {
        newVal = normalizeSystemRole(newVal);
      }
      if (String(oldVal || "") === String(newVal || "")) continue;

      logEmployeeChange_(
        eid,
        sheet.getName(),
        rowNum,
        targetCol,
        oldVal,
        newVal,
        txId
      );
      sheet.getRange(rowNum, targetCol).setValue(newVal);
      updated++;
    }

    if (updated && (fieldKey === "idNumber" || fieldKey === "phone")) {
      sheet
        .getRange(
          CONFIG.HEADER_ROW + 1,
          targetCol,
          sheet.getLastRow() - CONFIG.HEADER_ROW,
          1
        )
        .setNumberFormat("@");
    }

    return { ok: true, updated: updated };
  }

  /** יצירת עובד חדש מתוך שם בלבד */
  function createEmployeeByName_(name) {
    var nm = (name || "").trim();
    if (!nm) {
      return { ok: false, error: "שם עובד ריק" };
    }
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "פרטי עובדים"' };
    }

    ensureEmployeeIds_(sheet);

    var newId = generateUuid_();
    var lastRow = sheet.getLastRow();
    var newRow = lastRow + 1;

    sheet.getRange(newRow, CONFIG.COL.ACTIVE).setValue("פעיל");
    sheet.getRange(newRow, CONFIG.COL.ID).setValue(newId);
    sheet.getRange(newRow, CONFIG.COL.FULL_NAME).setValue(nm);

    return getEmployeeById_(newId);
  }

  /** מבטל Filter קיים בגליון (אם יש) */
  function revealAllRowsIfFiltered_() {
    var sheet = getEmployeesSheet_();
    if (!sheet) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "פרטי עובדים"' };
    }
    var filter = sheet.getFilter();
    if (!filter) {
      return { ok: true, changed: false };
    }
    filter.remove();
    return { ok: true, changed: true };
  }

  /**
   * בדיקה אם יש שורות של עובד שמוסתרות ע"י FILTER בגליון "פרטי עובדים".
   * מחזיר true/false.
   */
  function hasHiddenRowsForEmployee_(employeeId) {
    if (!employeeId) return false;

    var sheet = getEmployeesSheet_();
    if (!sheet) return false;

    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return false;

    var idCol = CONFIG.COL.ID;
    var range = sheet.getRange(
      CONFIG.HEADER_ROW + 1,
      idCol,
      lastRow - CONFIG.HEADER_ROW,
      1
    );
    var values = range.getValues();

    for (var i = 0; i < values.length; i++) {
      var rowIndex = CONFIG.HEADER_ROW + 1 + i;
      var val = String(values[i][0] || "").trim();
      if (val === String(employeeId)) {
        if (sheet.isRowHiddenByFilter && sheet.isRowHiddenByFilter(rowIndex)) {
          return true;
        }
      }
    }
    return false;
  }

  /** פתיחת הסטריפ בפועל */
  function showSidebar_(optEmployeeId) {
    var tmpl = HtmlService.createTemplateFromFile("EmployeeSidebar");
    tmpl.initialEmployeeId = optEmployeeId || "";
    var html = tmpl
      .evaluate()
      .setTitle("ניהול עובדים – בולדר חיפה")
      .setWidth(380);
    SpreadsheetApp.getUi().showSidebar(html);
  }

  /** ריצה ב־onOpen (דרך EMP_onOpen) */
  function handleOpen_(e) {
    if (!e || e.authMode !== ScriptApp.AuthMode.FULL) {
      debugLog_("handleOpen skipped: authMode=" + (e && e.authMode));
      return;
    }

    var sheet = getEmployeesSheet_();
    if (!sheet) return;

    var skipMenu = e && e.__skipEmpMenu;

    if (!skipMenu) {
      try {
        SpreadsheetApp.getUi()
          .createMenu("בולדר עובדים")
          .addItem("פתח סייד בר עובדים", "EMP_openSidebar")
          .addItem("רענן סייד בר", "EMP_reloadSidebar")
          .addSeparator()
          .addItem("בדיקת באקפיל IDs (DRY_RUN)", "EMP_menuBackfillIdsDryRun")
          .addItem(
            "באקפיל IDs לכל העובדים (EXECUTE)",
            "EMP_menuBackfillIdsExecute"
          )
          .addToUi();
      } catch (_menuErr) {
        // ignore menu errors so onOpen continues
      }
    }

    ensureEmployeeIds_(sheet);
    checkSchemaAndMaybeEmail_(sheet, "onOpen");

    showSidebar_("");
  }

  /** ריצה ב־onChange (אם נגדיר טריגר גלובלי) */
  function handleChange_(e) {
    var sheet = getEmployeesSheet_();
    if (!sheet) return;
    checkSchemaAndMaybeEmail_(sheet, "onChange");
  }

  /** ריצה ב־onSelectionChange – בחירת שורה בגליון */
  function handleSelectionChange_(e) {
    if (!e || !e.range) return;
    var range = e.range;
    var sheet = range.getSheet();
    if (!sheet) return;
    if (sheet.getName() !== CONFIG.SHEET_NAME_EMPLOYEES) return;

    var row = range.getRow();
    if (row <= CONFIG.HEADER_ROW) return;

    ensureEmployeeIds_(sheet);

    var id = sheet.getRange(row, CONFIG.COL.ID).getValue();
    if (!id) return;

    showSidebar_(String(id));
  }

  function handleEdit_(e) {
    var range = e && e.range ? e.range : null;
    if (!range) return;

    var sheet = range.getSheet();
    if (!sheet) return;

    if (sheet.getName() !== CONFIG.SHEET_NAME_EMPLOYEES) return;

    var startCol = range.getColumn();
    var endCol = startCol + range.getNumColumns() - 1;
    var idCol = CONFIG.COL.ID; // 2
    var nameCol = CONFIG.COL.FULL_NAME; // 3
    var touchesIdOrName = !(endCol < idCol || startCol > nameCol);
    if (touchesIdOrName) {
      ensureEmployeeIds_(sheet);
    }

    var colsResult = EMP_getEmployeeColumns_();
    if (!colsResult || !colsResult.ok) return;
    var jobTypeNameCol = colsResult.cols.jobTypeNameCol;

    // אם עודכן System Role בשורה – מיישרים לכל השורות של העובד
    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    var colSystemRole = colIndexByHeader_(headers, "System Role") || 0;
    var touchesSystemRole =
      colSystemRole && startCol <= colSystemRole && colSystemRole <= endCol;
    if (touchesSystemRole) {
      var firstRow = range.getRow();
      var numRows = range.getNumRows();
      for (var r = 0; r < numRows; r++) {
        var rowIndex = firstRow + r;
        if (rowIndex <= CONFIG.HEADER_ROW) continue;
        var empIdVal = sheet.getRange(rowIndex, CONFIG.COL.ID).getValue();
        if (!empIdVal) continue;
        var roleVal = sheet.getRange(rowIndex, colSystemRole).getValue();
        applyPersonalFieldChoice_(String(empIdVal), "systemRole", roleVal);
      }
    }

    if (!jobTypeNameCol) return;

    var touchesJobTypeName =
      startCol <= jobTypeNameCol && jobTypeNameCol <= endCol;
    if (!touchesJobTypeName) return;

    var mapRes = buildJobNameToIdMap_();
    if (!mapRes.ok) {
      try {
        Logger.log("[EMP_handleEdit] cannot build job map: " + mapRes.error);
      } catch (_ignored) {}
      return;
    }

    var firstRow = range.getRow();
    var numRows = range.getNumRows();
    for (var i = 0; i < numRows; i++) {
      var rowIndex = firstRow + i;
      if (rowIndex <= CONFIG.HEADER_ROW) continue;
      EMP_fillJobIdForRow_(sheet, rowIndex, colsResult.cols, mapRes.map);
    }
  }

  // חשיפה ל-EMP (API פנימי)
  EMP.getSidebarBootstrap = getSidebarBootstrap_;
  EMP.getEmployeeById = getEmployeeById_;
  EMP.saveEmployeePayload = saveEmployeePayload_;
  EMP.createEmployeeByName = createEmployeeByName_;
  EMP.revealAllRowsIfFiltered = revealAllRowsIfFiltered_;
  EMP.applyPersonalFieldChoice = applyPersonalFieldChoice_;
  EMP.handleOpen = handleOpen_;
  EMP.handleChange = handleChange_;
  EMP.handleSelectionChange = handleSelectionChange_;
  EMP.handleEdit = handleEdit_;
  EMP.hasHiddenRowsForEmployee = hasHiddenRowsForEmployee_;
  EMP.ensureEmployeeIds = ensureEmployeeIds_;
  EMP.showSidebar = showSidebar_;
  EMP.getEmployeesSheet = getEmployeesSheet_;
  EMP.getLogSheet = getLogSheet_;
  EMP.CONFIG = CONFIG;
  EMP.adminRunBulkAction = EmployeesModule_adminRunBulkAction_;
  EMP.adminReportBulkActionIssue = EmployeesModule_adminReportBulkActionIssue_;
  // Expose for legacy callers that reference the global name directly
  if (typeof globalThis !== "undefined") {
    globalThis.EMP_getEmployeeColumns_ = EMP_getEmployeeColumns_;
    globalThis.EmployeesModule_adminRunBulkAction_ =
      EmployeesModule_adminRunBulkAction_;
    globalThis.EmployeesModule_adminReportBulkActionIssue_ =
      EmployeesModule_adminReportBulkActionIssue_;
  } else {
    this.EMP_getEmployeeColumns_ = EMP_getEmployeeColumns_;
    this.EmployeesModule_adminRunBulkAction_ =
      EmployeesModule_adminRunBulkAction_;
    this.EmployeesModule_adminReportBulkActionIssue_ =
      EmployeesModule_adminReportBulkActionIssue_;
  }
  // Export column resolver for backfill functions defined outside the IIFE
  EMP.getEmployeeColumnsForBackfill = EMP_getEmployeeColumns_;
})();

// Legacy alias to avoid ReferenceError for callers that still expect global CONFIG
if (typeof CONFIG === "undefined" && typeof EMP !== "undefined" && EMP.CONFIG) {
  var CONFIG = EMP.CONFIG;
}

/** === עטיפות גלובליות לטריגרים ול-HTML === */

// L3 entrypoint wired from index.gs:onEdit to run employee onEdit logic.
function EMP_onEdit(e) {
  try {
    if (typeof EMP !== "undefined" && typeof EMP.handleEdit === "function") {
      EMP.handleEdit(e || {});
    }
  } catch (err) {
    Logger.log("EMP_onEdit error: " + err);
  }
}

function EMP_onOpen(e) {
  // אם הפונקציה הזו מוגדרת בקובץ אחר - מצוין. אם לא, אפשר להסיר.
  if (typeof OPT_onOpen === "function") {
    try {
      OPT_onOpen(e);
    } catch (err) {
      Logger.log("OPT_onOpen error: " + (err && err.stack ? err.stack : err));
    }
  }
  if (typeof EMP_checkEmployeesHeader_ === "function") {
    EMP_checkEmployeesHeader_();
  }
  if (typeof EMP !== "undefined" && EMP.handleOpen) {
    EMP.handleOpen(e || {});
  }
}

function EMP_onChange(e) {
  if (typeof EMP !== "undefined" && EMP.handleChange) {
    EMP.handleChange(e || {});
  }
}

function onChange(e) {
  // אם כבר יש אצלך EMP_onChange — תשאיר את זה
  try {
    if (typeof EMP_onChange === "function") EMP_onChange(e || {});
  } catch (err1) {
    Logger.log("EMP_onChange error: " + err1);
  }

  // Schema monitor
  try {
    SCH_onChange(e || {});
  } catch (err2) {
    Logger.log("SCH_onChange error: " + err2);
  }
}

function onSelectionChange(e) {
  if (typeof EMP !== "undefined" && EMP.handleSelectionChange) {
    EMP.handleSelectionChange(e || {});
  }
}

function EMP_debugLog(msg) {
  try {
    Logger.log("[EMP_DEBUG] " + msg);
  } catch (_e) {
    // ignore
  }
}

function EMP_getSidebarBootstrap() {
  try {
    EMP_debugLog("EMP_getSidebarBootstrap: calling getSidebarBootstrap");

    var bootstrap =
      EMP && typeof EMP.getSidebarBootstrap === "function"
        ? EMP.getSidebarBootstrap()
        : null;

    if (!bootstrap) {
      EMP_debugLog("EMP_getSidebarBootstrap got empty bootstrap");
      return { error: "EMP_getSidebarBootstrap returned empty" };
    }

    // Enforce serialization-safe plain object for the sidebar client
    var clean = JSON.parse(JSON.stringify(bootstrap));

    var empCount = (clean.employees || []).length;
    var jobCount = (clean.jobs || []).length;
    var payCount = (clean.payments || []).length;

    EMP_debugLog(
      "EMP_getSidebarBootstrap ok (CLEAN): employees=" +
        empCount +
        " jobs=" +
        jobCount +
        " payments=" +
        payCount
    );

    return clean;
  } catch (err) {
    EMP_debugLog("EMP_getSidebarBootstrap error: " + err);
    return { error: "EMP_getSidebarBootstrap failed: " + err };
  }
}

function EMP_openSidebar() {
  try {
    if (typeof EMP !== "undefined" && typeof EMP.showSidebar === "function") {
      EMP.showSidebar("");
    }
  } catch (err) {
    Logger.log("EMP_openSidebar error: " + err);
  }
}

function EMP_reloadSidebar() {
  try {
    EMP_openSidebar();
  } catch (err) {
    Logger.log("EMP_reloadSidebar error: " + err);
  }
}

function EMP_getEmployeeById(id) {
  try {
    EMP_debugLog("EMP_getEmployeeById: requesting id=" + id);

    var res =
      EMP && typeof EMP.getEmployeeById === "function"
        ? EMP.getEmployeeById(id)
        : null;

    if (!res) {
      EMP_debugLog("EMP_getEmployeeById empty result for id=" + id);
      return { error: "EMP_getEmployeeById returned empty for id=" + id };
    }

    var clean = JSON.parse(JSON.stringify(res));
    var hasEmployee = !!(clean && clean.employee);

    EMP_debugLog(
      "EMP_getEmployeeById ok (CLEAN) for id=" +
        id +
        " hasEmployee=" +
        (hasEmployee ? "yes" : "no")
    );

    return clean;
  } catch (err) {
    EMP_debugLog("EMP_getEmployeeById error: " + err);
    return { error: "EMP_getEmployeeById failed: " + err };
  }
}

/**
 * LEGACY: שמירת עובד מה-Sidebar:
 * - ממפה סוג עבודה ואופן תשלום ל-ID לפי "אופציות בחירה ו ID'S"
 * - עובד לפי כותרות ולא לפי אינדקסים קשיחים
 * - מחזיר אובייקט עובד מעודכן (כולל rowIndex לכל שורת עבודה) כדי שהסיידבר יתעדכן ולא יווצרו כפילויות.
 */
function EMP_saveEmployeePayload_LEGACY_(payload) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { ok: false, error: "לא ניתן לקבל נעילה למסמך לשמירה (נסה שוב)." };
  }

  try {
    if (!payload || !payload.name || !payload.rows || !payload.rows.length) {
      return { ok: false, error: "payload לא תקין מה-Sidebar" };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("פרטי עובדים");
    if (!sheet) {
      return { ok: false, error: 'לא נמצאה כרטיסייה "פרטי עובדים"' };
    }

    var headerRow = 1;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= headerRow) {
      return { ok: false, error: 'אין נתונים בכרטיסייה "פרטי עובדים"' };
    }

    var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    function colIndexByHeaderLocal(name) {
      return colIndexByHeader_(headers, name);
    }

    var colStatus = colIndexByHeaderLocal("סטטוס");
    var colName = colIndexByHeaderLocal("שם מלא");
    var colEmpId = colIndexByHeaderLocal("ID עובד");
    var colJobId = colIndexByHeaderLocal("ID סוגי עבודה");
    var colJobName = colIndexByHeaderLocal("סוג העבודה");
    var colDept = colIndexByHeaderLocal("מחלקה");
    var colAmount = colIndexByHeaderLocal("סכום");
    var colPayId = colIndexByHeaderLocal("ID אופן תשלום");
    var colPayName = colIndexByHeaderLocal("אופן תשלום");
    var colNotes = colIndexByHeaderLocal("הערות");
    var colGender = colIndexByHeaderLocal("מין");
    var colIdNum = colIndexByHeaderLocal("תז");
    var colPhone = colIndexByHeaderLocal("טלפון");
    var colBirthdate = colIndexByHeaderLocal("ת. לידה");
    var colEmail = colIndexByHeaderLocal("מייל");
    var colShirt = colIndexByHeaderLocal("מידת חולצה");
    var colTravel = colIndexByHeaderLocal("עלות החזרי נסיעות יומי");
    var colSystemRole = colIndexByHeaderLocal("System Role");

    if (!colName || !colJobName || !colStatus) {
      return {
        ok: false,
        error:
          'חסרות עמודות חובה (לפחות "שם מלא", "סוג העבודה", "סטטוס") – בדוק כותרות.',
      };
    }

    var dataRange = sheet.getRange(
      headerRow + 1,
      1,
      lastRow - headerRow,
      lastCol
    );
    var data = dataRange.getValues();
    var rowsToFormat = [];

    var rowNumToIndex = {};
    for (var i = 0; i < data.length; i++) {
      var rowNum = headerRow + 1 + i;
      rowNumToIndex[rowNum] = i;
    }

    var name = String(payload.name).trim();
    var employeeId = payload.id ? String(payload.id).trim() : "";

    function normalizeStr(val) {
      if (val === null || val === undefined) return "";
      return String(val).replace(/\s+/g, " ").trim();
    }

    function normalizeSystemRole(val) {
      var v = normalizeStr(val).toLowerCase();
      if (v === "admin" || v === "administrator") return "admin";
      if (v === "employee" || v === "user") return "employee";
      return "";
    }

    function resolveJob(jobTypeText) {
      var key = normalizeStr(jobTypeText);
      if (!key) return null;
      if (typeof OPT === "undefined" || !OPT.getJobByName) return null;
      return OPT.getJobByName(key);
    }

    function resolvePayment(paymentText) {
      var key = normalizeStr(paymentText);
      if (!key) return null;
      if (typeof OPT === "undefined" || !OPT.getPaymentByName) return null;
      return OPT.getPaymentByName(key);
    }

    var targetSystemRole =
      normalizeSystemRole(payload.systemRole) || "employee";

    // עדכון שורות קיימות + איסוף השורות החדשות (ללא rowIndex)
    var rowsPayload = payload.rows;
    var newRowsPayload = [];

    for (var r = 0; r < rowsPayload.length; r++) {
      var rp = rowsPayload[r];

      var allEmpty =
        !normalizeStr(rp.jobType) &&
        !normalizeStr(rp.department) &&
        !normalizeStr(rp.amount) &&
        !normalizeStr(rp.paymentMode) &&
        !normalizeStr(rp.notes);
      if (allEmpty) continue;

      var rowIndex = rp.rowIndex ? Number(rp.rowIndex) : null;

      if (rowIndex && rowNumToIndex.hasOwnProperty(rowIndex)) {
        var di = rowNumToIndex[rowIndex];
        var rowArr = data[di];
        rowsToFormat.push(rowIndex);

        // שם ו-ID עובד
        if (name) {
          rowArr[colName - 1] = name;
        }
        if (colEmpId && employeeId) {
          rowArr[colEmpId - 1] = employeeId;
        }

        // סטטוס
        if (colStatus) {
          rowArr[colStatus - 1] = rp.rowActive ? "פעיל" : "לא פעיל";
        }

        if (colGender) {
          rowArr[colGender - 1] = normalizeStr(payload.gender);
        }
        if (colIdNum) {
          rowArr[colIdNum - 1] = normalizeStr(payload.idNumber);
        }
        if (colPhone) {
          rowArr[colPhone - 1] = normalizeStr(payload.phone);
        }
        if (colBirthdate) {
          rowArr[colBirthdate - 1] = normalizeStr(payload.birthdate);
        }
        if (colEmail) {
          rowArr[colEmail - 1] = normalizeStr(payload.email);
        }
        if (colShirt) {
          rowArr[colShirt - 1] = normalizeStr(payload.shirtSize);
        }
        if (colTravel) {
          rowArr[colTravel - 1] = normalizeStr(payload.travelCost);
        }
        if (colSystemRole) {
          rowArr[colSystemRole - 1] = targetSystemRole;
        }

        // סכום
        if (colAmount) {
          var amountVal = normalizeStr(rp.amount);
          rowArr[colAmount - 1] = amountVal ? Number(amountVal) : "";
        }

        // הערות
        if (colNotes) {
          rowArr[colNotes - 1] = normalizeStr(rp.notes);
        }

        // סוג עבודה
        var jobRec = resolveJob(rp.jobType);
        if (jobRec) {
          if (colJobId) rowArr[colJobId - 1] = jobRec.id || "";
          if (colJobName) rowArr[colJobName - 1] = jobRec.name || "";
          if (colDept) rowArr[colDept - 1] = jobRec.department || "";
        } else {
          if (colJobId) rowArr[colJobId - 1] = "";
          if (colJobName) rowArr[colJobName - 1] = normalizeStr(rp.jobType);
          if (colDept) rowArr[colDept - 1] = normalizeStr(rp.department);
        }

        // אופן תשלום
        if (colPayName || colPayId) {
          var payRec = resolvePayment(rp.paymentMode);
          if (payRec) {
            if (colPayId) rowArr[colPayId - 1] = payRec.id || "";
            if (colPayName) rowArr[colPayName - 1] = payRec.name || "";
          } else {
            if (colPayId) rowArr[colPayId - 1] = "";
            if (colPayName)
              rowArr[colPayName - 1] = normalizeStr(rp.paymentMode);
          }
        }
      } else {
        // שורה חדשה – נטפל בה אחרי setValues
        newRowsPayload.push(rp);
      }
    }

    // כתיבה מחדש של הבלוק הקיים (שורות עם rowIndex)
    dataRange.setValues(data);

    // טיפול בשורות חדשות – הוספה בתחתית הטבלה + העתקת פורמט
    var appendedRowIndices = [];
    if (newRowsPayload.length > 0) {
      var appendRows = [];
      var baseRowTemplate = null;
      var baseRowTemplateRowNum = null;

      // ניסיון לבחור שורת תבנית של אותו עובד – לפי ID עובד
      if (colEmpId && employeeId) {
        for (var i2 = 0; i2 < data.length; i2++) {
          var rowEmpId = normalizeStr(data[i2][colEmpId - 1]);
          if (rowEmpId && rowEmpId === employeeId) {
            baseRowTemplate = data[i2].slice();
            baseRowTemplateRowNum = headerRow + 1 + i2;
            break;
          }
        }
      }

      // אם אין ID עובד, fallback לפי שם
      if (!baseRowTemplate && colName) {
        for (var i3 = 0; i3 < data.length; i3++) {
          if (normalizeStr(data[i3][colName - 1]) === name) {
            baseRowTemplate = data[i3].slice();
            baseRowTemplateRowNum = headerRow + 1 + i3;
            break;
          }
        }
      }

      // אם עדיין אין – נשתמש בשורת הנתונים הראשונה כטמפלט
      if (!baseRowTemplate && data.length > 0) {
        baseRowTemplate = data[0].slice();
        baseRowTemplateRowNum = headerRow + 1;
      }

      // אם גם זה אין (קיצון) – מייצרים מערך ריק
      if (!baseRowTemplate) {
        baseRowTemplate = new Array(lastCol);
        for (var c = 0; c < lastCol; c++) baseRowTemplate[c] = "";
      }

      for (var nr = 0; nr < newRowsPayload.length; nr++) {
        var rpNew = newRowsPayload[nr];
        var newRow = baseRowTemplate.slice();

        // שם ו-ID עובד
        if (colName && name) {
          newRow[colName - 1] = name;
        }
        if (colEmpId && employeeId) {
          newRow[colEmpId - 1] = employeeId;
        }

        // סטטוס
        if (colStatus) {
          newRow[colStatus - 1] = rpNew.rowActive ? "פעיל" : "לא פעיל";
        }

        if (colGender) {
          newRow[colGender - 1] = normalizeStr(payload.gender);
        }
        if (colIdNum) {
          newRow[colIdNum - 1] = normalizeStr(payload.idNumber);
        }
        if (colPhone) {
          newRow[colPhone - 1] = normalizeStr(payload.phone);
        }
        if (colBirthdate) {
          newRow[colBirthdate - 1] = normalizeStr(payload.birthdate);
        }
        if (colEmail) {
          newRow[colEmail - 1] = normalizeStr(payload.email);
        }
        if (colShirt) {
          newRow[colShirt - 1] = normalizeStr(payload.shirtSize);
        }
        if (colTravel) {
          newRow[colTravel - 1] = normalizeStr(payload.travelCost);
        }
        if (colSystemRole) {
          newRow[colSystemRole - 1] = targetSystemRole;
        }

        // סכום
        if (colAmount) {
          var amountNew = normalizeStr(rpNew.amount);
          newRow[colAmount - 1] = amountNew ? Number(amountNew) : "";
        }

        // הערות
        if (colNotes) {
          newRow[colNotes - 1] = normalizeStr(rpNew.notes);
        }

        // סוג עבודה
        var jobRecNew = resolveJob(rpNew.jobType);
        if (jobRecNew) {
          if (colJobId) newRow[colJobId - 1] = jobRecNew.id || "";
          if (colJobName) newRow[colJobName - 1] = jobRecNew.name || "";
          if (colDept) newRow[colDept - 1] = jobRecNew.department || "";
        } else {
          if (colJobId) newRow[colJobId - 1] = "";
          if (colJobName) newRow[colJobName - 1] = normalizeStr(rpNew.jobType);
          if (colDept) newRow[colDept - 1] = normalizeStr(rpNew.department);
        }

        // אופן תשלום
        if (colPayName || colPayId) {
          var payRecNew = resolvePayment(rpNew.paymentMode);
          if (payRecNew) {
            if (colPayId) newRow[colPayId - 1] = payRecNew.id || "";
            if (colPayName) newRow[colPayName - 1] = payRecNew.name || "";
          } else {
            if (colPayId) newRow[colPayId - 1] = "";
            if (colPayName)
              newRow[colPayName - 1] = normalizeStr(rpNew.paymentMode);
          }
        }

        appendRows.push(newRow);
      }

      if (appendRows.length > 0) {
        var startRow = lastRow + 1;
        sheet
          .getRange(startRow, 1, appendRows.length, lastCol)
          .setValues(appendRows);

        // העתקת פורמט ואימות נתונים מהתבנית (אם יש)
        if (baseRowTemplateRowNum) {
          var templateRange = sheet.getRange(
            baseRowTemplateRowNum,
            1,
            1,
            lastCol
          );
          for (var k = 0; k < appendRows.length; k++) {
            var destRow = startRow + k;
            templateRange.copyTo(sheet.getRange(destRow, 1, 1, lastCol), {
              formatOnly: true,
            });
            appendedRowIndices.push(destRow);
          }
        } else {
          for (var k2 = 0; k2 < appendRows.length; k2++) {
            appendedRowIndices.push(startRow + k2);
          }
        }
      }
    }

    // ריענון פורמט עבור עמודות ID (סוג עבודה / אופן תשלום) בשורות שעודכנו
    var formatCols = [];
    if (colJobId) formatCols.push(colJobId);
    if (colPayId) formatCols.push(colPayId);

    var templateRowForFormat = baseRowTemplateRowNum || headerRow + 1;
    if (formatCols.length && templateRowForFormat) {
      var allRowsNeedingFormat = rowsToFormat.concat(appendedRowIndices);
      if (allRowsNeedingFormat.length) {
        for (var fc = 0; fc < formatCols.length; fc++) {
          var colFmt = formatCols[fc];
          // העתקת פורמט מהשורה התבניתית
          sheet
            .getRange(templateRowForFormat, colFmt, 1, 1)
            .copyFormatToRange(
              sheet,
              colFmt,
              colFmt,
              Math.min.apply(null, allRowsNeedingFormat),
              Math.max.apply(null, allRowsNeedingFormat)
            );
          // שמירת ה-ID כטקסט כדי למנוע פורמט מספרי
          sheet
            .getRange(headerRow + 1, colFmt, sheet.getLastRow() - headerRow, 1)
            .setNumberFormat("@");
        }
      }
    }

    var rowsNeedingJobId = rowsToFormat.concat(appendedRowIndices);
    if (rowsNeedingJobId.length) {
      var colsResult = EMP_getEmployeeColumns_();
      var mapRes = buildJobNameToIdMap_();
      if (colsResult && colsResult.ok && mapRes && mapRes.ok) {
        var seenRows = {};
        for (var f = 0; f < rowsNeedingJobId.length; f++) {
          var rNum = rowsNeedingJobId[f];
          if (rNum <= headerRow) continue;
          if (seenRows[rNum]) continue;
          seenRows[rNum] = true;
          EMP_fillJobIdForRow_(sheet, rNum, colsResult.cols, mapRes.map);
        }
      }
    }

    // טעינת העובד המעודכן לתשובה – זה מה שהסיידבר ישתמש בו כדי לעדכן rowIndex
    var employee = null;
    if (employeeId && typeof getEmployeeById_ === "function") {
      employee = getEmployeeById_(employeeId);
    }

    return {
      ok: true,
      employee: employee,
      employeeId: employeeId,
      appendedRowIndices: appendedRowIndices,
    };
  } catch (err) {
    return {
      ok: false,
      error:
        "שגיאה ב-EMP_saveEmployeePayload: " +
        (err && err.message ? err.message : err),
    };
  } finally {
    lock.releaseLock();
  }
}

function EMP_saveEmployeePayload(payload) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { ok: false, error: "לא ניתן לקבל נעילה למסמך לשמירה (נסה שוב)." };
  }

  try {
    if (!payload || !payload.name || !payload.rows || !payload.rows.length) {
      return { ok: false, error: "payload לא תקין מה-Sidebar" };
    }

    var saveResult;
    if (
      typeof EMP !== "undefined" &&
      EMP &&
      typeof EMP.saveEmployeePayload === "function"
    ) {
      saveResult = EMP.saveEmployeePayload(payload);
    } else if (typeof saveEmployeePayload_ === "function") {
      saveResult = saveEmployeePayload_(payload);
    } else {
      saveResult = { ok: false, error: "saveEmployeePayload_ missing" };
    }

    // Phase 3 (disabled): backfill after sidebar save is paused until Gate2.
    // Backfill/normalize should be run via Admin menu only (DRY_RUN first).

    return saveResult;
  } catch (err) {
    return {
      ok: false,
      error:
        "שגיאה ב-EMP_saveEmployeePayload: " +
        (err && err.message ? err.message : err),
    };
  } finally {
    lock.releaseLock();
  }
}

function EMP_applyPersonalFieldChoice(employeeId, fieldKey, value) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    return { ok: false, error: "לא ניתן לקבל נעילה למסמך לעדכון (נסה שוב)." };
  }
  try {
    if (typeof EMP === "undefined" || !EMP.applyPersonalFieldChoice) {
      return { ok: false, error: "פונקציית עדכון שדה אישי אינה זמינה" };
    }
    return EMP.applyPersonalFieldChoice(employeeId, fieldKey, value);
  } catch (err) {
    return {
      ok: false,
      error:
        "שגיאה ב-EMP_applyPersonalFieldChoice: " +
        (err && err.message ? err.message : err),
    };
  } finally {
    lock.releaseLock();
  }
}

function EMP_createEmployeeByName(name) {
  return EMP.createEmployeeByName(name);
}

function EMP_normalizeJobAndPayRow_(
  sheet,
  rows,
  rowIndex,
  sheetRow,
  cols,
  missingJobs,
  missingPayments,
  actions,
  opts
) {
  var row = rows[rowIndex];
  var options = opts || {};
  var dryRun = options.dryRun !== false; // default DRY_RUN

  function norm_(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  function writeIfChanged_(colIdx, val) {
    if (row[colIdx - 1] === val) return;
    row[colIdx - 1] = val;
    if (!dryRun) {
      sheet.getRange(sheetRow, colIdx).setValue(val);
    }
    rowUpdated = true;
  }

  var rowUpdated = false;
  var employeeId = norm_(row[cols.employeeIdCol - 1]);

  // Job
  var jobTypeId = norm_(row[cols.jobTypeIdCol - 1]);
  var jobTypeName = norm_(row[cols.jobTypeNameCol - 1]);
  var department = norm_(row[cols.departmentCol - 1]);

  // If ID already exists, we leave it untouched (fill-missing-only).
  if (!jobTypeId && jobTypeName) {
    var jobByName =
      (typeof OPT !== "undefined" && OPT.getJobByNameAndDepartment
        ? OPT.getJobByNameAndDepartment(jobTypeName, department)
        : null) || null;
    if (jobByName && jobByName.id) {
      writeIfChanged_(cols.jobTypeIdCol, jobByName.id);
      if (actions) {
        actions.push({
          employeeId: employeeId,
          rowIndex: sheetRow,
          field: "jobTypeId",
          action: "FILL_MISSING_JOB_ID",
          value: jobByName.id,
        });
      }
    } else {
      missingJobs.push({
        employeeId: employeeId,
        rowIndex: sheetRow,
        jobName: jobTypeName,
        department: department,
        reason: "name-not-found",
      });
    }
  }

  // Payment
  var payTypeId = norm_(row[cols.payTypeIdCol - 1]);
  var payTypeName = norm_(row[cols.payTypeNameCol - 1]);

  if (!payTypeId && payTypeName) {
    var payByName =
      (typeof OPT !== "undefined" && OPT.getPaymentByName
        ? OPT.getPaymentByName(payTypeName)
        : null) || null;
    if (payByName && payByName.id) {
      writeIfChanged_(cols.payTypeIdCol, payByName.id);
      if (actions) {
        actions.push({
          employeeId: employeeId,
          rowIndex: sheetRow,
          field: "payTypeId",
          action: "FILL_MISSING_PAYMENT_ID",
          value: payByName.id,
        });
      }
    } else {
      missingPayments.push({
        employeeId: employeeId,
        rowIndex: sheetRow,
        payName: payTypeName,
        reason: "name-not-found",
      });
    }
  }

  return rowUpdated;
}

/**
 * סנכרון:
 * - J: "ID סוגי עבודה"
 * - L: "מחלקה"
 * - O: "ID אופן תשלום"
 * בכרטיסייה "פרטי עובדים"
 */
function EMP_backfillJobAndPaymentIdsForAllEmployees(opts) {
  var options = opts || {};
  var dryRun = options.dryRun !== false; // default DRY_RUN
  var empConfig = typeof EMP !== "undefined" ? EMP.CONFIG : null;
  if (!empConfig) {
    throw new Error("EMP.CONFIG is not available for backfill");
  }

  var sheet =
    typeof EMP !== "undefined" && EMP.getEmployeesSheet
      ? EMP.getEmployeesSheet()
      : null;
  if (!sheet) {
    throw new Error(
      'לא נמצאה כרטיסייה "' + empConfig.SHEET_NAME_EMPLOYEES + '"'
    );
  }

  var colsResult = null;
  if (typeof EMP !== "undefined" && EMP.getEmployeeColumnsForBackfill) {
    colsResult = EMP.getEmployeeColumnsForBackfill();
  } else if (typeof EMP_getEmployeeColumns_ === "function") {
    colsResult = EMP_getEmployeeColumns_();
  }
  if (!colsResult || !colsResult.ok) {
    var errMsg =
      "EMP_backfillJobAndPaymentIds: " +
      (colsResult && colsResult.error
        ? colsResult.error
        : "EMP_getEmployeeColumns_ missing");
    try {
      Logger.log(errMsg);
    } catch (_ignored) {}
    try {
      SpreadsheetApp.getActive().toast(
        "EMP backfill skipped: " + colsResult.error,
        "EMP_backfillJobAndPaymentIds",
        7
      );
    } catch (_ignored2) {}
    return {
      ok: false,
      updated: 0,
      missingJobs: [],
      missingPayments: [],
      error: colsResult.error,
    };
  }

  var cols = colsResult.cols;
  var HEADER_ROW = empConfig.HEADER_ROW;
  var COL_FULL_NAME = empConfig.COL.FULL_NAME;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= HEADER_ROW) {
    SpreadsheetApp.getActive().toast(
      "אין נתוני עובדים לסנכרון",
      "EMP_backfillJobAndPaymentIds",
      5
    );
    return { ok: true, updated: 0, missingJobs: [], missingPayments: [] };
  }

  var data = sheet
    .getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
    .getValues();

  var updated = 0;
  var missingJobs = [];
  var missingPayments = [];
  var actions = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var sheetRow = HEADER_ROW + 1 + i;

    var empId = row[cols.employeeIdCol - 1];
    var fullName = row[COL_FULL_NAME - 1];

    if (!empId && !fullName) continue;
    var rowUpdated = EMP_normalizeJobAndPayRow_(
      sheet,
      data,
      i,
      sheetRow,
      cols,
      missingJobs,
      missingPayments,
      actions,
      { dryRun: dryRun }
    );

    if (rowUpdated) updated++;
  }

  var msg =
    (dryRun ? "(DRY_RUN) " : "") +
    "עודכנו " +
    updated +
    " עובדים (ID סוג עבודה / מחלקה / ID אופן תשלום).";
  if (missingJobs.length)
    msg +=
      " אין התאמה ל-" + missingJobs.length + " סוגי עבודה (פירוט ב-Logger).";
  if (missingPayments.length)
    msg +=
      " אין התאמה ל-" +
      missingPayments.length +
      " אופני תשלום (פירוט ב-Logger).";

  SpreadsheetApp.getActive().toast(msg, "EMP_backfillJobAndPaymentIds", 7);

  return {
    ok: true,
    updated: updated,
    missingJobs: missingJobs,
    missingPayments: missingPayments,
    actions: actions,
    dryRun: dryRun,
  };
}

/**
 * Backfills job/payment IDs for a single employeeId.
 * @param {string} employeeId
 * @param {{dryRun?:boolean}=} opts
 * @return {{ ok: boolean, employeeId: string, updatedRows: number[], missingJobs: Object[], missingPayments: Object[], actions?: Object[], error?: string, dryRun?: boolean }}
 * @private
 */
function EMP_backfillJobAndPaymentIdsForEmployee(employeeId, opts) {
  var options = opts || {};
  var dryRun = options.dryRun !== false; // default DRY_RUN
  var empConfig = typeof EMP !== "undefined" ? EMP.CONFIG : null;
  if (!empConfig) {
    return {
      ok: false,
      employeeId: "",
      updatedRows: [],
      missingJobs: [],
      missingPayments: [],
      actions: [],
      dryRun: dryRun,
      error: "EMP.CONFIG is not available for backfill",
    };
  }
  var normalizedId = employeeId ? String(employeeId).trim() : "";
  if (!normalizedId) {
    return {
      ok: false,
      employeeId: normalizedId,
      updatedRows: [],
      missingJobs: [],
      missingPayments: [],
      actions: [],
      dryRun: dryRun,
      error: "missing employeeId",
    };
  }

  var sheet =
    typeof EMP !== "undefined" && EMP.getEmployeesSheet
      ? EMP.getEmployeesSheet()
      : null;
  if (!sheet) {
    return {
      ok: false,
      employeeId: normalizedId,
      updatedRows: [],
      missingJobs: [],
      missingPayments: [],
      actions: [],
      dryRun: dryRun,
      error: 'לא נמצאה כרטיסייה "' + empConfig.SHEET_NAME_EMPLOYEES + '"',
    };
  }

  var colsResult = null;
  if (typeof EMP !== "undefined" && EMP.getEmployeeColumnsForBackfill) {
    colsResult = EMP.getEmployeeColumnsForBackfill();
  } else if (typeof EMP_getEmployeeColumns_ === "function") {
    colsResult = EMP_getEmployeeColumns_();
  }
  if (!colsResult || !colsResult.ok) {
    return {
      ok: false,
      employeeId: normalizedId,
      updatedRows: [],
      missingJobs: [],
      missingPayments: [],
      actions: [],
      dryRun: dryRun,
      error:
        colsResult && colsResult.error
          ? colsResult.error
          : "EMP_getEmployeeColumns_ missing",
    };
  }

  var cols = colsResult.cols;
  var HEADER_ROW = empConfig.HEADER_ROW;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= HEADER_ROW) {
    return {
      ok: true,
      employeeId: normalizedId,
      updatedRows: [],
      missingJobs: [],
      missingPayments: [],
      actions: [],
      dryRun: dryRun,
    };
  }

  var data = sheet
    .getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
    .getValues();

  var updatedRows = [];
  var missingJobs = [];
  var missingPayments = [];
  var actions = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowEmpId = row[cols.employeeIdCol - 1];
    if (String(rowEmpId || "").trim() !== normalizedId) continue;

    var sheetRow = HEADER_ROW + 1 + i;
    var rowUpdated = EMP_normalizeJobAndPayRow_(
      sheet,
      data,
      i,
      sheetRow,
      cols,
      missingJobs,
      missingPayments,
      actions,
      { dryRun: dryRun }
    );
    if (rowUpdated && !dryRun) updatedRows.push(sheetRow);
  }

  return {
    ok: true,
    employeeId: normalizedId,
    updatedRows: updatedRows,
    missingJobs: missingJobs,
    missingPayments: missingPayments,
    actions: actions,
    dryRun: dryRun,
  };
}

function EMP_logBackfill_(
  mode,
  scope,
  updatedCount,
  missingJobsCount,
  missingPaymentsCount
) {
  var logSheet =
    typeof EMP !== "undefined" && EMP.getLogSheet ? EMP.getLogSheet() : null;
  if (!logSheet) return;
  var actor = "unknown";
  try {
    actor = Session.getActiveUser().getEmail() || "unknown";
  } catch (_e) {}

  logSheet.appendRow([
    new Date(),
    "EMP_ID_BACKFILL",
    mode,
    scope,
    updatedCount,
    missingJobsCount,
    missingPaymentsCount,
    actor,
  ]);
}

function EMP_menuBackfillIdsDryRun() {
  var result = EMP_backfillJobAndPaymentIdsForAllEmployees({ dryRun: true });
  var msg =
    "(DRY_RUN) עודכנו " +
    (result.updated || 0) +
    " עובדים; חסרים: " +
    (result.missingJobs.length || 0) +
    " סוגי עבודה, " +
    (result.missingPayments.length || 0) +
    " אופני תשלום.";
  try {
    SpreadsheetApp.getActive().toast(msg, "EMP_backfillJobAndPaymentIds", 7);
  } catch (_ignored) {}
  try {
    var actionsLen = Array.isArray(result.actions) ? result.actions.length : 0;
    Logger.log(msg + " actions_len=" + actionsLen);
  } catch (_ignored2) {}
}

function EMP_menuBackfillIdsExecute() {
  var ui = SpreadsheetApp.getUi();
  var dry = EMP_backfillJobAndPaymentIdsForAllEmployees({ dryRun: true });
  var prompt =
    "ימולאו IDs חסרים עבור " +
    (dry.updated || 0) +
    " עובדים. חסרים: " +
    (dry.missingJobs.length || 0) +
    " סוגי עבודה, " +
    (dry.missingPayments.length || 0) +
    " אופני תשלום. להמשיך?";
  var answer = ui.alert("EMP Backfill", prompt, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  var exec = EMP_backfillJobAndPaymentIdsForAllEmployees({ dryRun: false });
  EMP_logBackfill_(
    "EXECUTE",
    "ALL",
    exec.updated || 0,
    exec.missingJobs.length || 0,
    exec.missingPayments.length || 0
  );

  var msg =
    "(EXECUTE) מולאו IDs עבור " +
    (exec.updated || 0) +
    " עובדים; חסרים: " +
    (exec.missingJobs.length || 0) +
    " סוגי עבודה, " +
    (exec.missingPayments.length || 0) +
    " אופני תשלום.";
  try {
    SpreadsheetApp.getActive().toast(msg, "EMP_backfillJobAndPaymentIds", 7);
  } catch (_ignored3) {}
  try {
    var execActionsLen = Array.isArray(exec.actions) ? exec.actions.length : 0;
    Logger.log(msg + " actions_len=" + execActionsLen);
  } catch (_ignored4) {}
}
