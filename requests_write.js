// requests_write.gs
// כתיבת בקשות מהווב-אפ לגיליון "בקשות עובדים" + יצירת UUID אוטומטי לעמודת "ID בקשה"

const SHEET_NAME_REQUESTS = "בקשות עובדים";
const SHEET_NAME_EMPLOYEES_FOR_LOOKUP = "פרטי עובדים";

// eslint-disable-next-line no-unused-vars
function handleRequestPost(e) {
  var logger = null;
  var startMs = new Date().getTime();
  try {
    if (typeof ensureModuleLoggerDefined_ === "function") {
      logger = ensureModuleLoggerDefined_("REQUEST_POST");
    }
  } catch (_ignoredLogger) {
    // Logger unavailable, continue without it
  }

  function finish_(res, errorCode, err) {
    if (typeof logDuration_ === "function") {
      logDuration_(
        logger,
        "request.post",
        startMs,
        {
          success: !!(res && res.success),
          requestId: res && res.requestId,
          mode: res && res.mode,
          error: res && res.error,
        },
        res && res.success ? "info" : "warn",
        errorCode,
        err,
      );
    }
    return res;
  }

  var lock = LockService.getDocumentLock();
  var hasLock = lock.tryLock(30000);
  if (!hasLock) {
    return finish_(
      jsonResponse({
        success: false,
        error: "Lock unavailable (concurrent write)",
      }),
      "LOCK_UNAVAILABLE",
    );
  }

  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : "";
    var data = null;

    // Try JSON body first
    if (raw) {
      try {
        data = JSON.parse(raw);
      } catch (err) {
        return finish_(
          jsonResponse({
            success: false,
            error: "INVALID_JSON",
            details: String(err),
            rawSample: raw.slice(0, 200),
          }),
          "INVALID_JSON",
          err,
        );
      }
    }

    // Fallback: try e.parameter (legacy calls with query params)
    if (!data && e && e.parameter && typeof e.parameter === "object") {
      data = e.parameter;
    }

    // Still no data → fail
    if (!data || typeof data !== "object") {
      return finish_(
        jsonResponse({ success: false, error: "Missing body or parameter" }),
        "NO_POST_DATA",
      );
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_REQUESTS);
    if (!sheet) {
      return finish_(
        jsonResponse({
          success: false,
          error: "SHEET_NOT_FOUND",
          sheetName: SHEET_NAME_REQUESTS,
        }),
        "SHEET_NOT_FOUND",
      );
    }

    var headerRow = 1;
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) lastCol = 1;

    var headers = sheet
      .getRange(headerRow, 1, 1, lastCol)
      .getValues()[0]
      .map(function (h) {
        return String(h || "").trim();
      });

    function colByHeader_(name) {
      var idx = headers.indexOf(name);
      return idx >= 0 ? idx + 1 : null; // 1-based
    }

    function colByAny_(names) {
      for (var i = 0; i < names.length; i++) {
        var c = colByHeader_(names[i]);
        if (c) return c;
      }
      return null;
    }

    function normalize_(v) {
      if (v === null || v === undefined) return "";
      return String(v).replace(/\s+/g, " ").trim();
    }

    // עמודות לפי ה-Contract שלך
    var COL_STATUS = colByAny_(["סטטוס בקשה"]);
    var COL_REQUEST_ID = colByAny_(["ID בקשה"]);
    var COL_SHIFT_NOTE = colByAny_(["הערות למשמרת"]);
    var COL_SHIFT_ID = colByAny_(["ID משמרת"]);
    var COL_TIMESTAMP = colByAny_(["חותמת זמן"]);
    var COL_EMP_ID = colByAny_(["ID עובד"]);
    var COL_EMP_NAME = colByAny_(["שם מלא"]);
    var COL_DIRECTION = colByAny_(["כניסה / יציאה"]);
    var COL_FIX_DATE = colByAny_(["תיקון תאריך"]);
    var COL_FIX_TIME = colByAny_(["תיקון שעה"]);
    var COL_JOB_ID = colByAny_(["ID סוג עבודה", "ID סוגי עבודה"]);
    var COL_JOB_NAME = colByAny_(["סוג עבודה", "סוגי עבודה", "סוג העבודה"]);
    var COL_DEPT = colByAny_(["מחלקה"]);
    var COL_UNITS = colByAny_(["כמות היחידות"]);

    // מינימום חובה כדי לא לכתוב "שורות זבל"
    if (!COL_REQUEST_ID || !COL_TIMESTAMP) {
      return finish_(
        jsonResponse({
          success: false,
          error: "MISSING_REQUIRED_HEADERS",
          details:
            'חסרות כותרות חובה (לפחות "ID בקשה" ו-"חותמת זמן") בכרטיסייה "בקשות עובדים".',
        }),
        "MISSING_REQUIRED_HEADERS",
      );
    }

    // --- יצירת/שימוש ב-ID בקשה ---
    var requestId = normalize_(data.requestId);
    if (!requestId) {
      requestId = Utilities.getUuid(); // UUID
    }

    // --- הכנות נתונים ---
    var now = new Date();

    var status =
      data.status !== undefined ? normalize_(data.status) : "pending";

    var employeeId = normalize_(data.employeeId);
    var employeeName = normalize_(data.employeeName);

    // אם אין שם אבל יש ID עובד — ננסה לשלוף שם מ"פרטי עובדים" (B=ID, C=שם מלא)
    if (!employeeName && employeeId) {
      var lookedUp = lookupEmployeeNameById_(ss, employeeId);
      if (lookedUp) employeeName = lookedUp;
    }

    var shiftId = normalize_(data.shiftId);
    var shiftNote =
      data.shiftNote !== undefined ? normalize_(data.shiftNote) : "";

    var direction =
      data.direction !== undefined ? normalize_(data.direction) : "";
    var fixDate = data.fixDate !== undefined ? normalize_(data.fixDate) : "";
    var fixTime = data.fixTime !== undefined ? normalize_(data.fixTime) : "";

    var jobId = normalize_(data.jobId);
    var jobName = normalize_(data.jobName);
    var department = normalize_(data.department);

    // אם אין department ונוכל לפתור job מה-OPT — נשלים
    var jobRec = resolveJob_(jobId, jobName);
    if (jobRec) {
      if (!jobId) jobId = normalize_(jobRec.id);
      if (!jobName) jobName = normalize_(jobRec.name);
      if (!department) department = normalize_(jobRec.department);
    }

    var units =
      data.units !== undefined && data.units !== null && data.units !== ""
        ? Number(data.units)
        : "";

    // --- UPSERT: אם כבר קיימת בקשה עם אותו ID — נעדכן רק שדות שהגיעו בפועל ---
    var existingRow = findRowByRequestId_(sheet, COL_REQUEST_ID, requestId);
    if (existingRow) {
      var updates = [];

      // מעדכנים רק אם השדה קיים ב-payload (hasOwnProperty), כדי לא למחוק בטעות
      function maybeUpdate_(col, key, value) {
        if (!col) return;
        if (data.hasOwnProperty(key)) {
          updates.push({ col: col, val: value });
        }
      }

      maybeUpdate_(COL_STATUS, "status", status);
      maybeUpdate_(COL_SHIFT_NOTE, "shiftNote", shiftNote);
      maybeUpdate_(COL_SHIFT_ID, "shiftId", shiftId);
      // חותמת זמן: אם הלקוח שולח timestamp — נשתמש, אחרת נשאיר הקיים
      if (COL_TIMESTAMP && data.hasOwnProperty("timestamp")) {
        updates.push({
          col: COL_TIMESTAMP,
          val: normalize_(data.timestamp) || now,
        });
      }

      maybeUpdate_(COL_EMP_ID, "employeeId", employeeId);
      maybeUpdate_(COL_EMP_NAME, "employeeName", employeeName);
      maybeUpdate_(COL_DIRECTION, "direction", direction);
      maybeUpdate_(COL_FIX_DATE, "fixDate", fixDate);
      maybeUpdate_(COL_FIX_TIME, "fixTime", fixTime);
      maybeUpdate_(COL_JOB_ID, "jobId", jobId);
      maybeUpdate_(COL_JOB_NAME, "jobName", jobName);
      maybeUpdate_(COL_DEPT, "department", department);
      maybeUpdate_(COL_UNITS, "units", units);

      for (var u = 0; u < updates.length; u++) {
        sheet.getRange(existingRow, updates[u].col).setValue(updates[u].val);
      }

      return finish_(
        jsonResponse({
          success: true,
          requestId: requestId,
          mode: "updated",
          rowIndex: existingRow,
        }),
      );
    }

    // --- INSERT חדש ---
    var lastRow = sheet.getLastRow();
    var newRowIndex = lastRow + 1;

    // נייצר שורה מלאה באורך lastCol, ונמלא רק מה שיש לנו
    var rowArr = new Array(lastCol);
    for (var i = 0; i < lastCol; i++) rowArr[i] = "";

    if (COL_STATUS) rowArr[COL_STATUS - 1] = status;
    rowArr[COL_REQUEST_ID - 1] = requestId;
    if (COL_SHIFT_NOTE) rowArr[COL_SHIFT_NOTE - 1] = shiftNote;
    if (COL_SHIFT_ID) rowArr[COL_SHIFT_ID - 1] = shiftId;
    if (COL_TIMESTAMP) rowArr[COL_TIMESTAMP - 1] = now;
    if (COL_EMP_ID) rowArr[COL_EMP_ID - 1] = employeeId;
    if (COL_EMP_NAME) rowArr[COL_EMP_NAME - 1] = employeeName;
    if (COL_DIRECTION) rowArr[COL_DIRECTION - 1] = direction;
    if (COL_FIX_DATE) rowArr[COL_FIX_DATE - 1] = fixDate;
    if (COL_FIX_TIME) rowArr[COL_FIX_TIME - 1] = fixTime;
    if (COL_JOB_ID) rowArr[COL_JOB_ID - 1] = jobId;
    if (COL_JOB_NAME) rowArr[COL_JOB_NAME - 1] = jobName;
    if (COL_DEPT) rowArr[COL_DEPT - 1] = department;
    if (COL_UNITS) rowArr[COL_UNITS - 1] = units;

    sheet.getRange(newRowIndex, 1, 1, lastCol).setValues([rowArr]);

    return finish_(
      jsonResponse({
        success: true,
        requestId: requestId,
        mode: "inserted",
        rowIndex: newRowIndex,
      }),
    );
  } catch (err2) {
    return finish_(
      jsonResponse({
        success: false,
        error: "UNEXPECTED",
        details: String(err2),
      }),
      "UNEXPECTED",
      err2,
    );
  } finally {
    try {
      lock.releaseLock();
    } catch (e2) {
      // Lock release error, continue
    }
  }
}

function findRowByRequestId_(sheet, requestIdCol, requestId) {
  if (!sheet || !requestIdCol || !requestId) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var rng = sheet.getRange(2, requestIdCol, lastRow - 1, 1);
  var found = rng
    .createTextFinder(String(requestId))
    .matchEntireCell(true)
    .findNext();

  return found ? found.getRow() : null;
}

function lookupEmployeeNameById_(ss, employeeId) {
  try {
    var sh = ss.getSheetByName(SHEET_NAME_EMPLOYEES_FOR_LOOKUP);
    if (!sh) return "";

    // לפי המודול שלך: B=ID עובד, C=שם מלא
    var HEADER_ROW = 1;
    var COL_ID = 2;
    var COL_NAME = 3;

    var lastRow = sh.getLastRow();
    if (lastRow <= HEADER_ROW) return "";

    var rng = sh.getRange(HEADER_ROW + 1, COL_ID, lastRow - HEADER_ROW, 1);
    var found = rng
      .createTextFinder(String(employeeId))
      .matchEntireCell(true)
      .findNext();

    if (!found) return "";

    var r = found.getRow();
    return String(sh.getRange(r, COL_NAME).getValue() || "").trim();
  } catch (e) {
    return "";
  }
}

function resolveJob_(jobId, jobName) {
  try {
    // אם יש שם — הכי טוב
    if (jobName && typeof OPT !== "undefined" && OPT.getJobByName) {
      var rec = OPT.getJobByName(String(jobName).trim());
      if (rec) return rec;
    }

    // אם יש ID — ואין לנו getJobById חשוף, ננסה דרך getAllJobs
    if (jobId && typeof OPT !== "undefined" && OPT.getAllJobs) {
      var list = OPT.getAllJobs(false) || [];
      for (var i = 0; i < list.length; i++) {
        if (String(list[i].id || "").trim() === String(jobId).trim()) {
          return list[i];
        }
      }
    }
  } catch (e) {
    // Intentionally ignore OPT lookup errors
  }
  return null;
}
