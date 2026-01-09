/* global SHEET_NAME, OPT */
/* exported SHIFTS_list, SHIFTS_get, SHIFTS_getSelectedRef, SHIFTS_updateShift, SHIFTS_updateBonuses, BONUSES_listAll */
var SHIFTS = SHIFTS || {};
var BONUSES = BONUSES || {};

(function () {
  "use strict";

  var CONFIG = {
    SHEET_NAME: typeof SHEET_NAME !== "undefined" ? SHEET_NAME : "משמרות",
    HEADER_ROW: 1,
    BONUS_SEPARATOR: ",",
    COLS: {
      notes: {
        candidates: ["הערות למשמרת", "הערות", "notes"],
        fallbackIndex: 1,
      },
      shiftId: {
        candidates: ["ID משמרת", "shiftId", "Shift ID"],
        fallbackIndex: 2,
      },
      timestamp: {
        candidates: ["חותמת זמן", "timestamp", "Timestamp"],
        fallbackIndex: 3,
      },
      employeeId: {
        candidates: ["ID עובד", "employee_id", "Employee ID"],
        fallbackIndex: 4,
      },
      employeeName: {
        candidates: ["שם מלא", "employeeName", "Employee Name"],
        fallbackIndex: 5,
      },
      email: { candidates: ["מייל", "email", "Email"], fallbackIndex: null },
      direction: {
        candidates: ["כניסה / יציאה", "כיוון", "direction"],
        fallbackIndex: 6,
      },
      fixDate: {
        candidates: ["תיקון תאריך", "תאריך", "date"],
        fallbackIndex: 7,
      },
      fixTime: { candidates: ["תיקון שעה", "שעה", "time"], fallbackIndex: 8 },
      jobId: {
        candidates: ["ID סוג עבודה", "ID סוגי עבודה", "jobId"],
        fallbackIndex: 9,
      },
      jobName: {
        candidates: ["סוג עבודה", "סוג העבודה", "jobName", "workType"],
        fallbackIndex: 10,
      },
      department: { candidates: ["מחלקה", "department"], fallbackIndex: 11 },
      units: {
        candidates: ["כמות היחידות", "יחידות", "units"],
        fallbackIndex: 12,
      },
      status: {
        candidates: ["סטטוס משמרת", "סטטוס", "status"],
        fallbackIndex: null,
      },
      workDate: {
        candidates: ["תאריך", "תאריך משמרת", "Work Date"],
        fallbackIndex: null,
      },
      startTime: {
        candidates: ["כניסה", "שעת התחלה", "Start Time", "שעת כניסה"],
        fallbackIndex: null,
      },
      endTime: {
        candidates: ["יציאה", "שעת סיום", "End Time", "שעת יציאה"],
        fallbackIndex: null,
      },
      bonusIds: {
        candidates: ["BonusIds", "Bonus IDs", "בונוסים"],
        fallbackIndex: null,
      },
      manualEdited: {
        candidates: ["ManualEdited", "Manual Edited", "עריכה ידנית"],
        fallbackIndex: null,
      },
      manualNote: {
        candidates: ["ManualNote", "הערת מנהל", "הערת עריכה"],
        fallbackIndex: null,
      },
      lastUpdatedBySidebar: {
        candidates: ["LastUpdatedBySidebar", "עודכן בסיידבר על ידי"],
        fallbackIndex: null,
      },
      lastUpdatedAt: {
        candidates: [
          "LastUpdatedAt",
          "עודכן בסיידבר בתאריך",
          "SidebarUpdatedAt",
        ],
        fallbackIndex: null,
      },
      hasIssues: {
        candidates: ["HasIssues", "בעיה", "Issues"],
        fallbackIndex: null,
      },
    },
  };

  var RAW_WORK_LOG_SHEET_NAME = "דיווח שעות עבודה";
  var MAX_YESTERDAY_REBUILD_PAIRS = 200; // safety guard for SHIFTS_rebuildYesterday

  // התאמת shiftId לפי סכימה ידועה
  if (CONFIG.SHEET_NAME === "דיווח שעות עבודה") {
    CONFIG.COLS.shiftId = {
      candidates: ["ID דיווח"],
      fallbackIndex: null,
    };
  } else if (CONFIG.SHEET_NAME === "משמרות") {
    CONFIG.COLS.shiftId = {
      candidates: ["ID משמרת"],
      fallbackIndex: null,
    };
  }

  var NORMALIZED_SHIFT_CORE_HEADERS = [
    { field: "ShiftId", header: "ShiftId" },
    { field: "EmployeeId", header: "EmployeeId" },
    { field: "EmployeeName", header: "EmployeeName" },
    { field: "JobTypeId", header: "JobTypeId" },
    { field: "JobTypeName", header: "JobTypeName" },
    { field: "Department", header: "Department" },
    { field: "ShiftDate", header: "ShiftDate" },
    { field: "StartTime", header: "StartTime" },
    { field: "EndTime", header: "EndTime" },
    { field: "StartDateTime", header: "StartDateTime" },
    { field: "EndDateTime", header: "EndDateTime" },
    { field: "SpanHours", header: "SpanHours" },
    { field: "PayHours", header: "PayHours" },
    { field: "Status", header: "Status" },
    { field: "Note", header: "Note" },
    { field: "SourceReportIds", header: "SourceReportIds" },
  ];

  function getLogger_(operation) {
    try {
      if (typeof ensureModuleLoggerDefined_ === "function") {
        return ensureModuleLoggerDefined_(operation || "SHIFTS_MODULE");
      }
    } catch (_ignored) {}
    return null;
  }

  function ss_() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getSheet_() {
    var ctx = getShiftsSheetAndHeaderMap_();
    return ctx.sheet;
  }

  function getRawWorkLogSheet_() {
    var sh = ss_().getSheetByName(RAW_WORK_LOG_SHEET_NAME);
    if (!sh)
      throw new Error(
        'לא נמצאה כרטיסייה בשם "' + RAW_WORK_LOG_SHEET_NAME + '"'
      );
    return sh;
  }

  var SHIFT_HEADER_CANDIDATES = {
    shiftId: ["ID משמרת", "ShiftId", "shiftId"],
    employeeId: ["מזהה עובד", "ID עובד", "EmployeeId", "employeeId"],
    employeeName: ["שם מלא", "EmployeeName", "employeeName"],
    jobTypeId: ["ID סוג עבודה", "ID סוגי עבודה", "JobTypeId", "jobTypeId"],
    jobTypeName: ["סוג עבודה", "סוגי עבודה", "JobTypeName", "jobTypeName"],
    department: ["מחלקה", "מחלקות", "Department", "department"],
    shiftDate: ["תאריך", "תאריך משמרת", "ShiftDate", "shiftDate"],
    startTime: [
      "כניסה",
      "שעת התחלה",
      "Start Time",
      "StartTime",
      "startTime",
      "שעת כניסה",
    ],
    endTime: [
      "יציאה",
      "שעת סיום",
      "End Time",
      "EndTime",
      "endTime",
      "שעת יציאה",
    ],
    startDateTime: ["StartDateTime", "שעת התחלה מדויקת"],
    endDateTime: ["EndDateTime", "שעת סיום מדויקת"],
    spanHours: ["שעות", "SpanHours", "Hours"],
    payHours: ["שעות לשכר", "PayHours"],
    status: ["סטטוס משמרת", "סטטוס", "Status"],
    note: ["הערות", "הערות למשמרת", "Note"],
    sourceReportIds: ["מקור דיווחים", "SourceReportIds"],
  };

  var REQUIRED_SHIFT_HEADERS = Object.keys(SHIFT_HEADER_CANDIDATES).map(
    function (key) {
      return { key: key, header: SHIFT_HEADER_CANDIDATES[key][0] };
    }
  );

  function formatNumericColumn_(sheet, colIndex) {
    if (!colIndex) return;
    sheet.getRange(1, colIndex, sheet.getMaxRows(), 1).setNumberFormat("0.00");
  }

  function ensureRequiredShiftHeaders_(sheet) {
    void sheet;
    var ctx = getShiftsSheetAndHeaderMap_();
    return ctx.headerMap;
  }

  function getOrCreateShiftsSheet_() {
    var ctx = getShiftsSheetAndHeaderMap_();
    return ctx.sheet;
  }

  function getShiftsHeaderMap_() {
    var ctx = getShiftsSheetAndHeaderMap_();
    return ctx.headerMap;
  }

  function normalizeDirection_(raw) {
    var d = normKey_(raw);
    if (d === "in" || d === "כניסה" || d === "כניסה / יציאה") return "IN";
    if (d === "out" || d === "יציאה") return "OUT";
    return d ? d.toUpperCase() : "";
  }

  function buildWorkLogLayout_() {
    var sheet = getRawWorkLogSheet_();
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) lastCol = 1;
    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];
    var map = buildHeaderMap_(headers);

    return {
      sheet: sheet,
      headers: headers,
      map: map,
      cols: {
        notes: map[norm_("notes")] || map[norm_("הערות")],
        shiftId: map[norm_("ID דיווח")] || map[norm_("ID משמרת")],
        timestamp: map[norm_("חותמת זמן")],
        employeeId: map[norm_("ID עובד")] || map[norm_("מזהה עובד")],
        employeeName: map[norm_("שם מלא")],
        direction: map[norm_("כניסה / יציאה")],
        fixDate: map[norm_("תיקון תאריך")],
        fixTime: map[norm_("תיקון שעה")],
        jobTypeId: map[norm_("ID סוג עבודה")] || map[norm_("ID סוגי עבודה")],
        jobTypeName: map[norm_("סוג עבודה")] || map[norm_("סוגי עבודה")],
        department: map[norm_("מחלקה")],
        units: map[norm_("כמות היחידות")] || map[norm_("דיווח יחידות")],
      },
    };
  }

  function normalizeId_(val) {
    return String(val || "").trim();
  }

  function toDateOnly_(val) {
    var d = toDate_(val);
    if (!d) return null;
    d.setHours(0, 0, 0, 0);
    return d;
  }

  function addDays_(date, offset) {
    if (!date) return null;
    var d = new Date(date.getTime());
    d.setDate(d.getDate() + offset);
    return d;
  }

  function toWorkDateTime_(ts, fixDate, fixTime) {
    var tsDate = toDate_(ts);
    if (tsDate) return tsDate;
    var dt = buildDateTime_(fixDate, fixTime);
    return dt;
  }

  function readWorkLogsForEmployeeJobAndRange_(
    employeeId,
    jobTypeId,
    fromIso,
    toIso
  ) {
    var layout = buildWorkLogLayout_();
    var sheet = layout.sheet;
    var cols = layout.cols;
    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return [];
    var data = sheet
      .getRange(
        CONFIG.HEADER_ROW + 1,
        1,
        lastRow - CONFIG.HEADER_ROW,
        sheet.getLastColumn()
      )
      .getValues();
    var out = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var empIdVal = cols.employeeId ? norm_(row[cols.employeeId - 1]) : "";
      if (!empIdVal || empIdVal !== norm_(employeeId)) continue;

      var jobIdVal = cols.jobTypeId ? norm_(row[cols.jobTypeId - 1]) : "";
      if (!jobIdVal || jobIdVal !== norm_(jobTypeId)) continue;

      var workDateTime = toWorkDateTime_(
        cols.timestamp ? row[cols.timestamp - 1] : "",
        cols.fixDate ? row[cols.fixDate - 1] : "",
        cols.fixTime ? row[cols.fixTime - 1] : ""
      );
      if (!workDateTime) continue;

      var isoDate = toIsoDate_(workDateTime);
      if (fromIso && isoDate && isoDate < fromIso) continue;
      if (toIso && isoDate && isoDate > toIso) continue;

      out.push({
        reportId: cols.shiftId
          ? String(row[cols.shiftId - 1] || "").trim()
          : "",
        employeeId: empIdVal,
        employeeName: cols.employeeName
          ? String(row[cols.employeeName - 1] || "").trim()
          : "",
        jobTypeId: jobIdVal,
        jobTypeName: cols.jobTypeName
          ? String(row[cols.jobTypeName - 1] || "").trim()
          : "",
        department: cols.department
          ? String(row[cols.department - 1] || "").trim()
          : "",
        direction: normalizeDirection_(
          cols.direction ? row[cols.direction - 1] : ""
        ),
        workDateTime: workDateTime,
        workDateIso: isoDate,
        notesKind: cols.notes ? String(row[cols.notes - 1] || "").trim() : "",
        units: cols.units ? row[cols.units - 1] : "",
      });
    }

    out.sort(function (a, b) {
      if (
        a.workDateTime &&
        b.workDateTime &&
        a.workDateTime.getTime() !== b.workDateTime.getTime()
      )
        return a.workDateTime.getTime() - b.workDateTime.getTime();
      return String(a.reportId || "").localeCompare(String(b.reportId || ""));
    });

    return out;
  }

  function ensureSpanAndPayFormats_(sheet, headerMap) {
    var spanCol = headerMap["שעות"];
    var payCol = headerMap["שעות לשכר"];
    var rows = Math.max(1, sheet.getMaxRows() - CONFIG.HEADER_ROW);
    if (spanCol !== undefined && spanCol !== null)
      sheet
        .getRange(CONFIG.HEADER_ROW + 1, spanCol + 1, rows, 1)
        .setNumberFormat("0.00");
    if (payCol !== undefined && payCol !== null)
      sheet
        .getRange(CONFIG.HEADER_ROW + 1, payCol + 1, rows, 1)
        .setNumberFormat("0.00");
  }

  function getOrCreateShiftsSheet_() {
    var ctx = getShiftsSheetAndHeaderMap_();
    var sh = ctx.sheet;
    var map = ctx.headerMap;
    ensureSpanAndPayFormats_(sh, map);
    return sh;
  }

  function getShiftsHeaderMap_() {
    var ctx = getShiftsSheetAndHeaderMap_();
    return ctx.headerMap;
  }

  function norm_(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  function normKey_(v) {
    return norm_(v).toLowerCase();
  }

  function normalizeBool_(v) {
    if (v === true) return true;
    if (v === false) return false;
    var s = normKey_(v);
    if (!s) return false;
    return s === "true" || s === "yes" || s === "כן" || s === "y" || s === "1";
  }

  function buildHeaderMap_(headers) {
    var map = {};
    if (!headers || !headers.length) return map;
    for (var i = 0; i < headers.length; i++) {
      var key = norm_(headers[i]);
      if (!key) continue;
      map[key] = i + 1; // 1-based
    }
    return map;
  }

  function pickCol_(map, cfg) {
    if (!cfg) return null;
    var names = cfg.candidates || [];
    for (var i = 0; i < names.length; i++) {
      var key = norm_(names[i]);
      if (map[key]) return map[key];
    }
    return cfg.fallbackIndex || null;
  }

  var WRITE_ALLOWED_COLS = {
    bonusIds: true,
    manualEdited: true,
    lastUpdatedBySidebar: true,
    lastUpdatedAt: true,
  };

  function hasHeader_(map, cfg) {
    if (!cfg || !cfg.candidates) return false;
    for (var i = 0; i < cfg.candidates.length; i++) {
      var k = norm_(cfg.candidates[i]);
      if (map[k]) return true;
    }
    return false;
  }

  function ensureColumnsForWrite_(sheet, headerMap, keys, allowlist) {
    // Do not mutate sheet structure; just return current header map.
    return headerMap;
  }

  function toIsoDate_(val) {
    if (!val) return "";
    if (
      Object.prototype.toString.call(val) === "[object Date]" &&
      !isNaN(val.getTime())
    ) {
      var y = val.getFullYear();
      var m = ("0" + (val.getMonth() + 1)).slice(-2);
      var d = ("0" + val.getDate()).slice(-2);
      return y + "-" + m + "-" + d;
    }

    var s = norm_(val);
    if (!s) return "";

    // eslint-disable-next-line no-useless-escape
    var mIso = s.match(/^(\d{4})[\/.\-](\d{2})[\/.\-](\d{2})$/);
    if (mIso) return mIso[1] + "-" + mIso[2] + "-" + mIso[3];

    // eslint-disable-next-line no-useless-escape
    var mDmy = s.match(/^(\d{2})[\/.\-](\d{2})[\/.\-](\d{4})$/);
    if (mDmy) return mDmy[3] + "-" + mDmy[2] + "-" + mDmy[1];

    var parsedDate = new Date(s);
    if (!isNaN(parsedDate.getTime())) return toIsoDate_(parsedDate);
    return s;
  }

  function toTimeStr_(val) {
    if (!val && val !== 0) return "";
    if (
      Object.prototype.toString.call(val) === "[object Date]" &&
      !isNaN(val.getTime())
    ) {
      var hh = ("0" + val.getHours()).slice(-2);
      var mm = ("0" + val.getMinutes()).slice(-2);
      return hh + ":" + mm;
    }
    var s = norm_(val);
    var m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) {
      var h = ("0" + m[1]).slice(-2);
      var mi = m[2];
      return h + ":" + mi;
    }
    return s;
  }

  function splitBonusIds_(raw) {
    var out = [];
    if (!raw && raw !== 0) return out;
    var str = Array.isArray(raw)
      ? raw.join(CONFIG.BONUS_SEPARATOR)
      : String(raw);
    var parts = str.split(/[;,]/);
    var seen = {};
    for (var i = 0; i < parts.length; i++) {
      var p = norm_(parts[i]);
      if (!p) continue;
      if (seen[p]) continue;
      seen[p] = true;
      out.push(p);
    }
    return out;
  }

  function toDate_(val) {
    if (!val) return null;
    if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val))
      return val;
    var s = norm_(val);
    if (!s) return null;
    var d = new Date(s);
    return isNaN(d) ? null : d;
  }

  function buildDateTime_(dateVal, timeStr) {
    var d = toDate_(dateVal);
    if (!d) return null;
    var parts = norm_(timeStr || "").split(":");
    if (parts.length >= 2) {
      var hh = parseInt(parts[0], 10);
      var mm = parseInt(parts[1], 10);
      if (!isNaN(hh)) d.setHours(hh);
      if (!isNaN(mm)) d.setMinutes(mm);
      d.setSeconds(0);
      d.setMilliseconds(0);
    }
    return d;
  }

  function computeHoursDiff_(startDate, endDate) {
    if (!startDate || !endDate) return 0;
    var ms = endDate.getTime() - startDate.getTime();
    if (ms <= 0) return 0;
    var hours = ms / (1000 * 60 * 60);
    return Math.round(hours * 100) / 100;
  }

  function appendNote_(base, addition) {
    if (!addition) return base || "";
    if (!base) return addition;
    return base + " | " + addition;
  }

  function deriveDate_(row, layout) {
    var workDate = layout.workDate ? row[layout.workDate - 1] : "";
    var fixDate = layout.fixDate ? row[layout.fixDate - 1] : "";
    var ts = layout.timestamp ? row[layout.timestamp - 1] : "";
    return toIsoDate_(workDate || fixDate || ts);
  }

  function computeHasIssues_(status, notes) {
    var s = normKey_(status);
    if (
      s &&
      (s.indexOf("בעיה") !== -1 ||
        s.indexOf("issue") !== -1 ||
        s.indexOf("problem") !== -1)
    )
      return true;
    var n = normKey_(notes);
    if (
      n &&
      (n.indexOf("issue") !== -1 ||
        n.indexOf("בעיה") !== -1 ||
        n.indexOf("fix") !== -1)
    )
      return true;
    return false;
  }

  function translateShiftStatusAndNote_(status, note) {
    var statusMap = {
      OK: "תקין",
      MISSING_IN: "חסרה כניסה",
      MISSING_OUT: "חסרה יציאה",
      CONFLICT: "תקלה",
      MISSING_ENTRY: "חסרה כניסה",
      MISSING_EXIT: "חסרה יציאה",
    };

    var noteMap = {
      "Extra IN before OUT": "כניסה נוספת לפני יציאה",
      "Missing IN": "חסרה כניסה",
      "Missing OUT": "חסרה יציאה",
      "Unknown direction": "כיוון דיווח לא מזוהה",
      "issue: end time before or equal to start time":
        "שעת יציאה לפני או שווה לשעת כניסה",
      "issue: overlap with another shift for this employee/job":
        "חפיפה עם משמרת אחרת לאותו עובד/תפקיד",
      "issue: shift still open at end of range (missing exit)":
        "משמרת נשארה פתוחה ללא יציאה בסוף הטווח",
      "issue: OUT without matching IN (missing entry)": "יציאה ללא כניסה תואמת",
      "issue: second IN without matching OUT (missing exit)":
        "כניסה נוספת בלי יציאה קודמת",
    };

    return {
      statusDisplay: statusMap[status] || status || "",
      noteDisplay: noteMap[note] || note || "",
    };
  }

  function getLayout_(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) lastCol = 1;
    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];
    var map = buildHeaderMap_(headers);

    var layout = {};
    Object.keys(CONFIG.COLS).forEach(function (k) {
      layout[k] = pickCol_(map, CONFIG.COLS[k]);
    });

    return { map: map, layout: layout, lastCol: lastCol };
  }

  function readShiftRow_(row, layout) {
    var shift = {
      shiftId: layout.shiftId ? norm_(row[layout.shiftId - 1]) : "",
      employeeId: layout.employeeId ? norm_(row[layout.employeeId - 1]) : "",
      employeeName: layout.employeeName
        ? norm_(row[layout.employeeName - 1])
        : "",
      email: layout.email ? norm_(row[layout.email - 1]) : "",
      date: deriveDate_(row, layout),
      timestamp: layout.timestamp ? row[layout.timestamp - 1] : "",
      startTime: layout.startTime ? toTimeStr_(row[layout.startTime - 1]) : "",
      endTime: layout.endTime ? toTimeStr_(row[layout.endTime - 1]) : "",
      status: layout.status ? norm_(row[layout.status - 1]) : "",
      workType: layout.jobName ? norm_(row[layout.jobName - 1]) : "",
      workTypeId: layout.jobId ? norm_(row[layout.jobId - 1]) : "",
      department: layout.department ? norm_(row[layout.department - 1]) : "",
      direction: layout.direction ? norm_(row[layout.direction - 1]) : "",
      units: layout.units ? row[layout.units - 1] : "",
      bonusIds: layout.bonusIds ? splitBonusIds_(row[layout.bonusIds - 1]) : [],
      manualNote: layout.manualNote ? norm_(row[layout.manualNote - 1]) : "",
      hasIssues: computeHasIssues_(
        layout.status ? row[layout.status - 1] : "",
        layout.notes ? row[layout.notes - 1] : ""
      ),
      isManualEdited: layout.manualEdited
        ? normalizeBool_(row[layout.manualEdited - 1])
        : false,
      rawRow: null,
    };

    if (layout.notes) shift.notes = row[layout.notes - 1];

    return shift;
  }

  function buildShiftsForEmployeeJobAndDateRange(
    employeeId,
    jobTypeId,
    fromDate,
    toDate
  ) {
    var fromIso = toIsoDate_(fromDate || "");
    var toIso = toIsoDate_(toDate || "");
    var events = readWorkLogsForEmployeeJobAndRange_(
      employeeId,
      jobTypeId,
      fromIso,
      toIso
    );

    var shifts = [];
    var openIn = null;

    function pushShift_(opts) {
      var startDt = opts.startDateTime;
      var endDt = opts.endDateTime;
      var spanHours = computeHoursDiff_(startDt, endDt);
      var payHours = spanHours;
      var note = opts.note || "";
      var status = opts.status || "OK";

      // Detect overlap later; keep basic shape here.
      shifts.push({
        shiftId: opts.shiftId || "",
        employeeId: opts.employeeId || "",
        employeeName: opts.employeeName || "",
        jobTypeId: opts.jobTypeId || "",
        jobTypeName: opts.jobTypeName || "",
        department: opts.department || "",
        shiftDate: startDt ? toIsoDate_(startDt) : "",
        startTime: startDt ? toTimeStr_(startDt) : "",
        endTime: endDt ? toTimeStr_(endDt) : "",
        startDateTime: startDt,
        endDateTime: endDt,
        spanHours: spanHours,
        payHours: payHours,
        status: status,
        note: note,
        sourceReportIds: (opts.sourceReportIds || []).join(","),
      });
    }

    for (var i = 0; i < events.length; i++) {
      var ev = events[i];

      if (ev.direction === "IN") {
        if (!openIn) {
          openIn = ev;
        } else {
          // Second IN without closing the previous one.
          pushShift_({
            shiftId: ev.reportId || openIn.reportId || "",
            employeeId: ev.employeeId,
            employeeName: ev.employeeName,
            jobTypeId: ev.jobTypeId,
            jobTypeName: ev.jobTypeName,
            department: ev.department,
            startDateTime: openIn.workDateTime,
            endDateTime: ev.workDateTime,
            status: "MISSING_EXIT",
            note: appendNote_(
              "issue: second IN without matching OUT (missing exit)",
              ""
            ),
            sourceReportIds: [openIn.reportId, ev.reportId],
          });
          openIn = ev;
        }
        continue;
      }

      if (ev.direction === "OUT") {
        if (openIn) {
          var endBeforeStart =
            ev.workDateTime && openIn.workDateTime
              ? ev.workDateTime.getTime() <= openIn.workDateTime.getTime()
              : false;
          pushShift_({
            shiftId: openIn.reportId || ev.reportId || "",
            employeeId: openIn.employeeId || ev.employeeId,
            employeeName: openIn.employeeName || ev.employeeName,
            jobTypeId: openIn.jobTypeId || ev.jobTypeId,
            jobTypeName: openIn.jobTypeName || ev.jobTypeName,
            department: openIn.department || ev.department,
            startDateTime: openIn.workDateTime,
            endDateTime: ev.workDateTime,
            status: endBeforeStart ? "CONFLICT" : "OK",
            note: endBeforeStart
              ? appendNote_("issue: end time before or equal to start time", "")
              : "",
            sourceReportIds: [openIn.reportId, ev.reportId],
          });
          openIn = null;
        } else {
          // OUT without prior IN
          pushShift_({
            shiftId: ev.reportId || "",
            employeeId: ev.employeeId,
            employeeName: ev.employeeName,
            jobTypeId: ev.jobTypeId,
            jobTypeName: ev.jobTypeName,
            department: ev.department,
            startDateTime: ev.workDateTime,
            endDateTime: ev.workDateTime,
            status: "MISSING_ENTRY",
            note: appendNote_(
              "issue: OUT without matching IN (missing entry)",
              ""
            ),
            sourceReportIds: [ev.reportId],
          });
        }
      }
    }

    if (openIn) {
      pushShift_({
        shiftId: openIn.reportId || "",
        employeeId: openIn.employeeId,
        employeeName: openIn.employeeName,
        jobTypeId: openIn.jobTypeId,
        jobTypeName: openIn.jobTypeName,
        department: openIn.department,
        startDateTime: openIn.workDateTime,
        endDateTime: openIn.workDateTime,
        status: "MISSING_EXIT",
        note: appendNote_(
          "issue: shift still open at end of range (missing exit)",
          ""
        ),
        sourceReportIds: [openIn.reportId],
      });
    }

    // Detect overlaps among produced shifts for this employee+job
    shifts.sort(function (a, b) {
      if (a.startDateTime && b.startDateTime) {
        if (a.startDateTime.getTime() !== b.startDateTime.getTime())
          return a.startDateTime.getTime() - b.startDateTime.getTime();
      }
      return String(a.shiftId || "").localeCompare(String(b.shiftId || ""));
    });

    for (var j = 0; j < shifts.length; j++) {
      var sA = shifts[j];
      var endA = sA.endDateTime;
      for (var k = j + 1; k < shifts.length; k++) {
        var sB = shifts[k];
        if (!sB.startDateTime || !endA) break;
        if (sB.startDateTime.getTime() >= endA.getTime()) break;
        // overlap if startB < endA and startA < endB
        var overlap =
          sB.startDateTime.getTime() < endA.getTime() &&
          sA.startDateTime &&
          sA.startDateTime.getTime() <
            (sB.endDateTime || sB.startDateTime).getTime();
        if (overlap) {
          sA.status = sA.status === "OK" ? "CONFLICT" : sA.status;
          sB.status = sB.status === "OK" ? "CONFLICT" : sB.status;
          sA.note = appendNote_(
            sA.note,
            "issue: overlap with another shift for this employee/job"
          );
          sB.note = appendNote_(
            sB.note,
            "issue: overlap with another shift for this employee/job"
          );
        }
      }
    }

    return shifts;
  }

  function upsertShiftsForEmployeeJobAndRange(
    employeeId,
    jobTypeId,
    fromDate,
    toDate
  ) {
    var logger = getLogger_("SHIFTS_UPSERT_RANGE");
    var startMs = new Date().getTime();
    var empIdNorm = normalizeId_(employeeId);
    var jobIdNorm = normalizeId_(jobTypeId);
    var fromD = toDateOnly_(fromDate);
    var toD = toDateOnly_(toDate);
    var ctx = getShiftsSheetAndHeaderMap_();
    var sheet = ctx.sheet;
    var headerMap = ctx.headerMap;
    var headerRow = CONFIG.HEADER_ROW || 1;
    var lastRow = sheet.getLastRow();
    var lastCol = headerMap.length;

    var empCol = headerMap["מזהה עובד"];
    var jobCol = headerMap["ID סוג עבודה"];
    var dateCol = headerMap["תאריך משמרת"];
    if (empCol === undefined || jobCol === undefined || dateCol === undefined) {
      throw new Error("Shifts header map missing required columns");
    }

    if (lastRow > headerRow) {
      var data = sheet
        .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
        .getValues();
      var rowsToDelete = [];

      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var rEmp = normalizeId_(row[empCol]);
        var rJob = normalizeId_(row[jobCol]);
        if (rEmp !== empIdNorm) continue;
        if (rJob !== jobIdNorm) continue;

        var rDate = toDateOnly_(row[dateCol]);
        if (!rDate) continue;
        if (fromD && rDate.getTime() < fromD.getTime()) continue;
        if (toD && rDate.getTime() > toD.getTime()) continue;

        rowsToDelete.push(headerRow + 1 + i);
      }

      if (rowsToDelete.length) {
        rowsToDelete.sort(function (a, b) {
          return b - a;
        });
        for (var r = 0; r < rowsToDelete.length; r++) {
          sheet.deleteRow(rowsToDelete[r]);
        }
        if (typeof logDuration_ === "function") {
          logDuration_(logger, "shifts.upsert.delete", startMs, {
            deleted: rowsToDelete.length,
            emp: empIdNorm,
            job: jobIdNorm,
          });
        }
      }
    }

    var shifts = buildShiftsForEmployeeJobAndDateRange(
      employeeId,
      jobTypeId,
      fromDate,
      toDate
    );
    if (!shifts || !shifts.length) return;

    var startRow = sheet.getLastRow() + 1;
    var values = [];
    for (var j = 0; j < shifts.length; j++) {
      values[j] = new Array(lastCol).fill("");
    }

    function setByHeader_(rowArr, headerName, val) {
      var colIdx = headerMap[headerName];
      if (colIdx === undefined || colIdx === null) return;
      rowArr[colIdx] = val;
    }

    for (var s = 0; s < shifts.length; s++) {
      var shObj = shifts[s];
      var translated = translateShiftStatusAndNote_(shObj.status, shObj.note);
      var rowArr = values[s];

      setByHeader_(rowArr, "ID משמרת", shObj.shiftId || "");
      setByHeader_(rowArr, "מזהה עובד", shObj.employeeId || "");
      setByHeader_(rowArr, "שם מלא", shObj.employeeName || "");
      setByHeader_(rowArr, "ID סוג עבודה", shObj.jobTypeId || "");
      setByHeader_(rowArr, "סוג עבודה", shObj.jobTypeName || "");
      setByHeader_(rowArr, "מחלקה", shObj.department || "");

      var sd = shObj.shiftDate
        ? toDateOnly_(shObj.shiftDate)
        : toDateOnly_(shObj.startDateTime);
      setByHeader_(rowArr, "תאריך משמרת", sd);

      var st = shObj.startDateTime instanceof Date ? shObj.startDateTime : null;
      var et = shObj.endDateTime instanceof Date ? shObj.endDateTime : null;
      setByHeader_(rowArr, "שעת התחלה", st);
      setByHeader_(rowArr, "שעת סיום", et);

      setByHeader_(rowArr, "שעות", shObj.spanHours || 0);
      setByHeader_(rowArr, "שעות לשכר", shObj.payHours || shObj.spanHours || 0);
      setByHeader_(rowArr, "כמות יחידות", shObj.units);
      setByHeader_(rowArr, "סטטוס משמרת", translated.statusDisplay || "");

      if (headerMap.notesPrimary !== undefined && headerMap.notesPrimary !== null) {
        rowArr[headerMap.notesPrimary] = translated.noteDisplay || "";
      }
      if (
        headerMap.sourcePrimary !== undefined &&
        headerMap.sourcePrimary !== null
      ) {
        rowArr[headerMap.sourcePrimary] = shObj.sourceReportIds || "";
      }
    }

    if (values.length) {
      sheet.getRange(startRow, 1, values.length, lastCol).setValues(values);
      ensureSpanAndPayFormats_(sheet, headerMap);
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "shifts.upsert.write", startMs, {
          written: values.length,
          emp: empIdNorm,
          job: jobIdNorm,
        });
      }
    }
  }

  function SHIFTS_upsertAroundWorkLog_(employeeId, jobTypeId, workDate) {
    var empId = normalizeId_(employeeId);
    var jobId = normalizeId_(jobTypeId);
    if (!empId || !jobId) return;

    var wd = toDateOnly_(workDate);
    if (!wd) return;

    var fromDate = addDays_(wd, -1);
    var toDate = addDays_(wd, 1);

    upsertShiftsForEmployeeJobAndRange(empId, jobId, fromDate, toDate);
  }

  function SHIFTS_rebuildYesterday() {
    var today = toDateOnly_(new Date());
    if (!today) return;
    var yesterday = addDays_(today, -1);
    var headerRow = CONFIG.HEADER_ROW || 1;

    var sheet = getRawWorkLogSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return;

    var lastCol = sheet.getLastColumn();
    var data = sheet
      .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
      .getValues();

    var layout = buildWorkLogLayout_();
    var cols = layout.cols;
    var pairs = {};

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var workDt = toWorkDateTime_(
        cols.timestamp ? row[cols.timestamp - 1] : "",
        cols.fixDate ? row[cols.fixDate - 1] : "",
        cols.fixTime ? row[cols.fixTime - 1] : ""
      );
      var rowDate = toDateOnly_(workDt);
      if (!rowDate) continue;
      if (rowDate.getTime() !== yesterday.getTime()) continue;

      var empId = cols.employeeId ? normalizeId_(row[cols.employeeId - 1]) : "";
      var jobId = cols.jobTypeId ? normalizeId_(row[cols.jobTypeId - 1]) : "";
      if (!empId || !jobId) continue;

      pairs[empId + "|||" + jobId] = { employeeId: empId, jobTypeId: jobId };
    }

    var fromDate = addDays_(yesterday, -1);
    var toDate = addDays_(yesterday, 1);

    var keys = Object.keys(pairs);
    var totalPairs = keys.length;
    if (totalPairs > MAX_YESTERDAY_REBUILD_PAIRS) {
      Logger.log(
        "[SHIFTS_rebuildYesterday][GUARD] yesterday=" +
          toIsoDate_(yesterday) +
          " pairs=" +
          totalPairs +
          " limit=" +
          MAX_YESTERDAY_REBUILD_PAIRS +
          " — skipping upsert"
      );
      return;
    }

    var errorCount = 0;
    for (var k = 0; k < keys.length; k++) {
      var p = pairs[keys[k]];
      try {
        upsertShiftsForEmployeeJobAndRange(
          p.employeeId,
          p.jobTypeId,
          fromDate,
          toDate
        );
      } catch (e) {
        Logger.log(
          "SHIFTS_rebuildYesterday error for emp=" +
            p.employeeId +
            " job=" +
            p.jobTypeId +
            ": " +
            e
        );
        errorCount += 1;
      }
    }

    Logger.log(
      "[SHIFTS_rebuildYesterday][SUMMARY] yesterday=" +
        toIsoDate_(yesterday) +
        " pairs=" +
        totalPairs +
        " errors=" +
        errorCount
    );
  }

  function filterShifts_(list, filters) {
    filters = filters || {};
    var statusesFilter = Array.isArray(filters.statuses)
      ? filters.statuses.map(normKey_).filter(Boolean)
      : [];
    var statusSet = {};
    statusesFilter.forEach(function (s) {
      statusSet[s] = true;
    });

    var dateFrom = toIsoDate_(filters.dateFrom || "");
    var dateTo = toIsoDate_(filters.dateTo || "");
    var onlyWithIssues = !!filters.onlyWithIssues;
    var onlyWithBonuses = !!filters.onlyWithBonuses;
    var onlyManualChanges = !!filters.onlyManualChanges;
    var empId = norm_(filters.employeeId || "");
    var empName = normKey_(filters.employeeName || "");
    var email = normKey_(filters.email || "");
    var search = normKey_(filters.search || "");

    return list.filter(function (s) {
      if (empId && norm_(s.employeeId) !== empId) return false;
      if (empName && normKey_(s.employeeName) !== empName) return false;
      if (email && normKey_(s.email) !== email) return false;
      if (onlyWithIssues && !s.hasIssues) return false;
      if (onlyWithBonuses && (!s.bonusIds || s.bonusIds.length === 0))
        return false;
      if (onlyManualChanges && !s.isManualEdited) return false;
      if (statusesFilter.length) {
        var st = normKey_(s.status);
        if (!statusSet[st]) return false;
      }
      if (search) {
        var blob =
          (s.shiftId || "") +
          " " +
          (s.employeeName || "") +
          " " +
          (s.employeeId || "") +
          " " +
          (s.status || "") +
          " " +
          (s.workType || "") +
          " " +
          (s.department || "");
        if (normKey_(blob).indexOf(search) === -1) return false;
      }
      if (dateFrom && s.date && s.date < dateFrom) return false;
      if (dateTo && s.date && s.date > dateTo) return false;
      return true;
    });
  }

  function list_(filters) {
    var sheet = getSheet_();
    var layoutInfo = getLayout_(sheet);
    var layout = layoutInfo.layout;

    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) {
      return {
        ok: true,
        shifts: [],
        total: 0,
        returned: 0,
        hasMore: false,
        statuses: [],
      };
    }

    var data = sheet
      .getRange(
        CONFIG.HEADER_ROW + 1,
        1,
        lastRow - CONFIG.HEADER_ROW,
        layoutInfo.lastCol
      )
      .getValues();

    var allStatuses = {};
    var allShifts = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var shift = readShiftRow_(row, layout);
      if (!shift.shiftId && !shift.employeeId && !shift.employeeName) continue;
      var stKey = normKey_(shift.status);
      if (stKey) allStatuses[stKey] = shift.status;
      allShifts.push(shift);
    }

    var filtered = filterShifts_(allShifts, filters);

    filtered.sort(function (a, b) {
      if (a.date && b.date && a.date !== b.date)
        return a.date < b.date ? 1 : -1;
      return String(b.shiftId || "").localeCompare(String(a.shiftId || ""));
    });

    var limit = Number(filters && filters.limit);
    if (isNaN(limit) || limit <= 0) limit = 100;
    var offset = Number(filters && filters.offset);
    if (isNaN(offset) || offset < 0) offset = 0;

    var sliced = filtered.slice(offset, offset + limit);
    var hasMore = offset + limit < filtered.length;

    return {
      ok: true,
      shifts: sliced,
      total: filtered.length,
      returned: sliced.length,
      hasMore: hasMore,
      statuses: Object.keys(allStatuses).map(function (k) {
        return allStatuses[k];
      }),
    };
  }

  function findRowByShiftId_(sheet, layout, shiftId) {
    if (!shiftId) return null;
    var targetCol = layout.shiftId || 2;
    var lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return null;
    var range = sheet.getRange(
      CONFIG.HEADER_ROW + 1,
      targetCol,
      lastRow - CONFIG.HEADER_ROW,
      1
    );
    var finder = range.createTextFinder(String(shiftId)).matchEntireCell(true);
    var found = finder ? finder.findNext() : null;
    return found ? found.getRow() : null;
  }

  function get_(shiftId) {
    var sheet = getSheet_();
    var layoutInfo = getLayout_(sheet);
    var layout = layoutInfo.layout;

    var rowIndex = findRowByShiftId_(sheet, layout, shiftId);
    if (!rowIndex) {
      return { ok: false, error: "Shift not found for shiftId=" + shiftId };
    }

    var row = sheet.getRange(rowIndex, 1, 1, layoutInfo.lastCol).getValues()[0];
    var shift = readShiftRow_(row, layout);
    shift.rawRow = rowIndex;

    return { ok: true, shift: JSON.parse(JSON.stringify(shift)) };
  }

  function getSelectedRef_() {
    var ss = ss_();
    var sh = ss.getActiveSheet();
    if (!sh) return { ok: false, error: "לא נמצאה כרטיסייה פעילה" };

    var cell = sh.getActiveCell();
    if (!cell) return { ok: true, sourceSheet: sh.getName(), hasAny: false };

    var row = cell.getRow();
    if (row <= CONFIG.HEADER_ROW) {
      return { ok: true, sourceSheet: sh.getName(), row: row, hasAny: false };
    }

    var headers = sh
      .getRange(CONFIG.HEADER_ROW, 1, 1, sh.getLastColumn())
      .getValues()[0];
    var map = buildHeaderMap_(headers);
    var layout = {
      shiftId: pickCol_(map, CONFIG.COLS.shiftId),
      employeeId: pickCol_(map, CONFIG.COLS.employeeId),
      employeeName: pickCol_(map, CONFIG.COLS.employeeName),
      email: pickCol_(map, CONFIG.COLS.email),
    };

    function val_(col) {
      return col ? sh.getRange(row, col).getValue() : "";
    }

    var res = { ok: true, sourceSheet: sh.getName(), row: row, hasAny: false };
    var sid = norm_(val_(layout.shiftId));
    var eid = norm_(val_(layout.employeeId));
    var ename = norm_(val_(layout.employeeName));
    var email = norm_(val_(layout.email));

    if (sid) {
      res.shiftId = sid;
      res.hasAny = true;
    }
    if (eid) {
      res.employeeId = eid;
      res.hasAny = true;
    }
    if (ename) {
      res.employeeName = ename;
      res.hasAny = true;
    }
    if (email) {
      res.email = email;
      res.hasAny = true;
    }

    return res;
  }

  function updateBonuses_(shiftId, bonusIds) {
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      return { ok: false, error: "לא ניתן לנעול את הגיליון לעדכון בונוסים" };
    }

    try {
      var sheet = getSheet_();
      var layoutInfo = getLayout_(sheet);

      ensureColumnsForWrite_(
        sheet,
        layoutInfo.map,
        ["bonusIds", "manualEdited", "lastUpdatedBySidebar", "lastUpdatedAt"],
        WRITE_ALLOWED_COLS
      );
      layoutInfo = getLayout_(sheet);
      var layout = layoutInfo.layout;

      var rowIndex = findRowByShiftId_(sheet, layout, shiftId);
      if (!rowIndex)
        return { ok: false, error: "Shift not found for shiftId=" + shiftId };

      var uniqueIds = splitBonusIds_(bonusIds);
      var bonusStr = uniqueIds.join(CONFIG.BONUS_SEPARATOR);

      if (layout.bonusIds)
        sheet.getRange(rowIndex, layout.bonusIds).setValue(bonusStr);
      if (layout.manualEdited)
        sheet.getRange(rowIndex, layout.manualEdited).setValue(true);
      if (layout.lastUpdatedBySidebar)
        sheet
          .getRange(rowIndex, layout.lastUpdatedBySidebar)
          .setValue(Session.getActiveUser().getEmail() || "sidebar");
      if (layout.lastUpdatedAt)
        sheet.getRange(rowIndex, layout.lastUpdatedAt).setValue(new Date());

      var updated = get_(shiftId);
      if (updated && updated.ok) {
        return { ok: true, shift: updated.shift, bonusIds: uniqueIds };
      }
      return { ok: true, bonusIds: uniqueIds };
    } finally {
      try {
        lock.releaseLock();
      } catch (_e) {
        /* ignore lock release failures */
      }
    }
  }

  function updateShift_(payload) {
    if (!payload || !payload.shiftId) {
      return { ok: false, error: "shiftId is required" };
    }

    // TODO: verify with shift builder script once finalized to avoid overwriting calculated fields.

    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) {
      return { ok: false, error: "לא ניתן לנעול את הגיליון לעדכון משמרת" };
    }

    try {
      var ctx = getShiftsSheetAndHeaderMap_();
      var sheet = ctx.sheet;
      var headerMap = ctx.headerMap;

      var shiftIdCol = headerMap["ID משמרת"];
      if (shiftIdCol === undefined || shiftIdCol === null) {
        return { ok: false, error: "ShiftId column missing" };
      }

      var headerRow = CONFIG.HEADER_ROW || 1;
      var lastRow = sheet.getLastRow();
      if (lastRow <= headerRow) {
        return { ok: false, error: "No shift rows found" };
      }

      var colRange = sheet.getRange(
        headerRow + 1,
        shiftIdCol + 1,
        lastRow - headerRow,
        1
      );
      var colValues = colRange.getValues();
      var rowIndex = null;
      for (var i = 0; i < colValues.length; i++) {
        if (stringValue(colValues[i][0]) === stringValue(payload.shiftId)) {
          rowIndex = headerRow + 1 + i;
          break;
        }
      }

      if (!rowIndex)
        return {
          ok: false,
          error: "Shift not found for shiftId=" + payload.shiftId,
        };

      function setByHeader_(headerName, value) {
        var idx = headerMap[headerName];
        if (idx === undefined || idx === null) return;
        sheet.getRange(rowIndex, idx + 1).setValue(value);
      }

      if (payload.status !== undefined) {
        setByHeader_("סטטוס משמרת", payload.status);
      }
      if (payload.workType !== undefined) {
        setByHeader_("סוג עבודה", payload.workType);
      }
      if (payload.workTypeId !== undefined) {
        setByHeader_("ID סוג עבודה", payload.workTypeId);
      }
      if (payload.department !== undefined) {
        setByHeader_("מחלקה", payload.department);
      }
      if (payload.date !== undefined) {
        setByHeader_("תאריך משמרת", toIsoDate_(payload.date));
      }
      if (payload.startTime !== undefined) {
        setByHeader_("שעת התחלה", toTimeStr_(payload.startTime));
      }
      if (payload.endTime !== undefined) {
        setByHeader_("שעת סיום", toTimeStr_(payload.endTime));
      }
      if (payload.manualNote !== undefined) {
        if (
          headerMap.notesPrimary !== undefined &&
          headerMap.notesPrimary !== null
        ) {
          sheet
            .getRange(rowIndex, headerMap.notesPrimary + 1)
            .setValue(payload.manualNote);
        }
      }

      var updated = get_(payload.shiftId);
      if (updated && updated.ok) return updated;
      return { ok: true };
    } finally {
      try {
        lock.releaseLock();
      } catch (_e) {
        /* ignore lock release failures */
      }
    }
  }

  SHIFTS.list = function (filters) {
    var logger = getLogger_("SHIFTS_LIST");
    var startMs = new Date().getTime();
    try {
      var res = list_(filters || {});
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "shifts.list", startMs, {
          total: res && res.total,
          returned: res && res.returned,
        });
      }
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "shifts.list",
          startMs,
          { error: String(err) },
          "warn",
          "SHIFTS_LIST_FAIL",
          err
        );
      }
      Logger.log("SHIFTS.list error: " + err);
      return { ok: false, error: "SHIFTS.list failed: " + err };
    }
  };

  SHIFTS.get = function (shiftId) {
    var logger = getLogger_("SHIFTS_GET");
    var startMs = new Date().getTime();
    try {
      var res = get_(shiftId);
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "shifts.get", startMs, {
          shiftId: shiftId || "",
          ok: res && res.ok,
        });
      }
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "shifts.get",
          startMs,
          { shiftId: shiftId || "", error: String(err) },
          "warn",
          "SHIFTS_GET_FAIL",
          err
        );
      }
      Logger.log("SHIFTS.get error: " + err);
      return { ok: false, error: "SHIFTS.get failed: " + err };
    }
  };

  SHIFTS.getSelectedRef = function () {
    var logger = getLogger_("SHIFTS_GET_SELECTED");
    var startMs = new Date().getTime();
    try {
      var res = getSelectedRef_();
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "shifts.getSelectedRef", startMs, {
          sourceSheet: res && res.sourceSheet,
          hasAny: res && res.hasAny,
        });
      }
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "shifts.getSelectedRef",
          startMs,
          { error: String(err) },
          "warn",
          "SHIFTS_GET_SELECTED_FAIL",
          err
        );
      }
      Logger.log("SHIFTS.getSelectedRef error: " + err);
      return { ok: false, error: "SHIFTS.getSelectedRef failed: " + err };
    }
  };

  SHIFTS.updateBonuses = function (shiftId, bonusIds) {
    var logger = getLogger_("SHIFTS_UPDATE_BONUSES");
    var startMs = new Date().getTime();
    try {
      var res = updateBonuses_(shiftId, bonusIds || []);
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "shifts.updateBonuses", startMs, {
          shiftId: shiftId || "",
          bonuses: bonusIds ? bonusIds.length : 0,
          ok: res && res.ok,
        });
      }
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "shifts.updateBonuses",
          startMs,
          { shiftId: shiftId || "", error: String(err) },
          "warn",
          "SHIFTS_UPDATE_BONUSES_FAIL",
          err
        );
      }
      Logger.log("SHIFTS.updateBonuses error: " + err);
      return { ok: false, error: "SHIFTS.updateBonuses failed: " + err };
    }
  };

  SHIFTS.updateShift = function (payload) {
    var logger = getLogger_("SHIFTS_UPDATE");
    var startMs = new Date().getTime();
    try {
      var res = updateShift_(payload || {});
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "shifts.update", startMs, {
          shiftId: payload && payload.shiftId,
          ok: res && res.ok,
        });
      }
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "shifts.update",
          startMs,
          { shiftId: payload && payload.shiftId, error: String(err) },
          "warn",
          "SHIFTS_UPDATE_FAIL",
          err
        );
      }
      Logger.log("SHIFTS.updateShift error: " + err);
      return { ok: false, error: "SHIFTS.updateShift failed: " + err };
    }
  };

  // ---- Bonus catalog ----
  var BONUS_CONFIG = {
    SHEET_NAME_OPTIONS: "אופציות בחירה ו ID'S",
    HEADER_ROW: 1,
  };

  function getBonusSheet_() {
    var sh = ss_().getSheetByName(BONUS_CONFIG.SHEET_NAME_OPTIONS);
    if (!sh)
      throw new Error(
        'לא נמצאה כרטיסייה בשם "' + BONUS_CONFIG.SHEET_NAME_OPTIONS + '"'
      );
    return sh;
  }

  function bonusCols_(map) {
    return {
      status: map["סטטוס בונוס"] || map["סטטוס"] || map["BONUS_STATUS"] || 0,
      id: map["ID בונוס"] || map["BONUS_ID"] || map["Bonus Id"] || 0,
      name: map["סוג בונוס"] || map["BONUS_NAME"] || map["Bonus Name"] || 0,
      description:
        map["תיאור בונוס"] ||
        map["Bonus Description"] ||
        map["Description"] ||
        0,
      amount: map["סכום"] || map["Amount"] || map["Bonus Amount"] || 0,
      type: map["סוג"] || map["Type"] || map["Bonus Type"] || 0,
    };
  }

  function listAllBonuses_() {
    var sh = getBonusSheet_();

    // נסה לוודא IDs דרך מודול OPT אם קיים
    try {
      if (typeof OPT !== "undefined" && OPT.ensureCatalogIds) {
        OPT.ensureCatalogIds(sh);
      }
    } catch (_e) {
      /* ignore catalog ID enforcement failures */
    }

    var lastRow = sh.getLastRow();
    if (lastRow <= BONUS_CONFIG.HEADER_ROW) return { ok: true, bonuses: [] };

    var lastCol = sh.getLastColumn();
    var headers = sh
      .getRange(BONUS_CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];
    var map = buildHeaderMap_(headers);
    var cols = bonusCols_(map);

    var data = sh
      .getRange(
        BONUS_CONFIG.HEADER_ROW + 1,
        1,
        lastRow - BONUS_CONFIG.HEADER_ROW,
        lastCol
      )
      .getValues();
    var res = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var bonusId = cols.id ? norm_(row[cols.id - 1]) : "";
      var name = cols.name ? norm_(row[cols.name - 1]) : "";
      if (!bonusId && !name) continue;
      var status = cols.status ? norm_(row[cols.status - 1]) : "";
      var bonus = {
        bonusId: bonusId,
        name: name,
        description: cols.description ? norm_(row[cols.description - 1]) : "",
        amount: cols.amount ? row[cols.amount - 1] : "",
        type: cols.type ? norm_(row[cols.type - 1]) : "",
        status: status,
      };
      res.push(bonus);
    }

    res.sort(function (a, b) {
      return String(a.name || "").localeCompare(String(b.name || ""));
    });

    return { ok: true, bonuses: res };
  }

  BONUSES.listAll = function () {
    var logger = getLogger_("BONUSES_LIST_ALL");
    var startMs = new Date().getTime();
    try {
      var res = listAllBonuses_();
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "bonuses.listAll", startMs, {
          bonuses: res && res.bonuses ? res.bonuses.length : 0,
        });
      }
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "bonuses.listAll",
          startMs,
          { error: String(err) },
          "warn",
          "BONUSES_LIST_FAIL",
          err
        );
      }
      Logger.log("BONUSES.listAll error: " + err);
      return { ok: false, error: "BONUSES.listAll failed: " + err };
    }
  };
})();

// ---- Global wrappers for google.script.run ----
// eslint-disable-next-line no-unused-vars
function SHIFTS_list(filters) {
  if (typeof SHIFTS !== "undefined" && SHIFTS.list) {
    return SHIFTS.list(filters || {});
  }
  return { ok: false, error: "SHIFTS module missing" };
}

// eslint-disable-next-line no-unused-vars
function SHIFTS_get(shiftId) {
  if (typeof SHIFTS !== "undefined" && SHIFTS.get) {
    return SHIFTS.get(shiftId);
  }
  return { ok: false, error: "SHIFTS module missing" };
}

// eslint-disable-next-line no-unused-vars
function SHIFTS_getSelectedRef() {
  if (typeof SHIFTS !== "undefined" && SHIFTS.getSelectedRef) {
    return SHIFTS.getSelectedRef();
  }
  return { ok: false, error: "SHIFTS module missing" };
}

// eslint-disable-next-line no-unused-vars
function SHIFTS_updateShift(payload) {
  if (typeof SHIFTS !== "undefined" && SHIFTS.updateShift) {
    return SHIFTS.updateShift(payload || {});
  }
  return { ok: false, error: "SHIFTS module missing" };
}

// eslint-disable-next-line no-unused-vars
function SHIFTS_updateBonuses(shiftId, bonusIds) {
  if (typeof SHIFTS !== "undefined" && SHIFTS.updateBonuses) {
    return SHIFTS.updateBonuses(shiftId, bonusIds || []);
  }
  return { ok: false, error: "SHIFTS module missing" };
}

// eslint-disable-next-line no-unused-vars
function BONUSES_listAll() {
  if (typeof BONUSES !== "undefined" && BONUSES.listAll) {
    return BONUSES.listAll();
  }
  return { ok: false, error: "BONUSES module missing" };
}
