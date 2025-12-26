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
        candidates: ["תאריך משמרת", "Work Date"],
        fallbackIndex: null,
      },
      startTime: {
        candidates: ["שעת התחלה", "Start Time", "שעת כניסה"],
        fallbackIndex: null,
      },
      endTime: {
        candidates: ["שעת סיום", "End Time", "שעת יציאה"],
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

  function ss_() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getSheet_() {
    var sh = ss_().getSheetByName(CONFIG.SHEET_NAME);
    if (!sh)
      throw new Error('לא נמצאה כרטיסייה בשם "' + CONFIG.SHEET_NAME + '"');
    return sh;
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
      var sheet = getSheet_();
      var layoutInfo = getLayout_(sheet);

      ensureColumnsForWrite_(
        sheet,
        layoutInfo.map,
        [
          "bonusIds",
          "manualEdited",
          "manualNote",
          "lastUpdatedBySidebar",
          "lastUpdatedAt",
        ],
        WRITE_ALLOWED_COLS
      );

      layoutInfo = getLayout_(sheet);
      var layout = layoutInfo.layout;

      var rowIndex = findRowByShiftId_(sheet, layout, payload.shiftId);
      if (!rowIndex)
        return {
          ok: false,
          error: "Shift not found for shiftId=" + payload.shiftId,
        };

      if (
        payload.status !== undefined &&
        layout.status &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.status)
      ) {
        sheet.getRange(rowIndex, layout.status).setValue(payload.status);
      }
      if (
        payload.workType !== undefined &&
        layout.jobName &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.jobName)
      ) {
        sheet.getRange(rowIndex, layout.jobName).setValue(payload.workType);
      }
      if (
        payload.workTypeId !== undefined &&
        layout.jobId &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.jobId)
      ) {
        sheet.getRange(rowIndex, layout.jobId).setValue(payload.workTypeId);
      }
      if (
        payload.department !== undefined &&
        layout.department &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.department)
      ) {
        sheet
          .getRange(rowIndex, layout.department)
          .setValue(payload.department);
      }
      if (
        payload.date !== undefined &&
        layout.workDate &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.workDate)
      ) {
        sheet
          .getRange(rowIndex, layout.workDate)
          .setValue(toIsoDate_(payload.date));
      }
      if (
        payload.startTime !== undefined &&
        layout.startTime &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.startTime)
      ) {
        sheet
          .getRange(rowIndex, layout.startTime)
          .setValue(toTimeStr_(payload.startTime));
      }
      if (
        payload.endTime !== undefined &&
        layout.endTime &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.endTime)
      ) {
        sheet
          .getRange(rowIndex, layout.endTime)
          .setValue(toTimeStr_(payload.endTime));
      }
      if (
        payload.bonusIds !== undefined &&
        layout.bonusIds &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.bonusIds)
      ) {
        var ids = splitBonusIds_(payload.bonusIds);
        sheet
          .getRange(rowIndex, layout.bonusIds)
          .setValue(ids.join(CONFIG.BONUS_SEPARATOR));
      }
      if (
        payload.manualNote !== undefined &&
        layout.manualNote &&
        hasHeader_(layoutInfo.map, CONFIG.COLS.manualNote)
      ) {
        sheet
          .getRange(rowIndex, layout.manualNote)
          .setValue(payload.manualNote);
      }

      if (layout.manualEdited)
        sheet.getRange(rowIndex, layout.manualEdited).setValue(true);
      if (layout.lastUpdatedBySidebar)
        sheet
          .getRange(rowIndex, layout.lastUpdatedBySidebar)
          .setValue(Session.getActiveUser().getEmail() || "sidebar");
      if (layout.lastUpdatedAt)
        sheet.getRange(rowIndex, layout.lastUpdatedAt).setValue(new Date());

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
    try {
      var res = list_(filters || {});
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      Logger.log("SHIFTS.list error: " + err);
      return { ok: false, error: "SHIFTS.list failed: " + err };
    }
  };

  SHIFTS.get = function (shiftId) {
    try {
      var res = get_(shiftId);
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      Logger.log("SHIFTS.get error: " + err);
      return { ok: false, error: "SHIFTS.get failed: " + err };
    }
  };

  SHIFTS.getSelectedRef = function () {
    try {
      var res = getSelectedRef_();
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      Logger.log("SHIFTS.getSelectedRef error: " + err);
      return { ok: false, error: "SHIFTS.getSelectedRef failed: " + err };
    }
  };

  SHIFTS.updateBonuses = function (shiftId, bonusIds) {
    try {
      var res = updateBonuses_(shiftId, bonusIds || []);
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
      Logger.log("SHIFTS.updateBonuses error: " + err);
      return { ok: false, error: "SHIFTS.updateBonuses failed: " + err };
    }
  };

  SHIFTS.updateShift = function (payload) {
    try {
      var res = updateShift_(payload || {});
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
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
    try {
      var res = listAllBonuses_();
      return JSON.parse(JSON.stringify(res));
    } catch (err) {
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
