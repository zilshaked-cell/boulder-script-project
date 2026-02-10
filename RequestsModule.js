var REQ = REQ || {};

(function () {
  "use strict";

  var CONFIG = {
    SHEET_NAME_REQUESTS: "בקשות עובדים",
    HEADER_ROW: 1,
    COLS: {
      requestId: { candidates: ["ID בקשה", "Request ID", "RequestID"] },
      status: { candidates: ["סטטוס בקשה", "Status", "סטטוס"] },
      timestamp: { candidates: ["חותמת זמן", "Timestamp", "Created", "Time"] },
      employeeId: { candidates: ["ID עובד", "Employee ID"] },
      employeeName: { candidates: ["שם מלא", "Employee Name", "Name"] },
      shiftId: { candidates: ["ID משמרת", "Shift ID"] },
      direction: { candidates: ["כניסה / יציאה", "כיוון", "Direction"] },
      fixDate: { candidates: ["תיקון תאריך", "תאריך", "Date"] },
      fixTime: { candidates: ["תיקון שעה", "שעה", "Time Fix"] },
      jobId: { candidates: ["ID סוג עבודה", "ID סוגי עבודה", "Job ID"] },
      jobName: {
        candidates: ["סוג עבודה", "סוגי עבודה", "סוג העבודה", "Job Name"],
      },
      department: { candidates: ["מחלקה", "Department"] },
      units: { candidates: ["כמות היחידות", "יחידות", "Units"] },
      note: { candidates: ["הערות למשמרת", "הערות", "Notes", "Note"] },
    },
  };

  function ss_() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function sh_() {
    var sh = ss_().getSheetByName(CONFIG.SHEET_NAME_REQUESTS);
    if (!sh)
      throw new Error('לא נמצאה כרטיסייה "' + CONFIG.SHEET_NAME_REQUESTS + '"');
    return sh;
  }

  function uuid_() {
    return Utilities.getUuid();
  }

  function getLogger_(operation) {
    try {
      if (typeof ensureModuleLoggerDefined_ === "function") {
        return ensureModuleLoggerDefined_(operation || "REQUESTS_MODULE");
      }
    } catch (_ignored) {
      // Intentionally silent: logger not available, continue without logging
    }
    return null;
  }

  function norm_(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  function normalizeStatuses_(statuses, knownSet, logger, startMs) {
    var cleaned = [];
    var removed = [];

    if (!Array.isArray(statuses)) {
      return { normalized: cleaned, removed: removed };
    }

    for (var i = 0; i < statuses.length; i++) {
      var key = norm_(statuses[i]).toLowerCase();
      if (!key) continue;
      if (knownSet && !knownSet[key]) {
        if (removed.indexOf(key) === -1) removed.push(key);
        continue;
      }
      if (cleaned.indexOf(key) === -1) cleaned.push(key);
    }

    if (removed.length && typeof logDuration_ === "function") {
      logDuration_(
        logger,
        "requests.list",
        startMs || new Date().getTime(),
        {
          removedUnknownStatuses: removed.slice(0, 10),
          removedCount: removed.length,
        },
        "warn",
        "REQUEST_STATUS_UNKNOWN",
      );
    }

    return { normalized: cleaned, removed: removed };
  }

  function headersMap_(sheet) {
    var sh = sheet || sh_();
    var lastCol = sh.getLastColumn();
    if (lastCol < 1) return {};
    var headers = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var h = norm_(headers[i]);
      if (!h) continue;
      map[h] = i + 1; // 1-based
    }
    return map;
  }

  function col_(map, header) {
    return map[header] || null;
  }

  function pickCol_(map, cfg) {
    if (!cfg || !cfg.candidates) return null;
    for (var i = 0; i < cfg.candidates.length; i++) {
      var h = norm_(cfg.candidates[i]);
      if (map[h]) return map[h];
    }
    return null;
  }

  function readRow_(row, map) {
    if (!row) return null;
    function v(col) {
      return col ? norm_(row[col - 1]) : "";
    }

    var rec = {
      requestId: v(map.requestId),
      status: v(map.status),
      timestamp: map.timestamp ? row[map.timestamp - 1] : "",
      employeeId: v(map.employeeId),
      employeeName: v(map.employeeName),
      shiftId: v(map.shiftId),
      direction: v(map.direction),
      fixDate: v(map.fixDate),
      fixTime: v(map.fixTime),
      jobId: v(map.jobId),
      jobName: v(map.jobName),
      department: v(map.department),
      units: map.units ? row[map.units - 1] : "",
      note: map.note ? row[map.note - 1] : "",
    };

    if (rec.timestamp instanceof Date && !isNaN(rec.timestamp.getTime())) {
      rec.timestampIso = rec.timestamp.toISOString();
    } else if (rec.timestamp) {
      rec.timestampIso = String(rec.timestamp);
    } else {
      rec.timestampIso = "";
    }

    return rec;
  }

  var DEFAULT_OPEN_STATUS = [
    "pending",
    "פתוח",
    "פתוחה",
    "ממתין",
    "ממתינה",
    "open",
  ];

  function list_(filters) {
    var logger = getLogger_("REQUESTS_LIST");
    var startMs = new Date().getTime();
    var sh = sh_();
    var lastRow = sh.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) {
      var emptyRes = {
        ok: true,
        requests: [],
        total: 0,
        returned: 0,
        hasMore: false,
        statuses: [],
      };
      if (typeof logDuration_ === "function") {
        logDuration_(logger, "requests.list", startMs, { returned: 0 });
      }
      return emptyRes;
    }

    var map = headersMap_(sh);
    var cols = {};
    Object.keys(CONFIG.COLS).forEach(function (k) {
      cols[k] = pickCol_(map, CONFIG.COLS[k]);
    });

    var data = sh
      .getRange(
        CONFIG.HEADER_ROW + 1,
        1,
        lastRow - CONFIG.HEADER_ROW,
        sh.getLastColumn(),
      )
      .getValues();

    function toIsoDate_(d) {
      if (!d) return "";
      if (
        Object.prototype.toString.call(d) === "[object Date]" &&
        !isNaN(d.getTime())
      ) {
        var y = d.getFullYear();
        var m = ("0" + (d.getMonth() + 1)).slice(-2);
        var dd = ("0" + d.getDate()).slice(-2);
        return y + "-" + m + "-" + dd;
      }
      return norm_(d);
    }

    var search = norm_(filters && filters.search);
    var defaultOpenNormalized = normalizeStatuses_(
      DEFAULT_OPEN_STATUS,
      null,
      logger,
      startMs,
    ).normalized;
    var defaultOpenSet = {};
    defaultOpenNormalized.forEach(function (s) {
      defaultOpenSet[s] = true;
    });

    var rawStatusesFilter = Array.isArray(filters && filters.statuses)
      ? filters.statuses
          .map(function (s) {
            return norm_(s).toLowerCase();
          })
          .filter(Boolean)
      : [];

    var includeClosed = !!(filters && filters.includeClosed);
    var dateFrom = toIsoDate_(filters && filters.dateFrom);
    var dateTo = toIsoDate_(filters && filters.dateTo);

    var rows = [];
    var statusesFound = {};
    var oldestOpenDate = null;

    for (var i = 0; i < data.length; i++) {
      var rec = readRow_(data[i], cols);
      if (!rec) continue;
      var statusKey = norm_(rec.status).toLowerCase();
      if (statusKey && !statusesFound[statusKey]) {
        statusesFound[statusKey] = rec.status;
      }

      if (defaultOpenSet[statusKey]) {
        var ts =
          rec.timestamp instanceof Date
            ? rec.timestamp
            : rec.timestampIso
              ? new Date(rec.timestampIso)
              : null;
        if (ts instanceof Date && !isNaN(ts.getTime())) {
          if (!oldestOpenDate || ts.getTime() < oldestOpenDate.getTime()) {
            oldestOpenDate = ts;
          }
        }
      }

      rows.push({ rec: rec, statusKey: statusKey });
    }

    var normalizedStatuses = normalizeStatuses_(
      rawStatusesFilter,
      statusesFound,
      logger,
      startMs,
    );
    var statusesFilter = normalizedStatuses.normalized;
    var statusSet = {};
    statusesFilter.forEach(function (s) {
      statusSet[s] = true;
    });

    var all = [];

    for (var j = 0; j < rows.length; j++) {
      var row = rows[j];
      var recRow = row.rec;
      var statusKeyRow = row.statusKey;

      // default filter: only open statuses unless includeClosed or explicit statuses provided
      if (!includeClosed && statusesFilter.length === 0) {
        if (!defaultOpenSet[statusKeyRow]) continue;
      }

      if (statusesFilter.length && !statusSet[statusKeyRow]) continue;

      if (search) {
        var blob = (
          (recRow.requestId || "") +
          " " +
          (recRow.employeeName || "") +
          " " +
          (recRow.employeeId || "") +
          " " +
          (recRow.shiftId || "") +
          " " +
          (recRow.jobName || "") +
          " " +
          (recRow.department || "") +
          " " +
          (recRow.status || "") +
          " " +
          (recRow.note || "")
        ).toLowerCase();
        if (blob.indexOf(search.toLowerCase()) === -1) continue;
      }

      if (dateFrom && recRow.timestampIso && recRow.timestampIso < dateFrom)
        continue;
      if (dateTo && recRow.timestampIso && recRow.timestampIso > dateTo)
        continue;

      all.push(recRow);
    }

    // newest first by timestamp
    all.sort(function (a, b) {
      var at = a.timestamp instanceof Date ? a.timestamp.getTime() : 0;
      var bt = b.timestamp instanceof Date ? b.timestamp.getTime() : 0;
      if (at !== bt) return bt - at;
      return String(b.requestId || "").localeCompare(String(a.requestId || ""));
    });

    var limit = Number(filters && filters.limit);
    if (isNaN(limit) || limit <= 0) limit = 100;
    var offset = Number(filters && filters.offset);
    if (isNaN(offset) || offset < 0) offset = 0;

    var sliced = all.slice(offset, offset + limit);
    var hasMore = offset + limit < all.length;

    var res = {
      ok: true,
      requests: sliced,
      total: all.length,
      returned: sliced.length,
      hasMore: hasMore,
      statuses: Object.keys(statusesFound).map(function (k) {
        return statusesFound[k];
      }),
      oldestOpenTimestampIso: oldestOpenDate
        ? oldestOpenDate.toISOString()
        : "",
    };

    if (typeof logDuration_ === "function") {
      logDuration_(logger, "requests.list", startMs, {
        total: all.length,
        returned: sliced.length,
        offset: offset,
        limit: limit,
      });
    }

    return res;
  }

  /**
   * ממלא UUID בעמודות ID שיש להן כותרת שמכילה "ID" + שיש תוכן בשורה (כלומר לא שורה ריקה).
   * ברירת מחדל: רק עמודת "ID בקשה" (אם קיימת) כדי לא לגעת בשדות אחרים.
   */
  function ensureRequestIds_(sheet) {
    var logger = getLogger_("REQUESTS_ENSURE_IDS");
    var startMs = new Date().getTime();
    var warning = null; // הגדרת משתנה למניעת ReferenceError
    var sh = sheet || sh_();
    var lastRow = sh.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return { filled: 0 };

    var map = headersMap_(sh);

    // אנחנו תומכים בכותרת מדויקת "ID בקשה".
    // אם אין, ננסה למצוא עמודה שהכותרת שלה כוללת "ID" וגם "בקשה".
    var idCol = col_(map, "ID בקשה");
    if (!idCol) {
      // fallback חיפוש כותרת
      var keys = Object.keys(map);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (k.indexOf("ID") !== -1 && k.indexOf("בקשה") !== -1) {
          idCol = map[k];
          break;
        }
      }
    }

    if (!idCol) {
      // אין עמודת ID בקשה - לא עושים כלום
      warning = 'לא נמצאה עמודת "ID בקשה"';
      return { filled: 0, warning: warning };
    }

    var lastCol = sh.getLastColumn();
    var dataRange = sh.getRange(
      CONFIG.HEADER_ROW + 1,
      1,
      lastRow - CONFIG.HEADER_ROW,
      lastCol,
    );
    var data = dataRange.getValues();

    // נגדיר "שורה לא ריקה" ככזו שיש בה משהו חוץ מה-ID עצמו
    var filled = 0;
    var existing = {};
    for (var r = 0; r < data.length; r++) {
      var idv = norm_(data[r][idCol - 1]);
      if (idv) existing[idv] = true;
    }

    for (var r2 = 0; r2 < data.length; r2++) {
      var row = data[r2];
      var currId = norm_(row[idCol - 1]);
      if (currId) continue;

      // בדיקת שורה לא ריקה (יש תוכן כלשהו במישהו מהתאים חוץ מעמודת ה-ID)
      var hasContent = false;
      for (var c = 0; c < row.length; c++) {
        if (c === idCol - 1) continue;
        if (norm_(row[c])) {
          hasContent = true;
          break;
        }
      }
      if (!hasContent) continue;

      var newId;
      do {
        newId = uuid_();
      } while (existing[newId]);
      existing[newId] = true;
      row[idCol - 1] = newId;
      filled++;
    }

    if (filled > 0) {
      dataRange.setValues(data);
    }

    var result = { filled: filled };
    if (warning) result.warning = warning;
    if (typeof logDuration_ === "function") {
      logDuration_(logger, "requests.ensureIds", startMs, {
        filled: filled,
        warning: warning || "",
      });
    }
    return result;
  }

  function handleOpen_(e) {
    var logger = getLogger_("REQUESTS_ON_OPEN");
    var startMs = new Date().getTime();
    try {
      ensureRequestIds_();
    } catch (err) {
      if (typeof logDuration_ === "function") {
        logDuration_(
          logger,
          "requests.onOpen",
          startMs,
          {
            error: String(err),
          },
          "warn",
          "REQ_ON_OPEN_FAIL",
          err,
        );
      }
      return;
    }
    if (typeof logDuration_ === "function") {
      logDuration_(logger, "requests.onOpen", startMs, {});
    }
  }

  function handleEdit_(e) {
    var logger = getLogger_("REQUESTS_ON_EDIT");
    var startMs = new Date().getTime();
    var range = e && e.range ? e.range : null;
    if (!range) return;

    var sheet = range.getSheet();
    if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME_REQUESTS) return;

    // אם עריכה מתרחשת בטאב בקשות – נוודא IDs
    ensureRequestIds_(sheet);
    if (typeof logDuration_ === "function") {
      logDuration_(logger, "requests.onEdit", startMs, {
        sheet: sheet ? sheet.getName() : "",
      });
    }
  }

  // Export
  REQ.ensureRequestIds = ensureRequestIds_;
  REQ.listRequests = list_;
  REQ.handleOpen = handleOpen_;
  REQ.handleEdit = handleEdit_;

  // --- wrappers גלובליים ---
  function REQ_onOpen(e) {
    if (typeof REQ !== "undefined" && REQ.handleOpen) {
      REQ.handleOpen(e || {});
    }
  }
  function REQ_onEdit(e) {
    if (typeof REQ !== "undefined" && REQ.handleEdit) {
      REQ.handleEdit(e || {});
    }
  }

  // חשיפה של ה-wrappers לשימוש חיצוני
  REQ.REQ_onOpen = REQ_onOpen;
  REQ.REQ_onEdit = REQ_onEdit;
})();

// eslint-disable-next-line no-unused-vars
function REQ_onOpen(e) {
  if (typeof REQ !== "undefined" && REQ.REQ_onOpen) {
    REQ.REQ_onOpen(e || {});
  }
}

// eslint-disable-next-line no-unused-vars
function REQ_listRequests(filters) {
  if (typeof REQ !== "undefined" && REQ.listRequests) {
    return REQ.listRequests(filters || {});
  }
  return { ok: false, error: "REQ module missing" };
}

// eslint-disable-next-line no-unused-vars
function REQ_onEdit(e) {
  if (typeof REQ !== "undefined" && REQ.REQ_onEdit) {
    REQ.REQ_onEdit(e || {});
  }
}
