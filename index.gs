// Apps Script entrypoint for boulder-shifts-app. Paste into your Apps Script project.
// All requests arrive as POST JSON: { action: string, payload?: any }.

const EMPLOYEES_SHEET_NAME = "פרטי עובדים";
const EMPLOYEES_SHEET_NAMES = [
  EMPLOYEES_SHEET_NAME,
  "עובדים",
  "Employees",
  "employees",
];
const EMPLOYEE_HEADER_ROW = 1;
const WORK_LOGS_SHEET_NAME = "דיווח שעות עבודה";
const REQUESTS_SHEET_NAMES = ["בקשות עובדים"];
const OPTIONS_SHEET_NAMES = ["אופציות בחירה ו ID'S", "אופציות בחירה ו ID_S"];
const LEADERBOARD_SHEET_NAME = "leaderboard";
const LEADERBOARD_MAX_ROWS = 500;

// Canonical schema for the "משמרות" sheet.
// Source: _bundles/google SS boulder 5-1-26/משמרות.schema.json (exportedAt 2026-01-06T17:46:34.040Z)
// Headers in order (A1:M1):
// 1. ID משמרת
// 2. ID עובד
// 3. שם מלא
// 4. תאריך
// 5. כניסה
// 6. יציאה
// 7. ID סוג עבודה
// 8. סוג עבודה
// 9. מחלקה
// 10. שעות
// 11. כמות יחידות
// 12. הערות
// 13. סטטוס משמרת
const SHIFTS_HEADERS_CANONICAL = [
  "ID משמרת",
  "ID עובד",
  "שם מלא",
  "תאריך",
  "כניסה",
  "יציאה",
  "ID סוג עבודה",
  "סוג עבודה",
  "מחלקה",
  "שעות",
  "כמות יחידות",
  "הערות",
  "סטטוס משמרת",
];

const EMAIL_HEADER_CANDIDATES = [
  "מייל",
  "email",
  "אימייל",
  "mail",
  "כתובת מייל",
  "כתובת אימייל",
  "דואר אלקטרוני",
  'דוא"ל',
];

const HEADER_KEY_MAP = {
  "ID משמרת": "shiftId",
  "ID דיווח": "reportId",
  "Shift ID": "reportId",
  "Request ID": "id",
  requestId: "id",
  "ID בקשה": "id",
  "Employee ID": "employeeId",
  "ID עובד": "employeeId",
  "מזהה עובד": "employeeId",
  "ת.ז": "nationalId",
  "שם מלא": "employeeName",
  "ID סוג עבודה": "jobTypeId",
  "ID סוגי עבודה": "jobTypeId",
  "סוג עבודה": "jobName",
  "סוגי עבודה": "jobName",
  סטטוס: "status",
  "סטטוס סוגי עבודה": "jobStatus",
  "סטטוס משמרת": "status",
  סטטוס: "status",
  מחלקה: "department",
  מחלקות: "department",
  "דיווח שעות": "direction",
  "כניסה / יציאה": "direction",
  כניסה: "startTime",
  יציאה: "endTime",
  "תיקון תאריך": "fixDate",
  "תיקון שעה": "fixTime",
  "תאריך משמרת": "workDate",
  תאריך: "workDate",
  "חותמת זמן": "timestamp",
  שעות: "hoursDecimal",
  "שעות לשכר": "payHours",
  "כמות יחידות": "units",
  "כמות היחידות": "units",
  "דיווח יחידות": "units",
  הערות: "note",
  "הערה למנהל": "noteToManager",
  "תאריך נשלח": "submittedAt",
  "תאריך החלטה": "decidedAt",
  "מנהל מחליט": "managerDecision",
  תז: "nationalId",
  "סוג בקשה": "requestType",
  "סטטוס בקשה": "status",
  "הערות למשמרת": "note",
  "מקור דיווחים": "rawLogIds",
};

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = SCRIPT_PROPERTIES.getProperty("SPREADSHEET_ID") || "";
const DEPLOYMENT_FINGERPRINT =
  SCRIPT_PROPERTIES.getProperty("BUILD_ID") ||
  SCRIPT_PROPERTIES.getProperty("DEPLOYED_AT") ||
  "";
const ACCESS_ISSUE_RECIPIENTS = (
  SCRIPT_PROPERTIES.getProperty("ACCESS_ISSUE_RECIPIENTS") || ""
)
  .split(",")
  .map(function (s) {
    return stringValue(s);
  })
  .filter(function (s) {
    return !!s;
  });

const MONTHLY_ALERT_RECIPIENTS = (
  SCRIPT_PROPERTIES.getProperty("MONTHLY_ALERT_RECIPIENTS") || ""
)
  .split(",")
  .map(function (s) {
    return stringValue(s);
  })
  .filter(function (s) {
    return !!s;
  });

var LOG_LEVEL_SCRIPT = (
  SCRIPT_PROPERTIES.getProperty("LOG_LEVEL_SCRIPT") || "info"
)
  .toLowerCase()
  .trim();
var SYSTEM_LOG_SHEET_NAME = "system_logs";
var LOG_TO_SHEET_ENABLED =
  (SCRIPT_PROPERTIES.getProperty("LOG_TO_SHEET_ENABLED") || "true")
    .toLowerCase()
    .trim() === "true";
var LOG_LIST_TOKEN = (SCRIPT_PROPERTIES.getProperty("LOG_LIST_TOKEN") || "")
  .toString()
  .trim();
var __activeTraceContext = null;
var ERROR_CODES = {
  EMPLOYEE_NOT_FOUND: "EMPLOYEE_NOT_FOUND",
  EMPLOYEE_INACTIVE: "EMPLOYEE_INACTIVE",
  SHEET_MISSING: "SHEET_MISSING",
  SHEET_RANGE_EMPTY: "SHEET_RANGE_EMPTY",
  VALIDATION_FAILED: "VALIDATION_FAILED",
  APPS_SCRIPT_EXCEPTION: "APPS_SCRIPT_EXCEPTION",
  NETWORK_ERROR_TO_SCRIPT: "NETWORK_ERROR_TO_SCRIPT",
  UNAUTHORIZED: "UNAUTHORIZED",
  CONFLICTING_SHIFT: "CONFLICTING_SHIFT",
  UNKNOWN_ERROR: "UNKNOWN_ERROR",
  INTERNAL_ERROR: "INTERNAL_ERROR",
  PROFILE_SAVE_FAILED: "PROFILE_SAVE_FAILED",
  PROFILE_LOAD_FAILED: "PROFILE_LOAD_FAILED",
  REQUEST_LIST_FAILED: "REQUEST_LIST_FAILED",
  WORK_LOG_LOAD_FAILED: "WORK_LOG_LOAD_FAILED",
  SHIFT_CORRECTION_SAVE_FAILED: "SHIFT_CORRECTION_SAVE_FAILED",
  PREFERENCES_SAVE_FAILED: "PREFERENCES_SAVE_FAILED",
  MONTHLY_JOB_ERROR_REPORT_FAILED: "MONTHLY_JOB_ERROR_REPORT_FAILED",
  INVALID_JSON: "INVALID_JSON",
  EMPTY_RESPONSE: "EMPTY_RESPONSE",
};

// Ensure a module logger exists even if a deployed version is missing getModuleLogger_.
function ensureModuleLoggerDefined_(operation) {
  if (typeof getModuleLogger_ === "function") {
    return getModuleLogger_(operation);
  }
  var ctx = __activeTraceContext || {
    traceId: makeTraceId_(),
    operation: operation || "UNKNOWN",
  };
  if (operation) ctx.operation = operation;
  return createScriptLogger_(ctx);
}

// Contract/health metadata used by the status endpoint
const CONTRACT_SCHEMA_VERSION = 1;
const SHEET_CONTRACT_SPEC = [
  {
    key: "Employees",
    possibleNames: EMPLOYEES_SHEET_NAMES,
    requiredColumnsCritical: [
      ["מזהה עובד", "ID עובד", "Employee ID"],
      ["שם מלא", "name"],
      EMAIL_HEADER_CANDIDATES,
      ["סטטוס", "סטטוס פעיל", "active", "status"],
    ],
    requiredColumnsOptional: [],
  },
  {
    key: "WorkLogs",
    possibleNames: [WORK_LOGS_SHEET_NAME],
    requiredColumnsCritical: [
      ["ID דיווח", "Shift ID"],
      ["חותמת זמן", "timestamp"],
      ["ID עובד", "מזהה עובד", "Employee ID"],
      ["שם מלא", "name"],
      ["כניסה / יציאה", "דיווח שעות"],
      ["תיקון תאריך"],
      ["תיקון שעה"],
      ["ID סוג עבודה", "ID סוגי עבודה"],
      ["סוג עבודה", "סוגי עבודה"],
      ["מחלקה", "מחלקות"],
      ["כמות היחידות", "דיווח יחידות"],
    ],
    requiredColumnsOptional: [["הערות", "הערה"]],
  },
  {
    key: "Requests",
    possibleNames: REQUESTS_SHEET_NAMES,
    requiredColumnsCritical: [
      ["ID משמרת"],
      ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
      ["סטטוס", "סטטוס בקשה"],
      ["סוג עבודה", "סוגי עבודה"],
      ["תאריך משמרת", "חותמת זמן", "תיקון תאריך"],
      ["ID סוג עבודה", "ID סוגי עבודה"],
    ],
    requiredColumnsOptional: [["הערות", "הערה למנהל", "הערות למשמרת"]],
  },
];

function mapActionToOperation_(action) {
  if (!action) return "APPS_SCRIPT_PROXY_GENERIC";
  var map = {
    health: "HEALTH",
    "logs.list": "SYSTEM_LOG_LIST",
    "jobTypes.list": "REPORT_LOAD",
    "options.listPayments": "REPORT_LOAD",
    "employee.linkedJobs": "REPORT_LOAD",
    "admin.runBulkAction": "ADMIN_BULK_ACTION_RUN",
    "admin.listEmployees": "ADMIN_EMPLOYEES_LIST",
    "workLogs.listByEmployee": "WORK_LOG_LOAD",
    "requests.listByEmployee": "REQUEST_LIST",
    reportAccessIssue: "REPORT_SAVE",
    "shiftReport.submit": "WORK_LOG_SAVE",
    "shiftReport.deleteByIds": "WORK_LOG_SAVE",
    "shiftCorrection.submit": "SHIFT_CORRECTION_SAVE",
    "shiftReport.monthlyErrorNotify": "MONTHLY_JOB_ERROR_REPORT",
    "requests.approve": "REQUEST_DECIDE",
    "employee.save": "PROFILE_SAVE",
    employeeExistsByEmail: "PROFILE_LOAD",
    getCurrentEmployee: "PROFILE_LOAD",
    getWorkLogs: "WORK_LOG_LOAD",
    getEmployeeRequests: "REQUEST_LIST",
    "gameLeaderboard.list": "LEADERBOARD_LIST",
    "gameLeaderboard.submit": "LEADERBOARD_SAVE",
  };
  return map[action] || "APPS_SCRIPT_PROXY_GENERIC";
}

function getActionRegistry_() {
  return {
    health: function (_payload, _logger) {
      return healthCheck_();
    },
    "logs.list": function (payload, _logger) {
      return listSystemLogs_(payload || {});
    },
    "jobTypes.list": function (_payload, _logger) {
      return { jobTypes: listJobTypes_() };
    },
    "options.listPayments": function (_payload, logger) {
      return { payTypes: listPaymentOptions_(logger) };
    },
    "admin.runBulkAction": function (payload, logger) {
      return adminRunBulkAction_(payload || {}, logger);
    },
    "admin.reportBulkActionIssue": function (payload, logger) {
      return adminReportBulkActionIssue_(payload || {}, { logger: logger });
    },
    "admin.listEmployees": function (payload, logger) {
      return adminListEmployees_(payload || {}, logger);
    },
    "admin.listRequests": function (payload, logger) {
      return adminListRequests_(payload || {}, logger);
    },
    "admin.getRequestDetails": function (payload, logger) {
      return adminGetRequestDetails_(payload || {}, logger);
    },
    "admin.updateRequestStatus": function (payload, logger) {
      return adminUpdateRequestStatus_(payload || {}, logger);
    },
    "admin.createRequest": function (payload, logger) {
      return adminCreateRequest_(payload || {}, logger);
    },
    "admin.archiveRequest": function (payload, logger) {
      return adminArchiveRequest_(payload || {}, logger);
    },
    "employee.linkedJobs": function (payload, _logger) {
      return listEmployeeLinkedJobIds_(payload);
    },
    "workLogs.listByEmployee": function (payload, _logger) {
      return listWorkLogsByEmployee_(payload);
    },
    "requests.listByEmployee": function (payload, _logger) {
      return listRequestsByEmployee_(payload);
    },
    reportAccessIssue: function (payload, _logger) {
      return reportAccessIssue_(payload || {});
    },
    "shiftReport.submit": function (payload, logger) {
      return handleShiftReportSubmit_(payload, logger);
    },
    "shiftReport.deleteByIds": function (payload, logger) {
      return handleShiftReportDeleteByIds_(payload, logger);
    },
    "shiftCorrection.submit": function (payload, logger) {
      return handleShiftCorrectionSubmit_(payload, logger);
    },
    "shiftReport.monthlyErrorNotify": function (payload, _logger) {
      return handleShiftReportMonthlyErrorNotify_(payload);
    },
    "shifts.list": function (payload, _logger) {
      return listShifts_(payload || {});
    },
    "shifts.rebuildRange": function (payload, _logger) {
      return rebuildShiftsForRange_(payload);
    },
    "shifts.getHourlyOverlaps": function (payload, _logger) {
      var filter = (payload && payload.filter) || {};
      return SHIFTS_getHourlyOverlaps(filter);
    },
    "requests.approve": function (payload, _logger) {
      return handleRequestApprove_(payload);
    },
    "employee.save": function (payload, _logger) {
      return handleEmployeeSave_(payload);
    },
    employeeExistsByEmail: function (payload, _logger) {
      return employeeExistsByEmail_(payload || {});
    },
    // SPEC-EMP-001 / SPEC-EMP-002
    getCurrentEmployee: function (payload, _logger) {
      return getCurrentEmployeeData_(payload);
    },
    // Legacy stubs so existing screens stay alive until you wire real data
    getWorkLogs: function (_payload, _logger) {
      return { workLogs: [] };
    },
    getEmployeeRequests: function (_payload, _logger) {
      return { requests: [] };
    },
    "gameLeaderboard.list": function (payload, _logger) {
      return listGameLeaderboard_(payload || {});
    },
    "gameLeaderboard.submit": function (payload, _logger) {
      return submitGameLeaderboard_(payload || {});
    },
  };
}

function getSupportedActions_() {
  const registry = getActionRegistry_();
  return Object.keys(registry);
}

function makeTraceId_() {
  var shortUuid = Utilities.getUuid().split("-")[0];
  return new Date().getTime().toString(36) + "-" + shortUuid;
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action || (params.ping ? "ping" : "");
  try {
    switch (action) {
      case "health": {
        const expected = SCRIPT_PROPERTIES.getProperty("HEALTH_TOKEN");
        const provided = params.token || "";
        if (expected && expected !== provided) {
          return jsonResponse(
            {
              ok: false,
              error: "Unauthorized health check",
              errorCode: "HEALTH_TOKEN_MISMATCH",
            },
            401,
          );
        }
        return jsonResponse(healthCheck_());
      }
      case "ping":
        return jsonResponse(
          withOk_({ pong: true, timestamp: new Date().toISOString() }),
        );
      default:
        return jsonResponse(
          {
            ok: false,
            error: "GET not allowed for this action",
            errorCode: "METHOD_NOT_ALLOWED",
          },
          405,
        );
    }
  } catch (err) {
    return jsonResponse(
      {
        ok: false,
        error: err && err.message ? err.message : "Unexpected error",
        stack: err && err.stack ? String(err.stack).slice(0, 500) : undefined,
      },
      500,
    );
  }
}

// Ensure the web-app endpoints (doGet/doPost) never return HTML.
// Any unexpected exception should be converted into a JSON payload.
function createScriptTraceContext_(incomingMeta, action) {
  var meta =
    incomingMeta && typeof incomingMeta === "object" ? incomingMeta : {};
  var traceId = stringValue(meta.traceId) || makeTraceId_();
  var opFromClient = stringValue(meta.operation);
  var operation = opFromClient || mapActionToOperation_(stringValue(action));

  return {
    traceId: traceId,
    operation: operation || "APPS_SCRIPT_PROXY_GENERIC",
    source: stringValue(meta.source) || "apps-script-webapp",
    actor: meta.actor !== undefined ? meta.actor : null,
    version: stringValue(meta.version) || "1.0",
  };
}

function sanitizeForLogs_(obj) {
  if (!obj || typeof obj !== "object") return obj;
  var blockedKeys = {
    payload: true,
    headers: true,
    authorization: true,
    token: true,
    email: true,
    mail: true,
    spreadsheetid: true,
    spreadsheetId: true,
  };

  function sanitizeValue(value, depth) {
    if (depth > 2) return "[truncated]";
    if (value && typeof value === "object") {
      var out = Array.isArray(value) ? [] : {};
      var keys = Object.keys(value).slice(0, 20);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (blockedKeys[k.toLowerCase ? k.toLowerCase() : k]) {
          out[k] = "[redacted]";
        } else {
          out[k] = sanitizeValue(value[k], depth + 1);
        }
      }
      return out;
    }
    if (typeof value === "string") {
      return value.length > 120 ? value.slice(0, 120) + "…" : value;
    }
    return value;
  }

  return sanitizeValue(obj, 0);
}

function ensureSystemLogSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SYSTEM_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SYSTEM_LOG_SHEET_NAME);
    try {
      sheet.hideSheet();
    } catch (err) {
      // If hiding fails, continue without throwing to keep logging resilient.
    }
  }

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.appendRow([
      "timestamp",
      "traceId",
      "layer",
      "operation",
      "step",
      "severity",
      "actor",
      "errorCode",
      "details",
      "extra",
    ]);
  }

  return sheet;
}

function appendSystemLog_(entry) {
  if (!LOG_TO_SHEET_ENABLED) return;
  try {
    var sheet = ensureSystemLogSheet_();
    var details = "";
    if (entry && entry.details !== undefined) {
      try {
        details = JSON.stringify(sanitizeForLogs_(entry.details)).slice(0, 500);
      } catch (err) {
        details = String(entry.details).slice(0, 500);
      }
    }

    var extra = "";
    if (entry && entry.extra !== undefined) {
      try {
        extra = JSON.stringify(sanitizeForLogs_(entry.extra)).slice(0, 500);
      } catch (err2) {
        extra = String(entry.extra).slice(0, 500);
      }
    }

    sheet.appendRow([
      entry && entry.timestamp ? entry.timestamp : new Date().toISOString(),
      (entry && entry.traceId) || "",
      (entry && entry.layer) || "apps-script-router",
      (entry && entry.operation) || "",
      (entry && entry.step) || "log",
      (entry && entry.severity) || "info",
      entry && entry.actor !== undefined ? entry.actor : "",
      (entry && entry.errorCode) || "",
      details,
      extra,
    ]);
  } catch (err3) {
    try {
      console.warn("[trace] system_log_append_failed", {
        message: err3 && err3.message ? err3.message : String(err3),
      });
    } catch (ignore) {
      // Never throw from logging path.
    }
  }
}

function createScriptLogger_(traceContext) {
  var ctx = traceContext || { traceId: makeTraceId_(), operation: "UNKNOWN" };

  function emit(level, step, details, errorCode, err) {
    var errorStack =
      err && err.stack ? String(err.stack).slice(0, 500) : undefined;
    var errorMeta = err && err.meta ? sanitizeForLogs_(err.meta) : undefined;

    var payload = {
      timestamp: new Date().toISOString(),
      traceId: ctx.traceId,
      operation: ctx.operation,
      layer: "apps-script-router",
      step: step || "log",
      severity: level,
      actor: ctx.actor || null,
      errorCode: errorCode || null,
      details: sanitizeForLogs_(details) || null,
      errorMessage: err && err.message ? String(err.message) : undefined,
      errorStack: errorStack,
      errorMeta: errorMeta,
    };

    // Prefer console logging; sheet logging is optional and may not be configured.
    if (level === "error") {
      console.error("[trace]", payload);
    } else if (level === "warn") {
      console.warn("[trace]", payload);
    } else {
      console.log("[trace]", payload);
    }

    if (LOG_TO_SHEET_ENABLED) {
      appendSystemLog_({
        timestamp: payload.timestamp,
        traceId: payload.traceId,
        operation: payload.operation,
        layer: payload.layer,
        step: payload.step,
        severity: payload.severity,
        actor: payload.actor,
        errorCode: payload.errorCode,
        details: payload.details,
        extra: {
          errorMessage: payload.errorMessage,
          errorStack: payload.errorStack,
          errorMeta: payload.errorMeta,
        },
      });
    }
  }

  return {
    context: ctx,
    debug: function (step, details) {
      emit("debug", step, details);
    },
    info: function (step, details) {
      emit("info", step, details);
    },
    warn: function (step, details, errorCode) {
      emit("warn", step, details, errorCode);
    },
    error: function (step, details, errorCode, err) {
      emit("error", step, details, errorCode, err);
    },
  };
}

function safeDurationMs_(startMs) {
  var now = new Date().getTime();
  if (!startMs || isNaN(startMs)) return 0;
  var diff = now - startMs;
  return diff < 0 ? 0 : diff;
}

function logDuration_(
  logger,
  step,
  startMs,
  details,
  severity,
  errorCode,
  err,
) {
  if (!logger || typeof logger.info !== "function") return;
  var payload =
    details && typeof details === "object" ? Object.assign({}, details) : {};
  payload.durationMs = safeDurationMs_(startMs);

  var level = (severity || "info").toLowerCase();
  if (level === "error" && typeof logger.error === "function") {
    logger.error(step, payload, errorCode, err);
  } else if (level === "warn" && typeof logger.warn === "function") {
    logger.warn(step, payload, errorCode);
  } else if (level === "debug" && typeof logger.debug === "function") {
    logger.debug(step, payload);
  } else {
    logger.info(step, payload);
  }
}

// Lightweight module logger helper to avoid undefined crashes.
function getModuleLogger_(operation) {
  var ctx = __activeTraceContext || {
    traceId: makeTraceId_(),
    operation: operation || "UNKNOWN",
  };
  // If caller passed an operation string, prefer it.
  if (operation) ctx.operation = operation;
  return createScriptLogger_(ctx);
}

/**
 * Structured API log helper.
 *
 * @param {string} kind "request" | "success" | "error"
 * @param {string|null} traceId
 * @param {string|null} operation
 * @param {string|null} action
 * @param {Object|null} extra
 */
function logApiEvent_(kind, traceId, operation, action, extra) {
  try {
    var entry = {
      ts: new Date().toISOString(),
      kind: kind || null,
      traceId: traceId || null,
      operation: operation || null,
      action: action || null,
      extra: extra || null,
    };
    Logger.log("[API_TRACE] " + JSON.stringify(entry));
  } catch (err) {
    // לא רוצים שהלוג עצמו יפיל את הסקריפט
    Logger.log(
      "[API_TRACE_FALLBACK] " +
        (err && err.message ? String(err.message) : String(err)),
    );
  }
}

function doPost(e) {
  var logger = null;
  var action = "";
  var requestStartMs = new Date().getTime();
  var parseStartMs = requestStartMs;
  var traceId = null;
  var operation = null;

  try {
    const parsed = parseBody(e);
    const body = parsed.body || {};
    action = stringValue(body.action);
    const payload = body.payload || {};
    const traceContext = createScriptTraceContext_(body.meta, action);
    __activeTraceContext = traceContext;
    logger = createScriptLogger_(traceContext);
    traceId = traceContext && traceContext.traceId;
    operation = traceContext && traceContext.operation;

    logApiEvent_("request", traceId, operation, action, {
      hasPayload: !!(body && body.payload),
      hasMeta: !!(body && body.meta),
    });

    logger.info("http.action", {
      action: body && body.action,
      hasPayload: !!(body && body.payload),
      hasMeta: !!(body && body.meta),
    });

    var parseDetails = { hasPayload: !!(body && body.payload) };
    if (parsed.error) {
      parseDetails.parseError = String(parsed.error).slice(0, 120);
    }
    logDuration_(
      logger,
      "parse-body",
      parseStartMs,
      parseDetails,
      parsed.error ? "warn" : "debug",
      parsed.error ? "VALIDATION_FAILED" : null,
    );

    if (!body.meta || !body.meta.traceId) {
      logger.warn(
        "start",
        { reason: "Missing traceId from client" },
        "MISSING_TRACE_ID",
      );
    }

    if (parsed.error) {
      logger.warn(
        "validate-input",
        { reason: "Invalid JSON body", error: String(parsed.error) },
        "VALIDATION_FAILED",
      );
      return jsonResponse(
        withOk_({
          ok: false,
          error: "Invalid JSON body: " + String(parsed.error),
          errorCode: "VALIDATION_FAILED",
        }),
        400,
      );
    }

    logger.info("start", {
      action: action || "legacy",
      hasPayload: !!payload,
    });

    if (!action) {
      logger.info("dispatch", { handler: "legacy" });
      var legacyResponse = handleLegacyPost_(e, body, logger);
      logApiEvent_("success", traceId, operation, action || "legacy", {
        resultMeta: { type: "legacy" },
      });
      return legacyResponse;
    }

    logger.info("dispatch", { action: action });
    var dispatchStartMs = new Date().getTime();
    const response = handleNewPost_(action, payload, logger);
    logDuration_(logger, "response", dispatchStartMs, { action: action });
    return response;
  } catch (err) {
    var errorCode =
      err && err.errorCode ? err.errorCode : "APPS_SCRIPT_EXCEPTION";
    var errorDetails = {
      action: action,
      message: err && err.message ? err.message : "Unexpected error",
      stack: err && err.stack ? String(err.stack).slice(0, 500) : undefined,
    };
    if (err && err.meta) {
      errorDetails.meta = err.meta;
    }
    if (logger && logger.error) {
      logger.error("error", errorDetails, errorCode, err);
    } else {
      console.error("[apps-script] doPost failed", err);
    }

    logApiEvent_("error", traceId, operation, action, {
      errorMessage: err && err.message ? String(err.message) : String(err),
      errorName: err && err.name ? String(err.name) : null,
      stack: err && err.stack ? String(err.stack) : null,
    });

    return jsonResponse(
      {
        ok: false,
        error: err && err.message ? err.message : "Unexpected error",
        stack: err && err.stack ? String(err.stack).slice(0, 500) : undefined,
        meta: err && err.meta ? err.meta : undefined,
        errorCode: errorCode,
      },
      500,
    );
  } finally {
    if (logger) {
      logDuration_(logger, "request-total", requestStartMs, {
        action: action || "legacy",
      });
    }
    __activeTraceContext = null;
  }
}

function handleNewPost_(action, payload, logger) {
  var handlers = getActionRegistry_();
  var handler = handlers[action];

  if (!handler) {
    if (logger) logger.warn("unknown-action", { action: action });
    return jsonResponse({ ok: false, error: "Unknown action" }, 400);
  }

  var handlerStartMs = new Date().getTime();
  var result = handler(payload || {}, logger);
  logDuration_(logger, "handler", handlerStartMs, { action: action });
  var ctx = __activeTraceContext || {};
  var traceId = ctx.traceId || null;
  var operation = ctx.operation || null;
  var resultMeta;
  if (Array.isArray(result)) {
    resultMeta = { type: "array", length: result.length };
  } else if (result && typeof result === "object") {
    resultMeta = { type: "object", keys: Object.keys(result) };
  } else {
    resultMeta = { type: typeof result };
  }

  logApiEvent_("success", traceId, operation, action, {
    resultMeta: resultMeta,
  });

  return jsonResponse(withOk_(result));
}

function getCurrentEmployeeData_(payload) {
  const email =
    (payload && payload.email) ||
    (payload && payload.user && payload.user.email) ||
    "";

  if (!email) {
    return { ok: false, error: "Missing email" };
  }

  const result = employeeExistsByEmail_({ email: email });
  const found =
    result && result.ok === true && (result.found === true || result.exists);

  if (found) {
    const emp = result.employee || {};
    const idFromSheet = emp.id || result.employeeId || result.id || "";
    const nameFromSheet =
      emp.name || result.fullName || result.name || emp.fullName || "";
    return {
      employee: {
        ...emp,
        id: idFromSheet || email, // never return empty id
        employeeId: idFromSheet || email,
        name: nameFromSheet || emp.name || "",
        email: emp.email || email,
      },
    };
  }

  return { employee: null };
}

// SPEC-OPS-001
function healthCheck_() {
  if (!SPREADSHEET_ID) {
    return {
      ok: false,
      error: "Missing SPREADSHEET_ID",
      schemaVersion: CONTRACT_SCHEMA_VERSION,
      supportedActions: getSupportedActions_(),
      sheetChecks: [],
      buildId: DEPLOYMENT_FINGERPRINT || null,
      buildIdWarning: !DEPLOYMENT_FINGERPRINT,
    };
  }

  var ss;
  try {
    ss = getSpreadsheetForHealth_();
  } catch (err) {
    return {
      ok: false,
      error: "Failed to open spreadsheet",
      errorCode: "SPREADSHEET_OPEN_FAILED",
      errorMessage: String(err && err.message ? err.message : err).slice(
        0,
        300,
      ),
      schemaVersion: CONTRACT_SCHEMA_VERSION,
      supportedActions: getSupportedActions_(),
      sheetChecks: [],
      buildId: DEPLOYMENT_FINGERPRINT || null,
      buildIdWarning: !DEPLOYMENT_FINGERPRINT,
    };
  }

  const sheetChecks = SHEET_CONTRACT_SPEC.map(function (spec) {
    const entry = {
      key: spec.key,
      ok: false,
      foundSheetName: null,
      missingCriticalColumns: [],
      missingOptionalColumns: [],
    };

    try {
      const sheet = getSheetByNamesStrict_(ss, spec.possibleNames);
      entry.foundSheetName = sheet.getName();
      const headerMap = getHeaderMap_(sheet);

      const missingCritical = collectMissing_(
        headerMap,
        spec.requiredColumnsCritical,
        sheet.getName(),
      );
      const missingOptional = collectMissing_(
        headerMap,
        spec.requiredColumnsOptional,
        sheet.getName(),
      );

      entry.missingCriticalColumns = missingCritical;
      entry.missingOptionalColumns = missingOptional;
      entry.ok = missingCritical.length === 0;
    } catch (err) {
      entry.error = String(err && err.message ? err.message : err);
      entry.missingCriticalColumns = flattenColumnLabels_(
        spec.requiredColumnsCritical,
      );
      entry.missingOptionalColumns = flattenColumnLabels_(
        spec.requiredColumnsOptional,
      );
    }

    return entry;
  });

  return {
    ok: sheetChecks.every(function (c) {
      return c.ok;
    }),
    schemaVersion: CONTRACT_SCHEMA_VERSION,
    supportedActions: getSupportedActions_(),
    sheetChecks: sheetChecks,
    buildId: DEPLOYMENT_FINGERPRINT || null,
    buildIdWarning: !DEPLOYMENT_FINGERPRINT,
  };
}

function listSystemLogs_(payload) {
  var token = stringValue(payload && payload.token);
  if (!LOG_LIST_TOKEN) {
    return {
      ok: false,
      error: "logs_list_disabled",
      errorCode: ERROR_CODES.UNAUTHORIZED,
    };
  }

  if (!token || token !== LOG_LIST_TOKEN) {
    return {
      ok: false,
      error: "unauthorized",
      errorCode: ERROR_CODES.UNAUTHORIZED,
    };
  }

  var limit = Number(payload && payload.limit);
  if (!isFinite(limit) || limit <= 0) limit = 200;
  var MAX_LIMIT = 500;
  if (limit > MAX_LIMIT) limit = MAX_LIMIT;
  var offset = Number(payload && payload.offset);
  if (!isFinite(offset) || offset < 0) offset = 0;

  var sheet = null;
  try {
    sheet = getSpreadsheet_().getSheetByName(SYSTEM_LOG_SHEET_NAME);
  } catch (err) {
    sheet = null;
  }
  if (!sheet) {
    return { ok: true, logs: [], total: 0, limit: limit, offset: offset };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { ok: true, logs: [], total: 0, limit: limit, offset: offset };
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var total = lastRow - 1;
  var startRow = 2 + offset;
  if (startRow > lastRow) {
    return { ok: true, logs: [], total: total, limit: limit, offset: offset };
  }

  var rowsToFetch = Math.min(limit, lastRow - startRow + 1);
  var data = sheet
    .getRange(startRow, 1, rowsToFetch, headers.length)
    .getValues();
  var logs = data.map(function (row) {
    var obj = {};
    for (var i = 0; i < headers.length; i++) {
      obj[stringValue(headers[i])] = row[i];
    }
    return obj;
  });

  return {
    ok: true,
    logs: logs,
    total: total,
    limit: limit,
    offset: offset,
    hasMore: offset + logs.length < total,
  };
}

function handleLegacyPost_(e, body, _logger) {
  const paramAction = stringValue(
    e && e.parameter && e.parameter.action ? e.parameter.action : "",
  );

  if (paramAction === "request" || paramAction === "requests") {
    // Keep the legacy handler intact for old callers; it returns a JSON response.
    return handleRequestPost(e);
  }

  const legacyPayload = normalizeLegacyShiftPayload_(body);
  return jsonResponse(
    withOk_(handleShiftReportSubmit_(legacyPayload, _logger)),
  );
}

function parseBody(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return { body: {}, raw: "" };
  }

  const raw = e.postData.contents;
  try {
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      return { body: parsed, raw: raw };
    }
    return { body: {}, raw: raw };
  } catch (err) {
    return { body: {}, raw: raw, error: err };
  }
}

function normalizeLegacyShiftPayload_(body) {
  const src = body && typeof body === "object" ? body : {};
  return {
    employeeId: stringValue(src.employeeId || src.empId || src.id),
    employeeName: stringValue(src.employeeName || src.name),
    jobId: stringValue(src.jobId || src.jobTypeId),
    jobName: stringValue(
      src.jobName || (src.job && src.job.name ? src.job.name : ""),
    ),
    direction: stringValue(src.direction || src.dir),
    fixDate: stringValue(src.fixDate || src.date || src.workDate),
    fixTime: stringValue(src.fixTime || src.time),
    units: src.units,
    note: src.note || src.notes,
    timestamp: stringValue(src.timestamp) || stringValue(src.timeStamp),
  };
}

function jsonResponse(data, status) {
  var body = data || {};
  if (body && typeof body === "object" && !Array.isArray(body)) {
    var incomingMeta =
      body.meta && typeof body.meta === "object" ? body.meta : {};
    var mergedMeta = Object.assign({}, incomingMeta, {
      traceId:
        __activeTraceContext && __activeTraceContext.traceId
          ? __activeTraceContext.traceId
          : incomingMeta.traceId,
      operation:
        __activeTraceContext && __activeTraceContext.operation
          ? __activeTraceContext.operation
          : incomingMeta.operation,
    });

    if (typeof status === "number") {
      mergedMeta.statusCode = status;
    } else if (incomingMeta.statusCode !== undefined) {
      mergedMeta.statusCode = incomingMeta.statusCode;
    }

    body = Object.assign({}, body, {
      meta: mergedMeta,
    });
  }

  const output = ContentService.createTextOutput(JSON.stringify(body || {}));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function withOk_(data) {
  if (data && typeof data === "object" && data.ok !== undefined) return data;
  if (data && typeof data === "object" && !Array.isArray(data)) {
    return Object.assign({ ok: true }, data);
  }
  return { ok: true, data: data }; // fallback envelope
}

function listJobTypes_() {
  var logger = ensureModuleLoggerDefined_("REPORT_LOAD");
  logger.info("start", { sheetNames: OPTIONS_SHEET_NAMES });
  const sheet = getSheetByPossibleNames_(OPTIONS_SHEET_NAMES);
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const statusCol = getRequiredColumn_(
    headerMap,
    ["סטטוס סוגי עבודה"],
    sheetName,
  );
  const idCol = getRequiredColumn_(
    headerMap,
    ["ID סוגי עבודה", "ID סוג עבודה"],
    sheetName,
  );
  const nameCol = getRequiredColumn_(
    headerMap,
    ["סוגי עבודה", "סוג עבודה"],
    sheetName,
  );
  const deptCol = getRequiredColumn_(headerMap, ["מחלקות", "מחלקה"], sheetName);
  const payStatusCol = getOptionalColumn_(headerMap, ["סטטוס אופן תשלום"]);
  const payIdCol = getOptionalColumn_(headerMap, [
    "ID אופני תשלום",
    "ID אופן תשלום",
  ]);
  const payNameCol = getOptionalColumn_(headerMap, [
    "אופני תשלום",
    "אופן תשלום",
  ]);
  const payDefaultIdCol = getOptionalColumn_(headerMap, [
    "ID אופן תשלום דיפולט",
    "ID אופני תשלום דיפולט",
    "ID דיפולט אופן תשלום",
    "Default Pay Type ID",
    "Default Payment ID",
  ]);
  const payDefaultNameCol = getOptionalColumn_(headerMap, [
    "דיפולט אופן תשלום",
    "אופן תשלום דיפולט",
    "ברירת מחדל אופן תשלום",
    "דיפולט",
    "ברירת מחדל",
    "Default Pay Type",
    "Default Payment",
  ]);
  const dataRange = sheet.getDataRange().getValues();
  logger.info("read-sheet", {
    sheetName: sheetName,
    rows: dataRange.length - 1,
  });
  let blanks = 0;
  const results = [];
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const name = stringValue(row[nameCol - 1]);
    const status = stringValue(row[statusCol - 1]);
    const payStatus = payStatusCol ? stringValue(row[payStatusCol - 1]) : "";
    const payTypeIdRaw = payIdCol ? stringValue(row[payIdCol - 1]) : "";
    const payTypeId =
      payTypeIdRaw ||
      (payDefaultIdCol ? stringValue(row[payDefaultIdCol - 1]) : "");
    const payTypeNameRaw = payNameCol ? stringValue(row[payNameCol - 1]) : "";
    const payTypeName =
      payTypeNameRaw ||
      (payDefaultNameCol ? stringValue(row[payDefaultNameCol - 1]) : "");
    if (!name) {
      blanks++;
      if (blanks >= 20) break;
      continue;
    }
    blanks = 0;
    if (status !== "פעיל") continue;
    if (payStatusCol && payStatus && payStatus !== "פעיל") continue;
    let id = stringValue(row[idCol - 1]);
    if (!id) {
      id = Utilities.getUuid();
      sheet.getRange(i + 1, idCol).setValue(id);
    }
    results.push({
      id: id,
      name: name,
      department: stringValue(row[deptCol - 1]),
      isActive: true,
      payTypeId: payTypeId,
      payTypeName: payTypeName,
      payTypeStatus: payStatus,
    });
  }
  return results;
}

function listPaymentOptions_(logger) {
  var lg = logger || ensureModuleLoggerDefined_("OPTIONS_PAYMENTS_LIST");
  var startMs = new Date().getTime();

  try {
    var payments =
      typeof OPT !== "undefined" && OPT.getAllPayments
        ? OPT.getAllPayments(true)
        : [];

    var mapped = payments
      .filter(function (pay) {
        return !!pay;
      })
      .map(function (pay) {
        return {
          id: String(pay.id || ""),
          name: String(pay.name || ""),
          status: String(pay.status || "").trim(),
        };
      })
      .filter(function (pay) {
        var status = pay.status;
        return !status || status === "פעיל";
      });

    logDuration_(lg, "options.listPayments", startMs, {
      rows: mapped.length,
    });

    return mapped;
  } catch (err) {
    if (lg && typeof lg.error === "function") {
      lg.error("options.listPayments.error", {
        message: err && err.message,
      });
    }
    throw err;
  }
}

function listEmployeeLinkedJobIds_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("REPORT_LOAD");
  var startMs = new Date().getTime();
  const employeeId =
    payload && payload.employeeId ? String(payload.employeeId).trim() : "";
  if (!employeeId) {
    lg.warn(
      "validation",
      { reason: "Missing employeeId" },
      ERROR_CODES.VALIDATION_FAILED,
    );
    throw new Error("Missing employeeId");
  }
  const sheet = getEmployeesSheet_();
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const idCol = getRequiredColumn_(
    headerMap,
    ["מזהה עובד", "ID עובד"],
    sheetName,
  );
  const jobTypeCol = getRequiredColumn_(
    headerMap,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheetName,
  );
  const statusCol = getRequiredColumn_(headerMap, ["סטטוס"], sheetName);
  const jobTypeNameCol = getOptionalColumn_(
    headerMap,
    ["סוג העבודה", "סוג עבודה"],
    sheetName,
  );
  const jobPayTypeIdCol = getOptionalColumn_(
    headerMap,
    ["ID אופן תשלום", "ID אופני תשלום"],
    sheetName,
  );
  const jobPayTypeNameCol = getOptionalColumn_(
    headerMap,
    ["אופן תשלום"],
    sheetName,
  );
  const jobDepartmentCol = getOptionalColumn_(
    headerMap,
    ["מחלקה", "מחלקות"],
    sheetName,
  );
  const values = sheet.getDataRange().getValues();
  lg.info("read-sheet", { sheetName: sheetName, rows: values.length - 1 });
  const jobTypeIds = [];
  const jobTypeDetailMap = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmpId = stringValue(row[idCol - 1]);
    if (rowEmpId !== employeeId) continue;
    const status = stringValue(row[statusCol - 1]);
    if (status === "לא פעיל") continue;
    const jobTypeId = stringValue(row[jobTypeCol - 1]);
    if (!jobTypeId) continue;

    jobTypeIds.push(jobTypeId);

    const payTypeId = jobPayTypeIdCol
      ? stringValue(row[jobPayTypeIdCol - 1])
      : "";
    const payTypeName = jobPayTypeNameCol
      ? stringValue(row[jobPayTypeNameCol - 1])
      : "";
    const normalizedPayType = normalizePayType_(payTypeId || payTypeName);
    const existing = jobTypeDetailMap[jobTypeId] || {};
    const jobName = jobTypeNameCol ? stringValue(row[jobTypeNameCol - 1]) : "";
    const department = jobDepartmentCol
      ? stringValue(row[jobDepartmentCol - 1])
      : "";

    jobTypeDetailMap[jobTypeId] = {
      jobTypeId: jobTypeId,
      name: existing.name || jobName,
      department: existing.department || department,
      payTypeId: existing.payTypeId || payTypeId,
      payTypeName: existing.payTypeName || payTypeName,
      payType:
        normalizedPayType || existing.payType || payTypeName || payTypeId || "",
    };
  }

  const jobTypesDetailed = Object.keys(jobTypeDetailMap).map(function (key) {
    return jobTypeDetailMap[key];
  });
  logDuration_(lg, "employee.linkedJobs.total", startMs, {
    employeeId: employeeId,
    linkedCount: jobTypeIds.length,
    detailedCount: jobTypesDetailed.length,
  });
  return { employeeId, jobTypeIds, jobTypesDetailed };
}

function adminListEmployees_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_EMPLOYEES_LIST");
  var startMs = new Date().getTime();
  const sheet = getEmployeesSheet_();
  const headers = getHeaderMap_(sheet);
  const sheetName = sheet.getName();

  const idCol = getRequiredColumn_(
    headers,
    ["מזהה עובד", "ID עובד", "Employee ID"],
    sheetName,
  );
  const nameCol = getOptionalColumn_(headers, ["שם מלא", "name", "Full Name"]);
  const emailCol = getOptionalColumn_(headers, EMAIL_HEADER_CANDIDATES);
  const phoneCol = getOptionalColumn_(headers, [
    "טלפון",
    "פלאפון",
    "נייד",
    "phone",
    "mobile",
  ]);
  const genderCol = getOptionalColumn_(headers, ["מין", "gender"]);
  const nationalIdCol = getOptionalColumn_(headers, [
    "תז",
    "id",
    "id number",
    "תעודת זהות",
  ]);
  const birthDateCol = getOptionalColumn_(headers, [
    "ת. לידה",
    "תאריך לידה",
    "birth date",
    "birthday",
  ]);
  const shirtSizeCol = getOptionalColumn_(headers, [
    "מידת חולצה",
    "shirt size",
    "shirt",
  ]);
  const roleTitleCol = getOptionalColumn_(headers, [
    "תפקיד",
    "role title",
    "title",
    "תיאור תפקיד",
  ]);
  const notesCol = getOptionalColumn_(headers, ["הערות", "notes"]);
  const jobTypeCol = getOptionalColumn_(headers, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
    "Job Type ID",
  ]);
  const jobTypeNameCol = getOptionalColumn_(headers, [
    "סוג העבודה",
    "סוג עבודה",
  ]);
  const jobDepartmentCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const jobAmountCol = getOptionalColumn_(headers, ["סכום"]);
  const jobPayTypeIdCol = getOptionalColumn_(headers, [
    "ID אופן תשלום",
    "ID אופני תשלום",
  ]);
  const jobPayTypeNameCol = getOptionalColumn_(headers, ["אופן תשלום"]);
  const travelAllowanceCol = getOptionalColumn_(headers, [
    "עלות החזרי נסיעות יומי",
    "החזר נסיעות",
    "נסיעות",
    "travel allowance",
  ]);
  const statusCol = getOptionalColumn_(headers, [
    "סטטוס",
    "סטטוס פעיל",
    "active",
    "status",
  ]);
  const employmentStatusCol = getOptionalColumn_(headers, [
    "סטטוס העסקה",
    "employment status",
    "Employment Status",
  ]);
  const branchCol = getOptionalColumn_(headers, [
    "סניף",
    "branch",
    "Branch",
    "Branch Id",
    "branch id",
    "branchId",
  ]);
  const systemRoleCol = getOptionalColumn_(headers, [
    "System Role",
    "system role",
    "Role",
    "Admin Role",
  ]);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow <= EMPLOYEE_HEADER_ROW) {
    lg.info("read-sheet", {
      sheetName: sheetName,
      rows: 0,
    });
    return { employees: [] };
  }

  var dataRange = sheet.getRange(
    EMPLOYEE_HEADER_ROW + 1,
    1,
    lastRow - EMPLOYEE_HEADER_ROW,
    lastCol,
  );
  var values = dataRange.getValues();

  lg.info("read-sheet", {
    sheetName: sheetName,
    rows: values.length,
  });

  const jobTypes = listJobTypes_();
  const jobMap = {};
  const jobMapByName = {};
  jobTypes.forEach(function (jt) {
    jobMap[jt.id] = jt;
    const nameKey = stringValue(jt.name).toLowerCase();
    if (nameKey && !jobMapByName[nameKey]) {
      jobMapByName[nameKey] = jt;
    }
  });

  const statusFilterRaw = stringValue(payload.statusFilter).toUpperCase();
  const statusFilter =
    statusFilterRaw === "ACTIVE" || statusFilterRaw === "INACTIVE"
      ? statusFilterRaw
      : "ALL";
  const jobTypeFilter = stringValue(payload.jobTypeId);
  const search = stringValue(payload.search).toLowerCase();
  const employmentFilter = stringValue(payload.employmentStatus).toLowerCase();
  const branchFilter = stringValue(
    payload.branch || payload.branchId,
  ).toLowerCase();

  function normalizeSystemRole_(raw) {
    const val = stringValue(raw).toUpperCase();
    if (!val) return "EMPLOYEE";
    if (val === "NONE") return "NONE";
    if (val.indexOf("ADMIN") !== -1) return "ADMIN";
    if (val.indexOf("SHIFT") !== -1 || val.indexOf("MANAGER") !== -1)
      return "SHIFT_MANAGER";
    if (val.indexOf("VIEW") !== -1) return "VIEWER";
    if (val === "EMPLOYEE") return "EMPLOYEE";
    return "EMPLOYEE";
  }

  function rolePriority_(role) {
    switch (role) {
      case "ADMIN":
        return 3;
      case "SHIFT_MANAGER":
        return 2;
      case "VIEWER":
        return 1;
      default:
        return 0;
    }
  }

  function normalizeEmploymentStatus_(raw, inactiveFlag) {
    const val = stringValue(raw);
    if (!val && inactiveFlag) return "INACTIVE";
    if (!val) return "ACTIVE";
    const lower = val.toLowerCase();
    if (lower.indexOf("לא פעיל") !== -1 || lower.indexOf("inactive") !== -1)
      return "INACTIVE";
    if (lower.indexOf("terminated") !== -1 || lower.indexOf("פיט") !== -1)
      return "TERMINATED";
    if (lower.indexOf("leave") !== -1 || lower.indexOf("חופ") !== -1)
      return "ON_LEAVE";
    if (lower.indexOf("pending") !== -1 || lower.indexOf("start") !== -1)
      return "PENDING_START";
    const normalized = val.replace(/\s+/g, "_").toUpperCase();
    return normalized || (inactiveFlag ? "INACTIVE" : "ACTIVE");
  }

  const employees = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    const employeeId = stringValue(row[idCol - 1]);
    if (!employeeId) continue;

    const statusVal = statusCol ? stringValue(row[statusCol - 1]) : "";
    const statusLower = statusVal.toLowerCase();
    const rowInactive =
      statusLower.indexOf("לא פעיל") !== -1 ||
      statusLower.indexOf("inactive") !== -1 ||
      statusLower.indexOf("terminated") !== -1;
    const employmentVal = employmentStatusCol
      ? stringValue(row[employmentStatusCol - 1])
      : "";
    const branchVal = branchCol ? stringValue(row[branchCol - 1]) : "";

    let record = employees[employeeId];
    if (!record) {
      record = {
        id: employeeId,
        fullName: "",
        email: "",
        phone: "",
        gender: "",
        nationalId: "",
        birthDate: "",
        shirtSize: "",
        travelAllowanceDaily: "",
        roleTitle: "",
        notes: "",
        branch: "",
        employmentStatus: "",
        isActive: true,
        jobTypeIds: [],
        jobTypesDetailed: [],
        systemRole: "EMPLOYEE",
        _jobTypeDetailMap: {},
        _hasActiveRow: false,
        _hasInactiveRow: false,
        _roleRank: 0,
      };
      employees[employeeId] = record;
    }

    if (!record.fullName && nameCol)
      record.fullName = stringValue(row[nameCol - 1]);
    if (!record.email && emailCol)
      record.email = stringValue(row[emailCol - 1]);
    if (!record.phone && phoneCol)
      record.phone = stringValue(row[phoneCol - 1]);
    if (!record.gender && genderCol)
      record.gender = stringValue(row[genderCol - 1]);
    if (!record.nationalId && nationalIdCol)
      record.nationalId = stringValue(row[nationalIdCol - 1]);
    if (!record.birthDate && birthDateCol)
      record.birthDate = toIsoDate_(row[birthDateCol - 1]);
    if (!record.shirtSize && shirtSizeCol)
      record.shirtSize = stringValue(row[shirtSizeCol - 1]);
    if (!record.roleTitle && roleTitleCol)
      record.roleTitle = stringValue(row[roleTitleCol - 1]);
    if (!record.notes && notesCol)
      record.notes = stringValue(row[notesCol - 1]);
    if (!record.branch && branchVal) record.branch = branchVal;
    if (!record.travelAllowanceDaily && travelAllowanceCol)
      record.travelAllowanceDaily = stringValue(row[travelAllowanceCol - 1]);

    const normalizedEmployment = normalizeEmploymentStatus_(
      employmentVal,
      rowInactive,
    );
    if (!record.employmentStatus)
      record.employmentStatus = normalizedEmployment;
    if (
      record.employmentStatus === "ACTIVE" &&
      normalizedEmployment !== "ACTIVE"
    )
      record.employmentStatus = normalizedEmployment;

    if (rowInactive) {
      record._hasInactiveRow = true;
    } else {
      record._hasActiveRow = true;
    }

    var jobTypeId = jobTypeCol ? stringValue(row[jobTypeCol - 1]) : "";
    var jobTypeNameVal = jobTypeNameCol
      ? stringValue(row[jobTypeNameCol - 1])
      : "";
    var jobTypeNameKey = jobTypeNameVal.toLowerCase();

    if (!jobTypeId && jobTypeNameKey && jobMapByName[jobTypeNameKey]) {
      jobTypeId = jobMapByName[jobTypeNameKey].id || "";
    }

    if (jobTypeId) {
      if (record.jobTypeIds.indexOf(jobTypeId) === -1) {
        record.jobTypeIds.push(jobTypeId);
      }

      var jobTypeFromMap =
        jobMap[jobTypeId] ||
        (jobTypeNameKey ? jobMapByName[jobTypeNameKey] : {}) ||
        {};
      var detailPrev = record._jobTypeDetailMap[jobTypeId] || {};
      var jobDept = jobDepartmentCol
        ? stringValue(row[jobDepartmentCol - 1])
        : "";
      var payTypeIdRaw = jobPayTypeIdCol
        ? stringValue(row[jobPayTypeIdCol - 1])
        : "";
      var payTypeNameRaw = jobPayTypeNameCol
        ? stringValue(row[jobPayTypeNameCol - 1])
        : "";
      var amountRaw = jobAmountCol ? row[jobAmountCol - 1] : "";
      var amountValue =
        amountRaw !== null && amountRaw !== undefined && amountRaw !== ""
          ? amountRaw
          : detailPrev.amount !== undefined
            ? detailPrev.amount
            : jobTypeFromMap.amount || "";

      record._jobTypeDetailMap[jobTypeId] = {
        jobTypeId: jobTypeId,
        name: jobTypeNameVal || detailPrev.name || jobTypeFromMap.name || "",
        shortCode: detailPrev.shortCode || jobTypeFromMap.shortCode || "",
        colorHex: detailPrev.colorHex || jobTypeFromMap.colorHex || "",
        isActive:
          detailPrev.isActive !== undefined
            ? detailPrev.isActive
            : jobTypeFromMap.isActive !== false,
        department:
          jobDept || detailPrev.department || jobTypeFromMap.department || "",
        payTypeId:
          payTypeIdRaw ||
          detailPrev.payTypeId ||
          jobTypeFromMap.payTypeId ||
          "",
        payTypeName:
          payTypeNameRaw ||
          detailPrev.payTypeName ||
          jobTypeFromMap.payTypeName ||
          "",
        payType:
          payTypeNameRaw ||
          payTypeIdRaw ||
          detailPrev.payType ||
          detailPrev.payTypeName ||
          detailPrev.payTypeId ||
          jobTypeFromMap.payTypeName ||
          jobTypeFromMap.payTypeId ||
          "",
        amount: amountValue,
      };
    }

    const rawSystemRole = systemRoleCol
      ? stringValue(row[systemRoleCol - 1])
      : "";
    if (rawSystemRole) {
      const normalizedRole = normalizeSystemRole_(rawSystemRole);
      const rank = rolePriority_(normalizedRole);
      if (rank > record._roleRank) {
        record._roleRank = rank;
        record.systemRole = normalizedRole;
      }
    }
  }

  const allEmployees = Object.keys(employees).map(function (key) {
    const rec = employees[key];
    const isActive = rec._hasActiveRow
      ? true
      : rec._hasInactiveRow
        ? false
        : true;
    const employmentStatus = normalizeEmploymentStatus_(
      rec.employmentStatus,
      !isActive,
    );
    const jobTypesDetailed = rec.jobTypeIds.map(function (jobTypeId) {
      const jt = jobMap[jobTypeId] || {};
      const detail = rec._jobTypeDetailMap[jobTypeId] || {};
      const isActive =
        detail.isActive !== undefined ? detail.isActive : jt.isActive !== false;
      const amountValue =
        detail.amount !== undefined && detail.amount !== ""
          ? detail.amount
          : jt.amount || "";

      return {
        jobTypeId: jobTypeId,
        name: detail.name || jt.name || "",
        shortCode: detail.shortCode || jt.shortCode || "",
        colorHex: detail.colorHex || jt.colorHex || "",
        isActive: isActive,
        department: detail.department || jt.department || "",
        payTypeId: detail.payTypeId || jt.payTypeId || "",
        payTypeName: detail.payTypeName || jt.payTypeName || "",
        payType:
          detail.payType ||
          detail.payTypeName ||
          detail.payTypeId ||
          jt.payTypeName ||
          jt.payTypeId ||
          "",
        amount: amountValue,
      };
    });

    return {
      id: rec.id,
      fullName: rec.fullName,
      email: rec.email,
      phone: rec.phone,
      gender: rec.gender,
      nationalId: rec.nationalId,
      birthDate: rec.birthDate,
      shirtSize: rec.shirtSize,
      travelAllowanceDaily: rec.travelAllowanceDaily,
      roleTitle: rec.roleTitle,
      notes: rec.notes,
      branch: rec.branch,
      employmentStatus: employmentStatus,
      isActive: isActive,
      jobTypeIds: rec.jobTypeIds,
      jobTypesDetailed: jobTypesDetailed,
      systemRole: rec.systemRole,
    };
  });

  const filtered = allEmployees.filter(function (emp) {
    if (statusFilter === "ACTIVE" && !emp.isActive) return false;
    if (statusFilter === "INACTIVE" && emp.isActive) return false;
    if (jobTypeFilter && emp.jobTypeIds.indexOf(jobTypeFilter) === -1)
      return false;
    if (employmentFilter) {
      const empEmployment = stringValue(emp.employmentStatus).toLowerCase();
      if (empEmployment !== employmentFilter) return false;
    }
    if (branchFilter) {
      const empBranch = stringValue(emp.branch).toLowerCase();
      if (empBranch !== branchFilter) return false;
    }
    if (search) {
      const haystack = [
        emp.fullName,
        emp.email,
        emp.phone,
        emp.roleTitle,
        emp.id,
        emp.branch,
      ]
        .map(function (v) {
          return stringValue(v).toLowerCase();
        })
        .filter(function (v) {
          return !!v;
        });
      const matched = haystack.some(function (v) {
        return v.indexOf(search) !== -1;
      });
      if (!matched) return false;
    }
    return true;
  });

  logDuration_(lg, "admin.listEmployees.total", startMs, {
    sheet: sheetName,
    scannedRows: values.length,
    employees: allEmployees.length,
    returned: filtered.length,
    statusFilter: statusFilter,
    jobTypeFilter: !!jobTypeFilter,
    search: !!search,
  });

  return { employees: filtered };
}

function adminRunBulkAction_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_BULK_ACTION_RUN");
  var startMs = new Date().getTime();

  try {
    var data = EmployeesModule_adminRunBulkAction_(payload || {}, {
      logger: lg,
    });

    var targetCount =
      data &&
      data.result &&
      data.result.summary &&
      typeof data.result.summary.targetEmployeesCount === "number"
        ? data.result.summary.targetEmployeesCount
        : null;

    logDuration_(lg, "admin.runBulkAction.total", startMs, {
      mode: payload && payload.mode,
      bulkType: payload && payload.bulkType,
      targetEmployees: targetCount,
    });

    return {
      ok: true,
      data: data,
    };
  } catch (err) {
    if (lg && lg.error) {
      lg.error(
        "admin.runBulkAction.error",
        { message: err && err.message ? err.message : err },
        ERROR_CODES.INTERNAL_ERROR,
        err,
      );
    }

    return {
      ok: false,
      errorCode:
        err && err.code === "VALIDATION_ERROR"
          ? "VALIDATION_ERROR"
          : ERROR_CODES.INTERNAL_ERROR,
      errorMessage:
        err && err.message
          ? String(err.message)
          : "Unknown error in admin.runBulkAction",
    };
  }
}

/**
 * Handles admin.reportBulkActionIssue from the web app.
 * @param {Object} request BulkActionIssueReportPayload
 * @param {Object} context ActionContext
 * @return {Object} BulkActionIssueReportResponse
 */
function adminReportBulkActionIssue_(request, context) {
  return EmployeesModule_adminReportBulkActionIssue_(request, context);
}

// TODO(admin-requests): mapping based on existing sheet: id = requestId or shiftId; employeeId + employeeName; raw status from "סטטוס בקשה"; createdAt from "חותמת זמן"/"תאריך משמרת"; optional shiftId, jobTypeId/name, department, units, manager note.
function normalizeAdminRequestStatus_(raw) {
  var val = stringValue(raw).toUpperCase();
  if (!val) return "PENDING";
  if (val.indexOf("APPROVED") !== -1 || val.indexOf("מאושר") !== -1)
    return "APPROVED";
  if (val.indexOf("REJECT") !== -1 || val.indexOf("נדח") !== -1)
    return "REJECTED";
  if (val.indexOf("CANCELLED_BY_EMPLOYEE") !== -1 || val.indexOf("עובד") !== -1)
    return "CANCELLED_BY_EMPLOYEE";
  if (
    val.indexOf("CANCELLED_BY_MANAGER") !== -1 ||
    val.indexOf("CANCELLED") !== -1 ||
    val.indexOf("MANAGER") !== -1 ||
    val.indexOf("בוטל") !== -1
  )
    return "CANCELLED_BY_MANAGER";
  if (val.indexOf("AUTO") !== -1) return "AUTO_APPROVED";
  if (val.indexOf("PENDING") !== -1 || val.indexOf("ממתין") !== -1)
    return "PENDING";
  return "PENDING";
}

function normalizeIsoDateTimeValue_(val) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date && !isNaN(val.getTime())) return val.toISOString();
  var s = stringValue(val);
  if (!s) return "";
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return parsed.toISOString();
  return s;
}

function parseDateOnly_(val) {
  var s = stringValue(val);
  if (!s) return null;
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return parsed;
  return null;
}

function toBooleanCell_(val) {
  if (val === true) return true;
  if (val === false) return false;
  var s = stringValue(val).toLowerCase();
  if (!s) return false;
  if (s === "true" || s === "yes" || s === "1" || s === "כן") return true;
  return false;
}

function resolveRequestColumns_(options) {
  var sheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
  var headers = getHeaderMap_(sheet);
  var sheetName = sheet.getName();
  var ensureOptional = options && options.ensureOptional === true;

  function required(names) {
    return getRequiredColumn_(headers, names, sheetName);
  }

  function optional(names) {
    var col = getOptionalColumn_(headers, names);
    if (col || !ensureOptional) return col;
    var target = sheet.getLastColumn() + 1;
    sheet.getRange(1, target).setValue(names[0]);
    headers = getHeaderMap_(sheet);
    return target;
  }

  return {
    sheet: sheet,
    headers: headers,
    statusCol: required(["סטטוס בקשה", "סטטוס", "Status"]),
    requestIdCol: optional(["ID בקשה", "Request ID", "RequestID"]),
    shiftIdCol: optional(["ID משמרת", "Shift ID", "מזהה משמרת"]),
    createdAtCol: optional(["חותמת זמן", "תאריך משמרת", "תיקון תאריך"]),
    shiftDateCol: optional(["תאריך משמרת", "תיקון תאריך", "חותמת זמן"]),
    employeeIdCol: required([
      "מזהה עובד",
      "ID עובד",
      "Employee ID",
      "ת.ז",
      "תז",
    ]),
    employeeNameCol: optional(["שם מלא", "שם עובד", "Employee Name"]),
    requestTypeIdCol: optional([
      "ID סוג בקשה",
      "ID סוג עבודה",
      "ID סוגי עבודה",
      "Job Type ID",
    ]),
    requestTypeLabelCol: optional([
      "סוג בקשה",
      "תיאור בקשה",
      "סוג עבודה",
      "סוגי עבודה",
      "שם סוג עבודה",
    ]),
    departmentCol: optional(["מחלקה", "מחלקות", "Department"]),
    unitsCol: optional(["כמות היחידות", "דיווח יחידות", "Units"]),
    managerCommentCol: optional(["הערת מנהל", "הערות מנהל", "Manager Comment"]),
    employeeCommentCol: optional([
      "הערות למשמרת",
      "הערת עובד",
      "Employee Comment",
    ]),
    sourceCol: optional(["מקור", "Source"]),
    isArchivedCol: optional(["בארכיון", "ארכיון", "Archived", "isArchived"]),
    decisionAtCol: optional(["תאריך החלטה", "decisionAt"]),
    decidedByNameCol: optional(["מקבל החלטה", "הוחלט על ידי", "decidedByName"]),
    decidedByRoleCol: optional(["תפקיד מחליט", "decidedByRole"]),
    decidedByIdCol: optional([
      "מזהה מחליט",
      "decidedById",
      "decidedByEmployeeId",
    ]),
    updatedAtCol: optional(["עודכן", "תאריך עדכון", "updatedAt"]),
    dateFromCol: optional(["מתאריך", "dateFrom"]),
    dateToCol: optional(["עד תאריך", "dateTo"]),
    singleDateCol: optional(["תאריך יחיד", "singleDate"]),
  };
}

function mapRequestRow_(row, cols) {
  function cell(col) {
    if (!col) return "";
    return row[col - 1];
  }

  var statusRaw = stringValue(cell(cols.statusCol));
  var status = normalizeAdminRequestStatus_(statusRaw);
  var createdAt = cols.createdAtCol
    ? normalizeIsoDateTimeValue_(cell(cols.createdAtCol))
    : "";
  var shiftDate = cols.shiftDateCol
    ? normalizeIsoDateTimeValue_(cell(cols.shiftDateCol))
    : createdAt;
  var dateFrom = cols.dateFromCol
    ? normalizeIsoDateTimeValue_(cell(cols.dateFromCol))
    : "";
  var dateTo = cols.dateToCol
    ? normalizeIsoDateTimeValue_(cell(cols.dateToCol))
    : "";
  var singleDate = cols.singleDateCol
    ? normalizeIsoDateTimeValue_(cell(cols.singleDateCol))
    : shiftDate;

  var sourceRaw = stringValue(cell(cols.sourceCol)).toUpperCase();
  var source = sourceRaw ? sourceRaw : null;
  var isArchived = cols.isArchivedCol
    ? toBooleanCell_(cell(cols.isArchivedCol))
    : false;

  var idFromRequest = stringValue(cell(cols.requestIdCol));
  var idFromShift = stringValue(cell(cols.shiftIdCol));
  var id = idFromRequest || idFromShift;

  return {
    id: id,
    employeeId: stringValue(cell(cols.employeeIdCol)),
    employeeName: stringValue(cell(cols.employeeNameCol)),
    employeeEmail: null,
    status: status,
    rawStatus: statusRaw,
    statusLabel: statusRaw || status,
    createdAt: createdAt || null,
    updatedAt: cols.updatedAtCol
      ? normalizeIsoDateTimeValue_(cell(cols.updatedAtCol)) || null
      : null,
    decisionAt: cols.decisionAtCol
      ? normalizeIsoDateTimeValue_(cell(cols.decisionAtCol)) || null
      : null,
    dateFrom: dateFrom || null,
    dateTo: dateTo || null,
    singleDate: singleDate || null,
    shiftRefDate: shiftDate || null,
    shiftId: idFromShift || null,
    shiftRef: { shiftId: idFromShift || null, shiftDate: shiftDate || null },
    requestTypeId: stringValue(cell(cols.requestTypeIdCol)) || null,
    requestTypeLabel:
      stringValue(cell(cols.requestTypeLabelCol)) ||
      stringValue(cell(cols.requestTypeIdCol)) ||
      null,
    requestTypeName:
      stringValue(cell(cols.requestTypeLabelCol)) ||
      stringValue(cell(cols.requestTypeIdCol)) ||
      null,
    jobTypeId: stringValue(cell(cols.requestTypeIdCol)) || null,
    jobTypeName: stringValue(cell(cols.requestTypeLabelCol)) || null,
    department: stringValue(cell(cols.departmentCol)) || null,
    units: cell(cols.unitsCol) !== "" ? cell(cols.unitsCol) : null,
    source: source,
    isArchived: isArchived,
    managerNotes: stringValue(cell(cols.managerCommentCol)) || null,
    managerComment: stringValue(cell(cols.managerCommentCol)) || null,
    employeeComment: stringValue(cell(cols.employeeCommentCol)) || null,
    decidedById: stringValue(cell(cols.decidedByIdCol)) || null,
    decidedByName: stringValue(cell(cols.decidedByNameCol)) || null,
    decidedByRole: stringValue(cell(cols.decidedByRoleCol)) || null,
  };
}

function findRequestRowIndex_(values, cols, targetId) {
  var id = stringValue(targetId);
  if (!id) return -1;
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rid = cols.requestIdCol ? stringValue(row[cols.requestIdCol - 1]) : "";
    var sid = cols.shiftIdCol ? stringValue(row[cols.shiftIdCol - 1]) : "";
    if (rid === id || sid === id) return i;
  }
  return -1;
}

function requestValidationError_(msg) {
  var e = new Error(msg || "VALIDATION_ERROR");
  e.code = "VALIDATION_ERROR";
  return e;
}

function adminListRequests_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_REQUESTS_LIST");
  var startMs = new Date().getTime();
  var filters = (payload && payload.filters) || payload || {};

  var cols = resolveRequestColumns_({ ensureOptional: false });
  var sheet = cols.sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {
    lg.info("read-sheet", { sheetName: sheet.getName(), rows: 0 });
    return { requests: [] };
  }

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  lg.info("read-sheet", {
    sheetName: sheet.getName(),
    rows: values.length - 1,
  });

  var dateFromFilter = parseDateOnly_(filters.dateFrom);
  var dateToFilter = parseDateOnly_(filters.dateTo);
  var search = stringValue(filters.employeeSearch || filters.q).toLowerCase();

  var statusFilter = stringValue(filters.status || "ALL").toUpperCase();
  var includeArchived = filters && filters.includeArchived === true;

  var results = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var req = mapRequestRow_(row, cols);
    if (!req.id) continue;

    if (
      filters.employeeId &&
      req.employeeId !== stringValue(filters.employeeId)
    )
      continue;

    if (statusFilter === "OPEN_ONLY" && req.status !== "PENDING") continue;
    if (statusFilter !== "OPEN_ONLY" && statusFilter !== "ALL") {
      if (req.status !== statusFilter) continue;
    }

    if (!includeArchived && req.isArchived) continue;

    if (filters.requestTypeId) {
      if (req.requestTypeId !== stringValue(filters.requestTypeId)) continue;
    }

    if (filters.source && stringValue(filters.source).toUpperCase() !== "ALL") {
      var desiredSource = stringValue(filters.source).toUpperCase();
      if (!req.source || req.source !== desiredSource) continue;
    }

    if (dateFromFilter && req.createdAt) {
      var reqDate = new Date(req.createdAt);
      if (!isNaN(reqDate.getTime()) && reqDate < dateFromFilter) continue;
    }

    if (dateToFilter && req.createdAt) {
      var reqDateTo = new Date(req.createdAt);
      if (!isNaN(reqDateTo.getTime()) && reqDateTo > dateToFilter) continue;
    }

    if (search) {
      var haystack = [
        req.employeeName,
        req.employeeEmail,
        req.id,
        req.requestTypeLabel,
        req.managerComment,
        req.rawStatus,
      ]
        .map(function (v) {
          return stringValue(v).toLowerCase();
        })
        .filter(function (v) {
          return !!v;
        });
      var matched = haystack.some(function (v) {
        return v.indexOf(search) !== -1;
      });
      if (!matched) continue;
    }

    results.push(req);
  }

  logDuration_(lg, "admin.listRequests.total", startMs, {
    sheet: sheet.getName(),
    scannedRows: values.length - 1,
    returned: results.length,
    status: statusFilter,
    includeArchived: includeArchived,
  });

  return { requests: results };
}

function adminGetRequestDetails_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_REQUESTS_DETAILS");
  var startMs = new Date().getTime();
  var id = stringValue(payload && (payload.id || payload.requestId));
  if (!id) throw requestValidationError_("Missing requestId");

  var cols = resolveRequestColumns_({ ensureOptional: false });
  var sheet = cols.sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var rowIdx = findRequestRowIndex_(values, cols, id);
  if (rowIdx < 1) {
    throw requestValidationError_("REQUEST_NOT_FOUND");
  }

  var row = values[rowIdx];
  var mapped = mapRequestRow_(row, cols);

  logDuration_(lg, "admin.getRequestDetails.total", startMs, {
    id: id,
  });

  return { request: mapped };
}

function adminUpdateRequestStatus_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_REQUESTS_UPDATE");
  var startMs = new Date().getTime();
  var id = stringValue(payload && (payload.id || payload.requestId));
  var newStatus = normalizeAdminRequestStatus_(payload && payload.newStatus);
  if (!id) throw requestValidationError_("Missing requestId");
  if (!newStatus) throw requestValidationError_("Missing status");

  var allowed = [
    "PENDING",
    "APPROVED",
    "REJECTED",
    "CANCELLED_BY_EMPLOYEE",
    "CANCELLED_BY_MANAGER",
    "AUTO_APPROVED",
  ];
  if (allowed.indexOf(newStatus) === -1) {
    throw requestValidationError_("INVALID_STATUS");
  }

  var cols = resolveRequestColumns_({ ensureOptional: true });
  var sheet = cols.sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var rowIdx = findRequestRowIndex_(values, cols, id);
  if (rowIdx < 1) throw requestValidationError_("REQUEST_NOT_FOUND");
  var rowNumber = rowIdx + 1; // 1-based row index

  sheet.getRange(rowNumber, cols.statusCol).setValue(newStatus);

  if (
    cols.managerCommentCol &&
    payload &&
    payload.managerComment !== undefined
  ) {
    sheet
      .getRange(rowNumber, cols.managerCommentCol)
      .setValue(stringValue(payload.managerComment));
  }

  var nowIso = new Date().toISOString();
  if (cols.decisionAtCol)
    sheet.getRange(rowNumber, cols.decisionAtCol).setValue(nowIso);
  if (cols.updatedAtCol)
    sheet.getRange(rowNumber, cols.updatedAtCol).setValue(nowIso);

  if (cols.decidedByNameCol && payload && payload.decidedByName) {
    sheet
      .getRange(rowNumber, cols.decidedByNameCol)
      .setValue(stringValue(payload.decidedByName));
  }
  if (cols.decidedByRoleCol && payload && payload.decidedByRole) {
    sheet
      .getRange(rowNumber, cols.decidedByRoleCol)
      .setValue(stringValue(payload.decidedByRole));
  }
  if (cols.decidedByIdCol && payload && payload.decidedById) {
    sheet
      .getRange(rowNumber, cols.decidedByIdCol)
      .setValue(stringValue(payload.decidedById));
  }

  if (payload && payload.archiveAfterDecision && cols.isArchivedCol) {
    sheet.getRange(rowNumber, cols.isArchivedCol).setValue(true);
  }

  var updatedRow = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  var mapped = mapRequestRow_(updatedRow, cols);

  logDuration_(lg, "admin.updateRequestStatus.total", startMs, {
    id: id,
    newStatus: newStatus,
    archived: payload && payload.archiveAfterDecision,
  });

  return { request: mapped };
}

function adminCreateRequest_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_REQUESTS_CREATE");
  var startMs = new Date().getTime();

  var employeeId = stringValue(payload && payload.employeeId);
  var requestTypeId = stringValue(payload && payload.requestTypeId);
  var requestTypeLabel = stringValue(payload && payload.requestTypeLabel);
  var dateFrom = stringValue(payload && payload.dateFrom);
  var dateTo = stringValue(payload && payload.dateTo);
  var singleDate = stringValue(payload && payload.singleDate);
  var shiftId = stringValue(payload && payload.shiftId);
  var shiftDate = stringValue(
    payload && (payload.shiftDate || payload.shiftRefDate),
  );

  if (!employeeId || !requestTypeId) {
    throw requestValidationError_("Missing employeeId or requestTypeId");
  }

  var hasDate = dateFrom || dateTo || singleDate || shiftDate || shiftId;
  if (!hasDate) throw requestValidationError_("Missing date for request");

  var cols = resolveRequestColumns_({ ensureOptional: true });
  var sheet = cols.sheet;
  var lastCol = sheet.getLastColumn();
  var rowNumber = sheet.getLastRow() + 1;
  var nowIso = new Date().toISOString();
  var generatedId =
    payload && payload.requestId
      ? stringValue(payload.requestId)
      : "REQ-" + Utilities.getUuid().split("-")[0];

  var values = new Array(lastCol).fill("");
  function setCell(col, value) {
    if (!col) return;
    values[col - 1] = value;
  }

  setCell(cols.statusCol, "PENDING");
  setCell(cols.requestIdCol || cols.shiftIdCol, generatedId);
  setCell(cols.employeeIdCol, employeeId);
  setCell(cols.employeeNameCol, stringValue(payload && payload.employeeName));
  setCell(
    cols.createdAtCol,
    singleDate || dateFrom || dateTo || shiftDate || nowIso,
  );
  setCell(cols.shiftDateCol, shiftDate || singleDate || dateFrom || "");
  setCell(cols.shiftIdCol, shiftId || "");
  setCell(cols.dateFromCol, dateFrom || "");
  setCell(cols.dateToCol, dateTo || "");
  setCell(cols.singleDateCol, singleDate || "");
  setCell(cols.requestTypeIdCol, requestTypeId);
  setCell(cols.requestTypeLabelCol, requestTypeLabel || requestTypeId);
  setCell(
    cols.employeeCommentCol,
    stringValue(payload && payload.employeeComment),
  );
  setCell(
    cols.sourceCol,
    stringValue(payload && payload.sourceOverride) || "ADMIN",
  );
  setCell(cols.isArchivedCol, false);
  setCell(cols.updatedAtCol, nowIso);

  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);

  var newRow = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var mapped = mapRequestRow_(newRow, cols);

  logDuration_(lg, "admin.createRequest.total", startMs, {
    employeeId: employeeId,
    requestTypeId: requestTypeId,
  });

  return { request: mapped };
}

function adminArchiveRequest_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("ADMIN_REQUESTS_ARCHIVE");
  var startMs = new Date().getTime();
  var id = stringValue(payload && (payload.id || payload.requestId));
  var archivedFlag = payload && payload.archived === true;

  if (!id) throw requestValidationError_("Missing requestId");

  var cols = resolveRequestColumns_({ ensureOptional: true });
  if (!cols.isArchivedCol) {
    throw requestValidationError_("Archive column missing");
  }

  var sheet = cols.sheet;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var rowIdx = findRequestRowIndex_(values, cols, id);
  if (rowIdx < 1) throw requestValidationError_("REQUEST_NOT_FOUND");
  var rowNumber = rowIdx + 1;

  sheet.getRange(rowNumber, cols.isArchivedCol).setValue(archivedFlag);
  if (cols.updatedAtCol) {
    sheet
      .getRange(rowNumber, cols.updatedAtCol)
      .setValue(new Date().toISOString());
  }

  var updatedRow = sheet
    .getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var mapped = mapRequestRow_(updatedRow, cols);

  logDuration_(lg, "admin.archiveRequest.total", startMs, {
    id: id,
    archived: archivedFlag,
  });

  return { request: mapped };
}

function listWorkLogsByEmployee_(payload, logger) {
  // SPEC-SHIFT-001
  var lg = logger || ensureModuleLoggerDefined_("WORK_LOG_LOAD");
  var startMs = new Date().getTime();
  const employeeId =
    payload && payload.employeeId ? String(payload.employeeId).trim() : "";
  if (!employeeId) throw new Error("Missing employeeId");
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const reportIdCol = getRequiredColumn_(
    headerMap,
    ["ID דיווח", "Shift ID"],
    sheetName,
  );
  const tsCol = getRequiredColumn_(headerMap, ["חותמת זמן"], sheetName);
  const empIdCol = getRequiredColumn_(
    headerMap,
    ["ID עובד", "מזהה עובד"],
    sheetName,
  );
  const empNameCol = getRequiredColumn_(headerMap, ["שם מלא"], sheetName);
  const dirCol = getRequiredColumn_(
    headerMap,
    ["כניסה / יציאה", "דיווח שעות"],
    sheetName,
  );
  const fixDateCol = getRequiredColumn_(headerMap, ["תיקון תאריך"], sheetName);
  const fixTimeCol = getRequiredColumn_(headerMap, ["תיקון שעה"], sheetName);
  const jobTypeIdCol = getRequiredColumn_(
    headerMap,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheetName,
  );
  const jobNameCol = getRequiredColumn_(
    headerMap,
    ["סוג עבודה", "סוגי עבודה"],
    sheetName,
  );
  const deptCol = getRequiredColumn_(headerMap, ["מחלקה", "מחלקות"], sheetName);
  const unitsCol = getRequiredColumn_(
    headerMap,
    ["כמות היחידות", "דיווח יחידות"],
    sheetName,
  );
  const notesCol = getRequiredColumn_(headerMap, ["הערות", "הערה"], sheetName);
  const values = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (stringValue(row[empIdCol - 1]) !== employeeId) continue;
    logs.push({
      shiftId: stringValue(row[reportIdCol - 1]),
      timestamp: stringValue(row[tsCol - 1]),
      employeeId: stringValue(row[empIdCol - 1]),
      employeeName: stringValue(row[empNameCol - 1]),
      direction: stringValue(row[dirCol - 1]),
      jobTypeId: stringValue(row[jobTypeIdCol - 1]),
      jobName: stringValue(row[jobNameCol - 1]),
      department: stringValue(row[deptCol - 1]),
      units: stringValue(row[unitsCol - 1]),
      notes: stringValue(row[notesCol - 1]),
      fixDate: stringValue(row[fixDateCol - 1]),
      fixTime: stringValue(row[fixTimeCol - 1]),
    });
  }
  logDuration_(lg, "worklogs.list", startMs, {
    employeeId: employeeId,
    rowsScanned: Math.max(values.length - 1, 0),
    returned: logs.length,
  });
  return { workLogs: logs };
}

function listRequestsByEmployee_(payload, logger) {
  // SPEC-REQ-001
  var lg = logger || ensureModuleLoggerDefined_("REQUEST_LIST");
  var startMs = new Date().getTime();
  const employeeId =
    payload && payload.employeeId ? String(payload.employeeId).trim() : "";
  if (!employeeId) throw new Error("Missing employeeId");
  const employeesSheet = getEmployeesSheet_();
  const empHeaders = getHeaderMap_(employeesSheet);
  const empSheetName = employeesSheet.getName();
  const empIdCol = getRequiredColumn_(
    empHeaders,
    ["מזהה עובד", "ID עובד"],
    empSheetName,
  );
  const tzCol = getRequiredColumn_(empHeaders, ["ת.ז", "תז"], empSheetName);
  const nameCol = getRequiredColumn_(empHeaders, ["שם מלא"], empSheetName);
  const empRows = employeesSheet.getDataRange().getValues();
  let tzValue = "";
  let fullName = "";
  for (let i = 1; i < empRows.length; i++) {
    const row = empRows[i];
    if (stringValue(row[empIdCol - 1]) === employeeId) {
      tzValue = stringValue(row[tzCol - 1]);
      fullName = stringValue(row[nameCol - 1]);
      break;
    }
  }
  if (!tzValue && lg && lg.warn) {
    lg.warn("employee-lookup", { employeeId: employeeId, missing: "tz" });
  }

  const reqSheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
  const reqHeaders = getHeaderMap_(reqSheet);
  const reqSheetName = reqSheet.getName();
  const reqEmpCol = getRequiredColumn_(
    reqHeaders,
    ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
    reqSheetName,
  );
  const statusCol = getRequiredColumn_(
    reqHeaders,
    ["סטטוס", "סטטוס בקשה"],
    reqSheetName,
  );
  const requestIdCol = getOptionalColumn_(reqHeaders, [
    "ID בקשה",
    "Request ID",
    "RequestID",
  ]);
  const jobNameCol = getRequiredColumn_(
    reqHeaders,
    ["סוג עבודה", "סוגי עבודה"],
    reqSheetName,
  );
  const workDateCol = getOptionalColumn_(reqHeaders, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const directionCol = getOptionalColumn_(reqHeaders, [
    "דיווח שעות",
    "כניסה / יציאה",
  ]);
  const unitsCol = getOptionalColumn_(reqHeaders, [
    "כמות היחידות",
    "דיווח יחידות",
  ]);
  const noteCol = getOptionalColumn_(reqHeaders, [
    "הערה למנהל",
    "הערות",
    "הערות למשמרת",
  ]);
  const submittedCol = getOptionalColumn_(reqHeaders, [
    "תאריך נשלח",
    "חותמת זמן",
  ]);
  const decidedCol = getOptionalColumn_(reqHeaders, ["תאריך החלטה"]);
  const shiftIdCol = getRequiredColumn_(reqHeaders, ["ID משמרת"], reqSheetName);
  const requestTypeCol = getOptionalColumn_(reqHeaders, ["סוג בקשה"]);
  const jobTypeIdCol = getRequiredColumn_(
    reqHeaders,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    reqSheetName,
  );
  const fixDateCol = getOptionalColumn_(reqHeaders, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(reqHeaders, ["תיקון שעה"]);
  const payTypeCol = getOptionalColumn_(reqHeaders, [
    "אופן תשלום",
    "אופני תשלום",
  ]);
  const payTypeIdCol = getOptionalColumn_(reqHeaders, [
    "ID אופן תשלום",
    "ID אופני תשלום",
  ]);

  const jobTypes = listJobTypes_();
  const jobTypeNames = jobTypes.map((j) => j.name);

  const values = reqSheet.getDataRange().getValues();
  const requests = [];

  function pad2_(n) {
    return n < 10 ? "0" + n : String(n);
  }

  function normalizeTimeString_(val) {
    if (val === null || val === undefined || val === "") return "";
    const asDate = val instanceof Date ? val : new Date(val);
    if (!isNaN(asDate.getTime())) {
      return pad2_(asDate.getHours()) + ":" + pad2_(asDate.getMinutes());
    }
    const s = stringValue(val);
    const hhmm = s.match(/(\d{1,2}):(\d{2})/);
    if (hhmm) return pad2_(Number(hhmm[1])) + ":" + pad2_(Number(hhmm[2]));
    return "";
  }

  function normalizeIsoDateTime_(primary, datePart, timePart) {
    var candidate = primary instanceof Date ? primary : new Date(primary || "");
    if (!isNaN(candidate.getTime())) return candidate.toISOString();

    const baseDate = toIsoDate_(datePart || "");
    if (!baseDate) return "";

    const normalizedTime = normalizeTimeString_(timePart || "");
    if (normalizedTime) {
      const joined = new Date(baseDate + "T" + normalizedTime + ":00");
      if (!isNaN(joined.getTime())) return joined.toISOString();
    }

    const midnight = new Date(baseDate + "T00:00:00");
    return isNaN(midnight.getTime()) ? "" : midnight.toISOString();
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmpValue = stringValue(row[reqEmpCol - 1]);
    if (rowEmpValue !== tzValue && rowEmpValue !== employeeId) continue;
    const requestId = requestIdCol ? stringValue(row[requestIdCol - 1]) : "";
    const requestKey = requestId || stringValue(row[shiftIdCol - 1]);
    const jobName = stringValue(row[jobNameCol - 1]);
    let requestType = requestTypeCol
      ? stringValue(row[requestTypeCol - 1])
      : "";
    if (!requestType) {
      requestType =
        jobTypeNames.indexOf(jobName) >= 0 ? "job_not_linked" : "new_job_type";
    }
    const workDateRaw = workDateCol ? row[workDateCol - 1] : "";
    const submittedAt = normalizeIsoDateTime_(
      submittedCol ? row[submittedCol - 1] : "",
      workDateRaw,
      fixTimeCol ? row[fixTimeCol - 1] : "",
    );
    const requestedSummary =
      stringValue(row[directionCol - 1]) ||
      (unitsCol ? stringValue(row[unitsCol - 1]) : "");
    const noteToManager = noteCol ? stringValue(row[noteCol - 1]) : "";
    const rawPayType = payTypeCol ? stringValue(row[payTypeCol - 1]) : "";
    const payType = normalizePayType_(rawPayType);
    let correction = null;
    if (requestType === "shift_correction" && noteToManager) {
      try {
        const parsed = JSON.parse(noteToManager);
        if (parsed && typeof parsed === "object") {
          correction = parsed;
        }
      } catch (err) {
        correction = null;
      }
    }
    const workDate = workDateCol ? stringValue(row[workDateCol - 1]) : "";

    const fixDate = fixDateCol ? toIsoDate_(row[fixDateCol - 1]) : "";
    const fixTime = fixTimeCol ? normalizeTimeString_(row[fixTimeCol - 1]) : "";

    requests.push({
      id: requestKey,
      status: stringValue(row[statusCol - 1]),
      jobName: jobName,
      workDate: workDate,
      requestedSummary: requestedSummary,
      units: unitsCol ? stringValue(row[unitsCol - 1]) : "",
      noteToManager: noteToManager,
      submittedAt: submittedAt,
      decidedAt: normalizeIsoDateTime_(
        decidedCol ? row[decidedCol - 1] : "",
        workDateRaw,
        fixTimeCol ? row[fixTimeCol - 1] : "",
      ),
      type: requestType,
      employeeId: employeeId,
      employeeName: fullName,
      shiftId: stringValue(row[shiftIdCol - 1]),
      requestId: requestId,
      jobTypeId: stringValue(row[jobTypeIdCol - 1]),
      fixDate: fixDate,
      fixTime: fixTime,
      payType: payType,
      payTypeLabel: rawPayType,
      payTypeId: payTypeIdCol ? stringValue(row[payTypeIdCol - 1]) : "",
      createdAt: submittedAt,
      description:
        jobName && workDate ? jobName + " • " + workDate : jobName || workDate,
      correction: correction,
    });
  }
  logDuration_(lg, "requests.list", startMs, {
    employeeId: employeeId,
    rowsScanned: Math.max(values.length - 1, 0),
    returned: requests.length,
  });
  return { requests };
}

function findRequestById_(requestId) {
  const reqSheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
  const headers = getHeaderMap_(reqSheet);
  const reqSheetName = reqSheet.getName();
  const shiftIdCol = getRequiredColumn_(headers, ["ID משמרת"], reqSheetName);
  const statusCol = getRequiredColumn_(
    headers,
    ["סטטוס", "סטטוס בקשה"],
    reqSheetName,
  );
  const typeCol = getOptionalColumn_(headers, ["סוג בקשה"]);
  const empIdCol = getOptionalColumn_(headers, [
    "מזהה עובד",
    "ID עובד",
    "ת.ז",
    "תז",
  ]);
  const empNameCol = getOptionalColumn_(headers, ["שם מלא"]);
  const jobIdCol = getOptionalColumn_(headers, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const jobNameCol = getOptionalColumn_(headers, ["סוג עבודה", "סוגי עבודה"]);
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const workDateCol = getOptionalColumn_(headers, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const directionCol = getOptionalColumn_(headers, [
    "דיווח שעות",
    "כניסה / יציאה",
  ]);
  const fixDateCol = getOptionalColumn_(headers, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(headers, ["תיקון שעה"]);
  const unitsCol = getOptionalColumn_(headers, [
    "כמות היחידות",
    "דיווח יחידות",
  ]);
  const noteCol = getOptionalColumn_(headers, [
    "הערה למנהל",
    "הערות",
    "הערות למשמרת",
  ]);
  const payTypeCol = getOptionalColumn_(headers, ["אופן תשלום", "אופני תשלום"]);
  const payTypeIdCol = getOptionalColumn_(headers, [
    "ID אופן תשלום",
    "ID אופני תשלום",
  ]);
  const decidedAtCol = getOptionalColumn_(headers, ["תאריך החלטה"]);

  const values = reqSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (stringValue(row[shiftIdCol - 1]) !== stringValue(requestId)) continue;
    return {
      rowIndex: i + 1,
      headers: headers,
      sheet: reqSheet,
      record: {
        requestId: stringValue(row[shiftIdCol - 1]),
        status: statusCol ? stringValue(row[statusCol - 1]) : "",
        requestType: typeCol ? stringValue(row[typeCol - 1]) : "",
        employeeId: empIdCol ? stringValue(row[empIdCol - 1]) : "",
        employeeName: empNameCol ? stringValue(row[empNameCol - 1]) : "",
        jobTypeId: jobIdCol ? stringValue(row[jobIdCol - 1]) : "",
        jobName: jobNameCol ? stringValue(row[jobNameCol - 1]) : "",
        department: deptCol ? stringValue(row[deptCol - 1]) : "",
        workDate: workDateCol ? stringValue(row[workDateCol - 1]) : "",
        direction: directionCol ? stringValue(row[directionCol - 1]) : "",
        fixDate: fixDateCol ? stringValue(row[fixDateCol - 1]) : "",
        fixTime: fixTimeCol ? stringValue(row[fixTimeCol - 1]) : "",
        units: unitsCol ? row[unitsCol - 1] : "",
        noteToManager: noteCol ? stringValue(row[noteCol - 1]) : "",
        payType: payTypeCol ? stringValue(row[payTypeCol - 1]) : "",
        payTypeId: payTypeIdCol ? stringValue(row[payTypeIdCol - 1]) : "",
        decidedAtCol: decidedAtCol,
        statusCol: statusCol,
      },
    };
  }
  return null;
}

function handleShiftReportSubmit_(payload, logger) {
  if (!payload) throw new Error("Missing payload for shiftReport.submit");
  var fnStartMs = new Date().getTime();
  const nowIso = new Date().toISOString();

  if (logger) {
    logger.info("shift.submit.start", {
      reportMode: payload && payload.reportMode,
      employeeId: payload && payload.employeeId,
      rawShiftId: payload && payload.shiftId,
    });
  }
  const shiftId = payload.shiftId || Utilities.getUuid();
  const employeeId = stringValue(payload.employeeId);
  if (!employeeId) throw new Error("Missing employeeId");

  const employeesSheet = getEmployeesSheet_();
  const empHeaders = getHeaderMap_(employeesSheet);
  const empSheetName = employeesSheet.getName();
  const empIdCol = getRequiredColumn_(
    empHeaders,
    ["מזהה עובד", "ID עובד"],
    empSheetName,
  );
  const nameCol = getRequiredColumn_(empHeaders, ["שם מלא"], empSheetName);
  const tzCol = getRequiredColumn_(empHeaders, ["ת.ז", "תז"], empSheetName);
  const emailCol = getOptionalColumn_(empHeaders, EMAIL_HEADER_CANDIDATES);
  var employeeScanStartMs = new Date().getTime();
  const empValues = employeesSheet.getDataRange().getValues();
  let employeeName = "";
  let employeeTz = "";
  let employeeEmail = "";
  for (let i = 1; i < empValues.length; i++) {
    const row = empValues[i];
    if (stringValue(row[empIdCol - 1]) === employeeId) {
      employeeName = stringValue(row[nameCol - 1]);
      employeeTz = stringValue(row[tzCol - 1]);
      if (emailCol && !employeeEmail) {
        employeeEmail = stringValue(row[emailCol - 1]);
      }
      break;
    }
  }
  if (!employeeName) throw new Error("Employee not found: " + employeeId);
  logDuration_(logger, "shift.submit.resolve-employee", employeeScanStartMs, {
    sheet: empSheetName,
    found: !!employeeName,
  });

  const jobTypes = listJobTypes_();
  const jobMap = {};
  jobTypes.forEach((j) => (jobMap[j.id] = j));
  var linkedStartMs = new Date().getTime();
  const linked = listEmployeeLinkedJobIds_({ employeeId }).jobTypeIds;

  const isOther = payload.jobId === "__other__";
  const jobTypeId = isOther ? "" : stringValue(payload.jobId);
  const job = isOther ? null : jobMap[jobTypeId];
  const jobNameInput = stringValue(payload.jobName);
  const jobName = isOther
    ? jobNameInput
    : job
      ? stringValue(job.name)
      : jobNameInput;
  if (!jobName) throw new Error("Missing jobName");
  const department = isOther ? "" : job ? stringValue(job.department) : "";
  const isLinked = jobTypeId ? linked.indexOf(jobTypeId) >= 0 : false;
  const direction = normalizeDirectionToHebrew_(payload.direction);
  const fixDate = stringValue(payload.fixDate);
  const fixTime = stringValue(payload.fixTime);
  const units =
    payload.units !== null && payload.units !== undefined ? payload.units : "";
  const timestamp = stringValue(payload.timestamp) || nowIso;
  const workDate = stringValue(payload.workDate) || timestamp.split("T")[0];

  const employeePay = jobTypeId
    ? resolveEmployeeJobPayType_(
        employeesSheet,
        empHeaders,
        employeeId,
        jobTypeId,
      )
    : null;
  const payTypeName = employeePay
    ? employeePay.payTypeName
    : job
      ? stringValue(job.payTypeName)
      : "";
  const payTypeId = employeePay
    ? employeePay.payTypeId
    : job
      ? stringValue(job.payTypeId)
      : "";
  const payType = normalizePayType_(payTypeName);
  const payTypeForRequestRow = payType || payTypeName;
  logDuration_(logger, "shift.submit.resolve-job", linkedStartMs, {
    jobTypeCount: jobTypes.length,
    linkedCount: linked.length,
    isOther: isOther,
    payType: payType || payTypeName,
  });

  // Hourly/directional reports must include a valid direction; otherwise we
  // would write an ambiguous log that later surfaces as "כיוון דיווח לא מזוהה".
  const needsDirection =
    payType !== "unit" && payType !== "daily" && payType !== "monthly";
  if (needsDirection) {
    if (!direction) {
      throw new Error("Missing direction for directional report");
    }
    if (direction !== "כניסה" && direction !== "יציאה") {
      throw new Error("Invalid direction value: " + direction);
    }
  }

  if (payType === "monthly") {
    return {
      ok: true,
      status: "monthly_no_report",
      message: "סוג העבודה הזה לא דורש דיווח. אם זו טעות, לחץ כאן.",
    };
  }

  const timestampDate = timestamp.split("T")[0];
  const reportMode = stringValue(payload.mode || payload.reportMode);
  const manualDate = stringValue(payload.manualDate || payload.fixDate);
  const manualTime = stringValue(payload.manualTime || payload.fixTime);
  const isManualHourly =
    payType === "hourly" &&
    (reportMode.toLowerCase() === "manual" || !!manualDate || !!manualTime);

  var normalizeStartMs = new Date().getTime();
  let normalized = {
    workDate: workDate || timestampDate,
    direction: direction,
    fixDate: fixDate,
    fixTime: fixTime,
    units: units,
  };

  if (payType === "hourly") {
    normalized.units = "";
    if (isManualHourly) {
      normalized.fixDate = manualDate || workDate || timestampDate;
      normalized.fixTime = manualTime;
      normalized.workDate = normalized.fixDate || timestampDate;
    } else {
      normalized.fixDate = "";
      normalized.fixTime = "";
      normalized.workDate = timestampDate;
    }
  } else if (payType === "unit") {
    normalized.direction = "";
    normalized.fixDate = "";
    normalized.fixTime = "";
    normalized.workDate = workDate || timestampDate;
  } else if (payType === "daily") {
    normalized.direction = "";
    normalized.fixDate = "";
    normalized.fixTime = "";
    normalized.units = "";
    normalized.workDate = workDate || timestampDate;
  }

  if (payType === "unit" && normalized.units !== "") {
    const parsedUnits = Number(normalized.units);
    normalized.units = isNaN(parsedUnits) ? normalized.units : parsedUnits;
  }
  logDuration_(logger, "shift.submit.normalize", normalizeStartMs, {
    payType: payType || payTypeName,
    isManualHourly: isManualHourly,
    mode: reportMode.toLowerCase ? reportMode.toLowerCase() : reportMode,
  });

  const explicitRequestType = stringValue(payload.requestType);
  let requestType =
    explicitRequestType &&
    (explicitRequestType === "new_job_type" ||
      explicitRequestType === "job_not_linked")
      ? explicitRequestType
      : isOther
        ? "new_job_type"
        : "job_not_linked";

  const needsApproval =
    payload.requiresApproval === true || isOther || !isLinked;

  if (needsApproval) {
    const reqSheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
    const headers = getHeaderMap_(reqSheet);
    const reqSheetName = reqSheet.getName();
    getRequiredColumn_(headers, ["ID משמרת"], reqSheetName);
    getRequiredColumn_(headers, ["סטטוס"], reqSheetName);
    getRequiredColumn_(
      headers,
      ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
      reqSheetName,
    );
    getRequiredColumn_(headers, ["ID משמרת"], reqSheetName);
    getRequiredColumn_(headers, ["סטטוס", "סטטוס בקשה"], reqSheetName);
    getRequiredColumn_(
      headers,
      ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
      reqSheetName,
    );
    getRequiredColumn_(headers, ["שם מלא"], reqSheetName);
    getRequiredColumn_(headers, ["סוג עבודה", "סוגי עבודה"], reqSheetName);
    getRequiredColumn_(
      headers,
      ["ID סוג עבודה", "ID סוגי עבודה"],
      reqSheetName,
    );
    getOptionalColumn_(headers, ["תאריך משמרת", "תיקון תאריך", "חותמת זמן"]);
    getOptionalColumn_(headers, ["דיווח שעות", "כניסה / יציאה"]);
    getOptionalColumn_(headers, ["תיקון תאריך"]);
    getOptionalColumn_(headers, ["תיקון שעה"]);
    getOptionalColumn_(headers, ["כמות היחידות", "דיווח יחידות"]);
    getOptionalColumn_(headers, ["תאריך נשלח", "חותמת זמן"]);
    getOptionalColumn_(headers, ["הערה למנהל", "הערות", "הערות למשמרת"]);
    getOptionalColumn_(headers, ["סוג בקשה"]);
    getRequiredColumn_(headers, ["סוג בקשה"], reqSheetName);
    var requestAppendStartMs = new Date().getTime();
    const row = buildRowFromHeaders_(headers, {
      id: shiftId,
      shiftId: shiftId,
      status: "ממתין",
      employeeId: employeeId,
      employeeName: employeeName,
      nationalId: employeeTz,
      jobName: jobName,
      jobTypeId: jobTypeId,
      workDate: normalized.workDate,
      direction: normalized.direction,
      fixDate: normalized.fixDate,
      fixTime: normalized.fixTime,
      units: normalized.units,
      noteToManager: stringValue(
        payload.noteToManager || payload.note || payload.shiftNote,
      ),
      submittedAt: nowIso,
      timestamp: nowIso,
      decidedAt: "",
      managerDecision: "",
      requestType: requestType,
      payTypeId: payTypeId,
      payType: payTypeForRequestRow,
    });
    reqSheet.appendRow(row);
    logDuration_(logger, "shift.submit.request-append", requestAppendStartMs, {
      sheet: reqSheetName,
      requestType: requestType,
    });
    logDuration_(logger, "shift.submit.total", fnStartMs, {
      requiresApproval: true,
      jobTypeId: jobTypeId,
      payType: payType || payTypeName,
    });
    return {
      ok: true,
      success: true,
      requiresApproval: true,
      shiftId: shiftId,
      requestType: requestType,
      status: "saved_as_request",
    };
  }

  normalized.shiftId = shiftId;
  normalized.jobTypeId = jobTypeId;
  normalized.jobName = jobName;
  normalized.department = department;
  normalized.timestamp = timestamp;

  var writeStartMs = new Date().getTime();
  const writeResult = writeWorkLogFromNormalizedShift_(
    normalized,
    {
      employeeId: employeeId,
      employeeName: employeeName,
      note: payload.note,
    },
    logger,
  );
  logDuration_(logger, "shift.submit.write-worklog", writeStartMs, {
    jobTypeId: jobTypeId,
    payType: payType || payTypeName,
    direction: normalized.direction,
  });

  // Detect duplicates after writing (to include the new row in the list).
  var duplicateScanStartMs = new Date().getTime();
  const eventMsForCriteria = extractEventMs_(
    normalized.workDate,
    normalized.timestamp,
    normalized.fixDate,
    normalized.fixTime,
  );
  const duplicates = findDuplicateWorkLogs_(
    {
      employeeId: employeeId,
      jobTypeId: jobTypeId,
      jobName: jobName,
      department: department,
      direction: normalized.direction,
      workDate: normalized.workDate,
      fixDate: normalized.fixDate,
      fixTime: normalized.fixTime,
      timestamp: normalized.timestamp,
      eventMs: eventMsForCriteria,
    },
    shiftId,
  );
  logDuration_(logger, "shift.submit.duplicates-scan", duplicateScanStartMs, {
    count: duplicates.length,
  });
  const duplicatesFound = duplicates.length >= 2;

  appendSystemLog_({
    operation: "WORK_LOG_DUP_SCAN",
    step: "scan",
    severity: "info",
    errorCode: null,
    details: {
      foundCount: duplicates.length,
      includeShiftId: Boolean(shiftId),
      timeWindowMs: DUPLICATE_TIME_WINDOW_MS,
    },
  });

  if (duplicatesFound) {
    const cleanup = cleanupExactDuplicateWorkLogs_(duplicates);
    if (cleanup.deletedCount > 0) {
      logDuration_(logger, "shift.submit.total", fnStartMs, {
        requiresApproval: false,
        duplicatesFound: true,
        autoDeleted: cleanup.deletedCount,
        jobTypeId: jobTypeId,
        payType: payType || payTypeName,
      });
      return {
        ok: true,
        status: "duplicate_auto_deleted",
        shiftId: shiftId,
        deletedCount: cleanup.deletedCount,
        capped: cleanup.capped === true,
      };
    }

    logDuration_(logger, "shift.submit.total", fnStartMs, {
      requiresApproval: false,
      duplicatesFound: true,
      jobTypeId: jobTypeId,
      payType: payType || payTypeName,
    });
    return {
      ok: true,
      status: "duplicate_found",
      shiftId: shiftId,
      duplicates: duplicates,
    };
  }

  logDuration_(logger, "shift.submit.total", fnStartMs, {
    requiresApproval: false,
    duplicatesFound: false,
    jobTypeId: jobTypeId,
    payType: payType || payTypeName,
  });

  return writeResult;
}

function handleShiftReportMonthlyErrorNotify_(payload) {
  const employeeId = stringValue(payload.employeeId);
  const employeeEmail = stringValue(payload.email || payload.employeeEmail);
  let employeePhone = stringValue(
    payload.phone ||
      payload.employeePhone ||
      payload.mobile ||
      payload.employeeMobile,
  );
  const jobTypeId = stringValue(payload.jobTypeId);
  const jobName = stringValue(payload.jobName);

  if (!employeeId && !employeeEmail) {
    return { ok: false, error: "missing_employee_identifier" };
  }

  let resolvedName = stringValue(payload.employeeName);
  let resolvedEmail = employeeEmail;

  if (employeeId) {
    const sheet = getEmployeesSheet_();
    const headers = getHeaderMap_(sheet);
    const sheetName = sheet.getName();
    const empIdCol = getRequiredColumn_(
      headers,
      ["מזהה עובד", "ID עובד"],
      sheetName,
    );
    const nameCol = getOptionalColumn_(headers, ["שם מלא"]);
    const emailCol = getOptionalColumn_(headers, EMAIL_HEADER_CANDIDATES);
    const phoneCol = getOptionalColumn_(headers, [
      "טלפון",
      "פלאפון",
      "נייד",
      "phone",
      "mobile",
    ]);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (stringValue(rows[i][empIdCol - 1]) !== employeeId) continue;
      if (!resolvedName && nameCol)
        resolvedName = stringValue(rows[i][nameCol - 1]);
      if (!resolvedEmail && emailCol)
        resolvedEmail = stringValue(rows[i][emailCol - 1]);
      if (!employeePhone && phoneCol)
        employeePhone = stringValue(rows[i][phoneCol - 1]);
      break;
    }
  }

  sendMonthlyBlockEmail_({
    employeeId: employeeId,
    employeeName: resolvedName,
    employeeEmail: resolvedEmail,
    employeePhone: employeePhone,
    jobTypeId: jobTypeId,
    jobName: jobName,
    timestamp: new Date().toISOString(),
  });

  return { ok: true, status: "monthly_error_notified" };
}

function handleShiftCorrectionSubmit_(payload, logger) {
  if (!payload) throw new Error("Missing payload for shiftCorrection.submit");
  var fnStartMs = new Date().getTime();
  const nowIso = new Date().toISOString();
  const employeeId = stringValue(payload.employeeId);
  if (!employeeId) throw new Error("Missing employeeId");

  const employeesSheet = getEmployeesSheet_();
  const empHeaders = getHeaderMap_(employeesSheet);
  const empSheetName = employeesSheet.getName();
  const empIdCol = getRequiredColumn_(
    empHeaders,
    ["מזהה עובד", "ID עובד"],
    empSheetName,
  );
  const nameCol = getRequiredColumn_(empHeaders, ["שם מלא"], empSheetName);
  const tzCol = getRequiredColumn_(empHeaders, ["ת.ז", "תז"], empSheetName);
  var employeeScanStartMs = new Date().getTime();
  const empValues = employeesSheet.getDataRange().getValues();
  let employeeName = "";
  let employeeTz = "";
  for (let i = 1; i < empValues.length; i++) {
    const row = empValues[i];
    if (stringValue(row[empIdCol - 1]) === employeeId) {
      employeeName = stringValue(row[nameCol - 1]);
      employeeTz = stringValue(row[tzCol - 1]);
      break;
    }
  }
  if (!employeeName) throw new Error("Employee not found: " + employeeId);
  logDuration_(
    logger,
    "shift.correction.resolve-employee",
    employeeScanStartMs,
    {
      sheet: empSheetName,
      found: !!employeeName,
    },
  );

  const original = {
    workDate: stringValue(payload.originalWorkDate || payload.workDate),
    startTime: stringValue(payload.originalStartTime),
    endTime: stringValue(payload.originalEndTime),
    direction: normalizeDirectionToHebrew_(payload.originalDirection),
    jobTypeId: stringValue(payload.originalJobTypeId),
    units:
      payload.originalUnits !== undefined && payload.originalUnits !== null
        ? payload.originalUnits
        : "",
    shiftId: stringValue(payload.originalShiftId || payload.shiftId),
  };
  const updated = {
    workDate: stringValue(payload.newWorkDate),
    startTime: stringValue(payload.newStartTime),
    endTime: stringValue(payload.newEndTime),
    direction: normalizeDirectionToHebrew_(payload.newDirection),
    jobTypeId: stringValue(payload.newJobTypeId),
    units:
      payload.newUnits !== undefined && payload.newUnits !== null
        ? payload.newUnits
        : "",
  };

  const comparisonKeys = [
    "workDate",
    "startTime",
    "endTime",
    "direction",
    "jobTypeId",
    "units",
  ];
  const hasChange = comparisonKeys.some(function (key) {
    return stringValue(original[key]) !== stringValue(updated[key]);
  });
  if (!hasChange) {
    return { ok: false, error: "no_changes" };
  }

  const reqSheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
  const headers = getHeaderMap_(reqSheet);
  const reqSheetName = reqSheet.getName();

  getRequiredColumn_(headers, ["ID משמרת"], reqSheetName);
  getRequiredColumn_(headers, ["סטטוס", "סטטוס בקשה"], reqSheetName);
  getRequiredColumn_(
    headers,
    ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
    reqSheetName,
  );
  getRequiredColumn_(headers, ["שם מלא"], reqSheetName);
  getRequiredColumn_(headers, ["סוג עבודה", "סוגי עבודה"], reqSheetName);
  getRequiredColumn_(headers, ["ID סוג עבודה", "ID סוגי עבודה"], reqSheetName);
  getRequiredColumn_(headers, ["סוג בקשה"], reqSheetName);

  const newShiftId = original.shiftId || Utilities.getUuid();
  const jobTypeId = updated.jobTypeId || original.jobTypeId;
  const jobName = stringValue(
    payload.jobName || payload.newJobName || payload.originalJobName,
  );
  const workDate = updated.workDate || original.workDate;
  const direction =
    normalizeDirectionToHebrew_(updated.direction) ||
    normalizeDirectionToHebrew_(original.direction);
  const units = updated.units !== "" ? updated.units : original.units;

  const noteToManager = JSON.stringify({
    original: original,
    updated: updated,
  });

  const row = buildRowFromHeaders_(headers, {
    id: newShiftId,
    shiftId: newShiftId,
    status: "ממתין",
    employeeId: employeeId,
    employeeName: employeeName,
    nationalId: employeeTz,
    jobName: jobName,
    jobTypeId: jobTypeId,
    workDate: workDate,
    direction: direction,
    units: units,
    noteToManager: noteToManager,
    submittedAt: nowIso,
    timestamp: nowIso,
    decidedAt: "",
    managerDecision: "",
    requestType: "shift_correction",
    originalWorkDate: original.workDate,
    newWorkDate: updated.workDate,
    originalDirection: original.direction,
    newDirection: updated.direction,
    originalUnits: original.units,
    newUnits: updated.units,
    originalJobTypeId: original.jobTypeId,
    newJobTypeId: updated.jobTypeId,
  });

  var requestAppendStartMs = new Date().getTime();
  reqSheet.appendRow(row);

  logDuration_(
    logger,
    "shift.correction.append-request",
    requestAppendStartMs,
    {
      sheet: reqSheetName,
    },
  );
  logDuration_(logger, "shift.correction.total", fnStartMs, {
    requestType: "shift_correction",
    jobTypeId: jobTypeId,
  });

  return {
    ok: true,
    status: "saved_as_request",
    requestType: "shift_correction",
    shiftId: newShiftId,
  };
}

function handleRequestApprove_(payload, logger) {
  var lg = logger || ensureModuleLoggerDefined_("REQUEST_DECIDE");
  var startMs = new Date().getTime();
  const requestId = stringValue(
    payload.requestId || payload.shiftId || payload.id,
  );
  if (!requestId) throw new Error("Missing requestId");

  const found = findRequestById_(requestId);
  if (!found) return { ok: false, error: "request_not_found" };

  const rec = found.record;
  if (rec.requestType !== "job_not_linked") {
    return { ok: false, error: "unsupported_request_type" };
  }

  const updated = payload.updated || payload;
  const jobTypeId = stringValue(updated.jobTypeId || rec.jobTypeId);
  const jobName = stringValue(updated.jobName || rec.jobName);
  const department = stringValue(updated.department || rec.department);
  const workDate = stringValue(updated.workDate || rec.workDate);
  const direction = stringValue(updated.direction || rec.direction);
  const fixDate = stringValue(updated.fixDate || rec.fixDate);
  const fixTime = stringValue(updated.fixTime || rec.fixTime);
  const units =
    updated.units !== undefined && updated.units !== null
      ? updated.units
      : rec.units;
  const payTypeName = stringValue(updated.payType || rec.payType);
  const payTypeId = stringValue(updated.payTypeId || rec.payTypeId);

  const payType = normalizePayType_(payTypeName);

  linkJobTypeToEmployee_(
    rec.employeeId,
    rec.employeeName,
    jobTypeId,
    jobName,
    department,
    payTypeId,
    payTypeName,
  );

  const normalized = normalizeShiftForWrite_({
    payType: payType,
    timestamp: new Date().toISOString(),
    workDate: workDate,
    direction: direction,
    fixDate: fixDate,
    fixTime: fixTime,
    units: units,
    reportMode: updated.reportMode,
    manualDate: updated.manualDate,
    manualTime: updated.manualTime,
  });

  normalized.shiftId = rec.requestId || Utilities.getUuid();
  normalized.jobTypeId = jobTypeId;
  normalized.jobName = jobName;
  normalized.department = department;
  normalized.timestamp = new Date().toISOString();

  writeWorkLogFromNormalizedShift_(
    normalized,
    {
      employeeId: rec.employeeId,
      employeeName: rec.employeeName,
      note: updated.note,
    },
    lg,
  );

  if (rec.statusCol) {
    found.sheet.getRange(found.rowIndex, rec.statusCol).setValue("approved");
  }
  if (rec.decidedAtCol) {
    found.sheet
      .getRange(found.rowIndex, rec.decidedAtCol)
      .setValue(new Date().toISOString());
  }

  logDuration_(lg, "requests.approve", startMs, {
    requestId: requestId,
    jobTypeId: jobTypeId,
    employeeId: rec.employeeId,
  });
  return { ok: true, status: "approved", requestType: "job_not_linked" };
}

function handleEmployeeSave_(payload) {
  const employeeId = stringValue(payload.employeeId);
  if (!employeeId) throw new Error("Missing employeeId");

  const forbiddenName =
    payload.hasOwnProperty("fullName") ||
    payload.hasOwnProperty("name") ||
    payload.hasOwnProperty("employeeName");
  if (forbiddenName) return { ok: false, error: "fullName_read_only" };

  const sheet = getEmployeesSheet_();
  let headers = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const empIdCol = getRequiredColumn_(
    headers,
    ["מזהה עובד", "ID עובד"],
    sheetName,
  );
  const phoneCol = getOptionalColumn_(headers, ["טלפון"]);
  const birthCol = getOptionalColumn_(headers, ["ת. לידה", "תאריך לידה"]);

  const sizeInfo = ensureEmployeeSizeColumns_(sheet, headers);
  headers = sizeInfo.headers;
  const sizeCol = sizeInfo.sizeCol;
  const sizeSourceCol = sizeInfo.sourceCol;

  const values = sheet.getDataRange().getValues();
  let targetRow = null;
  for (let i = 1; i < values.length; i++) {
    if (stringValue(values[i][empIdCol - 1]) === employeeId) {
      targetRow = i + 1;
      break;
    }
  }
  if (!targetRow) return { ok: false, error: "employee_not_found" };

  if (payload.phone !== undefined && phoneCol) {
    sheet.getRange(targetRow, phoneCol).setValue(stringValue(payload.phone));
  }
  if (payload.birthDate !== undefined && birthCol) {
    sheet
      .getRange(targetRow, birthCol)
      .setValue(stringValue(payload.birthDate));
  }
  if (payload.size !== undefined && payload.size !== null) {
    sheet.getRange(targetRow, sizeCol).setValue(stringValue(payload.size));
    sheet.getRange(targetRow, sizeSourceCol).setValue("webapp");
  }

  const response = {
    id: employeeId,
    phone: payload.phone !== undefined ? stringValue(payload.phone) : undefined,
    birthDate:
      payload.birthDate !== undefined
        ? stringValue(payload.birthDate)
        : undefined,
    size:
      payload.size !== undefined && payload.size !== null
        ? stringValue(payload.size)
        : undefined,
    sizeSource:
      payload.size !== undefined && payload.size !== null
        ? "webapp"
        : undefined,
  };

  return { ok: true, status: "updated", employee: response };
}

// Returns employee match by email with normalized systemRole for auth.
function employeeExistsByEmail_(payload) {
  const email = stringValue(
    payload && (payload.email || payload.mail),
  ).toLowerCase();

  if (!email) {
    return { ok: false, success: false, error: "missing email" };
  }

  const sheet = getEmployeesSheet_();
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();

  const colEmail = getRequiredColumn_(
    headerMap,
    EMAIL_HEADER_CANDIDATES,
    sheetName,
  );
  const colStatus = getOptionalColumn_(headerMap, [
    "סטטוס",
    "סטטוס פעיל",
    "active",
    "status",
  ]);
  const colId = getOptionalColumn_(headerMap, [
    "מזהה עובד",
    "ID עובד",
    "Employee ID",
    "ID",
  ]);
  const colName = getOptionalColumn_(headerMap, [
    "שם מלא",
    "name",
    "שם",
    "שמות עובדים",
  ]);
  const colPhone = getOptionalColumn_(headerMap, [
    "טלפון",
    "פלאפון",
    "נייד",
    "phone",
    "mobile",
  ]);
  const colBirthDate = getOptionalColumn_(headerMap, [
    "ת. לידה",
    "תאריך לידה",
    "date of birth",
    "dob",
  ]);
  const colGender = getOptionalColumn_(headerMap, ["מין", "gender"]);
  const colShirtSize = getOptionalColumn_(headerMap, [
    "מידת חולצה",
    "חולצה",
    "shirt size",
    "shirt",
    "tshirt size",
  ]);
  const colDepartment = getOptionalColumn_(headerMap, [
    "מחלקה",
    "מחלקות",
    "department",
  ]);
  const colSize = getOptionalColumn_(headerMap, [
    "מידה (WebApp)",
    "מידה",
    "size",
  ]);
  const colSizeSource = getOptionalColumn_(headerMap, [
    "SOURCE_SIZE",
    "source size",
  ]);
  const colJobName = getOptionalColumn_(headerMap, [
    "סוג העבודה",
    "סוגי עבודה",
  ]);
  const colNotes = getOptionalColumn_(headerMap, ["הערות", "notes"]);
  const colSystemRole = getOptionalColumn_(headerMap, [
    "System Role",
    "system role",
    "Role",
    "Admin Role",
  ]);

  function normalizeSystemRole_(raw) {
    const val = stringValue(raw).toUpperCase();
    if (!val) return "EMPLOYEE";
    if (val === "NONE") return "NONE";
    if (val.indexOf("ADMIN") !== -1) return "ADMIN";
    if (val.indexOf("SHIFT") !== -1 || val.indexOf("MANAGER") !== -1)
      return "SHIFT_MANAGER";
    if (val.indexOf("VIEW") !== -1) return "VIEWER";
    if (val === "EMPLOYEE") return "EMPLOYEE";
    return "EMPLOYEE";
  }

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmail = stringValue(row[colEmail - 1]).toLowerCase();
    if (!rowEmail) continue;
    if (rowEmail !== email) continue;

    const statusVal = colStatus ? stringValue(row[colStatus - 1]) : "";
    const statusLower = statusVal.toLowerCase();
    const inactive =
      statusLower.indexOf("לא פעיל") !== -1 ||
      statusLower.indexOf("inactive") !== -1;
    const active = !inactive;

    const empId = colId ? stringValue(row[colId - 1]) : "";
    const empName = colName ? stringValue(row[colName - 1]) : "";
    const employee = {
      id: empId,
      name: empName,
      email: rowEmail,
      active: active,
      status: statusVal,
    };

    if (colPhone) employee.phone = stringValue(row[colPhone - 1]);
    if (colBirthDate) employee.birthDate = toIsoDate_(row[colBirthDate - 1]);
    if (colGender) employee.gender = stringValue(row[colGender - 1]);
    if (colShirtSize) employee.shirtSize = stringValue(row[colShirtSize - 1]);
    if (colDepartment)
      employee.department = stringValue(row[colDepartment - 1]);
    if (colSize) employee.size = stringValue(row[colSize - 1]);
    if (colSizeSource)
      employee.sizeSource = stringValue(row[colSizeSource - 1]);
    if (colJobName) {
      const jobName = stringValue(row[colJobName - 1]);
      employee.jobName = jobName;
      employee.jobTitle = jobName;
    }
    if (colNotes) employee.notes = stringValue(row[colNotes - 1]);
    const systemRoleRaw = colSystemRole ? row[colSystemRole - 1] : "";
    employee.systemRole = normalizeSystemRole_(systemRoleRaw);

    return {
      ok: true,
      success: true,
      found: true,
      exists: true, // legacy compatibility for callers that check "exists"
      employeeId: empId,
      fullName: empName,
      name: empName,
      employee: employee,
    };
  }

  return {
    ok: true,
    success: true,
    found: false,
    exists: false,
    employeeId: "",
    fullName: "",
    name: "",
    employee: null,
  };
}

function reportAccessIssue_(payload) {
  const email = stringValue(payload && payload.email);
  const name = stringValue(payload && payload.name);
  const ua = stringValue(payload && payload.userAgent);
  const ts =
    stringValue(payload && payload.timestamp) || new Date().toISOString();

  if (!ACCESS_ISSUE_RECIPIENTS.length) {
    return {
      success: false,
      error: "ACCESS_ISSUE_RECIPIENTS is not configured",
    };
  }

  if (!email) {
    return { success: false, error: "missing email" };
  }

  const subject = "דיווח הרשאה - Boulder Shifts";
  const body =
    "התקבלה בקשת גישה שאינה מזוהה בגיליון עובדים:\n" +
    "שם: " +
    (name || "לא ידוע") +
    "\n" +
    "אימייל: " +
    email +
    "\n" +
    "זמן: " +
    ts +
    "\n" +
    (ua ? "User-Agent: " + ua + "\n" : "") +
    "מקור: Apps Script reportAccessIssue";

  for (let i = 0; i < ACCESS_ISSUE_RECIPIENTS.length; i++) {
    MailApp.sendEmail({
      to: ACCESS_ISSUE_RECIPIENTS[i],
      subject: subject,
      body: body,
    });
  }

  return { success: true, sentTo: ACCESS_ISSUE_RECIPIENTS };
}

function normalizePayType_(name) {
  const key = stringValue(name).toLowerCase();
  if (!key) return "";
  if (key.indexOf("חודשי") !== -1 || key.indexOf("month") !== -1)
    return "monthly";
  if (key.indexOf("יחיד") !== -1 || key.indexOf("unit") !== -1) return "unit";
  if (key.indexOf("יומי") !== -1 || key.indexOf("day") !== -1) return "daily";
  if (key.indexOf("שעת") !== -1 || key.indexOf("hour") !== -1) return "hourly";
  return "";
}

function normalizeDirectionToHebrew_(value) {
  const key = stringValue(value).toLowerCase();
  if (!key) return "";
  if (
    key === "exit" ||
    key === "out" ||
    key === "end" ||
    key === "יציאה" ||
    key === "יצא"
  )
    return "יציאה";
  if (
    key === "entry" ||
    key === "enter" ||
    key === "in" ||
    key === "start" ||
    key === "כניסה" ||
    key === "נכנס"
  )
    return "כניסה";
  return stringValue(value);
}

function normalizeShiftForWrite_({
  payType,
  timestamp,
  workDate,
  direction,
  fixDate,
  fixTime,
  units,
  reportMode,
  manualDate,
  manualTime,
}) {
  const tsDate = (timestamp || "").split("T")[0] || workDate || "";
  const mode = stringValue(reportMode);

  const normalized = {
    workDate: workDate || tsDate,
    direction: direction,
    fixDate: fixDate,
    fixTime: fixTime,
    units: units,
  };

  if (payType === "hourly") {
    normalized.units = "";
    const isManual =
      mode.toLowerCase() === "manual" || !!manualDate || !!manualTime;
    if (isManual) {
      normalized.fixDate = manualDate || workDate || tsDate;
      normalized.fixTime = manualTime;
      normalized.workDate = normalized.fixDate || tsDate;
    } else {
      normalized.fixDate = "";
      normalized.fixTime = "";
      normalized.workDate = tsDate;
    }
  } else if (payType === "unit") {
    normalized.direction = "";
    normalized.fixDate = "";
    normalized.fixTime = "";
    normalized.workDate = workDate || tsDate;
  } else if (payType === "daily") {
    normalized.direction = "";
    normalized.fixDate = "";
    normalized.fixTime = "";
    normalized.units = "";
    normalized.workDate = workDate || tsDate;
  }

  if (payType === "unit" && normalized.units !== "") {
    const parsedUnits = Number(normalized.units);
    normalized.units = isNaN(parsedUnits) ? normalized.units : parsedUnits;
  }

  return normalized;
}

function resolveEmployeeJobPayType_(sheet, headers, employeeId, jobTypeId) {
  if (!jobTypeId) return null;
  const empIdCol = getRequiredColumn_(
    headers,
    ["מזהה עובד", "ID עובד"],
    sheet.getName(),
  );
  const jobTypeCol = getRequiredColumn_(
    headers,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheet.getName(),
  );
  const payIdCol = getOptionalColumn_(headers, [
    "ID אופן תשלום",
    "ID אופני תשלום",
  ]);
  const payNameCol = getOptionalColumn_(headers, ["אופן תשלום", "אופני תשלום"]);
  const statusCol = getOptionalColumn_(headers, ["סטטוס"]);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (stringValue(row[empIdCol - 1]) !== employeeId) continue;
    if (statusCol && stringValue(row[statusCol - 1]) === "לא פעיל") continue;
    if (stringValue(row[jobTypeCol - 1]) !== jobTypeId) continue;
    return {
      payTypeId: payIdCol ? stringValue(row[payIdCol - 1]) : "",
      payTypeName: payNameCol ? stringValue(row[payNameCol - 1]) : "",
    };
  }
  return null;
}

function sendMonthlyBlockEmail_(details) {
  const primaryRecipient = "ZIL.SHAKED@GMAIL.COM";
  const configuredRecipients = MONTHLY_ALERT_RECIPIENTS.length
    ? MONTHLY_ALERT_RECIPIENTS
    : ACCESS_ISSUE_RECIPIENTS;

  const dedup = {};
  const recipients = [];
  [primaryRecipient]
    .concat(configuredRecipients || [])
    .forEach(function (addr) {
      const clean = stringValue(addr);
      if (!clean) return;
      const key = clean.toLowerCase();
      if (dedup[key]) return;
      dedup[key] = true;
      recipients.push(clean);
    });

  if (!recipients.length) return;

  const jobLabel = details.jobName
    ? details.jobTypeId
      ? details.jobName + " (" + details.jobTypeId + ")"
      : details.jobName
    : details.jobTypeId;
  const reporterName = details.employeeName || "העובד";
  const contactEmail = details.employeeEmail || "";
  const contactPhone = details.employeePhone || details.phone || "";
  const subject = "דיווח חדש – חסימת סוג עבודה חודשי";

  const textLines = [
    subject,
    jobLabel ? "סוג עבודה: " + jobLabel : "",
    "הווב-אפליקציה חסמה דיווח כי סוג העבודה מוגדר לתשלום חודשי עבור העובד.",
    reporterName +
      " סימן שמדובר בטעות ומבקש לבדוק אם יש לאפשר דיווחים או לעדכן את אופן התשלום.",
    "",
    "מה צריך לבדוק?",
    "1) האם סוג העבודה צריך להיות חודשי עבור העובד הזה?",
    "2) אם לא, עדכנו את אופן התשלום או את שיוך העובד לסוג העבודה.",
    "",
    "ליצירת קשר עם העובד:",
    contactEmail ? "מייל: " + contactEmail : "",
    contactPhone ? "טלפון: " + contactPhone : "",
    details.timestamp ? "נשלח ב: " + details.timestamp : "",
  ].filter(function (line) {
    return stringValue(line);
  });

  var contactList = [];
  if (contactEmail) contactList.push("<li>מייל: " + contactEmail + "</li>");
  if (contactPhone) contactList.push("<li>טלפון: " + contactPhone + "</li>");
  const contactHtml =
    contactList.length > 0
      ? "<p>ליצירת קשר עם העובד:</p><ul>" + contactList.join("") + "</ul>"
      : "";

  const htmlBody =
    "<p><strong>" +
    subject +
    "</strong></p>" +
    (jobLabel ? "<p>סוג עבודה: " + jobLabel + "</p>" : "") +
    "<p>הווב-אפליקציה חסמה דיווח כי סוג העבודה מוגדר לתשלום חודשי עבור העובד.</p>" +
    "<p>" +
    reporterName +
    " סימן שמדובר בטעות ומבקש לבדוק אם יש לאפשר דיווחים או לעדכן את אופן התשלום.</p>" +
    "<p><strong>מה צריך לבדוק?</strong><br>" +
    "1) האם סוג העבודה צריך להיות חודשי עבור העובד הזה?<br>" +
    "2) אם לא, עדכנו את אופן התשלום או את שיוך העובד לסוג העבודה.</p>" +
    contactHtml +
    (details.timestamp ? "<p>נשלח ב: " + details.timestamp + "</p>" : "");

  const body = textLines.join("\n");

  recipients.forEach(function (to) {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: body,
      htmlBody: htmlBody,
    });
  });
}

function writeWorkLogFromNormalizedShift_(normalized, meta, logger) {
  var fnStartMs = new Date().getTime();
  const logsSheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  var headerStartMs = new Date().getTime();
  const headers = getHeaderMap_(logsSheet);
  const logsSheetName = logsSheet.getName();
  getRequiredColumn_(headers, ["ID דיווח", "Shift ID"], logsSheetName);
  getRequiredColumn_(headers, ["חותמת זמן"], logsSheetName);
  getRequiredColumn_(headers, ["ID עובד", "מזהה עובד"], logsSheetName);
  getRequiredColumn_(headers, ["שם מלא"], logsSheetName);
  getRequiredColumn_(headers, ["כניסה / יציאה", "דיווח שעות"], logsSheetName);
  getRequiredColumn_(headers, ["תיקון תאריך"], logsSheetName);
  getRequiredColumn_(headers, ["תיקון שעה"], logsSheetName);
  getRequiredColumn_(headers, ["ID סוג עבודה", "ID סוגי עבודה"], logsSheetName);
  getRequiredColumn_(headers, ["סוג עבודה", "סוגי עבודה"], logsSheetName);
  getRequiredColumn_(headers, ["מחלקה", "מחלקות"], logsSheetName);
  getRequiredColumn_(headers, ["כמות היחידות", "דיווח יחידות"], logsSheetName);
  getRequiredColumn_(headers, ["הערות", "הערה"], logsSheetName);
  if (logger) {
    logger.info("shift.worklog.header-map", {
      sheet: logsSheetName,
      headerKeys: Object.keys(headers),
    });
  }
  logDuration_(logger, "shift.worklog.headers", headerStartMs, {
    sheet: logsSheetName,
  });

  const row = buildRowFromHeaders_(headers, {
    note: stringValue(meta.note),
    reportId: normalized.shiftId,
    shiftId: normalized.shiftId,
    timestamp: normalized.timestamp,
    employeeId: meta.employeeId,
    employeeName: meta.employeeName,
    direction: normalized.direction,
    fixDate: normalized.fixDate,
    fixTime: normalized.fixTime,
    jobTypeId: normalized.jobTypeId,
    jobName: normalized.jobName,
    department: normalized.department,
    units: normalized.units,
    workDate: normalized.workDate,
  });

  if (logger) {
    logger.info("shift.worklog.row-preview", {
      hasShiftId: !!(normalized && normalized.shiftId),
      hasEmployeeId: !!(meta && meta.employeeId),
      workDate: normalized && normalized.workDate,
    });
  }

  var appendStartMs = new Date().getTime();
  logsSheet.appendRow(row);
  logDuration_(logger, "shift.worklog.append", appendStartMs, {
    sheet: logsSheetName,
  });

  try {
    var upsertStartMs = new Date().getTime();
    var workDateForUpsert = buildWorkDateForUpsert_(normalized);

    if (logger) {
      logger.info("shift.worklog.upsert.check-fn", {
        fnType: typeof SHIFTS_upsertAroundWorkLog_,
      });
    }

    if (workDateForUpsert) {
      if (logger) {
        logger.info("shift.worklog.upsert.start", {
          employeeId: String(meta.employeeId || normalized.employeeId || ""),
          jobTypeId: String(normalized.jobTypeId || ""),
          workDate: workDateForUpsert,
        });
      }

      SHIFTS_upsertAroundWorkLog_(
        String(meta.employeeId || normalized.employeeId || ""),
        String(normalized.jobTypeId || ""),
        workDateForUpsert,
      );
    }

    logDuration_(logger, "shift.worklog.upsert-shifts", upsertStartMs, {
      hasWorkDate: !!workDateForUpsert,
      workDate: workDateForUpsert,
    });
  } catch (e) {
    if (logger && logger.error) {
      logger.error(
        "shift.worklog.upsert.error",
        {
          employeeId: String(meta.employeeId || normalized.employeeId || ""),
          jobTypeId: String(normalized.jobTypeId || ""),
          workDate: workDateForUpsert,
          fnType: typeof SHIFTS_upsertAroundWorkLog_,
          errorMessage: String((e && e.message) || e),
        },
        "SHIFTS_UPSERT_FAILED",
        e,
      );
    } else {
      Logger.log("SHIFTS_upsertAroundWorkLog_ error: " + e);
    }
  }
  logDuration_(logger, "shift.worklog.total", fnStartMs, {
    sheet: logsSheetName,
  });
  return {
    ok: true,
    success: true,
    requiresApproval: false,
    shiftId: normalized.shiftId,
    status: "logged",
  };
}

function findDuplicateWorkLogs_(criteria, includeShiftId) {
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headers = getHeaderMap_(sheet);
  const reportIdCol = getRequiredColumn_(
    headers,
    ["ID דיווח", "Shift ID"],
    sheet.getName(),
  );
  const empCol = getRequiredColumn_(
    headers,
    ["ID עובד", "מזהה עובד"],
    sheet.getName(),
  );
  const empNameCol = getOptionalColumn_(headers, ["שם מלא", "שם", "עובד"]);
  const jobTypeIdCol = getOptionalColumn_(headers, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const jobNameCol = getOptionalColumn_(headers, ["סוג עבודה", "סוגי עבודה"]);
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const directionCol = getOptionalColumn_(headers, [
    "כניסה / יציאה",
    "דיווח שעות",
  ]);
  const unitsCol = getOptionalColumn_(headers, [
    "כמות היחידות",
    "כמות יחידות",
    "דיווח יחידות",
  ]);
  const workDateCol = getOptionalColumn_(headers, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const fixDateCol = getOptionalColumn_(headers, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(headers, ["תיקון שעה"]);
  const tsCol = getOptionalColumn_(headers, ["חותמת זמן"]);
  const noteCol = getOptionalColumn_(headers, [
    "הערות",
    "הערה",
    "הערות למשמרת",
  ]);

  function buildCanonical_(row) {
    const workDateRaw = workDateCol ? row[workDateCol - 1] : row[tsCol - 1];
    const workDateIso = toIsoDate_(workDateRaw) || "";
    const timestampRaw = tsCol ? row[tsCol - 1] : "";
    const fixDateRaw = fixDateCol ? row[fixDateCol - 1] : "";
    const fixTimeRaw = fixTimeCol ? row[fixTimeCol - 1] : "";
    return {
      shiftId: reportIdCol ? stringValue(row[reportIdCol - 1]) : "",
      employeeId: empCol ? stringValue(row[empCol - 1]) : "",
      jobTypeId: jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "",
      jobName: jobNameCol ? stringValue(row[jobNameCol - 1]) : "",
      department: deptCol ? stringValue(row[deptCol - 1]) : "",
      direction: directionCol ? stringValue(row[directionCol - 1]) : "",
      workDate: workDateIso,
      fixDate: fixDateCol ? toIsoDate_(fixDateRaw) : "",
      fixTime: fixTimeCol ? stringValue(fixTimeRaw) : "",
      units:
        unitsCol &&
        row[unitsCol - 1] !== undefined &&
        row[unitsCol - 1] !== null
          ? stringValue(row[unitsCol - 1])
          : "",
      note: noteCol ? stringValue(row[noteCol - 1]) : "",
      timestamp: tsCol ? stringValue(timestampRaw) : "",
      eventMs: extractEventMs_(
        workDateRaw,
        timestampRaw,
        fixDateRaw,
        fixTimeRaw,
      ),
    };
  }

  function canonicalKey_(rec) {
    return [
      rec.employeeId,
      rec.jobTypeId,
      rec.department,
      rec.direction,
      rec.workDate,
      rec.fixDate,
      rec.fixTime,
      rec.units,
      rec.note,
    ].join("__");
  }

  const values = sheet.getDataRange().getValues();
  const targetEmployee = stringValue(criteria.employeeId);
  const targetJobType = stringValue(criteria.jobTypeId);
  const targetDirection = stringValue(criteria.direction);
  const targetWorkDate = toIsoDate_(criteria.workDate) || "";
  const hasTimeInCriteria = Boolean(
    stringValue(criteria.fixTime) ||
    (criteria.timestamp && String(criteria.timestamp).match(/\d{1,2}:\d{2}/)) ||
    (criteria.workDate && String(criteria.workDate).match(/\d{1,2}:\d{2}/)),
  );
  const targetEventMs =
    typeof criteria.eventMs === "number"
      ? criteria.eventMs
      : hasTimeInCriteria
        ? extractEventMs_(
            criteria.workDate,
            criteria.timestamp,
            criteria.fixDate,
            criteria.fixTime,
          )
        : null;

  if (
    targetEventMs === null ||
    targetEventMs === undefined ||
    isNaN(targetEventMs)
  ) {
    return [];
  }

  const results = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rec = buildCanonical_(row);
    if (rec.eventMs === null || rec.eventMs === undefined || isNaN(rec.eventMs))
      continue;
    if (rec.employeeId !== targetEmployee) continue;
    if (rec.jobTypeId !== targetJobType) continue;
    if (rec.direction !== targetDirection) continue;
    if (rec.workDate !== targetWorkDate) continue;
    if (!isWithinDuplicateWindow_(rec.eventMs, targetEventMs)) continue;

    // includeShiftId is used to ensure the just-created row is not skipped
    if (includeShiftId && rec.shiftId === includeShiftId) {
      // always include
    }

    results.push({
      shiftId: rec.shiftId,
      employeeId: rec.employeeId,
      jobTypeId: rec.jobTypeId,
      jobName: rec.jobName,
      department: rec.department,
      direction: rec.direction,
      workDate: rec.workDate,
      timestamp: rec.timestamp,
      rowIndex: i + 1,
      canonicalKey: canonicalKey_(rec),
      eventMs: rec.eventMs,
    });
  }

  return results;
}

function cleanupExactDuplicateWorkLogs_(duplicates) {
  if (!duplicates || duplicates.length < 2) return { deletedCount: 0 };
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const foundCount = duplicates.length;
  const groups = {};
  duplicates.forEach(function (d) {
    if (!d || !d.canonicalKey) return;
    if (!groups[d.canonicalKey]) groups[d.canonicalKey] = [];
    groups[d.canonicalKey].push(d);
  });

  const rowsToDelete = [];
  Object.keys(groups).forEach(function (key) {
    const list = groups[key].slice().sort(function (a, b) {
      return (a.rowIndex || 0) - (b.rowIndex || 0);
    });
    if (list.length < 2) return;
    for (let i = 1; i < list.length; i++) {
      if (list[i].rowIndex) rowsToDelete.push(list[i].rowIndex);
    }
  });

  if (!rowsToDelete.length) return { deletedCount: 0 };

  const MAX_DELETE = 100;
  rowsToDelete.sort(function (a, b) {
    return b - a;
  });
  const limited = rowsToDelete.slice(0, MAX_DELETE);
  limited.forEach(function (rowIdx) {
    sheet.deleteRow(rowIdx);
  });

  appendSystemLog_({
    operation: "WORK_LOG_SAVE",
    step: "delete-exact-duplicates",
    severity: "info",
    errorCode: null,
    details: {
      foundCount: foundCount,
      deletedCount: limited.length,
      capped: rowsToDelete.length > MAX_DELETE,
      timeWindowMs: DUPLICATE_TIME_WINDOW_MS,
    },
  });

  return {
    deletedCount: limited.length,
    capped: rowsToDelete.length > MAX_DELETE,
  };
}

const DUPLICATE_TIME_WINDOW_MS = 10 * 60 * 1000; // 10 minutes
const FAULT_ACK_PROPERTY = "FAULT_ACK_MAP";

function normalizeTimeString_(val) {
  const s = stringValue(val);
  if (!s) return "";
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (!m) return "";
  const hh = ("0" + m[1]).slice(-2);
  const mm = ("0" + m[2]).slice(-2);
  const ss = ("0" + (m[3] || "00")).slice(-2);
  return hh + ":" + mm + ":" + ss;
}

function parseDateTimeFlexible_(datePart, timePart) {
  if (!datePart && !timePart) return null;

  if (datePart instanceof Date && !isNaN(datePart.getTime())) {
    if (timePart) {
      const timeNorm = normalizeTimeString_(timePart);
      if (timeNorm) {
        const isoDate = toIsoDate_(datePart);
        const composed = isoDate ? isoDate + "T" + timeNorm : null;
        const composedDate = composed ? new Date(composed) : null;
        if (composedDate && !isNaN(composedDate.getTime())) return composedDate;
      }
    }
    return datePart;
  }

  const dateStr = stringValue(datePart);
  const timeStr = normalizeTimeString_(timePart);

  if (dateStr && timeStr) {
    const isoDate = toIsoDate_(dateStr);
    const composed = isoDate ? isoDate + "T" + timeStr : "";
    if (composed) {
      const dt = new Date(composed);
      if (!isNaN(dt.getTime())) return dt;
    }
  }

  if (dateStr) {
    const mDmyTime = dateStr.match(
      /^(\d{2})[\/.\-](\d{2})[\/.\-](\d{4})[ T](\d{2}):(\d{2})(?::(\d{2}))?/,
    );
    if (mDmyTime) {
      const iso =
        mDmyTime[3] +
        "-" +
        mDmyTime[2] +
        "-" +
        mDmyTime[1] +
        "T" +
        (mDmyTime[4] || "00") +
        ":" +
        (mDmyTime[5] || "00") +
        ":" +
        (mDmyTime[6] || "00");
      const parsed = new Date(iso);
      if (!isNaN(parsed.getTime())) return parsed;
    }

    const parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) return parsed;
  }

  return null;
}

function extractEventMs_(workDateRaw, timestampRaw, fixDate, fixTime) {
  const fixDateTime = parseDateTimeFlexible_(
    fixDate || workDateRaw,
    fixTime || null,
  );
  if (fixDateTime && !isNaN(fixDateTime.getTime()))
    return fixDateTime.getTime();

  const tsParsed = parseDateTimeFlexible_(timestampRaw, null);
  if (tsParsed && !isNaN(tsParsed.getTime())) return tsParsed.getTime();

  const workParsed = parseDateTimeFlexible_(workDateRaw, null);
  if (workParsed && !isNaN(workParsed.getTime())) return workParsed.getTime();

  const isoDate = toIsoDate_(workDateRaw);
  if (isoDate) {
    const midnight = new Date(isoDate + "T00:00:00");
    if (!isNaN(midnight.getTime())) return midnight.getTime();
  }

  return null;
}

function isWithinDuplicateWindow_(msA, msB) {
  if (msA === null || msA === undefined || isNaN(msA)) return false;
  if (msB === null || msB === undefined || isNaN(msB)) return false;
  return Math.abs(msA - msB) <= DUPLICATE_TIME_WINDOW_MS;
}

// Lightweight in-memory check to validate duplicate window behavior without touching sheets.
function __testDuplicateDetectMock_() {
  const base = Date.parse("2024-01-01T10:00:00Z");
  const samples = [
    {
      employeeId: "E1",
      jobTypeId: "J1",
      direction: "IN",
      workDate: "2024-01-01",
      eventMs: base,
    },
    {
      employeeId: "E1",
      jobTypeId: "J1",
      direction: "IN",
      workDate: "2024-01-01",
      eventMs: base + 5 * 60 * 1000,
    },
    {
      employeeId: "E1",
      jobTypeId: "J1",
      direction: "IN",
      workDate: "2024-01-01",
      eventMs: base + 20 * 60 * 1000,
    },
  ];

  function key(rec) {
    const bucket = Math.floor((rec.eventMs || 0) / DUPLICATE_TIME_WINDOW_MS);
    return [
      rec.employeeId,
      rec.jobTypeId,
      rec.direction,
      rec.workDate,
      bucket,
    ].join("__");
  }

  const groups = {};
  samples.forEach(function (rec) {
    const k = key(rec);
    if (!groups[k]) groups[k] = [];
    groups[k].push(rec);
  });

  const withinWindow = Object.values(groups).some(function (list) {
    return Array.isArray(list) && list.length >= 2;
  });

  const outsideWindow = samples.filter(function (rec) {
    return !isWithinDuplicateWindow_(rec.eventMs, base);
  }).length;

  return {
    sampleCount: samples.length,
    buckets: Object.keys(groups).length,
    withinWindowDuplicateDetected: withinWindow,
    outsideWindowCount: outsideWindow,
  };
}

function buildDuplicateKey_(
  employeeId,
  jobTypeId,
  workDate,
  direction,
  eventMs,
) {
  const w = toIsoDate_(workDate) || "";
  const dir = stringValue(direction || "").trim();
  const bucket =
    typeof eventMs === "number" && !isNaN(eventMs)
      ? Math.floor(eventMs / DUPLICATE_TIME_WINDOW_MS)
      : "na";
  return [stringValue(employeeId), stringValue(jobTypeId), w, dir, bucket].join(
    "__",
  );
}

function buildFaultHash_(rec) {
  return [
    stringValue(rec.shiftId),
    stringValue(rec.employeeId),
    stringValue(rec.jobTypeId),
    stringValue(rec.workDate),
    stringValue(rec.direction),
    stringValue(rec.timestamp || rec.fixDate || ""),
  ].join("__");
}

function getFaultAckMap_() {
  try {
    const raw =
      PropertiesService.getDocumentProperties().getProperty(FAULT_ACK_PROPERTY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (e) {
    return {};
  }
}

function setFaultAck_(hash) {
  if (!hash) return;
  const props = PropertiesService.getDocumentProperties();
  const map = getFaultAckMap_();
  map[hash] = true;
  props.setProperty(FAULT_ACK_PROPERTY, JSON.stringify(map));
}

function handleShiftReportDeleteByIds_(payload, logger) {
  var fnStartMs = new Date().getTime();
  const shiftIds = (payload && payload.shiftIds) || [];
  if (!Array.isArray(shiftIds) || !shiftIds.length) {
    return { ok: false, error: "missing_shiftIds" };
  }

  if (shiftIds.length > 20) {
    return { ok: false, error: "DELETE_LIMIT_EXCEEDED", limit: 20 };
  }

  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headers = getHeaderMap_(sheet);
  const reportIdCol = getRequiredColumn_(
    headers,
    ["ID דיווח", "Shift ID"],
    sheet.getName(),
  );
  const values = sheet.getDataRange().getValues();

  const target = {};
  shiftIds.forEach(function (id) {
    target[stringValue(id)] = true;
  });

  const rowsToDelete = [];
  for (let i = 1; i < values.length; i++) {
    const rowShiftId = stringValue(values[i][reportIdCol - 1]);
    if (target[rowShiftId]) {
      rowsToDelete.push(i + 1); // 1-based row index
    }
  }

  rowsToDelete.sort(function (a, b) {
    return b - a;
  });

  rowsToDelete.forEach(function (rowIndex) {
    sheet.deleteRow(rowIndex);
  });

  appendSystemLog_({
    operation: "WORK_LOG_SAVE",
    step: "delete-duplicates",
    severity: "info",
    errorCode: null,
    details: {
      deletedCount: rowsToDelete.length,
      shiftIds: shiftIds,
    },
  });

  logDuration_(logger, "shift.delete-by-ids", fnStartMs, {
    deletedCount: rowsToDelete.length,
  });

  return {
    ok: true,
    status: "deleted",
    deletedCount: rowsToDelete.length,
  };
}

function refreshShiftsForWorkLog_(normalized, meta) {
  const employeeId = stringValue(meta && meta.employeeId);
  const workDateIso = toIsoDate_(
    (normalized && normalized.workDate) ||
      (normalized && normalized.fixDate) ||
      (normalized && normalized.timestamp),
  );
  if (!employeeId || !workDateIso) return { ok: false, skipped: true };

  const jobTypeId = stringValue(normalized && normalized.jobTypeId);

  return rebuildShiftsForRange_({
    dateFrom: workDateIso,
    dateTo: workDateIso,
    employeeId: employeeId,
    jobTypeId: jobTypeId,
  });
}

// Helpers
function getEmployeesSheet_() {
  function buildWorkDateForUpsert_(normalized) {
    if (!normalized) return null;

    if (normalized.workDate instanceof Date) return normalized.workDate;
    if (normalized.fixDate instanceof Date) return normalized.fixDate;
    if (normalized.timestamp instanceof Date) return normalized.timestamp;

    if (normalized.workDate) {
      var d1 = new Date(normalized.workDate);
      if (!isNaN(d1)) return d1;
    }

    if (normalized.fixDate && normalized.fixTime) {
      var d2 = new Date(normalized.fixDate);
      if (!isNaN(d2)) {
        var p = String(normalized.fixTime || "").split(":");
        if (p.length >= 2) {
          var hh = parseInt(p[0], 10);
          var mm = parseInt(p[1], 10);
          if (!isNaN(hh)) d2.setHours(hh);
          if (!isNaN(mm)) d2.setMinutes(mm);
          d2.setSeconds(0);
          d2.setMilliseconds(0);
        }
        if (!isNaN(d2)) return d2;
      }
    }

    if (normalized.timestamp) {
      var d3 = new Date(normalized.timestamp);
      if (!isNaN(d3)) return d3;
    }

    return null;
  }
  try {
    return getSheetByPossibleNames_(EMPLOYEES_SHEET_NAMES);
  } catch (err) {
    const fallback = findSheetByEmailHeader_();
    if (fallback) {
      Logger.log("[getEmployeesSheet] fallback sheet=" + fallback.getName());
      return fallback;
    }
    throw err;
  }
}

function findSheetByEmailHeader_() {
  const ss = getSpreadsheet_();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    try {
      const headerMap = getHeaderMap_(sheet);
      const col = getOptionalColumn_(headerMap, EMAIL_HEADER_CANDIDATES);
      if (col) return sheet;
    } catch (e) {
      // ignore sheets without headers
    }
  }
  return null;
}

function getSheetByPossibleNames_(names) {
  const ss = getSpreadsheet_();
  for (let i = 0; i < names.length; i++) {
    const sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  throw new Error("Sheet not found: " + names.join(", "));
}

function getSheetOrThrow_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  return sheet;
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const active = SpreadsheetApp.getActive();
  if (active) return active;
  throw new Error("Missing SPREADSHEET_ID and no active spreadsheet");
}

function getSpreadsheetForHealth_() {
  if (!SPREADSHEET_ID) {
    throw new Error("Missing SPREADSHEET_ID");
  }
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    var e = new Error(
      "SPREADSHEET_OPEN_FAILED: " +
        (err && err.message ? String(err.message) : String(err)),
    );
    e.code = "SPREADSHEET_OPEN_FAILED";
    throw e;
  }
}

function getSheetByNamesStrict_(ss, names) {
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  throw new Error("Sheet not found: " + names.join(", "));
}

function collectMissing_(headerMap, requiredMatrix, sheetName) {
  var missing = [];
  if (!requiredMatrix || !requiredMatrix.length) return missing;
  for (var i = 0; i < requiredMatrix.length; i++) {
    var colCandidates = requiredMatrix[i];
    try {
      getRequiredColumn_(headerMap, colCandidates, sheetName);
    } catch (err) {
      var label = Array.isArray(colCandidates)
        ? stringValue(colCandidates[0])
        : stringValue(colCandidates);
      if (label) missing.push(label);
    }
  }
  return missing;
}

function flattenColumnLabels_(matrix) {
  return (matrix || []).map(function (col) {
    return Array.isArray(col) ? stringValue(col[0]) : stringValue(col);
  });
}

function getHeaderMap_(sheet, headerRow) {
  const rowIndex = headerRow || 1;
  const headers = sheet
    .getRange(rowIndex, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const map = {};
  headers.forEach((h, idx) => {
    const key = stringValue(h);
    if (key) map[key] = idx + 1; // 1-based
  });
  if (Object.keys(map).length === 0) {
    throw new Error("No headers found in sheet '" + sheet.getName() + "'");
  }
  return map;
}

function normalizeHeaderKey_(header) {
  const key = stringValue(header);
  return HEADER_KEY_MAP.hasOwnProperty(key) ? HEADER_KEY_MAP[key] : key;
}

function getRequiredColumn_(headerMap, possibleHeaders, sheetName) {
  const candidates = Array.isArray(possibleHeaders)
    ? possibleHeaders
    : [possibleHeaders];
  for (let i = 0; i < candidates.length; i++) {
    const candidate = stringValue(candidates[i]);
    const normalized = normalizeHeaderKey_(candidate);
    for (const existing in headerMap) {
      const existingNormalized = normalizeHeaderKey_(existing);
      if (existing === candidate || existingNormalized === normalized) {
        return headerMap[existing];
      }
    }
  }
  var missingHeaderErr = new Error(
    "Missing required header '" +
      candidates[0] +
      "' in sheet '" +
      sheetName +
      "'",
  );
  missingHeaderErr.errorCode = ERROR_CODES.VALIDATION_FAILED;
  missingHeaderErr.meta = {
    sheetName: sheetName,
    missingHeader: candidates[0],
    normalizedHeader: normalizeHeaderKey_(candidates[0]),
  };
  throw missingHeaderErr;
}

function getOptionalColumn_(headerMap, possibleHeaders) {
  const candidates = Array.isArray(possibleHeaders)
    ? possibleHeaders
    : [possibleHeaders];
  for (let i = 0; i < candidates.length; i++) {
    const candidate = stringValue(candidates[i]);
    const normalized = normalizeHeaderKey_(candidate);
    for (const existing in headerMap) {
      const existingNormalized = normalizeHeaderKey_(existing);
      if (existing === candidate || existingNormalized === normalized) {
        return headerMap[existing];
      }
    }
  }
  return null;
}

function buildRowFromHeaders_(headerMap, data) {
  const headers = Object.keys(headerMap).sort(
    (a, b) => headerMap[a] - headerMap[b],
  );
  return headers.map((header) => {
    const normalized = HEADER_KEY_MAP[header] || header;

    // Allow callers that only pass submittedAt to still populate timestamp columns.
    if (normalized === "timestamp") {
      if (data.hasOwnProperty(header)) return data[header];
      if (data.hasOwnProperty("timestamp")) return data.timestamp;
      if (data.hasOwnProperty("submittedAt")) return data.submittedAt;
      return "";
    }

    return data.hasOwnProperty(header)
      ? data[header]
      : data.hasOwnProperty(normalized)
        ? data[normalized]
        : "";
  });
}

function stringValue(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function toIsoDate_(val) {
  if (!val) return "";
  if (
    Object.prototype.toString.call(val) === "[object Date]" &&
    !isNaN(val.getTime())
  ) {
    const y = val.getFullYear();
    const m = ("0" + (val.getMonth() + 1)).slice(-2);
    const d = ("0" + val.getDate()).slice(-2);
    return y + "-" + m + "-" + d;
  }

  const s = stringValue(val);
  if (!s) return "";

  const mIso = s.match(/^(\d{4})[\/.\-](\d{2})[\/.\-](\d{2})$/);
  if (mIso) return mIso[1] + "-" + mIso[2] + "-" + mIso[3];

  const mDmy = s.match(/^(\d{2})[\/.\-](\d{2})[\/.\-](\d{4})$/);
  if (mDmy) return mDmy[3] + "-" + mDmy[2] + "-" + mDmy[1];

  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return toIsoDate_(parsed);
  return s;
}

function linkJobTypeToEmployee_(
  employeeId,
  employeeName,
  jobTypeId,
  jobName,
  department,
  payTypeId,
  payTypeName,
) {
  const sheet = getEmployeesSheet_();
  const headers = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const empIdCol = getRequiredColumn_(
    headers,
    ["מזהה עובד", "ID עובד"],
    sheetName,
  );
  const empNameCol = getRequiredColumn_(headers, ["שם מלא"], sheetName);
  const jobTypeIdCol = getRequiredColumn_(
    headers,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheetName,
  );
  const jobNameCol = getRequiredColumn_(
    headers,
    ["סוג עבודה", "סוגי עבודה"],
    sheetName,
  );
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const payIdCol = getOptionalColumn_(headers, [
    "ID אופן תשלום",
    "ID אופני תשלום",
  ]);
  const payNameCol = getOptionalColumn_(headers, ["אופן תשלום", "אופני תשלום"]);

  const values = sheet.getDataRange().getValues();
  let targetRow = null;
  for (let i = 1; i < values.length; i++) {
    if (stringValue(values[i][empIdCol - 1]) === employeeId) {
      targetRow = i + 1;
      break;
    }
  }

  if (!targetRow) {
    targetRow = sheet.getLastRow() + 1;
    const rowArr = new Array(sheet.getLastColumn()).fill("");
    sheet.getRange(targetRow, 1, 1, rowArr.length).setValues([rowArr]);
  }

  if (employeeName)
    sheet.getRange(targetRow, empNameCol).setValue(employeeName);
  sheet.getRange(targetRow, empIdCol).setValue(employeeId);
  sheet.getRange(targetRow, jobTypeIdCol).setValue(jobTypeId || "");
  sheet.getRange(targetRow, jobNameCol).setValue(jobName || "");
  if (deptCol) sheet.getRange(targetRow, deptCol).setValue(department || "");
  if (payIdCol) sheet.getRange(targetRow, payIdCol).setValue(payTypeId || "");
  if (payNameCol)
    sheet.getRange(targetRow, payNameCol).setValue(payTypeName || "");
}

function ensureEmployeeSizeColumns_(sheet, headers) {
  let sizeCol = getOptionalColumn_(headers, [
    "מידה (WebApp)",
    "מידת חולצה (WebApp)",
  ]);
  let sourceCol = getOptionalColumn_(headers, ["SOURCE_SIZE", "מקור מידה"]);
  if (sizeCol && sourceCol) return { sizeCol, sourceCol, headers };

  const lastCol = sheet.getLastColumn();
  if (!sizeCol) {
    sizeCol = lastCol + 1;
    sheet.getRange(1, sizeCol).setValue("מידה (WebApp)");
  }
  if (!sourceCol) {
    sourceCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, sourceCol).setValue("SOURCE_SIZE");
  }
  const newHeaders = getHeaderMap_(sheet);
  return { sizeCol, sourceCol, headers: newHeaders };
}

function normalizeShiftHeaderValue_(v) {
  return stringValue(v);
}

function validateShiftsHeaders_(headers) {
  const logger = ensureModuleLoggerDefined_
    ? ensureModuleLoggerDefined_("SHIFTS_HEADERS_VALIDATE")
    : null;

  const expected = SHIFTS_HEADERS_CANONICAL.map(normalizeShiftHeaderValue_);
  const actual = (headers || []).map(normalizeShiftHeaderValue_);

  const issues = [];
  if (actual.length < expected.length) {
    issues.push(
      "Shifts sheet has fewer columns (" +
        actual.length +
        ") than expected (" +
        expected.length +
        ")",
    );
  }

  for (let i = 0; i < expected.length; i++) {
    const exp = expected[i];
    const act = actual[i] || "";
    if (exp !== act) {
      issues.push(
        "Mismatch at col " +
          (i + 1) +
          ": expected '" +
          exp +
          "' got '" +
          act +
          "'",
      );
    }
  }

  if (issues.length) {
    // עדיין נרשום לוג לשגיאות כותרת, אבל לא נחסום את האגרגטור.
    logger.error({
      layer: "shifts",
      step: "headers.validate",
      issues: issues.slice(0, 5),
    });

    // ממשיכים לעבוד עם ה-header הנוכחי כפי שהוא, כדי לא לחסום הרכבת משמרות.
    // normalized headers כבר ב-actual.
    return actual;
  }

  return actual;
}

function buildShiftsHeaderMap_(headers) {
  const headerMap = {};
  const notesIdx = [];
  const sourceIdx = [];

  for (let i = 0; i < headers.length; i++) {
    const h = normalizeShiftHeaderValue_(headers[i]);
    if (!h) continue;
    if (headerMap[h] === undefined) headerMap[h] = i; // zero-based index
    if (h === "ID עובד" && headerMap["מזהה עובד"] === undefined)
      headerMap["מזהה עובד"] = i;
    if (h === "תאריך" && headerMap["תאריך משמרת"] === undefined)
      headerMap["תאריך משמרת"] = i;
    if (h === "כניסה" && headerMap["שעת התחלה"] === undefined)
      headerMap["שעת התחלה"] = i;
    if (h === "יציאה" && headerMap["שעת סיום"] === undefined)
      headerMap["שעת סיום"] = i;
    if (h === "הערות" && headerMap["הערות למשמרת"] === undefined)
      headerMap["הערות למשמרת"] = i;
    if (h === "הערות") notesIdx.push(i);
    if (h === "מקור דיווחים") sourceIdx.push(i);
  }

  headerMap.notesLegacy = notesIdx.length ? notesIdx[0] : null;
  headerMap.notesPrimary =
    notesIdx.length > 1 ? notesIdx[1] : headerMap.notesLegacy;
  headerMap.sourceLegacy = sourceIdx.length ? sourceIdx[0] : null;
  headerMap.sourcePrimary =
    sourceIdx.length > 1 ? sourceIdx[1] : headerMap.sourceLegacy;
  headerMap.length = SHIFTS_HEADERS_CANONICAL.length;

  return headerMap;
}

function getShiftsSheetAndHeaderMap_() {
  const ss = getSpreadsheet_();
  const sheet = getSheetOrThrow_("משמרות");

  const maxCols = SHIFTS_HEADERS_CANONICAL.length;
  const headersRange = sheet.getRange(1, 1, 1, maxCols);
  const headers = headersRange.getValues()[0];

  const normalized = validateShiftsHeaders_(headers);
  const headerMap = buildShiftsHeaderMap_(normalized);

  return { sheet: sheet, headerMap: headerMap };
}

function getShiftsSheet_() {
  const ctx = getShiftsSheetAndHeaderMap_();
  return ctx.sheet;
}

function listShifts_(payload) {
  const filters = payload || {};
  const sheet = getSheetOrThrow_("משמרות");
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { shifts: [], total: 0 };

  const headerMap = getHeaderMap_(sheet, 1);

  const shiftIdCol = getRequiredColumn_(headerMap, ["ID משמרת"], "משמרות");
  const employeeCol = getRequiredColumn_(
    headerMap,
    ["ID עובד", "מזהה עובד"],
    "משמרות",
  );
  const employeeNameCol = getOptionalColumn_(headerMap, ["שם מלא"]);
  const workDateCol = getRequiredColumn_(
    headerMap,
    ["תאריך", "תאריך משמרת"],
    "משמרות",
  );
  const startCol = getOptionalColumn_(headerMap, ["כניסה", "שעת התחלה"]);
  const endCol = getOptionalColumn_(headerMap, ["יציאה", "שעת סיום"]);
  const jobTypeIdCol = getOptionalColumn_(headerMap, ["ID סוג עבודה"]);
  const jobNameCol = getOptionalColumn_(headerMap, ["סוג עבודה"]);
  const deptCol = getOptionalColumn_(headerMap, ["מחלקה"]);
  const hoursCol = getOptionalColumn_(headerMap, ["שעות"]);
  const payHoursCol = getOptionalColumn_(headerMap, ["שעות לשכר"]);
  const unitsCol = getOptionalColumn_(headerMap, ["כמות יחידות"]);
  const noteCol = getOptionalColumn_(headerMap, ["הערות", "הערות למשמרת"]);
  const statusCol = getRequiredColumn_(
    headerMap,
    ["סטטוס משמרת", "סטטוס"],
    "משמרות",
  );
  const sourceCol = getOptionalColumn_(headerMap, ["מקור דיווחים"]);

  const jobMap = {};
  try {
    listJobTypes_().forEach(function (jt) {
      jobMap[stringValue(jt.id)] = jt;
    });
  } catch (_err) {
    /* ignore – fall back to sheet-only data */
  }

  let employeeJobPayMap = {};
  const employeeFilterNorm = stringValue(filters.employeeId);
  if (employeeFilterNorm) {
    try {
      const linked = listEmployeeLinkedJobIds_({
        employeeId: employeeFilterNorm,
      });
      const details = (linked && linked.jobTypesDetailed) || [];
      details.forEach(function (d) {
        const id = stringValue(d.jobTypeId || d.id);
        if (!id) return;
        employeeJobPayMap[id] = d;
      });
    } catch (_err) {
      /* ignore – not critical for listing */
    }
  }

  function normalizePayTypeValue_(raw) {
    const key = stringValue(raw).toLowerCase();
    if (!key) return "";
    if (key.indexOf("חודשי") !== -1 || key.indexOf("month") !== -1)
      return "monthly";
    if (key.indexOf("יחיד") !== -1 || key.indexOf("unit") !== -1) return "unit";
    if (key.indexOf("יומי") !== -1 || key.indexOf("day") !== -1) return "daily";
    if (key.indexOf("שעת") !== -1 || key.indexOf("hour") !== -1)
      return "hourly";
    return "";
  }

  function resolvePayMeta_(jobTypeId) {
    const detail = employeeJobPayMap[jobTypeId] || {};
    const job = jobMap[jobTypeId] || {};
    const payTypeRaw =
      detail.payType ||
      detail.payTypeName ||
      detail.payTypeId ||
      job.payType ||
      job.payTypeName ||
      job.payTypeId ||
      "";
    const payType = normalizePayTypeValue_(payTypeRaw);
    const payTypeName = stringValue(detail.payTypeName || job.payTypeName);
    const payTypeId = stringValue(detail.payTypeId || job.payTypeId);
    const payTypeLabel =
      payTypeName ||
      payTypeId ||
      stringValue(job.payTypeLabel || job.payType) ||
      payTypeRaw ||
      "";
    return { payType, payTypeName, payTypeId, payTypeLabel };
  }

  const dateFrom = toIsoDate_(filters.dateFrom || "");
  const dateTo = toIsoDate_(filters.dateTo || "");
  const employeeFilter = stringValue(filters.employeeId);
  const statusFilters = Array.isArray(filters.statuses)
    ? filters.statuses
        .map(function (s) {
          return stringValue(s);
        })
        .filter(function (s) {
          return !!s;
        })
    : [];
  const jobTypeFilters = Array.isArray(filters.jobTypeIds)
    ? filters.jobTypeIds
        .map(function (s) {
          return stringValue(s);
        })
        .filter(function (s) {
          return !!s;
        })
    : [];

  var limit = Number(filters.limit);
  if (!isFinite(limit) || limit <= 0) limit = 200;
  var offset = Number(filters.offset);
  if (!isFinite(offset) || offset < 0) offset = 0;

  function parseNumberOrNull_(value) {
    if (value === null || value === undefined || value === "") return null;
    const n = Number(value);
    return isNaN(n) ? null : n;
  }

  function formatTime_(value) {
    if (
      Object.prototype.toString.call(value) === "[object Date]" &&
      !isNaN(value.getTime())
    ) {
      const hh = ("0" + value.getHours()).slice(-2);
      const mm = ("0" + value.getMinutes()).slice(-2);
      return hh + ":" + mm;
    }
    const s = stringValue(value);
    if (!s) return "";
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return ("0" + m[1]).slice(-2) + ":" + m[2];
    return s;
  }

  const results = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const workDate = toIsoDate_(row[workDateCol - 1]);
    if (dateFrom && workDate && workDate < dateFrom) continue;
    if (dateTo && workDate && workDate > dateTo) continue;

    const employeeId = stringValue(row[employeeCol - 1]);
    if (employeeFilter && employeeId !== employeeFilter) continue;

    const statusRaw = stringValue(row[statusCol - 1]);
    if (statusFilters.length && statusFilters.indexOf(statusRaw) === -1)
      continue;

    const jobTypeId = jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "";
    if (jobTypeFilters.length && jobTypeFilters.indexOf(jobTypeId) === -1)
      continue;

    const hoursDecimal = hoursCol
      ? parseNumberOrNull_(row[hoursCol - 1])
      : null;
    const payHoursVal = payHoursCol
      ? parseNumberOrNull_(row[payHoursCol - 1])
      : null;
    const unitsVal = unitsCol ? parseNumberOrNull_(row[unitsCol - 1]) : null;
    const payMeta = resolvePayMeta_(jobTypeId);

    results.push({
      shiftId: stringValue(row[shiftIdCol - 1]),
      employeeId: employeeId,
      employeeName: employeeNameCol
        ? stringValue(row[employeeNameCol - 1])
        : "",
      workDate: workDate,
      startTime: startCol ? formatTime_(row[startCol - 1]) : "",
      endTime: endCol ? formatTime_(row[endCol - 1]) : "",
      jobTypeId: jobTypeId,
      jobName: jobNameCol ? stringValue(row[jobNameCol - 1]) : "",
      department: deptCol ? stringValue(row[deptCol - 1]) : "",
      payType: payMeta.payType,
      payTypeName: payMeta.payTypeName,
      payTypeId: payMeta.payTypeId,
      payTypeLabel: payMeta.payTypeLabel,
      hoursDecimal: hoursDecimal,
      payHours: payHoursVal,
      units: unitsVal,
      status: statusRaw,
      note: noteCol ? stringValue(row[noteCol - 1]) : "",
      sourceReportIds: sourceCol ? stringValue(row[sourceCol - 1]) : "",
    });
  }

  const total = results.length;
  const sliced = results.slice(offset, offset + limit);
  return { shifts: sliced, total: total };
}

function SHIFTS_getHourlyOverlaps(filter) {
  const src = filter && typeof filter === "object" ? filter : {};

  const dateFrom = toIsoDate_(src.dateFrom || "");
  const dateTo = toIsoDate_(src.dateTo || "");
  const employeeId = stringValue(src.employeeId || "");
  const statuses = Array.isArray(src.statuses)
    ? src.statuses
        .map(function (s) {
          return stringValue(s).toUpperCase();
        })
        .filter(function (s) {
          return !!s;
        })
    : [];
  const jobTypeIds = Array.isArray(src.jobTypeIds)
    ? src.jobTypeIds
        .map(function (s) {
          return stringValue(s);
        })
        .filter(function (s) {
          return !!s;
        })
    : [];

  let bucketSizeMinutes = Number(src.bucketSizeMinutes);
  if (!isFinite(bucketSizeMinutes) || bucketSizeMinutes <= 0)
    bucketSizeMinutes = 60;

  function parseIsoDateLocal_(iso) {
    const s = stringValue(iso);
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    if (!isFinite(y) || !isFinite(mo) || !isFinite(d)) return null;
    return new Date(y, mo - 1, d, 0, 0, 0, 0);
  }

  function parseHm_(value) {
    const s = stringValue(value);
    const m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return null;
    const hh = Number(m[1]);
    const mm = Number(m[2]);
    if (!isFinite(hh) || !isFinite(mm)) return null;
    if (hh < 0 || hh > 23) return null;
    if (mm < 0 || mm > 59) return null;
    return { hh: hh, mm: mm };
  }

  function makeDateTime_(isoDate, hm) {
    const base = parseIsoDateLocal_(isoDate);
    if (!base || !hm) return null;
    base.setHours(hm.hh, hm.mm, 0, 0);
    return base;
  }

  function fmtDateTime_(dt) {
    try {
      return Utilities.formatDate(
        dt,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd'T'HH:mm:ss",
      );
    } catch (_err) {
      return String(dt);
    }
  }

  const shiftsResponse = listShifts_({
    dateFrom: dateFrom,
    dateTo: dateTo,
    employeeId: employeeId || undefined,
    statuses: statuses.length ? statuses : undefined,
    jobTypeIds: jobTypeIds.length ? jobTypeIds : undefined,
    limit: 100000,
    offset: 0,
  });

  const shifts =
    shiftsResponse && Array.isArray(shiftsResponse.shifts)
      ? shiftsResponse.shifts
      : [];

  const bucketMs = bucketSizeMinutes * 60 * 1000;
  const bucketMap = Object.create(null);

  let missingTimes = 0;
  let invalidTimes = 0;
  let crossMidnight = 0;

  function getOrCreateBucket_(bucketStartMs) {
    const key = String(bucketStartMs);
    if (bucketMap[key]) return bucketMap[key];

    const start = new Date(bucketStartMs);
    const end = new Date(bucketStartMs + bucketMs);
    const bucket = {
      bucketStart: fmtDateTime_(start),
      bucketEnd: fmtDateTime_(end),
      count: 0,
      byJob: [],
      __byJobMap: Object.create(null),
    };
    bucketMap[key] = bucket;
    return bucket;
  }

  function bumpJob_(bucket, jobTypeId) {
    const jobId = stringValue(jobTypeId);
    if (!jobId) return;
    const cur = bucket.__byJobMap[jobId] || 0;
    bucket.__byJobMap[jobId] = cur + 1;
  }

  for (let i = 0; i < shifts.length; i++) {
    const sh = shifts[i] || {};
    const isoDate = stringValue(sh.workDate);
    const startHm = parseHm_(sh.startTime);
    const endHm = parseHm_(sh.endTime);
    if (!isoDate || !startHm || !endHm) {
      missingTimes++;
      continue;
    }

    const startDt = makeDateTime_(isoDate, startHm);
    const endDt = makeDateTime_(isoDate, endHm);
    if (!startDt || !endDt) {
      invalidTimes++;
      continue;
    }

    const startMs = startDt.getTime();
    const endMs = endDt.getTime();
    if (!isFinite(startMs) || !isFinite(endMs)) {
      invalidTimes++;
      continue;
    }

    if (endMs <= startMs) {
      // Cross-midnight or invalid; the normalized shifts sheet should ideally avoid this.
      crossMidnight++;
      continue;
    }

    const jobTypeId = stringValue(sh.jobTypeId);

    // Walk buckets that overlap the shift interval.
    let cursor = Math.floor(startMs / bucketMs) * bucketMs;
    const lastBucketStart = Math.floor((endMs - 1) / bucketMs) * bucketMs;
    while (cursor <= lastBucketStart) {
      const bucketStart = cursor;
      const bucketEnd = cursor + bucketMs;
      const overlapMs =
        Math.min(endMs, bucketEnd) - Math.max(startMs, bucketStart);
      if (overlapMs > 0) {
        const bucket = getOrCreateBucket_(bucketStart);
        bucket.count += 1;
        bumpJob_(bucket, jobTypeId);
      }
      cursor += bucketMs;
    }
  }

  const buckets = Object.keys(bucketMap)
    .map(function (k) {
      return bucketMap[k];
    })
    .sort(function (a, b) {
      // bucketStart is formatted consistently; lexical compare works.
      return String(a.bucketStart).localeCompare(String(b.bucketStart));
    })
    .map(function (b) {
      const jobIds = Object.keys(b.__byJobMap);
      b.byJob = jobIds
        .map(function (jobId) {
          return { jobTypeId: jobId, count: b.__byJobMap[jobId] || 0 };
        })
        .sort(function (x, y) {
          return String(x.jobTypeId).localeCompare(String(y.jobTypeId));
        });
      delete b.__byJobMap;
      return b;
    });

  return {
    buckets: buckets,
    totalBuckets: buckets.length,
    totalShifts: shifts.length,
    skipped: {
      missingTimes: missingTimes,
      invalidTimes: invalidTimes,
      crossMidnight: crossMidnight,
    },
    filter: {
      dateFrom: dateFrom || null,
      dateTo: dateTo || null,
      employeeId: employeeId || null,
      statuses: statuses,
      jobTypeIds: jobTypeIds,
      bucketSizeMinutes: bucketSizeMinutes,
    },
  };
}

function rebuildShiftsDaily_() {
  const today = new Date();
  const dateTo = toIsoDate_(today);
  const dateFromDate = new Date(today.getTime());
  dateFromDate.setDate(today.getDate() - 31);
  const dateFrom = toIsoDate_(dateFromDate);
  return rebuildShiftsForRange_({ dateFrom: dateFrom, dateTo: dateTo });
}

function ensureDailyShiftsRebuildTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function (t) {
    return (
      t.getHandlerFunction && t.getHandlerFunction() === "rebuildShiftsDaily_"
    );
  });
  if (exists) return { ok: true, status: "exists" };

  ScriptApp.newTrigger("rebuildShiftsDaily_")
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .create();

  return { ok: true, status: "created" };
}

// Keep shifts up to date when work-log rows change directly in the sheet (limited scope to avoid heavy runs).
// Also delegate to other domain handlers (employees/options/requests) so single onEdit covers all.
function onEdit(e) {
  // Delegate to other modules first (non-blocking)
  try {
    if (typeof EMP_onEdit === "function") EMP_onEdit(e || {});
  } catch (errEmp) {
    Logger.log("EMP_onEdit error: " + errEmp);
  }
  try {
    if (typeof OPT_onEdit === "function") OPT_onEdit(e || {});
  } catch (errOpt) {
    Logger.log("OPT_onEdit error: " + errOpt);
  }
  try {
    if (typeof REQ_onEdit === "function") REQ_onEdit(e || {});
  } catch (errReq) {
    Logger.log("REQ_onEdit error: " + errReq);
  }

  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== WORK_LOGS_SHEET_NAME) return;

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  if (startRow < 2) return; // skip header row
  if (numRows > 5) return; // avoid bulk edits triggering many rebuilds

  var logger = getModuleLogger_("ON_EDIT");
  const headerMap = getHeaderMap_(sheet);
  if (logger) {
    logger.info("worklogs.onEdit", {
      sheet: sheet.getName(),
      startRow: startRow,
      numRows: numRows,
    });
  }
  const empCol = getRequiredColumn_(
    headerMap,
    ["ID עובד", "מזהה עובד"],
    sheet.getName(),
  );
  const tsCol = getRequiredColumn_(headerMap, ["חותמת זמן"], sheet.getName());
  const fixDateCol = getOptionalColumn_(headerMap, ["תיקון תאריך"]);
  const workDateCol = getOptionalColumn_(headerMap, ["תאריך משמרת"]);
  const jobTypeIdCol = getOptionalColumn_(headerMap, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);

  const values = sheet
    .getRange(startRow, 1, numRows, sheet.getLastColumn())
    .getValues();

  const rebuildRequests = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const employeeId = stringValue(row[empCol - 1]);
    const fixDateRaw = fixDateCol ? row[fixDateCol - 1] : "";
    const workDateRaw = workDateCol ? row[workDateCol - 1] : "";
    const tsRaw = row[tsCol - 1];
    const workDateIso = toIsoDate_(fixDateRaw || workDateRaw || tsRaw);
    if (!workDateIso) continue;

    const jobTypeId = jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "";
    const key = workDateIso + "__" + employeeId + "__" + jobTypeId;
    rebuildRequests[key] = {
      date: workDateIso,
      employeeId: employeeId,
      jobTypeId: jobTypeId,
    };
  }

  const keys = Object.keys(rebuildRequests);
  if (logger) {
    logger.info("worklogs.onEdit.rebuild-requests", {
      requestsCount: keys.length,
    });
  }
  for (let i = 0; i < keys.length; i++) {
    const req = rebuildRequests[keys[i]];
    try {
      if (logger) {
        logger.info("worklogs.onEdit.rebuild-call", {
          dateFrom: req.date,
          dateTo: req.date,
          employeeId: req.employeeId,
          jobTypeId: req.jobTypeId,
        });
      }
      rebuildShiftsForRange_({
        dateFrom: req.date,
        dateTo: req.date,
        employeeId: req.employeeId,
        jobTypeId: req.jobTypeId,
      });
    } catch (err) {
      // Non-blocking: rely on rebuildShiftsForRange_ logging and guards.
    }
  }

  try {
    checkAndPromptDuplicateOnEdit_(sheet, headerMap, startRow, numRows);
  } catch (err) {
    Logger.log("duplicate check failed: " + (err && err.message));
  }
}

function checkAndPromptDuplicateOnEdit_(sheet, headerMap, startRow, numRows) {
  const shiftIdCol = getRequiredColumn_(
    headerMap,
    ["ID דיווח", "Shift ID"],
    sheet.getName(),
  );
  const empCol = getRequiredColumn_(
    headerMap,
    ["ID עובד", "מזהה עובד"],
    sheet.getName(),
  );
  const jobTypeIdCol = getOptionalColumn_(headerMap, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const directionCol = getOptionalColumn_(headerMap, [
    "כניסה / יציאה",
    "דיווח שעות",
  ]);
  const workDateCol = getOptionalColumn_(headerMap, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const fixDateCol = getOptionalColumn_(headerMap, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(headerMap, ["תיקון שעה"]);
  const tsCol = getOptionalColumn_(headerMap, ["חותמת זמן"]);

  const values = sheet
    .getRange(startRow, 1, numRows, sheet.getLastColumn())
    .getValues();

  const duplicatesByKey = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const employeeId = stringValue(row[empCol - 1]);
    if (!employeeId) continue;
    const jobTypeId = jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "";
    const direction = directionCol ? stringValue(row[directionCol - 1]) : "";
    const workDateRaw = workDateCol ? row[workDateCol - 1] : row[tsCol - 1];
    const workDate = toIsoDate_(workDateRaw) || "";
    const timestampRaw = tsCol ? row[tsCol - 1] : "";
    const fixDateRaw = fixDateCol ? row[fixDateCol - 1] : "";
    const fixTimeRaw = fixTimeCol ? row[fixTimeCol - 1] : "";
    const eventMs = extractEventMs_(
      workDateRaw,
      timestampRaw,
      fixDateRaw,
      fixTimeRaw,
    );
    if (!workDate) continue;

    const shiftId = stringValue(row[shiftIdCol - 1]);
    const key = buildDuplicateKey_(
      employeeId,
      jobTypeId,
      workDate,
      direction,
      eventMs,
    );
    const dupList = findDuplicateWorkLogs_(
      {
        employeeId: employeeId,
        jobTypeId: jobTypeId,
        workDate: workDate,
        direction: direction,
        timestamp: timestampRaw,
        fixDate: fixDateRaw ? toIsoDate_(fixDateRaw) : "",
        fixTime: stringValue(fixTimeRaw),
      },
      shiftId,
    );
    if (dupList && dupList.length >= 2) {
      duplicatesByKey[key] = dupList;
    }
  }

  const aggregated = [];
  const seenIds = {};
  Object.keys(duplicatesByKey).forEach(function (k) {
    duplicatesByKey[k].forEach(function (d) {
      if (seenIds[d.shiftId]) return;
      seenIds[d.shiftId] = true;
      aggregated.push(d);
    });
  });

  if (!aggregated.length) return;

  showDuplicateDialog_(aggregated.slice(0, 20));
}

function showDuplicateDialog_(duplicates) {
  try {
    const html = HtmlService.createHtmlOutput(
      buildDuplicateDialogHtml_(duplicates),
    )
      .setWidth(520)
      .setHeight(520);
    SpreadsheetApp.getUi().showModalDialog(html, "דיווחים כפולים");
  } catch (err) {
    Logger.log("showDuplicateDialog_ failed: " + (err && err.message));
  }
}

function buildDuplicateDialogHtml_(duplicates) {
  const safePayload = JSON.stringify(duplicates || []).replace(/</g, "\\u003c");
  return `
<html dir="rtl">
<head>
  <style>
    body{font-family:Arial,sans-serif;background:#f7f7f7;margin:0;padding:16px;}
    .card{background:#fff;border-radius:12px;box-shadow:0 8px 24px rgba(0,0,0,0.15);padding:16px;}
    .row{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid #eee;}
    .btn{border:none;border-radius:8px;padding:10px 14px;font-weight:700;cursor:pointer;}
    .danger{background:#d32f2f;color:#fff;}
    .ghost{border:1px solid #ccc;background:#fff;color:#333;}
    .pill{width:32px;height:32px;border-radius:50%;border:2px solid #ccc;background:#fff;font-weight:700;cursor:pointer;}
    .pill.on{background:#fdecea;border-color:#b00020;color:#b00020;}
    .line-through{text-decoration:line-through;color:#888;}
    #confirm{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.35);justify-content:center;align-items:center;}
    #confirm .box{background:#fff;border-radius:12px;padding:16px;box-shadow:0 10px 32px rgba(0,0,0,0.2);width:340px;}
  </style>
</head>
<body>
  <div class="card">
    <h3 style="margin:0 0 8px;">בטוח שאין פה טעות?</h3>
    <p style="margin:0 0 12px;color:#555;">נמצאו דיווחים כפולים. סמנו ב-X למחיקה.</p>
    <div id="list"></div>
    <div style="display:flex;justify-content:space-between;gap:8px;margin-top:14px;">
      <button class="btn ghost" id="noError">אין טעות</button>
      <button class="btn danger" id="done" style="display:none;">סיימתי</button>
    </div>
  </div>
  <div id="confirm">
    <div class="box">
      <h4 style="margin:0 0 10px;">אתה בטוח שאתה רוצה למחוק <span id="count"></span> דיווחים?</h4>
      <div id="status" style="margin:0 0 10px;font-weight:600;color:#1b5e20;"></div>
      <div style="display:flex;gap:8px;">
        <button class="btn ghost" id="back">חזרה</button>
        <button class="btn danger" id="yes">כן</button>
      </div>
    </div>
  </div>
  <script>
    const duplicates = ${safePayload};
    const list = document.getElementById("list");
    const doneBtn = document.getElementById("done");
    const noErr = document.getElementById("noError");
    const confirm = document.getElementById("confirm");
    const count = document.getElementById("count");
    const status = document.getElementById("status");
    const state = new Map();
    duplicates.forEach(d => state.set(d.shiftId, false));

    function render() {
      list.innerHTML = '';
      duplicates.forEach(d => {
        const selected = state.get(d.shiftId);
        const row = document.createElement('div');
        row.className = 'row';
        const btn = document.createElement('button');
        btn.className = 'pill' + (selected ? ' on' : '');
        btn.textContent = selected ? '↺' : 'X';
        btn.onclick = () => {
          state.set(d.shiftId, !selected);
          render();
        };
        const info = document.createElement('div');
        info.style.flex = '1';
        const title = document.createElement('div');
        title.textContent = d.jobName || 'ללא שם עבודה';
        title.className = selected ? 'line-through' : '';
        const meta = document.createElement('div');
        meta.style.fontSize = '13px';
        meta.style.color = '#666';
        meta.className = selected ? 'line-through' : '';
        meta.textContent = (d.workDate || 'תאריך חסר') + (d.direction ? ' · ' + d.direction : '') + (d.department ? ' · ' + d.department : '');
        info.appendChild(title);
        info.appendChild(meta);
        const undo = document.createElement('span');
        undo.style.fontSize = '12px';
        undo.style.color = '#b00020';
        undo.textContent = selected ? 'בטל את מחיקת הדיווח' : '';
        row.appendChild(btn);
        row.appendChild(info);
        row.appendChild(undo);
        list.appendChild(row);
      });
      const selectedCount = Array.from(state.values()).filter(Boolean).length;
      doneBtn.style.display = selectedCount ? 'inline-block' : 'none';
    }

    render();

    noErr.onclick = () => google.script.host.close();

    doneBtn.onclick = () => {
      const selected = Array.from(state.entries()).filter(([,v]) => v).map(([k]) => k);
      if (!selected.length) return;
      count.textContent = selected.length;
      confirm.style.display = 'flex';
      status.textContent = '';
    };

    document.getElementById('back').onclick = () => {
      confirm.style.display = 'none';
    };

    document.getElementById('yes').onclick = () => {
      const selected = Array.from(state.entries()).filter(([,v]) => v).map(([k]) => k);
      if (!selected.length) return;
      google.script.run.withSuccessHandler((res) => {
        if (res && res.ok) {
          status.textContent = 'השינויים נשמרו.';
          setTimeout(() => google.script.host.close(), 1200);
        } else {
          status.style.color = '#b00020';
          status.textContent = (res && res.error) || 'מחיקה נכשלה.';
        }
      }).withFailureHandler((err) => {
        status.style.color = '#b00020';
        status.textContent = err && err.message ? err.message : 'שגיאה במחיקה.';
      }).shiftReport_deleteDuplicatesFromUi(selected);
    };
  </script>
</body>
</html>`;
}

function shiftReport_deleteDuplicatesFromUi(shiftIds) {
  return handleShiftReportDeleteByIds_({ shiftIds: shiftIds });
}

function shiftReport_showFaultsDialog() {
  const ui = SpreadsheetApp.getUi();

  let faults = [];
  try {
    faults = listFaultyWorkLogsForMenu_(30);
  } catch (err) {
    ui.alert("שגיאה בסריקת תקלות: " + (err && err.message ? err.message : err));
    return;
  }

  if (!faults || !faults.length) {
    ui.alert("אין תקלות להצגה.");
    return;
  }

  const html = HtmlService.createHtmlOutput(buildFaultsDialogHtml_(faults))
    .setWidth(720)
    .setHeight(760);

  ui.showModalDialog(html, "תיקון תקלות משמרות");
}

function shiftReport_handleFaultAction(fault, action) {
  if (!fault || !fault.hash || !action) {
    return { ok: false, error: "missing_fault_or_action" };
  }

  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headers = getHeaderMap_(sheet);
  const shiftIdCol = getRequiredColumn_(
    headers,
    ["ID דיווח", "Shift ID"],
    sheet.getName(),
  );
  const empCol = getRequiredColumn_(
    headers,
    ["ID עובד", "מזהה עובד"],
    sheet.getName(),
  );
  const empNameCol = getOptionalColumn_(headers, ["שם מלא", "שם", "עובד"]);
  const jobTypeIdCol = getOptionalColumn_(headers, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const jobNameCol = getOptionalColumn_(headers, ["סוג עבודה", "סוגי עבודה"]);
  const directionCol = getOptionalColumn_(headers, [
    "כניסה / יציאה",
    "דיווח שעות",
  ]);
  const unitsCol = getOptionalColumn_(headers, [
    "כמות היחידות",
    "כמות יחידות",
    "דיווח יחידות",
  ]);
  const workDateCol = getOptionalColumn_(headers, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const tsCol = getOptionalColumn_(headers, ["חותמת זמן"]);

  const values = sheet.getDataRange().getValues();
  let targetRowIndex = null;

  function rowMatchesHash(row, idx) {
    const workDateRaw = workDateCol ? row[workDateCol - 1] : row[tsCol - 1];
    const rec = {
      rowIndex: idx + 1,
      shiftId: stringValue(row[shiftIdCol - 1]),
      employeeId: stringValue(row[empCol - 1]),
      jobTypeId: jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "",
      direction: directionCol ? stringValue(row[directionCol - 1]) : "",
      workDate: toIsoDate_(workDateRaw) || "",
      timestamp: tsCol ? stringValue(row[tsCol - 1]) : "",
    };
    const hash = buildFaultHash_(rec);
    return hash === fault.hash;
  }

  for (let i = 1; i < values.length; i++) {
    if (rowMatchesHash(values[i], i)) {
      targetRowIndex = i + 1;
      break;
    }
  }

  if (!targetRowIndex) {
    return { ok: false, error: "row_not_found" };
  }

  if (action === "assignShiftId") {
    const newId = Utilities.getUuid();
    sheet.getRange(targetRowIndex, shiftIdCol).setValue(newId);
    setFaultAck_(fault.hash);
    appendSystemLog_({
      operation: "WORK_LOG_FAULT",
      step: "assign-shiftId",
      severity: "info",
      errorCode: null,
      details: { rowIndex: targetRowIndex, newId: newId },
    });
    return { ok: true, message: "נוצר ID חדש" };
  }

  if (action === "saveEdits") {
    const edits = (fault && fault.edits) || {};
    const mutations = [];
    function setIf(col, key, normalizer) {
      if (!col || !edits.hasOwnProperty(key)) return;
      const val = normalizer ? normalizer(edits[key]) : edits[key];
      mutations.push({ col: col, val: val });
    }

    setIf(empCol, "employeeId", stringValue);
    setIf(jobTypeIdCol, "jobTypeId", stringValue);
    setIf(jobNameCol, "jobName", stringValue);
    setIf(directionCol, "direction", stringValue);
    setIf(workDateCol, "workDate", toIsoDate_);
    setIf(tsCol, "timestamp", stringValue);
    setIf(unitsCol, "units", stringValue);

    // Optional payType, note columns if exist
    const payTypeCol = getOptionalColumn_(headers, [
      "אופן תשלום",
      "payType",
      "סוג שכר",
    ]);
    const noteCol = getOptionalColumn_(headers, [
      "הערות",
      "הערה",
      "הערה למנהל",
      "הערות למשמרת",
    ]);
    setIf(payTypeCol, "payType", stringValue);
    setIf(noteCol, "note", stringValue);

    mutations.forEach(function (m) {
      if (m.col) {
        sheet.getRange(targetRowIndex, m.col).setValue(m.val || "");
      }
    });

    setFaultAck_(fault.hash);
    appendSystemLog_({
      operation: "WORK_LOG_FAULT",
      step: "save-edits",
      severity: "info",
      errorCode: null,
      details: {
        rowIndex: targetRowIndex,
        editedKeys: Object.keys(edits || {}),
      },
    });
    return { ok: true, message: "השינויים נשמרו" };
  }

  if (action === "ack") {
    setFaultAck_(fault.hash);
    return { ok: true, message: "סומן כתקין" };
  }

  if (action === "deleteRow") {
    sheet.deleteRow(targetRowIndex);
    appendSystemLog_({
      operation: "WORK_LOG_FAULT",
      step: "delete-row",
      severity: "warning",
      errorCode: null,
      details: { rowIndex: targetRowIndex },
    });
    return { ok: true, message: "השורה נמחקה" };
  }

  return { ok: false, error: "unknown_action" };
}

// --- Duplicate scan from menu ---

function SHIFTS_showDuplicateCleanupDialog() {
  const ui = SpreadsheetApp.getUi();

  let groups = [];
  try {
    groups = listDuplicateWorkLogGroupsForMenu_(30);
  } catch (err) {
    ui.alert(
      "שגיאה בסריקת כפילויות: " + (err && err.message ? err.message : err),
    );
    return;
  }

  if (!groups || !groups.length) {
    ui.alert("לא נמצאו כפילויות בין דיווחי העובדים.");
    return;
  }

  const html = HtmlService.createHtmlOutput(
    buildDuplicateCleanupDialogHtml_(groups),
  )
    .setWidth(680)
    .setHeight(760);

  ui.showModalDialog(html, "בדיקת כפילויות משמרות");
}

function listDuplicateWorkLogGroupsForMenu_(maxGroups) {
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headers = getHeaderMap_(sheet);

  const shiftIdCol = getRequiredColumn_(
    headers,
    ["ID דיווח", "Shift ID"],
    sheet.getName(),
  );
  const empCol = getRequiredColumn_(
    headers,
    ["ID עובד", "מזהה עובד"],
    sheet.getName(),
  );
  const empNameCol = getOptionalColumn_(headers, ["שם מלא", "שם", "עובד"]);
  const jobTypeIdCol = getOptionalColumn_(headers, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const jobNameCol = getOptionalColumn_(headers, ["סוג עבודה", "סוגי עבודה"]);
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const directionCol = getOptionalColumn_(headers, [
    "כניסה / יציאה",
    "דיווח שעות",
  ]);
  const unitsCol = getOptionalColumn_(headers, [
    "כמות היחידות",
    "כמות יחידות",
    "דיווח יחידות",
  ]);
  const workDateCol = getOptionalColumn_(headers, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const fixDateCol = getOptionalColumn_(headers, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(headers, ["תיקון שעה"]);
  const tsCol = getOptionalColumn_(headers, ["חותמת זמן"]);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  function toRec(row, rowIndex) {
    const workDateRaw = workDateCol ? row[workDateCol - 1] : row[tsCol - 1];
    const timestampRaw = tsCol ? row[tsCol - 1] : "";
    const fixDateRaw = fixDateCol ? row[fixDateCol - 1] : "";
    const fixTimeRaw = fixTimeCol ? row[fixTimeCol - 1] : "";
    return {
      shiftId: stringValue(row[shiftIdCol - 1]),
      employeeId: stringValue(row[empCol - 1]),
      employeeName: empNameCol ? stringValue(row[empNameCol - 1]) : "",
      jobTypeId: jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "",
      jobName: jobNameCol ? stringValue(row[jobNameCol - 1]) : "",
      department: deptCol ? stringValue(row[deptCol - 1]) : "",
      direction: directionCol ? stringValue(row[directionCol - 1]) : "",
      workDate: toIsoDate_(workDateRaw) || "",
      fixDate: fixDateCol ? toIsoDate_(fixDateRaw) : "",
      fixTime: fixTimeCol ? stringValue(fixTimeRaw) : "",
      timestamp: tsCol ? stringValue(timestampRaw) : "",
      eventMs: extractEventMs_(
        workDateRaw,
        timestampRaw,
        fixDateRaw,
        fixTimeRaw,
      ),
      rowIndex: rowIndex,
    };
  }

  const byKey = {};
  for (let i = 1; i < values.length; i++) {
    const rec = toRec(values[i], i + 1);
    const employeeKey = rec.employeeId || rec.employeeName;
    if (!employeeKey || !rec.workDate) continue;
    const baseKey =
      stringValue(employeeKey) +
      "__" +
      stringValue(rec.jobTypeId) +
      "__" +
      stringValue(rec.direction) +
      "__" +
      rec.workDate;
    if (!baseKey) continue;
    if (!byKey[baseKey]) byKey[baseKey] = [];
    byKey[baseKey].push(rec);
  }

  const groups = [];
  Object.keys(byKey).forEach(function (k) {
    const list = (byKey[k] || []).slice().filter(function (r) {
      return r && r.shiftId;
    });
    if (list.length < 2) return;
    list.sort(function (a, b) {
      if (a.eventMs !== b.eventMs) {
        return (a.eventMs || 0) - (b.eventMs || 0);
      }
      return (a.rowIndex || 0) - (b.rowIndex || 0);
    });

    let cluster = [list[0]];
    for (let i = 1; i < list.length; i++) {
      const current = list[i];
      const prev = cluster[cluster.length - 1];
      if (isWithinDuplicateWindow_(current.eventMs, prev.eventMs)) {
        cluster.push(current);
      } else {
        if (cluster.length > 1) {
          groups.push({
            key: k + "::" + (groups.length + 1),
            keep: cluster[0],
            candidates: cluster.slice(1),
          });
        }
        cluster = [current];
      }
    }

    if (cluster.length > 1) {
      groups.push({
        key: k + "::" + (groups.length + 1),
        keep: cluster[0],
        candidates: cluster.slice(1),
      });
    }
  });

  if (maxGroups && groups.length > maxGroups) {
    return groups.slice(0, maxGroups);
  }
  return groups;
}

function workLogs_auditHeadersForDebug_() {
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
  const headers = headerRow.map(function (h) {
    return stringValue(h);
  });
  const required = [
    "ID דיווח",
    "חותמת זמן",
    "ID עובד",
    "כניסה / יציאה",
    "תיקון תאריך",
    "תיקון שעה",
    "ID סוג עבודה",
    "סוג עבודה",
    "מחלקה",
    "כמות היחידות",
    "הערות",
  ];
  const missing = required.filter(function (name) {
    return headers.indexOf(name) === -1;
  });

  return {
    sheetName: sheet.getName(),
    requiredCount: required.length,
    presentCount: headers.length,
    missingRequired: missing,
    headers: headers,
  };
}

// --- Fault scan from menu ---

function listFaultyWorkLogsForMenu_(maxRows) {
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headers = getHeaderMap_(sheet);
  let employeesSheet = null;
  let empHeaders = null;
  try {
    employeesSheet = getEmployeesSheet_();
    empHeaders = getHeaderMap_(employeesSheet);
  } catch (e) {
    employeesSheet = null;
    empHeaders = null;
  }

  const shiftIdCol = getRequiredColumn_(
    headers,
    ["ID דיווח", "Shift ID"],
    sheet.getName(),
  );
  const empCol = getRequiredColumn_(
    headers,
    ["ID עובד", "מזהה עובד"],
    sheet.getName(),
  );
  const empNameCol = getOptionalColumn_(headers, ["שם מלא", "שם", "עובד"]);
  const jobTypeIdCol = getOptionalColumn_(headers, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const jobNameCol = getOptionalColumn_(headers, ["סוג עבודה", "סוגי עבודה"]);
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const directionCol = getOptionalColumn_(headers, [
    "כניסה / יציאה",
    "דיווח שעות",
  ]);
  const unitsCol = getOptionalColumn_(headers, [
    "כמות היחידות",
    "כמות יחידות",
    "דיווח יחידות",
  ]);
  const workDateCol = getOptionalColumn_(headers, [
    "תאריך משמרת",
    "תיקון תאריך",
    "חותמת זמן",
  ]);
  const tsCol = getOptionalColumn_(headers, ["חותמת זמן"]);

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const ack = getFaultAckMap_();
  const faults = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const workDateRaw = workDateCol ? row[workDateCol - 1] : row[tsCol - 1];
    const rec = {
      rowIndex: i + 1,
      shiftId: stringValue(row[shiftIdCol - 1]),
      employeeId: stringValue(row[empCol - 1]),
      employeeName: empNameCol ? stringValue(row[empNameCol - 1]) : "",
      jobTypeId: jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "",
      jobName: jobNameCol ? stringValue(row[jobNameCol - 1]) : "",
      department: deptCol ? stringValue(row[deptCol - 1]) : "",
      direction: directionCol ? stringValue(row[directionCol - 1]) : "",
      workDate: toIsoDate_(workDateRaw) || "",
      timestamp: tsCol ? stringValue(row[tsCol - 1]) : "",
      units: unitsCol ? stringValue(row[unitsCol - 1]) : "",
      payType: "",
    };

    if (employeesSheet && rec.employeeId && rec.jobTypeId) {
      const payMeta = resolveEmployeeJobPayType_(
        employeesSheet,
        empHeaders,
        rec.employeeId,
        rec.jobTypeId,
      );
      if (payMeta) {
        rec.payType = payMeta.payTypeName || payMeta.payTypeId || "";
      }
    }

    const hash = buildFaultHash_(rec);
    if (ack[hash]) continue;

    const issues = [];
    if (!rec.shiftId) issues.push("missing_shiftId");
    if (!rec.employeeId) issues.push("missing_employeeId");
    if (jobTypeIdCol && !rec.jobTypeId) issues.push("missing_jobTypeId");
    if (!rec.workDate) issues.push("missing_workDate");

    if (!issues.length) continue;

    issues.forEach(function (issue) {
      faults.push({
        hash: hash,
        issue: issue,
        rowIndex: rec.rowIndex,
        shiftId: "",
        employeeId: "",
        employeeName: rec.employeeName,
        jobTypeId: rec.jobTypeId,
        jobName: rec.jobName,
        department: rec.department,
        direction: rec.direction,
        workDate: rec.workDate,
        timestamp: rec.timestamp,
        units: rec.units,
        payType: rec.payType,
        note: rec.note,
      });
    });

    if (maxRows && faults.length >= maxRows) break;
  }

  return faults;
}

function buildFaultsDialogHtml_(faults) {
  const safe = JSON.stringify(faults || []).replace(/</g, "\\u003c");
  return `
<html dir="rtl">
<head>
  <style>
    body{font-family:Arial,sans-serif;background:#f5f6fa;margin:0;padding:16px;}
    .card{background:#fff;border-radius:12px;box-shadow:0 8px 24px rgba(0,0,0,0.12);padding:16px;}
    .item{border:1px solid #e0e0e0;border-radius:10px;padding:12px;margin-bottom:10px;}
    .title{font-weight:700;margin:0 0 4px;font-size:15px;color:#111;display:flex;gap:6px;flex-wrap:wrap;align-items:center;}
    .title .dot{color:#9aa0b5;font-weight:400;}
    .meta{font-size:12px;color:#555;line-height:1.6;}
    .field{display:grid;grid-template-columns:120px 1fr;gap:8px 10px;font-size:12px;color:#333;align-items:center;margin-top:8px;}
    .field label{color:#444;font-weight:700;}
    .field input,.field select{width:100%;padding:6px 8px;border:1px solid #ccc;border-radius:6px;font-size:12px;}
    .field input[disabled]{background:#f3f4f6;color:#666;cursor:not-allowed;}
    .pill{padding:4px 8px;border-radius:6px;font-size:12px;background:#ffebee;color:#b00020;display:inline-block;}
    .reason{margin-top:10px;font-size:12px;color:#c62828;font-weight:700;}
    .actions{display:flex;gap:8px;margin-top:10px;flex-wrap:wrap;}
    .btn{border:none;border-radius:8px;padding:8px 12px;font-weight:700;cursor:pointer;font-size:12px;}
    .ghost{border:1px solid #ccc;background:#fff;color:#333;}
    .primary{background:#3949ab;color:#fff;}
    .danger{background:#d32f2f;color:#fff;}
    #status{margin-top:10px;font-weight:700;color:#1b5e20;}
  </style>
</head>
<body>
  <div class="card">
    <h3 style="margin:0 0 6px;">תקלות דיווחי עובדים</h3>
    <p style="margin:0 0 10px;color:#555;font-size:13px;">בחרו פעולה לכל תקלה: תיקון אוטומטי כשאפשר, סימון כתקין (לא נציג שוב), או מחיקה. ניתן לערוך שדות ולתקן לפני סימון.</p>
    <div id="list"></div>
    <div id="status"></div>
    <div style="display:flex;gap:10px;margin-top:12px;">
      <button class="btn ghost" id="closeBtn">סגירה</button>
      <button class="btn primary" id="refreshBtn">רענון</button>
    </div>
  </div>

  <script>
    const faults = ${safe};
    const listEl = document.getElementById('list');
    const statusEl = document.getElementById('status');
    document.getElementById('closeBtn').onclick = () => google.script.host.close();
    document.getElementById('refreshBtn').onclick = () => google.script.host.close();

    const separator = ' \u00b7 ';
    const issueText = {
      missing_shiftId: 'חסר מזהה דיווח',
      missing_employeeId: 'חסר עובד',
      missing_jobTypeId: 'חסר סוג עבודה',
      missing_workDate: 'חסר תאריך דיווח',
    };

    function normalizePayKind(raw) {
      const s = String(raw || '').toLowerCase();
      if (!s) return '';
      if (s.includes('unit') || s.includes('יחיד')) return 'unit';
      if (s.includes('שעת')) return 'hourly';
      if (s.includes('יומ')) return 'daily';
      if (s.includes('חוד')) return 'monthly';
      return '';
    }

    function extractDate(val) {
      if (!val) return '';
      const match = String(val).match(/(\d{4}-\d{2}-\d{2})/);
      return match && match[1] ? match[1] : '';
    }

    function extractTime(val) {
      if (!val) return '';
      const raw = String(val);
      const iso = raw.match(/T(\d{2}:\d{2})/);
      if (iso && iso[1]) return iso[1];
      const hhmm = raw.match(/\b(\d{1,2}:\d{2})\b/);
      if (hhmm && hhmm[1]) {
        const parts = hhmm[1].split(':');
        return parts[0].padStart(2, '0') + ':' + parts[1];
      }
      return '';
    }

    function updateTimestampFromParts(fault, dateVal, timeVal) {
      const datePart = dateVal || extractDate(fault.timestamp);
      const timePart = timeVal || extractTime(fault.timestamp);
      if (datePart && timePart) {
        fault.timestamp = datePart + 'T' + timePart;
      } else if (datePart) {
        fault.timestamp = datePart;
      } else if (timePart) {
        fault.timestamp = timePart;
      }
    }

    function formatTitle(fault) {
      const employee = fault.employeeName || fault.employeeId || 'עובד לא מזוהה';
      const role = fault.jobName || 'תפקיד חסר';
      const date = fault.workDate || extractDate(fault.timestamp) || 'תאריך חסר';
      return [employee, role, date].filter(Boolean).join(separator);
    }

    function buildMeta(fault) {
      const payKind = normalizePayKind(fault.payType);
      const time = extractTime(fault.timestamp);
      if (payKind === 'unit') {
        return [fault.department || '', fault.units ? fault.units + ' יחידות' : '', fault.payType || '']
          .filter(Boolean)
          .join(separator);
      }
      return [fault.department || '', fault.direction || '', time || '', fault.payType || '']
        .filter(Boolean)
        .join(separator);
    }

    function render() {
      listEl.innerHTML = '';
      faults.forEach((f, idx) => {
        const currentDate = extractDate(f.workDate || f.timestamp);
        const currentTime = extractTime(f.timestamp);
        if (currentDate && !f.workDate) f.workDate = currentDate;
        if (currentTime && !f.fixTime) f.fixTime = currentTime;

        const div = document.createElement('div');
        div.className = 'item';
        const title = document.createElement('div');
        title.className = 'title';
        title.textContent = formatTitle(f);
        div.appendChild(title);

        const meta = document.createElement('div');
        meta.className = 'meta';
        meta.textContent = buildMeta(f);
        div.appendChild(meta);

        const pill = document.createElement('div');
        pill.className = 'pill';
        pill.textContent = 'שורה ' + f.rowIndex;
        div.appendChild(pill);

        const form = document.createElement('div');
        form.className = 'field';

        const addRow = (label, control) => {
          const lab = document.createElement('label');
          lab.textContent = label;
          form.appendChild(lab);
          form.appendChild(control);
        };

        const payKind = normalizePayKind(f.payType);

        const jobInput = document.createElement('input');
        jobInput.value = f.jobName || '';
        jobInput.placeholder = 'שם תפקיד';
        jobInput.oninput = () => {
          f.jobName = jobInput.value;
          title.textContent = formatTitle(f);
        };
        addRow('תפקיד', jobInput);

        const payInput = document.createElement('input');
        payInput.value = f.payType || '';
        payInput.disabled = true;
        payInput.placeholder = 'לא לשינוי ממסך זה';
        addRow('אופן תשלום', payInput);

        const dateInput = document.createElement('input');
        dateInput.type = 'date';
        dateInput.value = f.workDate || '';
        dateInput.oninput = () => {
          f.workDate = dateInput.value;
          updateTimestampFromParts(f, dateInput.value, timeInput ? timeInput.value : '');
          title.textContent = formatTitle(f);
        };

        addRow('תאריך', dateInput);

        let timeInput = null;
        if (payKind !== 'unit') {
          timeInput = document.createElement('input');
          timeInput.type = 'time';
          timeInput.value = f.fixTime || '';
          timeInput.oninput = () => {
            f.fixTime = timeInput.value;
            updateTimestampFromParts(f, f.workDate, timeInput.value);
          };
          addRow('שעה', timeInput);
        }

        if (payKind === 'unit') {
          const unitsInput = document.createElement('input');
          unitsInput.type = 'number';
          unitsInput.min = '0';
          unitsInput.step = 'any';
          unitsInput.value = f.units || '';
          unitsInput.placeholder = 'כמות יחידות';
          unitsInput.oninput = () => {
            f.units = unitsInput.value;
            meta.textContent = buildMeta(f);
          };
          addRow('כמות יחידות', unitsInput);
        } else {
          const directionSelect = document.createElement('select');
          const directions = [
            { value: '', label: 'בחר כיוון' },
            { value: 'כניסה', label: 'כניסה' },
            { value: 'יציאה', label: 'יציאה' },
          ];
          if (f.direction && !directions.find((d) => d.value === f.direction)) {
            directions.splice(1, 0, { value: f.direction, label: f.direction });
          }
          directions.forEach((opt) => {
            const option = document.createElement('option');
            option.value = opt.value;
            option.textContent = opt.label;
            directionSelect.appendChild(option);
          });
          directionSelect.value = f.direction || '';
          directionSelect.onchange = () => {
            f.direction = directionSelect.value;
            meta.textContent = buildMeta(f);
          };
          addRow('כיוון', directionSelect);
        }

        const noteInput = document.createElement('input');
        noteInput.value = f.note || '';
        noteInput.placeholder = 'הוספת הערה';
        noteInput.oninput = () => {
          f.note = noteInput.value;
        };
        addRow('הערה', noteInput);

        div.appendChild(form);

        const reason = document.createElement('div');
        reason.className = 'reason';
        reason.textContent = 'סיבת התקלה: ' + (issueText[f.issue] || f.issue);
        div.appendChild(reason);

        const actions = document.createElement('div');
        actions.className = 'actions';

        if (f.issue === 'missing_shiftId') {
          const fix = document.createElement('button');
          fix.className = 'btn primary';
          fix.textContent = 'צור מזהה דיווח';
          fix.onclick = () => act(f, 'assignShiftId');
          actions.appendChild(fix);
        }

        const save = document.createElement('button');
        save.className = 'btn primary';
        save.textContent = 'שמור שינויים';
        save.onclick = () => act(f, 'saveEdits');
        actions.appendChild(save);

        const ack = document.createElement('button');
        ack.className = 'btn ghost';
        ack.textContent = 'סמן כתקין';
        ack.onclick = () => act(f, 'ack');
        actions.appendChild(ack);

        const del = document.createElement('button');
        del.className = 'btn danger';
        del.textContent = 'מחק שורה';
        del.onclick = () => act(f, 'deleteRow');
        actions.appendChild(del);

        const skip = document.createElement('button');
        skip.className = 'btn ghost';
        skip.textContent = 'דלג';
        skip.onclick = () => removeFault(f.hash, f.issue);
        actions.appendChild(skip);

        div.appendChild(actions);
        listEl.appendChild(div);
      });
      if (!faults.length) {
        listEl.innerHTML = '<div class="meta">אין תקלות להצגה.</div>';
      }
    }

    function removeFault(hash, issue) {
      const idx = faults.findIndex((f) => f.hash === hash && f.issue === issue);
      if (idx >= 0) faults.splice(idx, 1);
      render();
    }

    function act(fault, action) {
      statusEl.style.color = '#1b5e20';
      statusEl.textContent = 'מבצע פעולה...';
      google.script.run
        .withSuccessHandler((res) => {
          if (res && res.ok) {
            statusEl.style.color = '#1b5e20';
            statusEl.textContent = res.message || 'בוצע';
            removeFault(fault.hash, fault.issue);
          } else {
            statusEl.style.color = '#b00020';
            statusEl.textContent = (res && res.error) || 'שגיאה בפעולה';
          }
        })
        .withFailureHandler((err) => {
          statusEl.style.color = '#b00020';
          statusEl.textContent = err && err.message ? err.message : 'שגיאה בפעולה';
        })
        .shiftReport_handleFaultAction({ hash: fault.hash, issue: fault.issue, rowIndex: fault.rowIndex, edits: fault }, action);
    }

    render();
  </script>
</body>
</html>`;
}

function buildDuplicateCleanupDialogHtml_(groups) {
  const safePayload = JSON.stringify(groups || []).replace(/</g, "\\u003c");
  return `
<html dir="rtl">
<head>
  <style>
    body{font-family:Arial,sans-serif;background:#f5f6fa;margin:0;padding:16px;}
    .card{background:#fff;border-radius:12px;box-shadow:0 8px 24px rgba(0,0,0,0.12);padding:16px;}
    .group{border:1px solid #e0e0e0;border-radius:10px;padding:12px;margin-bottom:12px;}
    .title{font-weight:700;margin:0 0 6px;font-size:15px;}
    .row{display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #eee;}
    .row:last-child{border-bottom:none;}
    .pill{border-radius:6px;padding:2px 8px;font-size:12px;background:#eef2ff;color:#1a237e;}
    .keep{color:#1b5e20;font-weight:700;}
    .delete{color:#b00020;font-weight:700;}
    .meta{font-size:12px;color:#555;}
    .actions{display:flex;gap:10px;margin-top:14px;align-items:center;}
    .btn{border:none;border-radius:8px;padding:10px 14px;font-weight:700;cursor:pointer;}
    .ghost{border:1px solid #ccc;background:#fff;color:#333;}
    .danger{background:#d32f2f;color:#fff;}
    .primary{background:#3949ab;color:#fff;}
    #confirm{display:none;position:fixed;inset:0;background:rgba(0,0,0,0.35);justify-content:center;align-items:center;}
    #confirm .box{background:#fff;border-radius:12px;padding:16px;box-shadow:0 10px 32px rgba(0,0,0,0.2);width:420px;}
    #status{margin:8px 0 0;font-weight:700;color:#1b5e20;}
    .cap{font-size:12px;color:#b00020;margin-bottom:8px;}
  </style>
</head>
<body>
  <div class="card">
    <h3 style="margin:0 0 6px;">בדיקת כפילויות משמרות</h3>
    <p style="margin:0 0 10px;color:#555;font-size:13px;">בחרו את כל השורות שאתם רוצים למחוק (ללא הגבלת כמות בקבוצה). אחר כך נבקש אישור פרטני לכל שורה כדי למנוע טעויות.</p>
    <div id="cap" class="cap"></div>
    <div id="list"></div>
    <div class="actions">
      <button class="btn ghost" id="closeBtn">סגירה</button>
      <button class="btn primary" id="scanBtn">המשך למחיקה</button>
    </div>
  </div>

  <div id="confirm">
    <div class="box">
      <h4 style="margin:0 0 8px;">אישור מחיקה</h4>
      <div id="confirmBody" style="font-size:13px;color:#333;"></div>
      <div id="status"></div>
      <div class="actions" style="margin-top:12px;">
        <button class="btn ghost" id="skipBtn">דלג על שורה זו</button>
        <button class="btn ghost" id="cancelConfirm">בטל תהליך המחיקה</button>
        <button class="btn danger" id="confirmYes">מחק שורה זו</button>
      </div>
    </div>
  </div>

  <script>
    const groups = ${safePayload};
    const selected = new Set();
    const listEl = document.getElementById('list');
    const capEl = document.getElementById('cap');
    const confirmEl = document.getElementById('confirm');
    const confirmBody = document.getElementById('confirmBody');
    const statusEl = document.getElementById('status');
    const closeBtn = document.getElementById('closeBtn');
    const scanBtn = document.getElementById('scanBtn');
    const cancelConfirm = document.getElementById('cancelConfirm');
    const confirmYes = document.getElementById('confirmYes');
    const skipBtn = document.getElementById('skipBtn');

    function timeText(rec) {
      if (rec.fixTime) return rec.fixTime;
      if (rec.timestamp) {
        const t = String(rec.timestamp);
        const isoMatch = t.match(/T(\d{2}:\d{2})/);
        if (isoMatch && isoMatch[1]) return isoMatch[1];
        const parts = t.split(' ');
        if (parts.length > 1 && parts[1].match(/\d{1,2}:\d{2}/)) {
          return parts[1].slice(0,5);
        }
      }
      return '';
    }

    function fmt(rec) {
      const time = timeText(rec);
      return [rec.workDate || 'תאריך חסר', time, rec.direction || '', rec.jobName || 'ללא שם עבודה', rec.department || '']
        .filter(Boolean)
        .join(' · ');
    }

    function render() {
      listEl.innerHTML = '';
      groups.forEach((g, idx) => {
        const wrap = document.createElement('div');
        wrap.className = 'group';
        const title = document.createElement('div');
        title.className = 'title';
        title.textContent = 'קבוצה #' + (idx + 1);
        wrap.appendChild(title);

        const keepRow = document.createElement('div');
        keepRow.className = 'row';
        const keepLabel = document.createElement('div');
        keepLabel.className = 'keep';
        keepLabel.textContent = 'יישמר';
        const keepMeta = document.createElement('div');
        keepMeta.className = 'meta';
        keepMeta.textContent = fmt(g.keep);
        keepRow.appendChild(keepLabel);
        keepRow.appendChild(keepMeta);
        wrap.appendChild(keepRow);

        (g.candidates || []).forEach((c) => {
          const row = document.createElement('div');
          row.className = 'row';
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.name = 'del-' + g.key + '-' + c.shiftId;
          checkbox.checked = selected.has(c.shiftId);
          checkbox.onchange = () => {
            if (checkbox.checked) {
              selected.add(c.shiftId);
            } else {
              selected.delete(c.shiftId);
            }
            updateCap();
          };

          const delLabel = document.createElement('div');
          delLabel.className = 'delete';
          delLabel.textContent = 'מועמד למחיקה';
          const meta = document.createElement('div');
          meta.className = 'meta';
          meta.textContent = fmt(c);

          row.appendChild(checkbox);
          row.appendChild(delLabel);
          row.appendChild(meta);
          wrap.appendChild(row);
        });

        listEl.appendChild(wrap);
      });

      updateCap();
    }

    function updateCap() {
      const planned = selected.size;
      if (planned > 20) {
        capEl.textContent = 'אפשר למחוק עד 20 שורות בכל הרצה. בחרו עד 20.';
      } else if (planned === 0) {
        capEl.textContent = 'סמנו שורות למחיקה, ואז לחצו המשך.';
      } else {
        capEl.textContent = 'מסומנות ' + planned + ' שורות למחיקה.';
      }
    }

    function buildQueue() {
      const queue = [];
      groups.forEach(g => {
        (g.candidates || []).forEach(c => {
          if (selected.has(c.shiftId)) {
            queue.push({ group: g, del: c });
          }
        });
      });
      return queue;
    }

    function showConfirm(item) {
      statusEl.textContent = '';
      statusEl.style.color = '#1b5e20';
      const others = (item.group.candidates || []).filter(c => c.shiftId !== item.del.shiftId);
      const othersHtml = others.length
        ? '<div style="margin-top:8px;font-size:12px;color:#555;">חלופות בקבוצה: ' + others.map(fmt).join(' | ') + '</div>'
        : '';
      confirmBody.innerHTML =
        '<div style="margin-bottom:8px;">' +
        '<div class="keep">יישאר:</div>' +
        '<div class="meta">' + fmt(item.group.keep) + '</div>' +
        '</div>' +
        '<div>' +
        '<div class="delete">יימחק:</div>' +
        '<div class="meta">' + fmt(item.del) + '</div>' +
        '</div>' +
        othersHtml;
      confirmEl.style.display = 'flex';
    }

    function processQueue(queue, idx) {
      if (idx >= queue.length) {
        statusEl.textContent = 'הכל הסתיים. לא נותרו שורות לסריקה.';
        setTimeout(() => google.script.host.close(), 700);
        return;
      }
      const item = queue[idx];
      showConfirm(item);

      const doDelete = () => {
        google.script.run
          .withSuccessHandler(res => {
            if (!res || res.ok !== true) {
              statusEl.style.color = '#b00020';
              statusEl.textContent = (res && res.error) || 'מחיקה נכשלה.';
              return;
            }
            confirmEl.style.display = 'none';
            processQueue(queue, idx + 1);
          })
          .withFailureHandler(err => {
            statusEl.style.color = '#b00020';
            statusEl.textContent = (err && err.message) || 'שגיאה במחיקה.';
          })
          .shiftReport_deleteDuplicatesFromUi([item.del.shiftId]);
      };

      const doSkip = () => {
        confirmEl.style.display = 'none';
        processQueue(queue, idx + 1);
      };

      const doCancelAll = () => {
        confirmEl.style.display = 'none';
        statusEl.style.color = '#b00020';
        statusEl.textContent = 'תהליך המחיקה בוטל. שורות שנותרו לא נמחקו.';
      };

      confirmYes.onclick = doDelete;
      skipBtn.onclick = doSkip;
      cancelConfirm.onclick = doCancelAll;
    }

    scanBtn.onclick = () => {
      const queue = buildQueue();
      if (!queue.length) {
        capEl.textContent = 'לא נבחרו שורות למחיקה.';
        return;
      }
      if (queue.length > 20) {
        capEl.textContent = 'בחרו עד 20 שורות למחיקה בכל הרצה.';
        return;
      }
      processQueue(queue, 0);
    };

    closeBtn.onclick = () => google.script.host.close();

    render();
  </script>
</body>
</html>`;
}

// UI helper: build shifts submenu (chunked refresh actions).
function buildShiftsSubMenu_(ui) {
  return ui
    .createMenu("משמרות")
    .addItem("ריענון היום", "rebuildShiftsTodayMenuAction")
    .addItem("ריענון 7 ימים", "rebuildShiftsLast7MenuAction")
    .addItem("ריענון 14 ימים", "rebuildShiftsLast14MenuAction")
    .addItem("ריענון 30 ימים", "rebuildShiftsLast30MenuAction");
}

// Helper: add a menu item only when its handler exists to avoid breaking onOpen.
function maybeAddMenuItem_(menu, caption, handlerName) {
  try {
    if (!menu || !handlerName) return;

    var fn = null;
    try {
      if (typeof globalThis !== "undefined") {
        fn = globalThis[handlerName];
      } else {
        // Apps Script global fallback
        fn = this[handlerName];
      }
    } catch (_e) {
      // ignore lookup errors
    }

    if (typeof fn === "function") {
      menu.addItem(caption, handlerName);
    }
  } catch (err) {
    try {
      Logger.log("maybeAddMenuItem_ error: " + err);
    } catch (_ignored) {}
  }
}

function appendLegacyBoulderMenuItems_(menu) {
  // These handlers are added only if they exist, so we do not break missing functions.
  maybeAddMenuItem_(menu, "פתח סייד בר עובדים", "EMP_openSidebar");
  maybeAddMenuItem_(menu, "רענן סייד בר", "EMP_reloadSidebar");
  maybeAddMenuItem_(
    menu,
    "בדיקת באקפיל IDs (DRY_RUN)",
    "EMP_menuBackfillIdsDryRun",
  );
  maybeAddMenuItem_(
    menu,
    "באקפיל IDs לכל העובדים (EXECUTE)",
    "EMP_menuBackfillIdsExecute",
  );
}

// --- Game leaderboard (best-per-employee, capped) ---

function ensureLeaderboardSheet_() {
  const ss = getSpreadsheet_();
  const headers = [
    "timestamp",
    "employeeId",
    "employeeName",
    "email",
    "score",
    "bestTimeMs",
  ];

  let sheet = ss.getSheetByName(LEADERBOARD_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LEADERBOARD_SHEET_NAME);
    sheet.appendRow(headers);
    sheet.hideSheet();
    return { sheet: sheet, headers: headers };
  }

  // Normalize header row and keep the sheet hidden.
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (!sheet.isSheetHidden()) {
    try {
      sheet.hideSheet();
    } catch (_err) {
      // Ignore hide errors.
    }
  }
  return { sheet: sheet, headers: headers };
}

function readLeaderboardRows_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  return values
    .map(function (row) {
      return {
        timestamp: stringValue(row[0]) || new Date().toISOString(),
        employeeId: stringValue(row[1]),
        employeeName: stringValue(row[2]),
        email: stringValue(row[3]),
        score: Number(row[4]) || 0,
        bestTimeMs: Number(row[5]) || 0,
      };
    })
    .filter(function (item) {
      return item.employeeId;
    });
}

function sortLeaderboardRows_(rows) {
  return rows.sort(function (a, b) {
    const scoreDiff =
      (b && b.score ? b.score : 0) - (a && a.score ? a.score : 0);
    if (scoreDiff !== 0) return scoreDiff;

    const timeA = a && a.bestTimeMs ? a.bestTimeMs : Number.POSITIVE_INFINITY;
    const timeB = b && b.bestTimeMs ? b.bestTimeMs : Number.POSITIVE_INFINITY;
    if (timeA !== timeB) return timeA - timeB;

    const tsA = a && a.timestamp ? new Date(a.timestamp).getTime() : 0;
    const tsB = b && b.timestamp ? new Date(b.timestamp).getTime() : 0;
    return tsB - tsA;
  });
}

function writeLeaderboardRows_(sheet, headers, rows) {
  // Clear existing body and rewrite to keep things compact.
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (!rows || !rows.length) return;
  const matrix = rows.map(function (r) {
    return [
      r.timestamp || new Date().toISOString(),
      r.employeeId || "",
      r.employeeName || "",
      r.email || "",
      Number(r.score) || 0,
      Number(r.bestTimeMs) || 0,
    ];
  });
  sheet.getRange(2, 1, matrix.length, headers.length).setValues(matrix);
}

function listGameLeaderboard_(payload) {
  Logger.log(
    "LEADERBOARD_LIST called with payload: %s",
    JSON.stringify(payload),
  );
  const logger = ensureModuleLoggerDefined_("LEADERBOARD_LIST");
  const meta = ensureLeaderboardSheet_();
  const rows = sortLeaderboardRows_(
    readLeaderboardRows_(meta.sheet, meta.headers),
  );
  const top5 = rows.slice(0, 5).map(function (row, idx) {
    return {
      rank: idx + 1,
      name: row.employeeName || row.email || "Unknown",
      score: row.score,
      bestTimeMs: row.bestTimeMs,
      employeeId: row.employeeId,
      email: row.email,
      timestamp: row.timestamp,
    };
  });

  logger.info("leaderboard.list", { rows: rows.length, top: top5.length });
  return { ok: true, entries: top5 };
}

function submitGameLeaderboard_(payload) {
  Logger.log(
    "LEADERBOARD_SAVE called with payload: %s",
    JSON.stringify(payload),
  );
  const logger = ensureModuleLoggerDefined_("LEADERBOARD_SAVE");
  const employeeId = stringValue(payload.employeeId);
  const employeeName = stringValue(payload.employeeName);
  const email = stringValue(payload.email);
  const score = Number(payload.score);
  const bestTimeMs = Number(payload.bestTimeMs || payload.timeMs || 0);

  if (!employeeId || !employeeName || !email) {
    return {
      ok: false,
      error: "missing required fields",
      errorCode: "VALIDATION_FAILED",
    };
  }
  if (!isFinite(score) || score < 0) {
    return {
      ok: false,
      error: "invalid score",
      errorCode: "VALIDATION_FAILED",
    };
  }

  const meta = ensureLeaderboardSheet_();
  const existing = sortLeaderboardRows_(
    readLeaderboardRows_(meta.sheet, meta.headers),
  );

  let changed = false;
  const nowIso = new Date().toISOString();
  const updatedRows = existing.map(function (row) {
    if (row.employeeId !== employeeId) return row;
    const priorTime =
      row.bestTimeMs > 0 ? row.bestTimeMs : Number.POSITIVE_INFINITY;
    const candidateTime =
      bestTimeMs > 0 ? bestTimeMs : Number.POSITIVE_INFINITY;
    const isBetter =
      score > row.score ||
      (Math.abs(score - row.score) < 1e-9 && candidateTime < priorTime);
    if (!isBetter) return row;
    changed = true;
    return {
      timestamp: nowIso,
      employeeId: employeeId,
      employeeName: employeeName,
      email: email,
      score: score,
      bestTimeMs: bestTimeMs,
    };
  });

  if (
    !existing.some(function (r) {
      return r.employeeId === employeeId;
    })
  ) {
    changed = true;
    updatedRows.push({
      timestamp: nowIso,
      employeeId: employeeId,
      employeeName: employeeName,
      email: email,
      score: score,
      bestTimeMs: bestTimeMs,
    });
  }

  if (!changed) {
    return { ok: true, updated: false };
  }

  const sorted = sortLeaderboardRows_(updatedRows);
  const pruned = sorted.slice(0, LEADERBOARD_MAX_ROWS);
  writeLeaderboardRows_(meta.sheet, meta.headers, pruned);
  logger.info("leaderboard.upsert", {
    employeeId: employeeId,
    total: pruned.length,
  });

  const top5 = pruned.slice(0, 5).map(function (row, idx) {
    return {
      rank: idx + 1,
      name: row.employeeName || row.email || "Unknown",
      score: row.score,
      bestTimeMs: row.bestTimeMs,
      employeeId: row.employeeId,
      email: row.email,
      timestamp: row.timestamp,
    };
  });

  return { ok: true, updated: true, entries: top5 };
}

function rebuildShiftsTodayMenuAction() {
  const todayIso = toIsoDate_(new Date());
  return rebuildShiftsChunked_(todayIso, todayIso, 1);
}

function rebuildShiftsLast7MenuAction() {
  return rebuildShiftsSlidingWindow_(7);
}

function rebuildShiftsLast14MenuAction() {
  return rebuildShiftsSlidingWindow_(14);
}

function rebuildShiftsLast30MenuAction() {
  return rebuildShiftsSlidingWindow_(30);
}

function rebuildShiftsSlidingWindow_(daysBack) {
  const today = new Date();
  const dateTo = toIsoDate_(today);
  const dateFromDate = new Date(today.getTime());
  dateFromDate.setDate(today.getDate() - Math.max(0, Number(daysBack) || 0));
  const dateFrom = toIsoDate_(dateFromDate);
  return rebuildShiftsChunked_(dateFrom, dateTo, 7);
}

// Rebuild shifts in small slices to reduce runtime/memory.
function rebuildShiftsChunked_(dateFromIso, dateToIso, chunkDays) {
  const start = isoToDate_(dateFromIso);
  const end = isoToDate_(dateToIso);
  if (!start || !end || start > end) {
    return { ok: false, error: "invalid range" };
  }

  const maxSpanDays = 40;
  const spanDays =
    Math.floor((end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000)) + 1;
  if (spanDays > maxSpanDays) {
    return { ok: false, error: "range too large", spanDays: spanDays };
  }

  const step = Math.max(1, Number(chunkDays) || 7);
  const summaries = [];
  let cursor = new Date(start.getTime());

  while (cursor <= end) {
    const chunkStart = new Date(cursor.getTime());
    const chunkEnd = new Date(cursor.getTime());
    chunkEnd.setDate(chunkEnd.getDate() + step - 1);
    if (chunkEnd > end) chunkEnd.setTime(end.getTime());

    const fromIso = toIsoDate_(chunkStart);
    const toIso = toIsoDate_(chunkEnd);

    try {
      const res = rebuildShiftsForRange_({ dateFrom: fromIso, dateTo: toIso });
      summaries.push({ from: fromIso, to: toIso, result: res });
    } catch (err) {
      summaries.push({ from: fromIso, to: toIso, error: String(err) });
    }

    cursor.setDate(cursor.getDate() + step);
  }

  return { ok: true, chunks: summaries.length, summaries: summaries };
}

function showHourlyOverlapsDialog() {
  var html = HtmlService.createHtmlOutputFromFile("OverlapsDialog")
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "חפיפות בשכר שעתי");
}
