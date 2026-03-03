/* global EMP_getEmployeeById, EMP_saveEmployeePayload */
/* global SHIFTS_updateShift, SHIFTS_upsertAroundWorkLog_, handleShiftReportSubmit_ */
/* global REQ_listRequests, AUDIT_logEvent, getAllJobs */
/* global SHIFTS_list, EMP_getSidebarBootstrap, handleRequestPost */
/* exported UI_saveEmployeeDialogChanges, UI_saveShiftChanges */
/* exported UI_getRequestsDialogData, UI_saveRequestChanges */
/* exported UI_debugOpenEmployeeDialog, UI_debugOpenShiftsDialog, UI_debugOpenRequestsDialog */

function getSidebarLogger_(operation) {
  try {
    if (typeof ensureModuleLoggerDefined_ === "function") {
      return ensureModuleLoggerDefined_(operation || "SIDEBAR_DIALOG");
    }
    if (typeof createScriptLogger_ === "function") {
      return createScriptLogger_({
        traceId:
          (typeof makeTraceId_ === "function" && makeTraceId_()) ||
          (typeof Utilities !== "undefined" && Utilities.getUuid
            ? Utilities.getUuid()
            : ""),
        operation: operation || "SIDEBAR_DIALOG",
        source: "apps-script-sidebar",
      });
    }
  } catch (_ignored) {}
  return null;
}

/** Opens the extended employee dialog. */
function UI_openEmployeeDialog(params) {
  var safeParams = params && typeof params === "object" ? params : {};

  var template = HtmlService.createTemplateFromFile("ExtendedDialog");
  template.mode = "employee";
  template.title = "כרטיס עובד מורחב";
  template.paramsJson = JSON.stringify(safeParams || {});
  var html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, "כרטיס עובד");
}

/** Returns data for the extended employee dialog. */
function UI_getEmployeeDialogData(params) {
  var logger = getSidebarLogger_("SIDEBAR_EMPLOYEE_DIALOG");
  var startMs = new Date().getTime();
  var traceId =
    (typeof Utilities !== "undefined" &&
      Utilities.getUuid &&
      Utilities.getUuid()) ||
    "";
  var requestedAt = new Date();
  var actorEmail = "";
  try {
    actorEmail = Session.getActiveUser().getEmail() || "";
  } catch (_e) {
    actorEmail = "";
  }
  var actorRole = "MANAGER"; // placeholder until role logic is added

  var errors = [];
  var data = null;

  var safeParams = params && typeof params === "object" ? params : {};
  var employeeId = safeParams.employeeId
    ? String(safeParams.employeeId).trim()
    : "";

  if (!employeeId) {
    errors.push({
      code: "NO_EMPLOYEE_ID",
      message: "לא נבחר עובד להצגה.",
      field: "employeeId",
    });
  } else {
    var empRes = null;
    try {
      if (typeof EMP_getEmployeeById === "function") {
        empRes = EMP_getEmployeeById(employeeId);
      }
    } catch (_empErr) {
      empRes = { ok: false, error: "EMP_getEmployeeById threw" };
    }

    if (!empRes || !empRes.ok || !empRes.employee || !empRes.employee.id) {
      errors.push({
        code: "EMPLOYEE_NOT_FOUND",
        message: "העובד המבוקש לא נמצא.",
        field: "employeeId",
      });
    } else {
      var employee = empRes.employee;

      var allJobTypes = [];
      try {
        if (typeof OPT !== "undefined" && OPT.getAllJobs) {
          allJobTypes = OPT.getAllJobs(false) || [];
        } else if (typeof getAllJobs === "function") {
          allJobTypes = getAllJobs(false) || [];
        }
      } catch (_optErr) {
        allJobTypes = [];
      }

      var employeeJobs = buildEmployeeJobsFromEmployee_(employee);

      data = {
        employee: employee,
        employeeJobs: employeeJobs,
        allJobTypes: allJobTypes,
        stats: {},
        recentShifts: [],
        recentRequests: [],
      };
    }
  }

  var result = {
    meta: {
      traceId: traceId,
      requestedAt: requestedAt,
      actorEmail: actorEmail,
      actorRole: actorRole,
      canEdit: errors.length === 0,
    },
    data: errors.length ? null : data,
    errors: errors,
  };

  logDuration_(logger, "sidebar.employee.get", startMs, {
    employeeId: employeeId,
    errors: errors.length,
    source: "extended-dialog",
  });

  return result;
}

function buildEmployeeJobsFromEmployee_(employee) {
  if (!employee || !employee.rows || !employee.rows.length) return [];
  return employee.rows.map(function (r) {
    return {
      rowIndex: r.rowIndex,
      active: !!r.rowActive,
      jobType: r.jobType || "",
      department: r.department || "",
      amount: r.amount || "",
      paymentMode: r.paymentMode || "",
      notes: r.notes || "",
    };
  });
}

/** Saves changes made in the extended employee dialog. */
// eslint-disable-next-line no-unused-vars
function UI_saveEmployeeDialogChanges(envelope) {
  var logger = getSidebarLogger_("SIDEBAR_EMPLOYEE_SAVE");
  var startMs = new Date().getTime();
  var traceId =
    (typeof Utilities !== "undefined" &&
      Utilities.getUuid &&
      Utilities.getUuid()) ||
    "";
  var requestedAt = new Date();
  var actorEmail = "";
  try {
    actorEmail = Session.getActiveUser().getEmail() || "";
  } catch (_e) {
    actorEmail = "";
  }
  var actorRole = "MANAGER";

  var errors = [];
  var safeEnvelope = envelope && typeof envelope === "object" ? envelope : {};
  var candidate =
    safeEnvelope.payload ||
    safeEnvelope.employee ||
    safeEnvelope.data ||
    safeEnvelope;
  var employeePayload =
    candidate && typeof candidate === "object" && candidate.employee
      ? candidate.employee
      : candidate;

  var toSave = null;
  if (!employeePayload || typeof employeePayload !== "object") {
    errors.push({
      code: "INVALID_PAYLOAD",
      message: "נתוני השמירה אינם תקינים.",
    });
  } else {
    toSave = Object.assign({}, employeePayload);
    if (!toSave.id && toSave.employeeId) {
      toSave.id = String(toSave.employeeId).trim();
    }

    if (!toSave.id) {
      errors.push({
        code: "NO_EMPLOYEE_ID",
        message: "חסר מזהה עובד לשמירה.",
        field: "employeeId",
      });
    }

    var nameVal = toSave.name ? String(toSave.name).trim() : "";
    if (!nameVal) {
      errors.push({
        code: "NO_EMPLOYEE_NAME",
        message: "שם עובד חובה לשמירה.",
        field: "name",
      });
    } else {
      toSave.name = nameVal;
    }

    if (!Array.isArray(toSave.rows) || !toSave.rows.length) {
      errors.push({
        code: "NO_JOB_ROWS",
        message: "יש לספק לפחות שורת סוג עבודה אחת.",
        field: "rows",
      });
    }
  }

  var saveResult = null;
  if (!errors.length) {
    try {
      if (typeof EMP_saveEmployeePayload === "function") {
        saveResult = EMP_saveEmployeePayload(toSave);
      } else {
        saveResult = { ok: false, error: "EMP_saveEmployeePayload missing" };
      }
    } catch (errSave) {
      saveResult = { ok: false, error: String(errSave) };
    }

    if (!saveResult || saveResult.ok !== true) {
      errors.push({
        code: "SAVE_FAILED",
        message:
          "שמירת העובד נכשלה" +
          (saveResult && saveResult.error ? ": " + saveResult.error : ""),
      });
    }
  }

  var dialogRes = null;
  if (!errors.length) {
    try {
      dialogRes = UI_getEmployeeDialogData({ employeeId: toSave.id });
    } catch (errReload) {
      errors.push({
        code: "RELOAD_FAILED",
        message: "טעינת הנתונים לאחר שמירה נכשלה: " + errReload,
      });
    }

    if (
      dialogRes &&
      Array.isArray(dialogRes.errors) &&
      dialogRes.errors.length
    ) {
      errors = errors.concat(dialogRes.errors);
    }
  }

  if (!errors.length && typeof AUDIT_logEvent === "function") {
    try {
      AUDIT_logEvent({
        eventType: "employee_dialog_save",
        entityType: "employee",
        entityId: toSave.id || "",
        actorEmail: actorEmail,
        actorRole: actorRole,
        traceId: traceId,
        summaryAfter: toSave.name || "",
        extra: {
          rows: Array.isArray(toSave.rows) ? toSave.rows.length : 0,
        },
      });
    } catch (_auditErr) {
      // best-effort logging only
    }
  }

  var result = {
    meta: {
      traceId: traceId,
      requestedAt: requestedAt,
      actorEmail: actorEmail,
      actorRole: actorRole,
      canEdit: errors.length === 0,
    },
    data:
      saveResult || dialogRes
        ? {
            employeeId: toSave && toSave.id ? toSave.id : "",
            saveResult: saveResult,
            dialogData: dialogRes ? dialogRes.data : null,
            dialogMeta: dialogRes ? dialogRes.meta : null,
          }
        : null,
    errors: errors,
  };

  logDuration_(logger, "sidebar.employee.save", startMs, {
    employeeId: toSave && toSave.id ? toSave.id : "",
    errors: errors.length,
    source: "extended-dialog",
  });

  return result;
}

/** Opens the extended shifts dialog. */
function UI_openShiftsDialog(params) {
  var safeParams = params && typeof params === "object" ? params : {};

  var template = HtmlService.createTemplateFromFile("ExtendedDialog");
  template.mode = "shifts";
  template.title = "משמרות מורחב";
  template.paramsJson = JSON.stringify(safeParams || {});
  var html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, "משמרות");
}

/** Returns data for the extended shifts dialog. */
function UI_getShiftsDialogData(params) {
  var logger = getSidebarLogger_("SIDEBAR_SHIFTS_DIALOG");
  var startMs = new Date().getTime();
  var traceId =
    (typeof Utilities !== "undefined" &&
      Utilities.getUuid &&
      Utilities.getUuid()) ||
    "";
  var requestedAt = new Date();
  var actorEmail = "";
  try {
    actorEmail = Session.getActiveUser().getEmail() || "";
  } catch (_e) {
    actorEmail = "";
  }
  var actorRole = "MANAGER"; // placeholder until role logic is added

  var errors = [];
  var data = null;

  var safeParams = params && typeof params === "object" ? params : {};

  var range = normalizeDateRange_(safeParams.fromDate, safeParams.toDate);
  if (!range.ok) {
    errors.push(range.error);
  }

  var employeeId = safeParams.employeeId
    ? String(safeParams.employeeId).trim()
    : "";
  var jobId = safeParams.jobId ? String(safeParams.jobId).trim() : ""; // echoed only; Shifts filter does not use jobId currently
  var status = safeParams.status ? String(safeParams.status).trim() : "";

  var shifts = [];
  var employeesLookup = [];
  var jobTypesLookup = [];
  var statusesLookup = [];
  var totalShifts = 0;
  var returnedShifts = 0;
  var hasMore = false;

  if (errors.length === 0) {
    var filters = {
      dateFrom: range.fromIso,
      dateTo: range.toIso,
    };
    if (employeeId) filters.employeeId = employeeId;
    if (status) filters.statuses = [status];

    var shiftsRes = null;
    try {
      if (typeof SHIFTS_list === "function") {
        shiftsRes = SHIFTS_list(filters);
      }
    } catch (errShifts) {
      shiftsRes = { ok: false, error: String(errShifts) };
    }

    if (!shiftsRes || !shiftsRes.ok) {
      errors.push({
        code: "SHIFTS_LIST_FAILED",
        message:
          "קריאת משמרות נכשלה" +
          (shiftsRes && shiftsRes.error ? ": " + shiftsRes.error : ""),
      });
    } else {
      totalShifts = shiftsRes.total || 0;
      returnedShifts = shiftsRes.returned || 0;
      hasMore = !!shiftsRes.hasMore;
      statusesLookup = Array.isArray(shiftsRes.statuses)
        ? shiftsRes.statuses
        : [];

      shifts = (shiftsRes.shifts || []).map(function (s) {
        var logicalDate = buildLogicalDate_(s);
        return {
          id: s.shiftId || "",
          employeeId: s.employeeId || "",
          employeeName: s.employeeName || "",
          jobId: s.workTypeId || "",
          jobName: s.workType || "",
          rawDate: s.date || "",
          logicalDate: logicalDate,
          startTime: s.startTime || "",
          endTime: s.endTime || "",
          status: s.status || "",
          direction: s.direction || "",
          department: s.department || "",
          email: s.email || "",
          hasIssues: !!s.hasIssues,
          isManualEdited: !!s.isManualEdited,
          bonusIds: s.bonusIds || [],
          manualNote: s.manualNote || "",
          durationHours: s.spanHours !== undefined ? s.spanHours : null,
          payHours: s.payHours !== undefined ? s.payHours : null,
          source: s.source || "",
        };
      });
    }

    // Employees lookup
    try {
      if (typeof EMP_getSidebarBootstrap === "function") {
        var bootstrap = EMP_getSidebarBootstrap();
        if (
          bootstrap &&
          bootstrap.employees &&
          Array.isArray(bootstrap.employees)
        ) {
          employeesLookup = bootstrap.employees.map(function (e) {
            return {
              id: e.id || "",
              name: e.name || "",
              email: e.email || "",
            };
          });
        }
      }
    } catch (_empErr) {
      // ignore lookup failure; not critical for data payload
    }

    // Job types lookup
    try {
      if (typeof OPT !== "undefined" && OPT.getAllJobs) {
        jobTypesLookup = OPT.getAllJobs(false) || [];
      } else if (typeof getAllJobs === "function") {
        jobTypesLookup = getAllJobs(false) || [];
      }
      jobTypesLookup = (jobTypesLookup || []).map(function (j) {
        var statusVal = String(j.status || "").trim();
        var isActive =
          !statusVal ||
          statusVal === "פעיל" ||
          statusVal.toLowerCase() === "active";
        return {
          id: j.id || "",
          name: j.name || "",
          department: j.department || "",
          isActive: isActive,
          rawStatus: j.status || "",
        };
      });
    } catch (_optErr) {
      jobTypesLookup = [];
    }
  }

  data = errors.length
    ? null
    : {
        filtersEcho: {
          fromDate: range.fromIso,
          toDate: range.toIso,
          employeeId: employeeId || "",
          jobId: jobId || "",
          status: status || "",
        },
        shifts: shifts,
        employeesLookup: employeesLookup,
        jobTypesLookup: jobTypesLookup,
        statuses: statusesLookup,
        metaStats: {
          total: totalShifts,
          returned: returnedShifts,
          hasMore: hasMore,
        },
        rules: { crossMidnightCutoffHour: 4 },
      };

  var result = {
    meta: {
      traceId: traceId,
      requestedAt: requestedAt,
      actorEmail: actorEmail,
      actorRole: actorRole,
      canEdit: errors.length === 0,
    },
    data: data,
    errors: errors,
  };

  logDuration_(logger, "sidebar.shifts.get", startMs, {
    errors: errors.length,
    returnedShifts: returnedShifts,
    source: "extended-dialog",
  });

  return result;
}

function normalizeDateRange_(fromDateStr, toDateStr) {
  var today = new Date();
  var defaultFrom = new Date(today.getTime());
  defaultFrom.setDate(today.getDate() - 7);
  var defaultTo = new Date(today.getTime());
  defaultTo.setDate(today.getDate() + 1);

  var from = parseIsoDate_(fromDateStr) || defaultFrom;
  var to = parseIsoDate_(toDateStr) || defaultTo;

  if (from > to) {
    return {
      ok: false,
      error: {
        code: "INVALID_DATE_RANGE",
        message: "טווח התאריכים אינו תקין (מתחיל אחרי הסוף)",
        field: "fromDate",
      },
    };
  }

  return {
    ok: true,
    fromIso: formatIsoDate_(from),
    toIso: formatIsoDate_(to),
  };
}

function parseIsoDate_(val) {
  if (!val) return null;
  var s = String(val).trim();
  if (!s) return null;
  var parts = s.split("-");
  if (parts.length !== 3) return null;
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  var dt = new Date(y, m, d);
  return isNaN(dt.getTime()) ? null : dt;
}

function formatIsoDate_(dt) {
  var y = dt.getFullYear();
  var m = dt.getMonth() + 1;
  var d = dt.getDate();
  return (
    y +
    "-" +
    (m < 10 ? "0" + m : String(m)) +
    "-" +
    (d < 10 ? "0" + d : String(d))
  );
}

function resolveActorContext_() {
  var actorEmail = "";
  try {
    actorEmail = Session.getActiveUser().getEmail() || "";
  } catch (_e) {
    actorEmail = "";
  }
  var actorRole = actorEmail ? "MANAGER" : "UNKNOWN";
  return { actorEmail: actorEmail, actorRole: actorRole, actorId: "" };
}

function deriveActorRole_(actorEmail, employeesLookup) {
  if (!actorEmail) return "UNKNOWN";
  var emailKey = String(actorEmail).toLowerCase();
  if (Array.isArray(employeesLookup)) {
    for (var i = 0; i < employeesLookup.length; i++) {
      var e = employeesLookup[i];
      var eEmail = e && e.email ? String(e.email).toLowerCase() : "";
      if (eEmail && eEmail === emailKey) {
        return "EMPLOYEE";
      }
    }
  }
  return "MANAGER";
}

function buildLogicalDate_(shift) {
  if (!shift) return "";
  // The module's date field already derives a date from workDate/fixDate/timestamp.
  if (shift.date) return shift.date;
  return "";
}

function normRequestVal_(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s+/g, " ").trim();
}

function deriveRequestType_(rec) {
  if (!rec) return "";
  var direction = normRequestVal_(rec.direction);
  if (direction) return direction;
  return "";
}

function detectCancelledStatusValue_(statuses) {
  if (!Array.isArray(statuses)) return "";
  for (var i = 0; i < statuses.length; i++) {
    var s = statuses[i];
    var skey = normRequestVal_(s).toLowerCase();
    if (!skey) continue;
    if (skey.indexOf("בוטל") !== -1 || skey.indexOf("בטלה") !== -1) return s;
    if (skey.indexOf("cancel") !== -1) return s;
  }
  return "";
}

function isCancelledStatus_(status, cancelledStatusValue) {
  var skey = normRequestVal_(status).toLowerCase();
  if (!skey) return false;
  var cancelledKey = normRequestVal_(cancelledStatusValue).toLowerCase();
  if (cancelledKey && skey === cancelledKey) return true;
  return skey.indexOf("בוטל") !== -1 || skey.indexOf("cancel") !== -1;
}

function canEmployeeEditRequestStatus_(status, cancelledStatusValue) {
  var skey = normRequestVal_(status).toLowerCase();
  var cancelledKey = normRequestVal_(cancelledStatusValue).toLowerCase();
  var closedKeys = [
    "מאושרת",
    "אושר",
    "אושרה",
    "נדחתה",
    "נדחה",
    "דחוי",
    "דחויה",
    "approved",
    "rejected",
    "denied",
    "declined",
  ].map(function (s) {
    return normRequestVal_(s).toLowerCase();
  });
  if (cancelledKey) closedKeys.push(cancelledKey);
  return closedKeys.indexOf(skey) === -1;
}

function buildUniqueValuePairs_(items, key) {
  var seen = {};
  var out = [];
  if (!Array.isArray(items)) return out;
  for (var i = 0; i < items.length; i++) {
    var val = items[i] && items[i][key] !== undefined ? items[i][key] : "";
    var normVal = normRequestVal_(val);
    if (!normVal) continue;
    if (seen[normVal]) continue;
    seen[normVal] = true;
    out.push({ value: val || normVal, label: val || normVal });
  }
  return out;
}

/** Saves changes made in the extended shifts dialog. */
// eslint-disable-next-line no-unused-vars
function UI_saveShiftChanges(envelope) {
  var logger = getSidebarLogger_("SIDEBAR_SHIFTS_SAVE");
  var startMs = new Date().getTime();
  var traceId =
    (typeof Utilities !== "undefined" &&
      Utilities.getUuid &&
      Utilities.getUuid()) ||
    "";
  var requestedAt = new Date();
  var actorEmail = "";
  try {
    actorEmail = Session.getActiveUser().getEmail() || "";
  } catch (_e) {
    actorEmail = "";
  }
  var actorRole = "MANAGER"; // placeholder until role logic is added

  var errors = [];
  var safeEnvelope = envelope && typeof envelope === "object" ? envelope : {};
  var candidate =
    safeEnvelope.payload ||
    safeEnvelope.shift ||
    safeEnvelope.data ||
    safeEnvelope;

  var shiftPayload =
    candidate && typeof candidate === "object" && candidate.shift
      ? candidate.shift
      : candidate;

  var filtersEcho =
    safeEnvelope.filters && typeof safeEnvelope.filters === "object"
      ? safeEnvelope.filters
      : {};

  function normStr_(v) {
    if (v === null || v === undefined) return "";
    return String(v).trim();
  }

  function parseTimeToParts_(val) {
    var s = normStr_(val);
    if (!s) return null;
    var parts = s.split(":");
    var hh = parseInt(parts[0], 10);
    var mm = parts.length > 1 ? parseInt(parts[1], 10) : 0;
    if (isNaN(hh) || isNaN(mm)) return null;
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return { hours: hh, minutes: mm };
  }

  function buildDateTime_(dateIso, timeStr) {
    var dt = parseIsoDate_(dateIso);
    var t = parseTimeToParts_(timeStr);
    if (!dt || !t) return null;
    var out = new Date(dt.getTime());
    out.setHours(t.hours);
    out.setMinutes(t.minutes);
    out.setSeconds(0);
    out.setMilliseconds(0);
    return out;
  }

  function timeStr_(dt) {
    if (!(dt instanceof Date) || isNaN(dt)) return "";
    var h = dt.getHours();
    var m = dt.getMinutes();
    return (
      (h < 10 ? "0" + h : String(h)) + ":" + (m < 10 ? "0" + m : String(m))
    );
  }

  if (!shiftPayload || typeof shiftPayload !== "object") {
    errors.push({
      code: "INVALID_PAYLOAD",
      message: "נתוני המשמרת אינם תקינים.",
    });
  }

  var shiftId =
    shiftPayload && shiftPayload.id ? normStr_(shiftPayload.id) : "";
  var employeeId = shiftPayload ? normStr_(shiftPayload.employeeId) : "";
  var jobId = shiftPayload
    ? normStr_(shiftPayload.jobId || shiftPayload.workTypeId)
    : "";
  var jobName = shiftPayload
    ? normStr_(shiftPayload.jobName || shiftPayload.workType)
    : "";
  var department = shiftPayload ? normStr_(shiftPayload.department) : "";
  var status = shiftPayload ? normStr_(shiftPayload.status) : "";
  var shiftDate = shiftPayload
    ? normStr_(
        shiftPayload.date ||
          shiftPayload.shiftDate ||
          shiftPayload.rawDate ||
          shiftPayload.logicalDate
      )
    : "";
  var startTime = shiftPayload ? normStr_(shiftPayload.startTime) : "";
  var endTime = shiftPayload ? normStr_(shiftPayload.endTime) : "";
  var manualNote = shiftPayload
    ? normStr_(shiftPayload.manualNote || shiftPayload.note)
    : "";
  var bonusIds = Array.isArray(shiftPayload && shiftPayload.bonusIds)
    ? shiftPayload.bonusIds
    : [];

  if (!employeeId) {
    errors.push({ code: "NO_EMPLOYEE_ID", message: "יש לבחור עובד לשמירה." });
  }
  if (!jobId && !jobName) {
    errors.push({ code: "NO_JOB", message: "יש לבחור סוג עבודה." });
  }
  if (!shiftDate) {
    errors.push({ code: "NO_DATE", message: "יש לבחור תאריך למשמרת." });
  }

  if (!shiftId && (!startTime || !parsedStart)) {
    errors.push({
      code: "NO_START_TIME",
      message: "חסר זמן כניסה למשמרת חדשה.",
    });
  }
  if (!shiftId && (!endTime || !parsedEnd)) {
    errors.push({ code: "NO_END_TIME", message: "חסר זמן יציאה למשמרת חדשה." });
  }

  var parsedStart =
    shiftDate && startTime ? buildDateTime_(shiftDate, startTime) : null;
  var parsedEnd =
    shiftDate && endTime ? buildDateTime_(shiftDate, endTime) : null;

  if (parsedStart && parsedEnd && parsedEnd.getTime() < parsedStart.getTime()) {
    parsedEnd.setDate(parsedEnd.getDate() + 1);
  }

  var saveResult = null;

  if (!errors.length && shiftId) {
    var updatePayload = {
      shiftId: shiftId,
      workType: jobName,
      workTypeId: jobId,
      department: department,
      status: status,
      date: shiftDate,
      startTime: startTime,
      endTime: endTime,
      manualNote: manualNote,
      bonusIds: bonusIds,
    };

    try {
      if (typeof SHIFTS_updateShift === "function") {
        saveResult = SHIFTS_updateShift(updatePayload);
      } else {
        saveResult = { ok: false, error: "SHIFTS_updateShift missing" };
      }
    } catch (errUpdate) {
      saveResult = { ok: false, error: String(errUpdate) };
    }

    if (!saveResult || saveResult.ok !== true) {
      errors.push({
        code: "UPDATE_FAILED",
        message:
          "עדכון המשמרת נכשל" +
          (saveResult && saveResult.error ? ": " + saveResult.error : ""),
      });
    }
  } else if (!errors.length && !shiftId) {
    var newShiftId =
      (typeof Utilities !== "undefined" &&
        Utilities.getUuid &&
        Utilities.getUuid()) ||
      "";
    var startIso = parsedStart ? formatIsoDate_(parsedStart) : shiftDate;
    var endIso = parsedEnd ? formatIsoDate_(parsedEnd) : shiftDate;
    var endTimeStr = parsedEnd ? timeStr_(parsedEnd) : endTime;

    function submitReport_(direction, fixDateVal, fixTimeVal) {
      if (typeof handleShiftReportSubmit_ !== "function") {
        return { ok: false, error: "handleShiftReportSubmit_ missing" };
      }
      try {
        return handleShiftReportSubmit_({
          shiftId: newShiftId,
          employeeId: employeeId,
          jobId: jobId,
          jobName: jobName,
          department: department,
          direction: direction,
          fixDate: fixDateVal,
          fixTime: fixTimeVal,
          workDate: shiftDate,
          manualDate: fixDateVal,
          manualTime: fixTimeVal,
          mode: "manual",
          requiresApproval: false,
        });
      } catch (errSubmit) {
        return { ok: false, error: String(errSubmit) };
      }
    }

    var startRes = submitReport_("כניסה", startIso, startTime);
    var endRes = submitReport_("יציאה", endIso, endTimeStr || endTime);

    if (!startRes || startRes.ok === false || startRes.success === false) {
      errors.push({
        code: "CREATE_FAILED",
        message:
          "שמירת שעת כניסה נכשלה" +
          (startRes && startRes.error ? ": " + startRes.error : ""),
      });
    }
    if (!endRes || endRes.ok === false || endRes.success === false) {
      errors.push({
        code: "CREATE_FAILED",
        message:
          "שמירת שעת יציאה נכשלה" +
          (endRes && endRes.error ? ": " + endRes.error : ""),
      });
    }

    if (!errors.length) {
      try {
        if (typeof SHIFTS_upsertAroundWorkLog_ === "function") {
          SHIFTS_upsertAroundWorkLog_(employeeId, jobId, shiftDate);
        }
      } catch (_upsertErr) {
        // best-effort refresh only
      }

      saveResult = { ok: true, shiftId: newShiftId };
    }
  }

  var dialogRes = null;
  if (!errors.length) {
    try {
      dialogRes = UI_getShiftsDialogData(filtersEcho);
    } catch (errReload) {
      errors.push({
        code: "RELOAD_FAILED",
        message: "טעינת הנתונים לאחר שמירה נכשלה: " + errReload,
      });
    }

    if (
      dialogRes &&
      Array.isArray(dialogRes.errors) &&
      dialogRes.errors.length
    ) {
      errors = errors.concat(dialogRes.errors);
    }
  }

  if (!errors.length && typeof AUDIT_logEvent === "function") {
    try {
      AUDIT_logEvent({
        eventType: shiftId ? "shift_dialog_update" : "shift_dialog_create",
        entityType: "shift",
        entityId: shiftId || (saveResult && saveResult.shiftId) || "",
        actorEmail: actorEmail,
        actorRole: actorRole,
        traceId: traceId,
        summaryAfter: status || "",
        extra: {
          employeeId: employeeId,
          jobId: jobId,
          hasManualNote: !!manualNote,
        },
      });
    } catch (_auditErr) {
      // best-effort logging only
    }
  }

  var result = {
    meta: {
      traceId: traceId,
      requestedAt: requestedAt,
      actorEmail: actorEmail,
      actorRole: actorRole,
      canEdit: errors.length === 0,
    },
    data:
      saveResult || dialogRes
        ? {
            shiftId: shiftId,
            saveResult: saveResult,
            dialogData: dialogRes ? dialogRes.data : null,
            dialogMeta: dialogRes ? dialogRes.meta : null,
          }
        : null,
    errors: errors,
  };

  logDuration_(logger, "sidebar.shifts.save", startMs, {
    shiftId: shiftId,
    errors: errors.length,
    source: "extended-dialog",
  });

  return result;
}

/** Opens the extended requests dialog. */
function UI_openRequestsDialog(params) {
  var safeParams = params && typeof params === "object" ? params : {};

  var template = HtmlService.createTemplateFromFile("ExtendedDialog");
  template.mode = "requests";
  template.title = "בקשות עובדים";
  template.paramsJson = JSON.stringify(safeParams || {});
  var html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, "בקשות");
}

/** Returns data for the extended requests dialog (placeholder). */
// eslint-disable-next-line no-unused-vars
function UI_getRequestsDialogData(params) {
  var logger = getSidebarLogger_("SIDEBAR_REQUESTS_DIALOG");
  var startMs = new Date().getTime();
  var traceId =
    (typeof Utilities !== "undefined" &&
      Utilities.getUuid &&
      Utilities.getUuid()) ||
    "";
  var requestedAt = new Date();

  var actorCtx = resolveActorContext_();
  var actorEmail = actorCtx.actorEmail;
  var actorRole = actorCtx.actorRole;

  var errors = [];

  var safeParams = params && typeof params === "object" ? params : {};
  var includeCancelled = !!safeParams.includeCancelled;
  var employeeId = safeParams.employeeId
    ? String(safeParams.employeeId).trim()
    : "";
  var jobId = safeParams.jobId ? String(safeParams.jobId).trim() : "";
  var requestTypeFilter = safeParams.requestType
    ? String(safeParams.requestType).trim()
    : "";
  var statusFilter = safeParams.status ? String(safeParams.status).trim() : "";
  var statusesFilterArray = Array.isArray(safeParams.statuses)
    ? safeParams.statuses
    : null;

  var fromDateStr = safeParams.fromDate || "";
  var toDateStr = safeParams.toDate || "";

  var range = null;
  var fromIso = "";
  var toIso = "";

  // מפעילים נירמול טווח תאריכים רק אם המשתמש הגדיר from/to מפורשות
  if (fromDateStr || toDateStr) {
    range = normalizeDateRange_(fromDateStr, toDateStr);
    if (!range.ok) {
      errors.push(range.error);
    } else {
      fromIso = range.fromIso;
      toIso = range.toIso;
    }
  }

  var rawRequests = [];
  var statusesCollected = [];
  var cancelledStatusValue = "";

  if (errors.length === 0) {
    var filters = {
      includeClosed: true,
    };

    if (fromIso) {
      filters.dateFrom = fromIso;
    }
    if (toIso) {
      filters.dateTo = toIso;
    }

    if (statusesFilterArray && statusesFilterArray.length) {
      filters.statuses = statusesFilterArray
        .filter(function (s) {
          return !!normRequestVal_(s);
        })
        .map(function (s) {
          return String(s).trim();
        });
    } else if (statusFilter) {
      filters.statuses = [statusFilter];
    }

    var reqRes = null;
    try {
      if (
        typeof REQ !== "undefined" &&
        typeof REQ.listRequests === "function"
      ) {
        reqRes = REQ.listRequests(filters);
      } else if (typeof REQ_listRequests === "function") {
        reqRes = REQ_listRequests(filters);
      }
    } catch (errReq) {
      reqRes = { ok: false, error: String(errReq) };
    }

    if (!reqRes || !reqRes.ok) {
      errors.push({
        code: "REQ_LIST_FAILED",
        message:
          "קריאת בקשות נכשלה" +
          (reqRes && reqRes.error ? ": " + reqRes.error : ""),
      });
    } else {
      rawRequests = reqRes.requests || [];
      statusesCollected = reqRes.statuses || [];
      cancelledStatusValue = detectCancelledStatusValue_(statusesCollected);
    }
  }

  if (!cancelledStatusValue) cancelledStatusValue = "מבוטלת";
  var cancelledKey = normRequestVal_(cancelledStatusValue);
  var requestedCancelled =
    cancelledKey && normRequestVal_(statusFilter) === cancelledKey;

  var filteredRequests = rawRequests.filter(function (rec) {
    if (
      employeeId &&
      normRequestVal_(rec.employeeId) !== normRequestVal_(employeeId)
    ) {
      return false;
    }
    if (jobId && normRequestVal_(rec.jobId) !== normRequestVal_(jobId)) {
      return false;
    }

    var requestTypeVal = deriveRequestType_(rec);
    if (
      requestTypeFilter &&
      normRequestVal_(requestTypeVal) !== normRequestVal_(requestTypeFilter)
    ) {
      return false;
    }

    if (
      !requestedCancelled &&
      !includeCancelled &&
      isCancelledStatus_(rec.status, cancelledStatusValue)
    ) {
      return false;
    }

    if (
      statusFilter &&
      normRequestVal_(rec.status) !== normRequestVal_(statusFilter)
    ) {
      return false;
    }

    return true;
  });

  var mappedRequests = filteredRequests.map(function (rec) {
    var statusVal = rec.status || "";
    var requestTypeVal = deriveRequestType_(rec);
    return {
      id: rec.requestId || "",
      employeeId: rec.employeeId || "",
      employeeName: rec.employeeName || "",
      jobId: rec.jobId || "",
      jobName: rec.jobName || "",
      requestType: requestTypeVal,
      status: statusVal,
      createdAt: rec.timestamp || rec.timestampIso || "",
      fromDate: rec.fixDate || "",
      toDate: rec.fixDate || "",
      notes: rec.note || "",
      managerNotes: rec.shiftNote || "",
      canEmployeeEdit: canEmployeeEditRequestStatus_(
        statusVal,
        cancelledStatusValue
      ),
      canManagerEdit: true,
    };
  });

  var employeesLookup = [];
  try {
    if (typeof EMP_getSidebarBootstrap === "function") {
      var bootstrap = EMP_getSidebarBootstrap();
      if (
        bootstrap &&
        bootstrap.employees &&
        Array.isArray(bootstrap.employees)
      ) {
        employeesLookup = bootstrap.employees.map(function (e) {
          return {
            id: e.id || "",
            name: e.name || "",
            email: e.email || "",
          };
        });
      }
    }
  } catch (_empErr) {
    errors.push({
      code: "EMP_LOOKUP_FAILED",
      message: "טעינת רשימת העובדים נכשלה.",
    });
  }

  actorRole = deriveActorRole_(actorEmail, employeesLookup);

  var requestTypes = buildUniqueValuePairs_(mappedRequests, "requestType");
  var statuses = buildUniqueValuePairs_(mappedRequests, "status");

  var data = {
    filtersEcho: {
      fromDate: fromIso,
      toDate: toIso,
      employeeId: employeeId || "",
      jobId: jobId || "",
      requestType: requestTypeFilter || "",
      status: statusFilter || "",
      includeCancelled: includeCancelled,
    },
    requests: mappedRequests,
    employeesLookup: employeesLookup,
    requestTypes: requestTypes,
    statuses: statuses,
    rules: {
      cancelledStatusValue: cancelledStatusValue,
      defaultHideCancelled: true,
    },
  };

  var result = {
    meta: {
      traceId: traceId,
      requestedAt: requestedAt,
      actorEmail: actorEmail,
      actorRole: actorRole,
      canEdit:
        errors.length === 0 &&
        actorRole !== "UNKNOWN" &&
        actorRole !== "EMPLOYEE",
    },
    data: errors.length ? null : data,
    errors: errors,
  };

  logDuration_(logger, "sidebar.requests.get", startMs, {
    errors: errors.length,
    returned: rawRequests.length,
    source: "extended-dialog",
  });

  return result;
}

/** Saves changes made in the extended requests dialog (placeholder). */
// eslint-disable-next-line no-unused-vars
function UI_saveRequestChanges(envelope) {
  var logger = getSidebarLogger_("SIDEBAR_REQUESTS_SAVE");
  var startMs = new Date().getTime();
  var traceId =
    (typeof Utilities !== "undefined" &&
      Utilities.getUuid &&
      Utilities.getUuid()) ||
    "";
  var requestedAt = new Date();

  var actorCtx = resolveActorContext_();
  var actorEmail = actorCtx.actorEmail;
  var actorRole = actorCtx.actorRole;
  try {
    var employeesLookupForActor = [];
    if (typeof EMP_listBasicDirectory_ === "function") {
      employeesLookupForActor = EMP_listBasicDirectory_() || [];
    }
    actorRole = deriveActorRole_(actorEmail, employeesLookupForActor);
  } catch (_actorRoleErr) {
    actorRole = deriveActorRole_(actorEmail, []);
  }

  var errors = [];
  var safeEnvelope = envelope && typeof envelope === "object" ? envelope : {};
  var candidate =
    safeEnvelope.payload !== undefined
      ? safeEnvelope.payload
      : safeEnvelope.request !== undefined
      ? safeEnvelope.request
      : safeEnvelope.requests !== undefined
      ? safeEnvelope.requests
      : safeEnvelope.data !== undefined
      ? safeEnvelope.data
      : safeEnvelope;

  var requestsList = [];
  if (Array.isArray(candidate)) {
    requestsList = candidate;
  } else if (candidate && typeof candidate === "object") {
    requestsList = [candidate];
  }

  if (!requestsList.length) {
    errors.push({
      code: "NO_REQUESTS",
      message: "לא התקבלו בקשות לשמירה.",
    });
  }

  if (actorRole === "UNKNOWN" || actorRole === "EMPLOYEE") {
    errors.push({
      code: "FORBIDDEN",
      message: "אין הרשאה לעדכן בקשות.",
    });
  }

  var cancelledStatusValue =
    safeEnvelope.rules && safeEnvelope.rules.cancelledStatusValue
      ? String(safeEnvelope.rules.cancelledStatusValue)
      : "";
  if (!cancelledStatusValue) cancelledStatusValue = "מבוטלת";

  function normalizeRequestForSave_(req) {
    if (!req || typeof req !== "object") {
      return {
        ok: false,
        error: { code: "INVALID_PAYLOAD", message: "נתוני הבקשה אינם תקינים." },
      };
    }

    function norm_(v) {
      if (v === null || v === undefined) return "";
      return String(v).replace(/\s+/g, " ").trim();
    }

    var requestId = norm_(req.requestId || req.id);
    if (!requestId) {
      return {
        ok: false,
        error: { code: "NO_REQUEST_ID", message: "חסר מזהה בקשה לשמירה." },
      };
    }

    var status = norm_(req.status || req.newStatus || req.state);
    if (!status) {
      return {
        ok: false,
        error: { code: "NO_STATUS", message: "חסר סטטוס בקשה לשמירה." },
      };
    }

    var payload = { requestId: requestId, status: status };

    function maybeSet_(key, value) {
      if (value !== undefined) payload[key] = value;
    }

    var timestamp =
      req.timestamp ||
      req.createdAt ||
      req.requestedAt ||
      req.updatedAt ||
      null;
    if (!timestamp) timestamp = new Date().toISOString();
    maybeSet_("timestamp", timestamp);

    maybeSet_("employeeId", norm_(req.employeeId));
    maybeSet_("employeeName", norm_(req.employeeName));
    maybeSet_("shiftId", norm_(req.shiftId));
    maybeSet_("direction", norm_(req.direction || req.requestType));
    var fixDateVal = norm_(
      req.fixDate || req.fromDate || req.toDate || req.workDate
    );
    if (fixDateVal) maybeSet_("fixDate", fixDateVal);
    maybeSet_("fixTime", norm_(req.fixTime || req.time));
    maybeSet_("jobId", norm_(req.jobId));
    maybeSet_("jobName", norm_(req.jobName || req.jobType));
    maybeSet_("department", norm_(req.department));
    if (req.units !== undefined) maybeSet_("units", req.units);
    var noteVal =
      req.managerNotes !== undefined
        ? req.managerNotes
        : req.notes !== undefined
        ? req.notes
        : req.shiftNote;
    if (noteVal !== undefined) {
      maybeSet_("shiftNote", noteVal);
    }
    maybeSet_("decidedByName", norm_(req.decidedByName || actorEmail));
    maybeSet_("decidedByRole", norm_(req.decidedByRole || actorRole));
    if (safeEnvelope.meta && safeEnvelope.meta.actorId) {
      maybeSet_("decidedById", norm_(safeEnvelope.meta.actorId));
    }

    return { ok: true, payload: payload, requestId: requestId, status: status };
  }

  function saveRequest_(payload) {
    if (typeof handleRequestPost !== "function") {
      return {
        ok: false,
        error: {
          code: "REQUEST_SAVE_MISSING",
          message: "handleRequestPost לא זמין בסקריפט.",
        },
      };
    }

    var rawRes;
    try {
      rawRes = handleRequestPost({
        postData: { contents: JSON.stringify(payload) },
      });
    } catch (errSave) {
      return {
        ok: false,
        error: {
          code: "REQUEST_SAVE_THROW",
          message: "שמירת הבקשה נכשלה: " + errSave,
        },
      };
    }

    var parsed = null;
    try {
      if (rawRes && typeof rawRes.getContent === "function") {
        parsed = JSON.parse(rawRes.getContent());
      } else if (rawRes && typeof rawRes.getContentText === "function") {
        parsed = JSON.parse(rawRes.getContentText());
      } else if (typeof rawRes === "string") {
        parsed = JSON.parse(rawRes);
      } else if (
        rawRes &&
        typeof rawRes === "object" &&
        rawRes.success !== undefined
      ) {
        parsed = rawRes;
      }
    } catch (parseErr) {
      return {
        ok: false,
        error: {
          code: "REQUEST_SAVE_PARSE_FAILED",
          message: "שמירת הבקשה נכשלה: " + parseErr,
        },
      };
    }

    if (!parsed || parsed.success !== true) {
      return {
        ok: false,
        error: {
          code: "REQUEST_SAVE_FAILED",
          message:
            "שמירת הבקשה נכשלה" +
            (parsed && parsed.error ? ": " + parsed.error : ""),
          details: parsed || rawRes || null,
        },
      };
    }

    return {
      ok: true,
      summary: {
        requestId: parsed.requestId || payload.requestId || "",
        mode: parsed.mode || "updated",
        rowIndex: parsed.rowIndex || null,
        status: payload.status || "",
      },
    };
  }

  var saveResults = [];

  if (!errors.length) {
    for (var i = 0; i < requestsList.length; i++) {
      var normRes = normalizeRequestForSave_(requestsList[i]);
      if (!normRes.ok) {
        errors.push(normRes.error);
        continue;
      }

      var saveRes = saveRequest_(normRes.payload);
      if (!saveRes.ok) {
        errors.push(saveRes.error);
        continue;
      }

      saveResults.push(saveRes.summary);

      if (typeof AUDIT_logEvent === "function") {
        try {
          AUDIT_logEvent({
            traceId: traceId,
            eventType: "REQUEST_UPDATED",
            entityType: "request",
            entityId: saveRes.summary.requestId,
            actorEmail: actorEmail,
            actorRole: actorRole,
            source: "requests_dialog",
            summary: normRes.status,
            extra: {
              status: normRes.status,
              requestId: normRes.requestId,
              shiftId: normRes.payload.shiftId || "",
              employeeId: normRes.payload.employeeId || "",
              cancelledStatusValue: cancelledStatusValue,
            },
          });
        } catch (_auditErr) {
          // best-effort audit
        }
      }
    }
  }

  logDuration_(logger, "sidebar.requests.save", startMs, {
    errors: errors.length,
    saved: saveResults.length,
    source: "extended-dialog",
  });

  return {
    meta: {
      traceId: traceId,
      requestedAt: requestedAt,
      actorEmail: actorEmail,
      actorRole: actorRole,
      canEdit:
        errors.length === 0 &&
        actorRole !== "UNKNOWN" &&
        actorRole !== "EMPLOYEE",
    },
    data: saveResults.length ? { saveResults: saveResults } : null,
    errors: errors,
  };
}

// Debug helpers for manual testing (not for production use).
// eslint-disable-next-line no-unused-vars
function UI_debugOpenEmployeeDialog() {
  UI_openEmployeeDialog({});
}

// eslint-disable-next-line no-unused-vars
function UI_debugOpenShiftsDialog() {
  UI_openShiftsDialog({});
}

// eslint-disable-next-line no-unused-vars
function UI_debugOpenRequestsDialog() {
  UI_openRequestsDialog({});
}
