/* global HtmlService, SpreadsheetApp, Session, Utilities, SHIFTS_list, EMP_getSidebarBootstrap */

/** Opens the extended employee dialog. */
function UI_openEmployeeDialog(params) {
  params = params || {};
  void params;

  var template = HtmlService.createTemplateFromFile("ExtendedDialog");
  template.mode = "employee";
  template.title = "כרטיס עובד מורחב";
  var html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, "כרטיס עובד");
}

/** Returns data for the extended employee dialog. */
function UI_getEmployeeDialogData(params) {
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

  return {
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
function UI_saveEmployeeDialogChanges(envelope) {
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

  return {
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
}

/** Opens the extended shifts dialog. */
function UI_openShiftsDialog(params) {
  params = params || {};
  void params;

  var template = HtmlService.createTemplateFromFile("ExtendedDialog");
  template.mode = "shifts";
  template.title = "משמרות מורחב";
  var html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, "משמרות");
}

/** Returns data for the extended shifts dialog. */
function UI_getShiftsDialogData(params) {
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
        rules: { crossMidnightCutoffHour: 4 },
      };

  return {
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

/** Saves changes made in the extended shifts dialog (placeholder). */
function UI_saveShiftChanges(envelope) {
  void envelope;
  return {
    meta: {
      traceId: "",
      requestedAt: new Date(),
      actorEmail: "",
      actorRole: "",
      canEdit: false,
    },
    data: null,
    errors: [],
  };
}

/** Opens the extended requests dialog. */
function UI_openRequestsDialog(params) {
  params = params || {};
  void params;

  var template = HtmlService.createTemplateFromFile("ExtendedDialog");
  template.mode = "requests";
  template.title = "בקשות עובדים";
  var html = template.evaluate().setWidth(1200).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, "בקשות");
}

/** Returns data for the extended requests dialog (placeholder). */
function UI_getRequestsDialogData(params) {
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

  var range = normalizeDateRange_(safeParams.fromDate, safeParams.toDate);
  if (!range.ok) {
    errors.push(range.error);
  }

  var fromIso = range.ok ? range.fromIso : "";
  var toIso = range.ok ? range.toIso : "";

  var rawRequests = [];
  var statusesCollected = [];
  var cancelledStatusValue = "";

  if (errors.length === 0) {
    var filters = {
      dateFrom: fromIso,
      dateTo: toIso,
      includeClosed: true, // we will handle cancel hiding ourselves
    };
    if (statusFilter) filters.statuses = [statusFilter];

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
      managerNotes: "",
      canEmployeeEdit: canEmployeeEditRequestStatus_(
        statusVal,
        cancelledStatusValue
      ),
      canManagerEdit: true, // placeholder until role logic is added
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

  return {
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
}

/** Saves changes made in the extended requests dialog (placeholder). */
function UI_saveRequestChanges(envelope) {
  void envelope;
  return {
    meta: {
      traceId: "",
      requestedAt: new Date(),
      actorEmail: "",
      actorRole: "",
      canEdit: false,
    },
    data: null,
    errors: [],
  };
}

// Debug helpers for manual testing (not for production use).
function UI_debugOpenEmployeeDialog() {
  UI_openEmployeeDialog({});
}

function UI_debugOpenShiftsDialog() {
  UI_openShiftsDialog({});
}

function UI_debugOpenRequestsDialog() {
  UI_openRequestsDialog({});
}
