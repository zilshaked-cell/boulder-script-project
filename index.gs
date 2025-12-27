// Apps Script entrypoint for boulder-shifts-app. Paste into your Apps Script project.
// All requests arrive as POST JSON: { action: string, payload?: any }.

const EMPLOYEES_SHEET_NAME = "פרטי עובדים";
const EMPLOYEES_SHEET_NAMES = [
  EMPLOYEES_SHEET_NAME,
  "עובדים",
  "Employees",
  "employees",
];
const WORK_LOGS_SHEET_NAME = "דיווח שעות עבודה";
const REQUESTS_SHEET_NAMES = ["בקשות עובדים"];
const OPTIONS_SHEET_NAMES = ["אופציות בחירה ו ID'S", "אופציות בחירה ו ID_S"];

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

const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();
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

const HEADER_KEY_MAP = {
  "ID משמרת": "shiftId",
  "ID דיווח": "shiftId",
  "Shift ID": "shiftId",
  "Request ID": "id",
  requestId: "id",
  "Employee ID": "employeeId",
  "ID עובד": "employeeId",
  "מזהה עובד": "employeeId",
  "ID אופן תשלום": "payTypeId",
  "ID אופני תשלום": "payTypeId",
  "אופן תשלום": "payType",
  "אופני תשלום": "payType",
  "ת.ז": "nationalId",
  "שם מלא": "employeeName",
  "ID סוג עבודה": "jobTypeId",
  "ID סוגי עבודה": "jobTypeId",
  "סוג עבודה": "jobName",
  "סוגי עבודה": "jobName",
  סטטוס: "status",
  "סטטוס סוגי עבודה": "jobStatus",
  מחלקה: "department",
  מחלקות: "department",
  "דיווח שעות": "direction",
  "כניסה / יציאה": "direction",
  "תיקון תאריך": "fixDate",
  "תיקון שעה": "fixTime",
  "כמות היחידות": "units",
  "דיווח יחידות": "units",
  הערות: "note",
  "הערה למנהל": "noteToManager",
  "תאריך משמרת": "workDate",
  "תיקון תאריך": "workDate",
  "חותמת זמן": "timestamp",
  "תאריך נשלח": "submittedAt",
  "תאריך החלטה": "decidedAt",
  "מנהל מחליט": "managerDecision",
  תז: "nationalId",
  "סוג בקשה": "requestType",
  "סטטוס בקשה": "status",
  "הערות למשמרת": "note",
};

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action =
    params.action || (params.ping ? "ping" : params.shifts ? "shifts" : "");
  try {
    switch (action) {
      case "ping":
        return jsonResponse(
          withOk_({ pong: true, timestamp: new Date().toISOString() })
        );
      case "shifts": {
        const employeeId = params.employeeId || params.empId || params.id || "";
        if (!employeeId)
          return jsonResponse({ ok: false, error: "Missing employeeId" }, 400);
        return jsonResponse(withOk_(listWorkLogsByEmployee_({ employeeId })));
      }
      case "employeeExistsByEmail": {
        const result = employeeExistsByEmail_(params);
        return jsonResponse(
          Object.assign({}, result, { success: result.ok === true })
        );
      }
      default:
        return jsonResponse({ ok: false, error: "Unknown action" }, 400);
    }
  } catch (err) {
    return jsonResponse(
      {
        ok: false,
        error: err && err.message ? err.message : "Unexpected error",
        stack: err && err.stack ? String(err.stack).slice(0, 500) : undefined,
      },
      500
    );
  }
}

function doPost(e) {
  try {
    const parsed = parseBody(e);
    if (parsed.error) {
      return jsonResponse(
        withOk_({
          ok: false,
          error: "Invalid JSON body: " + String(parsed.error),
        }),
        400
      );
    }

    const body = parsed.body || {};
    const action = stringValue(body.action);
    const payload = body.payload || {};

    if (!action) {
      return handleLegacyPost_(e, body);
    }

    return handleNewPost_(action, payload);
  } catch (err) {
    return jsonResponse(
      {
        ok: false,
        error: err && err.message ? err.message : "Unexpected error",
        stack: err && err.stack ? String(err.stack).slice(0, 500) : undefined,
      },
      500
    );
  }
}

function handleNewPost_(action, payload) {
  switch (action) {
    case "jobTypes.list":
      return jsonResponse(withOk_({ jobTypes: listJobTypes_() }));
    case "employee.linkedJobs":
      return jsonResponse(withOk_(listEmployeeLinkedJobIds_(payload)));
    case "workLogs.listByEmployee":
      return jsonResponse(withOk_(listWorkLogsByEmployee_(payload)));
    case "requests.listByEmployee":
      return jsonResponse(withOk_(listRequestsByEmployee_(payload)));
    case "reportAccessIssue":
      return jsonResponse(withOk_(reportAccessIssue_(payload || {})));
    case "shiftReport.submit":
      return jsonResponse(withOk_(handleShiftReportSubmit_(payload)));
    case "shiftCorrection.submit":
      return jsonResponse(withOk_(handleShiftCorrectionSubmit_(payload)));
    case "shiftReport.monthlyErrorNotify":
      return jsonResponse(withOk_(handleShiftReportMonthlyErrorNotify_(payload)));
    case "requests.approve":
      return jsonResponse(withOk_(handleRequestApprove_(payload)));
    case "employee.save":
      return jsonResponse(withOk_(handleEmployeeSave_(payload)));
    case "employeeExistsByEmail":
      return jsonResponse(withOk_(employeeExistsByEmail_(payload || {})));
    // Legacy stubs so existing screens stay alive until you wire real data
    case "getCurrentEmployee": {
      const email =
        (payload && payload.email) ||
        (payload && payload.user && payload.user.email) ||
        "";

      if (!email) {
        return jsonResponse({ ok: false, error: "Missing email" }, 400);
      }

      const result = employeeExistsByEmail_({ email: email });
      if (result && result.ok === true && result.exists === true) {
        return jsonResponse(
          withOk_({
            employee: {
              id: result.employeeId || "",
              name: result.fullName || result.name || "",
              email: email,
            },
          })
        );
      }

      return jsonResponse(withOk_({ employee: null }));
    }
    case "getWorkLogs":
      return jsonResponse(withOk_({ workLogs: [] }));
    case "getEmployeeRequests":
      return jsonResponse(withOk_({ requests: [] }));
    default:
      return jsonResponse({ ok: false, error: "Unknown action" }, 400);
  }
}

function handleLegacyPost_(e, body) {
  const paramAction = stringValue(
    e && e.parameter && e.parameter.action ? e.parameter.action : ""
  );

  if (paramAction === "request" || paramAction === "requests") {
    // Keep the legacy handler intact for old callers; it returns a JSON response.
    return handleRequestPost(e);
  }

  const legacyPayload = normalizeLegacyShiftPayload_(body);
  return jsonResponse(withOk_(handleShiftReportSubmit_(legacyPayload)));
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
      src.jobName || (src.job && src.job.name ? src.job.name : "")
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
  const output = ContentService.createTextOutput(JSON.stringify(data || {}));
  output.setMimeType(ContentService.MimeType.JSON);
  if (typeof status === "number") {
    output.setResponseCode(status);
  }
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
  const sheet = getSheetByPossibleNames_(OPTIONS_SHEET_NAMES);
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const statusCol = getRequiredColumn_(
    headerMap,
    ["סטטוס סוגי עבודה"],
    sheetName
  );
  const idCol = getRequiredColumn_(
    headerMap,
    ["ID סוגי עבודה", "ID סוג עבודה"],
    sheetName
  );
  const nameCol = getRequiredColumn_(
    headerMap,
    ["סוגי עבודה", "סוג עבודה"],
    sheetName
  );
  const deptCol = getRequiredColumn_(headerMap, ["מחלקות", "מחלקה"], sheetName);
  const payStatusCol = getOptionalColumn_(headerMap, ["סטטוס אופן תשלום"]);
  const payIdCol = getOptionalColumn_(headerMap, ["ID אופני תשלום", "ID אופן תשלום"]);
  const payNameCol = getOptionalColumn_(headerMap, ["אופני תשלום", "אופן תשלום"]);
  const dataRange = sheet.getDataRange().getValues();
  let blanks = 0;
  const results = [];
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const name = stringValue(row[nameCol - 1]);
    const status = stringValue(row[statusCol - 1]);
    const payStatus = payStatusCol ? stringValue(row[payStatusCol - 1]) : "";
    const payTypeId = payIdCol ? stringValue(row[payIdCol - 1]) : "";
    const payTypeName = payNameCol ? stringValue(row[payNameCol - 1]) : "";
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

function listEmployeeLinkedJobIds_(payload) {
  const employeeId =
    payload && payload.employeeId ? String(payload.employeeId).trim() : "";
  if (!employeeId) throw new Error("Missing employeeId");
  const sheet = getEmployeesSheet_();
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const idCol = getRequiredColumn_(
    headerMap,
    ["מזהה עובד", "ID עובד"],
    sheetName
  );
  const jobTypeCol = getRequiredColumn_(
    headerMap,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheetName
  );
  const statusCol = getRequiredColumn_(headerMap, ["סטטוס"], sheetName);
  const values = sheet.getDataRange().getValues();
  const jobTypeIds = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmpId = stringValue(row[idCol - 1]);
    if (rowEmpId !== employeeId) continue;
    const status = stringValue(row[statusCol - 1]);
    if (status === "לא פעיל") continue;
    const jobTypeId = stringValue(row[jobTypeCol - 1]);
    if (jobTypeId) jobTypeIds.push(jobTypeId);
  }
  return { employeeId, jobTypeIds };
}

function listWorkLogsByEmployee_(payload) {
  const employeeId =
    payload && payload.employeeId ? String(payload.employeeId).trim() : "";
  if (!employeeId) throw new Error("Missing employeeId");
  const sheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const idCol = getRequiredColumn_(
    headerMap,
    ["ID משמרת", "ID דיווח"],
    sheetName
  );
  const tsCol = getRequiredColumn_(headerMap, ["חותמת זמן"], sheetName);
  const empIdCol = getRequiredColumn_(
    headerMap,
    ["ID עובד", "מזהה עובד"],
    sheetName
  );
  const empNameCol = getRequiredColumn_(headerMap, ["שם מלא"], sheetName);
  const dirCol = getRequiredColumn_(
    headerMap,
    ["כניסה / יציאה", "דיווח שעות"],
    sheetName
  );
  const fixDateCol = getRequiredColumn_(headerMap, ["תיקון תאריך"], sheetName);
  const fixTimeCol = getRequiredColumn_(headerMap, ["תיקון שעה"], sheetName);
  const jobTypeIdCol = getRequiredColumn_(
    headerMap,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheetName
  );
  const jobNameCol = getRequiredColumn_(
    headerMap,
    ["סוג עבודה", "סוגי עבודה"],
    sheetName
  );
  const deptCol = getRequiredColumn_(headerMap, ["מחלקה", "מחלקות"], sheetName);
  const unitsCol = getRequiredColumn_(
    headerMap,
    ["כמות היחידות", "דיווח יחידות"],
    sheetName
  );
  const notesCol = getRequiredColumn_(headerMap, ["הערות", "הערה"], sheetName);
  const values = sheet.getDataRange().getValues();
  const logs = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (stringValue(row[empIdCol - 1]) !== employeeId) continue;
    logs.push({
      shiftId: stringValue(row[idCol - 1]),
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
  return { workLogs: logs };
}

function listRequestsByEmployee_(payload) {
  const employeeId =
    payload && payload.employeeId ? String(payload.employeeId).trim() : "";
  if (!employeeId) throw new Error("Missing employeeId");
  const employeesSheet = getEmployeesSheet_();
  const empHeaders = getHeaderMap_(employeesSheet);
  const empSheetName = employeesSheet.getName();
  const empIdCol = getRequiredColumn_(
    empHeaders,
    ["מזהה עובד", "ID עובד"],
    empSheetName
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
  if (!tzValue) return { requests: [] };

  const reqSheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
  const reqHeaders = getHeaderMap_(reqSheet);
  const reqSheetName = reqSheet.getName();
  const reqEmpCol = getRequiredColumn_(
    reqHeaders,
    ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
    reqSheetName
  );
  const statusCol = getRequiredColumn_(
    reqHeaders,
    ["סטטוס", "סטטוס בקשה"],
    reqSheetName
  );
  const jobNameCol = getRequiredColumn_(
    reqHeaders,
    ["סוג עבודה", "סוגי עבודה"],
    reqSheetName
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
    reqSheetName
  );
  const fixDateCol = getOptionalColumn_(reqHeaders, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(reqHeaders, ["תיקון שעה"]);
  const payTypeCol = getOptionalColumn_(reqHeaders, ["אופן תשלום", "אופני תשלום"]);
  const payTypeIdCol = getOptionalColumn_(reqHeaders, ["ID אופן תשלום", "ID אופני תשלום"]);

  const jobTypes = listJobTypes_();
  const jobTypeNames = jobTypes.map((j) => j.name);

  const values = reqSheet.getDataRange().getValues();
  const requests = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowEmpValue = stringValue(row[reqEmpCol - 1]);
    if (rowEmpValue !== tzValue && rowEmpValue !== employeeId) continue;
    const jobName = stringValue(row[jobNameCol - 1]);
    let requestType = requestTypeCol
      ? stringValue(row[requestTypeCol - 1])
      : "";
    if (!requestType) {
      requestType =
        jobTypeNames.indexOf(jobName) >= 0 ? "job_not_linked" : "new_job_type";
    }
    const submittedAt = submittedCol ? stringValue(row[submittedCol - 1]) : "";
    const requestedSummary = stringValue(row[directionCol - 1]) ||
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

    requests.push({
      id: stringValue(row[shiftIdCol - 1]),
      status: stringValue(row[statusCol - 1]),
      jobName: jobName,
      workDate: workDate,
      requestedSummary: requestedSummary,
      units: unitsCol ? stringValue(row[unitsCol - 1]) : "",
      noteToManager: noteToManager,
      submittedAt: submittedAt,
      decidedAt: decidedCol ? stringValue(row[decidedCol - 1]) : "",
      type: requestType,
      employeeId: employeeId,
      employeeName: fullName,
      shiftId: stringValue(row[shiftIdCol - 1]),
      jobTypeId: stringValue(row[jobTypeIdCol - 1]),
      fixDate: fixDateCol ? stringValue(row[fixDateCol - 1]) : "",
      fixTime: fixTimeCol ? stringValue(row[fixTimeCol - 1]) : "",
      payType: payType,
      payTypeLabel: rawPayType,
      payTypeId: payTypeIdCol ? stringValue(row[payTypeIdCol - 1]) : "",
      createdAt: submittedAt,
      description: jobName && workDate ? jobName + " • " + workDate : jobName || workDate,
      correction: correction,
    });
  }
  return { requests };
}

function findRequestById_(requestId) {
  const reqSheet = getSheetByPossibleNames_(REQUESTS_SHEET_NAMES);
  const headers = getHeaderMap_(reqSheet);
  const reqSheetName = reqSheet.getName();
  const shiftIdCol = getRequiredColumn_(headers, ["ID משמרת"], reqSheetName);
  const statusCol = getRequiredColumn_(headers, ["סטטוס", "סטטוס בקשה"], reqSheetName);
  const typeCol = getOptionalColumn_(headers, ["סוג בקשה"]);
  const empIdCol = getOptionalColumn_(headers, ["מזהה עובד", "ID עובד", "ת.ז", "תז"]);
  const empNameCol = getOptionalColumn_(headers, ["שם מלא"]);
  const jobIdCol = getOptionalColumn_(headers, ["ID סוג עבודה", "ID סוגי עבודה"]);
  const jobNameCol = getOptionalColumn_(headers, ["סוג עבודה", "סוגי עבודה"]);
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const workDateCol = getOptionalColumn_(headers, ["תאריך משמרת", "תיקון תאריך", "חותמת זמן"]);
  const directionCol = getOptionalColumn_(headers, ["דיווח שעות", "כניסה / יציאה"]);
  const fixDateCol = getOptionalColumn_(headers, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(headers, ["תיקון שעה"]);
  const unitsCol = getOptionalColumn_(headers, ["כמות היחידות", "דיווח יחידות"]);
  const noteCol = getOptionalColumn_(headers, ["הערה למנהל", "הערות", "הערות למשמרת"]);
  const payTypeCol = getOptionalColumn_(headers, ["אופן תשלום", "אופני תשלום"]);
  const payTypeIdCol = getOptionalColumn_(headers, ["ID אופן תשלום", "ID אופני תשלום"]);
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

function handleShiftReportSubmit_(payload) {
  if (!payload) throw new Error("Missing payload for shiftReport.submit");
  const nowIso = new Date().toISOString();
  const shiftId = payload.shiftId || Utilities.getUuid();
  const employeeId = stringValue(payload.employeeId);
  if (!employeeId) throw new Error("Missing employeeId");

  const employeesSheet = getEmployeesSheet_();
  const empHeaders = getHeaderMap_(employeesSheet);
  const empSheetName = employeesSheet.getName();
  const empIdCol = getRequiredColumn_(
    empHeaders,
    ["מזהה עובד", "ID עובד"],
    empSheetName
  );
  const nameCol = getRequiredColumn_(empHeaders, ["שם מלא"], empSheetName);
  const tzCol = getRequiredColumn_(empHeaders, ["ת.ז", "תז"], empSheetName);
  const emailCol = getOptionalColumn_(empHeaders, EMAIL_HEADER_CANDIDATES);
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

  const jobTypes = listJobTypes_();
  const jobMap = {};
  jobTypes.forEach((j) => (jobMap[j.id] = j));
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
  const direction = stringValue(payload.direction);
  const fixDate = stringValue(payload.fixDate);
  const fixTime = stringValue(payload.fixTime);
  const units =
    payload.units !== null && payload.units !== undefined ? payload.units : "";
  const timestamp = stringValue(payload.timestamp) || nowIso;
  const workDate = stringValue(payload.workDate) || timestamp.split("T")[0];

  const employeePay = jobTypeId
    ? resolveEmployeeJobPayType_(employeesSheet, empHeaders, employeeId, jobTypeId)
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

  if (payType === "monthly") {
    return {
      ok: true,
      status: "monthly_no_report",
      message: "סוג העבודה הזה לא דורש דיווח. אם זו טעות, לחץ כאן.",
    };
  }

  const timestampDate = timestamp.split("T")[0];
  const reportMode = stringValue(payload.mode || payload.reportMode);
  const manualDate = stringValue(
    payload.manualDate || payload.workDate || payload.fixDate || workDate
  );
  const manualTime = stringValue(payload.manualTime || payload.fixTime || fixTime);
  const isManualHourly =
    payType === "hourly" &&
    (reportMode.toLowerCase() === "manual" || !!manualDate || !!manualTime);

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
      reqSheetName
    );
    getRequiredColumn_(headers, ["ID משמרת"], reqSheetName);
    getRequiredColumn_(headers, ["סטטוס", "סטטוס בקשה"], reqSheetName);
    getRequiredColumn_(
      headers,
      ["מזהה עובד", "ID עובד", "ת.ז", "תז"],
      reqSheetName
    );
    getRequiredColumn_(headers, ["שם מלא"], reqSheetName);
    getRequiredColumn_(headers, ["סוג עבודה", "סוגי עבודה"], reqSheetName);
    getRequiredColumn_(
      headers,
      ["ID סוג עבודה", "ID סוגי עבודה"],
      reqSheetName
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
    const row = buildRowFromHeaders_(headers, {
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
        payload.noteToManager || payload.note || payload.shiftNote
      ),
      submittedAt: nowIso,
      decidedAt: "",
      managerDecision: "",
      requestType: requestType,
      payTypeId: payTypeId,
      payType: payTypeForRequestRow,
    });
    reqSheet.appendRow(row);
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

  return writeWorkLogFromNormalizedShift_(normalized, {
    employeeId: employeeId,
    employeeName: employeeName,
    note: payload.note,
  });
}

function handleShiftReportMonthlyErrorNotify_(payload) {
  const employeeId = stringValue(payload.employeeId);
  const employeeEmail = stringValue(payload.email || payload.employeeEmail);
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
    const empIdCol = getRequiredColumn_(headers, ["מזהה עובד", "ID עובד"], sheetName);
    const nameCol = getOptionalColumn_(headers, ["שם מלא"]);
    const emailCol = getOptionalColumn_(headers, EMAIL_HEADER_CANDIDATES);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (stringValue(rows[i][empIdCol - 1]) !== employeeId) continue;
      if (!resolvedName && nameCol) resolvedName = stringValue(rows[i][nameCol - 1]);
      if (!resolvedEmail && emailCol) resolvedEmail = stringValue(rows[i][emailCol - 1]);
      break;
    }
  }

  sendMonthlyBlockEmail_({
    employeeId: employeeId,
    employeeName: resolvedName,
    employeeEmail: resolvedEmail,
    jobTypeId: jobTypeId,
    jobName: jobName,
    timestamp: new Date().toISOString(),
  });

  return { ok: true, status: "monthly_error_notified" };
}

function handleShiftCorrectionSubmit_(payload) {
  if (!payload) throw new Error("Missing payload for shiftCorrection.submit");
  const nowIso = new Date().toISOString();
  const employeeId = stringValue(payload.employeeId);
  if (!employeeId) throw new Error("Missing employeeId");

  const employeesSheet = getEmployeesSheet_();
  const empHeaders = getHeaderMap_(employeesSheet);
  const empSheetName = employeesSheet.getName();
  const empIdCol = getRequiredColumn_(
    empHeaders,
    ["מזהה עובד", "ID עובד"],
    empSheetName
  );
  const nameCol = getRequiredColumn_(empHeaders, ["שם מלא"], empSheetName);
  const tzCol = getRequiredColumn_(empHeaders, ["ת.ז", "תז"], empSheetName);
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

  const original = {
    workDate: stringValue(payload.originalWorkDate || payload.workDate),
    startTime: stringValue(payload.originalStartTime),
    endTime: stringValue(payload.originalEndTime),
    direction: stringValue(payload.originalDirection),
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
    direction: stringValue(payload.newDirection),
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
  getRequiredColumn_(headers, ["מזהה עובד", "ID עובד", "ת.ז", "תז"], reqSheetName);
  getRequiredColumn_(headers, ["שם מלא"], reqSheetName);
  getRequiredColumn_(headers, ["סוג עבודה", "סוגי עבודה"], reqSheetName);
  getRequiredColumn_(headers, ["ID סוג עבודה", "ID סוגי עבודה"], reqSheetName);
  getRequiredColumn_(headers, ["סוג בקשה"], reqSheetName);

  const newShiftId = original.shiftId || Utilities.getUuid();
  const jobTypeId = updated.jobTypeId || original.jobTypeId;
  const jobName = stringValue(
    payload.jobName || payload.newJobName || payload.originalJobName
  );
  const workDate = updated.workDate || original.workDate;
  const direction = updated.direction || original.direction;
  const units = updated.units !== "" ? updated.units : original.units;

  const noteToManager = JSON.stringify({
    original: original,
    updated: updated,
  });

  const row = buildRowFromHeaders_(headers, {
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

  reqSheet.appendRow(row);

  return {
    ok: true,
    status: "saved_as_request",
    requestType: "shift_correction",
    shiftId: newShiftId,
  };
}

function handleRequestApprove_(payload) {
  const requestId = stringValue(payload.requestId || payload.shiftId || payload.id);
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
    payTypeName
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

  writeWorkLogFromNormalizedShift_(normalized, {
    employeeId: rec.employeeId,
    employeeName: rec.employeeName,
    note: updated.note,
  });

  if (rec.statusCol) {
    found.sheet.getRange(found.rowIndex, rec.statusCol).setValue("approved");
  }
  if (rec.decidedAtCol) {
    found.sheet
      .getRange(found.rowIndex, rec.decidedAtCol)
      .setValue(new Date().toISOString());
  }

  return { ok: true, status: "approved", requestType: "job_not_linked" };
}

function handleEmployeeSave_(payload) {
  const employeeId = stringValue(payload.employeeId);
  if (!employeeId) throw new Error("Missing employeeId");

  const forbiddenName = payload.hasOwnProperty("fullName") || payload.hasOwnProperty("name") || payload.hasOwnProperty("employeeName");
  if (forbiddenName) return { ok: false, error: "fullName_read_only" };

  const sheet = getEmployeesSheet_();
  let headers = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const empIdCol = getRequiredColumn_(headers, ["מזהה עובד", "ID עובד"], sheetName);
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
    sheet.getRange(targetRow, birthCol).setValue(stringValue(payload.birthDate));
  }
  if (payload.size !== undefined && payload.size !== null) {
    sheet.getRange(targetRow, sizeCol).setValue(stringValue(payload.size));
    sheet.getRange(targetRow, sizeSourceCol).setValue("webapp");
  }

  const response = {
    id: employeeId,
    phone: payload.phone !== undefined ? stringValue(payload.phone) : undefined,
    birthDate:
      payload.birthDate !== undefined ? stringValue(payload.birthDate) : undefined,
    size: payload.size !== undefined && payload.size !== null ? stringValue(payload.size) : undefined,
    sizeSource:
      payload.size !== undefined && payload.size !== null ? "webapp" : undefined,
  };

  return { ok: true, status: "updated", employee: response };
}

function employeeExistsByEmail_(payload) {
  const email = stringValue(
    payload && (payload.email || payload.mail)
  ).toLowerCase();
  Logger.log("[employeeExistsByEmail] email=" + email);

  if (!email) {
    return { ok: false, success: false, error: "missing email" };
  }

  const sheet = getEmployeesSheet_();
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  Logger.log(
    "[employeeExistsByEmail] sheet=" +
      sheetName +
      ", id=" +
      SpreadsheetApp.getActive().getId()
  );

  const colEmail = getRequiredColumn_(
    headerMap,
    EMAIL_HEADER_CANDIDATES,
    sheetName
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

    return {
      ok: true,
      success: true,
      found: true,
      employee: {
        id: colId ? stringValue(row[colId - 1]) : "",
        name: colName ? stringValue(row[colName - 1]) : "",
        email: rowEmail,
        active: active,
        status: statusVal,
      },
    };
  }

  return { ok: true, success: true, found: false, employee: null };
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
  if (key.indexOf("יחיד") !== -1 || key.indexOf("unit") !== -1)
    return "unit";
  if (key.indexOf("יומי") !== -1 || key.indexOf("day") !== -1)
    return "daily";
  if (key.indexOf("שעת") !== -1 || key.indexOf("hour") !== -1)
    return "hourly";
  return "";
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
    sheet.getName()
  );
  const jobTypeCol = getRequiredColumn_(
    headers,
    ["ID סוג עבודה", "ID סוגי עבודה"],
    sheet.getName()
  );
  const payIdCol = getOptionalColumn_(
    headers,
    ["ID אופן תשלום", "ID אופני תשלום"]
  );
  const payNameCol = getOptionalColumn_(
    headers,
    ["אופן תשלום", "אופני תשלום"]
  );
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
  const recipients = MONTHLY_ALERT_RECIPIENTS.length
    ? MONTHLY_ALERT_RECIPIENTS
    : ACCESS_ISSUE_RECIPIENTS;
  if (!recipients || !recipients.length) return;
  const subject = "דיווח נחסם – אופן תשלום חודשי";
  const bodyLines = [
    "דיווח נחסם בגלל אופן תשלום חודשי:",
    "זמן: " + (details.timestamp || ""),
    "עובד: " + (details.employeeName || "") + " (" + (details.employeeId || "") + ")",
    "אימייל: " + (details.employeeEmail || ""),
    "סוג עבודה: " + (details.jobName || "") + " (" + (details.jobTypeId || "") + ")",
  ];
  const body = bodyLines.filter(function (l) {
    return stringValue(l);
  }).join("\n");
  recipients.forEach(function (to) {
    if (!to) return;
    MailApp.sendEmail({ to: to, subject: subject, body: body });
  });
}

function writeWorkLogFromNormalizedShift_(normalized, meta) {
  const logsSheet = getSheetOrThrow_(WORK_LOGS_SHEET_NAME);
  const headers = getHeaderMap_(logsSheet);
  const logsSheetName = logsSheet.getName();
  getRequiredColumn_(headers, ["ID משמרת", "ID דיווח"], logsSheetName);
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

  const row = buildRowFromHeaders_(headers, {
    note: stringValue(meta.note),
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

  logsSheet.appendRow(row);
  return {
    ok: true,
    success: true,
    requiresApproval: false,
    shiftId: normalized.shiftId,
    status: "logged",
  };
}

// Helpers
function getEmployeesSheet_() {
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
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();
  for (let i = 0; i < names.length; i++) {
    const sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  throw new Error("Sheet not found: " + names.join(", "));
}

function getSheetOrThrow_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error("Sheet not found: " + name);
  return sheet;
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
  throw new Error(
    "Missing required header '" +
      candidates[0] +
      "' in sheet '" +
      sheetName +
      "'"
  );
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
    (a, b) => headerMap[a] - headerMap[b]
  );
  return headers.map((header) => {
    const normalized = HEADER_KEY_MAP[header] || header;
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

function linkJobTypeToEmployee_(employeeId, employeeName, jobTypeId, jobName, department, payTypeId, payTypeName) {
  const sheet = getEmployeesSheet_();
  const headers = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const empIdCol = getRequiredColumn_(headers, ["מזהה עובד", "ID עובד"], sheetName);
  const empNameCol = getRequiredColumn_(headers, ["שם מלא"], sheetName);
  const jobTypeIdCol = getRequiredColumn_(headers, ["ID סוג עבודה", "ID סוגי עבודה"], sheetName);
  const jobNameCol = getRequiredColumn_(headers, ["סוג עבודה", "סוגי עבודה"], sheetName);
  const deptCol = getOptionalColumn_(headers, ["מחלקה", "מחלקות"]);
  const payIdCol = getOptionalColumn_(headers, ["ID אופן תשלום", "ID אופני תשלום"]);
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

  if (employeeName) sheet.getRange(targetRow, empNameCol).setValue(employeeName);
  sheet.getRange(targetRow, empIdCol).setValue(employeeId);
  sheet.getRange(targetRow, jobTypeIdCol).setValue(jobTypeId || "");
  sheet.getRange(targetRow, jobNameCol).setValue(jobName || "");
  if (deptCol) sheet.getRange(targetRow, deptCol).setValue(department || "");
  if (payIdCol) sheet.getRange(targetRow, payIdCol).setValue(payTypeId || "");
  if (payNameCol) sheet.getRange(targetRow, payNameCol).setValue(payTypeName || "");
}

function ensureEmployeeSizeColumns_(sheet, headers) {
  let sizeCol = getOptionalColumn_(headers, ["מידה (WebApp)", "מידת חולצה (WebApp)"]); 
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
