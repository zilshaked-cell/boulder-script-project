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

const HEADER_KEY_MAP = {
  "ID משמרת": "shiftId",
  "ID דיווח": "shiftId",
  "Shift ID": "shiftId",
  "Request ID": "id",
  requestId: "id",
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
  const dataRange = sheet.getDataRange().getValues();
  let blanks = 0;
  const results = [];
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const name = stringValue(row[nameCol - 1]);
    const status = stringValue(row[statusCol - 1]);
    if (!name) {
      blanks++;
      if (blanks >= 20) break;
      continue;
    }
    blanks = 0;
    if (status !== "פעיל") continue;
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
    requests.push({
      id: stringValue(row[shiftIdCol - 1]),
      status: stringValue(row[statusCol - 1]),
      jobName: jobName,
      workDate: workDateCol ? stringValue(row[workDateCol - 1]) : "",
      requestedSummary: stringValue(row[directionCol - 1]),
      units: unitsCol ? stringValue(row[unitsCol - 1]) : "",
      noteToManager: noteCol ? stringValue(row[noteCol - 1]) : "",
      submittedAt: submittedCol ? stringValue(row[submittedCol - 1]) : "",
      decidedAt: decidedCol ? stringValue(row[decidedCol - 1]) : "",
      type: requestType,
      employeeId: employeeId,
      employeeName: fullName,
      shiftId: stringValue(row[shiftIdCol - 1]),
      jobTypeId: stringValue(row[jobTypeIdCol - 1]),
      fixDate: fixDateCol ? stringValue(row[fixDateCol - 1]) : "",
      fixTime: fixTimeCol ? stringValue(row[fixTimeCol - 1]) : "",
    });
  }
  return { requests };
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
      workDate: workDate,
      direction: direction,
      fixDate: fixDate,
      fixTime: fixTime,
      units: units,
      noteToManager: stringValue(
        payload.noteToManager || payload.note || payload.shiftNote
      ),
      submittedAt: nowIso,
      decidedAt: "",
      managerDecision: "",
      requestType: requestType,
    });
    reqSheet.appendRow(row);
    return {
      ok: true,
      success: true,
      requiresApproval: true,
      shiftId: shiftId,
      requestType: requestType,
    };
  }

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
    note: stringValue(payload.note),
    shiftId: shiftId,
    timestamp: timestamp,
    employeeId: employeeId,
    employeeName: employeeName,
    direction: direction,
    fixDate: fixDate,
    fixTime: fixTime,
    jobTypeId: jobTypeId,
    jobName: jobName,
    department: department,
    units: units,
  });
  logsSheet.appendRow(row);
  return { ok: true, success: true, requiresApproval: false, shiftId: shiftId };
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
