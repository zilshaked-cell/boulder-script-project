function normalizeLogDirection_(raw) {
  const key = stringValue(raw).toLowerCase();
  if (!key) return "UNKNOWN";
  if (key.indexOf("כניסה") !== -1 || key.indexOf("in") !== -1) return "IN";
  if (key.indexOf("יציאה") !== -1 || key.indexOf("out") !== -1) return "OUT";
  return "UNKNOWN";
}

function parseTimeParts_(val) {
  if (!val && val !== 0) return null;
  if (
    Object.prototype.toString.call(val) === "[object Date]" &&
    !isNaN(val.getTime())
  ) {
    return {
      hh: ("0" + val.getHours()).slice(-2),
      mm: ("0" + val.getMinutes()).slice(-2),
    };
  }
  const s = stringValue(val);
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return { hh: ("0" + m[1]).slice(-2), mm: ("0" + m[2]).slice(-2) };
}

function buildEffectiveDateTime_(dateVal, timeVal, timestampVal) {
  const dateIso = toIsoDate_(dateVal);
  const timeParts = parseTimeParts_(timeVal);
  if (dateIso && timeParts) {
    const combined = new Date(
      dateIso + "T" + timeParts.hh + ":" + timeParts.mm + ":00"
    );
    if (!isNaN(combined.getTime())) return combined;
  }

  if (
    Object.prototype.toString.call(timestampVal) === "[object Date]" &&
    !isNaN(timestampVal.getTime())
  ) {
    return timestampVal;
  }

  const tsStr = stringValue(timestampVal);
  if (tsStr) {
    const parsed = new Date(tsStr);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function compareLogs_(a, b) {
  const aTime = a && a.effectiveDateTime ? a.effectiveDateTime.getTime() : 0;
  const bTime = b && b.effectiveDateTime ? b.effectiveDateTime.getTime() : 0;
  if (aTime !== bTime) return aTime - bTime;
  return (
    (a && a.logRowIndex ? a.logRowIndex : 0) -
    (b && b.logRowIndex ? b.logRowIndex : 0)
  );
}

function sumUnitsFromLogs_(logs) {
  let total = 0;
  let hasValue = false;
  for (let i = 0; i < logs.length; i++) {
    const raw = logs[i] ? logs[i].unitsRaw : "";
    const num = typeof raw === "number" ? raw : parseFloat(stringValue(raw));
    if (isNaN(num)) continue;
    hasValue = true;
    total += num;
  }
  return hasValue ? total : "";
}

function collectRawLogIds_(logs) {
  const ids = [];
  for (let i = 0; i < logs.length; i++) {
    const candidate = stringValue(logs[i] ? logs[i].shiftIdRaw : "");
    const fallback =
      candidate ||
      "row" + (logs[i] && logs[i].logRowIndex ? logs[i].logRowIndex : "");
    const id = stringValue(fallback);
    if (id && ids.indexOf(id) === -1) ids.push(id);
  }
  return ids.join(",");
}

function pickFirst_(logs, getter) {
  for (let i = 0; i < logs.length; i++) {
    const val = getter(logs[i]);
    if (stringValue(val)) return val;
  }
  return "";
}

function computeHoursDiff_(start, end) {
  if (!start || !end) return null;
  const startMs = start.getTime();
  const endMs = end.getTime();
  let diff = endMs - startMs;
  if (diff < 0) diff += 24 * 60 * 60 * 1000;
  return Math.round((diff / (1000 * 60 * 60)) * 100) / 100;
}

function isoToDate_(iso) {
  const s = stringValue(iso);
  if (!s) return null;
  const parts = s.split("-");
  if (parts.length === 3) {
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const d = parseInt(parts[2], 10);
    const dt = new Date(y, m, d);
    if (!isNaN(dt.getTime())) return dt;
  }
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function buildShiftId_(shift, seq) {
  const job = stringValue(shift.jobTypeId) || "NOJOB";
  const emp = stringValue(shift.employeeId) || "";
  const date = stringValue(shift.workDateIso) || "";
  return emp + "__" + date + "__" + job + "__" + seq;
}

function buildShiftRecord_(logs, status, note, startLog, endLog) {
  const startTime =
    startLog && startLog.effectiveDateTime ? startLog.effectiveDateTime : null;
  const endTime =
    endLog && endLog.effectiveDateTime ? endLog.effectiveDateTime : null;

  const shift = {
    employeeId: pickFirst_(logs, function (l) {
      return l && l.employeeId;
    }),
    employeeName: pickFirst_(logs, function (l) {
      return l && l.employeeName;
    }),
    jobTypeId: pickFirst_(logs, function (l) {
      return l && l.jobTypeId;
    }),
    jobName: pickFirst_(logs, function (l) {
      return l && l.jobName;
    }),
    department: pickFirst_(logs, function (l) {
      return l && l.department;
    }),
    workDateIso: pickFirst_(logs, function (l) {
      return l && l.workDateIso;
    }),
    startTime: startTime,
    endTime: endTime,
    status: status,
    note: note || "",
    units: sumUnitsFromLogs_(logs),
    rawLogIds: collectRawLogIds_(logs),
  };

  shift.hoursDecimal =
    startTime && endTime ? computeHoursDiff_(startTime, endTime) : "";

  shift.payHours =
    shift.hoursDecimal === "" || shift.hoursDecimal === null
      ? ""
      : status === "OK"
      ? shift.hoursDecimal
      : "";

  return shift;
}

function readWorkLogsForRange_(opts) {
  const filters = opts || {};
  const dateFrom = toIsoDate_(filters.dateFrom || "");
  const dateTo = toIsoDate_(filters.dateTo || "");
  const employeeFilter = stringValue(filters.employeeId);

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
  const empNameCol = getOptionalColumn_(headerMap, ["שם מלא"]);
  const dirCol = getRequiredColumn_(
    headerMap,
    ["כניסה / יציאה", "דיווח שעות"],
    sheetName
  );
  const fixDateCol = getOptionalColumn_(headerMap, ["תיקון תאריך"]);
  const fixTimeCol = getOptionalColumn_(headerMap, ["תיקון שעה"]);
  const workDateCol = getOptionalColumn_(headerMap, ["תאריך משמרת"]);
  const jobTypeIdCol = getOptionalColumn_(headerMap, [
    "ID סוג עבודה",
    "ID סוגי עבודה",
  ]);
  const jobNameCol = getOptionalColumn_(headerMap, ["סוג עבודה", "סוגי עבודה"]);
  const deptCol = getOptionalColumn_(headerMap, ["מחלקה", "מחלקות"]);
  const unitsCol = getOptionalColumn_(headerMap, [
    "כמות יחידות",
    "כמות היחידות",
    "דיווח יחידות",
  ]);

  const values = sheet.getDataRange().getValues();
  const logs = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const employeeId = stringValue(row[empIdCol - 1]);
    if (employeeFilter && employeeId !== employeeFilter) continue;

    const rawTs = row[tsCol - 1];
    const tsStr =
      rawTs instanceof Date && !isNaN(rawTs.getTime())
        ? rawTs.toISOString()
        : stringValue(rawTs);
    const fixDateRaw = fixDateCol ? row[fixDateCol - 1] : "";
    const fixTimeRaw = fixTimeCol ? row[fixTimeCol - 1] : "";
    const workDateRaw = workDateCol ? row[workDateCol - 1] : "";
    const workDateIso = toIsoDate_(fixDateRaw || workDateRaw || tsStr);
    if (dateFrom && workDateIso && workDateIso < dateFrom) continue;
    if (dateTo && workDateIso && workDateIso > dateTo) continue;

    const log = {
      shiftIdRaw: stringValue(row[idCol - 1]),
      timestamp: tsStr,
      employeeId: employeeId,
      employeeName: empNameCol ? stringValue(row[empNameCol - 1]) : "",
      direction: normalizeLogDirection_(row[dirCol - 1]),
      directionRaw: stringValue(row[dirCol - 1]),
      fixDate: toIsoDate_(fixDateRaw),
      fixTime: stringValue(fixTimeRaw),
      workDateIso: workDateIso,
      jobTypeId: jobTypeIdCol ? stringValue(row[jobTypeIdCol - 1]) : "",
      jobName: jobNameCol ? stringValue(row[jobNameCol - 1]) : "",
      department: deptCol ? stringValue(row[deptCol - 1]) : "",
      unitsRaw: unitsCol ? row[unitsCol - 1] : "",
      logRowIndex: i,
      effectiveDateTime: buildEffectiveDateTime_(fixDateRaw, fixTimeRaw, rawTs),
    };

    if (!log.workDateIso) continue;
    logs.push(log);
  }

  return logs;
}

function buildAggregatedShifts_(logs) {
  const grouped = {};
  for (let i = 0; i < logs.length; i++) {
    const log = logs[i];
    const key =
      stringValue(log.employeeId) +
      "|" +
      stringValue(log.workDateIso) +
      "|" +
      stringValue(log.jobTypeId);
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(log);
  }

  const keys = Object.keys(grouped).sort();
  const results = [];

  for (let k = 0; k < keys.length; k++) {
    const groupLogs = grouped[keys[k]].slice().sort(compareLogs_);
    const built = [];
    let open = null;

    for (let i = 0; i < groupLogs.length; i++) {
      const log = groupLogs[i];
      if (log.direction === "IN") {
        if (open) {
          built.push(
            buildShiftRecord_(
              [open],
              "CONFLICT",
              "Extra IN before OUT",
              open,
              null
            )
          );
        }
        open = log;
        continue;
      }

      if (log.direction === "OUT") {
        if (open) {
          built.push(buildShiftRecord_([open, log], "OK", "", open, log));
          open = null;
        } else {
          built.push(
            buildShiftRecord_([log], "MISSING_IN", "Missing IN", null, log)
          );
        }
        continue;
      }

      if (open) {
        built.push(
          buildShiftRecord_(
            [open, log],
            "CONFLICT",
            "Unknown direction",
            open,
            log
          )
        );
        open = null;
      } else {
        built.push(
          buildShiftRecord_([log], "CONFLICT", "Unknown direction", log, null)
        );
      }
    }

    if (open) {
      built.push(
        buildShiftRecord_([open], "MISSING_OUT", "Missing OUT", open, null)
      );
    }

    for (let i = 0; i < built.length; i++) {
      built[i].shiftId = buildShiftId_(built[i], i + 1);
      results.push(built[i]);
    }
  }

  return results;
}

function buildShiftRows_(shifts, headerMap) {
  const sheet = getShiftsSheet_();
  const map = headerMap || getHeaderMap_(sheet);
  return shifts.map(function (shift) {
    const workDate = isoToDate_(shift.workDateIso);
    return buildRowFromHeaders_(map, {
      shiftId: shift.shiftId,
      employeeId: shift.employeeId,
      employeeName: shift.employeeName,
      workDate: workDate,
      startTime: shift.startTime || "",
      endTime: shift.endTime || "",
      jobTypeId: shift.jobTypeId,
      jobName: shift.jobName,
      department: shift.department,
      hoursDecimal: shift.hoursDecimal === "" ? "" : Number(shift.hoursDecimal),
      payHours: shift.payHours === "" ? "" : Number(shift.payHours),
      units: shift.units,
      status: shift.status,
      note: shift.note,
      rawLogIds: shift.rawLogIds,
    });
  });
}

function rebuildShiftsForRange_(opts) {
  const filters = opts || {};
  const sheet = getShiftsSheet_();
  const headerMap = getHeaderMap_(sheet);
  const sheetName = sheet.getName();
  const workDateCol = getRequiredColumn_(headerMap, ["תאריך משמרת"], sheetName);
  const employeeCol = getRequiredColumn_(
    headerMap,
    ["מזהה עובד", "ID עובד"],
    sheetName
  );
  const startCol = getRequiredColumn_(headerMap, ["שעת התחלה"], sheetName);
  const endCol = getRequiredColumn_(headerMap, ["שעת סיום"], sheetName);
  const hoursCol = getRequiredColumn_(headerMap, ["שעות"], sheetName);
  const payCol = getRequiredColumn_(headerMap, ["שעות לשכר"], sheetName);

  const shifts = buildAggregatedShifts_(readWorkLogsForRange_(filters));
  const rows = buildShiftRows_(shifts, headerMap);

  const dateFrom = toIsoDate_(filters.dateFrom || "");
  const dateTo = toIsoDate_(filters.dateTo || "");
  const employeeFilter = stringValue(filters.employeeId);
  const fullRebuild = !dateFrom && !dateTo && !employeeFilter;

  if (fullRebuild) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
  } else if (sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    const toDelete = [];
    for (let i = 1; i < data.length; i++) {
      const rowDate = toIsoDate_(data[i][workDateCol - 1]);
      if (!rowDate) continue;
      if (dateFrom && rowDate < dateFrom) continue;
      if (dateTo && rowDate > dateTo) continue;
      if (
        employeeFilter &&
        stringValue(data[i][employeeCol - 1]) !== employeeFilter
      )
        continue;
      toDelete.push(i + 1);
    }
    for (let idx = toDelete.length - 1; idx >= 0; idx--) {
      sheet.deleteRow(toDelete[idx]);
    }
  }

  if (!rows.length) return { ok: true, shiftsWritten: 0 };

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(startRow, hoursCol, rows.length, 1).setNumberFormat("0.00");
  sheet.getRange(startRow, payCol, rows.length, 1).setNumberFormat("0.00");

  const startBg = [];
  const endBg = [];
  for (let i = 0; i < shifts.length; i++) {
    const status = shifts[i].status;
    const startColor =
      status === "MISSING_IN"
        ? "#FCD34D"
        : status === "CONFLICT"
        ? "#FCA5A5"
        : "";
    const endColor =
      status === "MISSING_OUT"
        ? "#FCD34D"
        : status === "CONFLICT"
        ? "#FCA5A5"
        : "";
    startBg.push([startColor]);
    endBg.push([endColor]);
  }

  sheet.getRange(startRow, startCol, rows.length, 1).setBackgrounds(startBg);
  sheet.getRange(startRow, endCol, rows.length, 1).setBackgrounds(endBg);

  return { ok: true, shiftsWritten: rows.length };
}
