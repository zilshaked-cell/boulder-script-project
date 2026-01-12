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

function aggLogger_(operation, loggerOpt) {
  if (loggerOpt) return loggerOpt;
  try {
    if (typeof ensureModuleLoggerDefined_ === "function") {
      return ensureModuleLoggerDefined_(operation || "SHIFTS_AGGREGATOR");
    }
  } catch (_ignored) {}
  return null;
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

function daysBetweenInclusive_(fromIso, toIso) {
  const from = isoToDate_(fromIso);
  const to = isoToDate_(toIso);
  if (!from || !to) return null;
  const oneDay = 24 * 60 * 60 * 1000;
  const diff = Math.floor((to.getTime() - from.getTime()) / oneDay);
  return diff + 1; // inclusive
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

function translateShiftStatusAndNote_(status, note) {
  const statusMap = {
    OK: "תקין",
    MISSING_IN: "חסרה כניסה",
    MISSING_OUT: "חסרה יציאה",
    CONFLICT: "תקלה",
    MISSING_ENTRY: "חסרה כניסה",
    MISSING_EXIT: "חסרה יציאה",
  };

  const noteMap = {
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

  const statusDisplay = statusMap[status] || status || "";
  const normalizedNote = noteMap[note] || note || "";

  return { statusDisplay: statusDisplay, noteDisplay: normalizedNote };
}

function readWorkLogsForRange_(opts) {
  const logger = aggLogger_("AGG_READ_LOGS");
  const startMs = new Date().getTime();
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

  if (typeof logDuration_ === "function") {
    logDuration_(logger, "shifts.readWorkLogsForRange", startMs, {
      logs: logs.length,
      from: dateFrom || "",
      to: dateTo || "",
      employeeId: employeeFilter || "",
    });
  }

  return logs;
}

function buildAggregatedShifts_(logs, debugExplain, loggerOpt) {
  const logger = aggLogger_("AGG_BUILD_SHIFTS", loggerOpt);
  const startMs = new Date().getTime();
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
  const statusCountsTotal = {};

  for (let k = 0; k < keys.length; k++) {
    const groupLogs = grouped[keys[k]].slice().sort(compareLogs_);
    const built = [];
    let open = null;
    const groupKey =
      "emp:" +
      String((groupLogs[0] && groupLogs[0].employeeId) || "") +
      "|date:" +
      String((groupLogs[0] && groupLogs[0].workDateIso) || "") +
      "|job:" +
      String((groupLogs[0] && groupLogs[0].jobTypeId) || "none");

    if (debugExplain && logger) {
      logger.info("shifts.agg.debug.timeline", {
        groupKey: groupKey,
        logsCount: groupLogs.length,
        logs: groupLogs.map(function (log) {
          var dt = log && log.effectiveDateTime;
          return {
            rawLogId:
              log && log.rawLogId !== undefined && log.rawLogId !== null
                ? log.rawLogId
                : log && log.shiftIdRaw,
            direction: log && log.direction,
            effectiveDateTimeIso: dt && dt.toISOString ? dt.toISOString() : "",
            logRowIndex: log && log.logRowIndex,
          };
        }),
      });
    }

    for (let i = 0; i < groupLogs.length; i++) {
      const log = groupLogs[i];
      if (log.direction === "IN") {
        if (open) {
          var shiftExtraIn = buildShiftRecord_(
            [open],
            "CONFLICT",
            "Extra IN before OUT",
            open,
            null
          );
          built.push(shiftExtraIn);
          if (debugExplain && logger) {
            logger.info("shifts.agg.debug.shift", {
              groupKey: groupKey,
              rule: "EXTRA_IN_BEFORE_OUT",
              status: shiftExtraIn.status,
              note: shiftExtraIn.note,
              rawLogIds: shiftExtraIn.rawLogIds,
              startTimeIso:
                shiftExtraIn.startTime && shiftExtraIn.startTime.toISOString
                  ? shiftExtraIn.startTime.toISOString()
                  : "",
              endTimeIso:
                shiftExtraIn.endTime && shiftExtraIn.endTime.toISOString
                  ? shiftExtraIn.endTime.toISOString()
                  : "",
              hoursDecimal: shiftExtraIn.hoursDecimal,
              payHours: shiftExtraIn.payHours,
            });
          }
        }
        open = log;
        continue;
      }

      if (log.direction === "OUT") {
        if (open) {
          var shiftPair = buildShiftRecord_([open, log], "OK", "", open, log);
          built.push(shiftPair);
          if (debugExplain && logger) {
            logger.info("shifts.agg.debug.shift", {
              groupKey: groupKey,
              rule: "PAIR_IN_OUT",
              status: shiftPair.status,
              note: shiftPair.note,
              rawLogIds: shiftPair.rawLogIds,
              startTimeIso:
                shiftPair.startTime && shiftPair.startTime.toISOString
                  ? shiftPair.startTime.toISOString()
                  : "",
              endTimeIso:
                shiftPair.endTime && shiftPair.endTime.toISOString
                  ? shiftPair.endTime.toISOString()
                  : "",
              hoursDecimal: shiftPair.hoursDecimal,
              payHours: shiftPair.payHours,
            });
          }
          open = null;
        } else {
          var shiftMissingIn = buildShiftRecord_(
            [log],
            "MISSING_IN",
            "Missing IN",
            null,
            log
          );
          built.push(shiftMissingIn);
          if (debugExplain && logger) {
            logger.info("shifts.agg.debug.shift", {
              groupKey: groupKey,
              rule: "MISSING_IN",
              status: shiftMissingIn.status,
              note: shiftMissingIn.note,
              rawLogIds: shiftMissingIn.rawLogIds,
              startTimeIso:
                shiftMissingIn.startTime && shiftMissingIn.startTime.toISOString
                  ? shiftMissingIn.startTime.toISOString()
                  : "",
              endTimeIso:
                shiftMissingIn.endTime && shiftMissingIn.endTime.toISOString
                  ? shiftMissingIn.endTime.toISOString()
                  : "",
              hoursDecimal: shiftMissingIn.hoursDecimal,
              payHours: shiftMissingIn.payHours,
            });
          }
        }
        continue;
      }

      if (open) {
        var shiftUnknownWithOpen = buildShiftRecord_(
          [open, log],
          "CONFLICT",
          "Unknown direction",
          open,
          log
        );
        built.push(shiftUnknownWithOpen);
        if (debugExplain && logger) {
          logger.info("shifts.agg.debug.shift", {
            groupKey: groupKey,
            rule: "UNKNOWN_DIRECTION",
            status: shiftUnknownWithOpen.status,
            note: shiftUnknownWithOpen.note,
            rawLogIds: shiftUnknownWithOpen.rawLogIds,
            startTimeIso:
              shiftUnknownWithOpen.startTime &&
              shiftUnknownWithOpen.startTime.toISOString
                ? shiftUnknownWithOpen.startTime.toISOString()
                : "",
            endTimeIso:
              shiftUnknownWithOpen.endTime &&
              shiftUnknownWithOpen.endTime.toISOString
                ? shiftUnknownWithOpen.endTime.toISOString()
                : "",
            hoursDecimal: shiftUnknownWithOpen.hoursDecimal,
            payHours: shiftUnknownWithOpen.payHours,
          });
        }
        open = null;
      } else {
        var shiftUnknown = buildShiftRecord_(
          [log],
          "CONFLICT",
          "Unknown direction",
          log,
          null
        );
        built.push(shiftUnknown);
        if (debugExplain && logger) {
          logger.info("shifts.agg.debug.shift", {
            groupKey: groupKey,
            rule: "UNKNOWN_DIRECTION",
            status: shiftUnknown.status,
            note: shiftUnknown.note,
            rawLogIds: shiftUnknown.rawLogIds,
            startTimeIso:
              shiftUnknown.startTime && shiftUnknown.startTime.toISOString
                ? shiftUnknown.startTime.toISOString()
                : "",
            endTimeIso:
              shiftUnknown.endTime && shiftUnknown.endTime.toISOString
                ? shiftUnknown.endTime.toISOString()
                : "",
            hoursDecimal: shiftUnknown.hoursDecimal,
            payHours: shiftUnknown.payHours,
          });
        }
      }
    }

    if (open) {
      var shiftMissingOut = buildShiftRecord_(
        [open],
        "MISSING_OUT",
        "Missing OUT",
        open,
        null
      );
      built.push(shiftMissingOut);
      if (debugExplain && logger) {
        logger.info("shifts.agg.debug.shift", {
          groupKey: groupKey,
          rule: "MISSING_OUT",
          status: shiftMissingOut.status,
          note: shiftMissingOut.note,
          rawLogIds: shiftMissingOut.rawLogIds,
          startTimeIso:
            shiftMissingOut.startTime && shiftMissingOut.startTime.toISOString
              ? shiftMissingOut.startTime.toISOString()
              : "",
          endTimeIso:
            shiftMissingOut.endTime && shiftMissingOut.endTime.toISOString
              ? shiftMissingOut.endTime.toISOString()
              : "",
          hoursDecimal: shiftMissingOut.hoursDecimal,
          payHours: shiftMissingOut.payHours,
        });
      }
    }

    var statusCounts = {};
    for (let i = 0; i < built.length; i++) {
      const s = built[i].status || "UNKNOWN";
      statusCounts[s] = (statusCounts[s] || 0) + 1;
    }

    if (debugExplain && logger) {
      var allLogIds = {};
      groupLogs.forEach(function (log) {
        var rawId =
          log && log.rawLogId !== undefined && log.rawLogId !== null
            ? log.rawLogId
            : log && log.shiftIdRaw;
        if (rawId !== undefined && rawId !== null && rawId !== "") {
          allLogIds[String(rawId)] = true;
        }
      });

      var usedLogIds = {};
      built.forEach(function (shift) {
        if (!shift || !shift.rawLogIds) return;
        String(shift.rawLogIds)
          .split(",")
          .map(function (s) {
            return stringValue(s);
          })
          .filter(function (s) {
            return !!s;
          })
          .forEach(function (id) {
            usedLogIds[id] = true;
          });
      });

      var unusedLogIds = [];
      Object.keys(allLogIds).forEach(function (id) {
        if (!usedLogIds[id]) unusedLogIds.push(id);
      });

      logger.info("shifts.agg.debug.group-summary", {
        groupKey: groupKey,
        logsCount: groupLogs.length,
        shiftsCount: built.length,
        statusCounts: statusCounts,
        usedLogIds: Object.keys(usedLogIds),
        unusedLogIds: unusedLogIds,
      });
    }

    for (let i = 0; i < built.length; i++) {
      var shiftRecord = built[i];
      shiftRecord.shiftId = buildShiftId_(shiftRecord, i + 1);
      var statusKey = shiftRecord.status || "UNKNOWN";
      statusCountsTotal[statusKey] = (statusCountsTotal[statusKey] || 0) + 1;
      results.push(shiftRecord);
    }
  }

  if (typeof logDuration_ === "function") {
    logDuration_(logger, "shifts.buildAggregated", startMs, {
      inputLogs: logs ? logs.length : 0,
      shifts: results.length,
      debugExplain: !!debugExplain,
    });
  }

  if (logger) {
    logger.info("shifts.agg.aggregate.summary", {
      groups: keys.length,
      shifts: results.length,
      statusCounts: statusCountsTotal,
      debugExplain: !!debugExplain,
    });
  }

  return results;
}

function buildShiftRows_(shifts, headerMap) {
  const map = headerMap || {};
  return shifts.map(function (shift) {
    const workDate = isoToDate_(shift.workDateIso);
    const translated = translateShiftStatusAndNote_(shift.status, shift.note);

    const row = new Array(SHIFTS_HEADERS_CANONICAL.length).fill("");
    row[map["ID משמרת"]] = shift.shiftId;
    row[map["מזהה עובד"]] = shift.employeeId;
    row[map["שם מלא"]] = shift.employeeName;
    row[map["תאריך משמרת"]] = workDate;
    row[map["שעת התחלה"]] = shift.startTime || "";
    row[map["שעת סיום"]] = shift.endTime || "";
    row[map["ID סוג עבודה"]] = shift.jobTypeId;
    row[map["סוג עבודה"]] = shift.jobName;
    row[map["מחלקה"]] = shift.department;
    row[map["שעות"]] =
      shift.hoursDecimal === "" ? "" : Number(shift.hoursDecimal);
    if (map["שעות לשכר"] !== undefined && map["שעות לשכר"] !== null) {
      row[map["שעות לשכר"]] =
        shift.payHours === "" ? "" : Number(shift.payHours);
    }
    row[map["כמות יחידות"]] = shift.units;
    row[map["סטטוס משמרת"]] = translated.statusDisplay;

    // Prefer primary duplicates for writes; leave legacy columns empty to avoid ambiguity.
    if (map.notesPrimary !== undefined && map.notesPrimary !== null) {
      row[map.notesPrimary] = translated.noteDisplay;
    }
    if (map.sourcePrimary !== undefined && map.sourcePrimary !== null) {
      row[map.sourcePrimary] = shift.rawLogIds || "";
    }

    return row;
  });
}

function rebuildShiftsForRange_(opts, loggerOpt) {
  var logger = aggLogger_("SHIFTS_AGGREGATOR", loggerOpt);
  var debugExplain = !!(opts && opts.debugExplain);
  var fnStartMs = new Date().getTime();

  try {
    const filters = opts || {};
    const dateFrom = toIsoDate_(filters.dateFrom || "");
    const dateTo = toIsoDate_(filters.dateTo || "");
    const employeeFilter = stringValue(filters.employeeId);
    if (!dateFrom || !dateTo) {
      throw new Error("rebuildShiftsForRange_: dateFrom/dateTo are required");
    }

    const daysRequested = daysBetweenInclusive_(dateFrom, dateTo);
    if (daysRequested === null) {
      throw new Error("rebuildShiftsForRange_: invalid date range");
    }
    const MAX_DAYS = 40;
    if (daysRequested > MAX_DAYS) {
      throw new Error(
        "rebuildShiftsForRange_: range too large (" +
          daysRequested +
          " days > " +
          MAX_DAYS +
          ")"
      );
    }

    if (logger) {
      logger.info("shifts.agg.rebuild.start", {
        dateFrom: dateFrom,
        dateTo: dateTo,
        employeeId: employeeFilter,
        jobTypeId: stringValue(filters.jobTypeId),
        debugExplain: debugExplain,
      });
    }

    // Lock previous month after the 9th of current month: do not edit/write rows before the start of the current month.
    const today = new Date();
    const currentMonthStart = new Date(
      today.getFullYear(),
      today.getMonth(),
      1
    );
    const lockActive = today.getDate() >= 9;
    const lockedCutoffIso = lockActive ? toIsoDate_(currentMonthStart) : "";

    const ctx = getShiftsSheetAndHeaderMap_();
    const sheet = ctx.sheet;
    const headerMap = ctx.headerMap;
    const workDateCol = headerMap["תאריך משמרת"];
    const employeeCol = headerMap["מזהה עובד"];
    const startCol = headerMap["שעת התחלה"];
    const endCol = headerMap["שעת סיום"];
    const hoursCol = headerMap["שעות"];

    var readStartMs = new Date().getTime();
    const workLogs = readWorkLogsForRange_(filters);
    logDuration_(logger, "shifts.agg.read-logs", readStartMs, {
      logs: workLogs.length,
      dateFrom: dateFrom,
      dateTo: dateTo,
      employeeId: employeeFilter,
    });

    var aggregateStartMs = new Date().getTime();
    const shifts = buildAggregatedShifts_(workLogs, debugExplain, logger);
    logDuration_(logger, "shifts.agg.aggregate", aggregateStartMs, {
      logs: workLogs.length,
      shifts: shifts.length,
      debugExplain: debugExplain,
    });

    const rows = buildShiftRows_(shifts, headerMap);

    // Delete only matching range/employee, and never touch locked dates.
    var deleteStartMs = new Date().getTime();
    var deleted = 0;
    if (sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const toDelete = [];
      for (let i = 1; i < data.length; i++) {
        const rowDate = toIsoDate_(data[i][workDateCol]);
        if (!rowDate) continue;
        if (lockedCutoffIso && rowDate < lockedCutoffIso) continue; // locked past period
        if (rowDate < dateFrom || rowDate > dateTo) continue;
        if (
          employeeFilter &&
          stringValue(data[i][employeeCol]) !== employeeFilter
        )
          continue;
        toDelete.push(i + 1);
      }
      deleted = toDelete.length;
      for (let idx = toDelete.length - 1; idx >= 0; idx--) {
        sheet.deleteRow(toDelete[idx]);
      }
    }
    logDuration_(logger, "shifts.agg.delete-old", deleteStartMs, {
      deleted: deleted,
      lockedCutoff: lockedCutoffIso || "",
    });

    // Filter out shifts that fall into locked periods.
    const writableRows = [];
    for (let i = 0; i < rows.length; i++) {
      const workDateIso = shifts[i] ? shifts[i].workDateIso : "";
      if (lockedCutoffIso && workDateIso && workDateIso < lockedCutoffIso) {
        continue; // skip locked
      }
      writableRows.push(rows[i]);
    }

    const skippedLocked = rows.length - writableRows.length;
    var writeStartMs = new Date().getTime();
    var result;

    if (!writableRows.length) {
      logDuration_(logger, "shifts.agg.write-shifts", writeStartMs, {
        rowsToWrite: 0,
        skippedLocked: skippedLocked,
        startRow: null,
      });
      result = {
        ok: true,
        shiftsWritten: 0,
        skippedLocked: skippedLocked,
        daysRequested: daysRequested,
      };
      logDuration_(logger, "shifts.agg.total", fnStartMs, {
        dateFrom: dateFrom,
        dateTo: dateTo,
        employeeId: employeeFilter,
        shiftsWritten: result.shiftsWritten,
        skippedLocked: result.skippedLocked,
        debugExplain: debugExplain,
      });
      return result;
    }

    const startRow = sheet.getLastRow() + 1;
    sheet
      .getRange(startRow, 1, writableRows.length, writableRows[0].length)
      .setValues(writableRows);
    sheet
      .getRange(startRow, hoursCol + 1, writableRows.length, 1)
      .setNumberFormat("0.00");
    const payCol = headerMap["שעות לשכר"];
    if (payCol !== undefined && payCol !== null) {
      sheet
        .getRange(startRow, payCol + 1, writableRows.length, 1)
        .setNumberFormat("0.00");
    }

    const startBg = [];
    const endBg = [];
    for (let i = 0; i < shifts.length; i++) {
      const workDateIso = shifts[i] ? shifts[i].workDateIso : "";
      if (lockedCutoffIso && workDateIso && workDateIso < lockedCutoffIso) {
        continue; // skipped locked
      }
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

    sheet
      .getRange(startRow, startCol + 1, startBg.length, 1)
      .setBackgrounds(startBg);
    sheet.getRange(startRow, endCol + 1, endBg.length, 1).setBackgrounds(endBg);

    result = {
      ok: true,
      shiftsWritten: writableRows.length,
      skippedLocked: skippedLocked,
      daysRequested: daysRequested,
    };

    logDuration_(logger, "shifts.agg.write-shifts", writeStartMs, {
      rowsToWrite: writableRows.length,
      skippedLocked: skippedLocked,
      startRow: startRow,
    });

    logDuration_(logger, "shifts.agg.total", fnStartMs, {
      dateFrom: dateFrom,
      dateTo: dateTo,
      employeeId: employeeFilter,
      shiftsWritten: result.shiftsWritten,
      skippedLocked: result.skippedLocked,
      debugExplain: debugExplain,
    });

    return result;
  } catch (err) {
    if (logger && logger.error) {
      logger.error(
        "shifts.agg.error",
        {
          filters: opts,
          errorMessage: String((err && err.message) || err),
        },
        "SHIFTS_AGG_ERROR",
        err
      );
    }
    throw err;
  }
}

function explainShiftById_(shiftId, loggerOpt) {
  var logger = aggLogger_("SHIFTS_EXPLAIN", loggerOpt);
  var ss = ss_();
  void ss; // keep reference explicit for clarity

  var ctx = getShiftsSheetAndHeaderMap_();
  var sheet = ctx && ctx.sheet ? ctx.sheet : null;
  var headerMap = ctx && ctx.headerMap ? ctx.headerMap : {};
  var sheetName = sheet ? sheet.getName() : "";

  if (!sheet) {
    if (logger) {
      logger.info("shifts.explain.not-found", {
        shiftId: shiftId,
        reason: "no-sheet",
      });
    }
    return;
  }

  var shiftIdCol = getRequiredColumn_(headerMap, ["ID משמרת"], sheetName);

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    if (logger) {
      logger.info("shifts.explain.no-rows", { shiftId: shiftId });
    }
    return;
  }

  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var values = range.getValues();

  var foundRowIndex = null;
  var foundEmployeeId = null;
  var foundWorkDate = null;
  var foundJobTypeId = null;

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (String(row[shiftIdCol - 1]) === String(shiftId)) {
      foundRowIndex = i + 2;
      // TODO: השלם חילוץ employeeId, workDate, jobTypeId לפי העמודות הידועות
      break;
    }
  }

  if (!foundRowIndex) {
    if (logger) {
      logger.info("shifts.explain.not-found", { shiftId: shiftId });
    }
    return;
  }

  if (logger) {
    logger.info("shifts.explain.found-shift", {
      shiftId: shiftId,
      rowIndex: foundRowIndex,
      // employeeId: foundEmployeeId,
      // workDate: foundWorkDate,
      // jobTypeId: foundJobTypeId,
    });
  }
}
