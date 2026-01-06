/* global SpreadsheetApp */
/* exported AUDIT_logEvent */

(function () {
  "use strict";

  var SHEET_NAME = "AUDIT_LOG";
  var DEFAULT_HEADERS = [
    "timestamp",
    "eventType",
    "entityType",
    "entityId",
    "actorEmail",
    "actorRole",
    "traceId",
    "summaryBefore",
    "summaryAfter",
    "extraJson",
  ];

  function getAuditSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(DEFAULT_HEADERS);
      try {
        sh.setFrozenRows(1);
      } catch (_e) {
        // ignore freeze errors in restricted contexts
      }
      return sh;
    }

    var lastRow = sh.getLastRow();
    if (lastRow < 1) {
      sh.appendRow(DEFAULT_HEADERS);
      try {
        sh.setFrozenRows(1);
      } catch (_e2) {}
    }

    return sh;
  }

  function buildHeaderIndex_(headers) {
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var key = String(headers[i] || "").trim();
      if (key) map[key] = i;
    }
    return map;
  }

  function safeStringify_(obj) {
    try {
      return JSON.stringify(obj == null ? {} : obj);
    } catch (_e) {
      return "{}";
    }
  }

  // eslint-disable-next-line no-unused-vars
  function AUDIT_logEvent(event) {
    var sh = getAuditSheet_();
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var hIndex = buildHeaderIndex_(headers);

    var row = new Array(headers.length);
    function set_(name, value) {
      if (hIndex.hasOwnProperty(name)) {
        row[hIndex[name]] = value;
      }
    }

    var evt = event || {};
    set_("timestamp", new Date());
    set_("eventType", evt.eventType || "");
    set_("entityType", evt.entityType || "");
    set_("entityId", evt.entityId || "");
    set_("actorEmail", evt.actorEmail || "");
    set_("actorRole", evt.actorRole || "");
    set_("traceId", evt.traceId || "");
    set_("summaryBefore", evt.summaryBefore || "");
    set_("summaryAfter", evt.summaryAfter || "");
    set_("extraJson", safeStringify_(evt.extra));

    // Backfill any missing cells with empty strings to match column count.
    for (var i = 0; i < row.length; i++) {
      if (row[i] === undefined) row[i] = "";
    }

    sh.appendRow(row);
    return { ok: true, written: true, sheet: SHEET_NAME };
  }

  if (typeof globalThis !== "undefined") {
    globalThis.AUDIT_logEvent = AUDIT_logEvent;
  } else {
    this.AUDIT_logEvent = AUDIT_logEvent;
  }
})();
