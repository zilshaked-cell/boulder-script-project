/* exported AUDIT_logEvent */

/**
 * מחזיר את גיליון ה-AUDIT_LOG, ויוצר אותו אם לא קיים.
 */
function getAuditSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheetName = "AUDIT_LOG";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow([
      "timestamp",
      "traceId",
      "eventType",
      "entityType",
      "entityId",
      "actorEmail",
      "actorRole",
      "source",
      "summary",
      "extraJson",
    ]);
  }
  return sheet;
}

/**
 * רישום אירוע Audit כללי.
 *
 * evt = {
 *   traceId,
 *   eventType,
 *   entityType,
 *   entityId,
 *   actorEmail,
 *   actorRole,
 *   source,
 *   summary,
 *   extra: {before, after, ...} // אופציונלי
 * }
 */
// eslint-disable-next-line no-unused-vars
function AUDIT_logEvent(evt) {
  if (!evt) {
    return;
  }

  var sheet = getAuditSheet_();
  var timestamp = new Date();

  var traceId = evt.traceId || "";
  var eventType = evt.eventType || "";
  var entityType = evt.entityType || "";
  var entityId = evt.entityId || "";
  var actorEmail = evt.actorEmail || "";
  var actorRole = evt.actorRole || "";
  var source = evt.source || "";
  var summary = evt.summary || "";
  var extra = evt.extra || null;

  var extraJson = "";
  if (extra) {
    try {
      extraJson = JSON.stringify(extra);
    } catch (e) {
      extraJson = "";
    }
  }

  sheet.appendRow([
    timestamp,
    traceId,
    eventType,
    entityType,
    entityId,
    actorEmail,
    actorRole,
    source,
    summary,
    extraJson,
  ]);
}
