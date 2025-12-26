var SCH = SCH || {};

(function () {
  "use strict";

  var CONFIG = {
    HEADER_ROW: 1,
    PROP_PREFIX: "SCHEMA_SNAPSHOT__",
    MAIL: {
      TO: "ZIL.SHAKED@GMAIL.COM",
      SUBJECT_PREFIX: "שכר בולדר חיפה – שינוי סכימה",
    },
    WATCH_SHEETS: [
      // שים לב: "פרטי עובדים" כבר מנוטר אצלך ב-EmployeesModule.
      // אם אתה רוצה איחוד מלא ולמנוע מייל כפול – השאר פה false או תמחק את המנגנון הישן שם.
      // 'פרטי עובדים',

      "אופציות בחירה ו ID'S",
      "דיווח שעות עבודה",
      "בקשות עובדים",
    ],
  };

  function ss_() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function props_() {
    return PropertiesService.getDocumentProperties();
  }

  function snapshot_(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) lastCol = 1;

    var headers = sheet
      .getRange(CONFIG.HEADER_ROW, 1, 1, lastCol)
      .getValues()[0];
    var out = [];
    for (var i = 0; i < headers.length; i++) {
      out.push(String(headers[i] || "").trim() + "#" + (i + 1));
    }
    return out.join("||");
  }

  function sendEmail_(sheetName, prev, current, reason) {
    // ⚠️ MailApp עלול להיכשל ב-simple trigger אם אין הרשאות/טריגר מותקן.
    // לכן אנחנו עוטפים ב-try/catch ולא מפילים כלום.
    try {
      var user =
        (Session.getActiveUser && Session.getActiveUser().getEmail()) ||
        "unknown";
      var ss = ss_();

      var body = "";
      body +=
        "זוהה שינוי במבנה (כותרות/מיקומים) בכרטיסייה: " + sheetName + "\n";
      body += "קובץ: " + ss.getName() + "\n";
      body += "סיבה/טריגר: " + (reason || "לא ידוע") + "\n";
      body += "מבצע: " + user + "\n\n";
      body += "Snapshot קודם:\n" + prev + "\n\n";
      body += "Snapshot נוכחי:\n" + current + "\n";

      MailApp.sendEmail({
        to: CONFIG.MAIL.TO,
        subject:
          CONFIG.MAIL.SUBJECT_PREFIX +
          " – " +
          sheetName +
          " (" +
          (reason || "change") +
          ")",
        body: body,
      });
    } catch (err) {
      Logger.log('SCH sendEmail failed for "%s": %s', sheetName, err);
    }
  }

  function checkOne_(sheetName, reason) {
    var sheet = ss_().getSheetByName(sheetName);
    if (!sheet) return { ok: false, sheet: sheetName, status: "missing" };

    var key = CONFIG.PROP_PREFIX + sheetName;
    var p = props_();
    var prev = p.getProperty(key) || "";
    var current = snapshot_(sheet);

    if (!prev) {
      p.setProperty(key, current);
      return { ok: true, sheet: sheetName, status: "initialized" };
    }

    if (prev === current) {
      return { ok: true, sheet: sheetName, status: "unchanged" };
    }

    p.setProperty(key, current);
    sendEmail_(sheetName, prev, current, reason);
    return { ok: true, sheet: sheetName, status: "changed" };
  }

  function checkAll_(reason) {
    var res = [];
    for (var i = 0; i < CONFIG.WATCH_SHEETS.length; i++) {
      res.push(checkOne_(CONFIG.WATCH_SHEETS[i], reason));
    }
    return res;
  }

  SCH.checkAll = checkAll_;
  SCH.onOpen = function (e) {
    return checkAll_("onOpen");
  };
  SCH.onChange = function (e) {
    var t = e && e.changeType ? String(e.changeType) : "onChange";
    return checkAll_(t);
  };
})();

// wrappers גלובליים
function SCH_onOpen(e) {
  if (typeof SCH !== "undefined" && SCH.onOpen) return SCH.onOpen(e || {});
}
function SCH_onChange(e) {
  if (typeof SCH !== "undefined" && SCH.onChange) return SCH.onChange(e || {});
}
