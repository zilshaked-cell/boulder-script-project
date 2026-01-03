function _getLib_() {
  var lib =
    typeof ExportGSS !== "undefined"
      ? ExportGSS
      : typeof exportgoogleSS !== "undefined"
      ? exportgoogleSS
      : null;
  if (!lib) {
    throw new Error(
      'Library not found. פתח "Libraries": הוסף את ה-Script ID של הספרייה,' +
        " בחר Version/Development, וודא שה-Identifier הוא ExportGSS או exportgoogleSS."
    );
  }
  return lib;
}

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // GPT Export (ראשי)
  try {
    ui.createMenu("GPT Export")
      .addSubMenu(
        ui
          .createMenu("Export")
          .addItem("All Sheets (Auto)", "gpt_all")
          .addItem("Selected Sheet…", "gpt_selected")
          .addItem("SCHEMA only (All)", "gpt_schema")
      )
      .addSeparator()
      .addItem("Refresh Sheet List", "gpt_refresh")
      .addItem("Open Spreadsheet Folder", "gpt_openFolder")
      .addToUi();
  } catch (err) {
    Logger.log("onOpen GPT Export error: " + err);
  }

  // GPT Export (LIB) – מושבת כדי למנוע כפילות תפריט

  // תפריט מאוחד: בולדר חיפה (עובדים + משמרות)
  try {
    const root = ui.createMenu("בולדר חיפה");
    if (typeof appendLegacyBoulderMenuItems_ === "function") {
      appendLegacyBoulderMenuItems_(root);
    }
    if (typeof buildShiftsSubMenu_ === "function") {
      root.addSubMenu(buildShiftsSubMenu_(ui));
    }
    root.addToUi();
  } catch (errMenu) {
    Logger.log("onOpen Boulder menu error: " + errMenu);
  }

  // מודול אופציות (UUID לאופציות) – אוטומטי
  try {
    if (typeof OPT_onOpen === "function") {
      OPT_onOpen(e);
    }
  } catch (errOpt) {
    Logger.log("onOpen OPT_onOpen error: " + errOpt);
  }

  // מודול עובדים – מפעיל לוגיקה אבל מדלג על יצירת תפריט פנימי (כדי לא להכפיל)
  try {
    if (
      e &&
      e.authMode === ScriptApp.AuthMode.FULL &&
      typeof EMP_onOpen === "function"
    ) {
      var empEvent = e || {};
      empEvent.__skipEmpMenu = true;
      EMP_onOpen(empEvent);
    }
  } catch (errEmp) {
    Logger.log("onOpen EMP_onOpen error: " + errEmp);
  }

  try {
    if (typeof REQ_onOpen === "function") {
      REQ_onOpen(e || {});
    }
  } catch (errReq) {
    Logger.log("onOpen REQ_onOpen error: " + errReq);
  }

  try {
    if (typeof SCH_onOpen === "function") {
      SCH_onOpen(e || {});
    }
  } catch (errSch) {
    Logger.log("onOpen SCH_onOpen error: " + errSch);
  }
}

// הפעלה פעם אחת (Run) כדי ליצור טריגר onOpen בר-אישור שמקבל הרשאות UI
function ensureInstallableOnOpenTrigger() {
  var has = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction && t.getHandlerFunction() === "onOpen";
  });
  if (!has) {
    ScriptApp.newTrigger("onOpen")
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  }
}

// הפעלה פעם אחת (Run) כדי ליצור טריגר יומי לבנייה מחדש של אתמול
function ensureShiftsDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (
      triggers[i].getHandlerFunction &&
      triggers[i].getHandlerFunction() === "SHIFTS_rebuildYesterday"
    ) {
      return;
    }
  }

  ScriptApp.newTrigger("SHIFTS_rebuildYesterday")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

/** ידיות תפריט → ספרייה */
function gpt_all() {
  _getLib_().exportAllSheets();
}
function gpt_selected() {
  _getLib_().openExportSheetDialog();
}
function gpt_schema() {
  _getLib_().exportSchemaOnlyAll();
}
function gpt_refresh() {
  _getLib_().refreshSheetList();
}
function gpt_openFolder() {
  _getLib_().openSpreadsheetFolder();
}

/** (ה-HTML רץ מהספרייה וקורא לפונקציות בספרייה; העטיפות כאן לא נדרשות אבל לא מזיקות) */
function exportSingleSheet(name) {
  return _getLib_().exportSingleSheet(name);
}
function getSheetList() {
  return _getLib_().getSheetList();
}

/** דיאגנוסטיקה (אופציונלי): Run פעם אחת ובדוק הלוג */
function _debugPingLibrary() {
  Logger.log("lib type = " + typeof _getLib_());
}
