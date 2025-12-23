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
  // קודם כל – התפריט של GPT Export
  const ui = SpreadsheetApp.getUi();
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

  // מודול אופציות (UUID לאופציות) – אוטומטי
  if (typeof OPT_onOpen === "function") {
    OPT_onOpen(e);
  }

  // מודול עובדים
  EMP_onOpen(e);

  try {
    if (typeof REQ_onOpen === "function") {
      REQ_onOpen(e || {});
    }
  } catch (err) {
    Logger.log("REQ_onOpen error: " + err);
  }

  try {
    SCH_onOpen(e || {});
  } catch (err) {
    Logger.log("SCH_onOpen error: " + err);
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
