/** @type {import('eslint').Linter.Config} */
module.exports = {
  env: {
    es2021: true,
    node: true,
  },
  extends: ["eslint:recommended"],
  parserOptions: {
    ecmaVersion: 12,
    sourceType: "script",
  },
  globals: {
    // ─── Google Apps Script built-in globals ───
    SpreadsheetApp: "readonly",
    ContentService: "readonly",
    HtmlService: "readonly",
    LockService: "readonly",
    PropertiesService: "readonly",
    ScriptApp: "readonly",
    MailApp: "readonly",
    Logger: "readonly",
    Session: "readonly",
    Utilities: "readonly",
    GoogleAppsScript: "readonly",

    // ─── Project module namespaces (var X = X || {}) ───
    EMP: "writable",
    OPT: "writable",
    REQ: "writable",
    SCH: "writable",

    // ─── Constants (index.gs) ───
    SHEET_NAME: "readonly",
    SHIFTS_SHEET_NAME: "readonly",
    WORK_LOGS_SHEET_NAME: "readonly",
    SHIFTS_HEADERS_CANONICAL: "readonly",
    ACCESS_ISSUE_RECIPIENTS: "readonly",

    // ─── Utility functions (index.gs) ───
    jsonResponse: "readonly",
    stringValue: "readonly",
    toIsoDate_: "readonly",
    getSheetOrThrow_: "readonly",
    getHeaderMap_: "readonly",
    getRequiredColumn_: "readonly",
    getOptionalColumn_: "readonly",
    getShiftsSheetAndHeaderMap_: "readonly",
    colIndexByHeader_: "readonly",

    // ─── Logging & tracing (index.gs) ───
    ensureModuleLoggerDefined_: "readonly",
    logDuration_: "readonly",
    appendSystemLog_: "readonly",
    createScriptLogger_: "readonly",
    makeTraceId_: "readonly",

    // ─── Cross-file functions (EmployeesModule.js) ───
    EMP_getEmployeeColumns_: "readonly",
    EMP_fillJobIdForRow_: "readonly",
    EMP_checkEmployeesHeader_: "readonly",
    getEmployeeById_: "readonly",
    buildJobNameToIdMap_: "readonly",
    normalizeKey_: "readonly",

    // ─── Cross-file functions (OptionsCatalog.js) ───
    buildJobMaps_: "readonly",
    buildPaymentMaps_: "readonly",

    // ─── Module event handlers (cross-file triggers) ───
    OPT_onOpen: "readonly",
    OPT_onEdit: "readonly",
    REQ_onEdit: "readonly",
    SCH_onChange: "readonly",

    // ─── API handlers (separate files) ───
    handleGetEmployees: "readonly",
    handleEmployeeExistsByEmail: "readonly",
    handleGetShifts: "readonly",
    handleShiftPost: "readonly",
    handleRequestPost: "readonly",
  },
  rules: {
    // GAS runtime invokes many global entrypoints dynamically (menus/triggers/doGet/doPost),
    // so static unused-vars detection produces high false-positive noise in this project.
    "no-unused-vars": "off",
    "no-undef": "error",
    "no-inner-declarations": "off",
    "no-prototype-builtins": "off",
    "no-empty": ["error", { allowEmptyCatch: true }],
    "no-redeclare": "off",
  },
};
