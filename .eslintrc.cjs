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
    // Google Apps Script globals
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
    OPT: "readonly",
    REQ: "readonly",
    SCH: "readonly",
    jsonResponse: "readonly",
  },
  rules: {
    // השארתי יחסית מינימלי כדי לא לשבור את הקוד הקיים
    "no-unused-vars": ["warn", { args: "none", vars: "all" }],
    "no-undef": "error",
    "no-inner-declarations": "off",
    "no-prototype-builtins": "off",
  },
};
