/* exported jsonResponse */

// eslint-disable-next-line no-unused-vars
function jsonResponse(obj, statusCode) {
  var output = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
  return output;
}

// Global sheet name for the shift report sheet ("דיווח שעות עבודה").
// Used by shifts_write.js, shifts_read.js and test_harness.js via SHEET_NAME.
if (typeof SHEET_NAME === "undefined") {
  var SHEET_NAME = "דיווח שעות עבודה";
}
