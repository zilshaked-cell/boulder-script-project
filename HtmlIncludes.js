/**
 * Shared HtmlService include helper.
 * Usage inside templates: <?!= include_("UiTheme"); ?>
 */
function include_(filename) {
  if (!filename) return "";
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
