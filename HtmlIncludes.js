/**
 * Shared HtmlService include helper.
 * Usage inside templates: <?!= include_("UiTheme"); ?>
 */
// eslint-disable-next-line no-unused-vars
function include_(filename) {
  if (!filename) return "";
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
