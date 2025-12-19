function jsonResponse(obj, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // כותרות בסיסיות ל־CORS, כדי שהווב־אפ יוכל לקרוא את ה־JSON
  output.setHeader("Access-Control-Allow-Origin", "*");
  output.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  output.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // סטטוס לא באמת משפיע ב־Apps Script, אבל נשאיר אותו לשימוש עתידי
  return output;
}
