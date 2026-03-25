/* LEGACY SHADOW FILE
 * This file is kept for visibility and comparison only.
 * Canonical active counterpart: boulder-script-project/entrypoints.js
 * Do not treat this file as the primary entrypoint path.
 */
// entrypoints.gs

/**
 * נקודת כניסה ל-POST מהווב-אפ (דיווח משמרת חדשה)
 */
function doPost_legacy(e) {
  try {
    var action =
      e && e.parameter && e.parameter.action ? String(e.parameter.action) : "";

    // אם מגיע ?action=request → כותבים ל"בקשות עובדים"
    if (action === "request" || action === "requests") {
      return handleRequestPost(e);
    }

    // ברירת מחדל: משמרות (ההתנהגות הקיימת שלך)
    return handleShiftPost(e);
  } catch (err) {
    return jsonResponse({
      success: false,
      error: "UNEXPECTED",
      details: String(err),
    });
  }
}

function doGet_legacy(e) {
  var params = (e && e.parameter) || {};
  // תומך גם ב-action (החדש) וגם ב-mode (הישן)
  var action = params.action || params.mode || "ping";

  if (action === "ping" || action === "health") {
    return jsonResponse({ success: true, message: "ok" });
  }

  if (action === "employeeExistsByEmail") {
    return handleEmployeeExistsByEmail(params);
  }

  if (action === "shifts" || action === "listShifts") {
    return handleGetShifts(params);
  }

  if (action === "employees" || action === "listEmployees") {
    return handleGetEmployees(params);
  }

  return jsonResponse({
    success: false,
    error: "Unknown action: " + action,
  });
}
