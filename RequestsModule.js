var REQ = REQ || {};

(function () {
  'use strict';

  var CONFIG = {
    SHEET_NAME_REQUESTS: 'בקשות עובדים',
    HEADER_ROW: 1
  };

  function ss_() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function sh_() {
    var sh = ss_().getSheetByName(CONFIG.SHEET_NAME_REQUESTS);
    if (!sh) throw new Error('לא נמצאה כרטיסייה "' + CONFIG.SHEET_NAME_REQUESTS + '"');
    return sh;
  }

  function uuid_() {
    return Utilities.getUuid();
  }

  function norm_(v) {
    if (v === null || v === undefined) return '';
    return String(v).replace(/\s+/g, ' ').trim();
  }

  function headersMap_(sheet) {
    var sh = sheet || sh_();
    var lastCol = sh.getLastColumn();
    if (lastCol < 1) return {};
    var headers = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
    var map = {};
    for (var i = 0; i < headers.length; i++) {
      var h = norm_(headers[i]);
      if (!h) continue;
      map[h] = i + 1; // 1-based
    }
    return map;
  }

  function col_(map, header) {
    return map[header] || null;
  }

  /**
   * ממלא UUID בעמודות ID שיש להן כותרת שמכילה "ID" + שיש תוכן בשורה (כלומר לא שורה ריקה).
   * ברירת מחדל: רק עמודת "ID בקשה" (אם קיימת) כדי לא לגעת בשדות אחרים.
   */
  function ensureRequestIds_(sheet) {
    var sh = sheet || sh_();
    var lastRow = sh.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return { filled: 0 };

    var map = headersMap_(sh);

    // אנחנו תומכים בכותרת מדויקת "ID בקשה".
    // אם אין, ננסה למצוא עמודה שהכותרת שלה כוללת "ID" וגם "בקשה".
    var idCol = col_(map, 'ID בקשה');
    if (!idCol) {
      // fallback חיפוש כותרת
      var keys = Object.keys(map);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        if (k.indexOf('ID') !== -1 && k.indexOf('בקשה') !== -1) {
          idCol = map[k];
          break;
        }
      }
    }

    if (!idCol) {
      // אין עמודת ID בקשה - לא עושים כלום
      return { filled: 0, warning: 'לא נמצאה עמודת "ID בקשה"' };
    }

    var lastCol = sh.getLastColumn();
    var dataRange = sh.getRange(CONFIG.HEADER_ROW + 1, 1, lastRow - CONFIG.HEADER_ROW, lastCol);
    var data = dataRange.getValues();

    // נגדיר "שורה לא ריקה" ככזו שיש בה משהו חוץ מה-ID עצמו
    var filled = 0;
    var existing = {};
    for (var r = 0; r < data.length; r++) {
      var idv = norm_(data[r][idCol - 1]);
      if (idv) existing[idv] = true;
    }

    for (var r2 = 0; r2 < data.length; r2++) {
      var row = data[r2];
      var currId = norm_(row[idCol - 1]);
      if (currId) continue;

      // בדיקת שורה לא ריקה (יש תוכן כלשהו במישהו מהתאים חוץ מעמודת ה-ID)
      var hasContent = false;
      for (var c = 0; c < row.length; c++) {
        if (c === (idCol - 1)) continue;
        if (norm_(row[c])) { hasContent = true; break; }
      }
      if (!hasContent) continue;

      var newId;
      do { newId = uuid_(); } while (existing[newId]);
      existing[newId] = true;
      row[idCol - 1] = newId;
      filled++;
    }

    if (filled > 0) {
      dataRange.setValues(data);
    }

    return { filled: filled };
  }

  function handleOpen_(e) {
    try { ensureRequestIds_(); } catch (err) {}
  }

  function handleEdit_(e) {
    var range = e && e.range ? e.range : null;
    if (!range) return;

    var sheet = range.getSheet();
    if (!sheet || sheet.getName() !== CONFIG.SHEET_NAME_REQUESTS) return;

    // אם עריכה מתרחשת בטאב בקשות – נוודא IDs
    ensureRequestIds_(sheet);
  }

  // Export
  REQ.ensureRequestIds = ensureRequestIds_;
  REQ.handleOpen = handleOpen_;
  REQ.handleEdit = handleEdit_;

  // --- wrappers גלובליים ---
  function REQ_onOpen(e) {
    if (typeof REQ !== 'undefined' && REQ.handleOpen) {
      REQ.handleOpen(e || {});
    }
  }
  function REQ_onEdit(e) {
    if (typeof REQ !== 'undefined' && REQ.handleEdit) {
      REQ.handleEdit(e || {});
    }
  }

  // חשיפה של ה-wrappers לשימוש חיצוני
  REQ.REQ_onOpen = REQ_onOpen;
  REQ.REQ_onEdit = REQ_onEdit;

})();

function REQ_onOpen(e) {
  if (typeof REQ !== 'undefined' && REQ.REQ_onOpen) {
    REQ.REQ_onOpen(e || {});
  }
}

function REQ_onEdit(e) {
  if (typeof REQ !== 'undefined' && REQ.REQ_onEdit) {
    REQ.REQ_onEdit(e || {});
  }
}
