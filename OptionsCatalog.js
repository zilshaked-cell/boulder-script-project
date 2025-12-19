var OPT = OPT || {};

(function () {
  var CONFIG = {
    SHEET_NAME_OPTIONS: "אופציות בחירה ו ID'S",
    HEADER_ROW: 1,
    COL: {
      JOB_STATUS: 1,       // A – סטטוס סוגי עבודה
      JOB_ID: 2,           // B – ID סוגי עבודה
      JOB_NAME: 3,         // C – סוגי עבודה
      JOB_DEPARTMENT: 4,   // D – מחלקות

      BONUS_STATUS: 5,     // E – סטטוס בונוס
      BONUS_ID: 6,         // F – ID בונוס
      BONUS_NAME: 7,       // G – סוג בונוס

      SENIORITY_STATUS: 8, // H – סטטוס רמת ותק
      SENIORITY_ID: 9,     // I – ID רמות ותק
      SENIORITY_NAME: 10,  // J – רמות ותק

      TRAIN_STATUS: 11,    // K – סטטוס סוג הכשרה
      TRAIN_ID: 12,        // L – ID סוג הכשרה
      TRAIN_NAME: 13,      // M – סוגי הכשרה

      PAY_STATUS: 14,      // N – סטטוס אופן תשלום
      PAY_ID: 15,          // O – ID אופן תשלום
      PAY_NAME: 16         // P – אופן תשלום
    }
  };

  var cache = null;

  function getSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG.SHEET_NAME_OPTIONS);
    if (!sh) {
      throw new Error('לא נמצאה כרטיסייה בשם "' + CONFIG.SHEET_NAME_OPTIONS + '"');
    }
    return sh;
  }

  function normalize_(val) {
    if (val === null || val === undefined) return '';
    return String(val).replace(/\s+/g, ' ').trim();
  }

  /**
   * קורא את כל טבלת האופציות ובונה קאש אחד:
   * - jobsByName / jobsById / jobsList
   * - paysByName / paysById / paysList
   */
  function buildCache_() {
    if (cache) return cache;

    var sh = getSheet_();
    var headerRow = CONFIG.HEADER_ROW;
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();

    var result = {
      jobsByName: {},
      jobsById: {},
      jobsList: [],
      paysByName: {},
      paysById: {},
      paysList: []
    };

    if (lastRow <= headerRow) {
      cache = result;
      return result;
    }

    var data = sh
      .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
      .getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // ---- סוגי עבודה ----
      var jobStatus = row[CONFIG.COL.JOB_STATUS - 1];
      var jobId = row[CONFIG.COL.JOB_ID - 1];
      var jobName = row[CONFIG.COL.JOB_NAME - 1];
      var dept = row[CONFIG.COL.JOB_DEPARTMENT - 1];

      if (jobId || jobName) {
        var job = {
          id: jobId || '',
          name: jobName || '',
          department: dept || '',
          status: normalize_(jobStatus)
        };

        var keyName = normalize_(job.name);
        if (keyName && !result.jobsByName[keyName]) {
          result.jobsByName[keyName] = job;
        }
        if (job.id && !result.jobsById[String(job.id)]) {
          result.jobsById[String(job.id)] = job;
        }
        result.jobsList.push(job);
      }

      // ---- אופני תשלום ----
      var payStatus = row[CONFIG.COL.PAY_STATUS - 1];
      var payId = row[CONFIG.COL.PAY_ID - 1];
      var payName = row[CONFIG.COL.PAY_NAME - 1];

      if (payId || payName) {
        var pay = {
          id: payId || '',
          name: payName || '',
          status: normalize_(payStatus)
        };

        var pKey = normalize_(pay.name);
        if (pKey && !result.paysByName[pKey]) {
          result.paysByName[pKey] = pay;
        }
        if (pay.id && !result.paysById[String(pay.id)]) {
          result.paysById[String(pay.id)] = pay;
        }
        result.paysList.push(pay);
      }
    }

    cache = result;
    return result;
  }

   function getJobByName_(name) {
    var maps = buildJobMaps_();
    var key = normalizeKey_(name);
    if (!key) return null;
    return maps.byName[key] || null;
  }

  function getPaymentByName_(name) {
    var maps = buildPaymentMaps_();
    var key = normalizeKey_(name);
    if (!key) return null;
    return maps.byName[key] || null;
  }

  /**
   * מחזירה את כל סוגי העבודה כ-Array של:
   * { id, name, department, status }
   * אם onlyActive = true – מחזירה רק סטטוס "פעיל" (או ריק).
   */
  function getAllJobs_(onlyActive) {
    var maps = buildJobMaps_();
    var list = [];
    for (var key in maps.byName) {
      if (!maps.byName.hasOwnProperty(key)) continue;
      var rec = maps.byName[key];
      if (!rec) continue;
      var status = String(rec.status || '').trim();
      if (onlyActive && status && status !== 'פעיל') {
        continue;
      }
      list.push(rec);
    }
    list.sort(function (a, b) {
      return String(a.name || '').localeCompare(String(b.name || ''));
    });
    return list;
  }

  /**
   * מחזירה את כל אופני התשלום כ-Array של:
   * { id, name, status }
   * אם onlyActive = true – מחזירה רק סטטוס "פעיל" (או ריק).
   */
  function getAllPayments_(onlyActive) {
    var maps = buildPaymentMaps_();
    var list = [];
    for (var key in maps.byName) {
      if (!maps.byName.hasOwnProperty(key)) continue;
      var rec = maps.byName[key];
      if (!rec) continue;
      var status = String(rec.status || '').trim();
      if (onlyActive && status && status !== 'פעיל') {
        continue;
      }
      list.push(rec);
    }
    list.sort(function (a, b) {
      return String(a.name || '').localeCompare(String(b.name || ''));
    });
    return list;
  }

   // חשיפה החוצה
  OPT.getJobByName      = getJobByName_;
  OPT.getPaymentByName  = getPaymentByName_;
  OPT.getAllJobs        = getAllJobs_;
  OPT.getAllPayments    = getAllPayments_;
})();
