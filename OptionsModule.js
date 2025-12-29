var OPT = OPT || {};

(function () {
  "use strict";

  var CONFIG = {
    SHEET_NAME_OPTIONS: "אופציות בחירה ו ID'S",
    HEADER_ROW: 1,
    COL: {
      // לפי SHEETS_CONTRACT.md (pay columns are H-J)
      JOB_STATUS: 1, // A
      JOB_ID: 2, // B
      JOB_NAME: 3, // C
      JOB_DEPARTMENT: 4, // D

      BONUS_STATUS: 5, // E
      BONUS_ID: 6, // F
      BONUS_NAME: 7, // G

      PAY_STATUS: 8, // H
      PAY_ID: 9, // I
      PAY_NAME: 10, // J

      SENIORITY_STATUS: 11, // K
      SENIORITY_ID: 12, // L
      SENIORITY_NAME: 13, // M

      TRAIN_STATUS: 14, // N
      TRAIN_ID: 15, // O
      TRAIN_NAME: 16, // P
    },
  };

  var cache = null;

  function getSheet_() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG.SHEET_NAME_OPTIONS);
    if (!sh)
      throw new Error(
        'לא נמצאה כרטיסייה בשם "' + CONFIG.SHEET_NAME_OPTIONS + '"'
      );
    return sh;
  }

  function uuid_() {
    return Utilities.getUuid();
  }

  function norm_(v) {
    if (v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  function normKey_(v) {
    return norm_(v).toLowerCase();
  }

  function isActive_(statusVal) {
    var s = norm_(statusVal);
    if (!s) return true; // ריק = פעיל
    if (s === "פעיל") return true;
    if (String(s).toUpperCase() === "TRUE") return true;
    return false;
  }

  /**
   * מייצר UUID לכל קטלוג (jobs/bonus/pay/seniority/train) רק כשיש ערך בשם ואין ID.
   * לא מוחק/לא דורס IDs קיימים.
   * @return {{created:number}} סטטיסטיקה
   */
  function ensureCatalogIds_(sheet) {
    var sh = sheet || getSheet_();
    var headerRow = CONFIG.HEADER_ROW;
    var lastRow = sh.getLastRow();
    if (lastRow <= headerRow) return { created: 0 };

    var lastCol = sh.getLastColumn();
    var range = sh.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol);
    var values = range.getValues();

    var created = 0;
    var existingIds = {};

    // PASS 1: אוספים IDs קיימים כדי לצמצם סיכוי לדופליקציה בתוך אותו גיליון
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var idCols = [
        CONFIG.COL.JOB_ID,
        CONFIG.COL.BONUS_ID,
        CONFIG.COL.PAY_ID,
        CONFIG.COL.SENIORITY_ID,
        CONFIG.COL.TRAIN_ID,
      ];
      for (var k = 0; k < idCols.length; k++) {
        var idv = norm_(row[idCols[k] - 1]);
        if (idv) existingIds[idv] = true;
      }
    }

    function ensureOne_(idCol, nameCol) {
      for (var r = 0; r < values.length; r++) {
        var row = values[r];
        var name = norm_(row[nameCol - 1]);
        if (!name) continue;
        var id = norm_(row[idCol - 1]);
        if (id) continue;

        var newId;
        do {
          newId = uuid_();
        } while (existingIds[newId]);

        existingIds[newId] = true;
        row[idCol - 1] = newId;
        created++;
      }
    }

    ensureOne_(CONFIG.COL.JOB_ID, CONFIG.COL.JOB_NAME);
    ensureOne_(CONFIG.COL.BONUS_ID, CONFIG.COL.BONUS_NAME);
    ensureOne_(CONFIG.COL.PAY_ID, CONFIG.COL.PAY_NAME);
    ensureOne_(CONFIG.COL.SENIORITY_ID, CONFIG.COL.SENIORITY_NAME);
    ensureOne_(CONFIG.COL.TRAIN_ID, CONFIG.COL.TRAIN_NAME);

    if (created > 0) {
      range.setValues(values);
      cache = null; // invalidate
    }

    return { created: created };
  }

  /**
   * בונה קאש על כל הקטלוגים.
   * NOTE: הקאש נבנה *אחרי* שמוודאים IDs.
   */
  function buildCache_() {
    if (cache) return cache;

    var sh = getSheet_();
    ensureCatalogIds_(sh);

    var headerRow = CONFIG.HEADER_ROW;
    var lastRow = sh.getLastRow();
    if (lastRow <= headerRow) {
      cache = {
        jobsByName: {},
        jobsById: {},
        jobs: [],
        bonusesByName: {},
        bonusesById: {},
        bonuses: [],
        paysByName: {},
        paysById: {},
        pays: [],
        seniorityByName: {},
        seniorityById: {},
        seniority: [],
        trainsByName: {},
        trainsById: {},
        trains: [],
      };
      return cache;
    }

    var lastCol = sh.getLastColumn();
    var headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    function findCol_(candidates) {
      for (var i = 0; i < headers.length; i++) {
        var h = norm_(headers[i]);
        for (var c = 0; c < candidates.length; c++) {
          if (h === norm_(candidates[c])) return i + 1; // 1-based
        }
      }
      return null;
    }

    var payStatusCol =
      findCol_(["סטטוס אופן תשלום", "סטטוס אופני תשלום"]) ||
      CONFIG.COL.PAY_STATUS;
    var payIdCol =
      findCol_(["ID אופן תשלום", "ID אופני תשלום"]) || CONFIG.COL.PAY_ID;
    var payNameCol =
      findCol_(["אופן תשלום", "אופני תשלום"]) || CONFIG.COL.PAY_NAME;

    var data = sh
      .getRange(headerRow + 1, 1, lastRow - headerRow, lastCol)
      .getValues();

    var out = {
      jobsByName: {},
      jobsById: {},
      jobs: [],
      bonusesByName: {},
      bonusesById: {},
      bonuses: [],
      paysByName: {},
      paysById: {},
      pays: [],
      seniorityByName: {},
      seniorityById: {},
      seniority: [],
      trainsByName: {},
      trainsById: {},
      trains: [],
    };

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      // jobs
      var jobName = norm_(row[CONFIG.COL.JOB_NAME - 1]);
      var jobId = norm_(row[CONFIG.COL.JOB_ID - 1]);
      var jobStatus = norm_(row[CONFIG.COL.JOB_STATUS - 1]);
      var dept = norm_(row[CONFIG.COL.JOB_DEPARTMENT - 1]);
      if (jobName || jobId) {
        var job = {
          id: jobId,
          name: jobName,
          department: dept,
          status: jobStatus,
        };
        out.jobs.push(job);
        if (jobId && !out.jobsById[jobId]) out.jobsById[jobId] = job;
        var jKey = normKey_(jobName);
        if (jKey && !out.jobsByName[jKey]) out.jobsByName[jKey] = job;
      }

      // bonuses
      var bonusName = norm_(row[CONFIG.COL.BONUS_NAME - 1]);
      var bonusId = norm_(row[CONFIG.COL.BONUS_ID - 1]);
      var bonusStatus = norm_(row[CONFIG.COL.BONUS_STATUS - 1]);
      if (bonusName || bonusId) {
        var bonus = { id: bonusId, name: bonusName, status: bonusStatus };
        out.bonuses.push(bonus);
        if (bonusId && !out.bonusesById[bonusId])
          out.bonusesById[bonusId] = bonus;
        var bKey = normKey_(bonusName);
        if (bKey && !out.bonusesByName[bKey]) out.bonusesByName[bKey] = bonus;
      }

      // pays
      var payName = norm_(row[payNameCol - 1]);
      var payId = norm_(row[payIdCol - 1]);
      var payStatus = norm_(row[payStatusCol - 1]);
      if (payName || payId) {
        var pay = { id: payId, name: payName, status: payStatus };
        out.pays.push(pay);
        if (payId && !out.paysById[payId]) out.paysById[payId] = pay;
        var pKey = normKey_(payName);
        if (pKey && !out.paysByName[pKey]) out.paysByName[pKey] = pay;
      }

      // seniority
      var sName = norm_(row[CONFIG.COL.SENIORITY_NAME - 1]);
      var sId = norm_(row[CONFIG.COL.SENIORITY_ID - 1]);
      var sStatus = norm_(row[CONFIG.COL.SENIORITY_STATUS - 1]);
      if (sName || sId) {
        var sen = { id: sId, name: sName, status: sStatus };
        out.seniority.push(sen);
        if (sId && !out.seniorityById[sId]) out.seniorityById[sId] = sen;
        var sKey = normKey_(sName);
        if (sKey && !out.seniorityByName[sKey]) out.seniorityByName[sKey] = sen;
      }

      // trains
      var tName = norm_(row[CONFIG.COL.TRAIN_NAME - 1]);
      var tId = norm_(row[CONFIG.COL.TRAIN_ID - 1]);
      var tStatus = norm_(row[CONFIG.COL.TRAIN_STATUS - 1]);
      if (tName || tId) {
        var tr = { id: tId, name: tName, status: tStatus };
        out.trains.push(tr);
        if (tId && !out.trainsById[tId]) out.trainsById[tId] = tr;
        var tKey = normKey_(tName);
        if (tKey && !out.trainsByName[tKey]) out.trainsByName[tKey] = tr;
      }
    }

    cache = out;
    return out;
  }

  function getJobById_(id) {
    var key = norm_(id);
    if (!key) return null;
    return buildCache_().jobsById[key] || null;
  }

  function getJobByNameAndDepartment_(name, department) {
    var nameKey = normKey_(name);
    var deptKey = normKey_(department);
    if (!nameKey) return null;

    // Prefer exact name + department match when department provided.
    if (deptKey) {
      var jobs = buildCache_().jobs;
      for (var i = 0; i < jobs.length; i++) {
        var j = jobs[i];
        if (normKey_(j.name) === nameKey && normKey_(j.department) === deptKey) {
          return j;
        }
      }
    }

    // Fallback to name-only match.
    return buildCache_().jobsByName[nameKey] || null;
  }

  function getPaymentById_(id) {
    var key = norm_(id);
    if (!key) return null;
    return buildCache_().paysById[key] || null;
  }

  function getPaymentByName_(name) {
    var key = normKey_(name);
    if (!key) return null;
    return buildCache_().paysByName[key] || null;
  }
  function getPaymentByName_(name) {
    var key = normKey_(name);
    if (!key) return null;
    return buildCache_().paysByName[key] || null;
  }
    if (!key) return null;
    return buildCache_().jobsByName[key] || null;
  }

  function getPaymentByName_(name) {
    var key = normKey_(name);
    if (!key) return null;
    return buildCache_().paysByName[key] || null;
  }

  function getAllJobs_(onlyActive) {
    var list = buildCache_().jobs.slice();
    if (onlyActive) {
      list = list.filter(function (r) {
        return isActive_(r.status);
      });
    }
    list.sort(function (a, b) {
      return String(a.name || "").localeCompare(String(b.name || ""));
    });
    return list;
  }

  function getAllPayments_(onlyActive) {
    var list = buildCache_().pays.slice();
    if (onlyActive) {
      list = list.filter(function (r) {
        return isActive_(r.status);
      });
    }
    list.sort(function (a, b) {
      return String(a.name || "").localeCompare(String(b.name || ""));
    });
    return list;
  }

  // ---------- Triggers ----------

  function handleOpen_(e) {
    ensureCatalogIds_();
  }

  function handleEdit_(e) {
    var range = e && e.range ? e.range : null;
    if (!range) return;

    var sh = range.getSheet();
    if (!sh || sh.getName() !== CONFIG.SHEET_NAME_OPTIONS) return;

    // אם העריכה בכלל לא נוגעת לעמודות שם של קטלוגים – אין סיבה לעבוד.
    var startCol = range.getColumn();
    var endCol = startCol + range.getNumColumns() - 1;

    var nameCols = [
      CONFIG.COL.JOB_NAME,
      CONFIG.COL.BONUS_NAME,
      CONFIG.COL.PAY_NAME,
      CONFIG.COL.SENIORITY_NAME,
      CONFIG.COL.TRAIN_NAME,
    ];

    var touches = false;
    for (var i = 0; i < nameCols.length; i++) {
      var c = nameCols[i];
      if (c >= startCol && c <= endCol) {
        touches = true;
        break;
      }
    }

    if (!touches) return;

    ensureCatalogIds_(sh);
  }

  // ---------- export ----------

  OPT.ensureCatalogIds = ensureCatalogIds_;
  OPT.getJobByName = getJobByName_;
  OPT.getPaymentByName = getPaymentByName_;
  OPT.getAllJobs = getAllJobs_;
  OPT.getJobById = getJobById_;
  OPT.getJobByNameAndDepartment = getJobByNameAndDepartment_;
  OPT.getPaymentById = getPaymentById_;
  OPT.getPaymentByName = getPaymentByName_;
  OPT.getAllPayments = getAllPayments_;
  OPT.handleOpen = handleOpen_;
  OPT.handleEdit = handleEdit_;
})();

// עטיפות גלובליות (נוחות לקריאה ממקומות אחרים)
function OPT_onOpen(e) {
  if (typeof OPT !== "undefined" && OPT.handleOpen) {
    OPT.handleOpen(e || {});
  }
}

function OPT_onEdit(e) {
  if (typeof OPT !== "undefined" && OPT.handleEdit) {
    OPT.handleEdit(e || {});
  }
}
