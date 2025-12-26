/* =========================================================
   א׳) הקוד להדבקה בספרייה "export google SS"
   קובץ: export GSS.gs  (החלף את כל התוכן בקיים)
   ========================================================= */

/** ===== FIXED TARGET FOLDER (Drive) ===== */
// Parent folder: "קבצי סיכום GSS ל GPT". The script will create/use a per-spreadsheet subfolder under this.
const TARGET_FOLDER_ID = "1uALJXfY-9BTibw0BNI6JeFU4QXAMjG7y";
// Project-specific child folder: "שכר בולדר חיפה" (fallback if parent is inaccessible)
const TARGET_FOLDER_ID_FALLBACK = "1jBhLxYjWPLPmR99voDSsiTqm1kbLgDN5";

/** ===== Menu (standalone library only; not auto-wired to avoid overriding host onOpen) ===== */
function exportGSS_onOpenStandalone() {
  SpreadsheetApp.getUi()
    .createMenu("GPT Export (LIB)")
    .addItem("All Sheets (Auto)", "exportAllSheets")
    .addItem("Selected Sheet…", "openExportSheetDialog")
    .addItem("SCHEMA only (All)", "exportSchemaOnlyAll")
    .addToUi();
}

/** ===== Public entry points (נקראות מהגיליון) ===== */
function exportAllSheets() {
  const cfg = defaultConfig_({ mode: "perSheet", exportAppsScript: true });
  return exportWorkbookAll_(cfg);
}
function exportSchemaOnlyAll() {
  const cfg = defaultConfig_({
    includeCells: false,
    mode: "schemaOnly",
    exportAppsScript: false,
  });
  return exportWorkbookAll_(cfg);
}
/* ✅ חדש: פותח את הדיאלוג של בחירת כרטיסייה ומגיש את הקובץ EXPORT GSS.html */
function openExportSheetDialog() {
  const t = HtmlService.createTemplateFromFile("EXPORT GSS"); // חייב להתאים לשם הקובץ
  const html = t.evaluate().setWidth(420).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, "Export Selected Sheet");
}
/* נדרש ע"י ה-HTML */
function exportSingleSheet(sheetName) {
  const cfg = defaultConfig_({ exportAppsScript: false });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet not found: " + sheetName);

  const target = getFixedTargetFolder_();
  const spFolder = getOrCreateSpreadsheetFolderUnderTarget_(target, ss);

  removeSheetArtifacts_(spFolder, sanitize_(sheetName));

  const unit = collectSheetCore_(sh, cfg);
  writeJsonFileReplacing_(
    spFolder,
    `${unit.fileBase}.schema.json`,
    sheetSchemaPayload_(unit)
  );

  let filesRef = {
    schema: `${unit.fileBase}.schema.json`,
    data: null,
    chunks: [],
  };
  if (cfg.includeCells && unit.rows > 0 && unit.cols > 0) {
    const estBytes = unit.rows * unit.cols * cfg.approxBytesPerCell;
    const canSingle =
      estBytes <= cfg.perSheetMaxBytesForSingleData &&
      unit.rows * unit.cols <= cfg.maxCellsPerSheetForInMemory;
    if (canSingle) {
      writeJsonFileReplacing_(
        spFolder,
        `${unit.fileBase}.data.json`,
        sheetDataPayloadAll_(unit, cfg)
      );
      filesRef.data = `${unit.fileBase}.data.json`;
    } else {
      const chunkNames = writeSheetChunksFromCore_(
        spFolder,
        unit,
        cfg,
        Date.now(),
        { partial: false, warnings: [] }
      );
      filesRef.chunks = chunkNames;
    }
  }
  upsertIndexForSheet_(spFolder, ss, unit, filesRef);
  return { ok: true };
}
function refreshSheetList() {
  const names = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map((s) => s.getName());
  PropertiesService.getDocumentProperties().setProperty(
    "gptExport.sheetList",
    JSON.stringify(names)
  );
  SpreadsheetApp.getUi().alert("Sheet list refreshed (" + names.length + ").");
}
function getSheetList() {
  const prop = PropertiesService.getDocumentProperties().getProperty(
    "gptExport.sheetList"
  );
  if (prop)
    try {
      return JSON.parse(prop);
    } catch (_) {}
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map((s) => s.getName());
}
/* ✅ תוקן: טוען בדיוק את "EXPORT GSS - OPEN FOLDER.html" */
function openSpreadsheetFolder() {
  const target = getFixedTargetFolder_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spFolder = getOrCreateSpreadsheetFolderUnderTarget_(target, ss);
  const t = HtmlService.createTemplateFromFile("EXPORT GSS - OPEN FOLDER");
  t.url = spFolder.getUrl();
  const html = t.evaluate().setWidth(360).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(html, "Open Spreadsheet Folder");
}

/** ===== Config ===== */
function defaultConfig_(overrides) {
  return Object.assign(
    {
      includeCells: true,
      includeFormulas: true,
      dateAsISO: true,
      trimTrailingEmpty: true,
      maxExamplesPerColumn: 5,
      inferHeaderByFrozenRows: true,
      inferHeaderHeuristic: true,
      approxBytesPerCell: 24,
      perSheetMaxBytesForSingleData: 4 * 1024 * 1024,
      maxCellsPerSheetForInMemory: 120000,
      chunkRowsTarget: 2000,
      maxRuntimeMs: 5.5 * 60 * 1000,
      limitSheets: null,
      exportAppsScript: true,
      mode: "perSheet",
      version: "gpt-sheet-export@final-2.1.0",
    },
    overrides || {}
  );
}

/** ===== All-sheets orchestrator ===== */
function exportWorkbookAll_(cfg) {
  const t0 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const target = getFixedTargetFolder_();
  const spFolder = getOrCreateSpreadsheetFolderUnderTarget_(target, ss);
  purgeExportArtifacts_(spFolder);

  const meta = collectSpreadsheetMeta_(ss);
  const index = {
    _export: {
      version: cfg.version,
      mode: cfg.mode,
      exportedAt: new Date().toISOString(),
      timeZone: Session.getScriptTimeZone(),
      spreadsheetId: meta.id,
      spreadsheetUrl: meta.url,
    },
    spreadsheet: meta,
    sheets: [],
    apps_script: null,
    warnings: [],
    partial: false,
  };

  for (const sh of ss.getSheets()) {
    if (cfg.limitSheets && !cfg.limitSheets.includes(sh.getName())) continue;
    if (Date.now() - t0 > cfg.maxRuntimeMs) {
      index.partial = true;
      break;
    }
    try {
      const unit = collectSheetCore_(sh, cfg);
      writeJsonFileReplacing_(
        spFolder,
        `${unit.fileBase}.schema.json`,
        sheetSchemaPayload_(unit)
      );

      if (cfg.includeCells && unit.rows > 0 && unit.cols > 0) {
        const estBytes = unit.rows * unit.cols * cfg.approxBytesPerCell;
        const canSingle =
          estBytes <= cfg.perSheetMaxBytesForSingleData &&
          unit.rows * unit.cols <= cfg.maxCellsPerSheetForInMemory;
        if (canSingle) {
          writeJsonFileReplacing_(
            spFolder,
            `${unit.fileBase}.data.json`,
            sheetDataPayloadAll_(unit, cfg)
          );
          unit.refs = {
            schema: `${unit.fileBase}.schema.json`,
            data: `${unit.fileBase}.data.json`,
            chunks: [],
          };
        } else {
          const chunkNames = writeSheetChunksFromCore_(
            spFolder,
            unit,
            cfg,
            t0,
            index
          );
          unit.refs = {
            schema: `${unit.fileBase}.schema.json`,
            data: null,
            chunks: chunkNames,
          };
        }
      } else {
        unit.refs = {
          schema: `${unit.fileBase}.schema.json`,
          data: null,
          chunks: [],
        };
      }

      index.sheets.push({
        sheet: {
          name: unit.name,
          gid: unit.gid,
          index: unit.index,
          hidden: unit.hidden,
          frozen: unit.frozen,
          range: unit.a1,
        },
        header_rows: unit.headerRows,
        columns: unit.columns,
        files: unit.refs,
      });
    } catch (e) {
      index.warnings.push(
        `Sheet "${sh.getName()}": ${e && e.message ? e.message : e}`
      );
    }
  }

  if (cfg.exportAppsScript) {
    const apps = exportAppsScript_(spFolder);
    index.apps_script = apps;
    if (apps && apps.warnings && apps.warnings.length) {
      index.warnings = index.warnings.concat(
        apps.warnings.map((w) => `Apps Script: ${w}`)
      );
    }
  }

  const idx = writeJsonFileReplacing_(spFolder, "index.json", index);
  showResultDialog_({
    folderUrl: spFolder.getUrl(),
    indexFileUrl: idx.getUrl(),
    sheetCount: index.sheets.length,
    partial: index.partial,
    warnings: index.warnings,
  });
  return { folderUrl: spFolder.getUrl(), indexFileUrl: idx.getUrl() };
}

/** ===== Index upsert for single-sheet ===== */
function upsertIndexForSheet_(folder, ss, unit, filesRef) {
  const name = "index.json";
  const current = readJsonFile_(folder, name) || {
    _export: {
      version: defaultConfig_().version,
      mode: "perSheet",
      exportedAt: new Date().toISOString(),
      timeZone: Session.getScriptTimeZone(),
      spreadsheetId: ss.getId(),
      spreadsheetUrl: ss.getUrl(),
    },
    spreadsheet: collectSpreadsheetMeta_(ss),
    sheets: [],
    apps_script: null,
    warnings: [],
    partial: false,
  };
  const i = current.sheets.findIndex(
    (s) => s.sheet && (s.sheet.name === unit.name || s.sheet.gid === unit.gid)
  );
  const entry = {
    sheet: {
      name: unit.name,
      gid: unit.gid,
      index: unit.index,
      hidden: unit.hidden,
      frozen: unit.frozen,
      range: unit.a1,
    },
    header_rows: unit.headerRows,
    columns: unit.columns,
    files: filesRef,
  };
  if (i >= 0) current.sheets[i] = entry;
  else current.sheets.push(entry);
  current._export.exportedAt = new Date().toISOString();
  writeJsonFileReplacing_(folder, name, current);
}

/** ===== Apps Script export (All Sheets only) ===== */
function exportAppsScript_(parentFolder) {
  const out = { scriptId: ScriptApp.getScriptId(), files: [], warnings: [] };
  try {
    if (
      typeof Script === "undefined" ||
      !Script.Projects ||
      !Script.Projects.getContent
    ) {
      out.warnings.push('Enable Advanced service "Apps Script API" + GCP API.');
      return out;
    }
    const content = Script.Projects.getContent(out.scriptId);
    if (!content || !content.files || !content.files.length) {
      out.warnings.push("No Apps Script files found.");
      return out;
    }
    const folder = getOrCreateCleanSubfolder_(parentFolder, "apps-script");
    for (const f of content.files) {
      const { filename, mimeType, data } = encodeAppsScriptFile_(f);
      writeTextFileReplacing_(folder, filename, data, mimeType);
      out.files.push({ name: filename, type: f.type });
    }
  } catch (e) {
    out.warnings.push(e && e.message ? e.message : String(e));
  }
  return out;
}
function encodeAppsScriptFile_(f) {
  const name = f.name || "Code";
  const type = f.type || "SERVER_JS";
  let ext = "gs",
    mime = "text/plain",
    data = f.source || "";
  if (type === "HTML") {
    ext = "html";
    mime = "text/html";
  } else if (type === "JSON") {
    ext = "json";
    mime = "application/json";
    try {
      data = JSON.stringify(JSON.parse(data), null, 2);
    } catch (_) {}
  }
  return { filename: `${sanitize_(name)}.${ext}`, mimeType: mime, data };
}
function getOrCreateCleanSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : parent.createFolder(name);
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      file.setTrashed(true);
    } catch (_) {
      try {
        folder.removeFile(file);
      } catch (__) {}
    }
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    try {
      sub.setTrashed(true);
    } catch (_) {
      try {
        folder.removeFolder(sub);
      } catch (__) {}
    }
  }
  return folder;
}

/** ===== Collectors & builders ===== */
function collectSpreadsheetMeta_(ss) {
  let ownerEmail = null;
  try {
    ownerEmail = ss.getOwner()?.getEmail() || null;
  } catch (_) {}
  return {
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl(),
    timeZone: ss.getSpreadsheetTimeZone(),
    locale: ss.getSpreadsheetLocale(),
    authorEmail: ownerEmail,
  };
}
function collectSheetCore_(sheet, cfg) {
  const name = sheet.getName();
  const fileBase = sanitize_(name);
  const lastRow = sheet.getLastRow(),
    lastCol = sheet.getLastColumn();
  const hasData = lastRow > 0 && lastCol > 0;

  let values = [],
    formulas = [];
  if (hasData) {
    const r = sheet.getRange(1, 1, lastRow, lastCol);
    values = r.getValues();
    formulas = r.getFormulas();
  }

  const bounds = cfg.trimTrailingEmpty
    ? detectLastNonEmptyBounds_(values, formulas)
    : { rows: values.length, cols: (values[0] || []).length };

  const rows = bounds.rows,
    cols = bounds.cols;
  const V = slice2D_(values, rows, cols);
  const F = slice2D_(formulas, rows, cols);

  const headerRows = inferHeaderRows_(sheet, V);
  const columns = buildColumnsSchema_({
    values: V,
    formulas: F,
    headerRows,
    maxExamples: cfg.maxExamplesPerColumn,
  });
  const a1 = rows && cols ? toA1_(1, 1) + ":" + toA1_(rows, cols) : null;

  return {
    name,
    fileBase,
    gid: sheet.getSheetId(),
    index: sheet.getIndex(),
    hidden: sheet.isSheetHidden(),
    frozen: { rows: sheet.getFrozenRows(), cols: sheet.getFrozenColumns() },
    rows,
    cols,
    a1,
    headerRows,
    columns,
    V,
    F,
  };
}
function sheetSchemaPayload_(u) {
  return {
    sheet: {
      name: u.name,
      gid: u.gid,
      index: u.index,
      hidden: u.hidden,
      frozen: u.frozen,
    },
    range: { rows: u.rows, cols: u.cols, a1: u.a1 },
    header_rows: u.headerRows,
    columns: u.columns,
  };
}
function sheetDataPayloadAll_(u, cfg) {
  return {
    sheet: u.name,
    gid: u.gid,
    range: { rows: u.rows, cols: u.cols, a1: u.a1 },
    header_rows: u.headerRows,
    cells: buildCellsDataFrom_(u.V, u.F, 0, u.rows, cfg.dateAsISO),
  };
}
function writeSheetChunksFromCore_(folder, u, cfg, t0, index) {
  const names = [];
  const rows = u.rows,
    cols = u.cols;
  if (rows <= 0 || cols <= 0) return names;
  const chunk = Math.max(1, Math.min(cfg.chunkRowsTarget, rows));
  let part = 1;
  for (let start = 0; start < rows; start += chunk) {
    if (Date.now() - t0 > cfg.maxRuntimeMs) {
      index.partial = true;
      break;
    }
    const end = Math.min(rows, start + chunk);
    const obj = {
      sheet: u.name,
      gid: u.gid,
      range: {
        rows: end - start,
        cols,
        rowStart: start + 1,
        rowEnd: end,
        a1: toA1_(start + 1, 1) + ":" + toA1_(end, cols),
      },
      header_rows: u.headerRows,
      cells: buildCellsDataFrom_(u.V, u.F, start, end, true),
    };
    const fname = `${u.fileBase}.rows.${pad4_(part)}.json`;
    writeJsonFileReplacing_(folder, fname, obj);
    names.push(fname);
    part++;
  }
  return names;
}
function buildCellsDataFrom_(V, F, start, end, dateAsISO) {
  const rows = end - start,
    cols = V[0] ? V[0].length : 0;
  const out = new Array(rows);
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) {
      const raw = V[start + r][c],
        f = F[start + r][c];
      const t = cellType_(raw);
      const v = raw instanceof Date && dateAsISO ? toISO_(raw) : raw;
      row[c] =
        f && f.trim()
          ? { v: v === "" ? null : v, t, f }
          : { v: v === "" ? null : v, t };
    }
    out[r] = row;
  }
  return out;
}
function buildColumnsSchema_({ values, formulas, headerRows, maxExamples }) {
  const rows = values.length,
    cols = rows ? values[0].length : 0;
  const out = [];
  for (let c = 0; c < cols; c++) {
    const header = [];
    for (let hr = 0; hr < headerRows && hr < rows; hr++)
      header.push(str_(values[hr][c]));
    let blanks = 0,
      formulaCount = 0;
    const examples = [],
      colVals = [];
    for (let r = headerRows; r < rows; r++) {
      const v = values[r][c],
        f = formulas[r][c];
      colVals.push(v);
      if (isEmpty_(v) && !f) blanks++;
      if (f && f.trim()) formulaCount++;
      if (examples.length < maxExamples && !isEmpty_(v))
        examples.push(example_(v));
    }
    out.push({
      index: c + 1,
      letter: toColLetter_(c + 1),
      header: header.length ? header : null,
      inferred_type: inferType_(colVals),
      counts: {
        total_rows: Math.max(0, rows - headerRows),
        blanks,
        formulas: formulaCount,
      },
      examples: examples.length ? examples : null,
    });
  }
  return out;
}

/** ===== Bounds, types, utils ===== */
function detectLastNonEmptyBounds_(values, formulas) {
  const rCount = values.length,
    cCount = rCount ? values[0].length : 0;
  let lastRow = 0,
    lastCol = 0;
  for (let r = 0; r < rCount; r++)
    for (let c = 0; c < cCount; c++) {
      const v = values[r][c],
        f = formulas[r][c];
      if (!isEmpty_(v) || (f && f.trim())) {
        if (r + 1 > lastRow) lastRow = r + 1;
        if (c + 1 > lastCol) lastCol = c + 1;
      }
    }
  return { rows: lastRow, cols: lastCol };
}
function inferHeaderRows_(sheet, values) {
  const frozen = sheet.getFrozenRows();
  if (frozen && frozen > 0) return frozen;
  if (!values.length) return 0;
  const row = values[0],
    n = row.length || 0;
  let strCount = 0;
  for (let c = 0; c < n; c++) if (cellType_(row[c]) === "string") strCount++;
  return strCount / (n || 1) >= 0.6 ? 1 : 0;
}
function isEmpty_(v) {
  return v === "" || v === null || v === undefined;
}
function cellType_(v) {
  if (isEmpty_(v)) return "blank";
  if (v instanceof Date) return "date";
  const t = typeof v;
  if (t === "number") return Number.isInteger(v) ? "integer" : "number";
  if (t === "boolean") return "boolean";
  return "string";
}
function inferType_(colVals) {
  const counts = {
    integer: 0,
    number: 0,
    boolean: 0,
    date: 0,
    string: 0,
    blank: 0,
  };
  for (const v of colVals)
    counts[cellType_(v)] = (counts[cellType_(v)] || 0) + 1;
  const nonBlank = colVals.length - (counts.blank || 0);
  if (!nonBlank) return "unknown";
  const f = (k) => (counts[k] || 0) / nonBlank;
  if (f("date") >= 0.5) return "date";
  if (f("integer") + f("number") >= 0.5)
    return counts.number ? "number" : "integer";
  if (f("boolean") >= 0.5) return "boolean";
  if (f("string") >= 0.5) return "string";
  return "mixed";
}
function example_(v) {
  return v instanceof Date ? toISO_(v) : v;
}
function str_(v) {
  return v instanceof Date ? toISO_(v) : v == null ? "" : String(v);
}
function toISO_(d) {
  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}
function toColLetter_(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function toA1_(row, col) {
  return toColLetter_(col) + row;
}
function slice2D_(arr, rows, cols) {
  const out = new Array(rows);
  for (let r = 0; r < rows; r++) {
    const row = arr[r] || [];
    out[r] = row.slice(0, cols);
  }
  return out;
}
function sanitize_(s) {
  return s.replace(/[\\/:*?"<>|]/g, "_").trim();
}
function pad4_(n) {
  return String(n).padStart(4, "0");
}
function escapeRegExp_(s) {
  return s.replace(/[.*+?^${}()|[\\]\\]/g, "\\$&");
}
function safeEmail_() {
  try {
    return Session.getActiveUser() && Session.getActiveUser().getEmail
      ? Session.getActiveUser().getEmail()
      : null;
  } catch (_) {
    try {
      return Session.getEffectiveUser && Session.getEffectiveUser().getEmail
        ? Session.getEffectiveUser().getEmail()
        : null;
    } catch (__) {
      return null;
    }
  }
}

/** ===== Drive I/O & folders ===== */
function getFixedTargetFolder_() {
  const ids = [TARGET_FOLDER_ID, TARGET_FOLDER_ID_FALLBACK].filter(Boolean);
  const errors = [];
  for (const id of ids) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      errors.push(id + ": " + (e && e.message ? e.message : e));
    }
  }
  const who = safeEmail_();
  throw new Error(
    "Target folder not found or no access. Tried: " +
      ids.join(", ") +
      (who ? " | User: " + who : "") +
      (errors.length ? " | Errors: " + errors.join(" || ") : "")
  );
}
function getOrCreateSpreadsheetFolderUnderTarget_(targetFolder, ss) {
  const name = sanitize_(ss.getName());
  const it = targetFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : targetFolder.createFolder(name);
}
function purgeExportArtifacts_(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const n = file.getName();
    if (/\.(json|zip)$/i.test(n)) {
      try {
        file.setTrashed(true);
      } catch (_) {
        try {
          folder.removeFile(file);
        } catch (__) {}
      }
    }
  }
}
function removeSheetArtifacts_(folder, fileBase) {
  const re = new RegExp(
    "^" +
      escapeRegExp_(fileBase) +
      "\\.(schema\\.json|data\\.json|rows\\.[0-9]{4}\\.json)$",
    "i"
  );
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (re.test(f.getName())) {
      try {
        f.setTrashed(true);
      } catch (_) {
        try {
          folder.removeFile(f);
        } catch (__) {}
      }
    }
  }
}
function writeJsonFileReplacing_(folder, name, obj) {
  const cur = folder.getFilesByName(name);
  while (cur.hasNext()) {
    const f = cur.next();
    try {
      f.setTrashed(true);
    } catch (_) {
      try {
        folder.removeFile(f);
      } catch (__) {}
    }
  }
  const blob = Utilities.newBlob(
    JSON.stringify(obj, null, 2),
    "application/json",
    name
  );
  return folder.createFile(blob);
}
function writeTextFileReplacing_(folder, name, content, mime) {
  const cur = folder.getFilesByName(name);
  while (cur.hasNext()) {
    const f = cur.next();
    try {
      f.setTrashed(true);
    } catch (_) {
      try {
        folder.removeFile(f);
      } catch (__) {}
    }
  }
  const blob = Utilities.newBlob(content, mime || "text/plain", name);
  return folder.createFile(blob);
}
function readJsonFile_(folder, name) {
  const it = folder.getFilesByName(name);
  if (!it.hasNext()) return null;
  try {
    return JSON.parse(it.next().getBlob().getDataAsString());
  } catch (_) {
    return null;
  }
}

/** ===== UI dialog result (All Sheets only) ===== */
function showResultDialog_(res) {
  const ui = SpreadsheetApp.getUi();
  const lines = [];
  if (res.folderUrl) lines.push(`Folder: ${res.folderUrl}`);
  if (res.indexFileUrl) lines.push(`Index: ${res.indexFileUrl}`);
  if (res.sheetCount) lines.push(`Sheets: ${res.sheetCount}`);
  if (res.partial) lines.push(`⚠ Partial export (time budget)`);
  if (res.warnings && res.warnings.length)
    lines.push(`Warnings:\n- ${res.warnings.join("\n- ")}`);
  ui.alert("GPT Export", lines.join("\n\n"), ui.ButtonSet.OK);
}

// Expose / override library API with this local code to avoid stale external versions
var ExportGSS = {
  exportAllSheets: exportAllSheets,
  exportSchemaOnlyAll: exportSchemaOnlyAll,
  openExportSheetDialog: openExportSheetDialog,
  exportSingleSheet: exportSingleSheet,
  refreshSheetList: refreshSheetList,
  getSheetList: getSheetList,
  openSpreadsheetFolder: openSpreadsheetFolder,
};
var exportgoogleSS = ExportGSS;
