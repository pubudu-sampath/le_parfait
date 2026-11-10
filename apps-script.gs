/**
 * Le Parfait - Equipment Sites Log
 * Google Apps Script backend for index.html
 *
 * Serves the web page:
 *   GET  ?action=logs      -> { success: true, rows: [ {state, site, equipmentName, equipmentCode, timestamp}, ... ] }
 *   GET  ?action=sites     -> { success: true, sites: ["Kealba Hotel", ...] }
 *   GET  ?action=sitelist  -> { success: true, sites: [ {id, site, state, supervisor, suburb, operating, remarks}, ... ] }
 *   POST {state, site, equipmentName, equipmentCode, timestamp} -> appends one equipment row
 *   POST {type:"site", op:"create"|"update"|"delete", ...}      -> maintains the Sites tab
 *
 * SETUP
 *   1. Put the new spreadsheet's ID in SHEET_ID below.
 *   2. Run setupSheet() once from the editor to create the tab and header row.
 *   3. Deploy > New deployment > Web app
 *        Execute as:      Me
 *        Who has access:  Anyone          <-- both of these matter, see README note
 *   4. Copy the /exec URL into SCRIPT_URL in index.html.
 */

// The long id in the sheet's URL:
// https://docs.google.com/spreadsheets/d/THIS_PART_HERE/edit
var SHEET_ID = "PASTE_NEW_SPREADSHEET_ID_HERE";

var SHEET_NAME = "Equipment Log";

// Second tab in the SAME spreadsheet - one openById call serves both.
var SITES_SHEET_NAME = "Sites";

var SITE_HEADERS = [
  "ID",
  "Site Name",
  "State",
  "Supervisor",
  "Suburb",
  "Still Operating",
  "Remarks"
];

// Maps a Sites header cell to the key the web page uses.
var SITE_FIELD_BY_HEADER = {
  "id": "id",
  "site name": "site",
  "site": "site",
  "state": "state",
  "supervisor": "supervisor",
  "suburb": "suburb",
  "still operating": "operating",
  "operating": "operating",
  "remarks": "remarks"
};

// Header row. Order can be changed in the sheet later - the script maps by
// name, not by column position.
var HEADERS = [
  "Timestamp",
  "State",
  "Site",
  "Equipment Name",
  "Equipment Code"
];

// Maps a header cell to the key the web page uses.
var FIELD_BY_HEADER = {
  "timestamp": "timestamp",
  "state": "state",
  "site": "site",
  "site name": "site",
  "equipment name": "equipmentName",
  "equipment code": "equipmentCode"
};

/* ------------------------------------------------------------------ */
/* Entry points                                                        */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || "logs";

    if (action === "sites") {
      return json({ success: true, sites: readSites() });
    }

    if (action === "sitelist") {
      return json({ success: true, sites: readSiteList() });
    }

    return json({ success: true, rows: readRows() });
  } catch (err) {
    return json({ success: false, message: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Empty request body");
    }

    var payload = JSON.parse(e.postData.contents);

    // Site records go through the same endpoint, tagged by type.
    if (payload && payload.type === "site") {
      lock.waitLock(20000);
      return json(handleSite(payload));
    }

    var entry = {
      state: trimmed(payload.state),
      site: trimmed(payload.site),
      equipmentName: trimmed(payload.equipmentName),
      equipmentCode: trimmed(payload.equipmentCode).toLowerCase(),
      timestamp: payload.timestamp ? new Date(payload.timestamp) : new Date()
    };

    if (!entry.site || !entry.equipmentCode) {
      throw new Error("site and equipmentCode are required");
    }

    if (isNaN(entry.timestamp.getTime())) {
      entry.timestamp = new Date();
    }

    // Two people can scan at the same moment; serialise the append so
    // neither row is lost.
    lock.waitLock(20000);

    var sheet = getSheet();
    var index = headerIndex(sheet);
    var row = [];

    for (var i = 0; i < index.headers.length; i++) {
      var field = FIELD_BY_HEADER[index.headers[i]];
      row.push(field ? entry[field] : "");
    }

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return json({ success: true });
  } catch (err) {
    return json({ success: false, message: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

/* ------------------------------------------------------------------ */
/* Reads                                                               */
/* ------------------------------------------------------------------ */

function readRows() {
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  var headers = values[0].map(normalise);
  var rows = [];

  for (var r = 1; r < values.length; r++) {
    var row = { state: "", site: "", equipmentName: "", equipmentCode: "", timestamp: "" };
    var hasData = false;

    for (var c = 0; c < headers.length; c++) {
      var field = FIELD_BY_HEADER[headers[c]];
      if (!field) continue;

      var value = values[r][c];
      if (value === "" || value === null) continue;

      hasData = true;

      if (field === "timestamp") {
        row.timestamp = (value instanceof Date)
          ? value.toISOString()
          : String(value);
      } else if (field === "equipmentCode") {
        row.equipmentCode = String(value).trim().toLowerCase();
      } else {
        row[field] = String(value).trim();
      }
    }

    if (hasData && row.equipmentCode) rows.push(row);
  }

  return rows;
}

function readSites() {
  var seen = {};
  var sites = [];
  var rows = readRows();

  for (var i = 0; i < rows.length; i++) {
    var site = rows[i].site;
    var key = site.toLowerCase();

    if (site && !seen[key]) {
      seen[key] = true;
      sites.push(site);
    }
  }

  sites.sort(function (a, b) { return a.localeCompare(b); });
  return sites;
}

/* ------------------------------------------------------------------ */
/* Sites tab                                                           */
/* ------------------------------------------------------------------ */

function readSiteList() {
  var sheet = getSitesSheet();
  var values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  var headers = values[0].map(normalise);
  var out = [];

  for (var r = 1; r < values.length; r++) {
    var row = {
      id: "", site: "", state: "", supervisor: "",
      suburb: "", operating: "", remarks: ""
    };

    for (var c = 0; c < headers.length; c++) {
      var field = SITE_FIELD_BY_HEADER[headers[c]];
      if (!field) continue;

      var value = values[r][c];
      if (value === "" || value === null) continue;

      if (field === "operating") {
        // Accept a checkbox, "yes"/"no", or TRUE/FALSE.
        var text = String(value).trim().toLowerCase();
        row.operating = (text === "true" || text === "yes" || value === true)
          ? "Yes"
          : "No";
      } else {
        row[field] = String(value).trim();
      }
    }

    // A row with no site name is blank or a stray note - skip it.
    if (row.site) {
      if (!row.operating) row.operating = "Yes";
      out.push(row);
    }
  }

  return out;
}

/** create / update / delete one site. Returns the response body. */
function handleSite(payload) {
  var op = String(payload.op || "create").toLowerCase();
  var sheet = getSitesSheet();

  if (op === "delete") {
    var rowNumber = findSiteRow(sheet, payload.id);
    if (rowNumber < 0) throw new Error("That site no longer exists");

    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { success: true };
  }

  if (op === "bulk") {
    return bulkAddSites(sheet, payload.rows);
  }

  var record = {
    id: trimmed(payload.id),
    site: trimmed(payload.site),
    state: trimmed(payload.state),
    supervisor: trimmed(payload.supervisor),
    suburb: trimmed(payload.suburb),
    operating: String(payload.operating).toLowerCase() === "no" ? "No" : "Yes",
    remarks: trimmed(payload.remarks)
  };

  if (!record.site) throw new Error("Site name is required");

  var headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getValues()[0]
    .map(normalise);

  if (op === "update") {
    var target = findSiteRow(sheet, record.id);
    if (target < 0) throw new Error("That site no longer exists");

    for (var c = 0; c < headers.length; c++) {
      var field = SITE_FIELD_BY_HEADER[headers[c]];
      if (!field || field === "id") continue;
      sheet.getRange(target, c + 1).setValue(record[field]);
    }

    SpreadsheetApp.flush();
    return { success: true, id: record.id };
  }

  record.id = Utilities.getUuid();

  var line = [];
  for (var i = 0; i < headers.length; i++) {
    var key = SITE_FIELD_BY_HEADER[headers[i]];
    line.push(key ? record[key] : "");
  }

  sheet.appendRow(line);
  SpreadsheetApp.flush();
  return { success: true, id: record.id };
}

/**
 * Appends many sites in a single write. Rows whose site name is blank or
 * already present are skipped rather than failing the whole import.
 */
function bulkAddSites(sheet, rows) {
  if (!rows || !rows.length) throw new Error("No rows to import");
  if (rows.length > 2000) throw new Error("Too many rows in one import (max 2000)");

  var headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getValues()[0]
    .map(normalise);

  // Existing names, so a re-run of the same file doesn't duplicate everything.
  var taken = {};
  var existing = readSiteList();
  for (var e = 0; e < existing.length; e++) {
    taken[existing[e].site.toLowerCase()] = true;
  }

  var lines = [];
  var skippedBlank = 0;
  var skippedDuplicate = 0;

  for (var i = 0; i < rows.length; i++) {
    var record = {
      id: Utilities.getUuid(),
      site: trimmed(rows[i].site),
      state: trimmed(rows[i].state).toUpperCase(),
      supervisor: trimmed(rows[i].supervisor),
      suburb: trimmed(rows[i].suburb),
      operating: String(rows[i].operating).toLowerCase() === "no" ? "No" : "Yes",
      remarks: trimmed(rows[i].remarks)
    };

    if (!record.site) { skippedBlank++; continue; }

    var key = record.site.toLowerCase();
    if (taken[key]) { skippedDuplicate++; continue; }
    taken[key] = true;

    var line = [];
    for (var c = 0; c < headers.length; c++) {
      var field = SITE_FIELD_BY_HEADER[headers[c]];
      line.push(field ? record[field] : "");
    }

    lines.push(line);
  }

  if (lines.length) {
    // One write for the whole batch - far cheaper than appendRow per row.
    sheet
      .getRange(sheet.getLastRow() + 1, 1, lines.length, headers.length)
      .setValues(lines);

    SpreadsheetApp.flush();
  }

  return {
    success: true,
    added: lines.length,
    skippedBlank: skippedBlank,
    skippedDuplicate: skippedDuplicate
  };
}

/** 1-based sheet row for a site id, or -1. */
function findSiteRow(sheet, id) {
  var wanted = trimmed(id);
  if (!wanted) return -1;

  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(normalise);
  var idColumn = -1;

  for (var c = 0; c < headers.length; c++) {
    if (SITE_FIELD_BY_HEADER[headers[c]] === "id") { idColumn = c; break; }
  }

  if (idColumn < 0) return -1;

  for (var r = 1; r < values.length; r++) {
    if (trimmed(values[r][idColumn]) === wanted) return r + 1;
  }

  return -1;
}

function getSitesSheet() {
  var book = SpreadsheetApp.openById(SHEET_ID);
  var sheet = book.getSheetByName(SITES_SHEET_NAME);

  // Created on demand so an existing deployment keeps working.
  if (!sheet) {
    sheet = book.insertSheet(SITES_SHEET_NAME);
    sheet.appendRow(SITE_HEADERS);
    sheet.getRange(1, 1, 1, SITE_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function getSheet() {
  if (!SHEET_ID || SHEET_ID.indexOf("PASTE_") === 0) {
    throw new Error("SHEET_ID has not been set in the script");
  }

  var book = SpreadsheetApp.openById(SHEET_ID);
  var sheet = book.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error('No tab named "' + SHEET_NAME + '" in that spreadsheet');
  }

  return sheet;
}

function headerIndex(sheet) {
  var headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getValues()[0]
    .map(normalise);

  return { headers: headers };
}

function normalise(value) {
  return String(value == null ? "" : value).trim().toLowerCase();
}

function trimmed(value) {
  return String(value == null ? "" : value).trim();
}

function json(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Run these from the editor                                           */
/* ------------------------------------------------------------------ */

/** Run once: creates the tab and the header row if they are missing. */
function setupSheet() {
  var book = SpreadsheetApp.openById(SHEET_ID);
  var sheet = book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  getSitesSheet();

  Logger.log('Ready: "%s" and "%s" in "%s"',
    SHEET_NAME, SITES_SHEET_NAME, book.getName());
}

/** Optional: confirms reads work before you wire up the web page. */
function testRead() {
  Logger.log("rows:  %s", JSON.stringify(readRows()));
  Logger.log("sites: %s", JSON.stringify(readSites()));
}
