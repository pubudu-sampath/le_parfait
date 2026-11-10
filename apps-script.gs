/**
 * Le Parfait - Equipment Sites Log
 * Google Apps Script backend for index.html
 *
 * Serves three things to the web page:
 *   GET  ?action=logs   -> { success: true, rows: [ {state, site, equipmentName, equipmentCode, timestamp}, ... ] }
 *   GET  ?action=sites  -> { success: true, sites: ["Kealba Hotel", ...] }
 *   POST {state, site, equipmentName, equipmentCode, timestamp} -> appends one row
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

  Logger.log('Ready: "%s" in "%s"', SHEET_NAME, book.getName());
}

/** Optional: confirms reads work before you wire up the web page. */
function testRead() {
  Logger.log("rows:  %s", JSON.stringify(readRows()));
  Logger.log("sites: %s", JSON.stringify(readSites()));
}
