/**
 * Scopus DOI Lookup Tool for Google Sheets
 * ------------------------------------------
 * Enter DOIs in column A; metadata is fetched from Elsevier's Scopus API.
 *
 * Replace the ENTIRE contents of your Apps Script file with this file
 * (don't paste alongside old code) — leftover duplicate onOpen()
 * functions are the usual cause of a doubled-up menu.
 */

var API_KEY = "094956ef75833c1369eccac0bf2555f0"; // Your Elsevier API key

var FIELDS = [
  { header: "Status" },
  { header: "Type",         key: "subtypeDescription" },
  { header: "Journal",      key: "prism:publicationName" },
  { header: "Title",        key: "dc:title" },
  { header: "Pub Date",     key: "prism:coverDate" },
  { header: "Citations",    key: "citedby-count", numeric: true },
  { header: "First Author", key: "dc:creator" },
  { header: "ISSN",         key: "prism:issn" },
  { header: "Open Access",  key: "openaccessFlag", flag: true },
  { header: "Scopus EID",   key: "eid" }
];

var DATA_START_COL = 2;      // Column B (Column A is the DOI)
var DATA_COL_COUNT = FIELDS.length;
var BATCH_SIZE = 20;         // chunk size for UrlFetchApp.fetchAll — keeps request bursts polite
var TRIGGER_HANDLER = 'onSheetEdit';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Scopus Tool')
    .addItem('🔄 Run Manual Batch (All Rows)', 'processScopusData')
    .addToUi();

  ensureAutoFetchTrigger_();
}

/**
 * Silently makes sure the live auto-fetch trigger exists. This is a no-op
 * (not an error) the very first time the sheet is opened, before the script
 * has been authorized — it will succeed automatically on the next onOpen()
 * after the user has authorized the script via any menu action.
 */
function ensureAutoFetchTrigger_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var alreadyInstalled = triggers.some(function (t) {
      return t.getHandlerFunction() === TRIGGER_HANDLER;
    });
    if (!alreadyInstalled) {
      ScriptApp.newTrigger(TRIGGER_HANDLER)
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onEdit()
        .create();
    }
  } catch (err) {
    // Not authorized yet — resolves itself after the first authorized run.
  }
}

/** Handles real-time cell edits in Column A. */
function onSheetEdit(e) {
  if (!e) return;
  var range = e.range;
  var sheet = range.getSheet();
  if (range.getColumn() !== 1 || range.getRow() < 2) return;

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return; // another edit is already being processed, skip rather than collide

  try {
    ensureHeaders(sheet);
    var startRow = range.getRow();
    var values = range.getValues();
    var dois = [];
    var rowNums = [];

    for (var i = 0; i < values.length; i++) {
      var doi = values[i][0];
      if (doi && doi.toString().trim() !== "") {
        dois.push(doi.toString());
        rowNums.push(startRow + i);
      } else {
        sheet.getRange(startRow + i, DATA_START_COL, 1, DATA_COL_COUNT).clearContent();
      }
    }
    if (dois.length === 0) return;

    var results = fetchScopusMetadataBatch(dois);
    for (var j = 0; j < rowNums.length; j++) {
      sheet.getRange(rowNums[j], DATA_START_COL, 1, DATA_COL_COUNT).setValues([results[j]]);
    }
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(err.message, 'Scopus Tool Error', 8);
  } finally {
    lock.releaseLock();
  }
}

/** Batch processes all existing DOIs in the sheet, fetched in parallel chunks. */
function processScopusData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  ensureHeaders(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Please add DOIs in Column A starting from row 2.");
    return;
  }

  var dois = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function (r) { return r[0]; });
  var outputData = new Array(dois.length);

  for (var start = 0; start < dois.length; start += BATCH_SIZE) {
    var chunk = dois.slice(start, start + BATCH_SIZE);
    var chunkResults = fetchScopusMetadataBatch(chunk);
    for (var i = 0; i < chunkResults.length; i++) {
      outputData[start + i] = chunkResults[i];
    }
    ss.toast('Processed ' + Math.min(start + BATCH_SIZE, dois.length) + ' / ' + dois.length + ' rows', 'Scopus Tool', 3);
  }

  sheet.getRange(2, DATA_START_COL, outputData.length, DATA_COL_COUNT).setValues(outputData);
  ui.alert('Done. Processed ' + outputData.length + ' rows.');
}

function ensureHeaders(sheet) {
  if (sheet.getRange("A1").getValue() !== "DOI") {
    var headers = ["DOI"].concat(FIELDS.map(function (f) { return f.header; }));
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#e8f0fe");
  }
}

/**
 * Fetches metadata for a batch of DOIs in parallel using UrlFetchApp.fetchAll,
 * which is far faster than calling UrlFetchApp.fetch once per row in a loop.
 */
function fetchScopusMetadataBatch(dois) {
  var validIndexes = [];
  var validRequests = [];

  dois.forEach(function (doi, idx) {
    var clean = cleanDoi_(doi);
    if (!clean) return; // leave blank rows out of the request batch entirely
    validIndexes.push(idx);
    validRequests.push({
      url: "https://api.elsevier.com/content/search/scopus?query=" + encodeURIComponent('doi(' + clean + ')'),
      method: "get",
      headers: { "X-ELS-APIKey": API_KEY, "Accept": "application/json" },
      muteHttpExceptions: true
    });
  });

  var responses = validRequests.length ? UrlFetchApp.fetchAll(validRequests) : [];
  var out = dois.map(function () { return emptyRow_(); });

  responses.forEach(function (response, i) {
    out[validIndexes[i]] = parseScopusResponse_(response);
  });

  return out;
}

function cleanDoi_(doi) {
  if (!doi) return "";
  return doi.toString()
    .replace(/^https?:\/\/(dx\.)?doi\.org\//i, "")
    .replace(/^doi:\s*/i, "")
    .trim();
}

function emptyRow_() {
  return FIELDS.map(function () { return ""; });
}

function statusRow_(status) {
  return [status].concat(FIELDS.slice(1).map(function () { return "-"; }));
}

function parseScopusResponse_(response) {
  var code = response.getResponseCode();
  if (code === 429) return statusRow_("Rate limited");
  if (code !== 200) return statusRow_("Error (" + code + ")");

  var data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {
    return statusRow_("Parse error");
  }

  var results = data && data["search-results"];
  var total = results ? parseInt(results["opensearch:totalResults"] || 0, 10) : 0;
  if (!total) return statusRow_("Not Indexed");

  var entry = results.entry[0];
  var row = ["Indexed"];
  FIELDS.slice(1).forEach(function (f) {
    if (f.flag) {
      row.push(entry[f.key] === "1" || entry[f.key] === true ? "Yes" : "No");
    } else if (f.numeric) {
      row.push(entry[f.key] !== undefined ? parseInt(entry[f.key], 10) : 0);
    } else {
      row.push(entry[f.key] || "N/A");
    }
  });
  return row;
}
