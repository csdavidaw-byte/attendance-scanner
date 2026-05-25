// ============================================================
//  Attendance Scanner — Google Apps Script
//  Paste this entire file into your Apps Script editor
//  in the spreadsheet: https://docs.google.com/spreadsheets/d/113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo
// ============================================================

// Your specific Google Sheet ID (already set)
var SPREADSHEET_ID = "113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo";

// Name of the sheet tab to write attendance into
var SHEET_NAME = "Attendance";

// ============================================================
//  DO NOT EDIT BELOW THIS LINE
// ============================================================

function doPost(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = JSON.parse(e.postData.contents);

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Name", "Phone", "Session", "Location", "Date", "Time", "Logged At"]);
      var headerRange = sheet.getRange(1, 1, 1, 7);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1a1916");
      headerRange.setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // Append the attendance record
    sheet.appendRow([
      data.name     || "",
      data.phone    || "",
      data.session  || "",
      data.location || "",
      data.date     || "",
      data.time     || "",
      new Date().toLocaleString("en-SG")
    ]);

    // Auto-resize columns
    sheet.autoResizeColumns(1, 7);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", message: "Logged: " + data.name }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handles GET requests — open this URL in browser to test
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Attendance Script is running! Sheet ID: " + SPREADSHEET_ID }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Gets or creates the Attendance sheet tab in YOUR specific spreadsheet
function getOrCreateSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}
