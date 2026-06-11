// ============================================================
//  Attendance Scanner — Google Apps Script
//  Sheet: https://docs.google.com/spreadsheets/d/113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo
//
//  DEPLOY SETTINGS (very important):
//  - Execute as: Me
//  - Who has access: Anyone (even anonymous)
// ============================================================

var SPREADSHEET_ID = "113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo";
var SHEET_GID      = 246908960;
var SHEET_NAME     = "Attendance";

function doGet(e) {
  var output;
  try {
    var p = e.parameter || {};

    if (!p.name) {
      // Test ping
      output = { status: "ok", message: "Script is running! Sheet connected." };
    } else {
      var sheet = getTargetSheet();

      if (sheet.getLastRow() === 0) {
        sheet.appendRow(["Name","Phone","Session","Location","Date","Time","Logged At"]);
        sheet.getRange(1,1,1,7).setFontWeight("bold").setBackground("#1a1916").setFontColor("#ffffff");
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        p.name     || "",
        p.phone    || "",
        p.session  || "",
        p.location || "",
        p.date     || "",
        p.time     || "",
        new Date().toLocaleString("en-SG")
      ]);
      sheet.autoResizeColumns(1,7);
      output = { status: "ok", message: "Logged: " + p.name };
    }
  } catch(err) {
    output = { status: "error", message: err.toString() };
  }

  // Support JSONP callback for script-tag technique
  var callback = (e.parameter || {}).callback;
  var json = JSON.stringify(output);
  var content = callback ? callback + "(" + json + ")" : json;
  return ContentService
    .createTextOutput(content)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function getTargetSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  var byName = ss.getSheetByName(SHEET_NAME);
  if (byName) return byName;
  return ss.insertSheet(SHEET_NAME);
}
