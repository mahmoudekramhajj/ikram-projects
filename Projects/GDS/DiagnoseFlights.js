/**
 * DiagnoseFlights.js — فحص هيكل شيت "الطيران" عندنا
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

function diagnoseFlightsSheet() {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
  var sheet = ss.getSheetByName(GDS2.Config.SHEET_FLIGHTS);
  if (!sheet) return { error: 'sheet not found: ' + GDS2.Config.SHEET_FLIGHTS };

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = {};
  for (var i = 0; i < headers.length; i++) {
    headerMap[i + 1] = { letter: colLetter_(i + 1), name: String(headers[i]) };
  }

  // صف عينة (أول صف بيانات)
  var sample = lastRow >= 2 ? sheet.getRange(2, 1, 1, lastCol).getValues()[0] : [];

  return {
    sheet_name: GDS2.Config.SHEET_FLIGHTS,
    col_count: lastCol,
    row_count: lastRow - 1,
    headers: headerMap,
    sample_row: sample
  };
}

function colLetter_(col) {
  var s = '';
  while (col > 0) {
    var r = (col - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}
