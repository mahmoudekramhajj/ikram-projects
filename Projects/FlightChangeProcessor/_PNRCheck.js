/**
 * _PNRCheck.js — فحص وجود PNR معين في V4
 */

function checkPNRInV4(args) {
  args = args || {};
  var pnr = (args.pnr || '').toUpperCase().trim();
  if (!pnr) return { error: 'ضع PNR' };

  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  if (!sh) return { error: 'sheet not found' };

  var data = sh.getDataRange().getValues();
  var headers = data[0];

  // نحصل على الأعمدة التي فيها PNR (قد تكون عدة أعمدة)
  var pnrCols = [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase();
    if (h.indexOf('pnr') !== -1) pnrCols.push(i);
  }

  var matches = [];
  for (var r = 1; r < data.length; r++) {
    for (var c = 0; c < pnrCols.length; c++) {
      var val = String(data[r][pnrCols[c]] || '').toUpperCase().trim();
      if (val === pnr || val.indexOf(pnr) !== -1) {
        matches.push({
          row: r + 1,
          pnrColumn: headers[pnrCols[c]],
          rowValues: summarizeRow_(headers, data[r])
        });
        break;
      }
    }
  }

  return {
    searchedPNR: pnr,
    foundInRows: matches.length,
    matches: matches
  };
}


function summarizeRow_(headers, row) {
  var keyFields = ['Passenger', 'Name', 'PNR', 'Booking', 'Incident', 'Route', 'Date', 'اسم', 'رقم', 'pnr', 'name'];
  var summary = {};
  for (var i = 0; i < headers.length && i < 20; i++) {
    var h = String(headers[i] || '');
    var v = row[i];
    if (v !== null && v !== undefined && v !== '') {
      summary[h] = String(v).substring(0, 60);
    }
  }
  return summary;
}
