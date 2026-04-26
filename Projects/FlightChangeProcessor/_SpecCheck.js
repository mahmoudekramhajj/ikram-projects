/**
 * فحص خلية محددة + كشف كل URL مخفي في أعمدة الأوقات
 */
function checkCell(args) {
  args = args || {};
  var row = args.row || 2;
  var col = args.col || 'AK';

  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var range = sh.getRange(col + row);
  var value = range.getValue();
  var display = range.getDisplayValue();
  var formula = range.getFormula();
  var note = range.getNote();

  return {
    cell: col + row,
    rawValue: String(value),
    rawType: typeof value,
    isDate: value instanceof Date,
    displayValue: display,
    formula: formula,
    note: note,
    length: String(value).length
  };
}


/**
 * يبحث عن نص معين في كامل الشيت
 */
/**
 * يقرأ صف كامل ويعرض كل خلاياه
 */
function readFullRow(args) {
  args = args || {};
  var row = args.row || 2;
  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  var dispVals = sh.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  var result = [];
  for (var i = 0; i < vals.length; i++) {
    result.push({
      col: colLetter_(i),
      header: String(headers[i] || ''),
      value: String(vals[i]).substring(0, 100),
      display: String(dispVals[i]).substring(0, 100),
      type: vals[i] instanceof Date ? 'Date' : typeof vals[i]
    });
  }
  return result;
}


function searchAnyText(args) {
  args = args || {};
  var q = (args.q || '').toLowerCase();
  if (!q) return { error: 'ضع q' };

  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  var hits = [];
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var v = String(data[r][c] || '').toLowerCase();
      if (v.indexOf(q) !== -1) {
        hits.push({
          row: r + 1,
          col: colLetter_(c),
          header: String(headers[c] || ''),
          value: String(data[r][c]).substring(0, 120)
        });
      }
    }
  }
  return { query: q, hits: hits };
}


/**
 * يبحث عن أي ذكر لـ mail.google.com في الشيت كله
 */
function findGmailUrlsAnywhere() {
  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  var hits = [];
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var v = String(data[r][c] || '');
      if (v.indexOf('mail.google.com') !== -1) {
        hits.push({
          row: r + 1,
          col: colLetter_(c),
          header: String(headers[c] || ''),
          value: v.substring(0, 120)
        });
      }
    }
  }

  return { totalHits: hits.length, hits: hits };
}


/**
 * يبحث عن أي خلية تحوي URL في أعمدة يجب أن تكون time/date فقط
 */
function findUrlsInWrongCols() {
  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  var timeCols = [];
  var dateCols = [];
  var flightCols = [];
  var airportCols = [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase();
    if (/time/i.test(h) || /وقت/.test(h)) timeCols.push({ index: i, name: headers[i] });
    else if (/^date /i.test(h) || /تاريخ/.test(h)) dateCols.push({ index: i, name: headers[i] });
    else if (/flight|رحلة/.test(h)) flightCols.push({ index: i, name: headers[i] });
    else if (/^to \d|^from \d|مطار/.test(h)) airportCols.push({ index: i, name: headers[i] });
  }

  var issues = [];
  var allSuspectCols = [].concat(timeCols, dateCols, flightCols, airportCols);

  for (var r = 1; r < data.length; r++) {
    for (var c = 0; c < allSuspectCols.length; c++) {
      var colIdx = allSuspectCols[c].index;
      var val = data[r][colIdx];
      var sv = String(val || '');
      if (/^https?:\/\//.test(sv)) {
        issues.push({
          row: r + 1,
          col: colLetter_(colIdx),
          header: allSuspectCols[c].name,
          value: sv.substring(0, 100)
        });
      }
    }
  }

  return {
    totalChecked: data.length - 1,
    suspectColumns: allSuspectCols.length,
    issuesFound: issues.length,
    issues: issues.slice(0, 50)
  };
}
