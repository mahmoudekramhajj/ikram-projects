function listAllSheets() {
  var result = {};
  ['MAIN_SPREADSHEET_ID', 'UNIFIED_SPREADSHEET_ID', 'CHANGES_SPREADSHEET_ID', 'COMPARISON_SPREADSHEET_ID'].forEach(function(k) {
    var id = CONFIG[k];
    if (!id) return;
    try {
      var ss = SpreadsheetApp.openById(id);
      result[k] = {
        id: id,
        name: ss.getName(),
        sheets: ss.getSheets().map(function(s) {
          return {
            name: s.getName(),
            rows: s.getLastRow(),
            cols: s.getLastColumn()
          };
        })
      };
    } catch (e) {
      result[k] = { id: id, error: e.message };
    }
  });
  return result;
}


function readCellFromSheet(args) {
  args = args || {};
  var ssId = args.ssId;
  var sheetName = args.sheetName;
  var cell = args.cell;

  var ss = SpreadsheetApp.openById(ssId);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { error: 'sheet not found: ' + sheetName };

  var range = sh.getRange(cell);
  return {
    ssName: ss.getName(),
    sheetName: sheetName,
    cell: cell,
    value: String(range.getValue()),
    display: range.getDisplayValue()
  };
}
