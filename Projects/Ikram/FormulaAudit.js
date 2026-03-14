function auditAllFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function(s) { return s.getName(); });
  
  var results = [['Sheet #', 'Sheet Name', 'Cell', 'Issue', 'Formula']];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var range = sheet.getDataRange();
    var formulas = range.getFormulas();
    var values = range.getDisplayValues();
    
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        var f = formulas[r][c];
        if (f === '') continue;
        
        var cell = range.getCell(r + 1, c + 1).getA1Notation();
        var val = values[r][c];
        
        // صيغ تحتوي #REF
        if (f.indexOf('#REF!') !== -1 || val === '#REF!') {
          results.push([i + 1, sheet.getName(), cell, '#REF! Error', f]);
        }
        // صيغ تحتوي أخطاء أخرى
        else if (['#ERROR!', '#N/A', '#NAME?', '#VALUE!', '#DIV/0!'].indexOf(val) !== -1) {
          results.push([i + 1, sheet.getName(), cell, val, f]);
        }
        // صيغ طويلة جداً
        else if (f.length > 500) {
          results.push([i + 1, sheet.getName(), cell, 'Long Formula (' + f.length + ' chars)', f.substring(0, 200) + '...']);
        }
      }
    }
  }
  
  var report = ss.getSheetByName('_FormulaAudit');
  if (!report) report = ss.insertSheet('_FormulaAudit');
  report.clearContents();
  if (results.length > 1) {
    report.getRange(1, 1, results.length, 5).setValues(results);
  }
  
  Logger.log('تم العثور على ' + (results.length - 1) + ' صيغة مشكلة');
}