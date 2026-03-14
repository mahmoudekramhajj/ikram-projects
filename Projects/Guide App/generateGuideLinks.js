/**
 * ══════════════════════════════════════════════════════════════════
 * توليد روابط المرشدين — شغّل هذه الدالة مرة واحدة
 * ستنشئ شيت "روابط المرشدين" بكل الأسماء وروابطهم
 * ══════════════════════════════════════════════════════════════════
 * 
 * ⚠️ عدّل DEPLOY_URL أدناه بعد كل Deploy جديد
 */

function generateGuideLinks() {
  var DEPLOY_URL = 'https://script.google.com/macros/s/AKfycbwirA68K9peWAsnAmcYZV3sDU1Bpy9TWQKt9lvtlF0kDGcvtxgqvAmJzhBdvD_eUf9Bdg/exec';
  
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var guideSheet = ss.getSheetByName('Tour Guide');
  var lastRow = guideSheet.getLastRow();
  var data = guideSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  // ── جمع أسماء المرشدين الفريدة (Registered + Unique فقط) ──
  var guidesMap = {};
  
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    var registered = String(data[i][9]);
    var unique = String(data[i][10]);
    
    if (!name) continue;
    if (registered.indexOf('✅') === -1 || unique.indexOf('✅') === -1) continue;
    
    if (!guidesMap[name]) {
      guidesMap[name] = { name: name, count: 0 };
    }
    guidesMap[name].count++;
  }
  
  // ── إنشاء أو مسح شيت الروابط ──
  var linkSheet = ss.getSheetByName('روابط المرشدين');
  if (!linkSheet) {
    linkSheet = ss.insertSheet('روابط المرشدين');
  } else {
    linkSheet.clear();
  }
  
  // ── رؤوس الأعمدة ──
  var headers = [['#', 'اسم المرشد', 'عدد الحجاج', 'الرابط']];
  linkSheet.getRange(1, 1, 1, 4).setValues(headers);
  linkSheet.getRange(1, 1, 1, 4)
    .setBackground('#1a5276')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontFamily('Tajawal');
  
  // ── بناء الصفوف ──
  var guides = [];
  for (var name in guidesMap) guides.push(guidesMap[name]);
  guides.sort(function(a, b) { return a.name.localeCompare(b.name); });
  
  var rows = [];
  for (var i = 0; i < guides.length; i++) {
    var g = guides[i];
    var url = DEPLOY_URL + '?g=' + encodeURIComponent(g.name);
    rows.push([i + 1, g.name, g.count, url]);
  }
  
  if (rows.length > 0) {
    linkSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  
  // ── تنسيق ──
  linkSheet.setColumnWidth(1, 40);
  linkSheet.setColumnWidth(2, 200);
  linkSheet.setColumnWidth(3, 100);
  linkSheet.setColumnWidth(4, 600);
  linkSheet.setRightToLeft(true);
  linkSheet.setFrozenRows(1);
  
  // ── تحويل عمود الرابط لروابط قابلة للنقر ──
  for (var i = 0; i < rows.length; i++) {
    var cell = linkSheet.getRange(i + 2, 4);
    var url = rows[i][3];
    cell.setFormula('=HYPERLINK("' + url + '", "افتح الرابط")');
    cell.setFontColor('#2980b9');
  }
  
  Logger.log('تم توليد ' + rows.length + ' رابط في شيت "روابط المرشدين"');

}