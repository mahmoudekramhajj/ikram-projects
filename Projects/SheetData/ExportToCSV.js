/**
 * ══════════════════════════════════════════════════════
 * تصدير شيت Ikram Abuown إلى CSV في Google Drive
 * ══════════════════════════════════════════════════════
 *
 * طريقة الاستخدام:
 * 1. افتح الشيت → Extensions → Apps Script
 * 2. أنشئ ملفاً جديداً والصق هذا الكود
 * 3. شغّل setupTrigger() مرة واحدة فقط لتفعيل التحديث التلقائي
 * 4. شغّل exportIkramAbuownToCSV() يدوياً للاختبار
 * ══════════════════════════════════════════════════════
 */

var SHEET_NAME = 'Ikram Abuown';
var FILE_NAME  = 'ikram_abuown.csv';
var FOLDER_NAME = 'اكرام الضيف';

// ─────────────────────────────────────────
// التصدير الرئيسي
// ─────────────────────────────────────────
function exportIkramAbuownToCSV() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('لم يتم العثور على الشيت: ' + SHEET_NAME);
    return;
  }

  var data = sheet.getDataRange().getValues();

  // تحويل البيانات إلى CSV
  var csv = data.map(function(row) {
    return row.map(function(cell) {
      var val = (cell instanceof Date)
        ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(cell);
      return '"' + val.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\n');

  // البحث عن مجلد اكرام الضيف أو الحفظ في الجذر
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  var folder  = folders.hasNext() ? folders.next() : DriveApp.getRootFolder();

  // تحديث الملف إن وجد أو إنشاؤه
  var files = folder.getFilesByName(FILE_NAME);
  if (files.hasNext()) {
    files.next().setContent(csv);
    Logger.log('تم تحديث الملف: ' + FILE_NAME);
  } else {
    folder.createFile(FILE_NAME, csv, MimeType.CSV);
    Logger.log('تم إنشاء الملف: ' + FILE_NAME);
  }
}

// ─────────────────────────────────────────
// إعداد التشغيل التلقائي كل ساعة
// شغّل هذه الدالة مرة واحدة فقط
// ─────────────────────────────────────────
function setupTrigger() {
  // حذف أي trigger سابق لتجنب التكرار
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'exportIkramAbuownToCSV') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // إنشاء trigger جديد كل ساعة
  ScriptApp.newTrigger('exportIkramAbuownToCSV')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('تم إعداد التحديث التلقائي كل ساعة');
}
