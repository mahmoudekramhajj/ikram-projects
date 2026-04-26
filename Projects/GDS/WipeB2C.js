/**
 * WipeB2C.js — مسح شامل لصفوف شيت B2C (عدا صف الرؤوس)
 *
 * تُستخدم مرة واحدة عند بدء النظام الجديد:
 *   1. يحفظ نسخة احتياطية من المستخدم (لا تلقائي)
 *   2. يمسح جميع الصفوف 2..N
 *   3. يُعيد لون الخلفية للافتراضي
 *
 * بعد المسح: استدعِ syncNames() لاستعادة البيانات الشخصية من PD.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.WipeB2C = {
  /**
   * مسح كل الصفوف عدا صف الرؤوس.
   */
  wipeAllRows: function() {
    var startTime = new Date();
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    if (!sheet) return { status: 'error', reason: 'sheet_missing' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { status: 'ok', rows_deleted: 0, reason: 'already_empty' };
    }

    var lastCol = sheet.getLastColumn();
    var rowsToDelete = lastRow - 1;

    // تجميد صف الرؤوس إن لم يكن مجمّداً (مطلوب لحذف باقي الصفوف)
    if (sheet.getFrozenRows() < 1) {
      sheet.setFrozenRows(1);
    }

    // حذف الصفوف 2..lastRow
    // GAS لا يسمح بحذف كل non-frozen rows. نترك صفاً واحداً ونفرّغه.
    if (rowsToDelete > 1) {
      sheet.deleteRows(3, rowsToDelete - 1);
    }
    // تفريغ الصف 2 المتبقي
    sheet.getRange(2, 1, 1, lastCol).clearContent().setBackground(null);

    var elapsed = (new Date() - startTime) / 1000;
    GDS2.Log.info('WipeB2C: wiped', { rows: rowsToDelete, elapsed_sec: elapsed });

    return {
      status: 'ok',
      rows_deleted: rowsToDelete,
      last_col: lastCol,
      elapsed_sec: elapsed
    };
  }
};

// ==================== Global Entry Points ====================

function wipeB2CRows() {
  return GDS2.WipeB2C.wipeAllRows();
}
