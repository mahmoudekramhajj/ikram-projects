/**
 * ResetFailures.js — إعادة تعيين الصفوف المتوقفة لإعادة المعالجة
 *
 * يُستخدم بعد تحسين prompt أو إصلاح bug.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.ResetFailures = {
  /**
   * إعادة تعيين كل الصفوف المتوقفة (BL >= MAX).
   * BL → 0، BJ → '' (لو موجود)، الصف يُلوَّن بلا لون.
   */
  resetStopped: function() {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { reset: 0 };

    var blCol = GDS2.Config.COL.FAIL_COUNT;  // 64
    var bjCol = GDS2.Config.COL.LAST_URL;    // 62

    // اقرأ BL + BJ لكل الصفوف
    var data = sheet.getRange(2, bjCol, lastRow - 1, 3).getValues(); // BJ, BK, BL
    var reset = [];

    for (var i = 0; i < data.length; i++) {
      var bl = Number(data[i][2]) || 0;
      if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) {
        var row = i + 2;
        sheet.getRange(row, blCol).setValue(0);
        sheet.getRange(row, bjCol).setValue(''); // امسح BJ حتى يعاد المعالجة
        // أزل لون الخلفية
        sheet.getRange(row, 1, 1, GDS2.Config.COL.NUSUK_AUTH).setBackground(null);
        reset.push(row);
      }
    }

    GDS2.Log.info('ResetFailures: reset stopped rows', { count: reset.length });
    return { reset_count: reset.length, rows: reset };
  },

  /**
   * إعادة تعيين كل صف فشل مرة واحدة (BL=1).
   */
  resetFailOnce: function() {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { reset: 0 };

    var blCol = GDS2.Config.COL.FAIL_COUNT;
    var data = sheet.getRange(2, blCol, lastRow - 1, 1).getValues();
    var reset = [];

    for (var i = 0; i < data.length; i++) {
      var bl = Number(data[i][0]) || 0;
      if (bl === 1) {
        var row = i + 2;
        sheet.getRange(row, blCol).setValue(0);
        sheet.getRange(row, 1, 1, GDS2.Config.COL.NUSUK_AUTH).setBackground(null);
        reset.push(row);
      }
    }

    return { reset_count: reset.length, rows: reset };
  }
};

// ==================== Global ====================

function resetStoppedRows() {
  return GDS2.ResetFailures.resetStopped();
}

function resetFailOnce() {
  return GDS2.ResetFailures.resetFailOnce();
}
