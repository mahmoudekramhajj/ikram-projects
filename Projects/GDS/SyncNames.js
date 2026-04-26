/**
 * SyncNames.js — مزامنة بيانات الحجاج من Presonal Details إلى B2C
 *
 * القاعدة: أي حاج في PD لديه ticketUrl (العمود Z) → يُضاف/يُحدَّث في B2C
 * - بغض النظر عن نوع العقد (B2B أو B2C)
 * - تُنقل أعمدة PD 1-27 إلى B2C 1-27 مباشرةً
 * - تُنقل أعمدة PD 29-33 (AC-AG) إلى B2C 28-32 (نقل البري)
 * - أعمدة B2C 33-65 (الطيران + metadata) لا تُلمس
 *
 * دوال عامة:
 *   syncNames()      — التشغيل الفعلي
 *   diagnosePD()     — تقرير هيكل PD (قراءة فقط)
 *   previewSync()    — معاينة الحجاج الذين سيُضافون/يُحدَّثون دون تعديل
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.SyncNames = {
  // عدد الأعمدة المنقولة من PD إلى B2C
  PD_DIRECT_COLS: 27,     // PD cols 1-27 → B2C cols 1-27
  PD_SHIFT_FROM: 29,      // PD col 29 (AC) يبدأ نقل النقل
  PD_SHIFT_TO: 33,        // PD col 33 (AG) آخر عمود نقل
  B2C_SYNC_WIDTH: 32,     // عدد الأعمدة المكتوبة في B2C (1-32)
  PD_TICKET_URL_COL: 26,  // Z
  PD_PASSPORT_COL: 6,     // F

  /**
   * تنفيذ المزامنة الفعلية: إضافة جدد + تحديث الموجودين.
   */
  run: function() {
    var startTime = new Date();
    var result = GDS2.SyncNames._analyze();

    if (result.error) return result;

    var b2cSheet = result.b2cSheet;

    // تطبيق التحديثات (batch: read-modify-write)
    if (result.updateOps.length > 0) {
      GDS2.SyncNames._applyUpdates(b2cSheet, result.updateOps, result.b2cLastRow);
    }

    // إضافة الصفوف الجديدة (batch)
    if (result.newRows.length > 0) {
      var startRow = b2cSheet.getLastRow() + 1;
      b2cSheet.getRange(startRow, 1, result.newRows.length, GDS2.SyncNames.B2C_SYNC_WIDTH)
        .setValues(result.newRows);
    }

    var elapsed = (new Date() - startTime) / 1000;
    var summary = {
      status: 'ok',
      added: result.added,
      updated: result.updated,
      skipped_no_url: result.skippedNoUrl,
      skipped_no_passport: result.skippedNoPassport,
      elapsed_sec: elapsed
    };
    GDS2.Log.info('SyncNames: done', summary);
    return summary;
  },

  /**
   * معاينة دون أي تعديل.
   */
  preview: function() {
    var result = GDS2.SyncNames._analyze();
    if (result.error) return result;

    // لا نكتب شيئاً
    return {
      status: 'preview',
      will_add: result.added,
      will_update: result.updated,
      skipped_no_url: result.skippedNoUrl,
      skipped_no_passport: result.skippedNoPassport,
      sample_new: result.newRows.slice(0, 3).map(function(r) {
        return { passport: r[5], name_ar: r[8] + ' ' + r[9] };
      })
    };
  },

  /**
   * تقرير هيكل PD (رؤوس الأعمدة).
   */
  diagnosePD: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var pdSheet = ss.getSheetByName(GDS2.Config.SHEET_PD);
    if (!pdSheet) return { error: 'PD sheet missing: ' + GDS2.Config.SHEET_PD };

    var lastCol = pdSheet.getLastColumn();
    var lastRow = pdSheet.getLastRow();
    var headers = pdSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var headerMap = {};
    for (var i = 0; i < headers.length; i++) {
      headerMap[i + 1] = headers[i];
    }

    // إحصاء الحجاج الذين لديهم ticketUrl
    var pdRange = pdSheet.getRange(2, 1, lastRow - 1, Math.min(lastCol, 33)).getValues();
    var withUrl = 0;
    var withoutUrl = 0;
    for (var r = 0; r < pdRange.length; r++) {
      var url = String(pdRange[r][GDS2.SyncNames.PD_TICKET_URL_COL - 1] || '').trim();
      if (url) withUrl++;
      else withoutUrl++;
    }

    return {
      sheetName: GDS2.Config.SHEET_PD,
      colCount: lastCol,
      rowCount: lastRow - 1,
      headers: headerMap,
      ticketUrlCol: GDS2.SyncNames.PD_TICKET_URL_COL,
      ticketUrlHeader: headerMap[GDS2.SyncNames.PD_TICKET_URL_COL] || null,
      passportCol: GDS2.SyncNames.PD_PASSPORT_COL,
      passportHeader: headerMap[GDS2.SyncNames.PD_PASSPORT_COL] || null,
      stats: {
        pilgrims_with_ticketUrl: withUrl,
        pilgrims_without_ticketUrl: withoutUrl
      }
    };
  },

  // ==================== Private ====================

  _analyze: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var pdSheet = ss.getSheetByName(GDS2.Config.SHEET_PD);
    var b2cSheet = ss.getSheetByName(GDS2.Config.SHEET_B2C);

    if (!pdSheet) return { error: 'PD sheet missing' };
    if (!b2cSheet) return { error: 'B2C sheet missing' };

    var pdLastRow = pdSheet.getLastRow();
    if (pdLastRow < 2) {
      return { error: 'PD is empty', b2cSheet: b2cSheet, updateOps: [], newRows: [], added: 0, updated: 0, skippedNoUrl: 0, skippedNoPassport: 0 };
    }

    // قراءة PD
    var pdLastCol = Math.min(pdSheet.getLastColumn(), 33);
    var pdData = pdSheet.getRange(2, 1, pdLastRow - 1, pdLastCol).getValues();

    // بناء خريطة جوازات B2C
    var b2cLastRow = b2cSheet.getLastRow();
    var b2cPassportMap = {};
    if (b2cLastRow >= 2) {
      var b2cPassports = b2cSheet.getRange(2, GDS2.SyncNames.PD_PASSPORT_COL, b2cLastRow - 1, 1).getValues();
      for (var i = 0; i < b2cPassports.length; i++) {
        var p = String(b2cPassports[i][0] || '').trim();
        if (p) b2cPassportMap[p] = i + 2;
      }
    }

    var added = 0, updated = 0, skippedNoUrl = 0, skippedNoPassport = 0;
    var newRows = [];
    var updateOps = [];

    for (var r = 0; r < pdData.length; r++) {
      var pdRow = pdData[r];

      var ticketUrl = String(pdRow[GDS2.SyncNames.PD_TICKET_URL_COL - 1] || '').trim();
      if (!ticketUrl) {
        skippedNoUrl++;
        continue;
      }

      var passport = String(pdRow[GDS2.SyncNames.PD_PASSPORT_COL - 1] || '').trim();
      if (!passport) {
        skippedNoPassport++;
        continue;
      }

      var b2cValues = GDS2.SyncNames._buildB2CRow(pdRow);

      if (b2cPassportMap[passport]) {
        updateOps.push({ row: b2cPassportMap[passport], values: b2cValues });
        updated++;
      } else {
        newRows.push(b2cValues);
        added++;
      }
    }

    return {
      b2cSheet: b2cSheet,
      b2cLastRow: b2cLastRow,
      updateOps: updateOps,
      newRows: newRows,
      added: added,
      updated: updated,
      skippedNoUrl: skippedNoUrl,
      skippedNoPassport: skippedNoPassport
    };
  },

  _buildB2CRow: function(pdRow) {
    var row = [];
    // PD cols 1-27 → B2C cols 1-27
    for (var j = 0; j < GDS2.SyncNames.PD_DIRECT_COLS; j++) {
      row.push(pdRow[j] !== undefined ? pdRow[j] : '');
    }
    // PD cols 29-33 → B2C cols 28-32 (تخطّي PD col 28)
    for (var k = GDS2.SyncNames.PD_SHIFT_FROM - 1; k < GDS2.SyncNames.PD_SHIFT_TO; k++) {
      row.push(pdRow[k] !== undefined ? pdRow[k] : '');
    }
    return row;
  },

  _applyUpdates: function(b2cSheet, updateOps, b2cLastRow) {
    if (b2cLastRow < 2) return;

    // read-modify-write لتحسين الأداء (1 setValues بدل N)
    var existingData = b2cSheet.getRange(2, 1, b2cLastRow - 1, GDS2.SyncNames.B2C_SYNC_WIDTH).getValues();

    for (var u = 0; u < updateOps.length; u++) {
      var op = updateOps[u];
      var idx = op.row - 2; // convert to 0-based
      if (idx >= 0 && idx < existingData.length) {
        existingData[idx] = op.values;
      }
    }

    b2cSheet.getRange(2, 1, b2cLastRow - 1, GDS2.SyncNames.B2C_SYNC_WIDTH).setValues(existingData);
  }
};

// ==================== Global Entry Points ====================

function syncNames() {
  return GDS2.SyncNames.run();
}

function previewSyncNames() {
  return GDS2.SyncNames.preview();
}

function diagnosePD() {
  return GDS2.SyncNames.diagnosePD();
}
