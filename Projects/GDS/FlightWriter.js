/**
 * FlightWriter.js — كتابة شاملة لبيانات الطيران + تلوين الصف
 *
 * قواعد الكتابة:
 *   1. قبل أي كتابة: فحص guards (BK=manual, BL=fails, BM=nusuk)
 *   2. كتابة شاملة: تُمسح AG-BI ثم تُكتب دفعة واحدة
 *   3. تحديث BJ (آخر رابط معالَج) = url الحالي
 *   4. إعادة BL إلى 0 عند النجاح
 *   5. عدم لمس BK (manual) أو BM (nusuk)
 *   6. تلوين الصف أخضر عند النجاح، أصفر/أحمر عند الفشل
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.FlightWriter = {
  /**
   * كتابة بيانات طيران لصف بعد نجاح التحقق.
   * @param {number} rowNumber
   * @param {Object} normalizedData
   * @param {string} url - رابط PD الحالي (سيُكتب في BJ)
   * @return {Object} { status, reason?, ... }
   */
  writeSuccess: function(rowNumber, normalizedData, url) {
    var sheet = GDS2.FlightWriter._getSheet();

    // فحص guards
    var guardCheck = GDS2.FlightWriter._checkGuards(sheet, rowNumber);
    if (guardCheck.blocked) {
      return { status: 'skipped', reason: guardCheck.reason, row: rowNumber };
    }

    // بناء مصفوفات بيانات الطيران
    var flightData = GDS2.Classifier.buildFlightArray(normalizedData); // 29 قيمة (AG-BI)
    var dep0Data = GDS2.Classifier.buildDep0Array(normalizedData);     // 7 قيم (BS-BY)
    var ret3Data = GDS2.Classifier.buildRet3Array(normalizedData);     // 7 قيم (BZ-CF)

    // كتابة شاملة: مسح + كتابة جديدة (AG-BI)
    var flightRange = sheet.getRange(rowNumber, GDS2.Config.FLIGHT_RANGE_START, 1, GDS2.Config.FLIGHT_RANGE_WIDTH);
    flightRange.clearContent();
    flightRange.setValues([flightData]);

    // كتابة DEP0 (BS-BY = 71-77) — حتى لو فارغة (للمسح الشامل)
    var dep0Range = sheet.getRange(rowNumber, GDS2.Config.FLIGHT_BLOCKS.DEP0.start, 1, GDS2.Config.FIELDS_PER_BLOCK);
    dep0Range.clearContent();
    dep0Range.setValues([dep0Data]);

    // كتابة RET3 (BZ-CF = 78-84) — حتى لو فارغة (للمسح الشامل)
    var ret3Range = sheet.getRange(rowNumber, GDS2.Config.FLIGHT_BLOCKS.RET3.start, 1, GDS2.Config.FIELDS_PER_BLOCK);
    ret3Range.clearContent();
    ret3Range.setValues([ret3Data]);

    // تحديث BJ (col 62) + BL (col 64)
    sheet.getRange(rowNumber, GDS2.Config.COL.LAST_URL).setValue(url);
    sheet.getRange(rowNumber, GDS2.Config.COL.FAIL_COUNT).setValue(0);

    // تلوين الصف أخضر
    GDS2.FlightWriter._colorRow(sheet, rowNumber, GDS2.Config.ROW_COLORS.SUCCESS);

    var hasDep0 = !!(normalizedData.segments && normalizedData.segments.dep0);
    var hasRet3 = !!(normalizedData.segments && normalizedData.segments.ret3);

    return {
      status: 'ok',
      row: rowNumber,
      wrote_cols: GDS2.Config.FLIGHT_RANGE_WIDTH + 14,
      has_dep0: hasDep0,
      has_ret3: hasRet3,
      bj_updated: true,
      bl_reset: true
    };
  },

  /**
   * تسجيل فشل: زيادة BL + تلوين الصف.
   */
  recordFailure: function(rowNumber, reason) {
    var sheet = GDS2.FlightWriter._getSheet();

    var current = sheet.getRange(rowNumber, GDS2.Config.COL.FAIL_COUNT).getValue();
    var currentBL = Number(current) || 0;
    var newBL = currentBL + 1;
    sheet.getRange(rowNumber, GDS2.Config.COL.FAIL_COUNT).setValue(newBL);

    // تلوين: أصفر إن < max، أحمر إن وصل للحد
    var color = newBL >= GDS2.Config.MAX_FAIL_ATTEMPTS
      ? GDS2.Config.ROW_COLORS.ERROR
      : GDS2.Config.ROW_COLORS.WARNING;
    GDS2.FlightWriter._colorRow(sheet, rowNumber, color);

    return {
      status: 'failure_recorded',
      row: rowNumber,
      reason: reason,
      new_fail_count: newBL,
      stopped: newBL >= GDS2.Config.MAX_FAIL_ATTEMPTS
    };
  },

  /**
   * تعليم BM=TRUE (رابط نسك يحتاج دخول) + تلوين رمادي.
   * لا يزيد BL (ليست فشلاً تقنياً).
   */
  markNusukAuthNeeded: function(rowNumber) {
    var sheet = GDS2.FlightWriter._getSheet();
    sheet.getRange(rowNumber, GDS2.Config.COL.NUSUK_AUTH).setValue(true);
    GDS2.FlightWriter._colorRow(sheet, rowNumber, GDS2.Config.ROW_COLORS.NUSUK);
    return { status: 'nusuk_marked', row: rowNumber };
  },

  /**
   * تعليم BK=TRUE (مُعدَّل يدوياً — قفل دائم) + تلوين أزرق.
   * يمنع الـ pipeline التلقائي من إعادة معالجة الصف.
   * @param {number} rowNumber
   * @return {Object} { status, row }
   */
  markManual: function(rowNumber) {
    var sheet = GDS2.FlightWriter._getSheet();
    sheet.getRange(rowNumber, GDS2.Config.COL.MANUAL).setValue(true);
    GDS2.FlightWriter._colorRow(sheet, rowNumber, GDS2.Config.ROW_COLORS.MANUAL);
    return { status: 'manual_marked', row: rowNumber };
  },

  /**
   * تعليم صف كـ "مُعدَّل يدوياً" بناءً على رقم الجواز.
   * @param {string} passport
   * @return {Object} { status, passport, row?, reason? }
   */
  markManualByPassport: function(passport) {
    var sheet = GDS2.FlightWriter._getSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'not_found', passport: passport, reason: 'empty_sheet' };

    var passports = sheet.getRange(2, GDS2.Config.COL.PASSPORT, lastRow - 1, 1).getValues();
    var target = String(passport).trim();

    for (var i = 0; i < passports.length; i++) {
      if (String(passports[i][0] || '').trim() === target) {
        var row = i + 2;
        return GDS2.FlightWriter.markManual(row);
      }
    }
    return { status: 'not_found', passport: passport };
  },

  /**
   * تلوين manual (BK=TRUE — أزرق).
   */
  applyManualColor: function(rowNumber) {
    var sheet = GDS2.FlightWriter._getSheet();
    GDS2.FlightWriter._colorRow(sheet, rowNumber, GDS2.Config.ROW_COLORS.MANUAL);
    return { status: 'manual_colored', row: rowNumber };
  },

  // ==================== Private ====================

  _getSheet: function() {
    return SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  },

  _checkGuards: function(sheet, rowNumber) {
    // قراءة BK (63), BL (64), BM (65) دفعة واحدة
    var guards = sheet.getRange(rowNumber, GDS2.Config.COL.MANUAL, 1, 3).getValues()[0];
    var bk = guards[0];
    var bl = Number(guards[1]) || 0;
    var bm = guards[2];

    if (bk === true || String(bk).toUpperCase() === 'TRUE') {
      return { blocked: true, reason: 'manual_edit_protected' };
    }
    if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) {
      return { blocked: true, reason: 'max_fail_attempts_reached' };
    }
    if (bm === true || String(bm).toUpperCase() === 'TRUE') {
      return { blocked: true, reason: 'nusuk_auth_needed' };
    }
    return { blocked: false };
  },

  _colorRow: function(sheet, rowNumber, color) {
    // تلوين كامل: من العمود 1 إلى آخر عمود (RET3_ARR_TIME = 84)
    sheet.getRange(rowNumber, 1, 1, GDS2.Config.COL.RET3_ARR_TIME).setBackground(color);
  }
};

// ==================== Global Entry Points ====================

/**
 * تعليم صف كـ مُعدَّل يدوياً (قفل دائم) بناءً على الجواز.
 * مثال: markManual("AP0316524")
 */
function markManual(passport) {
  return GDS2.FlightWriter.markManualByPassport(passport);
}

/**
 * مسح BM=TRUE لكل الصفوف (لإلغاء وضع "نسك محمي").
 * يُستخدم بعد إصلاح Drive URL transformation — الصفوف المُعلَّمة سابقاً ستُعالَج من جديد.
 */
function clearAllNusukAuth() {
  var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };

  var range = sheet.getRange(2, GDS2.Config.COL.NUSUK_AUTH, lastRow - 1, 1);
  var values = range.getValues();
  var cleared = 0;
  var clearedRows = [];

  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];
    if (v === true || String(v).toUpperCase() === 'TRUE') {
      values[i][0] = '';
      cleared++;
      clearedRows.push(i + 2);
    }
  }

  if (cleared > 0) {
    range.setValues(values);
    // إعادة تلوين الصفوف للحالة الافتراضية (أبيض)
    for (var j = 0; j < clearedRows.length; j++) {
      sheet.getRange(clearedRows[j], 1, 1, GDS2.Config.COL.RET3_ARR_TIME).setBackground(null);
    }
  }

  return { count: cleared, rows: clearedRows };
}

/**
 * تعليم عدة صفوف معاً.
 * مثال: markManualBatch(["AP0316524", "AA5710607", "AA6703604", "RD7621315"])
 */
function markManualBatch(passports) {
  var results = [];
  if (!passports || !passports.length) return { count: 0, results: [] };
  for (var i = 0; i < passports.length; i++) {
    results.push(GDS2.FlightWriter.markManualByPassport(passports[i]));
  }
  var success = results.filter(function(r) { return r.status === 'manual_marked'; }).length;
  return { count: passports.length, success: success, results: results };
}
