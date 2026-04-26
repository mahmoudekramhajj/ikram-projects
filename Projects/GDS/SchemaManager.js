/**
 * SchemaManager.js — إدارة هيكل شيت B2C
 *
 * المهمة: إنشاء الأعمدة الجديدة في نهاية شيت B2C تلقائياً.
 * BJ (62) = آخر رابط معالَج
 * BK (63) = مُعدَّل يدوياً
 * BL (64) = عدد محاولات الفشل
 * BM (65) = يحتاج دخول نسك
 * BN (66) = نوع عقد الطيران (B2B/B2C) [قائمة منسدلة]
 * BO (67) = نوع التغيير (6 أنواع) [قائمة منسدلة]
 * BP (68) = مصدر التغيير (إيميل/رابط/يدوي) [قائمة منسدلة]
 * BQ (69) = تاريخ آخر إيميل
 * BR (70) = معرّف الإيميل
 * BS-BY (71-77) = DEP0 — قطعة الذهاب الإضافية (للحالات 3+ قطع)
 * BZ-CF (78-84) = RET3 — قطعة العودة الإضافية (للحالات 3+ قطع)
 *
 * الدوال العامة (قابلة للاستدعاء عبر ClaudeAPI):
 *   runSchemaMigration()     — إنشاء/التحقق من الأعمدة + إضافة Data Validation
 *   getSchemaStatus()        — تقرير حالة الأعمدة دون تعديل
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.SchemaManager = {
  /**
   * يتحقق من وجود الأعمدة الجديدة، ويُنشئها إن كانت مفقودة.
   * يُصحّح العناوين إن كانت خاطئة.
   *
   * @return {Object} { status, colsAdded, headersFixed, lastCol }
   */
  ensureColumns: function() {
    var sheet = GDS2.SchemaManager._getSheet();
    var lastCol = sheet.getLastColumn();
    var maxRequired = GDS2.Config.COL.RET3_ARR_TIME; // 84 (CF)

    GDS2.Log.info('SchemaManager: checking', { lastCol: lastCol, required: maxRequired });

    var colsAdded = 0;
    if (lastCol < maxRequired) {
      var colsToAdd = maxRequired - lastCol;
      sheet.insertColumnsAfter(lastCol, colsToAdd);
      colsAdded = colsToAdd;
      GDS2.Log.info('SchemaManager: columns inserted', { added: colsToAdd });
    }

    // كتابة/تصحيح الرؤوس
    var headersFixed = GDS2.SchemaManager._ensureHeaders(sheet);

    // إضافة Data Validation للأعمدة الجديدة (BN, BO, BP)
    var dropdownsAdded = GDS2.SchemaManager._ensureDropdowns(sheet);

    // حفظ timestamp الترحيل
    GDS2.State.set(GDS2.Config.PROP.LAST_MIGRATION, new Date().toISOString());

    var result = {
      status: (colsAdded > 0 || headersFixed > 0 || dropdownsAdded > 0) ? 'updated' : 'ok',
      colsAdded: colsAdded,
      headersFixed: headersFixed,
      dropdownsAdded: dropdownsAdded,
      lastCol: sheet.getLastColumn(),
      timestamp: new Date().toISOString()
    };

    GDS2.Log.info('SchemaManager: done', result);
    return result;
  },

  /**
   * تقرير حالة الأعمدة دون أي تعديل.
   *
   * @return {Object} { lastCol, requiredCol, missingCols, headers }
   */
  getStatus: function() {
    var sheet = GDS2.SchemaManager._getSheet();
    var lastCol = sheet.getLastColumn();
    var requiredCol = GDS2.Config.COL.RET3_ARR_TIME; // 84

    var headers = {};
    for (var col = 62; col <= Math.min(84, lastCol); col++) {
      headers[GDS2.SchemaManager._colLetter(col)] = sheet.getRange(1, col).getValue();
    }

    var missingCols = [];
    for (var c = 62; c <= 84; c++) {
      if (c > lastCol) {
        missingCols.push(GDS2.SchemaManager._colLetter(c));
      }
    }

    return {
      lastCol: lastCol,
      lastColLetter: GDS2.SchemaManager._colLetter(lastCol),
      requiredCol: requiredCol,
      requiredColLetter: GDS2.SchemaManager._colLetter(requiredCol),
      missingCols: missingCols,
      headers: headers,
      lastMigration: GDS2.State.get(GDS2.Config.PROP.LAST_MIGRATION)
    };
  },

  // ==================== Private ====================

  _getSheet: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var sheet = ss.getSheetByName(GDS2.Config.SHEET_B2C);
    if (!sheet) {
      throw new Error('Sheet not found: ' + GDS2.Config.SHEET_B2C);
    }
    return sheet;
  },

  _ensureHeaders: function(sheet) {
    var expected = GDS2.Config.NEW_COL_HEADERS; // { 62: '...', ..., 84: '...' }
    var fixed = 0;
    var totalCols = 23; // 62..84 (BJ..CF)

    // قراءة الرؤوس الحالية دفعة واحدة
    var current = sheet.getRange(1, 62, 1, totalCols).getValues()[0];
    var needsUpdate = false;
    var newRow = [];

    for (var i = 0; i < totalCols; i++) {
      var col = 62 + i;
      var expectedHeader = expected[col];
      if (current[i] !== expectedHeader) {
        needsUpdate = true;
        fixed++;
      }
      newRow.push(expectedHeader);
    }

    if (needsUpdate) {
      var range = sheet.getRange(1, 62, 1, totalCols);
      range.setValues([newRow]);
      range.setFontWeight('bold').setHorizontalAlignment('center');
      GDS2.Log.info('SchemaManager: headers corrected', { count: fixed });
    }

    return fixed;
  },

  /**
   * إضافة Data Validation (قوائم منسدلة) للأعمدة BN, BO, BP.
   * يتم تطبيقها على نطاق الصفوف 2 إلى آخر صف.
   *
   * @return {number} عدد الأعمدة التي تم تحديثها
   */
  _ensureDropdowns: function(sheet) {
    var dropdowns = GDS2.Config.DROPDOWN_VALUES; // { 66: [...], 67: [...], 68: [...] }
    var lastRow = Math.max(sheet.getLastRow(), 2);
    var addedCount = 0;

    for (var col in dropdowns) {
      if (!dropdowns.hasOwnProperty(col)) continue;
      var values = dropdowns[col];
      var range = sheet.getRange(2, parseInt(col, 10), lastRow - 1, 1);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(values, true)
        .setAllowInvalid(true) // نسمح بقيم خارج القائمة (للحالات الاستثنائية)
        .build();
      range.setDataValidation(rule);
      addedCount++;
    }

    if (addedCount > 0) {
      GDS2.Log.info('SchemaManager: dropdowns added', { count: addedCount });
    }
    return addedCount;
  },

  _colLetter: function(col) {
    var s = '';
    while (col > 0) {
      var r = (col - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      col = Math.floor((col - 1) / 26);
    }
    return s;
  }
};

// ==================== Global Entry Points ====================
// (قابلة للاستدعاء عبر ClaudeAPI: ?action=run&fn=runSchemaMigration)

function runSchemaMigration() {
  return GDS2.SchemaManager.ensureColumns();
}

function getSchemaStatus() {
  return GDS2.SchemaManager.getStatus();
}
