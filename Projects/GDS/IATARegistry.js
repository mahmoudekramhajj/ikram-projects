/**
 * IATARegistry.js — سجل أكواد المطارات IATA
 *
 * يبني قائمة المطارات المعتمدة من البيانات التاريخية:
 * - شيت B2C (الأعمدة AG-BI) — قبل أي مسح
 * - شيت الطيران (الأعمدة 21-48) — B2B
 *
 * دوال عامة:
 *   extractIATAFromHistory()  — بناء/إعادة بناء القائمة
 *   getIATAList()             — القائمة الحالية
 *   isValidIATA(code)         — تحقق من صحة كود
 *   addIATA(code)             — إضافة كود جديد (تلقائي عند كشف جديد)
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.IATARegistry = {
  /**
   * يمسح أعمدة الطيران في شيتَي B2C والطيران ويستخرج كل كود IATA (3 أحرف كبيرة).
   * يحفظ النتيجة في Script Properties.
   */
  extractFromHistory: function() {
    var codes = {};
    var stats = {
      b2c: 0,
      flights: 0,
      total: 0
    };

    // B2C: AG-BI (33-61) — 29 عمود
    var b2cCount = GDS2.IATARegistry._extractFromSheet(
      GDS2.Config.SHEET_B2C,
      33,
      29,
      [3, 4, 10, 11, 17, 18, 24, 25], // From/To في كل block
      codes
    );
    stats.b2c = b2cCount;

    // شيت الطيران: 21-48 — 28 عمود
    // Blocks: 21-27 (transit out), 28-34 (main out), 35-41 (main ret), 42-48 (transit ret)
    // From/To في كل block = indices 3,4 (بدءاً من start)
    var flightsCount = GDS2.IATARegistry._extractFromSheet(
      GDS2.Config.SHEET_FLIGHTS,
      21,
      28,
      [3, 4, 10, 11, 17, 18, 24, 25],
      codes
    );
    stats.flights = flightsCount;

    var codeList = Object.keys(codes).sort();
    stats.total = codeList.length;

    GDS2.State.setJSON(GDS2.Config.PROP.IATA_REGISTRY, codeList);
    GDS2.Log.info('IATARegistry: extracted', stats);

    return {
      status: 'ok',
      stats: stats,
      codes: codeList
    };
  },

  /**
   * قائمة IATA الحالية من Script Properties.
   */
  getList: function() {
    return GDS2.State.getJSON(GDS2.Config.PROP.IATA_REGISTRY) || [];
  },

  /**
   * هل الكود موجود في القائمة؟
   */
  isValid: function(code) {
    if (!code) return false;
    var upper = String(code).toUpperCase().trim();
    if (!/^[A-Z]{3}$/.test(upper)) return false;
    var list = GDS2.IATARegistry.getList();
    return list.indexOf(upper) !== -1;
  },

  /**
   * إزالة قائمة من الأكواد (تنظيف OCR noise).
   * @param {string[]} codesToRemove
   * @return {Object} { removed: [...], notFound: [...] }
   */
  removeMany: function(codesToRemove) {
    var list = GDS2.IATARegistry.getList();
    var removed = [];
    var notFound = [];

    for (var i = 0; i < codesToRemove.length; i++) {
      var code = String(codesToRemove[i]).toUpperCase().trim();
      var idx = list.indexOf(code);
      if (idx === -1) {
        notFound.push(code);
      } else {
        list.splice(idx, 1);
        removed.push(code);
      }
    }

    GDS2.State.setJSON(GDS2.Config.PROP.IATA_REGISTRY, list);
    GDS2.Log.info('IATARegistry: cleanup', { removed: removed.length, total: list.length });
    return {
      removed: removed,
      notFound: notFound,
      remainingCount: list.length
    };
  },

  /**
   * إضافة كود جديد (تلقائي عند كشف مطار غير موجود).
   * @return {boolean} true إذا أُضيف، false إذا موجود مسبقاً
   */
  add: function(code) {
    if (!code) return false;
    var upper = String(code).toUpperCase().trim();
    if (!/^[A-Z]{3}$/.test(upper)) {
      GDS2.Log.warn('IATARegistry.add: invalid format', { code: code });
      return false;
    }

    var list = GDS2.IATARegistry.getList();
    if (list.indexOf(upper) !== -1) return false;

    list.push(upper);
    list.sort();
    GDS2.State.setJSON(GDS2.Config.PROP.IATA_REGISTRY, list);
    GDS2.Log.info('IATARegistry: added', { code: upper, total: list.length });
    return true;
  },

  // ==================== Private ====================

  _extractFromSheet: function(sheetName, startCol, width, fromToIndices, codes) {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      GDS2.Log.warn('IATARegistry: sheet missing', { sheet: sheetName });
      return 0;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;

    var lastCol = sheet.getLastColumn();
    var actualWidth = Math.min(width, lastCol - startCol + 1);
    if (actualWidth < 1) return 0;

    var range = sheet.getRange(2, startCol, lastRow - 1, actualWidth);
    var values = range.getValues();
    var found = 0;

    for (var r = 0; r < values.length; r++) {
      for (var i = 0; i < fromToIndices.length; i++) {
        var idx = fromToIndices[i];
        if (idx >= actualWidth) continue;
        var v = String(values[r][idx] || '').toUpperCase().trim();
        if (/^[A-Z]{3}$/.test(v)) {
          if (!codes[v]) found++;
          codes[v] = true;
        }
      }
    }

    return found;
  }
};

// ==================== Global Entry Points ====================

function extractIATAFromHistory() {
  return GDS2.IATARegistry.extractFromHistory();
}

function getIATAList() {
  return {
    count: GDS2.IATARegistry.getList().length,
    codes: GDS2.IATARegistry.getList()
  };
}

function isValidIATA(code) {
  return GDS2.IATARegistry.isValid(code);
}

function addIATA(code) {
  return GDS2.IATARegistry.add(code);
}

/**
 * تنظيف الأكواد المؤكَّد خطأها فقط — بعد فحص كل كود مشبوه عبر بحث Google.
 * كل كود في هذه القائمة تحقّقتُ يدوياً أنه ليس مطاراً حقيقياً.
 */
function cleanupIATANoise() {
  var blacklist = [
    // شركات طيران خلطها OCR كأكواد مطارات
    'ITA',  // ITA Airways (شركة)
    'KLM',  // KLM (شركة)
    // أكواد غير قياسية (لا توجد في IATA)
    'EID',  // الصحيح EDI (إدنبرة)
    'IIS',  // لا يوجد مطار بهذا الكود
    'NUR'   // أستانا غيّرت كودها إلى NQZ، NUR ليس IATA
  ];
  return GDS2.IATARegistry.removeMany(blacklist);
}
