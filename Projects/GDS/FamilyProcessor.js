/**
 * FamilyProcessor.js — ميزة: تطبيق نفس بيانات الطيران على أفراد العائلة
 *
 * متى تُستخدم؟
 *   بعد نجاح معالجة تذكرة حاج، ندير في all_passengers (من Claude).
 *   لكل راكب في القائمة:
 *     - نبحث عن صفه في B2C (مطابقة اسم)
 *     - إذا وجدناه و URL في PD = URL الذي عالجناه → نطبّق نفس البيانات
 *     - إذا URL مختلف → نتركه للمعالجة العادية
 *
 * النتيجة: عائلة من 5 بنفس URL = معالجة واحدة بدل 5 (توفير كبير)
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.FamilyProcessor = {
  /**
   * تطبيق نفس البيانات على أفراد العائلة الذين يشتركون في نفس URL.
   * @param {Object} normalizedData - بيانات الطيران المُطبَّعة
   * @param {string} url - الرابط الذي تمت معالجته
   * @param {number} primaryRow - صف الحاج الأصلي (تمت كتابته)
   * @return {Object} تقرير
   */
  applyToFamily: function(normalizedData, url, primaryRow) {
    var passengers = (normalizedData && normalizedData.all_passengers) || [];
    if (passengers.length <= 1) {
      return { applied_count: 0, skipped_count: 0, reason: 'single_passenger' };
    }

    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { applied_count: 0, skipped_count: 0, reason: 'empty_sheet' };

    // قراءة الأعمدة اللازمة للمطابقة + URL + BJ
    // نحتاج: cols 11, 12 (first/last EN), 9, 10 (first/last AR), 26 (URL), 62 (BJ)
    // سنقرأ 1..65 مرة واحدة
    var allData = sheet.getRange(2, 1, lastRow - 1, GDS2.Config.COL.NUSUK_AUTH).getValues();

    var applied = [];
    var skipped = [];

    for (var i = 0; i < passengers.length; i++) {
      var paxName = GDS2.FamilyProcessor._cleanName(passengers[i]);
      if (!paxName) continue;

      var match = GDS2.FamilyProcessor._findByName(allData, paxName);
      if (!match) {
        skipped.push({ name: paxName, reason: 'not_in_b2c' });
        continue;
      }

      // تخطَّ الصف الأصلي
      if (match.row === primaryRow) continue;

      // تحقق أن URL يتطابق (نفس التذكرة)
      if (match.url !== url) {
        skipped.push({ name: paxName, reason: 'different_url', row: match.row });
        continue;
      }

      // تحقق أن BJ ليس نفس URL (أي لم يُعالَج بعد)
      if (match.bj === url) {
        skipped.push({ name: paxName, reason: 'already_processed', row: match.row });
        continue;
      }

      // تطبيق البيانات
      var writeResult = GDS2.FlightWriter.writeSuccess(match.row, normalizedData, url);
      if (writeResult.status === 'ok') {
        applied.push({ name: paxName, row: match.row });
      } else {
        skipped.push({ name: paxName, reason: writeResult.reason, row: match.row });
      }
    }

    GDS2.Log.info('FamilyProcessor', { primary: primaryRow, applied: applied.length, skipped: skipped.length });

    return {
      applied_count: applied.length,
      skipped_count: skipped.length,
      applied: applied,
      skipped: skipped
    };
  },

  // ==================== Private ====================

  /**
   * تنظيف اسم: إزالة Mr./Mrs./Ms./etc + uppercase + trim
   */
  _cleanName: function(name) {
    if (!name) return '';
    return String(name).trim()
      .replace(/^(Mr\.?|Mrs\.?|Ms\.?|Miss|Dr\.?|Prof\.?|Sir|Sayed|Sayyid)\s+/i, '')
      .replace(/\s+/g, ' ')
      .toUpperCase();
  },

  /**
   * بحث في بيانات B2C عن اسم (EN أو AR).
   */
  _findByName: function(allData, cleanedName) {
    var ciAR_FIRST = GDS2.Config.COL.FIRST_NAME_AR - 1;   // 9-1=8
    var ciAR_LAST = GDS2.Config.COL.LAST_NAME_AR - 1;    // 10-1=9
    var ciEN_FIRST = GDS2.Config.COL.FIRST_NAME_EN - 1;  // 11-1=10
    var ciEN_LAST = GDS2.Config.COL.LAST_NAME_EN - 1;    // 12-1=11
    var ciURL = GDS2.Config.COL.TICKET_URL - 1;          // 26-1=25
    var ciBJ = GDS2.Config.COL.LAST_URL - 1;             // 62-1=61

    for (var i = 0; i < allData.length; i++) {
      var firstEn = String(allData[i][ciEN_FIRST] || '').trim().toUpperCase();
      var lastEn = String(allData[i][ciEN_LAST] || '').trim().toUpperCase();
      var firstAr = String(allData[i][ciAR_FIRST] || '').trim();
      var lastAr = String(allData[i][ciAR_LAST] || '').trim();

      var combinedEn = (firstEn + ' ' + lastEn).trim();
      var combinedAr = (firstAr + ' ' + lastAr).trim();

      if (combinedEn && combinedEn === cleanedName) {
        return {
          row: i + 2,
          name_en: combinedEn,
          name_ar: combinedAr,
          url: String(allData[i][ciURL] || '').trim(),
          bj: String(allData[i][ciBJ] || '').trim()
        };
      }

      // مطابقة احتياطية بالعربي (نادرة — التذاكر عادة بالإنجليزي)
      if (combinedAr && combinedAr === cleanedName) {
        return {
          row: i + 2,
          name_en: combinedEn,
          name_ar: combinedAr,
          url: String(allData[i][ciURL] || '').trim(),
          bj: String(allData[i][ciBJ] || '').trim()
        };
      }
    }
    return null;
  }
};
