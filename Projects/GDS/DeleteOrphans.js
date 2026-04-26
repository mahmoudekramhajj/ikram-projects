/**
 * DeleteOrphans.js — حذف الحجاج الأيتام من B2C
 *
 * يُحذف صف B2C إذا:
 *   - الحاج (بالجواز) غير موجود في PD، أو
 *   - موجود في PD لكن ticketUrl فارغ (إشارة أنه لم يعد يحتاج معالجة)
 *
 * إشعارات تيليغرام مؤجَّلة للمرحلة 6.
 *
 * دوال عامة:
 *   deleteOrphans()   — التشغيل الفعلي
 *   previewOrphans()  — معاينة الحجاج المرشحين للحذف دون تعديل
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.DeleteOrphans = {
  /**
   * تنفيذ الحذف.
   */
  run: function() {
    var startTime = new Date();
    var analysis = GDS2.DeleteOrphans._analyze();
    if (analysis.error) return analysis;

    var b2cSheet = analysis.b2cSheet;
    var orphans = analysis.orphans;

    // حذف من الأعلى للأسفل (بدءاً من أعلى رقم صف)
    orphans.sort(function(a, b) { return b.row - a.row; });

    for (var i = 0; i < orphans.length; i++) {
      b2cSheet.deleteRow(orphans[i].row);
      GDS2.Log.info('DeleteOrphans: deleted', orphans[i]);
      // TODO (Phase 6): TelegramNotifier.sendDeletion(orphans[i]);
    }

    var elapsed = (new Date() - startTime) / 1000;
    return {
      status: 'ok',
      deleted: orphans.length,
      reasons: analysis.reasonsCount,
      orphans_sample: orphans.slice(0, 5),
      elapsed_sec: elapsed
    };
  },

  /**
   * معاينة دون أي حذف.
   */
  preview: function() {
    var analysis = GDS2.DeleteOrphans._analyze();
    if (analysis.error) return analysis;
    return {
      status: 'preview',
      will_delete: analysis.orphans.length,
      reasons: analysis.reasonsCount,
      orphans_sample: analysis.orphans.slice(0, 10)
    };
  },

  // ==================== Private ====================

  _analyze: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var pdSheet = ss.getSheetByName(GDS2.Config.SHEET_PD);
    var b2cSheet = ss.getSheetByName(GDS2.Config.SHEET_B2C);

    if (!pdSheet) return { error: 'PD sheet missing' };
    if (!b2cSheet) return { error: 'B2C sheet missing' };

    // بناء مجموعة PD valid: جوازات لها ticketUrl
    var pdValidSet = {};       // passport → true (في PD + له ticketUrl)
    var pdExistsSet = {};      // passport → true (في PD، بغض النظر عن ticketUrl)

    var pdLastRow = pdSheet.getLastRow();
    if (pdLastRow >= 2) {
      var pdData = pdSheet.getRange(2, 1, pdLastRow - 1, 26).getValues();
      for (var r = 0; r < pdData.length; r++) {
        var passport = String(pdData[r][5] || '').trim(); // F = 6
        if (!passport) continue;
        pdExistsSet[passport] = true;
        var url = String(pdData[r][25] || '').trim(); // Z = 26
        if (url) pdValidSet[passport] = true;
      }
    }

    // فحص B2C للأيتام
    var b2cLastRow = b2cSheet.getLastRow();
    if (b2cLastRow < 2) {
      return {
        b2cSheet: b2cSheet,
        orphans: [],
        reasonsCount: { not_in_pd: 0, no_ticket_url: 0 }
      };
    }

    // قراءة جواز + اسم لكل صف B2C (للإشعار)
    var b2cData = b2cSheet.getRange(2, 1, b2cLastRow - 1, 12).getValues();
    var orphans = [];
    var reasonsCount = { not_in_pd: 0, no_ticket_url: 0 };

    for (var i = 0; i < b2cData.length; i++) {
      var passport = String(b2cData[i][5] || '').trim();
      if (!passport) continue; // صفوف فارغة

      if (pdValidSet[passport]) continue; // حاج شرعي، لا حذف

      var reason;
      if (!pdExistsSet[passport]) {
        reason = 'not_in_pd';
        reasonsCount.not_in_pd++;
      } else {
        reason = 'no_ticket_url';
        reasonsCount.no_ticket_url++;
      }

      orphans.push({
        row: i + 2,
        passport: passport,
        name_ar: (String(b2cData[i][8] || '') + ' ' + String(b2cData[i][9] || '')).trim(),
        name_en: (String(b2cData[i][10] || '') + ' ' + String(b2cData[i][11] || '')).trim(),
        reason: reason
      });
    }

    return {
      b2cSheet: b2cSheet,
      orphans: orphans,
      reasonsCount: reasonsCount
    };
  }
};

// ==================== Global Entry Points ====================

function deleteOrphans() {
  return GDS2.DeleteOrphans.run();
}

function previewOrphans() {
  return GDS2.DeleteOrphans.preview();
}
