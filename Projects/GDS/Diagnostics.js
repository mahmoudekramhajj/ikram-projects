/**
 * Diagnostics.js — دوال تشخيصية + عرض التقارير في واجهة المستخدم
 *
 * تقارير تُعرض في نوافذ منبثقة (Toast/Alert) من القائمة.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Diagnostics = {
  /**
   * تقرير الحالة العامة.
   */
  status: function() {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { total: 0 };

    // نقرأ: جواز (6), رابط (26), BJ (62), BK (63), BL (64), BM (65)
    var lastCol = GDS2.Config.COL.NUSUK_AUTH; // 65
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    var stats = {
      total: 0,
      with_url: 0,
      no_url: 0,
      processed: 0,
      pending: 0,
      manual: 0,
      stopped: 0,
      nusuk_auth: 0,
      fail_once: 0
    };

    for (var i = 0; i < data.length; i++) {
      var passport = String(data[i][GDS2.Config.COL.PASSPORT - 1] || '').trim();
      if (!passport) continue;
      stats.total++;

      var url = String(data[i][GDS2.Config.COL.TICKET_URL - 1] || '').trim();
      var bj = String(data[i][GDS2.Config.COL.LAST_URL - 1] || '').trim();
      var bk = data[i][GDS2.Config.COL.MANUAL - 1];
      var bl = Number(data[i][GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      var bm = data[i][GDS2.Config.COL.NUSUK_AUTH - 1];

      if (url) stats.with_url++;
      else stats.no_url++;

      if (!url) continue; // صفوف بلا URL: لا تُصنَّف في pending/processed

      if (bk === true || String(bk).toUpperCase() === 'TRUE') {
        stats.manual++;
      } else if (bm === true || String(bm).toUpperCase() === 'TRUE') {
        stats.nusuk_auth++;
      } else if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) {
        stats.stopped++;
      } else if (bj && bj === url) {
        stats.processed++;
      } else if (bl === 1) {
        stats.fail_once++;
        stats.pending++;
      } else {
        stats.pending++;
      }
    }

    return stats;
  },

  /**
   * قائمة بالحجاج الذين توقفت معالجتهم (BL >= max).
   */
  failures: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var bl = Number(row[GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      return bl >= GDS2.Config.MAX_FAIL_ATTEMPTS;
    });
  },

  manual: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var bk = row[GDS2.Config.COL.MANUAL - 1];
      return bk === true || String(bk).toUpperCase() === 'TRUE';
    });
  },

  nusuk: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var bm = row[GDS2.Config.COL.NUSUK_AUTH - 1];
      return bm === true || String(bm).toUpperCase() === 'TRUE';
    });
  },

  pending: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var url = String(row[GDS2.Config.COL.TICKET_URL - 1] || '').trim();
      var bj = String(row[GDS2.Config.COL.LAST_URL - 1] || '').trim();
      var bk = row[GDS2.Config.COL.MANUAL - 1];
      var bl = Number(row[GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      var bm = row[GDS2.Config.COL.NUSUK_AUTH - 1];

      if (!url) return false;
      if (bk === true || String(bk).toUpperCase() === 'TRUE') return false;
      if (bm === true || String(bm).toUpperCase() === 'TRUE') return false;
      if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) return false;
      if (bj === url) return false;
      return true;
    });
  },

  // ==================== UI helpers ====================

  alertStatus: function() {
    var s = GDS2.Diagnostics.status();
    var msg = [
      '📊 إحصائيات B2C',
      '',
      'إجمالي الصفوف: ' + s.total,
      'بـ ticketUrl: ' + s.with_url,
      '',
      '✅ معالجة ناجحة: ' + s.processed,
      '⏳ معلّقة: ' + s.pending,
      '🟡 فشل مرة واحدة: ' + s.fail_once,
      '🔴 متوقفة نهائياً: ' + s.stopped,
      '✍️ مُعدَّلة يدوياً: ' + s.manual,
      '🔒 تحتاج دخول نسك: ' + s.nusuk_auth
    ].join('\n');
    GDS2.Diagnostics._alert(msg, 'تقرير الحالة');
  },

  alertLastRun: function() {
    var status = GDS2.Pipeline.getStatus();
    if (!status.last_run) {
      GDS2.Diagnostics._alert('لم يُشغَّل Pipeline بعد.', 'آخر تشغيل');
      return;
    }
    var r = status.last_run;
    var msg = [
      '📋 آخر تشغيل',
      '',
      'بدأ: ' + r.started_at,
      'انتهى: ' + r.finished_at,
      'المدة: ' + (r.duration_sec || 0).toFixed(1) + ' ثانية',
      '',
      'صفوف مُفحوصة: ' + r.rows_scanned,
      '✅ نجحت: ' + r.success_count,
      '👨‍👩‍👧 عائلة: ' + r.family_applied,
      '❌ فشلت: ' + r.failure_count,
      '🚨 كوارث: ' + r.disaster_count,
      '🔒 nusuk auth: ' + r.nusuk_auth_count,
      '⏭️ متخطاة: ' + r.skipped_already,
      (r.timed_out ? '⏰ توقف عند الحد الزمني' : '')
    ].filter(Boolean).join('\n');
    GDS2.Diagnostics._alert(msg, 'آخر تشغيل');
  },

  alertList: function(list, title) {
    if (!list || list.length === 0) {
      GDS2.Diagnostics._alert('لا توجد صفوف مطابقة.', title);
      return;
    }
    var lines = [title + ' (' + list.length + ' حاج)', ''];
    var max = Math.min(list.length, 30);
    for (var i = 0; i < max; i++) {
      var item = list[i];
      lines.push('صف ' + item.row + ': ' + item.name + ' (' + item.passport + ')');
    }
    if (list.length > max) {
      lines.push('');
      lines.push('... و ' + (list.length - max) + ' آخرين');
    }
    GDS2.Diagnostics._alert(lines.join('\n'), title);
  },

  alertIATA: function() {
    var list = GDS2.IATARegistry.getList();
    var msg = [
      'قائمة IATA (' + list.length + ' كود)',
      '',
      list.join(', ')
    ].join('\n');
    GDS2.Diagnostics._alert(msg, 'IATA');
  },

  alertAirlines: function() {
    var list = GDS2.AirlineRegistry.getList();
    var lines = ['شركات الطيران (' + list.length + ')', ''];
    for (var i = 0; i < list.length; i++) {
      lines.push(list[i].code + ' — ' + list[i].name);
    }
    GDS2.Diagnostics._alert(lines.join('\n'), 'Airlines');
  },

  alertTelegramConfig: function() {
    var cfg = GDS2.TelegramNotifier.checkConfig();
    var msg = [
      'حالة Telegram',
      '',
      'Token: ' + (cfg.token_set ? '✅ مُعدّ' : '❌ غير مُعدّ'),
      'Chat ID: ' + (cfg.chat_id_set ? '✅ ' + cfg.chat_id : '❌ غير مُعدّ'),
      'جاهز: ' + (cfg.ready ? '✅' : '❌')
    ].join('\n');
    GDS2.Diagnostics._alert(msg, 'Telegram');
  },

  // ==================== Private ====================

  _listByFilter: function(filterFn) {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, GDS2.Config.COL.NUSUK_AUTH).getValues();
    var results = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!filterFn(row)) continue;

      var passport = String(row[GDS2.Config.COL.PASSPORT - 1] || '').trim();
      if (!passport) continue;

      var firstEn = String(row[GDS2.Config.COL.FIRST_NAME_EN - 1] || '').trim();
      var lastEn = String(row[GDS2.Config.COL.LAST_NAME_EN - 1] || '').trim();
      var firstAr = String(row[GDS2.Config.COL.FIRST_NAME_AR - 1] || '').trim();
      var lastAr = String(row[GDS2.Config.COL.LAST_NAME_AR - 1] || '').trim();
      var name = (firstEn + ' ' + lastEn).trim() || (firstAr + ' ' + lastAr).trim() || '(بلا اسم)';

      results.push({
        row: i + 2,
        passport: passport,
        name: name
      });
    }

    return results;
  },

  _alert: function(msg, title) {
    try {
      SpreadsheetApp.getUi().alert(title || 'GDS', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // لا UI متاح (مثلاً عبر ClaudeAPI)
      Logger.log('[' + (title || 'GDS') + '] ' + msg);
    }
  }
};

// ==================== Global Entry Points ====================

function gdsShowStatus() { GDS2.Diagnostics.alertStatus(); }

function gdsShowLastRun() { GDS2.Diagnostics.alertLastRun(); }

function gdsShowFailures() {
  var list = GDS2.Diagnostics.failures();
  GDS2.Diagnostics.alertList(list, '🔴 حجاج متوقفون نهائياً (BL=2)');
}

function gdsShowManual() {
  var list = GDS2.Diagnostics.manual();
  GDS2.Diagnostics.alertList(list, '✍️ حجاج مُعدَّلون يدوياً');
}

function gdsShowNusuk() {
  var list = GDS2.Diagnostics.nusuk();
  GDS2.Diagnostics.alertList(list, '🔒 حجاج يحتاجون دخول نسك');
}

function gdsShowPending() {
  var list = GDS2.Diagnostics.pending();
  GDS2.Diagnostics.alertList(list, '⏳ حجاج معلّقون (بدون معالجة)');
}

function gdsListTriggers() {
  var t = GDS2.Triggers.list();
  var lines = ['Triggers (' + t.count + ')', ''];
  for (var i = 0; i < t.triggers.length; i++) {
    lines.push(t.triggers[i].handler + ' — ' + t.triggers[i].event);
  }
  GDS2.Diagnostics._alert(lines.join('\n'), 'Triggers');
}

function gdsShowIATA() { GDS2.Diagnostics.alertIATA(); }

function gdsShowAirlines() { GDS2.Diagnostics.alertAirlines(); }

function gdsCheckTelegram() { GDS2.Diagnostics.alertTelegramConfig(); }

// للاستدعاء عبر ClaudeAPI (يُرجع JSON)
function getStatusJSON() { return GDS2.Diagnostics.status(); }
function getPendingJSON() { return GDS2.Diagnostics.pending(); }
function getFailuresJSON() { return GDS2.Diagnostics.failures(); }
