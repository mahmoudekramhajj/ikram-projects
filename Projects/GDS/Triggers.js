/**
 * Triggers.js — إدارة المشغّلات الزمنية
 *
 * المشغّلات:
 *   1. كل 30 دقيقة: runPipeline()
 *   2. يومياً الساعة 23 (11 مساءً): sendDailyReport()
 *
 * دوال عامة:
 *   installAllTriggers()   — تفعيل كل المشغّلات
 *   removeAllTriggers()    — إزالة كل مشغّلات GDS
 *   listTriggers()         — قائمة المشغّلات الحالية
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Triggers = {
  PIPELINE_FN: 'runPipeline',
  REPORT_FN: 'sendDailyReport',
  PIPELINE_EVERY_MIN: 30,
  REPORT_HOUR: 23,

  install: function() {
    // إزالة أي مشغّلات قديمة أولاً (لمنع التكرار)
    GDS2.Triggers.removeAll();

    // مشغّل Pipeline: كل 30 دقيقة
    ScriptApp.newTrigger(GDS2.Triggers.PIPELINE_FN)
      .timeBased()
      .everyMinutes(GDS2.Triggers.PIPELINE_EVERY_MIN)
      .create();

    // مشغّل تقرير يومي: الساعة 23
    ScriptApp.newTrigger(GDS2.Triggers.REPORT_FN)
      .timeBased()
      .everyDays(1)
      .atHour(GDS2.Triggers.REPORT_HOUR)
      .create();

    GDS2.Log.info('Triggers: installed');
    return GDS2.Triggers.list();
  },

  removeAll: function() {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    for (var i = 0; i < triggers.length; i++) {
      var t = triggers[i];
      var fn = t.getHandlerFunction();
      if (fn === GDS2.Triggers.PIPELINE_FN || fn === GDS2.Triggers.REPORT_FN) {
        ScriptApp.deleteTrigger(t);
        removed++;
      }
    }
    GDS2.Log.info('Triggers: removed', { count: removed });
    return { status: 'ok', removed: removed };
  },

  list: function() {
    var triggers = ScriptApp.getProjectTriggers();
    var result = [];
    for (var i = 0; i < triggers.length; i++) {
      var t = triggers[i];
      result.push({
        id: t.getUniqueId(),
        handler: t.getHandlerFunction(),
        event: t.getEventType().toString(),
        source: t.getTriggerSource().toString()
      });
    }
    return { count: result.length, triggers: result };
  }
};

// ==================== Global Entry Points ====================

function installAllTriggers() {
  return GDS2.Triggers.install();
}

function removeAllTriggers() {
  return GDS2.Triggers.removeAll();
}

function listTriggers() {
  return GDS2.Triggers.list();
}

/**
 * دالة التقرير اليومي (تُستدعى من مشغّل يومي الساعة 23).
 */
function sendDailyReport() {
  var status = GDS2.Pipeline.getStatus();
  var lastRun = status.last_run || {};

  var pendingCount = GDS2.Triggers._countPending();

  var stats = {
    date: Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd'),
    success_count: lastRun.success_count || 0,
    failure_count: lastRun.failure_count || 0,
    disaster_count: lastRun.disaster_count || 0,
    deletion_count: 0,
    nusuk_count: lastRun.nusuk_auth_count || 0,
    skipped_count: (lastRun.skipped_already || 0) + (lastRun.skipped_manual || 0) + (lastRun.skipped_max_fails || 0),
    cost_usd: GDS2.Triggers._estimateCost(lastRun.total_tokens),
    avg_sec: lastRun.rows_scanned ? (lastRun.duration_sec || 0) / lastRun.rows_scanned : 0,
    total_processed: (lastRun.success_count || 0) + (lastRun.family_applied || 0),
    pending: pendingCount
  };

  return GDS2.TelegramNotifier.sendDailyReport(stats);
}

GDS2.Triggers._countPending = function() {
  try {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    // BJ فارغ + BM ليس TRUE = قيد الانتظار
    var data = sheet.getRange(2, GDS2.Config.COL.LAST_URL, lastRow - 1, 4).getValues();
    var pending = 0;
    for (var i = 0; i < data.length; i++) {
      var bj = String(data[i][0] || '').trim();
      var bm = data[i][3];
      if (!bj && bm !== true && String(bm).toUpperCase() !== 'TRUE') pending++;
    }
    return pending;
  } catch (e) {
    return -1;
  }
};

GDS2.Triggers._estimateCost = function(tokens) {
  if (!tokens) return 0;
  // Claude Haiku 4.5: $1/M input, $5/M output (تقريبي)
  var input = (tokens.input || 0) * 1 / 1000000;
  var output = (tokens.output || 0) * 5 / 1000000;
  return input + output;
};
