/**
 * _Setup.js — إعدادات لمرة واحدة
 *
 * شغّل هذه الدوال من المحرر مرة واحدة ثم احذف المحتوى
 */


/**
 * ضع مفتاح Claude API هنا مرة واحدة ثم شغّل الدالة من المحرر
 * بعد النجاح، امسح القيمة من هنا واحفظ الملف
 */
function setupClaudeKey() {
  var API_KEY = 'ALREADY_SET_VIA_SCRIPT_PROPERTIES';

  if (!API_KEY || API_KEY === 'PASTE_KEY_HERE') {
    throw new Error('ضع المفتاح أولاً');
  }

  PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY', API_KEY);
  Logger.log('✅ تم حفظ CLAUDE_API_KEY في Script Properties');
  Logger.log('الآن: امسح القيمة من هذا الملف للحماية');
  return { ok: true };
}


/**
 * إعادة فتح الإيميلات التي تم تخطّيها (Skipped) لإعادة معالجتها بـ V5
 *
 * @param {Object} args - {limit: كم thread} (اختياري)
 */
function reprocessSkipped(args) {
  args = args || {};
  var limit = args.limit || 50;

  var skippedLabel = getOrCreateLabel_(CONFIG.SKIPPED_LABEL);
  var processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  var query = 'label:' + CONFIG.SKIPPED_LABEL;
  var threads = GmailApp.search(query, 0, limit);

  Logger.log('🔄 إزالة Skipped + Processed من ' + threads.length + ' thread');

  for (var i = 0; i < threads.length; i++) {
    threads[i].removeLabel(skippedLabel);
    threads[i].removeLabel(processedLabel);
  }

  return { removedFrom: threads.length, message: 'شغّل scanEmailsV3 الآن' };
}


/**
 * إنشاء trigger يعمل تلقائياً كل ساعة
 */
function setupAutoTrigger() {
  // احذف أي trigger قديم لنفس الدالة
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scanEmailsV3') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // أنشئ trigger جديد كل ساعة
  ScriptApp.newTrigger('scanEmailsV3')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ Trigger تلقائي: scanEmailsV3 كل ساعة');
  return { ok: true, frequency: 'hourly' };
}


/**
 * إلغاء Trigger التلقائي
 */
function removeAutoTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scanEmailsV3') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return { ok: true, removed: removed };
}


/**
 * عرض Triggers الحالية
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  return triggers.map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      type: t.getEventType().toString(),
      source: t.getTriggerSource().toString()
    };
  });
}
