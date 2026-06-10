/**
 * SeasonStop.js — إيقاف أي مشغّل آلي لانتهاء الموسم (2026-06-10)
 *
 * يحذف كل مشغّلات الوقت في هذا المشروع فيتوقف أي عمل تلقائي.
 * للاستئناف: أعد تشغيل دالة الإعداد الأصلية للمشروع.
 */
function removeAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = [];
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    ScriptApp.deleteTrigger(triggers[i]);
    removed.push(handler);
  }
  Logger.log('removed ' + removed.length + ' trigger(s): ' + removed.join(', '));
  return { removed_count: removed.length, removed: removed };
}

function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers().map(function(t) {
    return { handler: t.getHandlerFunction(), source: String(t.getTriggerSource()) };
  });
  return { count: triggers.length, triggers: triggers };
}
