/**
 * SeasonStop.js — إيقاف المزامنة الآلية لانتهاء الموسم (2026-06-10)
 *
 * يحذف مشغّل الوقت (syncCheck كل 5 دقائق) فيتوقف نسخ الملفات تلقائياً.
 * لا يلمس سجلات الربط (map_/foldermap_) — تبقى للاستئناف الموسم القادم.
 *
 * للاستئناف: شغّل setup() لإعادة إنشاء المشغّل.
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
