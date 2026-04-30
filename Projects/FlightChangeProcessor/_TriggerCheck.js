/**
 * _TriggerCheck.js — فحص قراءة-فقط لـ triggers المشروع
 * أُضيف 2026-04-28 ضمن خطة إنقاذ FCP (المرحلة 0).
 *
 * كيف تستخدمها:
 *   1) افتح المشروع في Apps Script Editor.
 *   2) من القائمة اختر الدالة listTriggers ثم Run.
 *   3) View → Logs لقراءة النتيجة.
 *
 * لا تكتب شيئاً، لا تُعدّل أي trigger. للحذف استخدم deleteAllTriggers يدوياً
 * بعد قراءة القائمة والتأكد.
 */

function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('====== FCP Trigger Check — ' + new Date().toISOString() + ' ======');
  Logger.log('READ_ONLY_MODE = ' + CONFIG.READ_ONLY_MODE);
  Logger.log('عدد triggers: ' + triggers.length);

  if (triggers.length === 0) {
    Logger.log('✅ لا يوجد أي trigger نشط — التطبيق متوقف فعلياً.');
    return { count: 0, triggers: [] };
  }

  var summary = [];
  for (var i = 0; i < triggers.length; i++) {
    var tr = triggers[i];
    var info = {
      idx: i + 1,
      handler: tr.getHandlerFunction(),
      type: String(tr.getEventType()),
      source: String(tr.getTriggerSource()),
      uid: tr.getUniqueId()
    };
    summary.push(info);
    Logger.log(
      '#' + info.idx + ' → handler=' + info.handler +
      ' | type=' + info.type +
      ' | source=' + info.source +
      ' | uid=' + info.uid
    );
  }
  Logger.log('⚠️ راجع آخر executions في Apps Script → Executions، ثم نفّذ deleteAllTriggers لو لزم.');
  return { count: triggers.length, triggers: summary };
}

/**
 * حذف كامل لكل triggers المشروع. لا تشغّلها إلا بعد listTriggers + قرار صريح.
 * تُسجَّل النتيجة في Logs قبل وبعد الحذف.
 */
function deleteAllTriggers() {
  var before = ScriptApp.getProjectTriggers();
  Logger.log('قبل الحذف: ' + before.length + ' trigger');
  for (var i = 0; i < before.length; i++) {
    Logger.log('  حذف: ' + before[i].getHandlerFunction() + ' (uid=' + before[i].getUniqueId() + ')');
    ScriptApp.deleteTrigger(before[i]);
  }
  var after = ScriptApp.getProjectTriggers();
  Logger.log('بعد الحذف: ' + after.length + ' trigger');
  return { deleted: before.length, remaining: after.length };
}
