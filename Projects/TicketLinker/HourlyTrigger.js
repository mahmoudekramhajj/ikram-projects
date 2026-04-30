/**
 * HourlyTrigger.js — تريجر ساعي لإكمال المراجعة اليدوية
 *
 * مهمته الوحيدة: فحص تبويب TL_UnresolvedTickets كل ساعة.
 * لو وُجدت صفوف PENDING_MANUAL بخانتي الجنسية والجواز ممتلئتين → يكمل المعالجة.
 */

function hourlyCheckUnresolvedTickets() {
  if (!acquireLock_()) {
    Logger.log('Another run in progress, skipping hourly check');
    return;
  }
  try {
    var result = processFilledUnresolvedRows_();
    Logger.log('Hourly check: ' + JSON.stringify(result));
    return result;
  } finally {
    releaseLock_();
  }
}

/**
 * يثبّت التريجر الساعي
 */
function installHourlyTrigger() {
  // احذف أي triggers قديمة لنفس الدالة
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'hourlyCheckUnresolvedTickets') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // ثبّت تريجر جديد كل ساعة
  ScriptApp.newTrigger('hourlyCheckUnresolvedTickets')
    .timeBased()
    .everyHours(1)
    .create();

  return { ok: true, message: 'Hourly trigger installed' };
}

/**
 * يعرض كل التريجرات الحالية (للتشخيص)
 */
function listTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  return triggers.map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      type: t.getEventType(),
      uniqueId: t.getUniqueId()
    };
  });
}

/**
 * يحذف كل التريجرات (للاختبار)
 */
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
  return { deleted: triggers.length };
}

/**
 * يثبّت تريجر كل 30 دقيقة لمعالجة التذاكر الجديدة تلقائياً
 */
function installNewTicketsTrigger() {
  // احذف أي trigger قديم لنفس الدالة
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scheduledProcessNewTickets') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('scheduledProcessNewTickets')
    .timeBased()
    .everyMinutes(30)
    .create();

  return { ok: true, message: 'Trigger installed: scheduledProcessNewTickets every 30 minutes' };
}
