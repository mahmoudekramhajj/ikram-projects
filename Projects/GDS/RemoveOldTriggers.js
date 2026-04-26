/**
 * RemoveOldTriggers.js — حذف triggers من الكود القديم فقط
 *
 * يحذف:
 *   - autoCheckNewRows (old)
 *   - processPDFBatch (old)
 *   - onOpen (old - إن كان installable)
 *   - stopPDFTrigger_ (old, احتياط)
 *   - setupTriggers (old, احتياط)
 *
 * يحافظ على:
 *   - runPipeline (جديد)
 *   - sendDailyReport (جديد)
 *   - showGDSMenu (جديد)
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.RemoveOldTriggers = {
  OLD_HANDLERS: [
    'autoCheckNewRows',
    'processPDFBatch',
    'onOpen',
    'stopPDFTrigger_',
    'setupTriggers'
  ],

  NEW_HANDLERS: ['runPipeline', 'sendDailyReport', 'showGDSMenu'],

  removeOldOnly: function() {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = [];
    var kept = [];

    for (var i = 0; i < triggers.length; i++) {
      var t = triggers[i];
      var handler = t.getHandlerFunction();

      if (GDS2.RemoveOldTriggers.OLD_HANDLERS.indexOf(handler) !== -1) {
        ScriptApp.deleteTrigger(t);
        removed.push(handler);
      } else {
        kept.push(handler);
      }
    }

    GDS2.Log.info('RemoveOldTriggers', { removed: removed, kept: kept });
    return {
      status: 'ok',
      removed_count: removed.length,
      removed: removed,
      kept_count: kept.length,
      kept: kept
    };
  }
};

// ==================== Global Entry Points ====================

function removeOldTriggers() {
  return GDS2.RemoveOldTriggers.removeOldOnly();
}
