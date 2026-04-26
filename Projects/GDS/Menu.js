/**
 * Menu.js — قائمة GDS الجديدة في Google Sheets
 *
 * ملاحظة: لا نعرّف onOpen() هنا لتجنّب التعارض مع الكود القديم.
 * بدلاً من ذلك نستخدم installable trigger لدالة showGDSMenu().
 *
 * الاستخدام لمرة واحدة:
 *   شغّل installGDSMenu() من GAS Editor أو عبر ClaudeAPI.
 *   بعدها عند فتح الشيت → تظهر قائمة "GDS" إضافية.
 *
 * عند حذف الكود القديم (Phase 9): نُعيد تسمية showGDSMenu إلى onOpen.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Menu = {
  /**
   * بناء قائمة GDS. تُستدعى من trigger عند فتح الشيت.
   */
  show: function() {
    try {
      SpreadsheetApp.getUi().createMenu('✈️ GDS')
        .addItem('📊 تقرير الحالة', 'gdsShowStatus')
        .addSeparator()
        .addSubMenu(SpreadsheetApp.getUi().createMenu('👤 مزامنة')
          .addItem('👁️ معاينة الأسماء', 'previewSyncNames')
          .addItem('▶️ تشغيل مزامنة PD → B2C', 'syncNames')
          .addSeparator()
          .addItem('👁️ معاينة الأيتام', 'previewOrphans')
          .addItem('🗑️ حذف الأيتام', 'deleteOrphans')
        )
        .addSubMenu(SpreadsheetApp.getUi().createMenu('🎫 معالجة التذاكر')
          .addItem('🐤 Canary (10 فقط)', 'runCanary')
          .addItem('▶️ تشغيل Pipeline كامل', 'runPipeline')
          .addSeparator()
          .addItem('🔍 حالة آخر تشغيل', 'gdsShowLastRun')
        )
        .addSubMenu(SpreadsheetApp.getUi().createMenu('📋 تقارير')
          .addItem('⚠️ الحجاج المتوقفين (BL=2)', 'gdsShowFailures')
          .addItem('✍️ الحجاج المُعدَّلين يدوياً', 'gdsShowManual')
          .addItem('🔒 الحجاج المحتاجين دخول نسك', 'gdsShowNusuk')
          .addItem('⏳ الحجاج المعلّقين (بدون معالجة)', 'gdsShowPending')
        )
        .addSubMenu(SpreadsheetApp.getUi().createMenu('⏰ Triggers')
          .addItem('✅ تفعيل التشغيل التلقائي', 'installAllTriggers')
          .addItem('📋 عرض الـ Triggers', 'gdsListTriggers')
          .addItem('❌ إيقاف التشغيل التلقائي', 'removeAllTriggers')
        )
        .addSubMenu(SpreadsheetApp.getUi().createMenu('🧠 السجلات')
          .addItem('📝 تحديث IATA من البيانات', 'extractIATAFromHistory')
          .addItem('🧹 تنظيف IATA noise', 'cleanupIATANoise')
          .addItem('📇 عرض قائمة IATA', 'gdsShowIATA')
          .addItem('✈️ عرض شركات الطيران', 'gdsShowAirlines')
        )
        .addSubMenu(SpreadsheetApp.getUi().createMenu('💬 Telegram')
          .addItem('✅ اختبار الاتصال', 'testTelegramMessage')
          .addItem('🔍 فحص الإعداد', 'gdsCheckTelegram')
        )
        .addToUi();
    } catch (e) {
      GDS2.Log.error('Menu.show failed', { error: e.message });
    }
  },

  /**
   * تثبيت trigger مرة واحدة يعمل عند كل فتح للشيت.
   */
  install: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var scriptId = ScriptApp.getScriptId();

    // إزالة أي triggers قديمة لـ showGDSMenu (منع التكرار)
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'showGDSMenu') {
        ScriptApp.deleteTrigger(triggers[i]);
        removed++;
      }
    }

    // تثبيت trigger جديد
    ScriptApp.newTrigger('showGDSMenu')
      .forSpreadsheet(ss)
      .onOpen()
      .create();

    // تشغيل الآن لمرة أولى (لو المستخدم يشاهد الشيت)
    try { GDS2.Menu.show(); } catch (e) {}

    return {
      status: 'ok',
      removed_old: removed,
      message: 'عند فتح الشيت القادمة، ستظهر قائمة GDS'
    };
  }
};

// ==================== Global Entry Points ====================

/**
 * Handler لـ installable onOpen trigger.
 */
function showGDSMenu() {
  return GDS2.Menu.show();
}

/**
 * تثبيت قائمة GDS — يُشغَّل مرة واحدة من Editor أو ClaudeAPI.
 */
function installGDSMenu() {
  return GDS2.Menu.install();
}
