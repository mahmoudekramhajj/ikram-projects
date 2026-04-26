/**
 * TelegramNotifier.js — إرسال إشعارات تيليغرام
 *
 * الاستخدام:
 *   1. setTelegramToken(token)      — حفظ BOT_TOKEN في Script Properties
 *   2. setTelegramChatId(chat_id)   — حفظ CHAT_ID للمجموعة المخصصة
 *   3. checkTelegramConfig()        — التحقق من الإعداد
 *   4. testTelegramMessage()        — إرسال رسالة اختبار
 *
 * الإرسال الفعلي يكون عبر Pipeline (فشل، كارثة، حذف، تقرير يومي).
 * إذا الإعداد غير مكتمل، الإشعارات تُتخطى بصمت (لا تُفشل المعالجة).
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.TelegramNotifier = {
  TELEGRAM_API: 'https://api.telegram.org/bot',
  RATE_LIMIT_MS: 100, // sleep بين الرسائل

  _lastSendTime: 0,

  /**
   * إرسال رسالة خام.
   * @return {Object} { status, message_id?, reason? }
   */
  send: function(text, options) {
    options = options || {};
    var token = GDS2.State.get(GDS2.Config.PROP.TELEGRAM_BOT_TOKEN);
    var chatId = GDS2.State.get(GDS2.Config.PROP.TELEGRAM_CHAT_ID);

    if (!token || !chatId) {
      return { status: 'skipped', reason: 'not_configured' };
    }

    GDS2.TelegramNotifier._enforceRateLimit();

    try {
      var response = UrlFetchApp.fetch(
        GDS2.TelegramNotifier.TELEGRAM_API + token + '/sendMessage',
        {
          method: 'post',
          payload: {
            chat_id: chatId,
            text: text,
            parse_mode: options.parseMode || '',
            disable_web_page_preview: 'true'
          },
          muteHttpExceptions: true
        }
      );

      GDS2.TelegramNotifier._lastSendTime = new Date().getTime();

      var code = response.getResponseCode();
      var body = response.getContentText();

      if (code === 200) {
        var json = JSON.parse(body);
        return { status: 'ok', message_id: json.result && json.result.message_id };
      }

      return { status: 'error', http_code: code, body: body.substring(0, 300) };
    } catch (e) {
      return { status: 'error', exception: e.message };
    }
  },

  // ==================== رسائل جاهزة ====================

  sendFailure: function(pilgrim, reason, detail) {
    var text = [
      '❌ فشل معالجة تذكرة',
      '',
      '🧳 الحاج: ' + (pilgrim.name || 'غير محدد'),
      '📔 الجواز: ' + (pilgrim.passport || '—'),
      '🔍 السبب: ' + (reason || 'غير محدد'),
      detail ? '📋 التفاصيل: ' + (typeof detail === 'string' ? detail : JSON.stringify(detail).substring(0, 200)) : '',
      '',
      '🔗 رابط التذكرة:',
      pilgrim.url || '—'
    ].filter(Boolean).join('\n');
    return GDS2.TelegramNotifier.send(text);
  },

  sendDisaster: function(pilgrim, disaster) {
    var passengers = (disaster.passengers_in_ticket || []).join(', ');
    var text = [
      '🚨 كارثة — الرابط لشخص آخر!',
      '',
      '🧳 الحاج: ' + (pilgrim.name || 'غير محدد'),
      '📔 الجواز: ' + (pilgrim.passport || '—'),
      '📋 الركاب الموجودون في التذكرة:',
      passengers ? '   ' + passengers : '   (لا يوجد)',
      '',
      '⚠️ راجع الرابط في Presonal Details — قد يكون مخلوطاً مع حاج آخر',
      '',
      '🔗 الرابط:',
      pilgrim.url || '—'
    ].join('\n');
    return GDS2.TelegramNotifier.send(text);
  },

  sendDeletion: function(orphan) {
    var text = [
      '🗑️ حذف تلقائي من B2C',
      '',
      '🧳 الحاج: ' + (orphan.name_en || orphan.name_ar || '—'),
      '📔 الجواز: ' + (orphan.passport || '—'),
      '🔍 السبب: ' + (orphan.reason === 'not_in_pd' ? 'غير موجود في Presonal Details (احتمال إلغاء)' : 'موجود في PD لكن بدون رابط تذكرة')
    ].join('\n');
    return GDS2.TelegramNotifier.send(text);
  },

  sendNusukAuth: function(pilgrim) {
    var text = [
      '🔒 رابط نسك محمي',
      '',
      '🧳 الحاج: ' + (pilgrim.name || '—'),
      '📔 الجواز: ' + (pilgrim.passport || '—'),
      '📋 الرابط يحتاج تسجيل دخول لنسك',
      '⏸️ المعالجة متوقفة — راجع يدوياً',
      '',
      '🔗 رابط التذكرة:',
      pilgrim.url || '—'
    ].join('\n');
    return GDS2.TelegramNotifier.send(text);
  },

  sendRisks: function(pilgrim, riskText) {
    if (!riskText) return { status: 'skipped', reason: 'no_risks' };
    return GDS2.TelegramNotifier.send('⚠️ ملاحظات معالجة\n\n' + riskText);
  },

  sendDailyReport: function(stats) {
    var text = [
      '📊 تقرير GDS اليومي — ' + stats.date,
      '',
      '✅ نجحت: ' + (stats.success_count || 0),
      '❌ فشلت: ' + (stats.failure_count || 0),
      '🚨 كوارث: ' + (stats.disaster_count || 0),
      '🗑️ محذوفة: ' + (stats.deletion_count || 0),
      '🔒 نسك auth: ' + (stats.nusuk_count || 0),
      '⏭️ مُتخطّاة: ' + (stats.skipped_count || 0),
      '',
      '💰 تكلفة اليوم: $' + (stats.cost_usd || 0).toFixed(4),
      '⏱️ متوسط المعالجة: ' + (stats.avg_sec || 0).toFixed(1) + 'ث',
      '',
      '📈 إجمالي المعالجة حتى الآن: ' + (stats.total_processed || 0),
      '📋 المتبقي بدون معالجة: ' + (stats.pending || 0)
    ].join('\n');
    return GDS2.TelegramNotifier.send(text);
  },

  // ==================== Setup ====================

  setToken: function(token) {
    if (!token || typeof token !== 'string' || token.length < 20) {
      return { status: 'error', reason: 'invalid_token' };
    }
    GDS2.State.set(GDS2.Config.PROP.TELEGRAM_BOT_TOKEN, token.trim());
    return { status: 'ok', length: token.length };
  },

  setChatId: function(chatId) {
    if (!chatId) return { status: 'error', reason: 'empty_chat_id' };
    GDS2.State.set(GDS2.Config.PROP.TELEGRAM_CHAT_ID, String(chatId).trim());
    return { status: 'ok', chat_id: String(chatId).trim() };
  },

  checkConfig: function() {
    var token = GDS2.State.get(GDS2.Config.PROP.TELEGRAM_BOT_TOKEN);
    var chatId = GDS2.State.get(GDS2.Config.PROP.TELEGRAM_CHAT_ID);
    return {
      token_set: !!token,
      chat_id_set: !!chatId,
      chat_id: chatId || null,
      token_prefix: token ? token.substring(0, 10) + '...' : null,
      ready: !!token && !!chatId
    };
  },

  testMessage: function() {
    return GDS2.TelegramNotifier.send(
      '✅ اختبار اتصال GDS\n\n' +
      'الوقت: ' + new Date().toISOString() + '\n' +
      'النظام جاهز لإرسال الإشعارات.'
    );
  },

  // ==================== Private ====================

  _enforceRateLimit: function() {
    var now = new Date().getTime();
    var elapsed = now - GDS2.TelegramNotifier._lastSendTime;
    if (GDS2.TelegramNotifier._lastSendTime > 0 && elapsed < GDS2.TelegramNotifier.RATE_LIMIT_MS) {
      Utilities.sleep(GDS2.TelegramNotifier.RATE_LIMIT_MS - elapsed);
    }
  }
};

// ==================== Global Entry Points ====================

function setTelegramToken(token) {
  return GDS2.TelegramNotifier.setToken(token);
}

function setTelegramChatId(chatId) {
  return GDS2.TelegramNotifier.setChatId(chatId);
}

function checkTelegramConfig() {
  return GDS2.TelegramNotifier.checkConfig();
}

function testTelegramMessage() {
  return GDS2.TelegramNotifier.testMessage();
}
