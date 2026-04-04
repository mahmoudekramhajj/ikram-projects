// ============================================
// Setup.js — إعداد الـ Webhook واختبار البوت
// ============================================

/**
 * إعداد Webhook لتيليغرام
 * شغّل هذه الدالة مرة واحدة بعد clasp deploy
 */
function setWebhook() {
  var url = ScriptApp.getService().getUrl();
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: url }),
    muteHttpExceptions: true
  });
  Logger.log('Webhook URL: ' + url);
  Logger.log('Response: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

/**
 * حذف Webhook
 */
function deleteWebhook() {
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/deleteWebhook', {
    method: 'post',
    muteHttpExceptions: true
  });
  Logger.log('deleteWebhook: ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

/**
 * معلومات Webhook الحالي
 */
function getWebhookInfo() {
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/getWebhookInfo', {
    muteHttpExceptions: true
  });
  var info = JSON.parse(res.getContentText());
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

/**
 * إرسال رسالة اختبار للأدمن
 */
function testBot() {
  var chatId = ADMIN_IDS[0];
  sendMessage_(chatId, '✅ <b>IkramAdmin Bot</b> يعمل بنجاح!\n\n' +
    '📅 ' + new Date().toLocaleString('ar-SA', { timeZone: 'Asia/Riyadh' }) + '\n' +
    '🔗 ' + ScriptApp.getService().getUrl());
  Logger.log('Test message sent to: ' + chatId);
}

/**
 * تسجيل المشروع في Claude Registry
 */
function registerBot() {
  return registerProject_();
}
