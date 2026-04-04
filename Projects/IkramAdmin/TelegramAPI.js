// ============================================
// Telegram API — دوال الإرسال
// ============================================

function sendMessage_(chatId, text, replyMarkup) {
  var payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText());
}

function editMessage_(chatId, messageId, text, replyMarkup) {
  var payload = {
    chat_id: chatId, message_id: messageId,
    text: text, parse_mode: 'HTML'
  };
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  UrlFetchApp.fetch(TELEGRAM_API + '/editMessageText', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
}

function answerCallback_(callbackQueryId, text) {
  var payload = { callback_query_id: callbackQueryId };
  if (text) payload.text = text;
  UrlFetchApp.fetch(TELEGRAM_API + '/answerCallbackQuery', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
}

function sendDocument_(chatId, docUrl, fileName, caption, replyMarkup) {
  var payload = { chat_id: chatId, document: docUrl, parse_mode: 'HTML' };
  if (caption) payload.caption = caption;
  if (fileName) payload.file_name = fileName;
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  UrlFetchApp.fetch(TELEGRAM_API + '/sendDocument', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
}

function sendMessageSafe_(chatId, text, replyMarkup) {
  try {
    var payload = { chat_id: chatId, text: text, parse_mode: 'HTML' };
    if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
    var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    if (json.ok) return { success: true, error: null };
    return { success: false, error: json.description || 'Unknown error' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// تحويل رابط Google Drive إلى رابط مباشر
function getDriveDirectUrl_(driveUrl) {
  if (!driveUrl) return '';
  var match = String(driveUrl).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return 'https://drive.google.com/uc?export=download&id=' + match[1];
  return driveUrl;
}
