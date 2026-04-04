// ============================================
// Telegram API — دوال الإرسال
// ============================================

function sendMessage_(chatId, text, replyMarkup) {
  var payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'HTML'
  };
  if (replyMarkup) {
    payload.reply_markup = JSON.stringify(replyMarkup);
  }
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('TG Response: ' + res.getContentText());
}

function answerCallback_(callbackQueryId) {
  UrlFetchApp.fetch(TELEGRAM_API + '/answerCallbackQuery', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ callback_query_id: callbackQueryId }),
    muteHttpExceptions: true
  });
}

// ============================================
// إرسال صورة عبر URL
// ============================================
function sendPhoto_(chatId, photoUrl, caption, replyMarkup) {
  var payload = {
    chat_id: chatId,
    photo: photoUrl,
    parse_mode: 'HTML'
  };
  if (caption) payload.caption = caption;
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendPhoto', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  Logger.log('TG Photo Response: ' + res.getContentText());
}

// ============================================
// إرسال ملف (PDF/Document) عبر URL
// ============================================
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

// ============================================
// إرسال آمن — يُرجع {success, error} بدل throw
// ============================================
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

function sendPhotoSafe_(chatId, photoUrl, caption, replyMarkup) {
  try {
    var payload = { chat_id: chatId, photo: photoUrl, parse_mode: 'HTML' };
    if (caption) payload.caption = caption;
    if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
    var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendPhoto', {
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

function sendDocumentSafe_(chatId, docUrl, fileName, caption, replyMarkup) {
  try {
    var payload = { chat_id: chatId, document: docUrl, parse_mode: 'HTML' };
    if (caption) payload.caption = caption;
    if (fileName) payload.file_name = fileName;
    if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
    var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendDocument', {
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

// ============================================
// تحويل رابط Google Drive إلى رابط مباشر
// ============================================
function getDriveDirectUrl_(driveUrl) {
  if (!driveUrl) return '';
  var match = String(driveUrl).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return 'https://drive.google.com/uc?export=download&id=' + match[1];
  return driveUrl;
}

// ============================================
// جلب معلومات ملف من تيليغرام (للصور)
// ============================================
function getFile_(fileId) {
  try {
    var res = UrlFetchApp.fetch(TELEGRAM_API + '/getFile', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ file_id: fileId }),
      muteHttpExceptions: true
    });
    var json = JSON.parse(res.getContentText());
    if (json.ok && json.result) {
      return json.result;
    }
    return null;
  } catch (e) {
    Logger.log('getFile_ error: ' + e.message);
    return null;
  }
}
