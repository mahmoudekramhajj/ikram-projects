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
