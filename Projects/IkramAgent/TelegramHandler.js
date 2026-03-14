/**
 * معالج Telegram لوكيل إكرام
 * الإصدار 2.0
 */

// رابط Web App المنشور
var WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbyaoTbo3iLa0ibr-UfUElR3DC1howS-m7s7Mzpl7muS4sArAm3ybTgK9YBfg_pgW5k/exec';

/**
 * معالجة طلبات GET
 */
function doGet(e) {
  return ContentService.createTextOutput('Ikram Agent v2.0 is running! 🕋')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Webhook لاستقبال الرسائل
 */
function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
    
    if (contents.message || contents.callback_query) {
      return handleTelegramWebhook(contents);
    }
    
    return ContentService.createTextOutput('OK');
  } catch (error) {
    try {
      if (e.parameter && e.parameter.From && e.parameter.Body) {
        return handleWhatsAppWebhook(e);
      }
    } catch (err) {
      Logger.log('WhatsApp error: ' + err.toString());
    }
    
    Logger.log('doPost Error: ' + error.toString());
    return ContentService.createTextOutput('Error');
  }
}

/**
 * معالجة Telegram Webhook
 */
function handleTelegramWebhook(contents) {
  try {
    var message = contents.message || (contents.callback_query ? contents.callback_query.message : null);
    var callbackData = contents.callback_query ? contents.callback_query.data : null;
    
    if (!message) return ContentService.createTextOutput('OK');
    
    var chatId = message.chat.id;
    var userId = contents.callback_query ? contents.callback_query.from.id : message.from.id;
    var userName = contents.callback_query ? contents.callback_query.from.first_name : (message.from.first_name || 'User');
    var messageText = callbackData || message.text || '';
    
    if (!messageText) return ContentService.createTextOutput('OK');
    
    // معالجة الرسالة
    var response = processMessage('Telegram', String(userId), userName, messageText);
    
    // إرسال الرد
    if (response.buttons && response.buttons.length > 0) {
      sendTelegramButtons(chatId, response.text, response.buttons);
    } else {
      sendTelegram(chatId, response.text);
    }
    
    return ContentService.createTextOutput('OK');
  } catch (error) {
    Logger.log('handleTelegramWebhook Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return ContentService.createTextOutput('Error');
  }
}

/**
 * إرسال رسالة نصية
 */
function sendTelegram(chatId, text) {
  try {
    var settings = getAllAgentSettings();
    var token = settings.TELEGRAM_BOT_TOKEN;
    
    if (!token) {
      Logger.log('Telegram token not found');
      return false;
    }
    
    var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    
    var payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'Markdown'
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      Logger.log('Telegram send error: ' + response.getContentText());
      // إعادة المحاولة بدون Markdown
      payload.parse_mode = undefined;
      options.payload = JSON.stringify(payload);
      UrlFetchApp.fetch(url, options);
    }
    
    return true;
  } catch (e) {
    Logger.log('sendTelegram Error: ' + e.toString());
    return false;
  }
}

/**
 * إرسال رسالة مع أزرار
 */
function sendTelegramButtons(chatId, text, buttons) {
  try {
    var settings = getAllAgentSettings();
    var token = settings.TELEGRAM_BOT_TOKEN;
    
    if (!token) return false;
    
    var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
    
    // تحويل الأزرار (زرين في كل صف)
    var keyboard = [];
    for (var i = 0; i < buttons.length; i += 2) {
      var row = [{ text: buttons[i], callback_data: buttons[i].substring(0, 60) }];
      if (buttons[i + 1]) {
        row.push({ text: buttons[i + 1], callback_data: buttons[i + 1].substring(0, 60) });
      }
      keyboard.push(row);
    }
    
    var payload = {
      chat_id: chatId,
      text: text,
      parse_mode: 'Markdown',
      reply_markup: {
        inline_keyboard: keyboard
      }
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      Logger.log('Telegram buttons error: ' + response.getContentText());
      return sendTelegram(chatId, text);
    }
    
    return true;
  } catch (e) {
    Logger.log('sendTelegramButtons Error: ' + e.toString());
    return sendTelegram(chatId, text);
  }
}

/**
 * تعيين Webhook
 */
function setTelegramWebhook() {
  var settings = getAllAgentSettings();
  var token = settings.TELEGRAM_BOT_TOKEN;
  
  var url = 'https://api.telegram.org/bot' + token + '/setWebhook?url=' + encodeURIComponent(WEBAPP_URL);
  
  var response = UrlFetchApp.fetch(url);
  Logger.log('setTelegramWebhook: ' + response.getContentText());
  Logger.log('Webhook URL: ' + WEBAPP_URL);
  
  return response.getContentText();
}

/**
 * حذف Webhook
 */
function deleteTelegramWebhook() {
  var settings = getAllAgentSettings();
  var token = settings.TELEGRAM_BOT_TOKEN;
  
  var url = 'https://api.telegram.org/bot' + token + '/deleteWebhook';
  
  var response = UrlFetchApp.fetch(url);
  Logger.log('deleteTelegramWebhook: ' + response.getContentText());
  
  return response.getContentText();
}

/**
 * معلومات Webhook
 */
function getTelegramWebhookInfo() {
  var settings = getAllAgentSettings();
  var token = settings.TELEGRAM_BOT_TOKEN;
  
  var url = 'https://api.telegram.org/bot' + token + '/getWebhookInfo';
  
  var response = UrlFetchApp.fetch(url);
  Logger.log('Webhook Info: ' + response.getContentText());
  
  return response.getContentText();
}
function quickTest() {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var sheet = ss.getSheetByName('الطيران');
  var lastRow = sheet.getLastRow();
  Logger.log('Last row: ' + lastRow);
  
  // جلب عينة من البيانات
  var sample = sheet.getRange(3, 5, 10, 2).getValues();
  Logger.log('Sample data (Country, City):');
  for (var i = 0; i < sample.length; i++) {
    Logger.log((i+3) + ': ' + sample[i][0] + ' | ' + sample[i][1]);
  }
}