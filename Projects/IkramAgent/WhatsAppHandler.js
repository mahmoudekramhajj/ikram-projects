/**
 * معالج WhatsApp (Twilio) لوكيل إكرام
 * الإصدار 1.0
 */

/**
 * معالجة Webhook من WhatsApp/Twilio
 */
function handleWhatsAppWebhook(e) {
  try {
    var params = e.parameter;
    
    var from = params.From || '';
    var body = params.Body || '';
    var profileName = params.ProfileName || 'User';
    
    // تنظيف رقم الهاتف
    var userId = from.replace('whatsapp:', '');
    
    if (!body) return ContentService.createTextOutput('<Response></Response>')
      .setMimeType(ContentService.MimeType.XML);
    
    // معالجة الرسالة
    var response = processMessage('WhatsApp', userId, profileName, body);
    
    // إرسال الرد
    sendWhatsApp(userId, response.text);
    
    // إذا كان طلب تسجيل بيانات
    if (response.requestLead) {
      // سيتم التعامل مع جمع البيانات في الرسائل القادمة
    }
    
    return ContentService.createTextOutput('<Response></Response>')
      .setMimeType(ContentService.MimeType.XML);
      
  } catch (error) {
    Logger.log('handleWhatsAppWebhook Error: ' + error.toString());
    return ContentService.createTextOutput('<Response></Response>')
      .setMimeType(ContentService.MimeType.XML);
  }
}

/**
 * إرسال رسالة WhatsApp عبر Twilio
 */
function sendWhatsApp(to, message) {
  try {
    var settings = getAllAgentSettings();
    
    var accountSid = settings.TWILIO_ACCOUNT_SID;
    var authToken = settings.TWILIO_AUTH_TOKEN;
    var fromNumber = settings.TWILIO_WHATSAPP_NUMBER;
    
    if (!accountSid || !authToken || !fromNumber) {
      Logger.log('Twilio settings not found');
      return false;
    }
    
    // تنسيق الأرقام
    var toFormatted = to.startsWith('+') ? to : '+' + to;
    var fromFormatted = fromNumber.startsWith('+') ? fromNumber : '+' + fromNumber;
    
    var url = 'https://api.twilio.com/2010-04-01/Accounts/' + accountSid + '/Messages.json';
    
    var payload = {
      'To': 'whatsapp:' + toFormatted,
      'From': 'whatsapp:' + fromFormatted,
      'Body': message
    };
    
    var options = {
      method: 'post',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(accountSid + ':' + authToken)
      },
      payload: payload,
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      Logger.log('WhatsApp sent successfully to: ' + toFormatted);
      return true;
    } else {
      Logger.log('WhatsApp send error: ' + responseCode + ' - ' + responseText);
      return false;
    }
    
  } catch (e) {
    Logger.log('sendWhatsApp Error: ' + e.toString());
    return false;
  }
}

/**
 * إرسال إشعار WhatsApp للفريق
 */
function sendWhatsAppNotification(message) {
  try {
    var settings = getAllAgentSettings();
    var notificationPhone = settings.NOTIFICATION_PHONE;
    
    if (!notificationPhone) {
      Logger.log('Notification phone not found');
      return false;
    }
    
    return sendWhatsApp(notificationPhone, message);
    
  } catch (e) {
    Logger.log('sendWhatsAppNotification Error: ' + e.toString());
    return false;
  }
}

/**
 * تعيين Webhook لـ Twilio WhatsApp
 * يجب تنفيذها بعد نشر Web App والحصول على الرابط
 */
function setTwilioWebhook() {
  var webAppUrl = ScriptApp.getService().getUrl();
  
  Logger.log('===========================================');
  Logger.log('Twilio WhatsApp Webhook Setup');
  Logger.log('===========================================');
  Logger.log('1. Go to: https://console.twilio.com');
  Logger.log('2. Navigate to: Messaging > Try it out > Send a WhatsApp message');
  Logger.log('3. Click on "Sandbox settings"');
  Logger.log('4. In "WHEN A MESSAGE COMES IN" field, paste this URL:');
  Logger.log(webAppUrl);
  Logger.log('5. Set method to: HTTP POST');
  Logger.log('6. Save');
  Logger.log('===========================================');
  
  return webAppUrl;
}

/**
 * اختبار إرسال WhatsApp
 */
function testWhatsAppSend() {
  var settings = getAllAgentSettings();
  var testPhone = settings.NOTIFICATION_PHONE;
  
  var testMessage = '🧪 *اختبار وكيل إكرام*\n\n' +
    'إذا وصلتك هذه الرسالة، فإن إعدادات WhatsApp تعمل بشكل صحيح!\n\n' +
    '✅ Twilio متصل\n' +
    '✅ الإعدادات صحيحة';
  
  var result = sendWhatsApp(testPhone, testMessage);
  
  if (result) {
    Logger.log('✅ Test message sent successfully!');
  } else {
    Logger.log('❌ Failed to send test message');
  }
  
  return result;
}

/**
 * اختبار استلام من Sandbox
 * ملاحظة: يجب أولاً إرسال "join return-you" من هاتفك للرقم +1 415 523 8886
 */
function sandboxInstructions() {
  Logger.log('===========================================');
  Logger.log('Twilio WhatsApp Sandbox Setup');
  Logger.log('===========================================');
  Logger.log('لتفعيل Sandbox على هاتفك:');
  Logger.log('1. افتح WhatsApp على هاتفك');
  Logger.log('2. أرسل رسالة إلى: +1 415 523 8886');
  Logger.log('3. اكتب: join return-you');
  Logger.log('4. انتظر رسالة التأكيد');
  Logger.log('===========================================');
}