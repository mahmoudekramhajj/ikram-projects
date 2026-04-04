// ============================================
// محرك البث — إرسال رسائل الإدارة للحجاج
// ============================================

// ============================================
// إنشاء/جلب الشيتات
// ============================================
function getAdminMessagesSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(ADMIN_MESSAGES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ADMIN_MESSAGES_SHEET);
    sheet.appendRow([
      'ID', 'Title', 'MessageAR', 'MessageEN', 'MessageFR', 'MessageDE', 'MessageIT', 'MessageES',
      'ImageURL', 'FileURL', 'FileName', 'Target', 'TargetValue',
      'Priority', 'Status', 'SentAt', 'SentCount', 'FailCount', 'CreatedBy'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDeliveryLogSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(DELIVERY_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DELIVERY_LOG_SHEET);
    sheet.appendRow(['MessageID', 'ChatID', 'Passport', 'Name', 'Language', 'Status', 'ErrorMsg', 'SentAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================
// معالجة الرسائل المعلّقة — Trigger كل 5 دقائق
// ============================================
function processAdminMessages_() {
  try {
    var sheet = getAdminMessagesSheet_();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][AM.STATUS]).toLowerCase().trim();
      if (status !== 'pending') continue;

      // حماية من تكرار: نغيّر الحالة فوراً
      sheet.getRange(i + 1, AM.STATUS + 1).setValue('sending');
      SpreadsheetApp.flush();

      var msgId = String(data[i][AM.ID] || (i));
      var msgRow = data[i];

      // جلب الجلسات المستهدفة
      var target = String(msgRow[AM.TARGET] || 'all').toLowerCase().trim();
      var targetValue = String(msgRow[AM.TARGET_VALUE] || '').trim();
      var sessions = getTargetSessions_(target, targetValue);

      // إرسال
      var result = sendBroadcast_(sessions, msgRow, msgId);

      // تحديث الشيت
      sheet.getRange(i + 1, AM.STATUS + 1).setValue('sent');
      sheet.getRange(i + 1, AM.SENT_AT + 1).setValue(new Date().toISOString());
      sheet.getRange(i + 1, AM.SENT_COUNT + 1).setValue(result.sent);
      sheet.getRange(i + 1, AM.FAIL_COUNT + 1).setValue(result.failed);

      // إرسال ملخص للمدراء
      var title = String(msgRow[AM.TITLE] || 'رسالة');
      sendBroadcastStats_(title, result.sent, result.failed);
    }
  } catch (e) {
    Logger.log('processAdminMessages_ error: ' + e.message);
  }
}

// ============================================
// جلب الجلسات المستهدفة
// ============================================
function getTargetSessions_(target, targetValue) {
  var sessionSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BotSessions');
  var sessionData = sessionSheet.getDataRange().getValues();

  var verified = [];
  for (var i = 1; i < sessionData.length; i++) {
    if (String(sessionData[i][4]) === 'verified' && sessionData[i][1]) {
      verified.push({
        chatId: String(sessionData[i][0]),
        passport: String(sessionData[i][1]),
        language: String(sessionData[i][3]) || 'en'
      });
    }
  }

  if (target === 'all') return verified;

  // نحتاج بيانات رحلة الحاج للاستهداف
  var journeySheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(JOURNEY_SHEET);
  var journeyData = journeySheet.getDataRange().getValues();

  // بناء خريطة: passport → rowData
  var passportMap = {};
  for (var j = 1; j < journeyData.length; j++) {
    var pp = String(journeyData[j][8] || '').toUpperCase().trim();
    if (pp) passportMap[pp] = journeyData[j];
  }

  var filtered = [];
  var tv = targetValue.toUpperCase().trim();

  for (var k = 0; k < verified.length; k++) {
    var s = verified[k];
    var row = passportMap[s.passport.toUpperCase().trim()];
    if (!row) continue;

    var match = false;
    if (target === 'package') {
      match = String(row[1] || '').toUpperCase().trim() === tv; // PackageId
    } else if (target === 'nationality') {
      match = String(row[12] || '').toUpperCase().trim().indexOf(tv) !== -1; // NationalityEn
    } else if (target === 'flight') {
      match = String(row[24] || '').toUpperCase().trim().indexOf(tv) !== -1; // ArrivalFlightNumber
    }

    if (match) filtered.push(s);
  }

  return filtered;
}

// ============================================
// إرسال البث لقائمة الجلسات
// ============================================
function sendBroadcast_(sessions, msgRow, msgId) {
  var logSheet = getDeliveryLogSheet_();
  var logBatch = [];
  var sent = 0;
  var failed = 0;

  var imageUrl = String(msgRow[AM.IMAGE_URL] || '').trim();
  var fileUrl = String(msgRow[AM.FILE_URL] || '').trim();
  var fileName = String(msgRow[AM.FILE_NAME] || '').trim();
  var priority = String(msgRow[AM.PRIORITY] || '').toLowerCase().trim();

  // تحويل روابط Drive
  if (imageUrl) imageUrl = getDriveDirectUrl_(imageUrl);
  if (fileUrl) fileUrl = getDriveDirectUrl_(fileUrl);

  for (var i = 0; i < sessions.length; i++) {
    var s = sessions[i];
    var lang = s.language || 'en';

    // جلب النص بلغة الحاج
    var langCol = LANG_TO_COL[lang] || AM.MSG_EN;
    var text = String(msgRow[langCol] || '').trim();
    if (!text) text = String(msgRow[AM.MSG_EN] || msgRow[AM.MSG_AR] || '').trim();

    // بناء الرسالة
    var header = priority === 'urgent' ? T_('ann_urgent', lang) + ' ' : '';
    var fullText = T_('ann_from_admin', lang) + header + text;

    var result;

    // إرسال نص
    result = sendMessageSafe_(s.chatId, fullText);

    if (result.success) {
      sent++;
      incrementMessagesSent_(s.chatId);
      incrementMessagesDelivered_(s.chatId);
      markActive_(s.chatId);

      // إرسال صورة إن وُجدت
      if (imageUrl) {
        sendPhotoSafe_(s.chatId, imageUrl, '');
      }

      // إرسال ملف إن وُجد
      if (fileUrl) {
        var caption = fileName || T_('ann_file_attached', lang);
        sendDocumentSafe_(s.chatId, fileUrl, fileName, caption);
      }

      logBatch.push([msgId, s.chatId, s.passport, '', lang, 'delivered', '', new Date().toISOString()]);
    } else {
      failed++;
      incrementMessagesSent_(s.chatId);

      // كشف حظر البوت
      if (result.error && result.error.indexOf('blocked') !== -1 || result.error && result.error.indexOf('403') !== -1) {
        markBlocked_(s.chatId);
      }

      logBatch.push([msgId, s.chatId, s.passport, '', lang, 'failed', result.error, new Date().toISOString()]);
    }

    // Rate limit protection
    Utilities.sleep(50);
  }

  // كتابة السجل دفعة واحدة
  if (logBatch.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logBatch.length, 8).setValues(logBatch);
  }

  return { sent: sent, failed: failed };
}

// ============================================
// إرسال ملخص الإحصائيات للمدراء
// ============================================
function sendBroadcastStats_(title, sentCount, failCount) {
  var msg = '📊 <b>تم إرسال رسالة</b>\n━━━━━━━━━━━━━━\n' +
    '📝 ' + title + '\n' +
    '✅ وصل: <b>' + sentCount + '</b>\n' +
    '❌ فشل: <b>' + failCount + '</b>';

  for (var i = 0; i < ADMIN_IDS.length; i++) {
    sendMessageSafe_(ADMIN_IDS[i], msg);
  }
}

// ============================================
// إضافة رسالة في AdminMessages من البوت
// ============================================
function addAdminMessage_(msgData) {
  var sheet = getAdminMessagesSheet_();
  var data = sheet.getDataRange().getValues();
  var nextId = data.length; // auto-increment

  sheet.appendRow([
    nextId,
    msgData.title || '',
    msgData.ar || '',
    msgData.en || '',
    msgData.fr || '',
    msgData.de || '',
    msgData.it || '',
    msgData.es || '',
    msgData.imageUrl || '',
    msgData.fileUrl || '',
    msgData.fileName || '',
    msgData.target || 'all',
    msgData.targetValue || '',
    msgData.priority || 'normal',
    'pending',
    '',
    0,
    0,
    'bot'
  ]);

  return nextId;
}

// ============================================
// جلب آخر الرسائل المرسلة (للحاج)
// ============================================
function getRecentMessages_(passport, maxCount) {
  maxCount = maxCount || 5;
  try {
    var sheet = getAdminMessagesSheet_();
    var data = sheet.getDataRange().getValues();

    // نحتاج بيانات الحاج للاستهداف
    var pilgrim = findPilgrimByPassport_(passport);
    var packageId = pilgrim && pilgrim.rowData ? String(pilgrim.rowData[1] || '') : '';
    var nationality = pilgrim && pilgrim.rowData ? String(pilgrim.rowData[12] || '') : '';
    var flight = pilgrim && pilgrim.rowData ? String(pilgrim.rowData[24] || '') : '';

    var messages = [];
    for (var i = data.length - 1; i >= 1 && messages.length < maxCount; i--) {
      if (String(data[i][AM.STATUS]).toLowerCase().trim() !== 'sent') continue;

      var target = String(data[i][AM.TARGET] || 'all').toLowerCase().trim();
      var tv = String(data[i][AM.TARGET_VALUE] || '').toUpperCase().trim();

      var match = false;
      if (target === 'all') {
        match = true;
      } else if (target === 'package') {
        match = packageId.toUpperCase().trim() === tv;
      } else if (target === 'nationality') {
        match = nationality.toUpperCase().trim().indexOf(tv) !== -1;
      } else if (target === 'flight') {
        match = flight.toUpperCase().trim().indexOf(tv) !== -1;
      }

      if (match) {
        messages.push({
          id: String(data[i][AM.ID] || i),
          title: String(data[i][AM.TITLE] || ''),
          priority: String(data[i][AM.PRIORITY] || ''),
          sentAt: String(data[i][AM.SENT_AT] || ''),
          rowData: data[i]
        });
      }
    }
    return messages;
  } catch (e) {
    Logger.log('getRecentMessages_ error: ' + e.message);
    return [];
  }
}

// ============================================
// جلب رسالة بالـ ID
// ============================================
function getMessageById_(msgId) {
  try {
    var sheet = getAdminMessagesSheet_();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][AM.ID]) === String(msgId)) {
        return data[i];
      }
    }
    return null;
  } catch (e) {
    Logger.log('getMessageById_ error: ' + e.message);
    return null;
  }
}

// ============================================
// إعداد Trigger تلقائي
// ============================================
function setupBroadcastTrigger() {
  // حذف triggers قديمة
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processAdminMessages_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // إنشاء trigger جديد — كل 5 دقائق
  ScriptApp.newTrigger('processAdminMessages_')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('Broadcast trigger set: every 5 minutes');
}
