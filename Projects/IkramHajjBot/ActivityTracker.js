// ============================================
// تتبع نشاط الحجاج — BotActivity
// أعمدة: ChatID | Passport | Name | PackageId | Nationality | Language |
//        RegisteredAt | LastActivity | BotStatus | MessagesSent | MessagesDelivered |
//        PhoneCollected | RoomsCollected | PhotoCollected
// ============================================

// ============================================
// إنشاء/جلب شيت BotActivity
// ============================================
function getBotActivitySheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(BOT_ACTIVITY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BOT_ACTIVITY_SHEET);
    sheet.appendRow([
      'ChatID', 'Passport', 'Name', 'PackageId', 'Nationality', 'Language',
      'RegisteredAt', 'LastActivity', 'BotStatus', 'MessagesSent', 'MessagesDelivered',
      'PhoneCollected', 'RoomsCollected', 'PhotoCollected'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================
// تسجيل حاج جديد عند أول تسجيل ناجح
// ============================================
function registerPilgrim_(chatId, passport, pilgrimRow) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    // تحقق: هل مسجّل مسبقاً؟
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        // تحديث فقط
        var row = i + 1;
        sheet.getRange(row, 2).setValue(String(passport).toUpperCase());
        sheet.getRange(row, 3).setValue(pilgrimRow ? pilgrimRow.name : '');
        sheet.getRange(row, 5).setValue(pilgrimRow && pilgrimRow.rowData ? String(pilgrimRow.rowData[12] || '') : '');
        sheet.getRange(row, 8).setValue(new Date().toISOString());
        sheet.getRange(row, 9).setValue('active');
        return;
      }
    }

    // تسجيل جديد
    var nationality = pilgrimRow && pilgrimRow.rowData ? String(pilgrimRow.rowData[12] || '') : '';
    var packageId = pilgrimRow && pilgrimRow.rowData ? String(pilgrimRow.rowData[1] || '') : '';
    var name = pilgrimRow ? pilgrimRow.name : '';

    sheet.appendRow([
      chatStr,
      String(passport).toUpperCase(),
      name,
      packageId,
      nationality,
      '',  // Language — يُحدّث لاحقاً
      new Date().toISOString(),
      new Date().toISOString(),
      'active',
      0, 0,
      'no', 'no', 'no'
    ]);
  } catch (e) {
    Logger.log('registerPilgrim_ error: ' + e.message);
  }
}

// ============================================
// تحديث آخر نشاط عند كل تفاعل
// ============================================
function updateBotActivity_(chatId) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        sheet.getRange(i + 1, 8).setValue(new Date().toISOString()); // LastActivity
        if (String(data[i][8]) !== 'active') {
          sheet.getRange(i + 1, 9).setValue('active'); // BotStatus
        }
        return;
      }
    }
  } catch (e) {
    Logger.log('updateBotActivity_ error: ' + e.message);
  }
}

// ============================================
// تحديث لغة الحاج
// ============================================
function updateActivityLanguage_(chatId, lang) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        sheet.getRange(i + 1, 6).setValue(lang);
        return;
      }
    }
  } catch (e) {
    Logger.log('updateActivityLanguage_ error: ' + e.message);
  }
}

// ============================================
// تسجيل حظر / إلغاء حظر البوت
// ============================================
function markBlocked_(chatId) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        sheet.getRange(i + 1, 9).setValue('blocked');
        return;
      }
    }
  } catch (e) {
    Logger.log('markBlocked_ error: ' + e.message);
  }
}

function markActive_(chatId) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        if (String(data[i][8]) === 'blocked') {
          sheet.getRange(i + 1, 9).setValue('active');
        }
        return;
      }
    }
  } catch (e) {
    Logger.log('markActive_ error: ' + e.message);
  }
}

// ============================================
// تحديث عدادات الرسائل
// ============================================
function incrementMessagesSent_(chatId) {
  updateActivityCounter_(chatId, 10); // MessagesSent = column J (index 9, 1-based 10)
}

function incrementMessagesDelivered_(chatId) {
  updateActivityCounter_(chatId, 11); // MessagesDelivered = column K
}

function updateActivityCounter_(chatId, col1Based) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        var current = Number(data[i][col1Based - 1]) || 0;
        sheet.getRange(i + 1, col1Based).setValue(current + 1);
        return;
      }
    }
  } catch (e) {
    Logger.log('updateActivityCounter_ error: ' + e.message);
  }
}

// ============================================
// تحديث حالة جمع البيانات
// ============================================
function updateDataCollection_(chatId, field) {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var chatStr = String(chatId);

    var colMap = { phone: 12, rooms: 13, photo: 14 }; // 1-based columns
    var col = colMap[field];
    if (!col) return;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === chatStr) {
        sheet.getRange(i + 1, col).setValue('yes');
        return;
      }
    }
  } catch (e) {
    Logger.log('updateDataCollection_ error: ' + e.message);
  }
}

// ============================================
// مزامنة الحالات — Trigger يومي
// ============================================
function syncBotActivity_() {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var sevenDaysAgo = new Date(now.getTime() - 7 * 86400000);

    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][8]);
      if (status === 'blocked') continue;

      var lastActivity = data[i][7] ? new Date(data[i][7]) : null;
      if (lastActivity && lastActivity < sevenDaysAgo && status !== 'inactive') {
        sheet.getRange(i + 1, 9).setValue('inactive');
      } else if (lastActivity && lastActivity >= sevenDaysAgo && status === 'inactive') {
        sheet.getRange(i + 1, 9).setValue('active');
      }
    }
    Logger.log('syncBotActivity_ completed');
  } catch (e) {
    Logger.log('syncBotActivity_ error: ' + e.message);
  }
}

// ============================================
// إحصائيات سريعة
// ============================================
function getBotStats_() {
  try {
    var sheet = getBotActivitySheet_();
    var data = sheet.getDataRange().getValues();

    var stats = { total: 0, active: 0, blocked: 0, inactive: 0, phone: 0, rooms: 0, photo: 0 };

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      stats.total++;
      var status = String(data[i][8]);
      if (status === 'active') stats.active++;
      else if (status === 'blocked') stats.blocked++;
      else if (status === 'inactive') stats.inactive++;

      if (String(data[i][11]) === 'yes') stats.phone++;
      if (String(data[i][12]) === 'yes') stats.rooms++;
      if (String(data[i][13]) === 'yes') stats.photo++;
    }
    return stats;
  } catch (e) {
    Logger.log('getBotStats_ error: ' + e.message);
    return { total: 0, active: 0, blocked: 0, inactive: 0, phone: 0, rooms: 0, photo: 0 };
  }
}
