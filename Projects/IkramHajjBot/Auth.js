// ============================================
// المصادقة — اللغة + الجواز + إدارة الجلسات
// ============================================

// ============================================
// تبديل الحاج — تسجيل خروج + إعادة طلب الجواز
// ============================================
function handleSwitchPilgrim_(chatId, session) {
  var lang = (session && session.language) || 'ar';

  // مسح كاش الحاج الحالي
  if (session.passport) {
    clearPilgrimCache_(session.passport);
  }

  // إعادة الجلسة لمرحلة إدخال الجواز مع الحفاظ على اللغة
  updateSession_(chatId, {
    passport: '',
    authStatus: 'awaiting_passport',
    authTimestamp: ''
  });

  sendMessage_(chatId, T_('switch_pilgrim', lang));
}

function handleLanguageSelection_(chatId, lang) {
  saveSession_(chatId, { language: lang, authStatus: 'awaiting_passport' });
  sendMessage_(chatId, T_('passport_prompt', lang));
}

function handlePassportInput_(chatId, text, session) {
  var lang = session.language || 'ar';

  if (text.length < 5 || text.length > 15) {
    sendMessage_(chatId, T_('passport_invalid', lang));
    return;
  }

  var pilgrim = findPilgrimByPassport_(text);

  if (pilgrim) {
    updateSession_(chatId, {
      passport: text.toUpperCase(),
      authStatus: 'verified',
      authTimestamp: new Date().toISOString()
    });

    var successMsgs = {
      ar: '✅ تم التحقق بنجاح!\n\nمرحباً <b>' + pilgrim.name + '</b> 🕋',
      en: '✅ Verified successfully!\n\nWelcome <b>' + pilgrim.name + '</b> 🕋',
      fr: '✅ Vérification réussie!\n\nBienvenue <b>' + pilgrim.name + '</b> 🕋',
      nl: '✅ Verificatie geslaagd!\n\nWelkom <b>' + pilgrim.name + '</b> 🕋'
    };

    sendMessage_(chatId, successMsgs[lang] || successMsgs['en']);
    sendMainMenu_(chatId, lang);

  } else {
    sendMessage_(chatId, T_('passport_not_found', lang));
  }
}

// ============================================
// إدارة الجلسات — BotSessions
// ============================================
function getSessionSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName('BotSessions');
}

function getSession_(chatId) {
  var cacheKey = 'session_' + String(chatId);

  var cached = getCache_(cacheKey);
  if (cached) return cached;

  var sheet = getSessionSheet_();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(chatId)) {
      var session = {
        row: i + 1,
        chatId: String(data[i][0]),
        passport: String(data[i][1]),
        inputState: String(data[i][2] || ''),
        language: String(data[i][3]),
        authStatus: String(data[i][4]),
        authTimestamp: String(data[i][5]),
        lastActivity: String(data[i][6]),
        messageCount: Number(data[i][7]) || 0
      };
      setCache_(cacheKey, session);
      return session;
    }
  }
  return null;
}

function saveSession_(chatId, fields) {
  clearSessionCache_(chatId);

  var existing = getSession_(chatId);
  if (existing) {
    updateSession_(chatId, fields);
  } else {
    var sheet = getSessionSheet_();
    sheet.appendRow([
      String(chatId),
      fields.passport || '',
      '',
      fields.language || 'ar',
      fields.authStatus || 'none',
      '',
      new Date().toISOString(),
      0
    ]);
    clearSessionCache_(chatId);
  }
}

function updateSession_(chatId, fields) {
  var sheet = getSessionSheet_();
  var session = getSession_(chatId);
  if (!session) return;

  var row = session.row;
  if (fields.passport !== undefined) sheet.getRange(row, 2).setValue(fields.passport);
  if (fields.inputState !== undefined) sheet.getRange(row, 3).setValue(fields.inputState);
  if (fields.language !== undefined) sheet.getRange(row, 4).setValue(fields.language);
  if (fields.authStatus !== undefined) sheet.getRange(row, 5).setValue(fields.authStatus);
  if (fields.authTimestamp !== undefined) sheet.getRange(row, 6).setValue(fields.authTimestamp);

  sheet.getRange(row, 7).setValue(new Date().toISOString());
  sheet.getRange(row, 8).setValue(session.messageCount + 1);

  clearSessionCache_(chatId);
}

function updateLastActivity_(chatId) {
  updateSession_(chatId, {});
}
