// ============================================
// إعدادات البوت
// ============================================
var BOT_TOKEN = '8694589281:AAHvT-anZgLDk6s5YO8WStP7Y2zhq6-BDIE';
var TELEGRAM_API = 'https://api.telegram.org/bot' + BOT_TOKEN;
var SHEET_ID = '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s';
var JOURNEY_SHEET = 'رحلة الحاج '; // مسافة في النهاية — مهم

// ============================================
// نقطة الدخول
// ============================================
function doPost(e) {
  try {
    var update = JSON.parse(e.postData.contents);
    
    if (update.message) {
      handleMessage_(update.message);
      return;
    }
    
    if (update.callback_query) {
      handleCallback_(update.callback_query);
      return;
    }
    
  } catch (err) {
    Logger.log('Bot Error: ' + err.message + '\n' + err.stack);
  }
}

// ============================================
// معالجة الرسائل النصية
// ============================================
function handleMessage_(message) {
  var chatId = message.chat.id;
  var text = (message.text || '').trim();
  var firstName = message.from.first_name || '';
  
  if (text === '/start') {
    sendWelcome_(chatId, firstName);
    return;
  }
  
  if (text === '/menu') {
    var session = getSession_(chatId);
    if (session && session.authStatus === 'verified') {
      sendMainMenu_(chatId, session.language);
    } else {
      sendMessage_(chatId, 'أرسل /start للبدء');
    }
    return;
  }
  
  var session = getSession_(chatId);
  
  if (!session) {
    sendMessage_(chatId, 'أرسل /start للبدء');
    return;
  }
  
  if (session.authStatus === 'awaiting_passport') {
    handlePassportInput_(chatId, text, session);
    return;
  }
  
  if (session.authStatus === 'verified') {
    sendMessage_(chatId, '📩 استلمنا سؤالك:\n«' + text + '»\n\nاستخدم القائمة للاستعلام عن بياناتك 👇');
    sendMainMenu_(chatId, session.language);
    updateLastActivity_(chatId);
    return;
  }
  
  sendMessage_(chatId, 'أرسل /start للبدء من جديد');
}

// ============================================
// معالجة ضغطات الأزرار
// ============================================
function handleCallback_(callback) {
  var chatId = callback.message.chat.id;
  var data = callback.data;
  
  answerCallback_(callback.id);
  
  if (data.indexOf('lang_') === 0) {
    var lang = data.replace('lang_', '');
    handleLanguageSelection_(chatId, lang);
    return;
  }
  
  var session = getSession_(chatId);
  if (!session || session.authStatus !== 'verified') {
    sendMessage_(chatId, 'أرسل /start للبدء');
    return;
  }
  
  if (data === 'my_flight') {
    handleMyFlight_(chatId, session);
    return;
  }
  
  if (data === 'my_hotel') {
    handleMyHotel_(chatId, session);
    return;
  }
  
  if (data === 'my_package') {
    handleMyPackage_(chatId, session);
    return;
  }
  
  if (data === 'my_transport') {
    handleMyTransport_(chatId, session);
    return;
  }
  
  if (data === 'emergency') {
    handleEmergency_(chatId, session);
    return;
  }
  
  if (data === 'contact_company') {
    handleContactCompany_(chatId, session);
    return;
  }

  if (data === 'general_faq') {
    sendMessage_(chatId, '⚙️ قسم الأسئلة العامة قيد التطوير...');
    return;
  }
  
  if (data === 'change_lang') {
    sendLanguageButtons_(chatId);
    return;
  }
  
  if (data === 'show_menu') {
    sendMainMenu_(chatId, session.language);
    return;
  }
}

// ============================================
// اختيار اللغة → طلب رقم الجواز
// ============================================
function handleLanguageSelection_(chatId, lang) {
  saveSession_(chatId, { language: lang, authStatus: 'awaiting_passport' });
  
  var msgs = {
    ar: '🔐 للوصول لبياناتك الشخصية:\n\nأرسل <b>رقم جواز السفر</b> الخاص بك',
    en: '🔐 To access your personal data:\n\nSend your <b>passport number</b>',
    fr: '🔐 Pour accéder à vos données:\n\nEnvoyez votre <b>numéro de passeport</b>',
    nl: '🔐 Om uw gegevens te bekijken:\n\nStuur uw <b>paspoortnummer</b>'
  };
  
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// ============================================
// التحقق من رقم الجواز
// ============================================
function handlePassportInput_(chatId, text, session) {
  var lang = session.language || 'ar';
  
  if (text.length < 5 || text.length > 15) {
    var errMsgs = {
      ar: '⚠️ رقم الجواز غير صحيح. تأكد وأعد الإرسال:',
      en: '⚠️ Invalid passport number. Please try again:',
      fr: '⚠️ Numéro de passeport invalide. Réessayez:',
      nl: '⚠️ Ongeldig paspoortnummer. Probeer opnieuw:'
    };
    sendMessage_(chatId, errMsgs[lang] || errMsgs['en']);
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
    var failMsgs = {
      ar: '❌ رقم الجواز غير موجود في النظام.\n\nتأكد من الرقم وأعد الإرسال:',
      en: '❌ Passport not found in our system.\n\nPlease check and try again:',
      fr: '❌ Passeport non trouvé.\n\nVérifiez et réessayez:',
      nl: '❌ Paspoort niet gevonden.\n\nControleer en probeer opnieuw:'
    };
    sendMessage_(chatId, failMsgs[lang] || failMsgs['en']);
  }
}

// ============================================
// البحث عن الحاج برقم الجواز
// ============================================
function findPilgrimByPassport_(passportNo) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(JOURNEY_SHEET);
  if (!sheet) {
    Logger.log('Sheet not found: ' + JOURNEY_SHEET);
    return null;
  }
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  
  var headers = data[1];
  var colPassport = 8;
  var colName = 7;
  
  var inputPassport = passportNo.toUpperCase().trim();
  
  for (var i = 2; i < data.length; i++) {
    var rowPassport = String(data[i][colPassport]).toUpperCase().trim();
    
    if (rowPassport === inputPassport) {
      var name = colName !== -1 ? String(data[i][colName]).trim() : 'حاج';
      return { name: name, row: i, rowData: data[i] };
    }
  }
  
  return null;
}

// ============================================
// جلب نوع التنقل من شيت الباقات
// الربط: عمود B في الباقات = عمود B في رحلة الحاج (PackageId)
// المخرج: "حافلة" / "قطار" أو "-"
// ============================================
function getTransportType_(packageId) {
  if (!packageId || packageId === '-') return '-';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('الباقات');
  if (!sheet) {
    Logger.log('Sheet not found: الباقات');
    return '-';
  }
  
  var data = sheet.getDataRange().getValues();
  var colB = 1;          // عمود B = index 1
  var colTransport = 65; // عمود BN = index 65 (التنقل)
  
  var inputId = String(packageId).trim();
  for (var i = 2; i < data.length; i++) {
    if (String(data[i][colB]).trim() === inputId) {
      return String(data[i][colTransport]).trim() || '-';
    }
  }
  return '-';
}

// ============================================
// 🏨 فندقي — بيانات الفنادق
// ============================================
function handleMyHotel_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);
  
  if (!pilgrim) {
    sendMessage_(chatId, lang === 'ar' ? '❌ لم يتم العثور على بياناتك' : '❌ Data not found');
    return;
  }
  
  var r = pilgrim.rowData;
  
  var house1 = String(r[36] || '-');
  var house1Start = formatDate_(r[37]);
  var house1End = formatDate_(r[38]);
  var house1Hotel = '';
  if (house1.toLowerCase().indexOf('madi') !== -1) {
    house1Hotel = lang === 'ar' ? String(r[46] || '-') : String(r[47] || '-');
  } else {
    house1Hotel = lang === 'ar' ? String(r[42] || '-') : String(r[43] || '-');
  }
  
  var house2 = String(r[39] || '-');
  var house2Start = formatDate_(r[40]);
  var house2End = formatDate_(r[41]);
  var house2Hotel = '';
  if (house2.toLowerCase().indexOf('madi') !== -1) {
    house2Hotel = lang === 'ar' ? String(r[46] || '-') : String(r[47] || '-');
  } else if (house2.toLowerCase().indexOf('shift') !== -1) {
    house2Hotel = lang === 'ar' ? String(r[44] || '-') : String(r[45] || '-');
  } else {
    house2Hotel = lang === 'ar' ? String(r[42] || '-') : String(r[43] || '-');
  }
  
  var text = '';
  
  if (lang === 'ar') {
    text = '🏨 <b>السكن الأول — ' + house1 + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 الفندق: <b>' + house1Hotel + '</b>\n' +
      '📅 الدخول: ' + house1Start + '\n' +
      '📅 الخروج: ' + house1End + '\n\n' +
      '🏨 <b>السكن الثاني — ' + house2 + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 الفندق: <b>' + house2Hotel + '</b>\n' +
      '📅 الدخول: ' + house2Start + '\n' +
      '📅 الخروج: ' + house2End;
  } else {
    text = '🏨 <b>Accommodation 1 — ' + house1 + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 Hotel: <b>' + house1Hotel + '</b>\n' +
      '📅 Check-in: ' + house1Start + '\n' +
      '📅 Check-out: ' + house1End + '\n\n' +
      '🏨 <b>Accommodation 2 — ' + house2 + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 Hotel: <b>' + house2Hotel + '</b>\n' +
      '📅 Check-in: ' + house2Start + '\n' +
      '📅 Check-out: ' + house2End;
  }
  
  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: lang === 'ar' ? '🔙 القائمة الرئيسية' : '🔙 Main Menu', callback_data: 'show_menu' }]]
  });
}

// ============================================
// 📦 باقتي + ⛺ المخيم
// ============================================
function handleMyPackage_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);
  
  if (!pilgrim) {
    sendMessage_(chatId, lang === 'ar' ? '❌ لم يتم العثور على بياناتك' : '❌ Data not found');
    return;
  }
  
  var r = pilgrim.rowData;
  
  var packageId = String(r[1] || '-');
  var campName = String(r[4] || '-');
  var groupNo = String(r[6] || '-');
  var nationality = lang === 'ar' ? String(r[13] || '-') : String(r[12] || '-');
  var country = lang === 'ar' ? String(r[15] || '-') : String(r[14] || '-');
  var flightType = String(r[25] || '-');
  
  var text = '';
  
  if (lang === 'ar') {
    text = '📦 <b>بيانات الباقة</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🔢 رقم الباقة: <b>' + packageId + '</b>\n' +
      '👥 رقم المجموعة: ' + groupNo + '\n' +
      '🎫 نوع التذكرة: ' + flightType + '\n' +
      '🌍 الجنسية: ' + nationality + '\n' +
      '🏠 بلد الإقامة: ' + country + '\n\n' +
      '⛺ <b>المخيم</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '📍 الموقع: <b>' + campName + '</b>';
  } else {
    text = '📦 <b>Package Details</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🔢 Package ID: <b>' + packageId + '</b>\n' +
      '👥 Group: ' + groupNo + '\n' +
      '🎫 Ticket Type: ' + flightType + '\n' +
      '🌍 Nationality: ' + nationality + '\n' +
      '🏠 Residence: ' + country + '\n\n' +
      '⛺ <b>Mina Camp</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '📍 Location: <b>' + campName + '</b>';
  }
  
  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: lang === 'ar' ? '🔙 القائمة الرئيسية' : '🔙 Main Menu', callback_data: 'show_menu' }]]
  });
}

// ============================================
// 🚌 مواصلاتي — مع نوع التنقل من شيت الباقات
// ============================================
function handleMyTransport_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);
  
  if (!pilgrim) {
    sendMessage_(chatId, lang === 'ar' ? '❌ لم يتم العثور على بياناتك' : '❌ Data not found');
    return;
  }
  
  var r = pilgrim.rowData;
  
  // بيانات النقل من رحلة الوصول والعودة
  var arrFrom = String(r[21] || '-');
  var arrTo = String(r[19] || '-');
  var arrDate = formatDate_(r[20]);
  var arrFlight = String(r[24] || '-');
  
  var retFrom = String(r[31] || '-');
  var retTo = String(r[29] || '-');
  var retDate = formatDate_(r[32]);
  var retFlight = String(r[34] || '-');
  
  // الفنادق
  var house1 = String(r[36] || '-');
  var house2 = String(r[39] || '-');
  var house1Start = formatDate_(r[37]);
  var house2Start = formatDate_(r[40]);
  
  // المخيم
  var campName = String(r[4] || '-');
  
  // نوع التنقل من شيت الباقات (حافلة / قطار)
  var packageId = String(r[1] || '');
  var transportType = getTransportType_(packageId);
  
  // الأيقونة حسب نوع التنقل
  var transportIcon = (transportType === 'قطار') ? '🚆' : '🚌';
  
  // ترجمة نوع التنقل للإنجليزية
  var transportTypeEn = '-';
  if (transportType === 'قطار') transportTypeEn = 'Train';
  else if (transportType === 'حافلة') transportTypeEn = 'Bus';
  else transportTypeEn = transportType;
  
  var text = '';
  
  if (lang === 'ar') {
    text = '🚌 <b>خطة التنقل</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '1️⃣ <b>الوصول — استقبال المطار</b>\n' +
      '✈️ ' + arrFlight + ' → ' + arrTo + '\n' +
      '📅 ' + arrDate + '\n' +
      '🚌 نقل من المطار إلى فندق ' + house1 + '\n\n' +
      '2️⃣ <b>التنقل بين المدن</b>\n' +
      transportIcon + ' من ' + house1 + ' إلى ' + house2 + ' (<b>' + transportType + '</b>)\n' +
      '📅 ' + house2Start + '\n\n' +
      '3️⃣ <b>⛺ المخيم — منى</b>\n' +
      '📍 <b>' + campName + '</b>\n\n' +
      '4️⃣ <b>المغادرة — توديع المطار</b>\n' +
      '✈️ ' + retFlight + ' → ' + retTo + '\n' +
      '📅 ' + retDate + '\n' +
      '🚌 نقل من الفندق إلى مطار ' + retFrom;
  } else {
    text = '🚌 <b>Transport Plan</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '1️⃣ <b>Arrival — Airport Transfer</b>\n' +
      '✈️ ' + arrFlight + ' → ' + arrTo + '\n' +
      '📅 ' + arrDate + '\n' +
      '🚌 Transfer to ' + house1 + ' hotel\n\n' +
      '2️⃣ <b>Intercity Transfer</b>\n' +
      transportIcon + ' From ' + house1 + ' to ' + house2 + ' (<b>' + transportTypeEn + '</b>)\n' +
      '📅 ' + house2Start + '\n\n' +
      '3️⃣ <b>⛺ Mina Camp</b>\n' +
      '📍 <b>' + campName + '</b>\n\n' +
      '4️⃣ <b>Departure — Airport Transfer</b>\n' +
      '✈️ ' + retFlight + ' → ' + retTo + '\n' +
      '📅 ' + retDate + '\n' +
      '🚌 Transfer to ' + retFrom + ' airport';
  }
  
  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: lang === 'ar' ? '🔙 القائمة الرئيسية' : '🔙 Main Menu', callback_data: 'show_menu' }]]
  });
}

// ============================================
// ✈️ رحلتي — بيانات الطيران
// ============================================
function handleMyFlight_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);
  
  if (!pilgrim) {
    sendMessage_(chatId, lang === 'ar' ? '❌ لم يتم العثور على بياناتك' : '❌ Data not found');
    return;
  }
  
  var r = pilgrim.rowData;
  
  var arrFlight = String(r[24] || '-');
  var arrAirline = lang === 'ar' ? String(r[16] || '-') : String(r[17] || '-');
  var arrFrom = String(r[21] || '-');
  var arrTo = String(r[19] || '-');
  var arrDate = formatDate_(r[20]);
  var arrDepTime = formatTime_(r[23]);
  var arrArvTime = formatTime_(r[18]);
  var arrType = String(r[25] || '-');
  
  var retFlight = String(r[34] || '-');
  var retAirline = lang === 'ar' ? String(r[26] || '-') : String(r[27] || '-');
  var retFrom = String(r[31] || '-');
  var retTo = String(r[29] || '-');
  var retDate = formatDate_(r[32]);
  var retDepTime = formatTime_(r[33]);
  var retArvTime = formatTime_(r[28]);
  
  var text = '';
  
  if (lang === 'ar') {
    text = '✈️ <b>رحلة الوصول</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🛫 الرحلة: <b>' + arrFlight + '</b>\n' +
      '🏢 الناقل: ' + arrAirline + '\n' +
      '📅 التاريخ: ' + arrDate + '\n' +
      '🛫 من: ' + arrFrom + ' (' + arrDepTime + ')\n' +
      '🛬 إلى: ' + arrTo + ' (' + arrArvTime + ')\n' +
      '🎫 النوع: ' + arrType + '\n\n' +
      '✈️ <b>رحلة العودة</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🛫 الرحلة: <b>' + retFlight + '</b>\n' +
      '🏢 الناقل: ' + retAirline + '\n' +
      '📅 التاريخ: ' + retDate + '\n' +
      '🛫 من: ' + retFrom + ' (' + retDepTime + ')\n' +
      '🛬 إلى: ' + retTo + ' (' + retArvTime + ')';
  } else {
    text = '✈️ <b>Arrival Flight</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🛫 Flight: <b>' + arrFlight + '</b>\n' +
      '🏢 Airline: ' + arrAirline + '\n' +
      '📅 Date: ' + arrDate + '\n' +
      '🛫 From: ' + arrFrom + ' (' + arrDepTime + ')\n' +
      '🛬 To: ' + arrTo + ' (' + arrArvTime + ')\n' +
      '🎫 Type: ' + arrType + '\n\n' +
      '✈️ <b>Return Flight</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🛫 Flight: <b>' + retFlight + '</b>\n' +
      '🏢 Airline: ' + retAirline + '\n' +
      '📅 Date: ' + retDate + '\n' +
      '🛫 From: ' + retFrom + ' (' + retDepTime + ')\n' +
      '🛬 To: ' + retTo + ' (' + retArvTime + ')';
  }
  
  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: lang === 'ar' ? '🔙 القائمة الرئيسية' : '🔙 Main Menu', callback_data: 'show_menu' }]]
  });
}

// ============================================
// دوال مساعدة — تنسيق التاريخ والوقت
// ============================================
function formatDate_(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    var d = val;
    return d.getFullYear() + '-' + pad2_(d.getMonth()+1) + '-' + pad2_(d.getDate());
  }
  var s = String(val);
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  return s || '-';
}

function formatTime_(val) {
  if (!val) return '-';
  var s = String(val);
  if (s.indexOf('.') !== -1) s = s.split('.')[0];
  var parts = s.split(':');
  if (parts.length >= 2) return parts[0] + ':' + parts[1];
  return s || '-';
}

function pad2_(n) {
  return n < 10 ? '0' + n : '' + n;
}

// ============================================
// 🚨 أرقام الطوارئ ووزارة الحج
// ============================================
function handleEmergency_(chatId, session) {
  var lang = (session && session.language) || 'ar';
  
  var text = '';
  
  if (lang === 'ar') {
    text = '🚨 <b>أرقام الطوارئ</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '🔴 <b>الطوارئ الموحد</b>\n' +
      '📞 <code>911</code>\n\n' +
      '🕋 <b>وزارة الحج — مساعدات الحجاج</b>\n' +
      '📞 <code>1966</code>\n\n' +
      '🏥 <b>الهلال الأحمر</b>\n' +
      '📞 <code>997</code>';
  } else if (lang === 'fr') {
    text = '🚨 <b>Numéros d\'urgence</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '🔴 <b>Urgences</b>\n' +
      '📞 <code>911</code>\n\n' +
      '🕋 <b>Ministère du Hajj</b>\n' +
      '📞 <code>1966</code>\n\n' +
      '🏥 <b>Croissant-Rouge</b>\n' +
      '📞 <code>997</code>';
  } else if (lang === 'nl') {
    text = '🚨 <b>Noodnummers</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '🔴 <b>Noodgevallen</b>\n' +
      '📞 <code>911</code>\n\n' +
      '🕋 <b>Ministerie van Hadj</b>\n' +
      '📞 <code>1966</code>\n\n' +
      '🏥 <b>Rode Halve Maan</b>\n' +
      '📞 <code>997</code>';
  } else {
    text = '🚨 <b>Emergency Numbers</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '🔴 <b>Unified Emergency</b>\n' +
      '📞 <code>911</code>\n\n' +
      '🕋 <b>Ministry of Hajj — Pilgrim Support</b>\n' +
      '📞 <code>1966</code>\n\n' +
      '🏥 <b>Red Crescent</b>\n' +
      '📞 <code>997</code>';
  }
  
  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: lang === 'ar' ? '🔙 القائمة الرئيسية' : '🔙 Main Menu', callback_data: 'show_menu' }]]
  });
}

// ============================================
// 📞 التواصل مع الشركة
// ============================================
function handleContactCompany_(chatId, session) {
  var lang = (session && session.language) || 'ar';
  
  var text = '';
  
  if (lang === 'ar') {
    text = '📞 <b>تواصل مع إكرام الضيف</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '📱 <b>الهاتف</b>\n' +
      '☎️ <code>8001111061</code>\n\n' +
      '💬 <b>واتساب</b>\n' +
      '📲 <code>+966125111940</code>';
  } else if (lang === 'fr') {
    text = '📞 <b>Contactez Ikram Al-Dayf</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '📱 <b>Téléphone</b>\n' +
      '☎️ <code>8001111061</code>\n\n' +
      '💬 <b>WhatsApp</b>\n' +
      '📲 <code>+966125111940</code>';
  } else if (lang === 'nl') {
    text = '📞 <b>Neem contact op met Ikram Al-Dayf</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '📱 <b>Telefoon</b>\n' +
      '☎️ <code>8001111061</code>\n\n' +
      '💬 <b>WhatsApp</b>\n' +
      '📲 <code>+966125111940</code>';
  } else {
    text = '📞 <b>Contact Ikram Al-Dayf</b>\n' +
      '━━━━━━━━━━━━━━\n\n' +
      '📱 <b>Phone</b>\n' +
      '☎️ <code>8001111061</code>\n\n' +
      '💬 <b>WhatsApp</b>\n' +
      '📲 <code>+966125111940</code>';
  }
  
  var backText = lang === 'ar' ? '🔙 القائمة الرئيسية' : '🔙 Main Menu';
  var waText = lang === 'ar' ? '💬 فتح واتساب' : '💬 Open WhatsApp';
  
  sendMessage_(chatId, text, {
    inline_keyboard: [
      [{ text: waText, url: 'https://wa.me/966125111940' }],
      [{ text: backText, callback_data: 'show_menu' }]
    ]
  });
}

// ============================================
// القائمة الرئيسية
// ============================================
function sendMainMenu_(chatId, lang) {
  var menus = {
    ar: {
      text: '📋 القائمة الرئيسية — اختر ما تريد:',
      buttons: [
        [{ text: '✈️ رحلتي', callback_data: 'my_flight' }, { text: '🏨 فندقي', callback_data: 'my_hotel' }],
        [{ text: '📦 باقتي', callback_data: 'my_package' }, { text: '🚌 مواصلاتي', callback_data: 'my_transport' }],
        [{ text: '🚨 طوارئ', callback_data: 'emergency' }, { text: '📞 تواصل معنا', callback_data: 'contact_company' }],
        [{ text: '❓ أسئلة عن الحج', callback_data: 'general_faq' }],
        [{ text: '🔄 تغيير اللغة', callback_data: 'change_lang' }]
      ]
    },
    en: {
      text: '📋 Main Menu — Choose an option:',
      buttons: [
        [{ text: '✈️ My Flight', callback_data: 'my_flight' }, { text: '🏨 My Hotel', callback_data: 'my_hotel' }],
        [{ text: '📦 My Package', callback_data: 'my_package' }, { text: '🚌 My Transport', callback_data: 'my_transport' }],
        [{ text: '🚨 Emergency', callback_data: 'emergency' }, { text: '📞 Contact Us', callback_data: 'contact_company' }],
        [{ text: '❓ Hajj FAQ', callback_data: 'general_faq' }],
        [{ text: '🔄 Change Language', callback_data: 'change_lang' }]
      ]
    },
    fr: {
      text: '📋 Menu Principal:',
      buttons: [
        [{ text: '✈️ Mon Vol', callback_data: 'my_flight' }, { text: '🏨 Mon Hôtel', callback_data: 'my_hotel' }],
        [{ text: '📦 Mon Forfait', callback_data: 'my_package' }, { text: '🚌 Mon Transport', callback_data: 'my_transport' }],
        [{ text: '🚨 Urgences', callback_data: 'emergency' }, { text: '📞 Nous Contacter', callback_data: 'contact_company' }],
        [{ text: '❓ FAQ Hajj', callback_data: 'general_faq' }],
        [{ text: '🔄 Changer la langue', callback_data: 'change_lang' }]
      ]
    },
    nl: {
      text: '📋 Hoofdmenu:',
      buttons: [
        [{ text: '✈️ Mijn Vlucht', callback_data: 'my_flight' }, { text: '🏨 Mijn Hotel', callback_data: 'my_hotel' }],
        [{ text: '📦 Mijn Pakket', callback_data: 'my_package' }, { text: '🚌 Mijn Vervoer', callback_data: 'my_transport' }],
        [{ text: '🚨 Noodgevallen', callback_data: 'emergency' }, { text: '📞 Contact', callback_data: 'contact_company' }],
        [{ text: '❓ Hadj FAQ', callback_data: 'general_faq' }],
        [{ text: '🔄 Taal wijzigen', callback_data: 'change_lang' }]
      ]
    }
  };
  
  var menu = menus[lang] || menus['en'];
  sendMessage_(chatId, menu.text, { inline_keyboard: menu.buttons });
}

// ============================================
// رسالة الترحيب + أزرار اللغة
// ============================================
function sendWelcome_(chatId, firstName) {
  sendMessage_(chatId, '🕋 مرحباً ' + firstName + '!\n\nأنا مساعد إكرام الضيف للحج.\nاختر لغتك / Choose your language:');
  sendLanguageButtons_(chatId);
}

function sendLanguageButtons_(chatId) {
  var keyboard = {
    inline_keyboard: [
      [{ text: '🇸🇦 العربية', callback_data: 'lang_ar' }, { text: '🇬🇧 English', callback_data: 'lang_en' }],
      [{ text: '🇫🇷 Français', callback_data: 'lang_fr' }, { text: '🇳🇱 Nederlands', callback_data: 'lang_nl' }]
    ]
  };
  sendMessage_(chatId, 'اختر لغتك / Choose your language:', keyboard);
}

// ============================================
// إدارة الجلسات — BotSessions
// ============================================
function getSessionSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName('BotSessions');
}

function getSession_(chatId) {
  var sheet = getSessionSheet_();
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(chatId)) {
      return {
        row: i + 1,
        chatId: String(data[i][0]),
        passport: String(data[i][1]),
        language: String(data[i][3]),
        authStatus: String(data[i][4]),
        authTimestamp: String(data[i][5]),
        lastActivity: String(data[i][6]),
        messageCount: Number(data[i][7]) || 0
      };
    }
  }
  return null;
}

function saveSession_(chatId, fields) {
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
  }
}

function updateSession_(chatId, fields) {
  var sheet = getSessionSheet_();
  var session = getSession_(chatId);
  if (!session) return;
  
  var row = session.row;
  if (fields.passport !== undefined) sheet.getRange(row, 2).setValue(fields.passport);
  if (fields.language !== undefined) sheet.getRange(row, 4).setValue(fields.language);
  if (fields.authStatus !== undefined) sheet.getRange(row, 5).setValue(fields.authStatus);
  if (fields.authTimestamp !== undefined) sheet.getRange(row, 6).setValue(fields.authTimestamp);
  
  sheet.getRange(row, 7).setValue(new Date().toISOString());
  sheet.getRange(row, 8).setValue(session.messageCount + 1);
}

function updateLastActivity_(chatId) {
  updateSession_(chatId, {});
}

// ============================================
// Telegram API
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
// إعداد Webhook — يُشغَّل يدوياً مرة واحدة
// ============================================
function setWebhook() {
  var webAppUrl = 'https://script.google.com/macros/s/AKfycbz_dppqAEGXvIpJjawZCjap4rpGokL6Hw2vSuF3Fd_Lg9uEH3ivgR2mWxuajEl_pJRK/exec';
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: webAppUrl }),
    muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

function getWebhookInfo() {
  var res = UrlFetchApp.fetch(TELEGRAM_API + '/getWebhookInfo', { muteHttpExceptions: true });
  Logger.log(res.getContentText());
}

/**
 * دالة اختبار — شغّلها يدوياً لفحص إرسال الرسائل
 */
function testBot() {
  Logger.log('=== Test 1: BotSessions ===');
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BotSessions');
    Logger.log('Sheet found: ' + (sheet ? 'YES' : 'NO'));
    Logger.log('Sheet name: ' + sheet.getName());
    sheet.appendRow(['TEST123', '', '', 'ar', 'test', '', new Date().toISOString(), 0]);
    Logger.log('Write OK');
  } catch(e) {
    Logger.log('Write ERROR: ' + e.message);
  }
  
  Logger.log('=== Test 2: Journey Sheet ===');
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var jSheet = ss.getSheetByName(JOURNEY_SHEET);
    Logger.log('Journey sheet found: ' + (jSheet ? 'YES' : 'NO'));
    if (jSheet) {
      var headers = jSheet.getRange(1, 1, 1, 15).getValues()[0];
      Logger.log('First 15 headers: ' + headers.join(' | '));
    }
  } catch(e) {
    Logger.log('Journey ERROR: ' + e.message);
  }
  
  Logger.log('=== Test 3: Telegram ===');
  try {
    var res = UrlFetchApp.fetch(TELEGRAM_API + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: '121527633',
        text: '✅ تجربة إرسال من testBot'
      }),
      muteHttpExceptions: true
    });
    Logger.log('TG Response: ' + res.getContentText());
  } catch(e) {
    Logger.log('TG ERROR: ' + e.message);
  }
}
function debugTransport() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  
  // 1. قراءة PackageId من رحلة الحاج لهذا الحاج
  var jSheet = ss.getSheetByName(JOURNEY_SHEET);
  var jData = jSheet.getDataRange().getValues();
  for (var i = 1; i < jData.length; i++) {
    if (String(jData[i][8]).toUpperCase().trim() === '674711081') { // جواز NOORAIN
      Logger.log('Journey PackageId: [' + jData[i][1] + '] type: ' + typeof jData[i][1]);
      break;
    }
  }
  
  // 2. قراءة أول 5 قيم من عمود B في الباقات + رأس التنقل
  var pSheet = ss.getSheetByName('الباقات');
  var pData = pSheet.getDataRange().getValues();
  Logger.log('Row0 headers count: ' + pData[0].length);
  Logger.log('Row1 BN value: [' + pData[1][65] + ']'); // BN = index 65
  for (var j = 2; j < 7 && j < pData.length; j++) {
    Logger.log('Row' + j + ' B=[' + pData[j][1] + '] type=' + typeof pData[j][1] + ' BN=[' + pData[j][65] + ']');
  }
}