// ============================================
// جمع بيانات الحاج — هاتف سعودي + أرقام غرف + صورة جواز
// شيت: Pilgrim Data
// أعمدة: Passport | Name | SaudiPhone | Room1 | Room2 | Room3 | PassportPhotoURL | UpdatedAt
// ============================================

// ============================================
// الحصول على شيت Pilgrim Data (إنشاء تلقائي إذا لم يوجد)
// ============================================
function getPilgrimDataSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(PILGRIM_DATA_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PILGRIM_DATA_SHEET);
    sheet.appendRow(['Passport', 'Name', 'SaudiPhone', 'Room1', 'Room2', 'Room3', 'PassportPhotoURL', 'UpdatedAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================
// قراءة بيانات الحاج المجمّعة
// ============================================
function getPilgrimData_(passport) {
  if (!passport) return null;
  var key = String(passport).toUpperCase().trim();
  var cacheKey = 'pdata_' + key;

  var cached = getCache_(cacheKey);
  if (cached) return cached;

  var sheet = getPilgrimDataSheet_();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase().trim() === key) {
      var result = {
        row: i + 1,
        passport: String(data[i][0]),
        name: String(data[i][1]),
        saudiPhone: String(data[i][2]),
        room1: String(data[i][3]),
        room2: String(data[i][4]),
        room3: String(data[i][5]),
        passportPhoto: String(data[i][6]),
        updatedAt: String(data[i][7])
      };
      setCache_(cacheKey, result);
      return result;
    }
  }
  return null;
}

// ============================================
// حفظ/تحديث بيان واحد في Pilgrim Data
// ============================================
function savePilgrimField_(passport, pilgrimName, fieldIndex, value) {
  var key = String(passport).toUpperCase().trim();
  var sheet = getPilgrimDataSheet_();
  var data = sheet.getDataRange().getValues();
  var row = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase().trim() === key) {
      row = i + 1;
      break;
    }
  }

  if (row === -1) {
    // إنشاء صف جديد
    var newRow = [key, pilgrimName, '', '', '', '', '', new Date().toISOString()];
    newRow[fieldIndex] = value;
    sheet.appendRow(newRow);
  } else {
    // تحديث الحقل المحدد
    sheet.getRange(row, fieldIndex + 1).setValue(value);
    sheet.getRange(row, 8).setValue(new Date().toISOString()); // UpdatedAt
  }

  // مسح الكاش
  var cache = CacheService.getScriptCache();
  cache.remove('pdata_' + key);
}

// ============================================
// حفظ رقم الهاتف السعودي
// ============================================
function handlePhoneInput_(chatId, text, session) {
  var lang = session.language || 'ar';
  var phone = text.replace(/[^0-9+]/g, '');

  // تحقق: يبدأ بـ 05 (10 أرقام) أو +966 (12-13 رقم)
  var valid = /^(05\d{8}|(\+?966)\d{9})$/.test(phone);
  if (!valid) {
    sendMessage_(chatId, T_('phone_invalid', lang));
    return;
  }

  // توحيد الصيغة: 05xxxxxxxx
  if (phone.indexOf('+966') === 0) phone = '0' + phone.substring(4);
  if (phone.indexOf('966') === 0) phone = '0' + phone.substring(3);

  var pilgrim = findPilgrimByPassport_(session.passport);
  var name = pilgrim ? pilgrim.name : '';

  savePilgrimField_(session.passport, name, 2, phone);
  updateSession_(chatId, { inputState: '' });

  sendMessage_(chatId, T_('phone_saved', lang, { phone: phone }));
  sendMainMenu_(chatId, lang);
}

// ============================================
// حفظ رقم الغرفة
// ============================================
function handleRoomInput_(chatId, text, session) {
  var lang = session.language || 'ar';
  var roomNo = text.trim();

  if (roomNo.length < 1 || roomNo.length > 10) {
    sendMessage_(chatId, T_('room_invalid', lang));
    return;
  }

  var inputState = session.inputState || '';
  var fieldIndex = 3; // Room1
  var hotelLabel = '1';
  if (inputState === 'awaiting_room_2') { fieldIndex = 4; hotelLabel = '2'; }
  if (inputState === 'awaiting_room_3') { fieldIndex = 5; hotelLabel = '3'; }

  var pilgrim = findPilgrimByPassport_(session.passport);
  var name = pilgrim ? pilgrim.name : '';

  savePilgrimField_(session.passport, name, fieldIndex, roomNo);
  updateSession_(chatId, { inputState: '' });

  sendMessage_(chatId, T_('room_saved', lang, { room: roomNo, hotel: hotelLabel }));
  sendMainMenu_(chatId, lang);
}

// ============================================
// حفظ صورة الجواز على Google Drive
// ============================================
function handlePassportPhoto_(chatId, fileId, session) {
  var lang = session.language || 'ar';

  try {
    // جلب معلومات الملف من تيليغرام
    var fileInfo = getFile_(fileId);
    if (!fileInfo || !fileInfo.file_path) {
      sendMessage_(chatId, T_('photo_error', lang));
      updateSession_(chatId, { inputState: '' });
      return;
    }

    // تحميل الصورة
    var fileUrl = 'https://api.telegram.org/file/bot' + BOT_TOKEN + '/' + fileInfo.file_path;
    var blob = UrlFetchApp.fetch(fileUrl).getBlob();

    // تحديد اسم الملف
    var passport = String(session.passport).toUpperCase().trim();
    var ext = fileInfo.file_path.split('.').pop() || 'jpg';
    blob.setName(passport + '.' + ext);

    // حفظ على Google Drive
    var folder = getDriveFolder_();
    if (!folder) {
      sendMessage_(chatId, T_('photo_error', lang));
      updateSession_(chatId, { inputState: '' });
      return;
    }

    // حذف نسخة قديمة إن وجدت
    var existing = folder.getFilesByName(passport + '.' + ext);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileLink = file.getUrl();

    // حفظ الرابط في الشيت
    var pilgrim = findPilgrimByPassport_(passport);
    var name = pilgrim ? pilgrim.name : '';
    savePilgrimField_(passport, name, 6, fileLink);

    updateSession_(chatId, { inputState: '' });
    sendMessage_(chatId, T_('photo_saved', lang));
    sendMainMenu_(chatId, lang);

  } catch (err) {
    Logger.log('Photo save error: ' + err.message);
    sendMessage_(chatId, T_('photo_error', lang));
    updateSession_(chatId, { inputState: '' });
  }
}

// ============================================
// الحصول على مجلد Google Drive
// ============================================
function getDriveFolder_() {
  if (DRIVE_FOLDER_ID) {
    try { return DriveApp.getFolderById(DRIVE_FOLDER_ID); } catch(e) {}
  }
  // إنشاء مجلد تلقائي إذا لم يُحدد
  var folders = DriveApp.getFoldersByName('IkramHajjBot_Passports');
  if (folders.hasNext()) return folders.next();
  var folder = DriveApp.createFolder('IkramHajjBot_Passports');
  Logger.log('Created Drive folder: ' + folder.getId() + ' — Update DRIVE_FOLDER_ID in Config.js');
  return folder;
}

// ============================================
// عرض بياناتي المجمّعة
// ============================================
function handleMyData_(chatId, session) {
  var lang = session.language || 'ar';
  var pdata = getPilgrimData_(session.passport);

  var phone = (pdata && pdata.saudiPhone) ? pdata.saudiPhone : '-';
  var room1 = (pdata && pdata.room1) ? pdata.room1 : '-';
  var room2 = (pdata && pdata.room2) ? pdata.room2 : '-';
  var room3 = (pdata && pdata.room3) ? pdata.room3 : '-';
  var photo = (pdata && pdata.passportPhoto) ? '✅' : '❌';

  var pilgrim = findPilgrimByPassport_(session.passport);
  var r = pilgrim ? pilgrim.rowData : null;

  // أسماء الفنادق
  var hotel1Name = '-';
  var hotel2Name = '-';
  var hotel3Name = '-';
  if (r) {
    var isAr = (lang === 'ar');
    var house1 = String(r[36] || '');
    if (house1.toLowerCase().indexOf('madi') !== -1) {
      hotel1Name = isAr ? String(r[46] || '-') : String(r[47] || '-');
    } else {
      hotel1Name = isAr ? String(r[42] || '-') : String(r[43] || '-');
    }
    var house2 = String(r[39] || '');
    if (house2.toLowerCase().indexOf('madi') !== -1) {
      hotel2Name = isAr ? String(r[46] || '-') : String(r[47] || '-');
    } else if (house2.toLowerCase().indexOf('shift') !== -1) {
      hotel2Name = isAr ? String(r[44] || '-') : String(r[45] || '-');
    } else {
      hotel2Name = isAr ? String(r[42] || '-') : String(r[43] || '-');
    }
    var shiftAr = String(r[44] || '').trim();
    var makkahAr = String(r[42] || '').trim();
    if (shiftAr && shiftAr !== '-' && shiftAr !== 'NULL' && shiftAr !== '' && shiftAr !== makkahAr) {
      hotel3Name = isAr ? shiftAr : (String(r[45] || '') || shiftAr);
    }
  }

  var text = '';
  if (lang === 'ar') {
    text = '📋 <b>بياناتي المسجّلة</b>\n━━━━━━━━━━━━━━\n\n' +
      '📱 الهاتف السعودي: <b>' + phone + '</b>\n\n' +
      '🏨 رقم الغرفة — ' + hotel1Name + ': <b>' + room1 + '</b>\n' +
      '🏨 رقم الغرفة — ' + hotel2Name + ': <b>' + room2 + '</b>\n';
    if (hotel3Name !== '-') text += '🏨 رقم الغرفة — ' + hotel3Name + ': <b>' + room3 + '</b>\n';
    text += '\n📸 صورة الجواز: ' + photo;
  } else {
    text = '📋 <b>My Registered Data</b>\n━━━━━━━━━━━━━━\n\n' +
      '📱 Saudi Phone: <b>' + phone + '</b>\n\n' +
      '🏨 Room — ' + hotel1Name + ': <b>' + room1 + '</b>\n' +
      '🏨 Room — ' + hotel2Name + ': <b>' + room2 + '</b>\n';
    if (hotel3Name !== '-') text += '🏨 Room — ' + hotel3Name + ': <b>' + room3 + '</b>\n';
    text += '\n📸 Passport Photo: ' + photo;
  }

  // أزرار تحديث البيانات
  var buttons = [];
  if (phone === '-') buttons.push([{ text: T_('btn_add_phone', lang), callback_data: 'add_phone' }]);
  buttons.push([{ text: T_('btn_add_room', lang), callback_data: 'add_room' }]);
  buttons.push([{ text: T_('btn_add_photo', lang), callback_data: 'add_passport_photo' }]);
  buttons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);

  sendMessage_(chatId, text, { inline_keyboard: buttons });
}

// ============================================
// بدء جمع رقم الهاتف
// ============================================
function promptPhone_(chatId, session) {
  var lang = session.language || 'ar';
  updateSession_(chatId, { inputState: 'awaiting_phone' });
  sendMessage_(chatId, T_('phone_prompt', lang));
}

// ============================================
// عرض اختيار الفندق لإدخال رقم الغرفة
// ============================================
function promptRoomSelection_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);
  if (!pilgrim) { sendMessage_(chatId, T_('data_not_found', lang)); return; }

  var r = pilgrim.rowData;
  var isAr = (lang === 'ar');

  var house1 = String(r[36] || '');
  var house2 = String(r[39] || '');
  var hotel1 = '', hotel2 = '', hotel3 = '';

  if (house1.toLowerCase().indexOf('madi') !== -1) {
    hotel1 = isAr ? String(r[46] || '-') : String(r[47] || '-');
  } else {
    hotel1 = isAr ? String(r[42] || '-') : String(r[43] || '-');
  }
  if (house2.toLowerCase().indexOf('madi') !== -1) {
    hotel2 = isAr ? String(r[46] || '-') : String(r[47] || '-');
  } else if (house2.toLowerCase().indexOf('shift') !== -1) {
    hotel2 = isAr ? String(r[44] || '-') : String(r[45] || '-');
  } else {
    hotel2 = isAr ? String(r[42] || '-') : String(r[43] || '-');
  }

  var shiftAr = String(r[44] || '').trim();
  var makkahAr = String(r[42] || '').trim();
  var hasThird = shiftAr && shiftAr !== '-' && shiftAr !== 'NULL' && shiftAr !== '' && shiftAr !== makkahAr;
  if (hasThird) {
    hotel3 = isAr ? shiftAr : (String(r[45] || '') || shiftAr);
  }

  var buttons = [
    [{ text: '🏨 ' + hotel1, callback_data: 'room_hotel_1' }],
    [{ text: '🏨 ' + hotel2, callback_data: 'room_hotel_2' }]
  ];
  if (hasThird) {
    buttons.push([{ text: '🏨 ' + hotel3, callback_data: 'room_hotel_3' }]);
  }
  buttons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);

  sendMessage_(chatId, T_('room_select_hotel', lang), { inline_keyboard: buttons });
}

// ============================================
// بدء جمع رقم الغرفة لفندق محدد
// ============================================
function promptRoom_(chatId, session, hotelNum) {
  var lang = session.language || 'ar';
  updateSession_(chatId, { inputState: 'awaiting_room_' + hotelNum });
  sendMessage_(chatId, T_('room_prompt', lang));
}

// ============================================
// بدء جمع صورة الجواز
// ============================================
function promptPassportPhoto_(chatId, session) {
  var lang = session.language || 'ar';
  updateSession_(chatId, { inputState: 'awaiting_passport_photo' });
  sendMessage_(chatId, T_('photo_prompt', lang));
}
