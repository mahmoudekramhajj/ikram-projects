// ============================================
// معالجات القائمة — طيران، فندق، باقة، مواصلات، طوارئ، تواصل، تحديث
// ============================================

// ============================================
// ✈️ رحلتي
// ============================================
function handleMyFlight_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;
  var isAr = (lang === 'ar');

  var arrFlight = String(r[24] || '-');
  var arrAirline = isAr ? String(r[16] || '-') : String(r[17] || '-');
  var arrFrom = String(r[21] || '-');
  var arrTo = String(r[19] || '-');
  var arrDate = formatDate_(r[20]);
  var arrDepTime = formatTime_(r[23]);
  var arrArvTime = formatTime_(r[18]);
  var arrType = String(r[25] || '-');

  var retFlight = String(r[34] || '-');
  var retAirline = isAr ? String(r[26] || '-') : String(r[27] || '-');
  var retFrom = String(r[31] || '-');
  var retTo = String(r[29] || '-');
  var retDate = formatDate_(r[32]);
  var retDepTime = formatTime_(r[33]);
  var retArvTime = formatTime_(r[28]);
  var retType = String(r[35] || '-');

  // --- Fetch transit details ---
  var arrTransit = null;
  var retTransit = null;
  var b2cData = null;

  if (arrType === 'B2B') {
    arrTransit = getB2BFlightDetails_(arrFlight, true);
  } else if (arrType === 'B2C') {
    b2cData = getB2CFlightDetails_(session.passport);
    if (b2cData) arrTransit = b2cData.arrival;
  }

  if (retType === 'B2B') {
    retTransit = getB2BFlightDetails_(retFlight, false);
  } else if (retType === 'B2C') {
    var b2cDataRet = b2cData ? b2cData : getB2CFlightDetails_(session.passport);
    if (b2cDataRet) retTransit = b2cDataRet['return'] || null;
  }

  var text = '';
  var legLbl = T_('lbl_leg', lang);
  var transitLbl = T_('lbl_transit', lang);
  var directLbl = T_('lbl_direct', lang);

  // --- Build arrival section ---
  var arrHasTransit = arrTransit && arrTransit.leg1 && arrTransit.leg2;

  if (isAr) {
    text = '✈️ <b>رحلة الوصول</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 الناقل: ' + arrAirline + '\n' +
      '🎫 النوع: ' + arrType + '\n';
  } else {
    text = '✈️ <b>Arrival Flight</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 ' + (isAr ? 'الناقل' : 'Airline') + ': ' + arrAirline + '\n' +
      '🎫 ' + (isAr ? 'النوع' : 'Type') + ': ' + arrType + '\n';
  }

  if (arrHasTransit) {
    var a1 = arrTransit.leg1;
    var a2 = arrTransit.leg2;
    text += '🛫 ' + legLbl + ' 1: <b>' + a1.flightNo + '</b> ' + a1.from + ' → ' + a1.to + ' (' + a1.timeDepart + ')\n' +
      '✋ ' + transitLbl + '\n' +
      '🛫 ' + legLbl + ' 2: <b>' + a2.flightNo + '</b> ' + a2.from + ' → ' + a2.to + ' (' + a2.timeDepart + ')\n' +
      '📅 ' + (isAr ? 'التاريخ' : 'Date') + ': ' + a1.dateDepart + '\n';
  } else {
    text += '🛫 ' + (isAr ? 'الرحلة' : 'Flight') + ': <b>' + arrFlight + '</b> (' + directLbl + ')\n' +
      '📅 ' + (isAr ? 'التاريخ' : 'Date') + ': ' + arrDate + '\n' +
      '🛫 ' + (isAr ? 'من' : 'From') + ': ' + arrFrom + ' (' + arrDepTime + ')\n' +
      '🛬 ' + (isAr ? 'إلى' : 'To') + ': ' + arrTo + ' (' + arrArvTime + ')\n';
  }

  // --- Build return section ---
  var retHasTransit = retTransit && retTransit.leg1 && retTransit.leg2;

  text += '\n';
  if (isAr) {
    text += '✈️ <b>رحلة العودة</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 الناقل: ' + retAirline + '\n';
  } else {
    text += '✈️ <b>Return Flight</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 ' + (isAr ? 'الناقل' : 'Airline') + ': ' + retAirline + '\n';
  }

  if (retHasTransit) {
    var rt1 = retTransit.leg1;
    var rt2 = retTransit.leg2;
    text += '🛫 ' + legLbl + ' 1: <b>' + rt1.flightNo + '</b> ' + rt1.from + ' → ' + rt1.to + ' (' + rt1.timeDepart + ')\n' +
      '✋ ' + transitLbl + '\n' +
      '🛫 ' + legLbl + ' 2: <b>' + rt2.flightNo + '</b> ' + rt2.from + ' → ' + rt2.to + ' (' + rt2.timeDepart + ')\n' +
      '📅 ' + (isAr ? 'التاريخ' : 'Date') + ': ' + rt1.dateDepart;
  } else {
    text += '🛫 ' + (isAr ? 'الرحلة' : 'Flight') + ': <b>' + retFlight + '</b> (' + directLbl + ')\n' +
      '📅 ' + (isAr ? 'التاريخ' : 'Date') + ': ' + retDate + '\n' +
      '🛫 ' + (isAr ? 'من' : 'From') + ': ' + retFrom + ' (' + retDepTime + ')\n' +
      '🛬 ' + (isAr ? 'إلى' : 'To') + ': ' + retTo + ' (' + retArvTime + ')';
  }

  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// 🏨 فندقي (2 أو 3 فنادق + خريطة)
// ============================================
function handleMyHotel_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;
  var isAr = (lang === 'ar');

  var house1 = String(r[36] || '-');
  var house1Start = formatDate_(r[37]);
  var house1End = formatDate_(r[38]);
  var house1Hotel = '';
  if (house1.toLowerCase().indexOf('madi') !== -1) {
    house1Hotel = isAr ? String(r[46] || '-') : String(r[47] || '-');
  } else {
    house1Hotel = isAr ? String(r[42] || '-') : String(r[43] || '-');
  }

  var house2 = String(r[39] || '-');
  var house2Start = formatDate_(r[40]);
  var house2End = formatDate_(r[41]);
  var house2Hotel = '';
  if (house2.toLowerCase().indexOf('madi') !== -1) {
    house2Hotel = isAr ? String(r[46] || '-') : String(r[47] || '-');
  } else if (house2.toLowerCase().indexOf('shift') !== -1) {
    house2Hotel = isAr ? String(r[44] || '-') : String(r[45] || '-');
  } else {
    house2Hotel = isAr ? String(r[42] || '-') : String(r[43] || '-');
  }

  var shiftHotelAr = String(r[44] || '').trim();
  var shiftHotelEn = String(r[45] || '').trim();
  var makkahHotelAr = String(r[42] || '').trim();
  var hasThirdHotel = shiftHotelAr && shiftHotelAr !== '-' && shiftHotelAr !== 'NULL' && shiftHotelAr !== 'null' && shiftHotelAr !== '' && shiftHotelAr !== makkahHotelAr;

  var accLbl = T_('lbl_acc', lang);
  var hotelLbl = T_('lbl_hotel', lang);
  var inLbl = T_('lbl_checkin', lang);
  var outLbl = T_('lbl_checkout', lang);

  var text = '🏨 <b>' + accLbl + ' 1 — ' + house1 + '</b>\n' +
    '━━━━━━━━━━━━━━\n' +
    '🏢 ' + hotelLbl + ': <b>' + house1Hotel + '</b>\n' +
    '📅 ' + inLbl + ': ' + house1Start + '\n' +
    '📅 ' + outLbl + ': ' + house1End + '\n\n';

  text += '🏨 <b>' + accLbl + ' 2 — ' + house2 + '</b>\n' +
    '━━━━━━━━━━━━━━\n' +
    '🏢 ' + hotelLbl + ': <b>' + house2Hotel + '</b>\n' +
    '📅 ' + inLbl + ': ' + house2Start + '\n' +
    '📅 ' + outLbl + ': ' + house2End;

  var thirdHotelName = '';
  if (hasThirdHotel) {
    thirdHotelName = isAr ? shiftHotelAr : (shiftHotelEn || shiftHotelAr);
    var packageId = String(r[1] || '').trim();
    var h3dates = getThirdHotelDates_(packageId);
    var h3In = '-';
    var h3Out = '-';

    if (h3dates) {
      h3In = formatDate_(h3dates.checkIn);

      // تحديد الحالة: هل الفندق الثالث هو آخر فندق قبل المطار؟
      // الحالة 2: المدينة → مكة 1 → مكة 2 → المطار (house1 يبدأ بالمدينة)
      var isLastBeforeAirport = house1.toLowerCase().indexOf('madi') !== -1;

      if (isLastBeforeAirport) {
        // حساب الخروج من رحلة العودة
        h3Out = calcThirdHotelCheckout_(r[32], r[33], h3dates.packageEnd, h3dates.checkOut);
      } else {
        // الحالة 1: نلتزم بتواريخ الباقات
        h3Out = formatDate_(h3dates.checkOut);
      }
    }

    text += '\n\n🏨 <b>' + accLbl + ' 3 — ' + T_('lbl_shifting', lang) + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 ' + hotelLbl + ': <b>' + thirdHotelName + '</b>\n' +
      '📅 ' + inLbl + ': ' + h3In + '\n' +
      '📅 ' + outLbl + ': ' + h3Out;
  }

  var mapButtons = [];
  var mapPfx = T_('lbl_map', lang);

  Logger.log('DEBUG hotel1=[' + house1Hotel + '] hotel2=[' + house2Hotel + ']');

  var map1 = getHotelMapLink_(house1Hotel);
  Logger.log('DEBUG map1=' + map1);
  if (map1) mapButtons.push([{ text: mapPfx + house1Hotel, url: map1 }]);

  var map2 = getHotelMapLink_(house2Hotel);
  Logger.log('DEBUG map2=' + map2);
  if (map2 && house2Hotel !== house1Hotel) mapButtons.push([{ text: mapPfx + house2Hotel, url: map2 }]);

  if (hasThirdHotel) {
    // نبحث بالعربي أولاً ثم بالإنجليزي
    var map3 = getHotelMapLink_(shiftHotelAr) || getHotelMapLink_(shiftHotelEn);
    if (map3) {
      mapButtons.push([{ text: mapPfx + thirdHotelName, url: map3 }]);
    }
  }

  Logger.log('DEBUG mapButtons=' + JSON.stringify(mapButtons));
  mapButtons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);
  sendMessage_(chatId, text, { inline_keyboard: mapButtons });
}

// ============================================
// 📦 باقتي + ⛺ المخيم
// ============================================
function handleMyPackage_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;
  var isAr = (lang === 'ar');

  var packageId = String(r[1] || '-');
  var campName = String(r[4] || '-');
  var groupNo = String(r[6] || '-');
  var nationality = isAr ? String(r[13] || '-') : String(r[12] || '-');
  var country = isAr ? String(r[15] || '-') : String(r[14] || '-');
  var flightType = String(r[25] || '-');

  var pkgNameObj = getPackageName_(packageId);
  var pkgName = '-';
  if (pkgNameObj) {
    pkgName = isAr ? pkgNameObj.nameAr : pkgNameObj.nameEn;
  }
  var pkgNameLbl = T_('lbl_pkg_name', lang);

  var guide = getTourGuide_(packageId);
  var guideLbl = T_('lbl_guide', lang);
  var guidePhoneLbl = T_('lbl_guide_phone', lang);

  var text = '';

  if (isAr) {
    text = '📦 <b>بيانات الباقة</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🔢 رقم الباقة: <b>' + packageId + '</b>\n' +
      '📋 ' + pkgNameLbl + ': <b>' + pkgName + '</b>\n' +
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
      '📋 ' + pkgNameLbl + ': <b>' + pkgName + '</b>\n' +
      '👥 Group: ' + groupNo + '\n' +
      '🎫 Ticket Type: ' + flightType + '\n' +
      '🌍 Nationality: ' + nationality + '\n' +
      '🏠 Residence: ' + country + '\n\n' +
      '⛺ <b>Mina Camp</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '📍 Location: <b>' + campName + '</b>';
  }

  text += '\n\n👨‍✈️ <b>' + guideLbl + '</b>\n' +
    '━━━━━━━━━━━━━━\n' +
    '🔄 ' + T_('lbl_under_review', lang);

  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// 🚌 مواصلاتي
// ============================================
function handleMyTransport_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;
  var isAr = (lang === 'ar');

  var arrFrom = String(r[21] || '-');
  var arrTo = String(r[19] || '-');
  var arrDate = formatDate_(r[20]);
  var arrFlight = String(r[24] || '-');

  var retFrom = String(r[31] || '-');
  var retTo = String(r[29] || '-');
  var retDate = formatDate_(r[32]);
  var retFlight = String(r[34] || '-');

  var house1 = String(r[36] || '-');
  var house2 = String(r[39] || '-');
  var house2Start = formatDate_(r[40]);
  var campName = String(r[4] || '-');

  var packageId = String(r[1] || '');
  var transportType = getTransportType_(packageId);
  var transportIcon = (transportType === 'قطار') ? '🚆' : '🚌';
  var transportTypeEn = transportType === 'قطار' ? 'Train' : transportType === 'حافلة' ? 'Bus' : transportType;

  var text = '';

  if (isAr) {
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
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// 🚨 أرقام الطوارئ
// ============================================
function handleEmergency_(chatId, session) {
  var lang = (session && session.language) || 'ar';

  var emergencyData = {
    ar: ['أرقام الطوارئ', 'الطوارئ الموحد', 'وزارة الحج — مساعدات الحجاج', 'الهلال الأحمر'],
    en: ['Emergency Numbers', 'Unified Emergency', 'Ministry of Hajj — Pilgrim Support', 'Red Crescent'],
    fr: ['Numéros d\'urgence', 'Urgences', 'Ministère du Hajj', 'Croissant-Rouge'],
    de: ['Notrufnummern', 'Einheitlicher Notruf', 'Hajj-Ministerium — Pilgerhilfe', 'Roter Halbmond'],
    it: ['Numeri di emergenza', 'Emergenza unificata', 'Ministero del Hajj — Supporto pellegrini', 'Mezzaluna Rossa'],
    es: ['Números de emergencia', 'Emergencia unificada', 'Ministerio del Hajj — Apoyo a peregrinos', 'Media Luna Roja']
  };

  var e = emergencyData[lang] || emergencyData['en'];

  var text = '🚨 <b>' + e[0] + '</b>\n' +
    '━━━━━━━━━━━━━━\n\n' +
    '🔴 <b>' + e[1] + '</b>\n📞 <code>911</code>\n\n' +
    '🕋 <b>' + e[2] + '</b>\n📞 <code>1966</code>\n\n' +
    '🏥 <b>' + e[3] + '</b>\n📞 <code>997</code>';

  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// 📞 التواصل مع الشركة
// ============================================
function handleContactCompany_(chatId, session) {
  var lang = (session && session.language) || 'ar';

  var contactData = {
    ar: ['تواصل مع إكرام الضيف', 'الهاتف', 'واتساب'],
    en: ['Contact Ikram Al-Dayf', 'Phone', 'WhatsApp'],
    fr: ['Contactez Ikram Al-Dayf', 'Téléphone', 'WhatsApp'],
    de: ['Kontaktieren Sie Ikram Al-Dayf', 'Telefon', 'WhatsApp'],
    it: ['Contatta Ikram Al-Dayf', 'Telefono', 'WhatsApp'],
    es: ['Contacta a Ikram Al-Dayf', 'Teléfono', 'WhatsApp']
  };

  var c = contactData[lang] || contactData['en'];

  var text = '📞 <b>' + c[0] + '</b>\n' +
    '━━━━━━━━━━━━━━\n\n' +
    '📱 <b>' + c[1] + '</b>\n☎️ <code>8001111061</code>\n\n' +
    '💬 <b>' + c[2] + '</b>\n📲 <code>+966125111940</code>';

  sendMessage_(chatId, text, {
    inline_keyboard: [
      [{ text: T_('btn_whatsapp', lang), url: 'https://wa.me/966125111940' }],
      [{ text: T_('btn_back', lang), callback_data: 'show_menu' }]
    ]
  });
}

// ============================================
// ✅ تأكيد الوصول
// ============================================
function handleConfirmArrival_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;
  var arriveDate = normDate_(r[20]); // ArrivalArriveDate
  var today = getTodayString_();
  var yesterday = getDateOffset_(today, -1);
  var tomorrow = getDateOffset_(today, 1);

  // التحقق من التاريخ (نافذة 3 أيام: أمس + اليوم + غداً)
  if (arriveDate !== today && arriveDate !== yesterday && arriveDate !== tomorrow) {
    sendMessage_(chatId, T_('arrival_not_today', lang, { date: formatDate_(r[20]) }), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
    return;
  }

  // التحقق إذا مؤكد مسبقاً
  var reception = getReceptionStatus_(pilgrim.row);
  if (reception && reception.status === 'تم') {
    sendMessage_(chatId, T_('arrival_already', lang, { time: reception.time }), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
    return;
  }

  // كتابة التأكيد في الشيت
  var source = 'Bot:' + String(r[7] || '').substring(0, 30);
  var success = confirmArrivalInSheet_(pilgrim.row, source);

  if (!success) {
    sendMessage_(chatId, '❌ حدث خطأ. حاول مرة أخرى.');
    return;
  }

  // مسح الكاش
  clearPilgrimCache_(session.passport);

  // إرسال رسالة نجاح للحاج
  sendMessage_(chatId, T_('arrival_confirmed', lang), {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });

  // إرسال إشعار لمجموعة العمليات المناسبة
  try {
    var arriveCity = String(r[19] || '');
    var airlineEn = String(r[17] || '');
    var groupChatId = getOpsGroupChatId_(arriveCity, airlineEn);

    var name = String(r[7] || '-');
    var passport = String(r[8] || '-');
    var flight = String(r[24] || '-');
    var nationality = String(r[12] || '-');
    var house1 = String(r[36] || '-');
    var hotel = '';
    if (house1.toLowerCase().indexOf('madi') !== -1) {
      hotel = String(r[47] || '-');
    } else {
      hotel = String(r[43] || '-');
    }

    var opsMsg = '✅ <b>تأكيد وصول (بوت)</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '👤 ' + name + '\n' +
      '🛂 <code>' + passport + '</code>\n' +
      '🌍 ' + nationality + '\n' +
      '✈️ ' + flight + ' → ' + arriveCity + '\n' +
      '🏨 ' + hotel + ' (' + house1 + ')';

    sendMessage_(groupChatId, opsMsg);
  } catch (e) {
    Logger.log('Ops notification error: ' + e.message);
  }
}

// ============================================
// 📱 QR Code — كود خاص لكل حاج
// ============================================
function handleMyQR_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;

  // بناء بيانات QR (JSON مضغوط — مفاتيح قصيرة)
  var qrData = {
    n: String(r[7] || ''),                      // Name
    p: String(r[8] || ''),                      // Passport
    pkg: String(r[1] || ''),                    // PackageId
    grp: String(r[6] || ''),                    // Group Number
    nat: String(r[12] || ''),                   // Nationality
    f: String(r[24] || ''),                     // Arrival Flight
    arr: normDate_(r[20]) || '',                // Arrival Date
    city: String(r[19] || ''),                  // Arrive City
    h1: String(r[36] || ''),                    // FirstHouse
    h2: String(r[39] || '')                     // LastHouse
  };

  var jsonStr = JSON.stringify(qrData);
  var encoded = encodeURIComponent(jsonStr);

  // QR Server API (مجاني وفعّال)
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&format=png&data=' + encoded;

  var caption = T_('qr_caption', lang);

  sendPhoto_(chatId, qrUrl, caption, {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// 📢 الإعلانات — قائمة آخر الرسائل
// ============================================
function handleAnnouncements_(chatId, session) {
  var lang = (session && session.language) || 'ar';
  var messages = getRecentMessages_(session.passport, 5);

  if (!messages || messages.length === 0) {
    sendMessage_(chatId, T_('ann_title', lang) + '\n\n' + T_('ann_empty', lang), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
    return;
  }

  var buttons = [];
  for (var i = 0; i < messages.length; i++) {
    var m = messages[i];
    var label = '';
    if (m.priority === 'urgent') label += '🚨 ';
    label += m.title || ('📢 #' + m.id);
    if (m.sentAt) {
      var d = m.sentAt.substring(0, 10);
      label += ' — ' + d;
    }
    buttons.push([{ text: label, callback_data: 'ann_detail_' + m.id }]);
  }
  buttons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);

  sendMessage_(chatId, T_('ann_title', lang), { inline_keyboard: buttons });
}

// ============================================
// 📢 تفاصيل إعلان واحد
// ============================================
function handleAnnDetail_(chatId, session, msgId) {
  var lang = (session && session.language) || 'ar';
  var msgRow = getMessageById_(msgId);

  if (!msgRow) {
    sendMessage_(chatId, T_('ann_empty', lang));
    return;
  }

  // جلب النص بلغة الحاج
  var langCol = LANG_TO_COL[lang] || AM.MSG_EN;
  var text = String(msgRow[langCol] || '').trim();
  if (!text) text = String(msgRow[AM.MSG_EN] || msgRow[AM.MSG_AR] || '').trim();

  var priority = String(msgRow[AM.PRIORITY] || '').toLowerCase().trim();
  var header = priority === 'urgent' ? T_('ann_urgent', lang) + ' ' : '';
  var fullText = T_('ann_from_admin', lang) + header + text;

  var backBtn = { inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'announcements' }]] };
  sendMessage_(chatId, fullText, backBtn);

  // إرسال صورة إن وُجدت
  var imageUrl = String(msgRow[AM.IMAGE_URL] || '').trim();
  if (imageUrl) {
    sendPhoto_(chatId, getDriveDirectUrl_(imageUrl), '');
  }

  // إرسال ملف إن وُجد
  var fileUrl = String(msgRow[AM.FILE_URL] || '').trim();
  if (fileUrl) {
    var fileName = String(msgRow[AM.FILE_NAME] || '').trim();
    sendDocument_(chatId, getDriveDirectUrl_(fileUrl), fileName, T_('ann_file_attached', lang));
  }
}

// ============================================
// 🔄 تحديث البيانات
// ============================================
function handleRefreshData_(chatId, session) {
  var lang = (session && session.language) || 'ar';
  var passport = session.passport;

  if (passport) {
    // مسح كل الكاش المتعلق بالحاج
    clearPilgrimCache_(passport);

    // إعادة جلب من الشيت مباشرة (تجاوز الكاش)
    var pilgrim = findPilgrimByPassport_(passport, true);
    if (pilgrim) {
      var pkgId = String(pilgrim.rowData[1] || '');
      clearTransportCache_(pkgId);
      clearHotelMapCache_(pilgrim.rowData);
    }
  }

  sendMessage_(chatId, T_('data_refreshed', lang));
  sendMainMenu_(chatId, lang);
}

// ============================================
// 🎫 التذكرة والتأشيرة
// ============================================
function handleVisaTicket_(chatId, session) {
  var lang = session.language || 'ar';
  var isAr = (lang === 'ar');

  var pilgrim = findPilgrimByPassport_(session.passport);
  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var info = getVisaAndTicket_(session.passport);
  var visaLink = getVisaImage_(session.passport);

  var visaStatusLbl = T_('lbl_visa_status', lang);
  var ticketNoLbl = T_('lbl_ticket_no', lang);

  var visaStatus = (info && info.visaStatus && info.visaStatus !== '-') ? info.visaStatus : '-';
  var ticketNo = (info && info.ticketNo) ? info.ticketNo : '';
  var ticketLink = (info && info.ticketLink) ? info.ticketLink : '';

  var text = '';
  if (isAr) {
    text = '🎫 <b>التذكرة والتأشيرة</b>\n━━━━━━━━━━━━━━\n\n';
    text += '📋 <b>' + visaStatusLbl + '</b>\n' + visaStatus + '\n\n';
    if (ticketNo) {
      text += '🎟️ <b>' + ticketNoLbl + '</b>\n<code>' + ticketNo + '</code>';
    } else {
      text += '🎟️ ' + T_('lbl_ticket_not_ready', lang);
    }
    if (!visaLink) {
      text += '\n\n📋 ' + T_('lbl_visa_not_ready', lang);
    }
  } else {
    text = '🎫 <b>Ticket & Visa</b>\n━━━━━━━━━━━━━━\n\n';
    text += '📋 <b>' + visaStatusLbl + '</b>\n' + visaStatus + '\n\n';
    if (ticketNo) {
      text += '🎟️ <b>' + ticketNoLbl + '</b>\n<code>' + ticketNo + '</code>';
    } else {
      text += '🎟️ ' + T_('lbl_ticket_not_ready', lang);
    }
    if (!visaLink) {
      text += '\n\n📋 ' + T_('lbl_visa_not_ready', lang);
    }
  }

  // بناء الأزرار
  var buttons = [];
  if (ticketLink) {
    buttons.push([{ text: T_('btn_download_ticket', lang), url: ticketLink }]);
  }
  if (visaLink) {
    buttons.push([{ text: T_('btn_view_visa', lang), url: visaLink }]);
  }
  buttons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);

  sendMessage_(chatId, text, { inline_keyboard: buttons });
}
