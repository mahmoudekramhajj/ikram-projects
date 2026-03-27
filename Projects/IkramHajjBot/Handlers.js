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
    text += '\n\n🏨 <b>' + accLbl + ' 3 — ' + T_('lbl_shifting', lang) + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '🏢 ' + hotelLbl + ': <b>' + thirdHotelName + '</b>';
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
    var map3 = getHotelMapLink_(thirdHotelName);
    if (map3 && thirdHotelName !== house1Hotel && thirdHotelName !== house2Hotel) {
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

  if (guide) {
    var guideName = isAr ? guide.nameAr : guide.nameEn;
    text += '\n\n👨‍✈️ <b>' + guideLbl + '</b>\n' +
      '━━━━━━━━━━━━━━\n' +
      '👤 ' + guideName + '\n' +
      '📞 ' + guidePhoneLbl + ': <code>' + guide.phone + '</code>';
  }

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
    nl: ['Noodnummers', 'Noodgevallen', 'Ministerie van Hadj', 'Rode Halve Maan']
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
    nl: ['Neem contact op met Ikram Al-Dayf', 'Telefoon', 'WhatsApp']
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
// 🔄 تحديث البيانات
// ============================================
function handleRefreshData_(chatId, session) {
  var lang = (session && session.language) || 'ar';
  var passport = session.passport;

  if (passport) {
    clearPilgrimCache_(passport);
    var pilgrim = findPilgrimByPassport_(passport);
    if (pilgrim) {
      var pkgId = String(pilgrim.rowData[1] || '');
      clearTransportCache_(pkgId);
      clearHotelMapCache_(pilgrim.rowData);
    }
  }

  sendMessage_(chatId, T_('data_refreshed', lang));
  sendMainMenu_(chatId, lang);
}
