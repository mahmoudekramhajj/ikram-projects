// ============================================
// معالجات القائمة — طيران، فندق، باقة، مواصلات، طوارئ، تواصل، تحديث
// ============================================

// ============================================
// ✈️ رحلتي
// ============================================
function handleMyFlight_(chatId, session) {
  var lang = session.language || 'ar';
  var isAr = (lang === 'ar');

  // جلب بيانات الطيران من المصدر الصحيح (الطيران / B2C)
  var data = getPilgrimFlightData_(session.passport);
  if (!data || (data.arrival.legs.length === 0 && data.return.legs.length === 0)) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var airlineName = isAr ? (data.airline.nameAr || data.airline.code) : (data.airline.nameEn || data.airline.code);
  var legLbl = T_('lbl_leg', lang);
  var transitLbl = T_('lbl_transit', lang);

  var text = '';

  // ============ رحلة الوصول ============
  text += isAr
    ? '✈️ <b>رحلة الوصول</b>\n━━━━━━━━━━━━━━\n🏢 الناقل: ' + airlineName + '\n🎫 النوع: ' + data.flightType + '\n'
    : '✈️ <b>Arrival Flight</b>\n━━━━━━━━━━━━━━\n🏢 Airline: ' + airlineName + '\n🎫 Type: ' + data.flightType + '\n';

  text += buildLegsSection_(data.arrival.legs, legLbl, transitLbl, isAr);

  // ============ رحلة العودة ============
  text += '\n';
  text += isAr
    ? '✈️ <b>رحلة العودة</b>\n━━━━━━━━━━━━━━\n🏢 الناقل: ' + airlineName + '\n'
    : '✈️ <b>Return Flight</b>\n━━━━━━━━━━━━━━\n🏢 Airline: ' + airlineName + '\n';

  text += buildLegsSection_(data.return.legs, legLbl, transitLbl, isAr);

  // ============ أزرار التذكرة والبلاغ ============
  var ticketInfo = getVisaAndTicket_(session.passport);
  var ticketLink = (ticketInfo && ticketInfo.ticketLink) ? ticketInfo.ticketLink : '';

  var flightButtons = [];
  if (ticketLink) {
    flightButtons.push([{ text: T_('btn_download_ticket', lang), url: ticketLink }]);
  }
  flightButtons.push([{ text: T_('btn_report_error', lang), callback_data: 'report_error' }]);
  flightButtons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);

  sendMessage_(chatId, text, { inline_keyboard: flightButtons });
}

// ============================================
// بناء نص مراحل الرحلة (مباشرة أو ترانزيت)
// ============================================
function buildLegsSection_(legs, legLbl, transitLbl, isAr) {
  var s = '';
  if (!legs || legs.length === 0) {
    s += isAr ? '⚠️ لم تُحدّد بيانات الرحلة بعد.\n' : '⚠️ Flight data not yet available.\n';
    return s;
  }

  var fmtDate = isAr ? formatArabicDate_ : formatDate_;
  var fmtTime = isAr ? formatArabicTime_ : formatTime_;
  var numFmt  = isAr ? toArabicDigits_ : function(x) { return x; };
  var lang    = isAr ? 'ar' : 'en';
  var fromLbl = isAr ? 'من'  : 'From';
  var toLbl   = isAr ? 'إلى' : 'To';

  if (legs.length === 1) {
    // رحلة مباشرة — عرض مبسّط
    var leg = legs[0];
    s += '🛫 ' + (isAr ? 'الرحلة' : 'Flight') + ': <b>' + leg.flightNo + '</b>\n';
    s += '📍 ' + fromLbl + ': ' + getAirportDisplay_(leg.from, lang) + '\n';
    s += '🎯 ' + toLbl   + ': ' + getAirportDisplay_(leg.to,   lang) + '\n';
    s += '📅 ' + fmtDate(leg.takeoffDate) + '  🕐 ' + fmtTime(leg.takeoffTime) + '\n';
    s += '📅 ' + fmtDate(leg.landDate)    + '  🕐 ' + fmtTime(leg.landTime)    + '\n';
    return s;
  }

  // رحلة ترانزيت — عرض كل المراحل
  for (var i = 0; i < legs.length; i++) {
    var lg = legs[i];
    if (i > 0) {
      s += '\n✋ ' + transitLbl + '\n\n';
    }
    s += '🛫 ' + legLbl + ' ' + numFmt(i + 1) + ': <b>' + lg.flightNo + '</b>\n';
    s += '📍 ' + fromLbl + ': ' + getAirportDisplay_(lg.from, lang) + '\n';
    s += '🎯 ' + toLbl   + ': ' + getAirportDisplay_(lg.to,   lang) + '\n';
    s += '📅 ' + fmtDate(lg.takeoffDate) + '  🕐 ' + fmtTime(lg.takeoffTime) + '\n';
    s += '📅 ' + fmtDate(lg.landDate)    + '  🕐 ' + fmtTime(lg.landTime)    + '\n';
  }
  return s;
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
  var fmtDate2 = isAr ? formatArabicDate_ : formatDate_;
  var numFmt2  = isAr ? toArabicDigits_ : function(x) { return x; };

  var packageId = String(r[1] || '').trim();

  // 1) جلب جميع الفنادق من شيت الباقات (مصدر الحقيقة الموحَّد)
  var packageHotels = getPackageHotels_(packageId);
  if (!packageHotels || !packageHotels.length) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  // 2) جلب تاريخ وصول الحاج (آخر leg) وتطبيق قواعد التعديل
  var arrival = getPilgrimArrival_(session.passport, lang);
  var arrivalDate = arrival ? arrival.date : null;
  var hotels = applyArrivalOverride_(packageHotels, arrivalDate);

  if (!hotels || !hotels.length) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  // 3) بناء النص ديناميكياً (1 إلى N فندق)
  var accLbl = T_('lbl_acc', lang);
  var hotelLbl = T_('lbl_hotel', lang);
  var inLbl = T_('lbl_checkin', lang);
  var outLbl = T_('lbl_checkout', lang);

  var text = '';
  for (var i = 0; i < hotels.length; i++) {
    var h = hotels[i];
    var cityLbl = getCityDisplayName_(h.city, lang);
    var hotelName = isAr ? (h.nameAr || h.nameEn || '-') : (h.nameEn || h.nameAr || '-');
    var inDate = fmtDate2(h.checkIn);
    var outDate = fmtDate2(h.checkOut);

    if (i > 0) text += '\n\n';
    text += '🏨 <b>' + accLbl + ' ' + numFmt2(i + 1) + ' — ' + cityLbl + '</b>\n';
    text += '━━━━━━━━━━━━━━\n';
    text += '🏢 ' + hotelLbl + ': <b>' + hotelName + '</b>\n';
    text += '📅 ' + inLbl + ': ' + inDate + '\n';
    text += '📅 ' + outLbl + ': ' + outDate;
  }

  // 4) أزرار خرائط الفنادق (دون تكرار)
  var mapButtons = [];
  var mapPfx = T_('lbl_map', lang);
  var seenMaps = {};

  for (var j = 0; j < hotels.length; j++) {
    var hh = hotels[j];
    var hKey = (hh.nameAr || hh.nameEn || '').trim();
    if (!hKey || seenMaps[hKey]) continue;
    seenMaps[hKey] = true;
    var hLabel = isAr ? (hh.nameAr || hh.nameEn || '-') : (hh.nameEn || hh.nameAr || '-');
    var mapLink = getHotelMapLink_(hh.nameAr) || getHotelMapLink_(hh.nameEn);
    if (mapLink) mapButtons.push([{ text: mapPfx + hLabel, url: mapLink }]);
  }

  mapButtons.push([{ text: T_('btn_report_error', lang), callback_data: 'report_error' }]);
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

  var text = '';
  if (isAr) {
    text = '📦 <b>بيانات الباقة</b>\n━━━━━━━━━━━━━━\n🔢 رقم الباقة: <b>' + packageId + '</b>\n📋 ' + pkgNameLbl + ': <b>' + pkgName + '</b>\n👥 رقم المجموعة: ' + groupNo + '\n🎫 نوع التذكرة: ' + flightType + '\n🌍 الجنسية: ' + nationality + '\n🏠 بلد الإقامة: ' + country + '\n\n⛺ <b>المخيم</b>\n━━━━━━━━━━━━━━\n📍 الموقع: <b>' + campName + '</b>';
  } else {
    text = '📦 <b>Package Details</b>\n━━━━━━━━━━━━━━\n🔢 Package ID: <b>' + packageId + '</b>\n📋 ' + pkgNameLbl + ': <b>' + pkgName + '</b>\n👥 Group: ' + groupNo + '\n🎫 Ticket Type: ' + flightType + '\n🌍 Nationality: ' + nationality + '\n🏠 Residence: ' + country + '\n\n⛺ <b>Mina Camp</b>\n━━━━━━━━━━━━━━\n📍 Location: <b>' + campName + '</b>';
  }

  text += '\n\n👨‍✈️ <b>' + T_('lbl_guide', lang) + '</b>\n━━━━━━━━━━━━━━\n🔄 ' + T_('lbl_under_review', lang);

  sendMessage_(chatId, text, {
    inline_keyboard: [
      [{ text: T_('btn_report_error', lang), callback_data: 'report_error' }],
      [{ text: T_('btn_back', lang), callback_data: 'show_menu' }]
    ]
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
  var fmtDate2 = isAr ? formatArabicDate_ : formatDate_;
  var packageId = String(r[1] || '').trim();

  // 1) جلب فنادق الباقة + تطبيق منطق الوصول
  var packageHotels = getPackageHotels_(packageId);
  if (!packageHotels || !packageHotels.length) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var arrival = getPilgrimArrival_(session.passport, lang);
  var departure = getPilgrimDeparture_(session.passport, lang);
  var arrivalDate = arrival ? arrival.date : null;
  var hotels = applyArrivalOverride_(packageHotels, arrivalDate);

  if (!hotels || !hotels.length) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  // 2) نوع النقل
  var transportType = getTransportType_(packageId);
  var transportIcon = (transportType === 'قطار') ? '🚆' : '🚌';
  var transportTypeAr = transportType;
  var transportTypeEn = transportType === 'قطار' ? 'Train' : transportType === 'حافلة' ? 'Bus' : transportType;
  var isTrain = (transportType === 'قطار');
  var campName = String(r[4] || '-');

  function nameOf_(h) {
    return isAr ? (h.nameAr || h.nameEn || '-') : (h.nameEn || h.nameAr || '-');
  }
  function cityOf_(h) {
    return getCityDisplayName_(h.city, lang);
  }
  function cityKey_(h) {
    return String(h.city || '').toLowerCase().trim();
  }

  var step = 1;
  var text = '';

  if (isAr) {
    text = '🚌 <b>خطة التنقل</b>\n━━━━━━━━━━━━━━\n\n';

    // ١) الوصول → أول فندق
    var first = hotels[0];
    var arrAirport = arrival ? arrival.airportAr : '-';
    var arrFlight = arrival ? arrival.flightNo : '-';
    var arrDate = arrival ? fmtDate2(arrival.date) : '-';
    text += step + '️⃣ <b>الوصول — استقبال المطار</b>\n';
    text += '✈️ رحلة <b>' + arrFlight + '</b>\n';
    text += '📍 الوصول إلى ' + arrAirport + '\n';
    text += '📅 ' + arrDate + '\n';
    text += '🚌 نقل إلى فندق <b>' + nameOf_(first) + '</b> في ' + cityOf_(first) + '\n\n';
    step++;

    // ٢..N) التنقلات بين الفنادق
    for (var i = 1; i < hotels.length; i++) {
      var prev = hotels[i - 1];
      var curr = hotels[i];
      var sameCity = cityKey_(prev) === cityKey_(curr);
      var inDate = fmtDate2(curr.checkIn);

      if (sameCity) {
        text += step + '️⃣ <b>تحويل داخل ' + cityOf_(curr) + '</b>\n';
        text += '🚌 الانتقال من فندق <b>' + nameOf_(prev) + '</b> إلى فندق <b>' + nameOf_(curr) + '</b>\n';
        text += '📅 ' + inDate + '\n\n';
      } else {
        text += step + '️⃣ <b>التنقل بين المدن (' + transportTypeAr + ')</b>\n';
        text += transportIcon + ' من ' + cityOf_(prev) + ' إلى ' + cityOf_(curr) + '\n';
        text += '🏨 من فندق <b>' + nameOf_(prev) + '</b> إلى فندق <b>' + nameOf_(curr) + '</b>\n';
        text += '📅 ' + inDate + '\n\n';
      }
      step++;
    }

    // المخيم
    text += step + '️⃣ <b>⛺ المخيم — منى</b>\n📍 <b>' + campName + '</b>\n\n';
    step++;

    // المغادرة
    var last = hotels[hotels.length - 1];
    var depAirport = departure ? departure.airportAr : '-';
    var depFlight = departure ? departure.flightNo : '-';
    var depDate = departure ? fmtDate2(departure.date) : '-';
    text += step + '️⃣ <b>المغادرة — توديع المطار</b>\n';
    text += '✈️ رحلة <b>' + depFlight + '</b>\n';
    text += '🚌 نقل من فندق <b>' + nameOf_(last) + '</b> في ' + cityOf_(last) + '\n';
    text += '📍 إلى ' + depAirport + '\n';
    text += '📅 ' + depDate;

  } else {
    text = '🚌 <b>Transport Plan</b>\n━━━━━━━━━━━━━━\n\n';

    var first2 = hotels[0];
    var arrAirport2 = arrival ? arrival.airportEn : '-';
    var arrFlight2 = arrival ? arrival.flightNo : '-';
    var arrDate2 = arrival ? fmtDate2(arrival.date) : '-';
    text += step + '️⃣ <b>Arrival — Airport Reception</b>\n';
    text += '✈️ Flight <b>' + arrFlight2 + '</b>\n';
    text += '📍 Arriving at ' + arrAirport2 + '\n';
    text += '📅 ' + arrDate2 + '\n';
    text += '🚌 Transfer to <b>' + nameOf_(first2) + '</b> hotel in ' + cityOf_(first2) + '\n\n';
    step++;

    for (var j = 1; j < hotels.length; j++) {
      var prev2 = hotels[j - 1];
      var curr2 = hotels[j];
      var sameCity2 = cityKey_(prev2) === cityKey_(curr2);
      var inDate2 = fmtDate2(curr2.checkIn);

      if (sameCity2) {
        text += step + '️⃣ <b>Hotel Shift in ' + cityOf_(curr2) + '</b>\n';
        text += '🚌 Transfer from <b>' + nameOf_(prev2) + '</b> to <b>' + nameOf_(curr2) + '</b>\n';
        text += '📅 ' + inDate2 + '\n\n';
      } else {
        text += step + '️⃣ <b>Intercity Transfer (' + transportTypeEn + ')</b>\n';
        text += transportIcon + ' From ' + cityOf_(prev2) + ' to ' + cityOf_(curr2) + '\n';
        text += '🏨 From <b>' + nameOf_(prev2) + '</b> to <b>' + nameOf_(curr2) + '</b>\n';
        text += '📅 ' + inDate2 + '\n\n';
      }
      step++;
    }

    text += step + '️⃣ <b>⛺ Mina Camp</b>\n📍 <b>' + campName + '</b>\n\n';
    step++;

    var last2 = hotels[hotels.length - 1];
    var depAirport2 = departure ? departure.airportEn : '-';
    var depFlight2 = departure ? departure.flightNo : '-';
    var depDate2 = departure ? fmtDate2(departure.date) : '-';
    text += step + '️⃣ <b>Departure — Airport Farewell</b>\n';
    text += '✈️ Flight <b>' + depFlight2 + '</b>\n';
    text += '🚌 Transfer from <b>' + nameOf_(last2) + '</b> hotel in ' + cityOf_(last2) + '\n';
    text += '📍 To ' + depAirport2 + '\n';
    text += '📅 ' + depDate2;
  }

  var transportButtons = [];
  if (isTrain) {
    transportButtons.push([{ text: T_('btn_train_ticket', lang), callback_data: 'train_ticket' }]);
  }
  transportButtons.push([{ text: T_('btn_report_error', lang), callback_data: 'report_error' }]);
  transportButtons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);

  sendMessage_(chatId, text, { inline_keyboard: transportButtons });
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

  var text = '🚨 <b>' + e[0] + '</b>\n━━━━━━━━━━━━━━\n\n🔴 <b>' + e[1] + '</b>\n📞 <code>911</code>\n\n🕋 <b>' + e[2] + '</b>\n📞 <code>1966</code>\n\n🏥 <b>' + e[3] + '</b>\n📞 <code>997</code>';

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

  var text = '📞 <b>' + c[0] + '</b>\n━━━━━━━━━━━━━━\n\n📱 <b>' + c[1] + '</b>\n☎️ <code>8001111061</code>\n\n💬 <b>' + c[2] + '</b>\n📲 <code>+966125111940</code>';

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
  var arriveDate = normDate_(r[20]);
  var today = getTodayString_();
  var yesterday = getDateOffset_(today, -1);
  var tomorrow = getDateOffset_(today, 1);

  if (arriveDate !== today && arriveDate !== yesterday && arriveDate !== tomorrow) {
    sendMessage_(chatId, T_('arrival_not_today', lang, { date: formatDate_(r[20]) }), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
    return;
  }

  var reception = getReceptionStatus_(session.passport);
  if (reception && reception.status === 'تم') {
    sendMessage_(chatId, T_('arrival_already', lang, { time: reception.time }), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
    return;
  }

  var source = 'Bot:' + String(r[7] || '').substring(0, 30);
  var success = confirmArrivalInSheet_(session.passport, source);

  if (!success) {
    sendMessage_(chatId, '❌ حدث خطأ. حاول مرة أخرى.');
    return;
  }

  clearPilgrimCache_(session.passport);

  sendMessage_(chatId, T_('arrival_confirmed', lang), {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });

  try {
    var arriveCity = String(r[19] || '');
    var airlineEn = String(r[17] || '');
    var groupChatId = getOpsGroupChatId_(arriveCity, airlineEn);

    var name = String(r[7] || '-');
    var passport = String(r[8] || '-');
    var flight = String(r[24] || '-');
    var nationality = String(r[12] || '-');
    var house1Raw = String(r[36] || "-");
    var house1 = getCityDisplayName_(house1Raw, lang);
    var hotel = "";
    if (isMadinahCity_(house1Raw)) {
      hotel = String(r[47] || '-');
    } else {
      hotel = String(r[43] || '-');
    }

    var opsMsg = '✅ <b>تأكيد وصول (بوت)</b>\n━━━━━━━━━━━━━━\n👤 ' + name + '\n🛂 <code>' + passport + '</code>\n🌍 ' + nationality + '\n✈️ ' + flight + ' → ' + arriveCity + '\n🏨 ' + hotel + ' (' + house1 + ')';

    sendMessage_(groupChatId, opsMsg);
  } catch (e) {
    Logger.log('Ops notification error: ' + e.message);
  }
}

// ============================================
// 📱 QR Code
// ============================================
function handleMyQR_(chatId, session) {
  var lang = session.language || 'ar';
  var pilgrim = findPilgrimByPassport_(session.passport);

  if (!pilgrim) {
    sendMessage_(chatId, T_('data_not_found', lang));
    return;
  }

  var r = pilgrim.rowData;
  var qrData = {
    n: String(r[7] || ''), p: String(r[8] || ''), pkg: String(r[1] || ''),
    grp: String(r[6] || ''), nat: String(r[12] || ''), f: String(r[24] || ''),
    arr: normDate_(r[20]) || '', city: String(r[19] || ''),
    h1: String(r[36] || ''), h2: String(r[39] || '')
  };

  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&format=png&data=' + encodeURIComponent(JSON.stringify(qrData));

  sendPhoto_(chatId, qrUrl, T_('qr_caption', lang), {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// 📢 الإعلانات
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
    if (m.sentAt) label += ' — ' + m.sentAt.substring(0, 10);
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

  var langCol = LANG_TO_COL[lang] || AM.MSG_EN;
  var text = String(msgRow[langCol] || '').trim();
  if (!text) text = String(msgRow[AM.MSG_EN] || msgRow[AM.MSG_AR] || '').trim();

  var priority = String(msgRow[AM.PRIORITY] || '').toLowerCase().trim();
  var header = priority === 'urgent' ? T_('ann_urgent', lang) + ' ' : '';
  var fullText = T_('ann_from_admin', lang) + header + text;

  sendMessage_(chatId, fullText, { inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'announcements' }]] });

  var imageUrl = String(msgRow[AM.IMAGE_URL] || '').trim();
  if (imageUrl) sendPhoto_(chatId, getDriveDirectUrl_(imageUrl), '');

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
    clearPilgrimCache_(passport);
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

  var text = isAr ? '🎫 <b>التذكرة والتأشيرة</b>\n━━━━━━━━━━━━━━\n\n' : '🎫 <b>Ticket & Visa</b>\n━━━━━━━━━━━━━━\n\n';
  text += '📋 <b>' + visaStatusLbl + '</b>\n' + visaStatus + '\n\n';
  if (ticketNo) {
    text += '🎟️ <b>' + ticketNoLbl + '</b>\n<code>' + ticketNo + '</code>';
  } else {
    text += '🎟️ ' + T_('lbl_ticket_not_ready', lang);
  }
  if (!visaLink) {
    text += '\n\n📋 ' + T_('lbl_visa_not_ready', lang);
  }

  var buttons = [];
  if (ticketLink) {
    buttons.push([{ text: T_('btn_download_ticket', lang), url: ticketLink }]);
  }
  buttons.push([{ text: T_('btn_report_error', lang), callback_data: 'report_error' }]);
  buttons.push([{ text: T_('btn_back', lang), callback_data: 'show_menu' }]);
  sendMessage_(chatId, text, { inline_keyboard: buttons });

  if (visaLink) {
    sendVisaDocument_(chatId, visaLink, T_('btn_view_visa', lang));
  }
}
