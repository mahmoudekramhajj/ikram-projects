// ============================================
// 📢 نظام الإشعارات الأوتوماتيكي الشامل
// 15 سيناريو — يعمل يومياً بـ Trigger
// ============================================

function sendDailyNotifications() {
  var sheet = getSessionSheet_();
  var sessions = sheet.getDataRange().getValues();

  for (var i = 1; i < sessions.length; i++) {
    var authStatus = String(sessions[i][4]);
    if (authStatus !== 'verified') continue;

    var chatId = String(sessions[i][0]);
    var passport = String(sessions[i][1]);
    var lang = String(sessions[i][3]) || 'ar';

    if (!passport) continue;

    try {
      var pilgrim = findPilgrimByPassport_(passport);
      if (!pilgrim) continue;

      var r = pilgrim.rowData;
      var today = getTodayString_();
      var tomorrow = getDateOffset_(today, 1);
      var in2Days = getDateOffset_(today, 2);

      var arriveDate = normDate_(r[20]);   // ArrivalArriveDate (U)
      var departDate = normDate_(r[32]);   // ReturnDepartureDate (AG)
      var house1End = normDate_(r[38]);    // FirstHouseEndDate (AM)
      var house2Start = normDate_(r[40]);  // LastHouseStartDate (AO)
      var house2End = normDate_(r[41]);    // LastHouseEndDate (AP)

      var packageId = String(r[1] || '');
      var guide = getTourGuide_(packageId);

      // ━━━━━━━ المرحلة 1: الوصول ━━━━━━━

      // #1 — تذكير قبل الوصول بيومين
      if (arriveDate === in2Days) {
        sendNotif_arrival_reminder_(chatId, lang, r, guide);
      }

      // #2 — تذكير قبل الوصول بيوم
      if (arriveDate === tomorrow) {
        sendNotif_arrival_tomorrow_(chatId, lang, r, guide);
      }

      // #3 — يوم الوصول — مرحباً + معلومات الفندق
      if (arriveDate === today) {
        sendNotif_welcome_(chatId, lang, r, guide);
        // طلب رقم الهاتف السعودي إذا لم يُسجّل
        var pdata = getPilgrimData_(passport);
        if (!pdata || !pdata.saudiPhone) {
          sendNotif_ask_phone_(chatId, lang);
        }
      }

      // ━━━━━━━ المرحلة 2: الإقامة ━━━━━━━

      // #4 — اليوم الأول بالفندق — طلب رقم الغرفة
      var house1Start = normDate_(r[37]); // FirstHouseStartDate (AL)
      if (house1Start === today) {
        var pdata1 = getPilgrimData_(passport);
        if (!pdata1 || !pdata1.room1) {
          sendNotif_ask_room_(chatId, lang, 1);
        }
      }

      // #5 — أرقام الطوارئ + المرشد (يوم بعد الوصول)
      var dayAfterArrival = getDateOffset_(arriveDate, 1);
      if (dayAfterArrival === today) {
        sendNotif_emergency_info_(chatId, lang, guide);
      }

      // ━━━━━━━ المرحلة 3: التنقل بين المدن ━━━━━━━

      // #6 — تذكير قبل تغيير الفندق بيوم
      if (house1End && house1End === tomorrow) {
        sendNotif_hotel_change_tomorrow_(chatId, lang, r);
      }

      // #7 — يوم تغيير الفندق
      if (house2Start && house2Start === today) {
        sendNotif_hotel_change_today_(chatId, lang, r);
      }

      // #8 — وصول الفندق الجديد — معلوماته + طلب رقم الغرفة
      if (house2Start === today) {
        sendNotif_new_hotel_info_(chatId, lang, r);
        var pdata2 = getPilgrimData_(passport);
        if (!pdata2 || !pdata2.room2) {
          sendNotif_ask_room_(chatId, lang, 2);
        }
      }

      // ━━━━━━━ المرحلة 4: المشاعر (مؤجلة — جاهزة للتفعيل) ━━━━━━━
      // سيناريوهات 9-11 تحتاج بيانات إضافية
      // عند إضافتها: منى، عرفة، مزدلفة، أيام التشريق

      // ━━━━━━━ المرحلة 5: المغادرة ━━━━━━━

      // #12 — تذكير قبل المغادرة بيومين
      if (departDate === in2Days) {
        sendNotif_departure_2days_(chatId, lang, r);
      }

      // #13 — تذكير قبل المغادرة بيوم
      if (departDate === tomorrow) {
        sendNotif_departure_tomorrow_(chatId, lang, r);
      }

      // #14 — يوم المغادرة
      if (departDate === today) {
        sendNotif_departure_today_(chatId, lang, r);
      }

      // #15 — طلب صورة الجواز (3 أيام بعد الوصول إن لم تُرسل)
      var day3AfterArrival = getDateOffset_(arriveDate, 3);
      if (day3AfterArrival === today) {
        var pdataPhoto = getPilgrimData_(passport);
        if (!pdataPhoto || !pdataPhoto.passportPhoto) {
          sendNotif_ask_passport_photo_(chatId, lang);
        }
      }

    } catch (err) {
      Logger.log('Notification error for ' + chatId + ': ' + err.message);
    }
  }
}

// ============================================
// === رسائل الإشعارات ===
// ============================================

// #1 — تذكير قبل الوصول بيومين
function sendNotif_arrival_reminder_(chatId, lang, r, guide) {
  var flight = String(r[24] || '-');
  var from = String(r[21] || '-');
  var to = String(r[19] || '-');
  var date = (lang === 'ar') ? formatArabicDate_(r[20]) : formatDate_(r[20]);
  var time = (lang === 'ar') ? formatArabicTime_(r[23]) : formatTime_(r[23]);

  var guideLine = '';
  if (guide) {
    var gName = (lang === 'ar') ? guide.nameAr : guide.nameEn;
    guideLine = '\n\n👤 ' + T_('lbl_guide', lang) + ': ' + gName + '\n📞 ' + guide.phone;
  }

  var msgs = {
    ar: '📢 <b>تذكير — رحلتك بعد يومين!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ الرحلة: <b>' + flight + '</b>\n🛫 من: ' + from + '\n🛬 إلى: ' + to +
      '\n📅 التاريخ: ' + date + '\n🕐 الإقلاع: ' + time + guideLine +
      '\n\n🤲 نسأل الله لك رحلة آمنة',
    en: '📢 <b>Reminder — Your flight is in 2 days!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flight: <b>' + flight + '</b>\n🛫 From: ' + from + '\n🛬 To: ' + to +
      '\n📅 Date: ' + date + '\n🕐 Departure: ' + time + guideLine +
      '\n\n🤲 Wishing you a safe journey',
    fr: '📢 <b>Rappel — Votre vol est dans 2 jours!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vol: <b>' + flight + '</b>\n🛫 De: ' + from + '\n🛬 À: ' + to +
      '\n📅 Date: ' + date + '\n🕐 Départ: ' + time + guideLine +
      '\n\n🤲 Bon voyage',
    de: '📢 <b>Erinnerung — Ihr Flug ist in 2 Tagen!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flug: <b>' + flight + '</b>\n🛫 Von: ' + from + '\n🛬 Nach: ' + to +
      '\n📅 Datum: ' + date + '\n🕐 Abflug: ' + time + guideLine +
      '\n\n🤲 Gute Reise',
    it: '📢 <b>Promemoria — Il tuo volo è tra 2 giorni!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Volo: <b>' + flight + '</b>\n🛫 Da: ' + from + '\n🛬 A: ' + to +
      '\n📅 Data: ' + date + '\n🕐 Partenza: ' + time + guideLine +
      '\n\n🤲 Buon viaggio',
    es: '📢 <b>Recordatorio — ¡Tu vuelo es en 2 días!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vuelo: <b>' + flight + '</b>\n🛫 Desde: ' + from + '\n🛬 A: ' + to +
      '\n📅 Fecha: ' + date + '\n🕐 Salida: ' + time + guideLine +
      '\n\n🤲 Buen viaje'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #2 — تذكير قبل الوصول بيوم
function sendNotif_arrival_tomorrow_(chatId, lang, r, guide) {
  var flight = String(r[24] || '-');
  var from = String(r[21] || '-');
  var to = String(r[19] || '-');
  var time = (lang === 'ar') ? formatArabicTime_(r[23]) : formatTime_(r[23]);
  var house1 = String(r[36] || '-');

  var hotel = '';
  if (house1.toLowerCase().indexOf('madi') !== -1) {
    hotel = (lang === 'ar') ? String(r[46] || '-') : String(r[47] || '-');
  } else {
    hotel = (lang === 'ar') ? String(r[42] || '-') : String(r[43] || '-');
  }

  var guideLine = '';
  if (guide) {
    var gName = (lang === 'ar') ? guide.nameAr : guide.nameEn;
    guideLine = '\n\n👤 ' + T_('lbl_guide', lang) + ': <b>' + gName + '</b>\n📞 <code>' + guide.phone + '</code>';
  }

  var msgs = {
    ar: '📢 <b>غداً موعد رحلتك!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ الرحلة: <b>' + flight + '</b>\n🛫 من: ' + from + '\n🛬 إلى: ' + to +
      '\n🕐 الإقلاع: ' + time + '\n\n🏨 فندقك: <b>' + hotel + '</b>' + guideLine +
      '\n\n📦 جهّز أمتعتك وثائقك\n🤲 في حفظ الله',
    en: '📢 <b>Your flight is tomorrow!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flight: <b>' + flight + '</b>\n🛫 From: ' + from + '\n🛬 To: ' + to +
      '\n🕐 Departure: ' + time + '\n\n🏨 Your hotel: <b>' + hotel + '</b>' + guideLine +
      '\n\n📦 Prepare your luggage and documents\n🤲 Safe travels',
    fr: '📢 <b>Votre vol est demain!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vol: <b>' + flight + '</b>\n🛫 De: ' + from + '\n🛬 À: ' + to +
      '\n🕐 Départ: ' + time + '\n\n🏨 Votre hôtel: <b>' + hotel + '</b>' + guideLine +
      '\n\n📦 Préparez vos bagages\n🤲 Bon voyage',
    de: '📢 <b>Ihr Flug ist morgen!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flug: <b>' + flight + '</b>\n🛫 Von: ' + from + '\n🛬 Nach: ' + to +
      '\n🕐 Abflug: ' + time + '\n\n🏨 Ihr Hotel: <b>' + hotel + '</b>' + guideLine +
      '\n\n📦 Bereiten Sie Ihr Gepäck vor\n🤲 Gute Reise',
    it: '📢 <b>Il tuo volo è domani!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Volo: <b>' + flight + '</b>\n🛫 Da: ' + from + '\n🛬 A: ' + to +
      '\n🕐 Partenza: ' + time + '\n\n🏨 Il tuo hotel: <b>' + hotel + '</b>' + guideLine +
      '\n\n📦 Prepara i bagagli e i documenti\n🤲 Buon viaggio',
    es: '📢 <b>¡Tu vuelo es mañana!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vuelo: <b>' + flight + '</b>\n🛫 Desde: ' + from + '\n🛬 A: ' + to +
      '\n🕐 Salida: ' + time + '\n\n🏨 Tu hotel: <b>' + hotel + '</b>' + guideLine +
      '\n\n📦 Prepara tu equipaje y documentos\n🤲 Buen viaje'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #3 — يوم الوصول
function sendNotif_welcome_(chatId, lang, r, guide) {
  var house1 = String(r[36] || '');
  var hotel = '';
  if (house1.toLowerCase().indexOf('madi') !== -1) {
    hotel = (lang === 'ar') ? String(r[46] || '-') : String(r[47] || '-');
  } else {
    hotel = (lang === 'ar') ? String(r[42] || '-') : String(r[43] || '-');
  }

  var mapLink = getHotelMapLink_(hotel);
  var mapLine = mapLink ? '\n📍 <a href="' + mapLink + '">' + T_('lbl_map', lang) + hotel + '</a>' : '';

  var guideLine = '';
  if (guide) {
    var gName = (lang === 'ar') ? guide.nameAr : guide.nameEn;
    guideLine = '\n\n👤 ' + T_('lbl_guide', lang) + ': <b>' + gName + '</b>\n📞 <code>' + guide.phone + '</code>';
  }

  var msgs = {
    ar: '🕋 <b>مرحباً بك في المملكة العربية السعودية!</b>\n━━━━━━━━━━━━━━\n' +
      '🏨 فندقك: <b>' + hotel + '</b>\n📍 السكن الأول: ' + house1 + mapLine + guideLine +
      '\n\n🚌 فريق الاستقبال بانتظارك في المطار\n🤲 حجاً مبروراً وذنباً مغفوراً',
    en: '🕋 <b>Welcome to Saudi Arabia!</b>\n━━━━━━━━━━━━━━\n' +
      '🏨 Your hotel: <b>' + hotel + '</b>\n📍 First stay: ' + house1 + mapLine + guideLine +
      '\n\n🚌 Our team is waiting at the airport\n🤲 May your Hajj be blessed',
    fr: '🕋 <b>Bienvenue en Arabie Saoudite!</b>\n━━━━━━━━━━━━━━\n' +
      '🏨 Votre hôtel: <b>' + hotel + '</b>\n📍 Premier séjour: ' + house1 + mapLine + guideLine +
      '\n\n🚌 Notre équipe vous attend à l\'aéroport\n🤲 Hajj béni',
    de: '🕋 <b>Willkommen in Saudi-Arabien!</b>\n━━━━━━━━━━━━━━\n' +
      '🏨 Ihr Hotel: <b>' + hotel + '</b>\n📍 Erster Aufenthalt: ' + house1 + mapLine + guideLine +
      '\n\n🚌 Unser Team erwartet Sie am Flughafen\n🤲 Möge Ihre Hadsch gesegnet sein',
    it: '🕋 <b>Benvenuto in Arabia Saudita!</b>\n━━━━━━━━━━━━━━\n' +
      '🏨 Il tuo hotel: <b>' + hotel + '</b>\n📍 Primo soggiorno: ' + house1 + mapLine + guideLine +
      '\n\n🚌 Il nostro team ti aspetta in aeroporto\n🤲 Che il tuo Hajj sia benedetto',
    es: '🕋 <b>¡Bienvenido a Arabia Saudita!</b>\n━━━━━━━━━━━━━━\n' +
      '🏨 Tu hotel: <b>' + hotel + '</b>\n📍 Primera estancia: ' + house1 + mapLine + guideLine +
      '\n\n🚌 Nuestro equipo te espera en el aeropuerto\n🤲 Que tu Hajj sea bendecido'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #4 — طلب رقم الهاتف السعودي
function sendNotif_ask_phone_(chatId, lang) {
  sendMessage_(chatId, T_('notif_ask_phone', lang), {
    inline_keyboard: [[{ text: T_('btn_add_phone', lang), callback_data: 'add_phone' }]]
  });
}

// #5 — أرقام الطوارئ + معلومات المرشد
function sendNotif_emergency_info_(chatId, lang, guide) {
  var guideLine = '';
  if (guide) {
    var gName = (lang === 'ar') ? guide.nameAr : guide.nameEn;
    guideLine = '\n\n👤 ' + T_('lbl_guide', lang) + ': <b>' + gName + '</b>\n📞 <code>' + guide.phone + '</code>';
  }

  var msgs = {
    ar: '📋 <b>معلومات مهمة — احتفظ بهذه الرسالة</b>\n━━━━━━━━━━━━━━\n\n' +
      '🔴 الطوارئ الموحد: <code>911</code>\n🕋 مساعدات الحجاج: <code>1966</code>\n🏥 الهلال الأحمر: <code>997</code>\n\n' +
      '📞 إكرام الضيف: <code>8001111061</code>' + guideLine,
    en: '📋 <b>Important Info — Save this message</b>\n━━━━━━━━━━━━━━\n\n' +
      '🔴 Emergency: <code>911</code>\n🕋 Pilgrim Support: <code>1966</code>\n🏥 Red Crescent: <code>997</code>\n\n' +
      '📞 Ikram Al-Dayf: <code>8001111061</code>' + guideLine,
    fr: '📋 <b>Infos importantes — Conservez ce message</b>\n━━━━━━━━━━━━━━\n\n' +
      '🔴 Urgences: <code>911</code>\n🕋 Aide aux pèlerins: <code>1966</code>\n🏥 Croissant-Rouge: <code>997</code>\n\n' +
      '📞 Ikram Al-Dayf: <code>8001111061</code>' + guideLine,
    de: '📋 <b>Wichtige Infos — Speichern Sie diese Nachricht</b>\n━━━━━━━━━━━━━━\n\n' +
      '🔴 Notruf: <code>911</code>\n🕋 Pilger-Hilfe: <code>1966</code>\n🏥 Roter Halbmond: <code>997</code>\n\n' +
      '📞 Ikram Al-Dayf: <code>8001111061</code>' + guideLine,
    it: '📋 <b>Informazioni importanti — Salva questo messaggio</b>\n━━━━━━━━━━━━━━\n\n' +
      '🔴 Emergenza: <code>911</code>\n🕋 Assistenza pellegrini: <code>1966</code>\n🏥 Mezzaluna Rossa: <code>997</code>\n\n' +
      '📞 Ikram Al-Dayf: <code>8001111061</code>' + guideLine,
    es: '📋 <b>Información importante — Guarda este mensaje</b>\n━━━━━━━━━━━━━━\n\n' +
      '🔴 Emergencias: <code>911</code>\n🕋 Ayuda al peregrino: <code>1966</code>\n🏥 Media Luna Roja: <code>997</code>\n\n' +
      '📞 Ikram Al-Dayf: <code>8001111061</code>' + guideLine
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #6 — تذكير قبل تغيير الفندق بيوم
function sendNotif_hotel_change_tomorrow_(chatId, lang, r) {
  var house1 = String(r[36] || '-');
  var house2 = String(r[39] || '-');

  var hotel2 = '';
  if (house2.toLowerCase().indexOf('madi') !== -1) {
    hotel2 = (lang === 'ar') ? String(r[46] || '-') : String(r[47] || '-');
  } else if (house2.toLowerCase().indexOf('shift') !== -1) {
    hotel2 = (lang === 'ar') ? String(r[44] || '-') : String(r[45] || '-');
  } else {
    hotel2 = (lang === 'ar') ? String(r[42] || '-') : String(r[43] || '-');
  }

  var packageId = String(r[1] || '');
  var transportType = getTransportType_(packageId);
  var transportIcon = (transportType === 'قطار') ? '🚆' : '🚌';
  var transportEn = transportType === 'قطار' ? 'Train' : 'Bus';

  var msgs = {
    ar: '🔔 <b>تنبيه — غداً موعد انتقالك!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 من: ' + house1 + '\n📍 إلى: ' + house2 + '\n🏨 الفندق الجديد: <b>' + hotel2 + '</b>\n' +
      transportIcon + ' وسيلة النقل: <b>' + transportType + '</b>\n\n' +
      '📦 جهّز أمتعتك وسلّم الغرفة قبل الساعة 12:00 ظهراً',
    en: '🔔 <b>Alert — Hotel transfer tomorrow!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 From: ' + house1 + '\n📍 To: ' + house2 + '\n🏨 New hotel: <b>' + hotel2 + '</b>\n' +
      transportIcon + ' Transport: <b>' + transportEn + '</b>\n\n' +
      '📦 Pack your bags and check out before 12:00 PM',
    fr: '🔔 <b>Alerte — Transfert demain!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 De: ' + house1 + '\n📍 À: ' + house2 + '\n🏨 Nouvel hôtel: <b>' + hotel2 + '</b>\n' +
      transportIcon + ' Transport: <b>' + transportEn + '</b>\n\n' +
      '📦 Préparez vos bagages, checkout avant 12h00',
    de: '🔔 <b>Hinweis — Hotelwechsel morgen!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 Von: ' + house1 + '\n📍 Nach: ' + house2 + '\n🏨 Neues Hotel: <b>' + hotel2 + '</b>\n' +
      transportIcon + ' Transport: <b>' + transportEn + '</b>\n\n' +
      '📦 Packen Sie Ihr Gepäck, Checkout vor 12:00 Uhr',
    it: '🔔 <b>Avviso — Trasferimento hotel domani!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 Da: ' + house1 + '\n📍 A: ' + house2 + '\n🏨 Nuovo hotel: <b>' + hotel2 + '</b>\n' +
      transportIcon + ' Trasporto: <b>' + transportEn + '</b>\n\n' +
      '📦 Prepara i bagagli, checkout entro le 12:00',
    es: '🔔 <b>¡Aviso — Traslado de hotel mañana!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 Desde: ' + house1 + '\n📍 A: ' + house2 + '\n🏨 Nuevo hotel: <b>' + hotel2 + '</b>\n' +
      transportIcon + ' Transporte: <b>' + transportEn + '</b>\n\n' +
      '📦 Prepara tu equipaje, checkout antes de las 12:00'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #7 — يوم تغيير الفندق
function sendNotif_hotel_change_today_(chatId, lang, r) {
  var house1 = String(r[36] || '-');
  var house2 = String(r[39] || '-');

  var packageId = String(r[1] || '');
  var transportType = getTransportType_(packageId);
  var transportIcon = (transportType === 'قطار') ? '🚆' : '🚌';
  var transportEn = transportType === 'قطار' ? 'Train' : 'Bus';

  var msgs = {
    ar: '🚌 <b>اليوم موعد انتقالك!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 من: ' + house1 + '\n📍 إلى: ' + house2 + '\n' +
      transportIcon + ' الوسيلة: <b>' + transportType + '</b>\n\n' +
      '⏰ تأكد من تسليم مفتاح الغرفة\n📦 تأكد من أمتعتك كاملة',
    en: '🚌 <b>Transfer day!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 From: ' + house1 + '\n📍 To: ' + house2 + '\n' +
      transportIcon + ' Via: <b>' + transportEn + '</b>\n\n' +
      '⏰ Return your room key\n📦 Make sure you have all your luggage',
    fr: '🚌 <b>Jour de transfert!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 De: ' + house1 + '\n📍 À: ' + house2 + '\n' +
      transportIcon + ' Via: <b>' + transportEn + '</b>\n\n' +
      '⏰ Rendez la clé de chambre\n📦 Vérifiez vos bagages',
    de: '🚌 <b>Umzugstag!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 Von: ' + house1 + '\n📍 Nach: ' + house2 + '\n' +
      transportIcon + ' Via: <b>' + transportEn + '</b>\n\n' +
      '⏰ Geben Sie Ihren Zimmerschlüssel ab\n📦 Überprüfen Sie Ihr gesamtes Gepäck',
    it: '🚌 <b>Giorno del trasferimento!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 Da: ' + house1 + '\n📍 A: ' + house2 + '\n' +
      transportIcon + ' Via: <b>' + transportEn + '</b>\n\n' +
      '⏰ Restituisci la chiave della camera\n📦 Controlla tutti i tuoi bagagli',
    es: '🚌 <b>¡Día de traslado!</b>\n━━━━━━━━━━━━━━\n' +
      '📍 Desde: ' + house1 + '\n📍 A: ' + house2 + '\n' +
      transportIcon + ' Via: <b>' + transportEn + '</b>\n\n' +
      '⏰ Entrega la llave de tu habitación\n📦 Asegúrate de tener todo tu equipaje'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #8 — معلومات الفندق الجديد
function sendNotif_new_hotel_info_(chatId, lang, r) {
  var house2 = String(r[39] || '-');
  var hotel2 = '';
  if (house2.toLowerCase().indexOf('madi') !== -1) {
    hotel2 = (lang === 'ar') ? String(r[46] || '-') : String(r[47] || '-');
  } else if (house2.toLowerCase().indexOf('shift') !== -1) {
    hotel2 = (lang === 'ar') ? String(r[44] || '-') : String(r[45] || '-');
  } else {
    hotel2 = (lang === 'ar') ? String(r[42] || '-') : String(r[43] || '-');
  }

  var checkIn = (lang === 'ar') ? formatArabicDate_(r[40]) : formatDate_(r[40]);
  var checkOut = (lang === 'ar') ? formatArabicDate_(r[41]) : formatDate_(r[41]);

  var mapLink = getHotelMapLink_(hotel2);
  var mapLine = mapLink ? '\n📍 <a href="' + mapLink + '">' + T_('lbl_map', lang) + hotel2 + '</a>' : '';

  var msgs = {
    ar: '🏨 <b>مرحباً بك في فندقك الجديد</b>\n━━━━━━━━━━━━━━\n' +
      '🏢 <b>' + hotel2 + '</b>\n📅 الدخول: ' + checkIn + '\n📅 الخروج: ' + checkOut + mapLine,
    en: '🏨 <b>Welcome to your new hotel</b>\n━━━━━━━━━━━━━━\n' +
      '🏢 <b>' + hotel2 + '</b>\n📅 Check-in: ' + checkIn + '\n📅 Check-out: ' + checkOut + mapLine,
    fr: '🏨 <b>Bienvenue dans votre nouvel hôtel</b>\n━━━━━━━━━━━━━━\n' +
      '🏢 <b>' + hotel2 + '</b>\n📅 Arrivée: ' + checkIn + '\n📅 Départ: ' + checkOut + mapLine,
    de: '🏨 <b>Willkommen in Ihrem neuen Hotel</b>\n━━━━━━━━━━━━━━\n' +
      '🏢 <b>' + hotel2 + '</b>\n📅 Check-in: ' + checkIn + '\n📅 Check-out: ' + checkOut + mapLine,
    it: '🏨 <b>Benvenuto nel tuo nuovo hotel</b>\n━━━━━━━━━━━━━━\n' +
      '🏢 <b>' + hotel2 + '</b>\n📅 Check-in: ' + checkIn + '\n📅 Check-out: ' + checkOut + mapLine,
    es: '🏨 <b>Bienvenido a tu nuevo hotel</b>\n━━━━━━━━━━━━━━\n' +
      '🏢 <b>' + hotel2 + '</b>\n📅 Check-in: ' + checkIn + '\n📅 Check-out: ' + checkOut + mapLine
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// طلب رقم الغرفة
function sendNotif_ask_room_(chatId, lang, hotelNum) {
  var msgs = {
    ar: '🏨 هل وصلت لغرفتك؟ أرسل لنا <b>رقم الغرفة</b> لنحفظه في ملفك:',
    en: '🏨 Have you reached your room? Send us your <b>room number</b>:',
    fr: '🏨 Êtes-vous dans votre chambre? Envoyez votre <b>numéro de chambre</b>:',
    de: '🏨 Sind Sie in Ihrem Zimmer? Senden Sie uns Ihre <b>Zimmernummer</b>:',
    it: '🏨 Sei arrivato nella tua stanza? Inviaci il <b>numero della camera</b>:',
    es: '🏨 ¿Has llegado a tu habitación? Envíanos el <b>número de habitación</b>:'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en'], {
    inline_keyboard: [[{ text: T_('btn_add_room', lang), callback_data: 'room_hotel_' + hotelNum }]]
  });
}

// #12 — تذكير قبل المغادرة بيومين
function sendNotif_departure_2days_(chatId, lang, r) {
  var flight = String(r[34] || '-');
  var date = (lang === 'ar') ? formatArabicDate_(r[32]) : formatDate_(r[32]);
  var time = (lang === 'ar') ? formatArabicTime_(r[33]) : formatTime_(r[33]);
  var from = String(r[31] || '-');
  var to = String(r[29] || '-');

  var msgs = {
    ar: '📢 <b>تذكير — رحلة العودة بعد يومين</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ الرحلة: <b>' + flight + '</b>\n📅 التاريخ: ' + date + '\n🕐 الموعد: ' + time +
      '\n🛫 من: ' + from + '\n🛬 إلى: ' + to +
      '\n\n📦 ابدأ بتجهيز أمتعتك',
    en: '📢 <b>Reminder — Return flight in 2 days</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flight: <b>' + flight + '</b>\n📅 Date: ' + date + '\n🕐 Time: ' + time +
      '\n🛫 From: ' + from + '\n🛬 To: ' + to +
      '\n\n📦 Start packing your luggage',
    fr: '📢 <b>Rappel — Vol retour dans 2 jours</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vol: <b>' + flight + '</b>\n📅 Date: ' + date + '\n🕐 Heure: ' + time +
      '\n🛫 De: ' + from + '\n🛬 À: ' + to +
      '\n\n📦 Commencez à préparer vos bagages',
    de: '📢 <b>Erinnerung — Rückflug in 2 Tagen</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flug: <b>' + flight + '</b>\n📅 Datum: ' + date + '\n🕐 Zeit: ' + time +
      '\n🛫 Von: ' + from + '\n🛬 Nach: ' + to +
      '\n\n📦 Beginnen Sie mit dem Packen',
    it: '📢 <b>Promemoria — Volo di ritorno tra 2 giorni</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Volo: <b>' + flight + '</b>\n📅 Data: ' + date + '\n🕐 Ora: ' + time +
      '\n🛫 Da: ' + from + '\n🛬 A: ' + to +
      '\n\n📦 Inizia a preparare i bagagli',
    es: '📢 <b>Recordatorio — Vuelo de regreso en 2 días</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vuelo: <b>' + flight + '</b>\n📅 Fecha: ' + date + '\n🕐 Hora: ' + time +
      '\n🛫 Desde: ' + from + '\n🛬 A: ' + to +
      '\n\n📦 Comienza a preparar tu equipaje'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #13 — تذكير قبل المغادرة بيوم
function sendNotif_departure_tomorrow_(chatId, lang, r) {
  var flight = String(r[34] || '-');
  var from = String(r[31] || '-');
  var to = String(r[29] || '-');
  var date = (lang === 'ar') ? formatArabicDate_(r[32]) : formatDate_(r[32]);
  var time = (lang === 'ar') ? formatArabicTime_(r[33]) : formatTime_(r[33]);

  var msgs = {
    ar: '📢 <b>غداً رحلة عودتك!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ الرحلة: <b>' + flight + '</b>\n🛫 من: ' + from + '\n🛬 إلى: ' + to +
      '\n📅 التاريخ: ' + date + '\n🕐 الإقلاع: ' + time +
      '\n\n🚌 سيتم نقلك من الفندق إلى المطار\n⏰ كن مستعداً في الوقت المحدد\n🤲 تقبّل الله حجك',
    en: '📢 <b>Your return flight is tomorrow!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flight: <b>' + flight + '</b>\n🛫 From: ' + from + '\n🛬 To: ' + to +
      '\n📅 Date: ' + date + '\n🕐 Departure: ' + time +
      '\n\n🚌 You will be transferred from hotel to airport\n⏰ Be ready on time\n🤲 May Allah accept your Hajj',
    fr: '📢 <b>Votre vol retour est demain!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vol: <b>' + flight + '</b>\n🛫 De: ' + from + '\n🛬 À: ' + to +
      '\n📅 Date: ' + date + '\n🕐 Départ: ' + time +
      '\n\n🚌 Transfert hôtel-aéroport assuré\n⏰ Soyez prêt à l\'heure\n🤲 Hajj accepté inchallah',
    de: '📢 <b>Ihr Rückflug ist morgen!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Flug: <b>' + flight + '</b>\n🛫 Von: ' + from + '\n🛬 Nach: ' + to +
      '\n📅 Datum: ' + date + '\n🕐 Abflug: ' + time +
      '\n\n🚌 Transfer Hotel-Flughafen\n⏰ Seien Sie rechtzeitig bereit\n🤲 Möge Allah Ihre Hadsch annehmen',
    it: '📢 <b>Il tuo volo di ritorno è domani!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Volo: <b>' + flight + '</b>\n🛫 Da: ' + from + '\n🛬 A: ' + to +
      '\n📅 Data: ' + date + '\n🕐 Partenza: ' + time +
      '\n\n🚌 Trasferimento hotel-aeroporto\n⏰ Sii pronto in orario\n🤲 Che Allah accetti il tuo Hajj',
    es: '📢 <b>¡Tu vuelo de regreso es mañana!</b>\n━━━━━━━━━━━━━━\n' +
      '✈️ Vuelo: <b>' + flight + '</b>\n🛫 Desde: ' + from + '\n🛬 A: ' + to +
      '\n📅 Fecha: ' + date + '\n🕐 Salida: ' + time +
      '\n\n🚌 Traslado hotel-aeropuerto\n⏰ Esté listo a tiempo\n🤲 Que Allah acepte tu Hajj'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #14 — يوم المغادرة
function sendNotif_departure_today_(chatId, lang, r) {
  var flight = String(r[34] || '-');
  var time = (lang === 'ar') ? formatArabicTime_(r[33]) : formatTime_(r[33]);
  var from = String(r[31] || '-');

  var msgs = {
    ar: '✈️ <b>اليوم موعد رحلة عودتك!</b>\n━━━━━━━━━━━━━━\n' +
      '🛫 الرحلة: <b>' + flight + '</b>\n🕐 الإقلاع: ' + time + '\n📍 المطار: ' + from +
      '\n\n⏰ تأكد من تسليم الغرفة\n📦 تأكد من جميع أمتعتك\n🤲 أعاده الله عليك بالخير',
    en: '✈️ <b>Your departure is today!</b>\n━━━━━━━━━━━━━━\n' +
      '🛫 Flight: <b>' + flight + '</b>\n🕐 Departure: ' + time + '\n📍 Airport: ' + from +
      '\n\n⏰ Check out of your room\n📦 Make sure you have everything\n🤲 May you return safely',
    fr: '✈️ <b>Votre départ est aujourd\'hui!</b>\n━━━━━━━━━━━━━━\n' +
      '🛫 Vol: <b>' + flight + '</b>\n🕐 Départ: ' + time + '\n📍 Aéroport: ' + from +
      '\n\n⏰ Checkout de la chambre\n📦 Vérifiez tous vos bagages\n🤲 Bon retour',
    de: '✈️ <b>Heute ist Ihr Abreisetag!</b>\n━━━━━━━━━━━━━━\n' +
      '🛫 Flug: <b>' + flight + '</b>\n🕐 Abflug: ' + time + '\n📍 Flughafen: ' + from +
      '\n\n⏰ Checken Sie aus Ihrem Zimmer aus\n📦 Überprüfen Sie Ihr gesamtes Gepäck\n🤲 Sichere Heimkehr',
    it: '✈️ <b>Oggi è il giorno della partenza!</b>\n━━━━━━━━━━━━━━\n' +
      '🛫 Volo: <b>' + flight + '</b>\n🕐 Partenza: ' + time + '\n📍 Aeroporto: ' + from +
      '\n\n⏰ Fai il check-out dalla camera\n📦 Controlla tutti i tuoi bagagli\n🤲 Buon rientro a casa',
    es: '✈️ <b>¡Hoy es tu día de partida!</b>\n━━━━━━━━━━━━━━\n' +
      '🛫 Vuelo: <b>' + flight + '</b>\n🕐 Salida: ' + time + '\n📍 Aeropuerto: ' + from +
      '\n\n⏰ Haz el checkout de tu habitación\n📦 Asegúrate de tener todo tu equipaje\n🤲 Que regreses con bien'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en']);
}

// #15 — طلب صورة الجواز
function sendNotif_ask_passport_photo_(chatId, lang) {
  var msgs = {
    ar: '📸 نرجو منك إرسال <b>صورة جواز السفر</b> لحفظها في ملفك.\n\nاضغط الزر أدناه ثم أرسل الصورة:',
    en: '📸 Please send a <b>photo of your passport</b> for our records.\n\nTap the button below then send the photo:',
    fr: '📸 Veuillez envoyer une <b>photo de votre passeport</b>.\n\nAppuyez sur le bouton ci-dessous:',
    de: '📸 Bitte senden Sie ein <b>Foto Ihres Reisepasses</b>.\n\nTippen Sie auf die Schaltfläche unten:',
    it: '📸 Per favore invia una <b>foto del tuo passaporto</b>.\n\nPremi il pulsante qui sotto:',
    es: '📸 Por favor envía una <b>foto de tu pasaporte</b>.\n\nPresiona el botón de abajo:'
  };
  sendMessage_(chatId, msgs[lang] || msgs['en'], {
    inline_keyboard: [[{ text: T_('btn_add_photo', lang), callback_data: 'add_passport_photo' }]]
  });
}

// ============================================
// إعداد Trigger يومي — يُشغَّل يدوياً مرة واحدة
// ============================================
function setupDailyNotificationTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyNotifications') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('sendDailyNotifications')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Riyadh')
    .create();

  Logger.log('Daily notification trigger created at 8:00 AM Riyadh time');
}
