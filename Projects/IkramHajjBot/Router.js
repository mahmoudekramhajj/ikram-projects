// ============================================
// نقطة الدخول + توجيه الرسائل والأزرار
// ============================================

function doGet(e) {
  return ContentService.createTextOutput('IkramHajjBot is running! v16');
}

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
// معالجة الرسائل النصية + الصور
// ============================================
function handleMessage_(message) {
  var chatId = message.chat.id;
  var text = (message.text || '').trim();
  var firstName = message.from.first_name || '';

  // === أمر /start ===
  if (text === '/start') {
    sendWelcome_(chatId, firstName);
    return;
  }

  // === أمر /menu ===
  if (text === '/menu') {
    var session = getSession_(chatId);
    if (session && session.authStatus === 'verified') {
      sendMainMenu_(chatId, session.language);
    } else {
      sendMessage_(chatId, T_('start_prompt', 'ar'));
    }
    return;
  }

  var session = getSession_(chatId);

  if (!session) {
    sendMessage_(chatId, T_('start_prompt', 'ar'));
    return;
  }

  // === مرحلة إدخال الجواز ===
  if (session.authStatus === 'awaiting_passport') {
    handlePassportInput_(chatId, text, session);
    return;
  }

  // === مرحلة verified — فحص حالة الإدخال ===
  if (session.authStatus === 'verified') {
    var inputState = session.inputState || '';

    // --- صورة مستلمة ---
    if (message.photo && message.photo.length > 0) {
      if (inputState === 'awaiting_passport_photo') {
        var fileId = message.photo[message.photo.length - 1].file_id;
        handlePassportPhoto_(chatId, fileId, session);
      } else {
        sendMessage_(chatId, T_('photo_not_received', session.language));
        sendMainMenu_(chatId, session.language);
      }
      return;
    }

    // --- جمع رقم الهاتف السعودي ---
    if (inputState === 'awaiting_phone') {
      handlePhoneInput_(chatId, text, session);
      return;
    }

    // --- جمع رقم الغرفة ---
    if (inputState.indexOf('awaiting_room') === 0) {
      handleRoomInput_(chatId, text, session);
      return;
    }

    // --- انتظار صورة الجواز — أرسل نص بدلاً من صورة ---
    if (inputState === 'awaiting_passport_photo') {
      sendMessage_(chatId, T_('photo_not_received', session.language));
      return;
    }

    // --- رسالة نصية عادية ---
    sendMessage_(chatId, '📩 استلمنا سؤالك:\n«' + text + '»\n\nاستخدم القائمة للاستعلام عن بياناتك 👇');
    sendMainMenu_(chatId, session.language);
    updateLastActivity_(chatId);
    return;
  }

  sendMessage_(chatId, T_('start_prompt_new', 'ar'));
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
    sendMessage_(chatId, T_('start_prompt', 'ar'));
    return;
  }

  var handlers = {
    'my_flight':       function() { handleMyFlight_(chatId, session); },
    'my_hotel':        function() { handleMyHotel_(chatId, session); },
    'my_package':      function() { handleMyPackage_(chatId, session); },
    'my_transport':    function() { handleMyTransport_(chatId, session); },
    'my_data':         function() { handleMyData_(chatId, session); },
    'emergency':       function() { handleEmergency_(chatId, session); },
    'contact_company': function() { handleContactCompany_(chatId, session); },
    'general_faq':     function() { sendMessage_(chatId, T_('faq_wip', session.language)); },
    'change_lang':     function() { sendLanguageButtons_(chatId); },
    'refresh_data':    function() { handleRefreshData_(chatId, session); },
    'switch_pilgrim':  function() { handleSwitchPilgrim_(chatId, session); },
    'show_menu':       function() { updateSession_(chatId, { inputState: '' }); sendMainMenu_(chatId, session.language); },
    'add_phone':       function() { promptPhone_(chatId, session); },
    'add_room':        function() { promptRoomSelection_(chatId, session); },
    'room_hotel_1':    function() { promptRoom_(chatId, session, '1'); },
    'room_hotel_2':    function() { promptRoom_(chatId, session, '2'); },
    'room_hotel_3':    function() { promptRoom_(chatId, session, '3'); },
    'add_passport_photo': function() { promptPassportPhoto_(chatId, session); }
  };

  if (handlers[data]) {
    handlers[data]();
  }
}
