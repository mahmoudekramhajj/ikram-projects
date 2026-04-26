// ============================================
// نقطة الدخول + توجيه الرسائل والأزرار
// ============================================

function doGet(e) {
  var params = e ? e.parameter : {};
  if (params.key) return handleClaudeAPI_(e);
  return ContentService.createTextOutput('IkramHajjBot is running! v18');
}

function doPost(e) {
  try {
    var update = JSON.parse(e.postData.contents);
    if (update.message) { handleMessage_(update.message); return; }
    if (update.callback_query) { handleCallback_(update.callback_query); return; }
  } catch (err) {
    Logger.log('Bot Error: ' + err.message + '\n' + err.stack);
  }
}

function handleMessage_(message) {
  var chatId = message.chat.id;
  var text = (message.text || '').trim();
  var firstName = message.from.first_name || '';

  if (text === '/start') { sendWelcome_(chatId, firstName); return; }

  if (text === '/broadcast' || text === '/stats') {
    if (isAdmin_(chatId)) { if (text === '/broadcast') startBroadcast_(chatId); else handleStats_(chatId); }
    else { sendMessage_(chatId, T_('admin_not_authorized', 'ar')); }
    return;
  }

  if (text === '/menu') {
    var session = getSession_(chatId);
    if (session && session.authStatus === 'verified') sendMainMenu_(chatId, session.language);
    else sendMessage_(chatId, T_('start_prompt', 'ar'));
    return;
  }

  // --- التحقق من رد المدير على بلاغ (قبل أي شيء آخر) ---
  if (text) {
    var pendingReport = getCache_('rpt_pending_' + chatId);
    if (pendingReport) {
      handleReportAdminInput_(chatId, text, pendingReport);
      return;
    }
  }

  // --- المشغّل في حالة كتابة رد لبلاغ TicketsMonitor ---
  if (text && isAdmin_(chatId)) {
    var pendingTm = getCache_('tm_pending_' + chatId);
    if (pendingTm) {
      handleTicketMonitorInput_(chatId, text, pendingTm);
      return;
    }
  }

  var session = getSession_(chatId);
  if (!session) { sendMessage_(chatId, T_('start_prompt', 'ar')); return; }

  if (session.authStatus === 'awaiting_passport') { handlePassportInput_(chatId, text, session); return; }

  if (session.authStatus === 'verified') {
    var inputState = session.inputState || '';

    if (message.photo && message.photo.length > 0) {
      if (inputState === 'awaiting_passport_photo') {
        handlePassportPhoto_(chatId, message.photo[message.photo.length - 1].file_id, session);
      } else {
        sendMessage_(chatId, T_('photo_not_received', session.language));
        sendMainMenu_(chatId, session.language);
      }
      return;
    }

    if (inputState === 'awaiting_phone') { handlePhoneInput_(chatId, text, session); return; }
    if (inputState.indexOf('awaiting_room') === 0) { handleRoomInput_(chatId, text, session); return; }
    if (inputState === 'awaiting_passport_photo') { sendMessage_(chatId, T_('photo_not_received', session.language)); return; }
    if (inputState === 'awaiting_report_correction') { handleReportCorrectionInput_(chatId, text, session); return; }
    if (inputState.indexOf('admin_') === 0) { handleAdminInput_(chatId, text, message, session); return; }

    sendMessage_(chatId, '📩 استلمنا سؤالك:\n«' + text + '»\n\nاستخدم القائمة للاستعلام عن بياناتك 👇');
    sendMainMenu_(chatId, session.language);
    updateLastActivity_(chatId);
    return;
  }

  sendMessage_(chatId, T_('start_prompt_new', 'ar'));
}

function handleCallback_(callback) {
  var chatId = callback.message.chat.id;
  var data = callback.data;

  answerCallback_(callback.id);

  if (data.indexOf('lang_') === 0) { handleLanguageSelection_(chatId, data.replace('lang_', '')); return; }

  // --- أزرار TicketsMonitor للمشغّلين (قبل التحقق من الجلسة) ---
  if (data.indexOf('tm_') === 0 && isAdmin_(chatId)) {
    handleTicketMonitorCallback_(chatId, data, callback);
    return;
  }

  // --- أزرار البلاغات الإدارية (قبل التحقق من الجلسة) ---
  if (data.indexOf('rpt_skip_') === 0) { handleReportSkip_(chatId, data); return; }
  if (data.indexOf('rpt_') === 0) { handleReportAction_(chatId, data, callback.from); return; }

  var session = getSession_(chatId);
  if (!session || session.authStatus !== 'verified') { sendMessage_(chatId, T_('start_prompt', 'ar')); return; }

  var lang = session.language || 'ar';

  var handlers = {
    'my_flight':          function() { handleMyFlight_(chatId, session); },
    'my_hotel':           function() { handleMyHotel_(chatId, session); },
    'my_package':         function() { handleMyPackage_(chatId, session); },
    'my_transport':       function() { handleMyTransport_(chatId, session); },
    'my_data':            function() { handleMyData_(chatId, session); },
    'emergency':          function() { handleEmergency_(chatId, session); },
    'contact_company':    function() { handleContactCompany_(chatId, session); },
    'general_faq':        function() { sendMessage_(chatId, T_('faq_wip', lang)); },
    'change_lang':        function() { sendLanguageButtons_(chatId); },
    'refresh_data':       function() { handleRefreshData_(chatId, session); },
    'switch_pilgrim':     function() { handleSwitchPilgrim_(chatId, session); },
    'show_menu':          function() { updateSession_(chatId, { inputState: '' }); sendMainMenu_(chatId, lang); },
    'add_phone':          function() { promptPhone_(chatId, session); },
    'add_room':           function() { promptRoomSelection_(chatId, session); },
    'room_hotel_1':       function() { promptRoom_(chatId, session, '1'); },
    'room_hotel_2':       function() { promptRoom_(chatId, session, '2'); },
    'room_hotel_3':       function() { promptRoom_(chatId, session, '3'); },
    'add_passport_photo': function() { promptPassportPhoto_(chatId, session); },
    'confirm_arrival':    function() { handleConfirmArrival_(chatId, session); },
    'my_qr':              function() { handleMyQR_(chatId, session); },
    'announcements':      function() { handleAnnouncements_(chatId, session); },
    'visa_ticket':        function() { handleVisaTicket_(chatId, session); },
    'report_error':       function() { handleReportError_(chatId, session); },
    'report_submit':      function() { handleReportSubmit_(chatId, session); },
    'report_cancel':      function() { handleReportCancel_(chatId, session); },
    'my_reports':         function() { handleMyReports_(chatId, session); },
    'train_ticket':       function() { sendMessage_(chatId, T_('train_ticket_unavailable', lang), { inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]] }); }
  };

  if (data.indexOf('ann_detail_') === 0) { handleAnnDetail_(chatId, session, data.replace('ann_detail_', '')); return; }
  if (data.indexOf('report_sec_') === 0) { handleReportSection_(chatId, session, data.replace('report_sec_', '')); return; }
  if (data.indexOf('report_issue_') === 0) {
    var parts = data.replace('report_issue_', '').split('_');
    handleReportIssueSelected_(chatId, session, parts[0], parts.slice(1).join('_'));
    return;
  }
  if (data.indexOf('admin_') === 0) { handleAdminCallback_(chatId, data, session); return; }

  if (handlers[data]) handlers[data]();

  updateBotActivity_(chatId);
}
