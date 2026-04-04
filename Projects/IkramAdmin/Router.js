// ============================================
// Router.js — نقطة الدخول الرئيسية (Webhook)
// ============================================

/**
 * doGet — يُستخدم لـ ClaudeAPI + اختبار الاتصال
 */
function doGet(e) {
  var params = e ? e.parameter : {};

  // ClaudeAPI
  if (params.action || params.key) {
    return handleClaudeAPI_(e);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    bot: 'IkramAdmin',
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost — Telegram Webhook
 */
function doPost(e) {
  try {
    var update = JSON.parse(e.postData.contents);

    if (update.callback_query) {
      handleCallback_(update.callback_query);
    } else if (update.message) {
      handleMessage_(update.message);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
  }

  return ContentService.createTextOutput('ok');
}

// ============================================
// توجيه الرسائل النصية
// ============================================
function handleMessage_(message) {
  var chatId = message.chat.id;
  var text = (message.text || '').trim();

  if (!text) return;

  var session = getSession_(chatId);
  var startTime = new Date();

  // --- أوامر مباشرة ---
  if (text === '/start') {
    sendWelcome_(chatId);
    logActivity_(chatId, 'start', '', '');
    return;
  }

  if (text === '/menu') {
    if (session.status === 'verified') {
      sendMainMenu_(chatId, session);
    } else {
      sendWelcome_(chatId);
    }
    return;
  }

  // --- حالة انتظار المصادقة ---
  if (session.status === 'awaiting_auth') {
    handleAuthInput_(chatId, text, session);
    logActivity_(chatId, 'auth', text, getDuration_(startTime));
    return;
  }

  // --- غير مصادق ---
  if (session.status !== 'verified') {
    sendWelcome_(chatId);
    return;
  }

  // --- حالات الإدخال الخاصة ---
  if (session.inputState === 'awaiting_search') {
    session.inputState = '';
    saveSession_(session);
    handlePilgrimSearch_(chatId, text, session);
    logActivity_(chatId, 'search', text, getDuration_(startTime));
    return;
  }

  if (session.inputState === 'awaiting_ai') {
    session.inputState = '';
    saveSession_(session);
    handleAIQuery_(chatId, text, session);
    logActivity_(chatId, 'ai_query', text, getDuration_(startTime));
    return;
  }

  // --- نص حر = AI تلقائي ---
  // أي نص لا يتطابق مع أوامر → يُعالج كسؤال AI
  handleAIQuery_(chatId, text, session);
  logActivity_(chatId, 'ai_auto', text, getDuration_(startTime));
}

// ============================================
// توجيه الأزرار (Callbacks)
// ============================================
function handleCallback_(callbackQuery) {
  var chatId = callbackQuery.message.chat.id;
  var data = callbackQuery.data;
  var callbackId = callbackQuery.id;

  answerCallback_(callbackId);

  var session = getSession_(chatId);
  var startTime = new Date();

  // --- اختيار اللغة ---
  if (data === 'lang_ar') {
    handleLanguageSelection_(chatId, 'ar');
    return;
  }
  if (data === 'lang_en') {
    handleLanguageSelection_(chatId, 'en');
    return;
  }

  // --- غير مصادق ---
  if (session.status !== 'verified') {
    sendWelcome_(chatId);
    return;
  }

  var lang = session.language || 'ar';

  // --- تنقل ---
  switch (data) {
    case 'back_main':
      sendMainMenu_(chatId, session);
      return;

    case 'settings':
      sendSettingsMenu_(chatId, session);
      return;

    case 'change_lang':
      sendMessage_(chatId, T_('welcome', lang), {
        inline_keyboard: [
          [{ text: '🇸🇦 العربية', callback_data: 'lang_ar' },
           { text: '🇬🇧 English', callback_data: 'lang_en' }]
        ]
      });
      return;

    case 'logout':
      handleLogout_(chatId, session);
      return;

    // --- الميزات الرئيسية ---
    case 'dash':
      handleDashboard_(chatId, session);
      logActivity_(chatId, 'dashboard', '', getDuration_(startTime));
      return;

    case 'search':
      session.inputState = 'awaiting_search';
      saveSession_(session);
      sendMessage_(chatId, T_('search_prompt', lang));
      return;

    case 'packages':
      handlePackageStatus_(chatId, session);
      logActivity_(chatId, 'packages', '', getDuration_(startTime));
      return;

    case 'flights':
      sendFlightsMenu_(chatId, session);
      return;

    case 'flights_today':
      handleFlightSchedule_(chatId, session, 'today');
      logActivity_(chatId, 'flights', 'today', getDuration_(startTime));
      return;

    case 'flights_week':
      handleFlightSchedule_(chatId, session, 'week');
      logActivity_(chatId, 'flights', 'week', getDuration_(startTime));
      return;

    case 'flights_all':
      handleFlightSchedule_(chatId, session, 'all');
      logActivity_(chatId, 'flights', 'all', getDuration_(startTime));
      return;

    case 'hotels':
      handleHotelOccupancy_(chatId, session);
      logActivity_(chatId, 'hotels', '', getDuration_(startTime));
      return;

    case 'camps':
      handleCampStatus_(chatId, session);
      logActivity_(chatId, 'camps', '', getDuration_(startTime));
      return;

    case 'sales':
      if (!hasPermission_(session, 'sales')) {
        sendMessage_(chatId, T_('no_permission', lang));
        return;
      }
      handleSalesReport_(chatId, session);
      logActivity_(chatId, 'sales', '', getDuration_(startTime));
      return;

    case 'export':
      sendExportMenu_(chatId, session);
      return;

    case 'ai':
      session.inputState = 'awaiting_ai';
      saveSession_(session);
      sendMessage_(chatId, T_('ai_prompt', lang));
      return;

    // --- الإدارة ---
    case 'broadcast':
      if (!hasPermission_(session, 'broadcast')) {
        sendMessage_(chatId, T_('no_permission', lang));
        return;
      }
      // TODO: Phase 4
      sendMessage_(chatId, '📢 ميزة البث قيد التطوير...', backButton_(lang));
      return;

    case 'admin_users':
      if (!hasPermission_(session, 'user_mgmt')) {
        sendMessage_(chatId, T_('no_permission', lang));
        return;
      }
      // TODO: Phase 4
      sendMessage_(chatId, '👥 إدارة المستخدمين قيد التطوير...', backButton_(lang));
      return;

    case 'bot_stats':
      // TODO: Phase 4
      sendMessage_(chatId, '📈 إحصائيات البوت قيد التطوير...', backButton_(lang));
      return;

    // --- تصدير ---
    case 'exp_packages':
    case 'exp_flights':
    case 'exp_pilgrims':
      // TODO: Phase 3
      sendMessage_(chatId, T_('export_generating', lang));
      sendMessage_(chatId, '📤 التصدير قيد التطوير — سيتوفر في المرحلة 3', backButton_(lang));
      return;

    // --- عرض تفاصيل حاج ---
    default:
      if (data.indexOf('pilgrim_') === 0) {
        var seq = data.replace('pilgrim_', '');
        handlePilgrimDetail_(chatId, seq, session);
        logActivity_(chatId, 'pilgrim_detail', seq, getDuration_(startTime));
        return;
      }
  }
}

// ============================================
// حساب المدة
// ============================================
function getDuration_(startTime) {
  return ((new Date() - startTime) / 1000).toFixed(1) + 's';
}

// ============================================
// Stub — سيتم تنفيذه في المرحلة 3
// ============================================

function handleAIQuery_(chatId, text, session) {
  var lang = session.language || 'ar';
  sendMessage_(chatId, T_('ai_thinking', lang));
  sendMessage_(chatId, '💬 ' + (lang === 'ar' ? 'سؤالك' : 'Your question') + ': <b>' + text + '</b>\n\n⏳ ' + (lang === 'ar' ? 'الذكاء الاصطناعي سيتوفر في المرحلة 3' : 'AI coming in Phase 3'), backButton_(lang));
}
