// ============================================
// Menu.js — القوائم والأزرار
// ============================================

/**
 * إرسال رسالة الترحيب مع أزرار اللغة
 */
function sendWelcome_(chatId) {
  sendMessage_(chatId, T_('welcome', 'ar'), {
    inline_keyboard: [
      [{ text: '🇸🇦 العربية', callback_data: 'lang_ar' },
       { text: '🇬🇧 English', callback_data: 'lang_en' }]
    ]
  });
}

/**
 * إرسال القائمة الرئيسية حسب الدور
 */
function sendMainMenu_(chatId, session) {
  var lang = session.language || 'ar';
  var role = session.role;
  var keyboard = [];

  // الصف 1: لوحة القيادة + بحث
  keyboard.push([
    { text: T_('btn_dashboard', lang), callback_data: 'dash' },
    { text: T_('btn_search', lang), callback_data: 'search' }
  ]);

  // الصف 2: الباقات + الرحلات
  keyboard.push([
    { text: T_('btn_packages', lang), callback_data: 'packages' },
    { text: T_('btn_flights', lang), callback_data: 'flights' }
  ]);

  // الصف 3: الفنادق + المخيمات
  keyboard.push([
    { text: T_('btn_hotels', lang), callback_data: 'hotels' },
    { text: T_('btn_camps', lang), callback_data: 'camps' }
  ]);

  // الصف 4: المبيعات + التصدير (Admin/Manager فقط)
  if (role === ROLES.ADMIN || role === ROLES.MANAGER) {
    keyboard.push([
      { text: T_('btn_sales', lang), callback_data: 'sales' },
      { text: T_('btn_export', lang), callback_data: 'export' }
    ]);
  }

  // الصف 5: AI
  keyboard.push([
    { text: T_('btn_ai', lang), callback_data: 'ai' }
  ]);

  // الصف 6: الإعدادات + (إدارة للأدمن)
  if (role === ROLES.ADMIN) {
    keyboard.push([
      { text: T_('btn_settings', lang), callback_data: 'settings' },
      { text: T_('btn_users', lang), callback_data: 'admin_users' }
    ]);
    keyboard.push([
      { text: T_('btn_broadcast', lang), callback_data: 'broadcast' },
      { text: T_('btn_stats', lang), callback_data: 'bot_stats' }
    ]);
  } else {
    keyboard.push([
      { text: T_('btn_settings', lang), callback_data: 'settings' }
    ]);
  }

  sendMessage_(chatId, T_('menu_title', lang), { inline_keyboard: keyboard });
}

/**
 * إرسال قائمة الإعدادات
 */
function sendSettingsMenu_(chatId, session) {
  var lang = session.language || 'ar';
  sendMessage_(chatId, T_('settings_title', lang), {
    inline_keyboard: [
      [{ text: T_('btn_change_lang', lang), callback_data: 'change_lang' }],
      [{ text: T_('btn_logout', lang), callback_data: 'logout' }],
      [{ text: T_('btn_back', lang), callback_data: 'back_main' }]
    ]
  });
}

/**
 * إرسال قائمة فرعية للرحلات
 */
function sendFlightsMenu_(chatId, session) {
  var lang = session.language || 'ar';
  sendMessage_(chatId, T_('flights_title', lang), {
    inline_keyboard: [
      [{ text: T_('btn_flights_today', lang), callback_data: 'flights_today' },
       { text: T_('btn_flights_week', lang), callback_data: 'flights_week' }],
      [{ text: T_('btn_flights_all', lang), callback_data: 'flights_all' }],
      [{ text: T_('btn_back', lang), callback_data: 'back_main' }]
    ]
  });
}

/**
 * إرسال قائمة التصدير
 */
function sendExportMenu_(chatId, session) {
  var lang = session.language || 'ar';
  sendMessage_(chatId, T_('export_prompt', lang), {
    inline_keyboard: [
      [{ text: T_('btn_packages', lang), callback_data: 'exp_packages' },
       { text: T_('btn_flights', lang), callback_data: 'exp_flights' }],
      [{ text: T_('btn_search', lang) + ' (Excel)', callback_data: 'exp_pilgrims' }],
      [{ text: T_('btn_back', lang), callback_data: 'back_main' }]
    ]
  });
}

/**
 * زر الرجوع للقائمة الرئيسية
 */
function backButton_(lang) {
  return {
    inline_keyboard: [
      [{ text: T_('btn_back', lang), callback_data: 'back_main' }]
    ]
  };
}
