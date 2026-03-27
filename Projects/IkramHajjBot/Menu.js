// ============================================
// القائمة الرئيسية + الترحيب + العد التنازلي
// ============================================

function sendMainMenu_(chatId, lang) {
  var countdown = getCountdownText_(chatId, lang);

  var buttons = [
    [{ text: T_('btn_flight', lang), callback_data: 'my_flight' }, { text: T_('btn_hotel', lang), callback_data: 'my_hotel' }],
    [{ text: T_('btn_package', lang), callback_data: 'my_package' }, { text: T_('btn_transport', lang), callback_data: 'my_transport' }],
    [{ text: T_('btn_my_data', lang), callback_data: 'my_data' }],
    [{ text: T_('btn_emergency', lang), callback_data: 'emergency' }, { text: T_('btn_contact', lang), callback_data: 'contact_company' }],
    [{ text: T_('btn_change_lang', lang), callback_data: 'change_lang' }, { text: T_('btn_refresh', lang), callback_data: 'refresh_data' }],
    [{ text: T_('btn_switch', lang), callback_data: 'switch_pilgrim' }]
  ];

  sendMessage_(chatId, countdown + T_('menu_title', lang), { inline_keyboard: buttons });
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
// ⏳ العد التنازلي
// ============================================
function getCountdownText_(chatId, lang) {
  var session = getSession_(chatId);
  if (!session || !session.passport) return '';

  var pilgrim = findPilgrimByPassport_(session.passport);
  if (!pilgrim) return '';

  var r = pilgrim.rowData;
  var arrDate = normDate_(r[20]);
  var depDate = normDate_(r[32]);

  if (!arrDate || !depDate) return '';

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  var arrive = parseDate_(arrDate);
  var depart = parseDate_(depDate);

  if (!arrive || !depart) return '';

  var msPerDay = 86400000;

  if (today < arrive) {
    var daysLeft = Math.ceil((arrive - today) / msPerDay);
    return T_('countdown_before', lang, { days: daysLeft }) + '\n\n';
  }

  if (today >= arrive && today <= depart) {
    var dayNum = Math.floor((today - arrive) / msPerDay) + 1;
    return T_('countdown_during', lang, { day: dayNum }) + '\n\n';
  }

  return T_('countdown_after', lang) + '\n\n';
}
