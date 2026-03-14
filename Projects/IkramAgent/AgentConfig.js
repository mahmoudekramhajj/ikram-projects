/**
 * إعدادات وكيل إكرام
 * الإصدار 2.1 - مع حفظ الجلسات
 */

var AGENT_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  
  SHEETS: {
    PACKAGES: 'الباقات',
    FLIGHTS: 'الطيران',
    HOTELS: 'الفنادق',
    SETTINGS: 'إعدادات_الوكيل',
    LEADS: 'طلبات_المتابعة',
    CHAT_LOG: 'سجل_المحادثات'
  },
  
  DATA_START_ROW: 3,
  HOTELS_START_ROW: 2
};

/**
 * جلب إعداد من شيت الإعدادات
 */
function getAgentSetting(key) {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.SETTINGS);
    if (!sheet) return null;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === key) {
        return String(data[i][1]).trim();
      }
    }
    return null;
  } catch (e) {
    Logger.log('Error getting setting: ' + e.toString());
    return null;
  }
}

/**
 * جلب جميع الإعدادات
 */
function getAllAgentSettings() {
  return {
    CLAUDE_API_KEY: getAgentSetting('CLAUDE_API_KEY'),
    TELEGRAM_BOT_TOKEN: getAgentSetting('TELEGRAM_BOT_TOKEN'),
    TWILIO_ACCOUNT_SID: getAgentSetting('TWILIO_ACCOUNT_SID'),
    TWILIO_AUTH_TOKEN: getAgentSetting('TWILIO_AUTH_TOKEN'),
    TWILIO_WHATSAPP_NUMBER: getAgentSetting('TWILIO_WHATSAPP_NUMBER'),
    NOTIFICATION_EMAIL: getAgentSetting('NOTIFICATION_EMAIL'),
    NOTIFICATION_PHONE: getAgentSetting('NOTIFICATION_PHONE')
  };
}

// ═══════════════════════════════════════════════════════════
// رسائل الترحيب واختيار اللغة
// ═══════════════════════════════════════════════════════════

var WELCOME_MESSAGE = '🕋 *مرحباً بك في إكرام الضيف للحج*\n' +
  '*Welcome to Ikram AlDyf for Hajj*\n' +
  '*Bienvenue chez Ikram AlDyf*\n\n' +
  '✨ اختر لغتك المفضلة\n' +
  'Choose your language\n' +
  'Choisissez votre langue';

var LANGUAGE_BUTTONS = [
  '🇸🇦 العربية',
  '🇬🇧 English',
  '🇫🇷 Français',
  '✍️ أخرى/Other'
];

// ═══════════════════════════════════════════════════════════
// رسائل القائمة الرئيسية بكل لغة
// ═══════════════════════════════════════════════════════════

var MESSAGES = {
  ar: {
    mainMenu: '🕋 *إكرام الضيف للحج*\n\nكيف يمكنني مساعدتك؟',
    chooseCountry: '🌍 اختر الدولة:',
    chooseHotel: '🏨 اختر الفندق:',
    chooseCity: '✈️ اختر مدينة المغادرة:',
    noPackages: '😔 عذراً، لا توجد باقات متاحة حالياً لهذا الاختيار.',
    contactUs: '📞 *تواصل معنا*\n\n' +
      '📱 واتساب: +966125111940\n' +
      '📧 البريد: info@ikramhajj.com\n\n' +
      'أو أرسل بياناتك وسنتواصل معك قريباً',
    packageDetails: '📦 *تفاصيل الباقة*',
    price: '💰 السعر',
    duration: '📅 المدة',
    hotels: '🏨 الفنادق',
    flights: '✈️ الطيران',
    available: '✅ متاح',
    soldOut: '❌ نفدت',
    remaining: 'متبقي',
    nights: 'ليالي',
    days: 'أيام',
    makkah: 'مكة',
    madinah: 'المدينة',
    backToMenu: '🔙 القائمة الرئيسية',
    otherLang: 'اكتب رسالتك بأي لغة تفضلها وسأرد عليك بنفس اللغة ✍️'
  },
  en: {
    mainMenu: '🕋 *Ikram AlDyf for Hajj*\n\nHow can I help you?',
    chooseCountry: '🌍 Choose country:',
    chooseHotel: '🏨 Choose hotel:',
    chooseCity: '✈️ Choose departure city:',
    noPackages: '😔 Sorry, no packages available for this selection.',
    contactUs: '📞 *Contact Us*\n\n' +
      '📱 WhatsApp: +966125111940\n' +
      '📧 Email: info@ikramhajj.com\n\n' +
      'Or send your details and we\'ll contact you soon',
    packageDetails: '📦 *Package Details*',
    price: '💰 Price',
    duration: '📅 Duration',
    hotels: '🏨 Hotels',
    flights: '✈️ Flights',
    available: '✅ Available',
    soldOut: '❌ Sold Out',
    remaining: 'remaining',
    nights: 'nights',
    days: 'days',
    makkah: 'Makkah',
    madinah: 'Madinah',
    backToMenu: '🔙 Main Menu',
    otherLang: 'Write your message in any language and I\'ll respond in the same ✍️'
  },
  fr: {
    mainMenu: '🕋 *Ikram AlDyf pour le Hajj*\n\nComment puis-je vous aider?',
    chooseCountry: '🌍 Choisissez le pays:',
    chooseHotel: '🏨 Choisissez l\'hôtel:',
    chooseCity: '✈️ Choisissez la ville de départ:',
    noPackages: '😔 Désolé, aucun forfait disponible pour cette sélection.',
    contactUs: '📞 *Contactez-nous*\n\n' +
      '📱 WhatsApp: +966125111940\n' +
      '📧 Email: info@ikramhajj.com\n\n' +
      'Ou envoyez vos coordonnées et nous vous contacterons bientôt',
    packageDetails: '📦 *Détails du forfait*',
    price: '💰 Prix',
    duration: '📅 Durée',
    hotels: '🏨 Hôtels',
    flights: '✈️ Vols',
    available: '✅ Disponible',
    soldOut: '❌ Épuisé',
    remaining: 'restants',
    nights: 'nuits',
    days: 'jours',
    makkah: 'La Mecque',
    madinah: 'Médine',
    backToMenu: '🔙 Menu principal',
    otherLang: 'Écrivez votre message dans n\'importe quelle langue et je répondrai dans la même ✍️'
  }
};

// ═══════════════════════════════════════════════════════════
// أزرار القائمة الرئيسية بكل لغة
// ═══════════════════════════════════════════════════════════

var MENU_BUTTONS = {
  ar: [
    '🌍 باقات حسب الدولة',
    '🏨 باقات حسب الفندق',
    '✈️ باقات حسب مدينة الطيران',
    '📞 تواصل معنا'
  ],
  en: [
    '🌍 Packages by Country',
    '🏨 Packages by Hotel',
    '✈️ Packages by Flight City',
    '📞 Contact Us'
  ],
  fr: [
    '🌍 Forfaits par pays',
    '🏨 Forfaits par hôtel',
    '✈️ Forfaits par ville de vol',
    '📞 Contactez-nous'
  ]
};

// ═══════════════════════════════════════════════════════════
// إدارة جلسات المستخدمين باستخدام Cache Service
// ═══════════════════════════════════════════════════════════

/**
 * جلب جلسة المستخدم
 */
function getUserSession(userId) {
  try {
    var cache = CacheService.getScriptCache();
    var sessionData = cache.get('session_' + userId);
    
    if (sessionData) {
      return JSON.parse(sessionData);
    }
    
    // إنشاء جلسة جديدة
    var newSession = {
      lang: null,
      stage: 'welcome',
      lastActivity: new Date().toISOString()
    };
    
    cache.put('session_' + userId, JSON.stringify(newSession), 21600); // 6 ساعات
    return newSession;
    
  } catch (e) {
    Logger.log('getUserSession Error: ' + e.toString());
    return {
      lang: null,
      stage: 'welcome',
      lastActivity: new Date().toISOString()
    };
  }
}

/**
 * حفظ جلسة المستخدم
 */
function saveUserSession(userId, session) {
  try {
    var cache = CacheService.getScriptCache();
    session.lastActivity = new Date().toISOString();
    cache.put('session_' + userId, JSON.stringify(session), 21600); // 6 ساعات
  } catch (e) {
    Logger.log('saveUserSession Error: ' + e.toString());
  }
}

/**
 * تعيين لغة المستخدم
 */
function setUserLanguage(userId, lang) {
  var session = getUserSession(userId);
  session.lang = lang;
  session.stage = 'main_menu';
  saveUserSession(userId, session);
}

/**
 * تعيين مرحلة المستخدم
 */
function setUserStage(userId, stage) {
  var session = getUserSession(userId);
  session.stage = stage;
  saveUserSession(userId, session);
}