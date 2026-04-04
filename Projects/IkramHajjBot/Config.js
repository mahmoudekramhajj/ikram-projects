// ============================================
// إعدادات البوت والثوابت
// ============================================
var BOT_TOKEN = '8694589281:AAHvT-anZgLDk6s5YO8WStP7Y2zhq6-BDIE';
var TELEGRAM_API = 'https://api.telegram.org/bot' + BOT_TOKEN;
var SHEET_ID = '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s';
var JOURNEY_SHEET = 'رحلة الحاج '; // مسافة في النهاية — مهم
var PILGRIM_DATA_SHEET = 'Pilgrim Data'; // شيت البيانات المجمّعة من الحجاج
var DRIVE_FOLDER_ID = ''; // TODO: أضف ID مجلد Google Drive لحفظ صور الجوازات
var CACHE_TTL = 1800; // مدة الكاش بالثواني (30 دقيقة)

// ============================================
// المدراء — Chat IDs المصرّح لهم باستخدام /broadcast و /stats
// ============================================
var ADMIN_IDS = ['8566760392']; // إكرام الضيف

// ============================================
// أسماء الشيتات الجديدة
// ============================================
var ADMIN_MESSAGES_SHEET = 'AdminMessages';
var DELIVERY_LOG_SHEET = 'DeliveryLog';
var BOT_ACTIVITY_SHEET = 'BotActivity';

// خريطة أعمدة AdminMessages
var AM = {
  ID: 0, TITLE: 1, MSG_AR: 2, MSG_EN: 3, MSG_FR: 4, MSG_DE: 5, MSG_IT: 6, MSG_ES: 7,
  IMAGE_URL: 8, FILE_URL: 9, FILE_NAME: 10, TARGET: 11, TARGET_VALUE: 12,
  PRIORITY: 13, STATUS: 14, SENT_AT: 15, SENT_COUNT: 16, FAIL_COUNT: 17, CREATED_BY: 18
};

// خريطة اللغة → عمود الرسالة في AdminMessages
var LANG_TO_COL = { ar: AM.MSG_AR, en: AM.MSG_EN, fr: AM.MSG_FR, de: AM.MSG_DE, it: AM.MSG_IT, es: AM.MSG_ES };

// ============================================
// أعمدة تأكيد الوصول في شيت "رحلة الحاج"
// ============================================
var COL_RECEPTION_STATUS = 49;  // AX — حالة الاستقبال ("تم")
var COL_RECEPTION_TIME = 50;    // AY — وقت التأكيد
var COL_RECEPTION_STAFF = 51;   // AZ — مصدر التأكيد (اسم الموظف أو Bot)

// ============================================
// مجموعات العمليات — المطارات (Telegram Chat IDs)
// ============================================
var OPS_GROUPS = {
  madinah: '-4849598886',        // مطار المدينة
  jeddah_t1: '-5220583519',      // مطار جدة — صالة 1
  jeddah_north: '-5267173490'    // مطار جدة — الصالة الشمالية
};

// مجموعات العمليات — المواقع
var OPS_LOCATION_GROUPS = {
  makkah: '-4916619724',         // عمليات مكة
  madinah: '-5284394785',        // عمليات المدينة
  mashaaer: '-5268778683'        // عمليات المشاعر
};

// ============================================
// جدول صالات مطار جدة حسب شركة الطيران
// ============================================
var AIRLINE_TERMINAL = {
  'Emirates': 'T1', 'Etihad': 'T1', 'Gulf Air': 'T1',
  'EgyptAir': 'T1', 'Qatar Airways': 'T1', 'Royal Jordanian': 'T1',
  'Saudia': 'T1', 'Turkish Airlines': 'T1', 'Ethiopian Airlines': 'T1',
  'Flyadeal': 'T1',
  'AJet': 'N', 'Wizz Air': 'N', 'Aegean Airlines': 'N',
  'Air Cairo': 'N', 'AnadoluJet': 'N', 'Flydubai': 'N',
  'Pegasus Airlines': 'N', 'WEST ISLE AIR INC.': 'N'
};

// ============================================
// CacheService — دوال مساعدة
// ============================================
function getCache_(key) {
  var cache = CacheService.getScriptCache();
  var data = cache.get(key);
  if (data) {
    try { return JSON.parse(data); } catch(e) { return null; }
  }
  return null;
}

function setCache_(key, value, ttl) {
  var cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(value), ttl || CACHE_TTL);
}

function clearPilgrimCache_(passport) {
  var cache = CacheService.getScriptCache();
  cache.remove('pilgrim_' + String(passport).toUpperCase().trim());
}

function clearSessionCache_(chatId) {
  var cache = CacheService.getScriptCache();
  cache.remove('session_' + String(chatId));
}

function clearTransportCache_(packageId) {
  var cache = CacheService.getScriptCache();
  cache.remove('transport_' + String(packageId).trim());
}

function clearHotelMapCache_(rowData) {
  var cache = CacheService.getScriptCache();
  var names = [
    String(rowData[42] || ''), String(rowData[43] || ''),
    String(rowData[44] || ''), String(rowData[45] || ''),
    String(rowData[46] || ''), String(rowData[47] || '')
  ];
  for (var i = 0; i < names.length; i++) {
    if (names[i] && names[i] !== '-') {
      cache.remove('hmap2_' + names[i]);
    }
  }
}
