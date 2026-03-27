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
