// ============================================
// IkramAdmin — إعدادات البوت والثوابت
// ============================================

var BOT_TOKEN = '8735744366:AAH2NNWBxA8IYb0HHlueixywL81GmffJLjQ';
var TELEGRAM_API = 'https://api.telegram.org/bot' + BOT_TOKEN;

var SHEET_ID = '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s';
var CACHE_TTL = 1800; // 30 دقيقة

// ============================================
// المدراء — Chat IDs
// ============================================
var ADMIN_IDS = ['8566760392'];

// ============================================
// الأدوار
// ============================================
var ROLES = {
  ADMIN: 'admin',
  MANAGER: 'manager',
  EMPLOYEE: 'employee'
};

// ============================================
// أسماء الشيتات
// ============================================
var SHEETS = {
  // مصادر حية (من نسك مباشرة)
  JOURNEY2: 'رحلة الحاج 2',
  PERSONAL: 'Presonal Details',
  // مصدر مدمج (JourneyMerger)
  JOURNEY: 'رحلة الحاج ',        // مسافة في النهاية — مهم!
  // مصادر ثابتة
  PACKAGES: 'الباقات',
  FLIGHTS: 'الطيران',
  HOTELS: 'الفنادق',
  CAMPS: 'مخيم مني',
  GUIDES: 'Tour Guide',
  USERS: 'المستخدمين',
  GUIDE_RABIH: 'Guide Rabih',
  // شيتات خاصة بالبوت (تُنشأ تلقائياً)
  BOT_SESSIONS: 'AdminBotSessions',
  BOT_LOG: 'AdminBotLog'
};

// ============================================
// فهارس أعمدة Presonal Details (المصدر الأحدث)
// ============================================
var PD = {
  SEQ: 0,              // الرقم التسلسلي
  GROUP: 1,            // رقم المجموعة (= ApplicationId)
  TYPE: 2,             // نوع الحاج (رئيسي/عضو)
  CATEGORY: 3,         // فئة الحجاج
  GENDER: 4,           // الجنس
  PASSPORT: 5,         // رقم جواز السفر
  PASSPORT_EXP: 6,     // تاريخ انتهاء الجواز
  PASSPORT_ISSUE: 7,   // تاريخ الإصدار
  FIRST_NAME_AR: 8,    // الاسم الأول (عربي)
  LAST_NAME_AR: 9,     // اسم العائلة (عربي)
  FIRST_NAME_EN: 10,   // الاسم الأول (إنجليزي)
  LAST_NAME_EN: 11,    // اسم العائلة (إنجليزي)
  DOB: 12,             // تاريخ الميلاد
  EMAIL: 13,           // البريد الإلكتروني
  PHONE: 14,           // رقم الجوال
  GUIDE_NAME: 15,      // اسم المرشد الديني (إنجليزي)
  COUNTRY_RESIDENCE: 16, // بلد الإقامة
  NATIONALITY: 17,     // الجنسية
  PACKAGE_NO: 18,      // رقم الباقة
  PACKAGE_NAME: 19,    // اسم الباقة
  FLIGHT_TYPE: 20,     // نوع عقد الطيران (B2B/B2C)
  CONTRACT_NAME: 21,   // اسم العقد
  VISA_STATUS: 22,     // حالة التأشيرة
  IN_KSA: 23,          // داخل المملكة
  TICKET_NO: 24,       // رقم التذكرة
  TICKET_URL: 25,      // رابط التذكرة
  INVOICE_NO: 26,      // رقم الفاتورة
  CAMP: 27,            // المخيم (AB)
  TRANSPORT_ARR_TYPE: 28,  // نوع النقل (وصول)
  TRANSPORT_ARR_TIME: 29,  // وقت الوصول - النقل
  TRANSPORT_DEP_TYPE: 30,  // نوع النقل (مغادرة)
  TRANSPORT_DEP_TIME: 31,  // وقت المغادرة - النقل
  BOOKING_DETAILS: 32     // رابط نسك
};

// ============================================
// فهارس أعمدة رحلة الحاج 2 (المصدر الأحدث للرحلة)
// ============================================
var J2 = {
  BOOKING_ID: 0,
  PACKAGE_ID: 1,
  SERVICE_PROVIDER: 2,
  PACKAGE_YEAR: 3,
  CAMP_NAME: 4,
  APPLICATION_ID: 5,   // = رقم المجموعة في PD
  IS_MAIN: 6,
  GENDER: 7,
  NATIONALITY_EN: 8,
  COUNTRY_RESIDENCE_EN: 9,
  // بيانات الطيران (وصول)
  ARR_AIRLINE_AR: 10,
  ARR_AIRLINE_EN: 11,
  ARR_ARRIVAL_TIME: 12,
  ARR_ARRIVE_CITY: 13,
  ARR_ARRIVE_DATE: 14,
  ARR_DEPART_CITY: 15,
  ARR_DEPART_DATE: 16,
  ARR_DEPART_TIME: 17,
  ARR_FLIGHT_NO: 18,
  ARR_FLIGHT_TYPE: 19,
  // بيانات الطيران (عودة)
  RET_AIRLINE_AR: 20,
  RET_AIRLINE_EN: 21,
  RET_ARRIVAL_TIME: 22,
  RET_ARRIVE_CITY: 23,
  RET_ARRIVE_DATE: 24,
  RET_DEPART_CITY: 25,
  RET_DEPART_DATE: 26,
  RET_DEPART_TIME: 27,
  RET_FLIGHT_NO: 28,
  RET_FLIGHT_TYPE: 29,
  // الفنادق
  FIRST_HOUSE: 30,
  FIRST_HOUSE_START: 31,
  FIRST_HOUSE_END: 32,
  LAST_HOUSE: 33,
  LAST_HOUSE_START: 34,
  LAST_HOUSE_END: 35,
  MAKKAH_AR: 36,
  MAKKAH_EN: 37,
  MAKKAH_SHIFT_AR: 38,
  MAKKAH_SHIFT_EN: 39,
  MADINAH_AR: 40,
  MADINAH_EN: 41
};

// ============================================
// فهارس أعمدة الباقات
// ============================================
var PKG = {
  NO: 0, NUSK_NO: 1, NAME_AR: 2, CATEGORY: 3, IKRAM_NO: 4,
  PRICE: 5, DATE_START: 6, DATE_END: 7, NO_DAYS: 8,
  CITY_START: 9, NO_PILGRIM: 10,
  // الفندق 1
  H1_CITY: 11, H1_NAME_AR: 12, H1_NAME_EN: 13,
  H1_CHECKIN: 14, H1_CHECKOUT: 15,
  H1_DBL: 16, H1_DBL_PRICE: 17,
  H1_TRI: 18, H1_TRI_PRICE: 19,
  H1_QUAD: 20, H1_QUAD_PRICE: 21,
  H1_ROOMS: 22, H1_BEDS: 23,
  // الفندق 2
  H2_DIFF: 25, H2_CITY: 26, H2_NAME_AR: 27, H2_NAME_EN: 28,
  H2_CHECKIN: 29, H2_CHECKOUT: 30,
  H2_DBL: 31, H2_DBL_PRICE: 32,
  H2_TRI: 33, H2_TRI_PRICE: 34,
  H2_QUAD: 35, H2_QUAD_PRICE: 36,
  H2_ROOMS: 37, H2_BEDS: 38,
  // الفندق 3
  H3_CITY: 41, H3_NAME_AR: 42, H3_NAME_EN: 43,
  H3_CHECKIN: 44, H3_CHECKOUT: 45,
  H3_DBL: 46, H3_DBL_PRICE: 47,
  H3_TRI: 48, H3_TRI_PRICE: 49,
  H3_QUAD: 50, H3_QUAD_PRICE: 51,
  H3_ROOMS: 52, H3_BEDS: 53,
  // إضافية
  PHOTO_LINK: 54, MEAL_START: 55, MEAL_END: 56,
  SALES: 57, REMAINING: 58, PERCENT: 59,
  NAME_EN: 60, BOOKING_URL: 61,
  TRANSPORT: 65,  // BN — التنقل بين المدن (حافلة/قطار)
  // عقود نسك
  NUSK_CONTRACT_H1: 66, NUSK_CONTRACT_H2: 67, NUSK_CONTRACT_H3: 68
};

// ============================================
// فهارس أعمدة الطيران
// ============================================
var FLT = {
  NO: 0, PNR: 1, SUPPLIER: 2, STATUS: 3, COUNTRY: 4,
  CITY: 5, AIRLINE: 6, PAX: 7, NO_DAYS: 8,
  // المالية
  FARE_SOURCE: 9, FARE: 10, CURRENCY: 11, FARE_SAR: 12,
  ADD: 13, TOTAL: 14, PRICE_NUSK: 15, DIFF1: 16,
  PRICE_NUSK2: 17, DIFF2: 18, PROFIT: 19, TOTAL_NUSUK: 20,
  // ذهاب 1
  GO1_FLIGHT: 21, GO1_TAKEOFF_DATE: 22, GO1_TAKEOFF_TIME: 23,
  GO1_FROM: 24, GO1_TO: 25, GO1_LAND_DATE: 26, GO1_LAND_TIME: 27,
  // ذهاب 2
  GO2_FLIGHT: 28, GO2_TAKEOFF_DATE: 29, GO2_TAKEOFF_TIME: 30,
  GO2_FROM: 31, GO2_TO: 32, GO2_LAND_DATE: 33, GO2_LAND_TIME: 34,
  // عودة 1
  RET1_FLIGHT: 35, RET1_TAKEOFF_DATE: 36, RET1_TAKEOFF_TIME: 37,
  RET1_FROM: 38, RET1_TO: 39, RET1_LAND_DATE: 40, RET1_LAND_TIME: 41,
  // عودة 2
  RET2_FLIGHT: 42, RET2_TAKEOFF_DATE: 43, RET2_TAKEOFF_TIME: 44,
  RET2_FROM: 45, RET2_TO: 46, RET2_LAND_DATE: 47, RET2_LAND_TIME: 48,
  // الباقات المرتبطة (10)
  PKG1_NAME: 49, PKG1_NO: 50,
  PKG2_NAME: 51, PKG2_NO: 52,
  PKG3_NAME: 53, PKG3_NO: 54,
  PKG4_NAME: 55, PKG4_NO: 56,
  PKG5_NAME: 57, PKG5_NO: 58,
  PKG6_NAME: 59, PKG6_NO: 60,
  PKG7_NAME: 61, PKG7_NO: 62,
  PKG8_NAME: 63, PKG8_NO: 64,
  PKG9_NAME: 65, PKG9_NO: 66,
  PKG10_NAME: 67, PKG10_NO: 68,
  // المبيعات
  SALES: 69, REMAINING: 70
};

// ============================================
// فهارس أعمدة رحلة الحاج (المدمج — للاستقبال فقط)
// ============================================
var JRN = {
  BOOKING_ID: 0, PACKAGE_ID: 1, NAME: 7, PASSPORT: 8,
  RECEPTION_STATUS: 49, RECEPTION_TIME: 50, RECEPTION_STAFF: 51
};

// ============================================
// CacheService — دوال مساعدة
// ============================================
function getCache_(key) {
  var cache = CacheService.getScriptCache();
  var data = cache.get(key);
  if (data) {
    try { return JSON.parse(data); } catch (e) { return null; }
  }
  return null;
}

function setCache_(key, value, ttl) {
  var cache = CacheService.getScriptCache();
  // CacheService max value = 100KB
  var json = JSON.stringify(value);
  if (json.length > 100000) {
    Logger.log('Cache value too large for key: ' + key + ' (' + json.length + ' chars)');
    return;
  }
  cache.put(key, json, ttl || CACHE_TTL);
}

function clearCache_(key) {
  CacheService.getScriptCache().remove(key);
}
