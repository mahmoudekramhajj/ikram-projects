/**
 * Hotel Management App — Backend
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v3.1 — Internal Room IDs + 4-Phase AutoAssign + Optimization + Split Files
 */

// ============================================================
// CONFIGURATION
// ============================================================

const HOTEL_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  
 SHEETS: {
    JOURNEY: 'رحلة الحاج ',
    ROOM_TYPE: 'Room Type Preview',
    PACKAGES: 'الباقات',
    TOUR_GUIDE: 'Tour Guide',
    ROOM_MAPPING: 'Room Mapping',
    PRE_ASSIGN_LOG: 'PA_Log'
  },
  // رحلة الحاج column indices (0-based)
  JOURNEY_COLS: {
    BOOKING_ID: 5,
    PACKAGE_ID: 1,
    GROUP_NUMBER: 6,
    NAME: 7,
    PASSPORT: 8,
    GENDER: 11,
    NATIONALITY_EN: 12,
    COUNTRY_RESIDENCE: 14,
    ARRIVAL_TIME: 18,
    ARRIVAL_CITY: 19,
    ARRIVAL_DATE: 20,
    ARRIVAL_FLIGHT: 24,
    RETURN_DEPT_CITY: 31,
    RETURN_DEPT_DATE: 32,
    RETURN_DEPT_TIME: 33,
    RETURN_FLIGHT: 34,
    FIRST_HOUSE: 36,
    FIRST_HOUSE_START: 37,
    FIRST_HOUSE_END: 38,
    LAST_HOUSE: 39,
    LAST_HOUSE_START: 40,
    LAST_HOUSE_END: 41,
    MAKKAH_AR: 42,
    MAKKAH_EN: 43,
    MAKKAH_SHIFT_AR: 44,
    MAKKAH_SHIFT_EN: 45,
    MADINAH_AR: 46,
    MADINAH_EN: 47
  },
  
  // Room Type Preview column indices (0-based)
  ROOM_COLS: {
    GROUP_NUMBER: 0,
    PACKAGE_NAME: 3,
    PILGRIMS: 4,
    MALE: 5,
    FEMALE: 6,
    NATIONALITY: 7,
    GROUP_LEADER: 8,
    PHONE: 9,
    TOUR_GUIDE: 10,
    MADINAH_HOTEL: 15,
    ROOM_TYPE_MED: 16,
    BEDS_MED: 17,
    FULL_ROOMS_MED: 18,
    SHARED_BEDS_MED: 19,
    MAKKAH1_HOTEL: 20,
    ROOM_TYPE_MAK1: 21,
    BEDS_MAK1: 22,
    FULL_ROOMS_MAK1: 23,
    SHARED_BEDS_MAK1: 24,
    MAKKAH2_HOTEL: 25,
    ROOM_TYPE_MAK2: 26,
    BEDS_MAK2: 27,
    FULL_ROOMS_MAK2: 28,
    SHARED_BEDS_MAK2: 29
  },
  
  // Tour Guide sheet (0-based) — reads columns A-K (11 columns)
  // Structure: Guide name (A) → Pilgrim passport (E) → Registered (J) → Unique (K)
  GUIDE_COLS: {
    GUIDE_NAME: 0,     // A: اسم المرشد
    PASSPORT: 4,        // E: رقم جواز الحاج
    REG_STATUS: 9,      // J: حالة التسجيل (✅ Registered)
    UNIQUE_CHECK: 10    // K: فحص التكرار (✅ Unique)
  },
  GUIDE_START_ROW: 2,
  
  // Hotel abbreviations for internal room IDs (City + Abbr + Type + Seq)
  HOTEL_ABBR: {
    // المدينة المنورة (M)
    'Concorde Dar Al Khair Hotel': 'CDK',
    'Al Shakreen Golden Tulip': 'SGT',
    'Movenpick Anwar almadinah hotel': 'MOV',
    'EMAAR ELITE ALMADINAH HOTEL': 'EEL',
    'EMAAR ROYAL HOTEL': 'ERO',
    'Millennium Alaqeeq COMPANY': 'MIL',
    'Madinah Hilton Hotel': 'HIL',
    'Diyar Al Taqwa Hotel': 'DYT',
    'Taiba Front Hotel': 'TFR',
    'Crowne Plaza Madinah': 'CRW',
    'DarAl Taqwa Hotel': 'DAT',
    // مكة المكرمة (K) — موحّد مع مكة تحويل
    'Anjum Hotel': 'ANJ',
    'Park Plaza Hotel': 'PPZ',
    'Emaar Legend Hotel': 'ELG',
    'Biak Otel': 'BIK',
    'Dar Almaqam Hotel': 'DAM',
    'MAAD INTERNATIONAL HOTEL CO LTD': 'MAD',
    'Swissotel Makkah Hotel': 'SWM',
    'Sheraton Makkah Company is a Single Person Company': 'SHR',
    'EMAAR AL RAWDA 2': 'ER2',
    'ALDANAH ALMASIYYAH HOTEL COMPANY': 'ADM',
    'Makkah Clock Royal Tower A Fairmont Hotel': 'FCT',
    'Holiday Inn Bakkah Hotel': 'HIB',
    'Jabal Omar Hyatt Regency Hotel': 'HYT',
    'Swissotel Al Maqam Makkah': 'SAM',
    'Emaar Al Noor Hotel': 'ENR'
  },
  
  // Time margins (hours)
  // ARRIVAL: مطار_مدينة الفندق — ساعات من هبوط الطائرة حتى الوصول للفندق
  // DEPARTURE: مدينة الفندق_مطار المغادرة — ساعات قبل الرحلة يجب المغادرة
  // INTERCITY: وسيلة النقل — ساعات السفر بين المدن
  MARGINS: {
    ARRIVAL: {
      'Madinah_Madina': 3,
      'Jeddah_Madina': 8,
      'Jeddah_Makkah': 4,
      'Jeddah_Makkah Shifting': 4,
      'Madinah_Makkah': 8
    },
    DEPARTURE: {
      'Makkah_Jeddah': 8,
      'Madina_Madinah': 3,
      'Madina_Jeddah': 12,
      'Makkah_Madinah': 12
    },
    INTERCITY: {
      'bus': 7,
      'train': null   // null = موعد غير محدد — يُحدد لاحقاً
    }
  },
  
  CHECKOUT_TIME: '12:00',
  CHECKIN_TIME: '13:00',
  
  ROOM_MAPPING_HEADERS: [
    'HotelName', 'InternalRoomId', 'RoomType', 'Capacity',
    'CheckIn', 'CheckOut', 'PackageRef', 'ActualRoomNo',
    'Status', 'OccupantIDs'
  ],
  ROOM_MAPPING_COLS: {
    HOTEL_NAME: 0, INTERNAL_ID: 1, ROOM_TYPE: 2, CAPACITY: 3,
    CHECK_IN: 4, CHECK_OUT: 5, PACKAGE_REF: 6, ACTUAL_ROOM_NO: 7,
    STATUS: 8, OCCUPANT_IDS: 9
  },
  // Hotel sheet column headers
  HOTEL_SHEET_HEADERS: [
    'ApplicantId', 'اسم الحاج', 'رقم الجواز', 'الجنس', 'الجنسية',
    'رقم المجموعة', 'اسم الباقة', 'المدينة', 'المرحلة',
    'بداية العقد', 'نهاية العقد',
    'رحلة الوصول', 'تاريخ الوصول', 'وقت الوصول المتوقع',
    'حالة الوصول', 'نوع الغرفة', 'RoomGroup_ID',
    'رقم الغرفة', 'حالة Check-in', 'وقت Check-in',
    'رحلة المغادرة', 'تاريخ المغادرة', 'وسيلة النقل'
  ]
};

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('HotelIndex');
  return template.evaluate()
    .setTitle('إدارة الفنادق | إكرام الضيف')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// DATA: GET HOTEL LIST
// ============================================================

function getHotelList() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var data = sheet.getDataRange().getValues();
  
  var hotels = { madinah: {}, makkah: {} };
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var medHotel = row[J.MADINAH_EN];
    if (medHotel && medHotel !== 'NULL') {
      if (!hotels.madinah[medHotel]) hotels.madinah[medHotel] = 0;
      hotels.madinah[medHotel]++;
    }
    
    var makHotel = row[J.MAKKAH_EN];
    if (makHotel && makHotel !== 'NULL') {
      if (!hotels.makkah[makHotel]) hotels.makkah[makHotel] = 0;
      hotels.makkah[makHotel]++;
    }
   var shiftHotel = row[J.MAKKAH_SHIFT_EN];
    if (shiftHotel && shiftHotel !== 'NULL') {
      if (!hotels.makkah[shiftHotel]) hotels.makkah[shiftHotel] = 0;
      hotels.makkah[shiftHotel]++;
    }
  }
  
  return hotels;
}
