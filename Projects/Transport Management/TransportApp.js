/**
 * Transport Management App — Backend
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0 — إدارة عمليات النقل (28 عملية)
 */

// ============================================================
// CONFIGURATION
// ============================================================

const TRANSPORT_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',

  SHEETS: {
    // شيتات موجودة (قراءة فقط)
    JOURNEY: 'رحلة الحاج ',
    PACKAGES: 'الباقات',
    PERSONAL_DETAILS: 'Presonal Details',
    FLIGHTS: 'الطيران',
    USERS: 'المستخدمين',
    TOUR_GUIDE: 'Guide Rabih',
    // شيتات جديدة (قراءة + كتابة)
    TRIPS: 'Transport_Trips',
    BOARDING: 'Transport_Boarding',
    VEHICLES: 'Transport_Vehicles',
    CONTRACTS: 'Transport_Contracts',
    INCIDENTS: 'Transport_Incidents',
    HAJJ_SCHEDULE: 'Transport_HajjSchedule',
    AIRLINE_TERMINALS: 'Airline_Terminals'
  },

  // رحلة الحاج column indices (0-based)
  JOURNEY_COLS: {
    BOOKING_ID: 0,
    PACKAGE_ID: 1,
    SERVICE_PROVIDER: 2,
    PACKAGE_YEAR: 3,
    CAMP_NAME: 4,
    APPLICANT_ID: 5,
    GROUP_NUMBER: 6,
    NAME: 7,
    PASSPORT: 8,
    EMAIL: 9,
    IS_MAIN: 10,
    GENDER: 11,
    NATIONALITY_EN: 12,
    NATIONALITY_AR: 13,
    COUNTRY_RESIDENCE_EN: 14,
    COUNTRY_RESIDENCE_AR: 15,
    ARRIVAL_AIRLINE_AR: 16,
    ARRIVAL_AIRLINE_EN: 17,
    ARRIVAL_TIME: 18,
    ARRIVAL_CITY: 19,
    ARRIVAL_DATE: 20,
    DEPARTURE_CITY: 21,
    DEPARTURE_DATE: 22,
    DEPARTURE_TIME: 23,
    ARRIVAL_FLIGHT: 24,
    ARRIVAL_FLIGHT_TYPE: 25,
    RETURN_AIRLINE_AR: 26,
    RETURN_AIRLINE_EN: 27,
    RETURN_ARRIVAL_TIME: 28,
    RETURN_ARRIVE_CITY: 29,
    RETURN_ARRIVE_DATE: 30,
    RETURN_DEPT_CITY: 31,
    RETURN_DEPT_DATE: 32,
    RETURN_DEPT_TIME: 33,
    RETURN_FLIGHT: 34,
    RETURN_FLIGHT_TYPE: 35,
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

  // Presonal Details column indices (0-based)
  PD_COLS: {
    SERIAL: 0,
    GROUP_NUMBER: 1,
    PILGRIM_TYPE: 2,
    PILGRIM_CATEGORY: 3,
    GENDER: 4,
    PASSPORT: 5,
    PASSPORT_EXPIRY: 6,
    ISSUE_DATE: 7,
    FIRST_NAME_AR: 8,
    LAST_NAME_AR: 9,
    FIRST_NAME_EN: 10,
    LAST_NAME_EN: 11,
    DOB: 12,
    EMAIL: 13,
    PHONE: 14,
    GUIDE_NAME: 15,
    COUNTRY_RESIDENCE: 16,
    NATIONALITY: 17,
    PACKAGE_ID: 18,
    PACKAGE_NAME: 19,
    FLIGHT_CONTRACT_TYPE: 20,
    CONTRACT_NAME: 21,
    VISA_STATUS: 22,
    IN_KINGDOM: 23,
    TICKET_NUMBER: 24,
    TICKET_LINK: 25,
    INVOICE_NUMBER: 26,
    CAMP: 27,
    TRANSPORT_TYPE_ARRIVAL: 28,
    TRANSPORT_TIME_ARRIVAL: 29,
    TRANSPORT_TYPE_DEPARTURE: 30,
    TRANSPORT_TIME_DEPARTURE: 31,
    BOOKING_DETAILS: 32
  },

  // الباقات column indices
  PKG_COLS: {
    NO: 0,
    NUSK_NO: 1,
    NAME_AR: 2,
    CATEGORY: 3,
    IKRAM_NO: 4,
    TRANSPORT: 65  // عمود BN — وسيلة النقل بين المدن
  },

  // Transport_Trips column indices
  TRIP_COLS: {
    TRIP_ID: 0,
    OPERATION_TYPE: 1,
    OPERATION_NAME: 2,
    DATE: 3,
    SCHEDULED_TIME: 4,
    ORIGIN: 5,
    ORIGIN_TYPE: 6,
    DESTINATION: 7,
    DESTINATION_TYPE: 8,
    VEHICLE_ID: 9,
    VEHICLE_PLATE: 10,
    DRIVER_NAME: 11,
    DRIVER_PHONE: 12,
    CAPACITY: 13,
    ASSIGNED_COUNT: 14,
    BOARDED_COUNT: 15,
    STATUS: 16,
    LINKED_FLIGHT: 17,
    LINKED_FLIGHT_TIME: 18,
    PACKAGE_IDS: 19,
    GUIDE_NAMES: 20,
    NOTES: 21,
    CREATED_BY: 22,
    CREATED_AT: 23,
    UPDATED_AT: 24
  },

  // Transport_Boarding column indices
  BOARDING_COLS: {
    BOARDING_ID: 0,
    TRIP_ID: 1,
    BOOKING_ID: 2,
    PASSPORT: 3,
    PILGRIM_NAME: 4,
    GROUP_NUMBER: 5,
    PACKAGE_ID: 6,
    SCAN_TIME: 7,
    SCAN_BY: 8,
    SCAN_METHOD: 9,
    DECLARED_STATUS: 10,
    BOARDING_STATUS: 11,
    NOTES: 12
  },

  // Time margins (hours) — from Hotel Management
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
      'train': null
    }
  },

  CHECKOUT_TIME: '12:00',
  CHECKIN_TIME: '13:00',

  CACHE_DURATION: 120,  // ثانية

  // Session
  SESSION_PREFIX: 'TRANSPORT_SESSION_',
  SESSION_DURATION: 21600  // 6 ساعات
};

// ============================================================
// تعريف الـ 28 عملية نقل
// ============================================================

const OPERATIONS = {
  // === A. استقبال المطار — وصول دولي (B2B فقط) ===
  1:  { name: 'مطار المدينة → فنادق المدينة', origin: 'مطار المدينة', originType: 'airport', dest: 'فنادق المدينة', destType: 'hotel', category: 'arrival', marginKey: 'Madinah_Madina', b2bOnly: true },
  2:  { name: 'مطار جدة صالة 1 → فنادق مكة', origin: 'مطار جدة صالة 1', originType: 'airport', dest: 'فنادق مكة', destType: 'hotel', category: 'arrival', marginKey: 'Jeddah_Makkah', b2bOnly: true },
  3:  { name: 'مطار جدة الشمالية → فنادق مكة', origin: 'مطار جدة الشمالية', originType: 'airport', dest: 'فنادق مكة', destType: 'hotel', category: 'arrival', marginKey: 'Jeddah_Makkah', b2bOnly: true },
  4:  { name: 'مطار جدة صالة 1 → فنادق المدينة', origin: 'مطار جدة صالة 1', originType: 'airport', dest: 'فنادق المدينة', destType: 'hotel', category: 'arrival', marginKey: 'Jeddah_Madina', b2bOnly: true },
  5:  { name: 'مطار جدة الشمالية → فنادق المدينة', origin: 'مطار جدة الشمالية', originType: 'airport', dest: 'فنادق المدينة', destType: 'hotel', category: 'arrival', marginKey: 'Jeddah_Madina', b2bOnly: true },

  // === B. النقل بين المدن — حافلة ===
  6:  { name: 'فنادق مكة → فنادق المدينة (حافلة)', origin: 'فنادق مكة', originType: 'hotel', dest: 'فنادق المدينة', destType: 'hotel', category: 'intercity', transport: 'bus', marginKey: 'bus' },
  7:  { name: 'فنادق المدينة → فنادق مكة (حافلة)', origin: 'فنادق المدينة', originType: 'hotel', dest: 'فنادق مكة', destType: 'hotel', category: 'intercity', transport: 'bus', marginKey: 'bus' },

  // === B. النقل بين المدن — أرجل القطار (حافلة من/إلى المحطة) ===
  8:  { name: 'فنادق مكة → محطة قطار مكة', origin: 'فنادق مكة', originType: 'hotel', dest: 'محطة قطار مكة', destType: 'station', category: 'intercity_train', transport: 'train' },
  9:  { name: 'محطة قطار المدينة → فنادق المدينة', origin: 'محطة قطار المدينة', originType: 'station', dest: 'فنادق المدينة', destType: 'hotel', category: 'intercity_train', transport: 'train' },
  10: { name: 'فنادق المدينة → محطة قطار المدينة', origin: 'فنادق المدينة', originType: 'hotel', dest: 'محطة قطار المدينة', destType: 'station', category: 'intercity_train', transport: 'train' },
  11: { name: 'محطة قطار مكة → فنادق مكة', origin: 'محطة قطار مكة', originType: 'station', dest: 'فنادق مكة', destType: 'hotel', category: 'intercity_train', transport: 'train' },

  // === C. التحويل داخل مكة (Shifting) ===
  12: { name: 'تحويل فنادق مكة (Shifting)', origin: 'فندق مكة 1', originType: 'hotel', dest: 'فندق مكة 2', destType: 'hotel', category: 'shifting' },

  // === D. الحج — الذهاب لمنى ===
  13: { name: 'فنادق مكة → مخيم المعيصم', origin: 'فنادق مكة', originType: 'hotel', dest: 'مخيم المعيصم', destType: 'camp', category: 'hajj' },
  14: { name: 'فنادق مكة → مخيم مجر الكبش', origin: 'فنادق مكة', originType: 'hotel', dest: 'مخيم مجر الكبش', destType: 'camp', category: 'hajj' },
  15: { name: 'فنادق مكة → مخيم 72', origin: 'فنادق مكة', originType: 'hotel', dest: 'مخيم 72', destType: 'camp', category: 'hajj' },

  // === D. الحج — المناسك (المعيصم فقط بالحافلة) ===
  16: { name: 'المعيصم → عرفة', origin: 'مخيم المعيصم', originType: 'camp', dest: 'عرفة', destType: 'camp', category: 'hajj_ritual' },
  17: { name: 'عرفة → مزدلفة', origin: 'عرفة', originType: 'camp', dest: 'مزدلفة', destType: 'camp', category: 'hajj_ritual' },
  18: { name: 'مزدلفة → المعيصم', origin: 'مزدلفة', originType: 'camp', dest: 'مخيم المعيصم', destType: 'camp', category: 'hajj_ritual' },

  // === D. الحج — العودة من منى ===
  19: { name: 'مخيم المعيصم → فنادق مكة', origin: 'مخيم المعيصم', originType: 'camp', dest: 'فنادق مكة', destType: 'hotel', category: 'hajj' },
  20: { name: 'مخيم مجر الكبش → فنادق مكة', origin: 'مخيم مجر الكبش', originType: 'camp', dest: 'فنادق مكة', destType: 'hotel', category: 'hajj' },
  21: { name: 'مخيم 72 → فنادق مكة', origin: 'مخيم 72', originType: 'camp', dest: 'فنادق مكة', destType: 'hotel', category: 'hajj' },

  // === E. المغادرة — عودة دولية ===
  22: { name: 'فنادق مكة → مطار جدة صالة 1', origin: 'فنادق مكة', originType: 'hotel', dest: 'مطار جدة صالة 1', destType: 'airport', category: 'departure', marginKey: 'Makkah_Jeddah' },
  23: { name: 'فنادق مكة → مطار جدة الشمالية', origin: 'فنادق مكة', originType: 'hotel', dest: 'مطار جدة الشمالية', destType: 'airport', category: 'departure', marginKey: 'Makkah_Jeddah' },
  24: { name: 'فنادق المدينة → مطار المدينة', origin: 'فنادق المدينة', originType: 'hotel', dest: 'مطار المدينة', destType: 'airport', category: 'departure', marginKey: 'Madina_Madinah' },
  25: { name: 'فنادق المدينة → مطار جدة صالة 1', origin: 'فنادق المدينة', originType: 'hotel', dest: 'مطار جدة صالة 1', destType: 'airport', category: 'departure', marginKey: 'Madina_Jeddah' },
  26: { name: 'فنادق المدينة → مطار جدة الشمالية', origin: 'فنادق المدينة', originType: 'hotel', dest: 'مطار جدة الشمالية', destType: 'airport', category: 'departure', marginKey: 'Madina_Jeddah' },

  // === E. المغادرة — محطات القطار ===
  27: { name: 'فنادق مكة → محطة قطار مكة (مغادرة)', origin: 'فنادق مكة', originType: 'hotel', dest: 'محطة قطار مكة', destType: 'station', category: 'departure_train' },
  28: { name: 'فنادق المدينة → محطة قطار المدينة (مغادرة)', origin: 'فنادق المدينة', originType: 'hotel', dest: 'محطة قطار المدينة', destType: 'station', category: 'departure_train' }
};

// ============================================================
// المواقع — قائمة كاملة لشاشة اختيار الموقع
// ============================================================

const LOCATIONS = {
  madinah: {
    name: 'المدينة المنورة',
    sites: [
      { id: 'madinah_airport', name: 'مطار المدينة', type: 'airport' },
      { id: 'madinah_station', name: 'محطة قطار المدينة', type: 'station' },
      { id: 'madinah_hotels', name: 'فنادق المدينة', type: 'hotel' }
    ]
  },
  makkah: {
    name: 'مكة المكرمة',
    sites: [
      { id: 'jeddah_t1', name: 'مطار جدة صالة 1', type: 'airport' },
      { id: 'jeddah_north', name: 'مطار جدة الشمالية', type: 'airport' },
      { id: 'makkah_station', name: 'محطة قطار مكة', type: 'station' },
      { id: 'makkah_hotels', name: 'فنادق مكة', type: 'hotel' }
    ]
  },
  mashaaer: {
    name: 'المشاعر المقدسة',
    sites: [
      { id: 'camp_maisam', name: 'مخيم المعيصم', type: 'camp' },
      { id: 'camp_mujar', name: 'مخيم مجر الكبش', type: 'camp' },
      { id: 'camp_72', name: 'مخيم 72', type: 'camp' },
      { id: 'arafat', name: 'عرفة', type: 'camp' },
      { id: 'muzdalifah', name: 'مزدلفة', type: 'camp' }
    ]
  }
};

// ============================================================
// الأدوار وصلاحياتها
// ============================================================

const ROLES = {
  ops_manager: {
    name: 'مدير العمليات',
    tabs: ['trips', 'scanner', 'schedule', 'incidents', 'dashboard'],
    canCreate: true,
    canEdit: true,
    canCancel: true,
    allLocations: true
  },
  guide: {
    name: 'المرشد',
    tabs: ['trips', 'scanner', 'incidents'],
    canCreate: false,
    canEdit: false,
    canCancel: false,
    allLocations: false
  },
  field: {
    name: 'موظف ميداني',
    tabs: ['trips', 'scanner', 'incidents'],
    canCreate: false,
    canEdit: false,
    canCancel: false,
    allLocations: false
  },
  airport_staff: {
    name: 'مسؤول مطار',
    tabs: ['trips', 'scanner', 'incidents'],
    canCreate: false,
    canEdit: false,
    canCancel: false,
    allLocations: false
  },
  hotel_staff: {
    name: 'مسؤول فندق',
    tabs: ['trips', 'scanner', 'incidents'],
    canCreate: false,
    canEdit: false,
    canCancel: false,
    allLocations: false
  }
};

// ============================================================
// Telegram — مجموعات العمليات
// ============================================================

const TELEGRAM_CONFIG = {
  BOT_TOKEN: '8694589281:AAHvT-anZgLDk6s5YO8WStF7Y2zhq6-BDIE',
  OPS_GROUPS: {
    madinah: '-4849598886',
    jeddah_t1: '-5220583519',
    jeddah_north: '-5267173490'
  }
};

// ============================================================
// رؤوس أعمدة الشيتات الجديدة
// ============================================================

const SHEET_HEADERS = {
  TRIPS: [
    'TripId', 'OperationType', 'OperationName', 'Date', 'ScheduledTime',
    'Origin', 'OriginType', 'Destination', 'DestinationType',
    'VehicleId', 'VehiclePlate', 'DriverName', 'DriverPhone',
    'Capacity', 'AssignedCount', 'BoardedCount', 'Status',
    'LinkedFlight', 'LinkedFlightTime', 'PackageIds', 'GuideNames',
    'Notes', 'CreatedBy', 'CreatedAt', 'UpdatedAt'
  ],
  BOARDING: [
    'BoardingId', 'TripId', 'BookingId', 'Passport', 'PilgrimName',
    'GroupNumber', 'PackageId', 'ScanTime', 'ScanBy', 'ScanMethod',
    'DeclaredStatus', 'BoardingStatus', 'Notes'
  ],
  VEHICLES: [
    'VehicleId', 'ContractId', 'Type', 'PlateNumber', 'Capacity',
    'DriverName', 'DriverPhone', 'CompanyName', 'Status'
  ],
  CONTRACTS: [
    'ContractId', 'CompanyName', 'ContactPerson', 'Phone', 'Email',
    'OperationTypes', 'VehicleCount', 'StartDate', 'EndDate', 'Status'
  ],
  INCIDENTS: [
    'IncidentId', 'TripId', 'Type', 'Severity', 'Description',
    'AffectedPilgrims', 'Resolution', 'ResolvedBy', 'ReportedBy',
    'ReportedAt', 'ResolvedAt', 'Status'
  ],
  HAJJ_SCHEDULE: [
    'ScheduleId', 'HajjDay', 'OperationType', 'Date', 'Time',
    'CampName', 'PackageIds', 'EstimatedPilgrims', 'BusCount', 'Notes'
  ]
};

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet(e) {
  // ClaudeAPI: إذا وُجد action + key → معالجة API
  if (e && e.parameter && e.parameter.action && e.parameter.key) {
    return handleClaudeAPI_(e);
  }

  var template = HtmlService.createTemplateFromFile('TransportIndex');
  return template.evaluate()
    .setTitle('إدارة النقل | إكرام الضيف')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// INIT: إنشاء الشيتات الجديدة إذا لم تكن موجودة
// ============================================================

function initTransportSheets() {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var created = [];

  var sheetDefs = [
    { name: TRANSPORT_CONFIG.SHEETS.TRIPS, headers: SHEET_HEADERS.TRIPS },
    { name: TRANSPORT_CONFIG.SHEETS.BOARDING, headers: SHEET_HEADERS.BOARDING },
    { name: TRANSPORT_CONFIG.SHEETS.VEHICLES, headers: SHEET_HEADERS.VEHICLES },
    { name: TRANSPORT_CONFIG.SHEETS.CONTRACTS, headers: SHEET_HEADERS.CONTRACTS },
    { name: TRANSPORT_CONFIG.SHEETS.INCIDENTS, headers: SHEET_HEADERS.INCIDENTS },
    { name: TRANSPORT_CONFIG.SHEETS.HAJJ_SCHEDULE, headers: SHEET_HEADERS.HAJJ_SCHEDULE }
  ];

  for (var i = 0; i < sheetDefs.length; i++) {
    var def = sheetDefs[i];
    var sheet = findSheet_(ss, def.name);

    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sheet.getRange(1, 1, 1, def.headers.length)
        .setFontWeight('bold')
        .setBackground('#0D4A3A')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      created.push(def.name);
    }
  }

  return { created: created, message: created.length > 0 ? 'تم إنشاء ' + created.length + ' شيت جديد' : 'جميع الشيتات موجودة' };
}

// ============================================================
// INIT: إنشاء شيت صالات الطيران بمطار جدة
// ============================================================

function initAirlineTerminals() {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheetName = TRANSPORT_CONFIG.SHEETS.AIRLINE_TERMINALS;
  var sheet = findSheet_(ss, sheetName);

  if (sheet) {
    return { success: true, message: 'شيت ' + sheetName + ' موجود مسبقاً' };
  }

  sheet = ss.insertSheet(sheetName);

  var headers = ['Airline Code', 'Airline Name EN', 'Airline Name AR', 'Jeddah Terminal'];

  // البيانات: صالة 1
  var terminal1 = [
    ['EK', 'Emirates', 'طيران الإمارات', 'Terminal 1'],
    ['EY', 'Etihad Airways', 'طيران الاتحاد', 'Terminal 1'],
    ['GF', 'Gulf Air', 'طيران الخليج', 'Terminal 1'],
    ['MS', 'EgyptAir', 'مصر للطيران', 'Terminal 1'],
    ['QR', 'Qatar Airways', 'الخطوط القطرية', 'Terminal 1'],
    ['RJ', 'Royal Jordanian', 'الملكية الأردنية', 'Terminal 1'],
    ['SV', 'Saudia', 'الخطوط السعودية', 'Terminal 1'],
    ['TK', 'Turkish Airlines', 'الخطوط التركية', 'Terminal 1'],
    ['ET', 'Ethiopian Airlines', 'الخطوط الإثيوبية', 'Terminal 1'],
    ['F3', 'Flyadeal', 'طيران أديل', 'Terminal 1']
  ];

  // البيانات: الصالة الشمالية
  var terminalN = [
    ['VF', 'AJet', 'أجيت', 'North Terminal'],
    ['W4', 'Wizz Air', 'ويز إير', 'North Terminal'],
    ['A3', 'Aegean Airlines', 'إيجيان', 'North Terminal'],
    ['SM', 'Air Cairo', 'إير كايرو', 'North Terminal'],
    ['TK', 'AnadoluJet', 'أناضولو جيت', 'North Terminal'],
    ['FZ', 'Flydubai', 'فلاي دبي', 'North Terminal'],
    ['PC', 'Pegasus Airlines', 'بيغاسوس', 'North Terminal'],
    ['WI', 'WEST ISLE AIR INC.', 'ويست آيل إير', 'North Terminal']
  ];

  var allData = [headers];
  allData = allData.concat(terminal1).concat(terminalN);

  sheet.getRange(1, 1, allData.length, 4).setValues(allData);

  // تنسيق الهيدر
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#0D4A3A')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  // تلوين الصفوف
  var t1Range = sheet.getRange(2, 1, terminal1.length, 4);
  t1Range.setBackground('#E8F5E9');
  var tNRange = sheet.getRange(2 + terminal1.length, 1, terminalN.length, 4);
  tNRange.setBackground('#FFF3E0');

  sheet.autoResizeColumns(1, 4);

  return { success: true, message: 'تم إنشاء شيت ' + sheetName + ' — ' + (terminal1.length + terminalN.length) + ' شركة طيران' };
}

// ============================================================
// INIT: إضافة شركات الطيران الناقصة + الأسماء البديلة
// ============================================================

function addMissingAirlines() {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.AIRLINE_TERMINALS);
  if (!sheet) return { success: false, error: 'شيت Airline_Terminals غير موجود' };

  var newRows = [
    // شركات جديدة
    ['BA', 'British Airways', 'الخطوط البريطانية', 'Terminal 1'],
    ['WY', 'Oman Air', 'الطيران العماني', 'Terminal 1'],
    ['XY', 'flynas', 'طيران ناس', 'Terminal 1'],
    // أسماء بديلة لشركات موجودة
    ['EK', 'ALAMARATYH', 'الاماراتية', 'Terminal 1'],
    ['SV', 'Saudi Airlines', 'الخطوط السعودية', 'Terminal 1'],
    ['W4', 'Wizz Air Hungary', 'ويز إير هنغاريا', 'North Terminal']
  ];

  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, newRows.length, 4).setValues(newRows);

  return { success: true, message: 'تمت إضافة ' + newRows.length + ' شركة طيران' };
}

// ============================================================
// DATA: الحصول على قائمة العمليات
// ============================================================

function getOperations() {
  return OPERATIONS;
}

function getLocations() {
  return LOCATIONS;
}

function getRoles() {
  return ROLES;
}
