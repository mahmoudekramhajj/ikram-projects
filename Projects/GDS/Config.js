/**
 * Config.js — الإعدادات المركزية (namespace: GDS2)
 *
 * ملاحظة: نستخدم namespace GDS2 لتجنّب التعارض مع CONFIG في الكود القديم.
 * في المرحلة 9، بعد حذف الكود القديم، نعيد تسميته إلى GDS.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Config = {
  // ==================== Version Marker ====================
  // يتغيّر مع كل deploy للتأكد من تحميل الكود الجديد
  PROJECT_VERSION: '@61-2026-04-26-version-marker',

  // ==================== Spreadsheet ====================
  SS_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',

  // ==================== Sheets ====================
  SHEET_PD: 'Presonal Details',
  SHEET_B2C: 'B2C',
  SHEET_FLIGHTS: 'الطيران',

  // ==================== B2C أعمدة أساسية (ثابتة) ====================
  // (أرقام الأعمدة 1-based)
  COL: {
    SERIAL: 1,            // A - الرقم التسلسلي
    GROUP: 2,             // B - رقم المجموعة
    TYPE: 3,              // C - نوع الحاج
    CATEGORY: 4,          // D - فئة الحجاج
    GENDER: 5,            // E - الجنس
    PASSPORT: 6,          // F - رقم جواز السفر (مفتاح)
    PASSPORT_EXPIRY: 7,   // G
    ISSUE_DATE: 8,        // H
    FIRST_NAME_AR: 9,     // I
    LAST_NAME_AR: 10,     // J
    FIRST_NAME_EN: 11,    // K
    LAST_NAME_EN: 12,     // L
    BIRTH_DATE: 13,       // M
    EMAIL: 14,            // N
    PHONE: 15,            // O
    GUIDE_NAME: 16,       // P
    RESIDENCE_COUNTRY: 17,// Q
    NATIONALITY: 18,      // R
    PACKAGE_NO: 19,       // S
    PACKAGE_NAME: 20,     // T
    CONTRACT_TYPE: 21,    // U - نوع عقد الطيران
    CONTRACT_NAME: 22,    // V
    VISA_STATUS: 23,      // W
    INSIDE_KSA: 24,       // X
    TICKET_NO: 25,        // Y - رقم التذكرة
    TICKET_URL: 26,       // Z - رابط التذكرة PDF
    INVOICE_NO: 27,       // AA

    // النقل البري (لا نلمسها)
    TRANSPORT_ARRIVAL: 28, // AB
    ARRIVAL_TIME: 29,      // AC
    TRANSPORT_DEP: 30,     // AD
    DEPARTURE_TIME: 31,    // AE
    BOOKING_DETAILS: 32,   // AF

    // PNR
    PNR: 61,  // BI

    // الأعمدة الجديدة (تُنشأ تلقائياً في أول تشغيل)
    LAST_URL: 62,         // BJ - آخر رابط معالَج
    MANUAL: 63,           // BK - مُعدَّل يدوياً
    FAIL_COUNT: 64,       // BL - عدد محاولات الفشل
    NUSUK_AUTH: 65,       // BM - يحتاج دخول نسك

    // أعمدة المرحلة 1: تتبّع تغييرات الإيميل (تُضاف تلقائياً)
    CONTRACT_TYPE_BN: 66, // BN - نوع عقد الطيران (B2B/B2C)
    CHANGE_TYPE: 67,      // BO - نوع التغيير (إلغاء/تعديل ذهاب/...)
    CHANGE_SOURCE: 68,    // BP - مصدر التغيير (إيميل/رابط/يدوي)
    EMAIL_DATE: 69,       // BQ - تاريخ آخر إيميل
    EMAIL_ID: 70,         // BR - معرّف الإيميل (Gmail Message ID)

    // DEP0: قطعة الذهاب الإضافية (الأبعد عن جدة) — 7 أعمدة BS-BY (71-77)
    DEP0_FLIGHT_NO: 71,   // BS
    DEP0_DEP_DATE: 72,    // BT
    DEP0_DEP_TIME: 73,    // BU
    DEP0_FROM: 74,        // BV
    DEP0_TO: 75,          // BW
    DEP0_ARR_DATE: 76,    // BX
    DEP0_ARR_TIME: 77,    // BY

    // RET3: قطعة العودة الإضافية (الأبعد عن جدة) — 7 أعمدة BZ-CF (78-84)
    RET3_FLIGHT_NO: 78,   // BZ
    RET3_DEP_DATE: 79,    // CA
    RET3_DEP_TIME: 80,    // CB
    RET3_FROM: 81,        // CC
    RET3_TO: 82,          // CD
    RET3_ARR_DATE: 83,    // CE
    RET3_ARR_TIME: 84     // CF
  },

  // ==================== Flight blocks ====================
  // كل block = 7 أعمدة (FlightNo, DepDate, DepTime, From, To, ArrDate, ArrTime)
  FLIGHT_BLOCKS: {
    DEP1: { start: 33, end: 39 },  // AG-AM — ترانزيت الذهاب
    DEP2: { start: 40, end: 46 },  // AN-AT — الذهاب الأساسي (دائماً)
    RET1: { start: 47, end: 53 },  // AU-BA — العودة الأساسي (دائماً)
    RET2: { start: 54, end: 60 },  // BB-BH — ترانزيت العودة
    DEP0: { start: 71, end: 77 },  // BS-BY — قطعة الذهاب الإضافية (3+ قطع)
    RET3: { start: 78, end: 84 }   // BZ-CF — قطعة العودة الإضافية (3+ قطع)
  },
  FIELDS_PER_BLOCK: 7,
  FLIGHT_RANGE_START: 33,  // AG
  FLIGHT_RANGE_END: 61,    // BI (بما فيه PNR)
  FLIGHT_RANGE_WIDTH: 29,  // 33 → 61

  // ==================== رؤوس الأعمدة الجديدة ====================
  NEW_COL_HEADERS: {
    62: 'آخر رابط معالَج',
    63: 'مُعدَّل يدوياً',
    64: 'عدد محاولات الفشل',
    65: 'يحتاج دخول نسك',
    66: 'نوع عقد الطيران',
    67: 'نوع التغيير',
    68: 'مصدر التغيير',
    69: 'تاريخ آخر إيميل',
    70: 'معرّف الإيميل',
    // DEP0 (قطعة الذهاب الإضافية)
    71: 'FlightNo 0',
    72: 'Date TAKEOFF 0',
    73: 'TAKEOFF TIME 0',
    74: 'From 0',
    75: 'To 0',
    76: 'DATE LANDING 0',
    77: 'LANDING TIME 0',
    // RET3 (قطعة العودة الإضافية)
    78: 'FlightNo 3',
    79: 'Date TAKEOFF 3',
    80: 'TAKEOFF TIME 3',
    81: 'From 3',
    82: 'To 3',
    83: 'DATE LANDING 3',
    84: 'LANDING TIME 3'
  },

  // ==================== قوائم منسدلة (Data Validation) ====================
  DROPDOWN_VALUES: {
    66: ['B2B', 'B2C'],  // نوع عقد الطيران
    67: [
      'إلغاء',
      'تعديل الذهاب',
      'تعديل الإياب',
      'تعديل الذهاب والإياب',
      'تأخّر إداري',
      'تغيير معدّات'
    ],
    68: ['إيميل', 'رابط', 'يدوي']
  },

  // ==================== حدود التشغيل ====================
  MAX_FAIL_ATTEMPTS: 2,
  MAX_RUNTIME_MS: 5.5 * 60 * 1000,
  CLAUDE_RATE_LIMIT_MS: 1500,
  TELEGRAM_RATE_LIMIT_MS: 100,
  CANARY_SIZE: 10,

  // ==================== موسم الحج ====================
  SEASON_START: '2026-04-01',
  SEASON_END: '2026-07-31',
  SAUDI_ARRIVAL_AIRPORTS: ['JED', 'MED'],

  // ==================== Script Properties keys ====================
  PROP: {
    ANTHROPIC_API_KEY: 'ANTHROPIC_API_KEY',
    TELEGRAM_BOT_TOKEN: 'TELEGRAM_BOT_TOKEN',
    TELEGRAM_CHAT_ID: 'TELEGRAM_GDS_CHAT_ID',
    IATA_REGISTRY: 'IATA_REGISTRY',
    AIRLINE_REGISTRY: 'AIRLINE_REGISTRY',
    LAST_MIGRATION: 'LAST_SCHEMA_MIGRATION',
    LAST_RUN: 'LAST_PIPELINE_RUN'
  },

  // ==================== Drive ====================
  PDF_FOLDER_NAME: 'GDS_Tickets',

  // ==================== Claude API ====================
  CLAUDE_MODEL: 'claude-haiku-4-5',
  CLAUDE_API_URL: 'https://api.anthropic.com/v1/messages',
  CLAUDE_ANTHROPIC_VERSION: '2023-06-01',
  CLAUDE_MAX_TOKENS: 1500,
  CLAUDE_MAX_RETRIES: 3,

  // ==================== ألوان الصفوف ====================
  ROW_COLORS: {
    SUCCESS: '#d9ead3', // أخضر فاتح
    WARNING: '#fff2cc', // أصفر فاتح
    ERROR: '#f4cccc',   // أحمر فاتح
    MANUAL: '#cfe2f3',  // أزرق فاتح
    NUSUK: '#d9d9d9'    // رمادي فاتح
  }
};
