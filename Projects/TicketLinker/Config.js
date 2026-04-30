/**
 * Config.js — TicketLinker
 *
 * المهمة: ربط ملفات PDF التذاكر بحجاجها في شيت B2C عبر العمود Z (TICKET_URL).
 * بعد كتابة الرابط، GDS Pipeline يلتقطه كل 30د ويحلّله ويفرّغ بياناته.
 *
 * نطاق هذا المشروع: M1 (PnrResolver) + M2 (Matcher) + M3 (Writer) فقط.
 * GDS مسؤول عن الـ parsing والكتابة في AG-CF/BI/BO/BP/BQ/BR — لا نلمسه.
 */

var TL = (typeof TL !== 'undefined') ? TL : {};

TL.Config = {
  // ==================== مصادر البيانات ====================

  // مجلد Drive الرئيسي للتذاكر (الوجهة C — كل شيء يُنسَخ إليه)
  TICKETS_FOLDER_ID: '154uCHRSpZnC5iGvdL1uGqVroy6txUIvG',

  // مجلدات Drive الأصلية (read-only) — تُزامَن دورياً إلى TICKETS_FOLDER_ID
  SOURCE_FOLDER_IDS: [
    { id: '1XkkGiaNjU58V11i7aL9-4I9LcAjHRw82', name: 'A_inbox' },
    { id: '1Yfr7w4R-lXgRwu66k7bnoYhmjzAlOki6', name: 'B_inbox' }
  ],

  // مجلدات فرعية في C يتجاوزها Pipeline (لا يستهلك Claude عليها)
  // Unresolved_Tickets: PDFs الفاشلة المُنسخة للمراجعة اليدوية
  // B_inbox: تذاكر مصدرها مختلف (لا تطابق PD) — تُراجَع يدوياً عبر TL_BInboxReview
  EXCLUDED_FOLDERS: ['Unresolved_Tickets', 'B_inbox'],

  // الشيت الرئيسي (نفس SHEET_ID لـ GDS)
  SS_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',

  // أسماء التبويبات في الشيت الرئيسي
  SHEET_B2C: 'B2C',
  SHEET_FLIGHTS: 'الطيران',
  SHEET_PD: 'Presonal Details',

  // ==================== Workbench Sheet (يُنشأ مرة واحدة) ====================

  // SHEET_ID للـ workbench (PdfManifest, Proposals, Audit, RunLog)
  // يُملأ عند إنشاء الـ workbench لأول مرة
  WORKBENCH_SS_ID: '',  // TODO: ضع SHEET_ID بعد إنشاء الـ workbench

  WB_TABS: {
    MANIFEST: 'PdfManifest',   // كل PDF + PNR المستخرَج
    PROPOSALS: 'Proposals',    // اقتراحات الكتابة قبل التنفيذ
    AUDIT: 'Audit',            // نتائج الـ self-audit
    RUN_LOG: 'RunLog',         // سجل التشغيل (totals)
    UNRESOLVED: 'Unresolved',  // PDFs بدون PNR صحيح للمراجعة اليدوية
    COPY_LOG: 'TL_CopyLog'     // سجل المزامنة من A/B إلى C
  },

  // ==================== أعمدة B2C (مرآة لـ GDS Config) ====================
  // 1-based (للتوافق مع SpreadsheetApp)

  B2C_COL: {
    PASSPORT: 6,        // F - رقم جواز السفر (مفتاح البحث)
    FIRST_NAME_EN: 11,  // K
    LAST_NAME_EN: 12,   // L
    CONTRACT_TYPE: 21,  // U - B2B/B2C في PD
    CONTRACT_NAME: 22,  // V
    TICKET_URL: 26,     // Z - المُستهدَف للكتابة
    PNR: 61             // BI - من شيت الطيران (يُستنسخ في B2C)
  },

  // أسماء الـ headers (للـ Header-based detection — حماية من column shift)
  B2C_HEADERS: {
    PASSPORT: 'رقم جواز السفر',
    FIRST_NAME_EN: 'الاسم الأول (إنجليزي)',
    LAST_NAME_EN: 'اسم العائلة (إنجليزي)',
    TICKET_URL: 'رابط التذكرة',
    PNR: 'PNR'
  },

  // ==================== أعمدة PD (المُستهدَف الجديد للكتابة) ====================
  // PD = Presonal Details. الكتابة الآن في PD، وسكريبت آخر ينقل لـ B2C

  PD_COL: {
    PASSPORT: 6,        // F - رقم جواز السفر (مفتاح البحث)
    FIRST_NAME_EN: 11,  // K
    LAST_NAME_EN: 12,   // L
    CONTRACT_NAME: 22,  // V - اسم العقد
    TICKET_NO: 25,      // Y
    TICKET_URL: 26      // Z - المُستهدَف للكتابة الجديد
  },

  PD_HEADERS: {
    PASSPORT: 'رقم جواز السفر',
    FIRST_NAME_EN: 'الاسم الأول (إنجليزي)',
    LAST_NAME_EN: 'اسم العائلة (إنجليزي)',
    TICKET_URL: 'رابط التذكرة'
  },

  // ==================== أعمدة شيت الطيران ====================
  // (للبحث عن pool بـ PNR)

  FLIGHTS_COL: {
    PNR: 2,             // B - PNR (قد يحوي عدة بفاصل " - ")
    CONTRACT_NAME: 91   // CM - اسم العقد (1-based: CM = 91)
  },

  // ==================== Claude API ====================

  CLAUDE_MODEL: 'claude-haiku-4-5',
  CLAUDE_API_URL: 'https://api.anthropic.com/v1/messages',
  CLAUDE_ANTHROPIC_VERSION: '2023-06-01',
  CLAUDE_MAX_TOKENS: 4000,  // 1500 كان يقطع JSON على التذاكر متعددة الركّاب → CLAUDE_ERROR دائم
  CLAUDE_MAX_RETRIES: 3,

  // ==================== المطابقة ====================

  // عتبة قبول auto-OK (تجنب false positives)
  MATCH_CONFIDENCE_THRESHOLD: 0.92,
  // عتبة اللجوء لـ Claude (بين هذه و الـ THRESHOLD)
  MATCH_LOW_CONFIDENCE_THRESHOLD: 0.80,
  AUDIT_SAMPLE_MIN: 5,
  AUDIT_SAMPLE_PERCENT: 0.10,  // 10% من proposals، حد أدنى 5

  // ==================== حدود التشغيل ====================

  MAX_RUNTIME_MS: 5.5 * 60 * 1000,  // 5.5 د (ترك هامش لـ GAS timeout)
  CLAUDE_RATE_LIMIT_MS: 1500,

  // ==================== Script Properties keys ====================

  PROP: {
    ANTHROPIC_API_KEY: 'Claude Api',
    LOCK_FLAG: 'TL_RUNNING',                // قفل التشغيل المتزامن
    CURSOR_M1: 'TL_CURSOR_M1',              // cursor لـ Manifest crawler
    CURSOR_M2: 'TL_CURSOR_M2',              // cursor لـ Matcher
    CURSOR_M3: 'TL_CURSOR_M3'               // cursor لـ Writer
  },

  // ==================== Regex helpers ====================

  // PNR = 6 أحرف/أرقام، uppercase
  PNR_REGEX: /\b([A-Z0-9]{6})\b/g,

  // PNR في اسم ملف (يبدأ بـ PNR)
  PNR_FILENAME_REGEX: /^([A-Z0-9]{6})(?:[._\s\-]|$)/,

  // PNR في نص PDF (بعد كلمات مفتاحية)
  PNR_TEXT_PATTERNS: [
    /Airline\s+Reference\s*\(?PNR\)?\s*:?\s*([A-Z0-9]{6})/i,
    /Record\s+Locator\s*:?\s*([A-Z0-9]{6})/i,
    /Booking\s+Reference\s*:?\s*([A-Z0-9]{6})/i,
    /PNR\s*:?\s*([A-Z0-9]{6})/i
  ]
};
