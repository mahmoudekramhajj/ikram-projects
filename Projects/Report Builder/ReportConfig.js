/**
 * ReportConfig.js — تعريف الشيتات والأعمدة لـ Report Builder v2.0
 * بحث شامل عبر كل الشيتات — كشف تلقائي للأعمدة
 */

var SPREADSHEET_ID = '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s';
var CACHE_TTL = 300; // 5 minutes

// ══════════════════════════════════════════════
//  إعدادات الشيتات المعروفة
//  (dataStartRow وأيقونات — الباقي يُكتشف تلقائياً)
// ══════════════════════════════════════════════

var SHEET_CONFIG = {
  'الباقات':          { dataStartRow: 3, icon: '📦', nameEn: 'Packages' },
  'الطيران':          { dataStartRow: 3, icon: '✈️', nameEn: 'Flights' },
  'رحلة الحاج ':      { dataStartRow: 2, icon: '🕋', nameEn: 'Pilgrim Journey' },
  'رحلة الحاج 2':     { dataStartRow: 2, icon: '🕋', nameEn: 'Journey v2' },
  'Presonal Details':  { dataStartRow: 2, icon: '👤', nameEn: 'Personal Details' },
  'الفنادق':          { dataStartRow: 2, icon: '🏨', nameEn: 'Hotels' },
  'مخيم مني':         { dataStartRow: 2, icon: '⛺', nameEn: 'Mina Camp' },
  'Tour Guide':        { dataStartRow: 2, icon: '🧭', nameEn: 'Tour Guide' },
  'Guide Rabih':       { dataStartRow: 2, icon: '🧭', nameEn: 'Guide Rabih' },
  'المستخدمين':       { dataStartRow: 2, icon: '👥', nameEn: 'Users' },
  'BotSessions':       { dataStartRow: 2, icon: '🤖', nameEn: 'Bot Sessions' },
  'Hotel Check-in':    { dataStartRow: 2, icon: '🏨', nameEn: 'Hotel Check-in' },
  'Room Mapping':      { dataStartRow: 2, icon: '🚪', nameEn: 'Room Mapping' },
  'Room Type':         { dataStartRow: 2, icon: '🛏️', nameEn: 'Room Type' },
  'ExchangeRates':     { dataStartRow: 2, icon: '💱', nameEn: 'Exchange Rates' },
  'PNR Target Countries': { dataStartRow: 2, icon: '🌍', nameEn: 'PNR Countries' }
};

// أيقونة افتراضية للشيتات غير المعروفة
var DEFAULT_ICON = '📄';
var DEFAULT_START_ROW = 2;

// ══════════════════════════════════════════════
//  روابط الشيتات (Join Relationships)
//  عند اختيار أعمدة من شيتين، هذه العلاقات تربطهما
// ══════════════════════════════════════════════

var SHEET_JOINS = [
  {
    sheets: ['رحلة الحاج ', 'Presonal Details'],
    nameAr: 'رحلة الحاج ↔ البيانات الشخصية',
    nameEn: 'Journey ↔ Personal Details',
    type: 'groupMatch',
    // رحلة الحاج.GroupNumber (6) = PD.رقم المجموعة (1)
    primaryKey: { sheet: 'رحلة الحاج ', col: 6 },
    secondaryKey: { sheet: 'Presonal Details', col: 1 },
    // داخل المجموعة: مطابقة بالجنس + نوع الحاج
    matchKeys: {
      primary: { gender: 11, isMain: 10 },
      secondary: { gender: 4, isMain: 2 }
    }
  },
  {
    sheets: ['رحلة الحاج ', 'الباقات'],
    nameAr: 'رحلة الحاج ↔ الباقات',
    nameEn: 'Journey ↔ Packages',
    type: 'simple',
    primaryKey: { sheet: 'رحلة الحاج ', col: 1 }, // PackageId
    secondaryKey: { sheet: 'الباقات', col: 1 }     // Nusk No
  },
  {
    sheets: ['Presonal Details', 'الباقات'],
    nameAr: 'البيانات الشخصية ↔ الباقات',
    nameEn: 'Personal Details ↔ Packages',
    type: 'simple',
    primaryKey: { sheet: 'Presonal Details', col: 18 }, // رقم الباقة
    secondaryKey: { sheet: 'الباقات', col: 1 }         // Nusk No
  },
  {
    sheets: ['رحلة الحاج 2', 'Presonal Details'],
    nameAr: 'رحلة الحاج 2 ↔ البيانات الشخصية',
    nameEn: 'Journey v2 ↔ Personal Details',
    type: 'simple',
    primaryKey: { sheet: 'رحلة الحاج 2', col: 5 },   // ApplicationId
    secondaryKey: { sheet: 'Presonal Details', col: 1 } // رقم المجموعة
  },
  {
    sheets: ['رحلة الحاج ', 'Guide Rabih'],
    nameAr: 'رحلة الحاج ↔ المرشدين',
    nameEn: 'Journey ↔ Guides',
    type: 'simple',
    primaryKey: { sheet: 'رحلة الحاج ', col: 8 }, // Passport
    secondaryKey: { sheet: 'Guide Rabih', col: 5 } // Passport column in Guide Rabih
  }
];

// ══════════════════════════════════════════════
//  كشف تلقائي — قراءة كل الشيتات وأعمدتها
// ══════════════════════════════════════════════

/**
 * يرجع قائمة كل الشيتات مع أعمدتها (headers)
 * يُخزّن في الكاش لمدة 5 دقائق
 */
function getAllSheetsWithColumns() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('allSheetsColumns');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allSheets = ss.getSheets();
  var result = [];

  for (var i = 0; i < allSheets.length; i++) {
    var sheet = allSheets[i];
    var name = sheet.getName();

    // تجاهل الشيتات الفارغة أو المخفية
    if (sheet.isSheetHidden()) continue;
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastCol === 0 || lastRow === 0) continue;

    var cfg = SHEET_CONFIG[name] || {};
    var startRow = cfg.dataStartRow || DEFAULT_START_ROW;
    var icon = cfg.icon || DEFAULT_ICON;
    var nameEn = cfg.nameEn || name;

    // قراءة الهيدر (الصف الأول أو ما قبل بداية البيانات)
    var headerRow = startRow - 1;
    if (headerRow < 1) headerRow = 1;
    var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

    var columns = [];
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').trim();
      if (!h) h = 'col_' + (c + 1);
      columns.push({
        index: c,
        name: h,
        // تخمين النوع من اسم العمود
        type: guessColumnType_(h)
      });
    }

    result.push({
      sheetName: name,
      nameAr: name,
      nameEn: nameEn,
      icon: icon,
      dataStartRow: startRow,
      rowCount: Math.max(0, lastRow - startRow + 1),
      columns: columns
    });
  }

  // ترتيب: الشيتات الكبيرة أولاً
  result.sort(function(a, b) { return b.rowCount - a.rowCount; });

  var json = JSON.stringify(result);
  // تقسيم الكاش إذا كبير
  if (json.length < 100000) {
    cache.put('allSheetsColumns', json, CACHE_TTL);
  }

  return result;
}

/**
 * تخمين نوع العمود من اسمه
 */
function guessColumnType_(name) {
  var n = name.toLowerCase();

  // تاريخ
  if (n.indexOf('date') > -1 || n.indexOf('تاريخ') > -1 ||
      n.indexOf('check-in') > -1 || n.indexOf('check-out') > -1 ||
      n.indexOf('start') > -1 || n.indexOf('end') > -1 ||
      n.indexOf('بداية') > -1 || n.indexOf('نهاية') > -1 ||
      n.indexOf('انتهاء') > -1 || n.indexOf('إصدار') > -1 ||
      n.indexOf('ميلاد') > -1 || n.indexOf('dob') > -1) {
    return 'date';
  }

  // نسبة
  if (n.indexOf('%') > -1 || n.indexOf('نسبة') > -1 || n.indexOf('pct') > -1 ||
      n.indexOf('percentage') > -1) {
    return 'percentage';
  }

  // عملة / سعر
  if (n.indexOf('price') > -1 || n.indexOf('سعر') > -1 || n.indexOf('fare') > -1 ||
      n.indexOf('amount') > -1 || n.indexOf('مبلغ') > -1 || n.indexOf('total') > -1 ||
      n.indexOf('إجمالي') > -1 || n.indexOf('sar') > -1 || n.indexOf('profit') > -1 ||
      n.indexOf('ربح') > -1 || n.indexOf('deposit') > -1 || n.indexOf('دفعة') > -1) {
    return 'currency';
  }

  // رقم
  if (n.indexOf('no.') > -1 || n.indexOf('عدد') > -1 || n.indexOf('count') > -1 ||
      n.indexOf('pax') > -1 || n.indexOf('rooms') > -1 || n.indexOf('beds') > -1 ||
      n.indexOf('sold') > -1 || n.indexOf('remaining') > -1 || n.indexOf('متبقي') > -1 ||
      n.indexOf('مبيعات') > -1 || n === 'no' || n === 'الرقم') {
    return 'number';
  }

  return 'text';
}

/**
 * البحث عن علاقة Join بين شيتين
 */
function findJoinRelation(sheet1, sheet2) {
  for (var i = 0; i < SHEET_JOINS.length; i++) {
    var join = SHEET_JOINS[i];
    if ((join.sheets[0] === sheet1 && join.sheets[1] === sheet2) ||
        (join.sheets[0] === sheet2 && join.sheets[1] === sheet1)) {
      return join;
    }
  }
  return null;
}
