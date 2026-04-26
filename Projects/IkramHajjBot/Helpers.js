// ============================================
// دوال مساعدة — تنسيق التاريخ والوقت
// ============================================

function formatDate_(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    var d = val;
    return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
  }
  var s = String(val);
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  return s || '-';
}

function formatTime_(val) {
  if (!val) return '-';
  // Date objects (GAS returns time cells as Date with 1899 epoch)
  if (typeof val === 'object' && val.getHours) {
    // استخدام Utilities.formatDate مع المنطقة الزمنية للشيت
    // لضمان عرض الوقت تماماً كما هو مخزّن في الخلية
    var tz = SpreadsheetApp.openById(SHEET_ID).getSpreadsheetTimeZone();
    return Utilities.formatDate(val, tz, 'HH:mm');
  }
  var s = String(val);
  // Catch stringified Date objects: "Sun Dec 31 1899 07:01:00 GMT+0300"
  if (s.indexOf('1899') !== -1 || s.indexOf('1900') !== -1) {
    var match = s.match(/(\d{1,2}):(\d{2})/);
    if (match) return pad2_(parseInt(match[1], 10)) + ':' + match[2];
  }
  if (s.indexOf('.') !== -1) s = s.split('.')[0];
  var parts = s.split(':');
  if (parts.length >= 2) return parts[0] + ':' + parts[1];
  return s || '-';
}

function pad2_(n) {
  return n < 10 ? '0' + n : '' + n;
}


// ============================================
// دوال عربية: تحويل الأرقام لعربية-هندية + تنسيق التاريخ/الوقت
// لحل مشكلة RTL/LTR mixing في تيليغرام
// ============================================

// ملاحظة: بقرار تصميمي (2026-04-23) — الأرقام تُعرض لاتينية حتى في الواجهة العربية
// لوضوحها على الحاج. هذه الدالة صارت no-op؛ يُحتفظ بها للتوافق مع كل المستدعين.
function toArabicDigits_(text) {
  if (text === null || text === undefined || text === '') return '';
  return String(text);
}

function formatArabicDate_(val) {
  if (!val) return '-';
  var monthsAr = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
  var monthsEn = { 'jan':0,'feb':1,'mar':2,'apr':3,'may':4,'jun':5,'jul':6,'aug':7,'sep':8,'oct':9,'nov':10,'dec':11 };
  var s = String(val).trim();

  // نمط ١: "Fri 15 May 2026" (من V3)
  var m = s.match(/(?:\w+\s+)?(\d{1,2})\s+(\w+)\s+(\d{4})/);
  if (m && monthsEn[m[2].toLowerCase().substring(0,3)] !== undefined) {
    return toArabicDigits_(m[1]) + ' ' + monthsAr[monthsEn[m[2].toLowerCase().substring(0,3)]] + ' ' + toArabicDigits_(m[3]);
  }
  // نمط ٢: "May 15, 2026"
  m = s.match(/(\w+)\s+(\d{1,2}),?\s+(\d{4})/);
  if (m && monthsEn[m[1].toLowerCase().substring(0,3)] !== undefined) {
    return toArabicDigits_(m[2]) + ' ' + monthsAr[monthsEn[m[1].toLowerCase().substring(0,3)]] + ' ' + toArabicDigits_(m[3]);
  }
  // نمط ٣: 'YYYY-MM-DD' من formatDate_
  var iso = formatDate_(val);
  if (iso === '-') return '-';
  var parts = iso.split('-');
  if (parts.length === 3) {
    var mIdx = parseInt(parts[1], 10) - 1;
    var mon = (mIdx >= 0 && mIdx < 12) ? monthsAr[mIdx] : parts[1];
    return toArabicDigits_(parts[2]) + ' ' + mon + ' ' + toArabicDigits_(parts[0]);
  }
  return toArabicDigits_(iso);
}

function formatArabicTime_(val) {
  if (!val) return '-';
  var s = String(val).trim();
  // صيغة 12-hour: "09:10 PM" / "9:10 AM"
  var ampm = s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (ampm) {
    var hh = pad2_(parseInt(ampm[1], 10));
    var mm = ampm[2];
    var suffix = ampm[3].toUpperCase() === 'PM' ? 'مساءً' : 'صباحاً';
    return toArabicDigits_(hh + ':' + mm) + ' ' + suffix;
  }
  // صيغة 24-hour: "HH:mm"
  var t = formatTime_(val);
  if (t === '-') return '-';
  return toArabicDigits_(t);
}

function parseDate_(dateStr) {
  if (!dateStr) return null;
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return null;
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setHours(0, 0, 0, 0);
  return d;
}

function normDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return val.getFullYear() + '-' + pad2_(val.getMonth() + 1) + '-' + pad2_(val.getDate());
  }
  var s = String(val);
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  return s;
}

function getTodayString_() {
  var d = new Date();
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

function getDateOffset_(dateStr, days) {
  var parts = dateStr.split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  d.setDate(d.getDate() + days);
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}

// ============================================
// تحويل رمز IATA للمطار إلى: اسم المطار (المدينة)
// يقع على الرمز الخام إن لم يكن مُسجّلاً في AIRPORTS
// ============================================
function getAirportDisplay_(code, lang) {
  var c = String(code || '').trim().toUpperCase();
  if (!c || c === '-') return '-';
  var isAr = (lang === 'ar');
  var info = (typeof AIRPORTS !== 'undefined') ? AIRPORTS[c] : null;
  if (!info) return c;  // fallback — الرمز كما هو
  var airport = isAr ? info.ar : info.en;
  var city = isAr ? info.cityAr : info.cityEn;
  return airport + ' (' + city + ')';
}
