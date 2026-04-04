/**
 * Transport Management App — Helper Functions
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0 — Cache, Sheet helpers, Formatters, ID generation
 */

// ============================================================
// CACHE: CacheService — TTL 2 دقيقة
// ============================================================

var CACHE_TTL = 120;

function getCachedData_(cacheKey) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {}
  return null;
}

function setCachedData_(cacheKey, data) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    if (json.length < 100000) {
      cache.put(cacheKey, json, CACHE_TTL);
    }
  } catch (e) {}
}

function clearCache_(cacheKey) {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove(cacheKey);
  } catch (e) {}
}

// ============================================================
// SHEET: findSheet_ مع trim — بحث آمن عن الشيت
// ============================================================

function findSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim() === name.trim()) return sheets[i];
  }
  return null;
}

// ============================================================
// FORMAT: Date & Time
// ============================================================

function formatDate_(dateVal) {
  if (!dateVal) return '';
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, 'Asia/Riyadh', 'yyyy-MM-dd');
  }
  var str = String(dateVal);
  if (str.length >= 10) return str.substring(0, 10);
  return str;
}

function formatTime_(timeVal) {
  if (!timeVal) return '';
  // Date object (GAS يخزن الوقت كتاريخ 1899)
  if (timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, 'Asia/Riyadh', 'HH:mm');
  }
  var str = String(timeVal);
  // إذا فيه 1899 أو Day name → هذا Date object تحوّل لنص
  if (str.indexOf('1899') >= 0 || str.indexOf('1900') >= 0 || str.match(/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)/)) {
    try {
      var d = new Date(str);
      if (!isNaN(d.getTime())) {
        var h = d.getHours();
        var m = d.getMinutes();
        return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
      }
    } catch (e) {}
  }
  if (str.includes(':')) {
    var parts = str.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
  }
  return str;
}

function formatDateTime_(date) {
  if (!date) return '';
  if (!(date instanceof Date)) date = new Date(date);
  return Utilities.formatDate(date, 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');
}

function formatTimeArabic_(timeStr) {
  if (!timeStr) return '--:--';
  var parts = String(timeStr).split(':');
  if (parts.length < 2) return timeStr;
  var hours = parseInt(parts[0]);
  var minutes = parts[1];
  var period = hours >= 12 ? 'م' : 'ص';
  var displayHour = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
  return String(displayHour).padStart(2, '0') + ':' + minutes + ' ' + period;
}

// ============================================================
// TIME: Subtract hours from time string
// ============================================================

function subtractHours_(timeStr, hours) {
  if (!timeStr) return null;
  var parts = timeStr.split(':');
  var h = parseInt(parts[0]);
  var m = parseInt(parts[1]) || 0;

  h -= Math.floor(hours);
  m -= (hours % 1) * 60;

  if (m < 0) { h--; m += 60; }
  if (h < 0) h += 24;

  return String(h).padStart(2, '0') + ':' + String(Math.round(m)).padStart(2, '0');
}

function addHoursToTime_(timeStr, hours) {
  if (!timeStr) return null;
  var parts = timeStr.split(':');
  var h = parseInt(parts[0]);
  var m = parseInt(parts[1]) || 0;

  h += Math.floor(hours);
  m += (hours % 1) * 60;

  if (m >= 60) { h++; m -= 60; }
  if (h >= 24) h -= 24;

  return String(h).padStart(2, '0') + ':' + String(Math.round(m)).padStart(2, '0');
}

// ============================================================
// CITY: Code helpers
// ============================================================

function getCityCode_(hotelCity) {
  if (!hotelCity) return '';
  if (hotelCity === 'Madina' || hotelCity === 'Madinah') return 'Madina';
  if (hotelCity === 'Makkah' || hotelCity === 'Makkah Shifting') return 'Makkah';
  return hotelCity;
}

function isTrainTransport_(transport) {
  if (!transport) return false;
  var str = String(transport).trim();
  return str === 'قطار' || str.toLowerCase() === 'train' || str.indexOf('قطار') >= 0;
}

// ============================================================
// ID: Generate unique IDs
// ============================================================

function generateTripId_() {
  var now = new Date();
  var date = Utilities.formatDate(now, 'Asia/Riyadh', 'yyMMdd');
  var seq = String(Math.floor(Math.random() * 9000) + 1000);
  return 'TR-' + date + '-' + seq;
}

function generateBoardingId_() {
  var now = new Date();
  var date = Utilities.formatDate(now, 'Asia/Riyadh', 'yyMMdd');
  var seq = String(Math.floor(Math.random() * 9000) + 1000);
  return 'BD-' + date + '-' + seq;
}

function generateIncidentId_() {
  var now = new Date();
  var date = Utilities.formatDate(now, 'Asia/Riyadh', 'yyMMdd');
  var seq = String(Math.floor(Math.random() * 9000) + 1000);
  return 'INC-' + date + '-' + seq;
}

function generateVehicleId_() {
  return 'VH-' + String(Math.floor(Math.random() * 90000) + 10000);
}

function generateContractId_() {
  return 'TC-' + String(Math.floor(Math.random() * 90000) + 10000);
}

// ============================================================
// AUTH: Session management
// ============================================================

function createSession_(userData) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put(
    TRANSPORT_CONFIG.SESSION_PREFIX + token,
    JSON.stringify(userData),
    TRANSPORT_CONFIG.SESSION_DURATION
  );
  return token;
}

function validateSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var sessionData = cache.get(TRANSPORT_CONFIG.SESSION_PREFIX + token);
  if (!sessionData) return null;
  try {
    var user = JSON.parse(sessionData);
    // Refresh session
    cache.put(TRANSPORT_CONFIG.SESSION_PREFIX + token, sessionData, TRANSPORT_CONFIG.SESSION_DURATION);
    return user;
  } catch (e) {
    return null;
  }
}

// ============================================================
// AUTH: Login
// ============================================================

function getLoginUser() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}

    var role = 'field';
    var name = 'موظف ميداني';

    // إذا تم التعرف على الإيميل — ابحث في شيت المستخدمين
    if (email) {
      var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
      var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.USERS);
      if (sheet) {
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          var userEmail = String(row[1] || '').trim().toLowerCase();
          if (userEmail === email.toLowerCase()) {
            name = String(row[0] || '').trim();
            role = String(row[2] || 'field').trim();
            break;
          }
        }
      }
      if (name === 'موظف ميداني') name = email.split('@')[0];
    }

    // السماح بالدخول دائماً — حتى بدون إيميل
    var userData = { email: email || 'guest', name: name, role: role, location: '' };
    var token = createSession_(userData);
    return { success: true, user: userData, token: token };
  } catch (e) {
    // حتى لو فشل كل شيء — ادخل كضيف
    var guestData = { email: 'guest', name: 'زائر', role: 'field', location: '' };
    return { success: true, user: guestData, token: createSession_(guestData) };
  }
}

// ============================================================
// TRANSPORT MAP: قراءة نوع النقل من الباقات
// ============================================================

function getTransportMap_(ss) {
  var cached = getCachedData_('transportMap_transport');
  if (cached) return cached;

  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.PACKAGES);
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  // قراءة حتى عمود BQ (69 عمود) لضمان الوصول لعمود التنقل BN (66)
  var data = sheet.getRange(1, 1, lastRow, 69).getValues();
  var map = {};

  for (var i = 2; i < data.length; i++) {
    // المطابقة بـ Nusk No (عمود B = index 1) — لأن PackageId في رحلة الحاج = Nusk No
    var nuskNo = data[i][TRANSPORT_CONFIG.PKG_COLS.NUSK_NO];
    var transport = data[i][TRANSPORT_CONFIG.PKG_COLS.TRANSPORT];
    if (nuskNo && transport) {
      map[String(nuskNo).trim()] = String(transport).trim();
    }
  }

  setCachedData_('transportMap_transport', map);
  return map;
}

// ============================================================
// DEPARTURE: حساب وقت مغادرة الحاج من الفندق
// ============================================================

function calculateDepartureTime_(returnTime, hotelCity, returnCity) {
  var marginKey = getCityCode_(hotelCity) + '_' + returnCity;
  var margin = TRANSPORT_CONFIG.MARGINS.DEPARTURE[marginKey] || 8;
  return subtractHours_(formatTime_(returnTime), margin);
}

// ============================================================
// ARRIVAL: حساب الوقت المتوقع للوصول للفندق
// ============================================================

function calculateExpectedArrival_(arrivalTime, arrivalCity, hotelCity) {
  var marginKey = arrivalCity + '_' + getCityCode_(hotelCity);
  var margin = TRANSPORT_CONFIG.MARGINS.ARRIVAL[marginKey] || 4;
  return addHoursToTime_(formatTime_(arrivalTime), margin);
}

// ============================================================
// CLEAN: تنظيف البيانات
// ============================================================

function cleanRow_(row) {
  return row.map(function(cell) {
    if (cell instanceof Date) {
      return Utilities.formatDate(cell, 'Asia/Riyadh', 'yyyy-MM-dd');
    }
    if (cell === null || cell === undefined) return '';
    var str = String(cell).trim();
    // تنظيف: NULL, null, undefined, #N/A → فارغ
    if (str === 'NULL' || str === 'null' || str === 'undefined' || str === '#N/A' || str === '#REF!') return '';
    return str;
  });
}

function cleanValue_(val) {
  if (!val) return '';
  var str = String(val).trim();
  if (str === 'NULL' || str === 'null' || str === 'undefined' || str === '#N/A') return '';
  return str;
}
