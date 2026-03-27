/**
 * =====================================================
 *  إكرام الضيف — بحث موظفي المطار
 *  Airport Staff Search — Code.gs (Backend)
 *  Version: 3.0 — مع نظام المصادقة + الفلاتر التفاعلية
 * =====================================================
 */

// ─────────────────────────────────────────────────────
//  ثوابت المشروع
// ─────────────────────────────────────────────────────

const CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'رحلة الحاج ',
  CACHE_DURATION: 300,
  CACHE_KEY: 'AIRPORT_SEARCH_DATA',
  EMAIL_RECIPIENTS: '',
  HOURS_BEFORE: 72
};

const COL = {
  BOOKING_ID:       0,
  PACKAGE_ID:       1,
  GROUP_NUMBER:     6,
  NAME:             7,
  PASSPORT:         8,
  GENDER:          11,
  NATIONALITY_EN:  12,
  NATIONALITY_AR:  13,
  RESIDENCE:       14,
  AIRLINE_AR:      16,
  AIRLINE_EN:      17,
  ARRIVAL_TIME:    18,
  ARRIVE_CITY:     19,
  ARRIVE_DATE:     20,
  DEPARTURE_CITY:  21,
  FLIGHT_NUMBER:   24,
  FLIGHT_TYPE:     25,
  FIRST_HOUSE:     36,
  MAKKAH_EN:       43,
  MAKKAH_SHIFT_EN: 45,
  MADINAH_EN:      47,
  HALL:            48
};


// ─────────────────────────────────────────────────────
//  ثوابت المصادقة
// ─────────────────────────────────────────────────────

const AUTH_CONFIG = {
  USERS_SHEET: 'المستخدمين',
  SESSION_DURATION: 21600,
  SESSION_PREFIX: 'AUTH_SESSION_'
};


// ─────────────────────────────────────────────────────
//  دوال المصادقة
// ─────────────────────────────────────────────────────

function hashPassword_(password) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return digest.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function validateLogin(username, password) {
  try {
    if (!username || !password) {
      return { success: false, error: 'أدخل اسم المستخدم وكلمة المرور' };
    }
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
    if (!sheet) {
      return { success: false, error: 'لم يتم إعداد نظام المستخدمين بعد' };
    }
    var data = sheet.getDataRange().getValues();
    var inputHash = hashPassword_(password);
    var inputUser = String(username).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var storedUser = String(row[0]).trim().toLowerCase();
      var storedHash = String(row[1]).trim();
      var isActive   = row[6];
      if (storedUser === inputUser && storedHash === inputHash) {
        if (isActive !== true && String(isActive).toLowerCase() !== 'true') {
          return { success: false, error: 'الحساب معطل — تواصل مع المشرف' };
        }
        var userData = {
          username:  storedUser,
          fullName:  String(row[2]).trim(),
          airport:   String(row[3]).trim(),
          hall:      String(row[4]).trim(),
          role:      String(row[5]).trim()
        };
        var token = createSession_(userData);
        return { success: true, token: token, user: userData };
      }
    }
    return { success: false, error: 'اسم المستخدم أو كلمة المرور غير صحيحة' };
  } catch (e) {
    Logger.log('validateLogin Error: ' + e.message);
    return { success: false, error: 'خطأ في النظام — حاول مرة أخرى' };
  }
}

function createSession_(userData) {
  var token = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put(AUTH_CONFIG.SESSION_PREFIX + token, JSON.stringify(userData), AUTH_CONFIG.SESSION_DURATION);
  return token;
}

function validateSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var sessionData = cache.get(AUTH_CONFIG.SESSION_PREFIX + token);
  if (!sessionData) return null;
  try {
    var user = JSON.parse(sessionData);
    cache.put(AUTH_CONFIG.SESSION_PREFIX + token, sessionData, AUTH_CONFIG.SESSION_DURATION);
    return user;
  } catch (e) {
    return null;
  }
}

function logout(token) {
  if (!token) return;
  var cache = CacheService.getScriptCache();
  cache.remove(AUTH_CONFIG.SESSION_PREFIX + token);
}


// ─────────────────────────────────────────────────────
//  أدوات إدارة المستخدمين
// ─────────────────────────────────────────────────────

function setupUsersSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var existing = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
  if (existing) { SpreadsheetApp.getUi().alert('شيت المستخدمين موجود بالفعل'); return; }
  var sheet = ss.insertSheet(AUTH_CONFIG.USERS_SHEET);
  var headers = ['اسم_المستخدم', 'كلمة_المرور', 'الاسم_الكامل', 'المطار', 'الصالة', 'الدور', 'فعال'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#0D4A3A').setFontColor('#FFFFFF').setFontWeight('bold');
  var defaultAdmin = ['admin', hashPassword_('admin123'), 'المشرف', 'All', 'all', 'admin', true];
  sheet.getRange(2, 1, 1, headers.length).setValues([defaultAdmin]);
  sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 280); sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 80); sheet.setColumnWidth(6, 80); sheet.setColumnWidth(7, 60);
  sheet.setRightToLeft(true); sheet.setFrozenRows(1);
  var protection = sheet.protect().setDescription('بيانات المستخدمين — محمي');
  protection.setWarningOnly(false);
  SpreadsheetApp.getUi().alert('✅ تم إنشاء شيت المستخدمين\n\nالمستخدم: admin\nكلمة المرور: admin123\n\n⚠️ غيّر كلمة المرور فوراً!');
}

function addUser(username, password, fullName, airport, hall, role) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
  if (!sheet) { Logger.log('❌ شيت المستخدمين غير موجود'); return; }
  var hash = hashPassword_(password);
  sheet.appendRow([username, hash, fullName, airport, hall || 'all', role || 'staff', true]);
  Logger.log('✅ تم إضافة المستخدم: ' + username);
}

function resetUserPassword(username, newPassword) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(AUTH_CONFIG.USERS_SHEET);
  if (!sheet) { Logger.log('❌ شيت المستخدمين غير موجود'); return; }
  var data = sheet.getDataRange().getValues();
  var target = String(username).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) {
      sheet.getRange(i + 1, 2).setValue(hashPassword_(newPassword));
      Logger.log('✅ تم تحديث كلمة مرور: ' + username);
      return;
    }
  }
  Logger.log('❌ المستخدم غير موجود: ' + username);
}


// ─────────────────────────────────────────────────────
//  نقطة دخول Web App — مع المصادقة
// ─────────────────────────────────────────────────────

function doGet(e) {
  var page  = (e && e.parameter && e.parameter.page)  || 'index';
  var token = (e && e.parameter && e.parameter.token) || '';

  if (page === 'login') {
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('تسجيل الدخول — إكرام الضيف')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var session = validateSession(token);

  if (!session) {
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('تسجيل الدخول — إكرام الضيف')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'search') {
    var airport = (e.parameter.airport) || 'Madinah';
    var hall    = (e.parameter.hall)    || '';
    var template = HtmlService.createTemplateFromFile('Search');
    template.injectedAirport = airport;
    template.injectedHall    = hall;
    template.injectedToken   = token;
    template.injectedUser    = JSON.stringify(session);
    return template.evaluate()
      .setTitle('بحث رحلات المطار — إكرام الضيف')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var template = HtmlService.createTemplateFromFile('Index');
  template.injectedToken = token;
  template.injectedUser  = JSON.stringify(session);
  return template.evaluate()
    .setTitle('إكرام الضيف — اختيار المطار')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}


// ─────────────────────────────────────────────────────
//  جلب البيانات مع التخزين المؤقت
// ─────────────────────────────────────────────────────

function getAllData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CONFIG.CACHE_KEY);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  
  var chunksCount = cache.get(CONFIG.CACHE_KEY + '_CHUNKS');
  if (chunksCount) {
    var parts = []; var valid = true;
    for (var i = 0; i < parseInt(chunksCount); i++) {
      var part = cache.get(CONFIG.CACHE_KEY + '_' + i);
      if (!part) { valid = false; break; }
      parts.push(part);
    }
    if (valid) { try { return JSON.parse(parts.join('')); } catch(e) {} }
  }
  
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var cleaned = data.map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return (cell === null || cell === undefined) ? '' : String(cell).trim();
    });
  });
  
  try {
    var jsonStr = JSON.stringify(cleaned);
    var chunkSize = 90000;
    if (jsonStr.length <= chunkSize) {
      cache.put(CONFIG.CACHE_KEY, jsonStr, CONFIG.CACHE_DURATION);
    } else {
      var chunks = Math.ceil(jsonStr.length / chunkSize);
      var cacheMap = {};
      cacheMap[CONFIG.CACHE_KEY + '_CHUNKS'] = String(chunks);
      for (var i = 0; i < chunks; i++) {
        cacheMap[CONFIG.CACHE_KEY + '_' + i] = jsonStr.substring(i * chunkSize, (i + 1) * chunkSize);
      }
      cache.putAll(cacheMap, CONFIG.CACHE_DURATION);
    }
  } catch (e) { Logger.log('Cache write failed: ' + e.message); }
  
  return cleaned;
}


// ─────────────────────────────────────────────────────
//  دوال المنطق الأساسية
// ─────────────────────────────────────────────────────

function getDestinationHotel_(row) {
  var firstHouse = String(row[COL.FIRST_HOUSE]).trim().toLowerCase();
  if (firstHouse === 'madina') return row[COL.MADINAH_EN] || 'غير محدد';
  if (firstHouse === 'makkah shifting') return row[COL.MAKKAH_SHIFT_EN] || 'غير محدد';
  if (firstHouse === 'makkah') return row[COL.MAKKAH_EN] || 'غير محدد';
  return 'غير محدد';
}

function getTransportLabel_(flightType) {
  var type = String(flightType).trim().toUpperCase();
  if (type === 'B2B') return 'إكرام الضيف';
  if (type === 'B2C') return 'تحالف النقل';
  return 'غير محدد';
}

function getDestinationLabel_(firstHouse) {
  var fh = String(firstHouse).trim().toLowerCase();
  if (fh === 'madina') return 'المدينة المنورة';
  if (fh === 'makkah') return 'مكة المكرمة';
  if (fh === 'makkah shifting') return 'مكة المكرمة (تحويلي)';
  return firstHouse;
}

function getAirportLabel_(city) {
  var c = String(city).trim().toLowerCase();
  if (c === 'madinah') return 'مطار المدينة المنورة';
  if (c === 'jeddah') return 'مطار جدة';
  return city;
}

function formatTime_(timeStr) {
  if (!timeStr) return '--:--';
  var cleaned = String(timeStr).replace(/\.0+$/, '').trim();
  var parts = cleaned.split(':');
  if (parts.length < 2) return timeStr;
  var hours = parseInt(parts[0]);
  var minutes = parts[1];
  var period = hours >= 12 ? 'م' : 'ص';
  var displayHour = hours > 12 ? hours - 12 : (hours === 0 ? 12 : hours);
  return String(displayHour).padStart(2, '0') + ':' + minutes + ' ' + period;
}

function getRawTime_(timeStr) {
  if (!timeStr) return '99:99';
  var cleaned = String(timeStr).replace(/\.0+$/, '').trim();
  var parts = cleaned.split(':');
  if (parts.length < 2) return '99:99';
  return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
}


// ─────────────────────────────────────────────────────
//  فلترة مشتركة
// ─────────────────────────────────────────────────────

function filterData_(allData, params) {
  var filtered = allData;
  
  if (params.airport) {
    var targetAirport = String(params.airport).trim().toLowerCase();
    filtered = filtered.filter(function(row) {
      var city = String(row[COL.ARRIVE_CITY]).trim().toLowerCase();
      if (targetAirport === 'madinah') return city === 'madinah' || city === 'madina';
      return city === targetAirport;
    });
  }

  // ═══ فلتر الصالة — جديد ═══
  if (params.hall && params.hall !== '') {
    var hallMap = { 'hall1': '1', 'north': 'N' };
    var targetHall = hallMap[params.hall] || params.hall;
    filtered = filtered.filter(function(row) {
      return String(row[COL.HALL]).trim() === targetHall;
    });
  }

  if (params.date && params.date !== '') {
    filtered = filtered.filter(function(row) {
      return String(row[COL.ARRIVE_DATE]).substring(0, 10) === params.date;
    });
  }
  if (params.flight && params.flight !== '') {
    var searchFlight = String(params.flight).trim().toUpperCase();
    filtered = filtered.filter(function(row) {
      return String(row[COL.FLIGHT_NUMBER]).toUpperCase().indexOf(searchFlight) > -1;
    });
  }
  if (params.dest && params.dest !== '') {
    var targetDest = String(params.dest).trim().toLowerCase();
    filtered = filtered.filter(function(row) {
      var fh = String(row[COL.FIRST_HOUSE]).trim().toLowerCase();
      if (targetDest === 'madina') return fh === 'madina';
      if (targetDest === 'makkah') return fh === 'makkah' || fh === 'makkah shifting';
      return true;
    });
  }
  if (params.transport && params.transport !== '') {
    var targetTransport = String(params.transport).trim().toUpperCase();
    filtered = filtered.filter(function(row) {
      return String(row[COL.FLIGHT_TYPE]).trim().toUpperCase() === targetTransport;
    });
  }
  if (params.hotel && params.hotel !== '') {
    var targetHotel = String(params.hotel).trim();
    filtered = filtered.filter(function(row) {
      var h = String(getDestinationHotel_(row)).replace(/NULL/gi, 'غير محدد').trim();
      return h === targetHotel;
    });
  }
  if (params.airline && params.airline !== '') {
    var targetAirline = String(params.airline).trim();
    filtered = filtered.filter(function(row) {
      var a = String(row[COL.AIRLINE_AR]).trim() || String(row[COL.AIRLINE_EN]).trim();
      return a === targetAirline;
    });
  }
  if (params.departureCity && params.departureCity !== '') {
    var targetDep = String(params.departureCity).trim();
    filtered = filtered.filter(function(row) {
      return String(row[COL.DEPARTURE_CITY]).trim() === targetDep;
    });
  }
  
  return filtered;
}


// ─────────────────────────────────────────────────────
//  الدوال المكشوفة للواجهة
// ─────────────────────────────────────────────────────

function getFlightSummary(params) {
  try {
    var allData = getAllData();
    var filtered = filterData_(allData, params);
    var groups = {};
    
    filtered.forEach(function(row) {
      var flightNum = String(row[COL.FLIGHT_NUMBER]).trim();
      var destination = String(row[COL.FIRST_HOUSE]).trim();
      var hotel = String(getDestinationHotel_(row)).replace(/NULL/gi, 'غير محدد').trim();
      var transport = String(row[COL.FLIGHT_TYPE]).trim().toUpperCase();
      var dateStr = String(row[COL.ARRIVE_DATE]).substring(0, 10);
      var groupKey = flightNum + '|' + dateStr + '|' + destination + '|' + hotel;
      
      if (!groups[groupKey]) {
        groups[groupKey] = {
          groupKey: groupKey, flightNumber: flightNum,
          airline: String(row[COL.AIRLINE_EN]).trim(),
          airlineAr: String(row[COL.AIRLINE_AR]).trim(),
          date: dateStr, time: formatTime_(row[COL.ARRIVAL_TIME]),
          timeRaw: getRawTime_(row[COL.ARRIVAL_TIME]),
          destination: getDestinationLabel_(destination),
          destinationRaw: destination, hotel: hotel,
          pilgrimCount: 0, b2bCount: 0, b2cCount: 0, transport: ''
        };
      }
      groups[groupKey].pilgrimCount++;
      if (transport === 'B2B') groups[groupKey].b2bCount++;
      else groups[groupKey].b2cCount++;
    });
    
    var sortOrder = (params.sortOrder && params.sortOrder === 'desc') ? 'desc' : 'asc';
    var result = Object.keys(groups).map(function(key) {
      var g = groups[key];
      if (g.b2bCount > 0 && g.b2cCount > 0) g.transport = 'مختلط';
      else if (g.b2bCount > 0) g.transport = 'إكرام الضيف';
      else g.transport = 'تحالف النقل';
      return g;
    });
    
    result.sort(function(a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      var timeCompare = a.timeRaw < b.timeRaw ? -1 : (a.timeRaw > b.timeRaw ? 1 : 0);
      return sortOrder === 'desc' ? -timeCompare : timeCompare;
    });
    
    var totalPilgrims = result.reduce(function(sum, g) { return sum + g.pilgrimCount; }, 0);
    return { success: true, data: result, totalPilgrims: totalPilgrims, totalFlights: result.length };
  } catch (e) {
    Logger.log('getFlightSummary Error: ' + e.message);
    return { success: false, error: e.message, data: [], totalPilgrims: 0, totalFlights: 0 };
  }
}

function getFlightDetails(params) {
  try {
    var allData = getAllData();
    var pilgrims = allData.filter(function(row) {
      var city = String(row[COL.ARRIVE_CITY]).trim().toLowerCase();
      var targetAirport = String(params.airport).trim().toLowerCase();
      var cityMatch = (targetAirport === 'madinah') ? (city === 'madinah' || city === 'madina') : city === targetAirport;
      var dateMatch = String(row[COL.ARRIVE_DATE]).substring(0, 10) === params.date;
      var flightMatch = String(row[COL.FLIGHT_NUMBER]).trim() === params.flightNumber;
      var destMatch = String(row[COL.FIRST_HOUSE]).trim() === params.destination;
      var hotelMatch = String(getDestinationHotel_(row)).replace(/NULL/gi, 'غير محدد').trim() === params.hotel;
      return cityMatch && dateMatch && flightMatch && destMatch && hotelMatch;
    });
    
    var details = pilgrims.map(function(row) {
      return {
        name: String(row[COL.NAME]).trim(),
        passport: String(row[COL.PASSPORT]).trim(),
        nationality: String(row[COL.NATIONALITY_EN]).trim(),
        nationalityAr: String(row[COL.NATIONALITY_AR]).trim(),
        groupNumber: String(row[COL.GROUP_NUMBER]).trim(),
        gender: String(row[COL.GENDER]).trim(),
        transport: getTransportLabel_(row[COL.FLIGHT_TYPE]),
        transportType: String(row[COL.FLIGHT_TYPE]).trim().toUpperCase(),
        departureCity: String(row[COL.DEPARTURE_CITY]).trim(),
        residence: String(row[COL.RESIDENCE]).trim()
      };
    });
    
    details.sort(function(a, b) {
      if (a.groupNumber !== b.groupNumber) return parseInt(a.groupNumber) - parseInt(b.groupNumber);
      return a.name < b.name ? -1 : 1;
    });
    return { success: true, data: details, count: details.length };
  } catch (e) {
    Logger.log('getFlightDetails Error: ' + e.message);
    return { success: false, error: e.message, data: [], count: 0 };
  }
}

function getAvailableDates(airport, hall) {
  try {
    var allData = getAllData();
    var filtered = filterData_(allData, { airport: airport, hall: hall || '' });
    var datesSet = {};
    filtered.forEach(function(row) {
      var date = String(row[COL.ARRIVE_DATE]).substring(0, 10);
      if (date && date.length === 10) datesSet[date] = true;
    });
    return { success: true, dates: Object.keys(datesSet).sort() };
  } catch (e) { return { success: false, dates: [], error: e.message }; }
}

function getAvailableHotels(airport, hall) {
  try {
    var allData = getAllData();
    var filtered = filterData_(allData, { airport: airport, hall: hall || '' });
    var hotelsSet = {};
    filtered.forEach(function(row) {
      var hotel = String(getDestinationHotel_(row)).replace(/NULL/gi, '').trim();
      if (hotel && hotel !== 'غير محدد' && hotel !== '') hotelsSet[hotel] = true;
    });
    return { success: true, hotels: Object.keys(hotelsSet).sort() };
  } catch (e) { return { success: false, hotels: [], error: e.message }; }
}

function getFilteredOptions(params) {
  try {
    var allData = getAllData();
    var filtered = filterData_(allData, params);
    var dates = {}, hotels = {}, airlines = {}, departureCities = {};
    
    filtered.forEach(function(row) {
      var d = String(row[COL.ARRIVE_DATE]).substring(0, 10);
      if (d && d.length === 10) dates[d] = true;
      var h = String(getDestinationHotel_(row)).replace(/NULL/gi, '').trim();
      if (h && h !== 'غير محدد' && h !== '') hotels[h] = true;
      var a = String(row[COL.AIRLINE_AR]).trim() || String(row[COL.AIRLINE_EN]).trim();
      if (a && a !== '') airlines[a] = true;
      var dc = String(row[COL.DEPARTURE_CITY]).trim();
      if (dc && dc !== '') departureCities[dc] = true;
    });
    
    return { success: true, dates: Object.keys(dates).sort(), hotels: Object.keys(hotels).sort(), airlines: Object.keys(airlines).sort(), departureCities: Object.keys(departureCities).sort() };
  } catch (e) {
    return { success: false, dates: [], hotels: [], airlines: [], departureCities: [], error: e.message };
  }
}

function getQuickStats(airport, date, hall) {
  try {
    var allData = getAllData();
    var filtered = filterData_(allData, { airport: airport, date: date, hall: hall || '' });
    var stats = { totalPilgrims: filtered.length, b2bCount: 0, b2cCount: 0, makkahCount: 0, madinahCount: 0, flights: {} };
    filtered.forEach(function(row) {
      var type = String(row[COL.FLIGHT_TYPE]).trim().toUpperCase();
      if (type === 'B2B') stats.b2bCount++; else stats.b2cCount++;
      var fh = String(row[COL.FIRST_HOUSE]).trim().toLowerCase();
      if (fh === 'madina') stats.madinahCount++; else stats.makkahCount++;
      var flightKey = String(row[COL.FLIGHT_NUMBER]).trim() + '|' + String(row[COL.ARRIVE_DATE]).substring(0, 10);
      stats.flights[flightKey] = true;
    });
    stats.totalFlights = Object.keys(stats.flights).length;
    delete stats.flights;
    return { success: true, stats: stats };
  } catch (e) { return { success: false, error: e.message }; }
}

function quickSearch(query) {
  try {
    if (!query || String(query).trim().length < 3) return { success: false, error: 'أدخل 3 أحرف على الأقل' };
    var searchTerm = String(query).trim();
    var allData = getAllData();

    // تحديد نوع البحث: عربي = اسم، غير ذلك = جواز أو رحلة
    var isArabic = /[\u0600-\u06FF]/.test(searchTerm);
    var upperTerm = searchTerm.toUpperCase();

    var results = allData.filter(function(row) {
      if (isArabic) {
        return String(row[COL.NAME]).indexOf(searchTerm) > -1;
      }
      return String(row[COL.PASSPORT]).toUpperCase().indexOf(upperTerm) > -1 ||
             String(row[COL.NAME]).toUpperCase().indexOf(upperTerm) > -1;
    }).map(function(row) {
      return {
        name: String(row[COL.NAME]).trim(), passport: String(row[COL.PASSPORT]).trim(),
        nationality: String(row[COL.NATIONALITY_EN]).trim(), groupNumber: String(row[COL.GROUP_NUMBER]).trim(),
        flightNumber: String(row[COL.FLIGHT_NUMBER]).trim(), airline: String(row[COL.AIRLINE_EN]).trim(),
        arriveCity: getAirportLabel_(row[COL.ARRIVE_CITY]),
        arriveDate: String(row[COL.ARRIVE_DATE]).substring(0, 10),
        arriveTime: formatTime_(row[COL.ARRIVAL_TIME]),
        destination: getDestinationLabel_(row[COL.FIRST_HOUSE]),
        hotel: String(getDestinationHotel_(row)).replace(/NULL/gi, 'غير محدد'),
        transport: getTransportLabel_(row[COL.FLIGHT_TYPE])
      };
    });

    if (results.length > 50) results = results.slice(0, 50);
    var searchType = isArabic ? 'name' : 'passport';
    return { success: true, data: results, count: results.length, searchType: searchType };
  } catch (e) { return { success: false, error: e.message, data: [], count: 0 }; }
}

function searchByPassport(passport) {
  return quickSearch(passport);
}


// ─────────────────────────────────────────────────────
//  تصدير Excel
// ─────────────────────────────────────────────────────

function exportToExcel(params, groupBy) {
  try {
    var allData = getAllData();
    var filtered = filterData_(allData, params);
    if (filtered.length === 0) return { success: false, error: 'لا توجد بيانات للتصدير' };
    
    var airportName = params.airport === 'Jeddah' ? 'جدة' : 'المدينة';
    var fileName = 'استقبال_' + airportName + '_' + groupBy + '_' + (params.date || 'all');
    var ss = SpreadsheetApp.create(fileName);
    var defaultSheet = ss.getSheets()[0];
    var pilgrims = buildPilgrimList_(filtered);
    
    if (groupBy === 'day') buildSheetByDay_(ss, pilgrims, airportName);
    else if (groupBy === 'flight') buildSheetByFlight_(ss, pilgrims, airportName);
    else if (groupBy === 'destination') buildSheetByDestination_(ss, pilgrims, airportName);
    else if (groupBy === 'hotel') buildSheetByHotel_(ss, pilgrims, airportName);
    
    if (ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
    SpreadsheetApp.flush();
    
    var fileId = ss.getId();
    var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    var blob = response.getBlob().setName(fileName + '.xlsx');
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var downloadUrl = file.getUrl();
    DriveApp.getFileById(fileId).setTrashed(true);
    return { success: true, fileUrl: downloadUrl, fileName: fileName + '.xlsx' };
  } catch (e) {
    Logger.log('exportToExcel Error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function buildPilgrimList_(filtered) {
  return filtered.map(function(row) {
    return {
      name: String(row[COL.NAME]).trim(), passport: String(row[COL.PASSPORT]).trim(),
      nationality: String(row[COL.NATIONALITY_AR]).trim() || String(row[COL.NATIONALITY_EN]).trim(),
      gender: String(row[COL.GENDER]).trim(), groupNumber: String(row[COL.GROUP_NUMBER]).trim(),
      flightNumber: String(row[COL.FLIGHT_NUMBER]).trim(),
      airline: String(row[COL.AIRLINE_AR]).trim() || String(row[COL.AIRLINE_EN]).trim(),
      date: String(row[COL.ARRIVE_DATE]).substring(0, 10),
      time: formatTime_(row[COL.ARRIVAL_TIME]), timeRaw: getRawTime_(row[COL.ARRIVAL_TIME]),
      destination: getDestinationLabel_(row[COL.FIRST_HOUSE]),
      hotel: String(getDestinationHotel_(row)).replace(/NULL/gi, 'غير محدد').trim(),
      transport: getTransportLabel_(row[COL.FLIGHT_TYPE]),
      transportType: String(row[COL.FLIGHT_TYPE]).trim().toUpperCase(),
      departureCity: String(row[COL.DEPARTURE_CITY]).trim()
    };
  }).sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    if (a.timeRaw !== b.timeRaw) return a.timeRaw < b.timeRaw ? -1 : 1;
    return a.flightNumber < b.flightNumber ? -1 : 1;
  });
}

function getExcelHeaders_() {
  return ['#', 'الاسم', 'رقم الجواز', 'الجنسية', 'الجنس', 'المجموعة', 'رقم الرحلة', 'شركة الطيران', 'التاريخ', 'الوقت', 'الوجهة', 'الفندق', 'المواصلات', 'مدينة المغادرة'];
}

function pilgrimToRow_(p, idx) {
  return [idx, p.name, p.passport, p.nationality, p.gender, p.groupNumber, p.flightNumber, p.airline, p.date, p.time, p.destination, p.hotel, p.transport, p.departureCity];
}

function formatSheet_(sheet, headerRow, dataStartRow, dataEndRow, numCols) {
  var headerRange = sheet.getRange(headerRow, 1, 1, numCols);
  headerRange.setBackground('#0D4A3A').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
  if (dataEndRow >= dataStartRow) {
    var rowCount = dataEndRow - dataStartRow + 1;
    var dataRange = sheet.getRange(dataStartRow, 1, rowCount, numCols);
    dataRange.setBorder(true, true, true, true, true, true).setFontSize(10).setVerticalAlignment('middle');

    // ── ألوان الصفوف المتبادلة — دفعة واحدة ──
    var bgColors = [];
    for (var r = 0; r < rowCount; r++) {
      var rowBg = [];
      for (var c = 0; c < numCols; c++) rowBg.push(r % 2 === 1 ? '#F5F5F5' : '#FFFFFF');
      bgColors.push(rowBg);
    }

    // ── ألوان عمود المواصلات — دمج مع الألوان ──
    var vals = dataRange.getValues();
    var fontColors = [];
    for (var r = 0; r < rowCount; r++) {
      var rowFont = [];
      for (var c = 0; c < numCols; c++) rowFont.push(null);
      var transport = String(vals[r][12]);
      if (transport === 'إكرام الضيف') { bgColors[r][12] = '#E8F5E9'; rowFont[12] = '#2E7D32'; }
      else if (transport === 'تحالف النقل') { bgColors[r][12] = '#E3F2FD'; rowFont[12] = '#1565C0'; }
      fontColors.push(rowFont);
    }
    dataRange.setBackgrounds(bgColors);
    dataRange.setFontColors(fontColors);
  }
  sheet.setColumnWidth(1, 40); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100); sheet.setColumnWidth(5, 60); sheet.setColumnWidth(6, 70);
  sheet.setColumnWidth(7, 90); sheet.setColumnWidth(8, 100); sheet.setColumnWidth(9, 90);
  sheet.setColumnWidth(10, 80); sheet.setColumnWidth(11, 120); sheet.setColumnWidth(12, 180);
  sheet.setColumnWidth(13, 100); sheet.setColumnWidth(14, 100);
  sheet.setRightToLeft(true); sheet.setFrozenRows(headerRow);
}

function addGroupTitle_(sheet, row, title, count, numCols) {
  sheet.getRange(row, 1, 1, numCols).merge();
  var cell = sheet.getRange(row, 1);
  cell.setValue(title + '  (' + count + ' حاج)').setBackground('#C8A84E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);
}

function buildSheetByDay_(ss, pilgrims, airportName) {
  var grouped = {};
  pilgrims.forEach(function(p) { if (!grouped[p.date]) grouped[p.date] = []; grouped[p.date].push(p); });
  var dates = Object.keys(grouped).sort();
  var headers = getExcelHeaders_();
  dates.forEach(function(date) {
    var sheetName = date.substring(5);
    var sheet = ss.insertSheet(sheetName);
    var list = grouped[date]; var row = 1;
    addGroupTitle_(sheet, row, airportName + ' — ' + date, list.length, headers.length); row++;
    sheet.getRange(row, 1, 1, headers.length).setValues([headers]); row++;
    var dataStart = row;
    var batchData = list.map(function(p, idx) { return pilgrimToRow_(p, idx + 1); });
    sheet.getRange(dataStart, 1, batchData.length, headers.length).setValues(batchData);
    row += batchData.length;
    formatSheet_(sheet, dataStart - 1, dataStart, row - 1, headers.length);
  });
}

function buildSheetByFlight_(ss, pilgrims, airportName) {
  var grouped = {};
  pilgrims.forEach(function(p) { var key = p.flightNumber + ' | ' + p.date; if (!grouped[key]) grouped[key] = []; grouped[key].push(p); });
  var sheet = ss.insertSheet('حسب الرحلة');
  var headers = getExcelHeaders_(); var row = 1;
  var keys = Object.keys(grouped).sort();
  keys.forEach(function(key) {
    var list = grouped[key];
    addGroupTitle_(sheet, row, key + ' — ' + list[0].airline, list.length, headers.length); row++;
    sheet.getRange(row, 1, 1, headers.length).setValues([headers]); var headerRow = row; row++;
    var dataStart = row;
    var batchData = list.map(function(p, idx) { return pilgrimToRow_(p, idx + 1); });
    sheet.getRange(dataStart, 1, batchData.length, headers.length).setValues(batchData);
    row += batchData.length;
    formatSheet_(sheet, headerRow, dataStart, row - 1, headers.length); row++;
  });
  sheet.setRightToLeft(true);
}

function buildSheetByDestination_(ss, pilgrims, airportName) {
  var grouped = {};
  pilgrims.forEach(function(p) { if (!grouped[p.destination]) grouped[p.destination] = []; grouped[p.destination].push(p); });
  var headers = getExcelHeaders_();
  Object.keys(grouped).sort().forEach(function(dest) {
    var sheet = ss.insertSheet(dest.substring(0, 30));
    var list = grouped[dest]; var row = 1;
    addGroupTitle_(sheet, row, dest, list.length, headers.length); row++;
    sheet.getRange(row, 1, 1, headers.length).setValues([headers]); row++;
    var dataStart = row;
    var batchData = list.map(function(p, idx) { return pilgrimToRow_(p, idx + 1); });
    sheet.getRange(dataStart, 1, batchData.length, headers.length).setValues(batchData);
    row += batchData.length;
    formatSheet_(sheet, dataStart - 1, dataStart, row - 1, headers.length);
  });
}

function buildSheetByHotel_(ss, pilgrims, airportName) {
  var grouped = {};
  pilgrims.forEach(function(p) { if (!grouped[p.hotel]) grouped[p.hotel] = []; grouped[p.hotel].push(p); });
  var headers = getExcelHeaders_(); var sheetIdx = 0;
  Object.keys(grouped).sort().forEach(function(hotel) {
    sheetIdx++;
    var sheetName = String(sheetIdx) + '-' + hotel.substring(0, 28);
    var sheet = ss.insertSheet(sheetName);
    var list = grouped[hotel]; var row = 1;
    addGroupTitle_(sheet, row, hotel, list.length, headers.length); row++;
    sheet.getRange(row, 1, 1, headers.length).setValues([headers]); row++;
    var dataStart = row;
    var batchData = list.map(function(p, idx) { return pilgrimToRow_(p, idx + 1); });
    sheet.getRange(dataStart, 1, batchData.length, headers.length).setValues(batchData);
    row += batchData.length;
    formatSheet_(sheet, dataStart - 1, dataStart, row - 1, headers.length);
  });
}


// ─────────────────────────────────────────────────────
//  تريجر الإيميل
// ─────────────────────────────────────────────────────

function setupDailyEmailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { if (t.getHandlerFunction() === 'sendArrivalEmail') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendArrivalEmail').timeBased().everyDays(1).atHour(7).create();
  Logger.log('✅ تم تفعيل التريجر اليومي — 7 صباحاً');
  SpreadsheetApp.getActive().toast('تم تفعيل الإرسال اليومي الساعة 7 صباحاً', '✅ تم', 5);
}

function sendArrivalEmail() {
  try {
    var recipients = CONFIG.EMAIL_RECIPIENTS;
    if (!recipients || recipients === '') recipients = Session.getEffectiveUser().getEmail();
    var targetDate = new Date();
    targetDate.setHours(targetDate.getHours() + CONFIG.HOURS_BEFORE);
    var targetDateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var allData = getAllData();
    var filtered = allData.filter(function(row) { return String(row[COL.ARRIVE_DATE]).substring(0, 10) === targetDateStr; });
    if (filtered.length === 0) { Logger.log('لا توجد رحلات في ' + targetDateStr); return; }
    var pilgrims = buildPilgrimList_(filtered);
    var blobs = [];
    var filePrefix = 'استقبال_' + targetDateStr;
    blobs.push(createExcelBlob_(pilgrims, 'day', filePrefix + '_حسب_اليوم'));
    blobs.push(createExcelBlob_(pilgrims, 'flight', filePrefix + '_حسب_الرحلة'));
    blobs.push(createExcelBlob_(pilgrims, 'destination', filePrefix + '_حسب_الوجهة'));
    blobs.push(createExcelBlob_(pilgrims, 'hotel', filePrefix + '_حسب_الفندق'));
    
    var flights = {}; var b2bCount = 0, b2cCount = 0, makkahCount = 0, madinahCount = 0, jeddahCount = 0, madinahAirportCount = 0;
    filtered.forEach(function(row) {
      flights[String(row[COL.FLIGHT_NUMBER]).trim()] = true;
      var type = String(row[COL.FLIGHT_TYPE]).trim().toUpperCase();
      if (type === 'B2B') b2bCount++; else b2cCount++;
      var fh = String(row[COL.FIRST_HOUSE]).trim().toLowerCase();
      if (fh === 'madina') madinahCount++; else makkahCount++;
      var city = String(row[COL.ARRIVE_CITY]).trim().toLowerCase();
      if (city === 'jeddah') jeddahCount++; else madinahAirportCount++;
    });
    
    var subject = '✈️ استقبال ' + targetDateStr + ' — ' + filtered.length + ' حاج / ' + Object.keys(flights).length + ' رحلة';
    var body = '<div dir="rtl" style="font-family:Tahoma,Arial;font-size:14px;color:#333;">';
    body += '<h2 style="color:#0D4A3A;border-bottom:3px solid #C8A84E;padding-bottom:8px;">تقرير الاستقبال — ' + targetDateStr + '</h2>';
    body += '<table style="border-collapse:collapse;width:100%;max-width:500px;margin:16px 0;">';
    body += '<tr><td style="padding:8px;background:#0D4A3A;color:#fff;font-weight:bold;">إجمالي الحجاج</td><td style="padding:8px;background:#E8F5E9;font-weight:bold;font-size:18px;text-align:center;">' + filtered.length + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#f5f5f5;">عدد الرحلات</td><td style="padding:8px;text-align:center;">' + Object.keys(flights).length + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#f5f5f5;">مطار جدة</td><td style="padding:8px;text-align:center;">' + jeddahCount + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#f5f5f5;">مطار المدينة</td><td style="padding:8px;text-align:center;">' + madinahAirportCount + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#E8F5E9;color:#2E7D32;">إكرام الضيف (B2B)</td><td style="padding:8px;text-align:center;">' + b2bCount + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#E3F2FD;color:#1565C0;">تحالف النقل (B2C)</td><td style="padding:8px;text-align:center;">' + b2cCount + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#f5f5f5;">وجهة مكة</td><td style="padding:8px;text-align:center;">' + makkahCount + '</td></tr>';
    body += '<tr><td style="padding:8px;background:#f5f5f5;">وجهة المدينة</td><td style="padding:8px;text-align:center;">' + madinahCount + '</td></tr>';
    body += '</table>';
    body += '<p style="color:#888;font-size:12px;">مرفق 4 ملفات Excel: حسب اليوم / الرحلة / الوجهة / الفندق</p>';
    body += '<p style="color:#C8A84E;font-size:11px;">شركة إكرام الضيف للسياحة — موسم حج 1447</p>';
    body += '</div>';
    
    MailApp.sendEmail({ to: recipients, subject: subject, htmlBody: body, attachments: blobs });
    Logger.log('✅ تم إرسال الإيميل إلى: ' + recipients + ' — ' + filtered.length + ' حاج');
  } catch (e) { Logger.log('sendArrivalEmail Error: ' + e.message); }
}

function createExcelBlob_(pilgrims, groupBy, fileName) {
  var ss = SpreadsheetApp.create(fileName + '_temp');
  var defaultSheet = ss.getSheets()[0];
  var airportName = 'إكرام الضيف';
  if (groupBy === 'day') buildSheetByDay_(ss, pilgrims, airportName);
  else if (groupBy === 'flight') buildSheetByFlight_(ss, pilgrims, airportName);
  else if (groupBy === 'destination') buildSheetByDestination_(ss, pilgrims, airportName);
  else if (groupBy === 'hotel') buildSheetByHotel_(ss, pilgrims, airportName);
  if (ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  SpreadsheetApp.flush();
  var fileId = ss.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  var blob = response.getBlob().setName(fileName + '.xlsx');
  DriveApp.getFileById(fileId).setTrashed(true);
  return blob;
}

function clearCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY); cache.remove(CONFIG.CACHE_KEY + '_CHUNKS');
  for (var i = 0; i < 20; i++) cache.remove(CONFIG.CACHE_KEY + '_' + i);
  return 'Cache cleared';
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('✈️ بحث المطار')
    .addItem('👤 إعداد شيت المستخدمين', 'setupUsersSheet')
    .addSeparator()
    .addItem('📧 تفعيل الإرسال اليومي', 'setupDailyEmailTrigger')
    .addItem('📧 إرسال تجريبي الآن', 'sendArrivalEmail')
    .addItem('🗑️ مسح الكاش', 'clearCache')
    .addToUi();
}
// ═══════════════════════════════════════════════════════
//  مسح جواز السفر — Google Drive OCR
// ═══════════════════════════════════════════════════════

function scanPassportMRZ(base64Data) {
  try {
    if (!base64Data) return { success: false, error: 'لم يتم استلام صورة' };

    // ── تحويل Base64 إلى Blob ──
    var parts = base64Data.split(',');
    var raw = parts.length > 1 ? parts[1] : parts[0];
    var decoded = Utilities.base64Decode(raw);
    var blob = Utilities.newBlob(decoded, 'image/jpeg', 'passport_scan.jpg');

    // ── رفع لـ Drive مع OCR ──
    var file = Drive.Files.insert(
      { title: 'MRZ_SCAN_' + new Date().getTime(), mimeType: 'image/jpeg' },
      blob,
      { ocr: true, ocrLanguage: 'en' }
    );

    // ── قراءة النص ──
    var doc = DocumentApp.openById(file.id);
    var text = doc.getBody().getText();

    // ── حذف الملف فوراً ──
    DriveApp.getFileById(file.id).setTrashed(true);

    if (!text || text.trim().length < 10) {
      return { success: false, error: 'لم يتم التعرف على نص في الصورة' };
    }

    // ── استخراج من MRZ ──
    var passport = extractPassportFromMRZ_(text);
    if (passport) return { success: true, passport: passport };

    // ── محاولة بديلة ──
    var fallback = extractPassportFallback_(text);
    if (fallback) return { success: true, passport: fallback };

    return { success: false, error: 'لم يتم التعرف على رقم الجواز — أدخله يدوياً' };

  } catch (e) {
    Logger.log('scanPassportMRZ Error: ' + e.message);
    return { success: false, error: 'خطأ في المعالجة: ' + e.message };
  }
}

function extractPassportFromMRZ_(text) {
  var lines = text.split(/\n/).map(function(l) { return l.replace(/\s/g, ''); });

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/[^A-Z0-9<]/gi, '');
    if (line.length >= 30 && /^P/.test(line)) {
      if (i + 1 < lines.length) {
        var next = lines[i + 1].replace(/[^A-Z0-9<]/gi, '');
        if (next.length >= 28) {
          var pNum = next.substring(0, 9).replace(/<+$/g, '');
          if (pNum.length >= 5) return pNum;
        }
      }
      break;
    }
  }

  // بحث بديل — سطر 44+ حرف
  for (var j = 0; j < lines.length; j++) {
    var cleaned = lines[j].replace(/[^A-Z0-9<]/gi, '');
    if (cleaned.length >= 44 && cleaned.indexOf('<') > -1 && j + 1 < lines.length) {
      var nextLine = lines[j + 1].replace(/[^A-Z0-9<]/gi, '');
      if (nextLine.length >= 28) {
        var num = nextLine.substring(0, 9).replace(/<+$/g, '');
        if (num.length >= 5) return num;
      }
    }
  }
  return null;
}

function extractPassportFallback_(text) {
  var patterns = [
    /passport\s*(?:no|number|#)?\s*[:\-]?\s*([A-Z0-9]{6,9})/i,
    /\b([A-Z]{1,2}\d{6,8})\b/,
    /\b(\d{9})\b/
  ];
  for (var i = 0; i < patterns.length; i++) {
    var match = text.match(patterns[i]);
    if (match && match[1]) return match[1].toUpperCase();
  }
  return null;
}