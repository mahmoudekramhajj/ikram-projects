// ============================================================
// CACHE: CacheService للشيتات الصغيرة — TTL 2 دقيقة
// يقلل عدد القراءات من الشيت عند الاستدعاءات المتكررة
// ============================================================

var CACHE_TTL = 120; // ثانية

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
    if (json.length < 100000) { // حد CacheService = 100KB
      cache.put(cacheKey, json, CACHE_TTL);
    }
  } catch (e) {}
}

// ============================================================
// HELPER: findSheet_ مع trim — بحث آمن عن الشيت
// يعالج مشكلة المسافات الزائدة في أسماء الشيتات
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
// HELPER: Match pilgrim to hotel — unified Makkah
// يفحص عمودي مكة + مكة تحويل معاً عند طلب مكة
// يُعيد المدينة الفعلية للحاج أو null إذا لا يتبع هذا الفندق
// ============================================================

function getActualHotelCity_(row, hotelName, requestedCity) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  
  if (requestedCity === 'Madina') {
    return row[J.MADINAH_EN] === hotelName ? 'Madina' : null;
  }
  
  if (requestedCity === 'Makkah') {
    if (row[J.MAKKAH_EN] === hotelName) return 'Makkah';
    if (row[J.MAKKAH_SHIFT_EN] === hotelName) return 'Makkah Shifting';
    return null;
  }
  
  return null;
}
/**
 * Hotel Management App — Helper Functions
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v3.1 — Hotel sheets, helpers, room ID generation, guide map, formatters
 */


// ============================================================
// HOTEL SHEET: Get or Create (Individual per hotel)
// ============================================================

function getHotelSheet_(ss, hotelName) {
  var sheetName = sanitizeSheetName_(hotelName);
  return ss.getSheetByName(sheetName);
}
// ============================================================
// MIGRATION: إعادة تسمية شيتات الفنادق من الاسم الكامل → الاختصار
// يُشغَّل مرة واحدة فقط ثم يُحذف أو يُترك
// ============================================================

function migrateHotelSheetNames() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheets = ss.getSheets();
  var abbr = HOTEL_CONFIG.HOTEL_ABBR;
  var log = [];
  
  // بناء خريطة عكسية: اختصار → اسم كامل (لتجنب التكرار)
  var abbrExists = {};
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    // هل هذا الشيت بالفعل اختصار؟
    for (var fullName in abbr) {
      if (abbr[fullName] === name) abbrExists[name] = true;
    }
  }
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var oldName = sheet.getName();
    
    // تحقق: هل الاسم الحالي يطابق اسم فندق كامل أو مقصوص؟
    for (var fullName in abbr) {
      var newName = abbr[fullName];
      
      // تطابق كامل أو تطابق مقصوص (الاسم الحالي هو بداية الاسم الكامل)
      if (oldName === fullName || 
          (fullName.indexOf(oldName) === 0 && oldName.length >= 20)) {
        
        // لا نعيد التسمية إذا الاختصار موجود بالفعل كشيت آخر
        if (abbrExists[newName] && oldName !== newName) {
          // شيت مكرر — نحذفه إذا فارغ (هيدر فقط)
          if (sheet.getLastRow() <= 1) {
            ss.deleteSheet(sheet);
            log.push('🗑️ حُذف (مكرر فارغ): ' + oldName);
          } else {
            log.push('⚠️ مكرر غير فارغ: ' + oldName + ' → ' + newName + ' موجود');
          }
        } else {
          sheet.setName(newName);
          abbrExists[newName] = true;
          log.push('✅ ' + oldName + ' → ' + newName);
        }
        break;
      }
    }
  }
  
  Logger.log('=== نتيجة الهجرة ===\n' + log.join('\n'));
  return log;
}
function getOrCreateHotelSheet_(ss, hotelName) {
  var sheetName = sanitizeSheetName_(hotelName);
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(HOTEL_CONFIG.HOTEL_SHEET_HEADERS);
    sheet.getRange(1, 1, 1, HOTEL_CONFIG.HOTEL_SHEET_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

function sanitizeSheetName_(hotelName) {
  if (!hotelName) return 'Unknown';
  // استخدام الاختصار من HOTEL_ABBR كاسم التبويب
  var abbr = HOTEL_CONFIG.HOTEL_ABBR[hotelName];
  if (abbr) return abbr;
  // Fallback: أول 3 حروف كبيرة (لفنادق غير معرّفة)
  Logger.log('⚠️ فندق بدون اختصار: ' + hotelName);
  return String(hotelName).substring(0, 20).replace(/[\/\\?*\[\]]/g, '_');
}

// ============================================================
// HOTEL SHEET: Populate with pilgrim data
// ============================================================

function populateHotelSheet(hotelName, hotelCity) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = getOrCreateHotelSheet_(ss, hotelName);
  
  var pilgrims = getPilgrimsForHotel(hotelName, hotelCity);
  var data = sheet.getDataRange().getValues();
  
  // Build existing booking IDs set
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    existing[String(data[i][0])] = i + 1; // row number
  }
  
  var newRows = [];
  
  for (var p = 0; p < pilgrims.length; p++) {
    var pg = pilgrims[p];
    var rowData = [
      pg.bookingId, pg.name, pg.passport, pg.gender, pg.nationality,
      pg.groupNumber, pg.packageName, pg.hotelCity, pg.phase,
      pg.hotelCheckIn, pg.hotelCheckOut,
      pg.arrivalFlight, pg.arrivalDate, pg.expectedArrival,
      pg.arrivalStatus, pg.roomType, '',
      '', 'pending', '',
      pg.returnFlight, pg.returnDate, pg.transport
    ];
    
    if (existing[String(pg.bookingId)]) {
      // Update only static fields (don't overwrite RoomGroup, Room#, Status, Time)
      var rowNum = existing[String(pg.bookingId)];
      sheet.getRange(rowNum, 1, 1, 16).setValues([rowData.slice(0, 16)]);
      sheet.getRange(rowNum, 21, 1, 3).setValues([rowData.slice(20, 23)]);
    } else {
      newRows.push(rowData);
    }
  }
  
  // Batch append all new rows at once
  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  
  SpreadsheetApp.flush();
  return { success: true, total: pilgrims.length };
}

// ============================================================
// HOTEL SHEET: Get Check-in Data
// ============================================================

function getCheckinData_(ss, hotelName) {
  var sheet = getHotelSheet_(ss, hotelName);
  if (!sheet) return {};
  
  var data = sheet.getDataRange().getValues();
  var map = {};
  
  for (var i = 1; i < data.length; i++) {
    var bookingId = String(data[i][0]);
    map[bookingId] = {
      roomGroupId: data[i][16] || '',   // RoomGroup_ID (Q)
      roomNumber: data[i][17] || '',     // رقم الغرفة (R)
      status: data[i][18] || 'pending',  // حالة Check-in (S)
      checkInTime: data[i][19] || ''     // وقت Check-in (T)
    };
  }
  
  return map;
}

// ============================================================
// HELPER: Get Room Capacity
// ============================================================

function getRoomCapacity_(roomType) {
  if (!roomType) return 4;
  var rt = String(roomType).toLowerCase();
  if (rt.indexOf('quad') >= 0 || rt.indexOf('رباعي') >= 0 || rt.indexOf('4') >= 0) return 4;
  if (rt.indexOf('tri') >= 0 || rt.indexOf('ثلاثي') >= 0 || rt.indexOf('3') >= 0) return 3;
  if (rt.indexOf('dbl') >= 0 || rt.indexOf('double') >= 0 || rt.indexOf('ثنائي') >= 0 || rt.indexOf('2') >= 0) return 2;
  return 4;
}

// ============================================================
// HELPER: Get Hotel Name from Journey Row
// ============================================================

function getHotelName_(row, hotelCity) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  if (hotelCity === 'Madina') return row[J.MADINAH_EN];
  if (hotelCity === 'Makkah') return row[J.MAKKAH_EN];
  if (hotelCity === 'Makkah Shifting') return row[J.MAKKAH_SHIFT_EN];
  return null;
}

// ============================================================
// HELPER: Get Hotel Dates
// ============================================================

function getHotelDates_(row, hotelCity) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var firstHouse = row[J.FIRST_HOUSE];
  var lastHouse = row[J.LAST_HOUSE];
  
  if (hotelCity === firstHouse || 
      (hotelCity === 'Madina' && firstHouse === 'Madina') ||
      (hotelCity === 'Makkah' && firstHouse === 'Makkah') ||
      (hotelCity === 'Makkah Shifting' && firstHouse === 'Makkah Shifting')) {
    return {
      checkIn: formatDate_(row[J.FIRST_HOUSE_START]),
      checkOut: formatDate_(row[J.FIRST_HOUSE_END]),
      phase: 'first'
    };
  }
  
  if (hotelCity === lastHouse || 
      (hotelCity === 'Madina' && lastHouse === 'Madina') ||
      (hotelCity === 'Makkah' && lastHouse === 'Makkah') ||
      (hotelCity === 'Makkah Shifting' && lastHouse === 'Makkah Shifting')) {
    return {
      checkIn: formatDate_(row[J.LAST_HOUSE_START]),
      checkOut: formatDate_(row[J.LAST_HOUSE_END]),
      phase: 'last'
    };
  }
  
  return {
    checkIn: formatDate_(row[J.FIRST_HOUSE_END]),
    checkOut: formatDate_(row[J.LAST_HOUSE_START]),
    phase: 'middle'
  };
}

// ============================================================
// HELPER: Calculate Expected Arrival — 11 سيناريو وصول
// Phase first → وصول من المطار (5 حالات: مطار_مدينة الفندق)
// Phase last  → انتقال بين المدن أو تحويل (6 حالات)
// Phase middle → تحويل فندقي (وقت check-in 13:00)
// ============================================================

function calculateExpectedArrival_(row, hotelCity) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var hotelDates = getHotelDates_(row, hotelCity);

  // ===== المرحلة الأولى: وصول من المطار =====
  if (hotelDates.phase === 'first') {
    var arrivalCity = row[J.ARRIVAL_CITY];
    var arrivalTime = row[J.ARRIVAL_TIME];
    var firstHouse = row[J.FIRST_HOUSE];

    var marginKey = arrivalCity + '_' + firstHouse;
    var margin = HOTEL_CONFIG.MARGINS.ARRIVAL[marginKey];

    if (!margin || !arrivalTime) return null;
    return addHoursToTime_(formatTime_(arrivalTime), margin);
  }

  // ===== المرحلة الأخيرة: انتقال من فندق آخر =====
  if (hotelDates.phase === 'last') {
    // الحاج قادم من المدينة الأخرى — الوصول ≈ وقت check-in الفندق
    return HOTEL_CONFIG.CHECKIN_TIME;
  }

  // ===== المرحلة الوسطى: تحويل فندقي =====
  if (hotelDates.phase === 'middle') {
    return HOTEL_CONFIG.CHECKIN_TIME;
  }

  return null;
}

function addHoursToTime_(timeStr, hours) {
  if (!timeStr) return null;
  var parts = timeStr.split(':');
  var h = parseInt(parts[0]);
  var m = parseInt(parts[1]) || 0;

  h += Math.floor(hours);
  m += (hours % 1) * 60;

  if (m >= 60) { h += Math.floor(m / 60); m = m % 60; }
  if (h >= 24) h -= 24;

  return String(h).padStart(2, '0') + ':' + String(Math.round(m)).padStart(2, '0');
}

// ============================================================
// HELPER: Check Arrival Status (Early/Late/Normal/NoShow)
// ============================================================

function checkArrivalStatus_(row, hotelCity, expectedArrival) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var firstHouse = row[J.FIRST_HOUSE];
  var arrivalDate = row[J.ARRIVAL_DATE];
  var contractStart = row[J.FIRST_HOUSE_START];
  var contractEnd = row[J.FIRST_HOUSE_END];
  
  var result = { status: 'normal', earlyHours: 0, lateDays: 0 };
  
  var hotelDates = getHotelDates_(row, hotelCity);
  if (hotelDates.phase !== 'first') return result;
  
  if (!arrivalDate || !contractStart) return result;
  
  var arrival = new Date(arrivalDate);
  var start = new Date(contractStart);
  arrival.setHours(0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);
  
  // --- NO SHOW: Madinah only — arrival after contract END ---
  if (hotelCity === 'Madina' && firstHouse === 'Madina' && contractEnd) {
    var end = new Date(contractEnd);
    end.setHours(0, 0, 0, 0);
    if (arrival > end) {
      result.status = 'noshow';
      return result;
    }
  }
  
  // --- LATE: Arrival date AFTER contract start ---
  var diffDays = Math.round((arrival - start) / (1000 * 60 * 60 * 24));
  if (diffDays > 0) {
    result.status = 'late';
    result.lateDays = diffDays;
    return result;
  }
  
  // --- EARLY: Same day + expected arrival before 13:00 (Madinah only) ---
  if (diffDays === 0 && hotelCity === 'Madina' && firstHouse === 'Madina' && expectedArrival) {
    var parts = expectedArrival.split(':');
    var expHour = parseInt(parts[0]);
    var expMin = parseInt(parts[1]) || 0;
    
    if (expHour < 13) {
      result.status = 'early';
      result.earlyHours = 13 - expHour;
      if (expMin > 0) result.earlyHours--;
      if (result.earlyHours < 1) result.earlyHours = 1;
      return result;
    }
  }
  
  return result;
}

// ============================================================
// HELPER: Calculate Departure — 13 سيناريو مغادرة
// Phase last  → مطار (رحلة عودة)
// Phase first → انتقال بين المدن (حافلة / قطار)
// Phase middle → تحويل فندقي (نفس المدينة)
// ============================================================

function calculateDeparture_(row, hotelCity, hotelDates, transport) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;

  var departure = {
    destination: '',
    destinationHotel: '',
    transport: transport,
    time: null,
    type: 'normal',
    linkedFlight: null,
    linkedTime: null
  };

  // ===== المرحلة الأخيرة: مغادرة للمطار (رحلة العودة) =====
  // المدينة → مطار المدينة | المدينة → مطار جدة
  // مكة → مطار جدة | مكة → مطار المدينة
  if (hotelDates.phase === 'last') {
    var returnCity = row[J.RETURN_DEPT_CITY] || 'Jeddah';
    departure.destination = 'مطار ' + (returnCity === 'Madinah' ? 'المدينة' : 'جدة');
    departure.linkedFlight = row[J.RETURN_FLIGHT];
    departure.linkedTime = formatTime_(row[J.RETURN_DEPT_TIME]);
    departure.transport = 'حافلة'; // المطار دائماً حافلة

    var marginKey = getCityCode_(hotelCity) + '_' + returnCity;
    var margin = HOTEL_CONFIG.MARGINS.DEPARTURE[marginKey] || 8;
    departure.time = subtractHours_(departure.linkedTime, margin);

  // ===== المرحلة الأولى: انتقال بين المدن =====
  // المدينة → فندق مكة (حافلة/قطار)
  // مكة → فندق المدينة (حافلة/قطار)
  } else if (hotelDates.phase === 'first') {
    var nextHotel = getNextHotelName_(row, hotelCity);
    departure.destinationHotel = nextHotel;

    if (getCityCode_(hotelCity) === 'Madina') {
      departure.destination = 'مكة المكرمة';
    } else {
      departure.destination = 'المدينة المنورة';
    }

    // تحديد وسيلة النقل: قطار أو حافلة
    var isTrain = transport && (
      String(transport).indexOf('قطار') >= 0 ||
      String(transport).toLowerCase().indexOf('train') >= 0
    );

    if (isTrain) {
      departure.transport = 'قطار';
      departure.time = null; // موعد غير محدد — قطار
    } else {
      departure.transport = 'حافلة';
      var interMargin = HOTEL_CONFIG.MARGINS.INTERCITY['bus'] || 7;
      departure.time = HOTEL_CONFIG.CHECKOUT_TIME;
    }

  // ===== المرحلة الوسطى: تحويل فندقي (مكة1 ← مكة2) =====
  } else {
    var nextHotel = getNextHotelName_(row, hotelCity);
    departure.destinationHotel = nextHotel;
    departure.destination = 'تحويل فندقي';
    departure.transport = 'حافلة';
    departure.time = HOTEL_CONFIG.CHECKOUT_TIME;
    departure.type = 'shifting';
  }

  // تحديد نوع المغادرة حسب الوقت
  if (departure.time && departure.type !== 'shifting') {
    var deptHour = parseInt(departure.time.split(':')[0]);
    if (deptHour < 6) departure.type = 'early';
  }

  return departure;
}

// ============================================================
// HELPER: Get Transport Map
// ============================================================

function getTransportMap_(ss) {
  var cached = getCachedData_('transportMap');
  if (cached) return cached;

  var sheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.PACKAGES);
  var data = sheet.getDataRange().getValues();
  var map = {};

  for (var i = 2; i < data.length; i++) {
    var ikramNo = data[i][4];
    var transport = data[i][65];
    if (ikramNo && transport) {
      map[ikramNo] = transport;
    }
  }

  setCachedData_('transportMap', map);
  return map;
}

// ============================================================
// HELPER: Room Type & Shared Beds
// ============================================================

function getRoomType_(groupInfo, hotelCity) {
  if (hotelCity === 'Madina') return groupInfo.roomTypeMed || '';
  if (hotelCity === 'Makkah') return groupInfo.roomTypeMak1 || '';
  if (hotelCity === 'Makkah Shifting') return groupInfo.roomTypeMak2 || '';
  return '';
}

function getSharedBeds_(groupInfo, hotelCity) {
  if (hotelCity === 'Madina') return groupInfo.sharedBedsMed || 0;
  if (hotelCity === 'Makkah') return groupInfo.sharedBedsMak1 || 0;
  if (hotelCity === 'Makkah Shifting') return groupInfo.sharedBedsMak2 || 0;
  return 0;
}

function getFullRooms_(groupInfo, hotelCity) {
  if (hotelCity === 'Madina') return groupInfo.fullRoomsMed || 0;
  if (hotelCity === 'Makkah') return groupInfo.fullRoomsMak1 || 0;
  if (hotelCity === 'Makkah Shifting') return groupInfo.fullRoomsMak2 || 0;
  return 0;
}

// ============================================================
// INTERNAL ROOM ID GENERATION (v3 — replaces RG-timestamp)
// Format: [City 1char][Hotel 3chars][Type 1char][Seq 3digits]
// Example: MCRWQ001 = Madinah, Crowne Plaza, Quad, Room #1
// ============================================================

function getHotelAbbr_(hotelName) {
  // Exact match first
  if (HOTEL_CONFIG.HOTEL_ABBR[hotelName]) return HOTEL_CONFIG.HOTEL_ABBR[hotelName];
  
  // Fuzzy match — normalize and compare
  var norm = normalizeHotelName_(hotelName);
  for (var key in HOTEL_CONFIG.HOTEL_ABBR) {
    if (normalizeHotelName_(key) === norm) return HOTEL_CONFIG.HOTEL_ABBR[key];
  }
  
  // Fallback — first 3 uppercase letters
  var letters = hotelName.replace(/[^A-Za-z]/g, '').toUpperCase();
  return letters.substring(0, 3) || 'UNK';
}

function getCityPrefix_(hotelCity) {
  if (hotelCity === 'Madina') return 'M';
  return 'K'; // Makkah + Makkah Shifting = same physical hotel
}

function getRoomTypeCode_(roomType) {
  var cap = getRoomCapacity_(roomType);
  if (cap === 4) return 'Q';
  if (cap === 3) return 'T';
  if (cap === 2) return 'D';
  return 'Q'; // default
}

function generateInternalRoomId_(hotelName, hotelCity, roomType, existingIds) {
  var city = getCityPrefix_(hotelCity);
  var abbr = getHotelAbbr_(hotelName);
  var type = getRoomTypeCode_(roomType);
  var prefix = city + abbr + type;
  
  // Find next available sequence
  var maxSeq = 0;
  if (existingIds && existingIds.length > 0) {
    for (var i = 0; i < existingIds.length; i++) {
      var id = String(existingIds[i]);
      if (id.indexOf(prefix) === 0) {
        var seqStr = id.substring(prefix.length);
        var seq = parseInt(seqStr) || 0;
        if (seq > maxSeq) maxSeq = seq;
      }
    }
  }
  
  var nextSeq = String(maxSeq + 1).padStart(3, '0');
  return prefix + nextSeq;
}

// Legacy wrapper — for manual grouping compatibility
function generateRoomGroupId() {
  var now = new Date();
  var ts = Utilities.formatDate(now, 'Asia/Riyadh', 'yyMMddHHmmssSSS');
  return 'RG-' + ts + '-' + Math.floor(Math.random() * 10000);
}

// ============================================================
// HELPER: Format Utilities
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
  var str = String(timeVal);
  if (str.includes(':')) {
    var parts = str.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
  }
  if (timeVal instanceof Date) {
    return Utilities.formatDate(timeVal, 'Asia/Riyadh', 'HH:mm');
  }
  return str;
}

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

function getCityCode_(hotelCity) {
  if (hotelCity === 'Madina') return 'Madina';
  if (hotelCity === 'Makkah' || hotelCity === 'Makkah Shifting') return 'Makkah';
  return hotelCity;
}

function normalizeHotelName_(name) {
  if (!name) return '';
  return String(name).toLowerCase()
    .replace(/[-_]/g, '')
    .replace(/\s+/g, '')
    .replace(/hotel|company|co|ltd|madinah|madina|makkah|mecca|medina/gi, '');
}

// ============================================================
// GUIDE MAP: Build from Tour Guide sheet
// Returns: { passport → guideName } for Registered + Unique only
// ============================================================

function buildGuideMap_(ss) {
  var cached = getCachedData_('guideMap');
  if (cached) return cached;

  var map = {};
  try {
    var sheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.TOUR_GUIDE);
    if (!sheet) return map;
    
    var lastRow = sheet.getLastRow();
    var startRow = HOTEL_CONFIG.GUIDE_START_ROW || 2;
    if (lastRow < startRow) return map;
    
    var numRows = lastRow - startRow + 1;
    var data = sheet.getRange(startRow, 1, numRows, 11).getValues(); // A-K
    var G = HOTEL_CONFIG.GUIDE_COLS;
    
    for (var i = 0; i < data.length; i++) {
      var guideName = String(data[i][G.GUIDE_NAME]).trim();
      var passport = String(data[i][G.PASSPORT]).trim().toUpperCase();
      var regStatus = String(data[i][G.REG_STATUS]).trim();
      var dupCheck = String(data[i][G.UNIQUE_CHECK]).trim();
      
      if (guideName && passport
          && regStatus.indexOf('Registered') !== -1
          && dupCheck.indexOf('Unique') !== -1) {
        map[passport] = guideName;
      }
    }
  } catch (e) {
    Logger.log('buildGuideMap_ ERROR: ' + e.toString());
  }
  setCachedData_('guideMap', map);
  return map;
}

// ============================================================
// GUIDE MAP: Get guide name for a pilgrim by passport
// ============================================================

function getGuideForPilgrim_(guideMap, passport) {
  if (!guideMap || !passport) return '';
  var key = String(passport).trim().toUpperCase();
  return guideMap[key] || '';
}