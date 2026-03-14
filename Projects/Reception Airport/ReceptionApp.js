/**
 * =====================================================
 *  إكرام الضيف — تأكيد استقبال الحجاج بالمطار
 *  Airport Reception App — ReceptionApp.gs (Backend)
 *  Version: 1.0
 * =====================================================
 */

// ─────────────────────────────────────────────────────
//  ثوابت المشروع
// ─────────────────────────────────────────────────────

const CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'رحلة الحاج ',  // لاحظ المسافة بعد "الحاج"
  CACHE_DURATION: 300,         // 5 دقائق
  CACHE_KEY: 'RECEPTION_APP_DATA',
  DELAY_ALERT_MINUTES: 30     // تنبيه التأخر بعد 30 دقيقة
};

const COL = {
  BOOKING_ID:        0,   // A
  PACKAGE_ID:        1,   // B
  GROUP_NUMBER:      6,   // G
  NAME:              7,   // H
  PASSPORT:          8,   // I
  GENDER:           11,   // L
  NATIONALITY_EN:   12,   // M
  NATIONALITY_AR:   13,   // N
  AIRLINE_AR:       16,   // Q
  AIRLINE_EN:       17,   // R
  ARRIVAL_TIME:     18,   // S
  ARRIVE_CITY:      19,   // T
  ARRIVE_DATE:      20,   // U
  DEPARTURE_CITY:   21,   // V
  FLIGHT_NUMBER:    24,   // Y
  FLIGHT_TYPE:      25,   // Z
  FIRST_HOUSE:      36,   // AK
  MAKKAH_EN:        43,   // AR
  MAKKAH_SHIFT_EN:  45,   // AT
  MADINAH_EN:       47,   // AV
  HALL:             48,   // AW — الصالة/الترمنال
  RECEPTION_STATUS: 49,   // AX — حالة الاستقبال
  RECEPTION_TIME:   50,   // AY — وقت التأكيد
  RECEPTION_STAFF:  51    // AZ — اسم الموظف
};


// ─────────────────────────────────────────────────────
//  نقطة دخول Web App
// ─────────────────────────────────────────────────────

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'index';

  if (page === 'main') {
    var t = HtmlService.createTemplateFromFile('ReceptionMain');
    return t.evaluate()
      .setTitle('استقبال الحجاج — إكرام الضيف')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('ReceptionIndex')
    .setTitle('إكرام الضيف — تسجيل دخول الاستقبال')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/** إدراج ملفات HTML الفرعية */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ─────────────────────────────────────────────────────
//  جلب البيانات مع التخزين المؤقت
// ─────────────────────────────────────────────────────

function getAllData() {
  var cache = CacheService.getScriptCache();

  // محاولة قراءة من الكاش (قطعة واحدة)
  var cached = cache.get(CONFIG.CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  // كاش مقسّم
  var chunksCount = cache.get(CONFIG.CACHE_KEY + '_CHUNKS');
  if (chunksCount) {
    var parts = [];
    var valid = true;
    for (var i = 0; i < parseInt(chunksCount); i++) {
      var part = cache.get(CONFIG.CACHE_KEY + '_' + i);
      if (!part) { valid = false; break; }
      parts.push(part);
    }
    if (valid) {
      try { return JSON.parse(parts.join('')); } catch(e) {}
    }
  }

  // قراءة من الشيت
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), COL.RECEPTION_STAFF + 1);

  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var cleaned = data.map(function(row, idx) {
    var mapped = row.map(function(cell) {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return (cell === null || cell === undefined) ? '' : String(cell).trim();
    });
    // إضافة رقم الصف الفعلي في الشيت (1-based، بعد الهيدر)
    mapped.push(idx + 2);
    return mapped;
  });

  // تخزين مؤقت
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
  } catch (e) {
    Logger.log('Cache write failed: ' + e.message);
  }

  return cleaned;
}


// ─────────────────────────────────────────────────────
//  البحث بالجواز
// ─────────────────────────────────────────────────────

/**
 * searchByPassport — البحث عن حاج برقم الجواز
 * @param {string} passport — رقم الجواز (كامل أو جزئي — 3 أحرف كحد أدنى)
 * @param {string} airport — اسم المطار (jeddah/madinah)
 * @param {string} hall — الصالة (للفلترة)
 * @returns {Object} — بيانات الحاج أو خطأ
 */
function searchByPassport(passport, airport, hall) {
  try {
    if (!passport || passport.length < 3) {
      return { success: false, error: 'أدخل 3 أحرف على الأقل' };
    }

    var searchVal = String(passport).trim().toUpperCase();
    var allData = getAllData();

    // فلترة حسب المطار والصالة
    var filtered = filterByAirportAndHall_(allData, airport, hall);

    // البحث في رقم الجواز
    var results = filtered.filter(function(row) {
      return String(row[COL.PASSPORT]).toUpperCase().indexOf(searchVal) > -1;
    });

    if (results.length === 0) {
      return { success: false, error: 'لم يُعثر على حاج بهذا الرقم' };
    }

    // تحويل النتائج
    var pilgrims = results.map(function(row) {
      return buildPilgrimObject_(row);
    });

    return { success: true, data: pilgrims };

  } catch (e) {
    Logger.log('searchByPassport Error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ─────────────────────────────────────────────────────
//  تأكيد الاستقبال
// ─────────────────────────────────────────────────────

/**
 * confirmReception — تأكيد استقبال حاج أو مجموعة حجاج
 * @param {number[]} sheetRows — أرقام الصفوف في الشيت (1-based)
 * @param {string} staffName — اسم الموظف
 * @returns {Object} — نتيجة العملية
 */
function confirmReception(sheetRows, staffName) {
  try {
    if (!sheetRows || sheetRows.length === 0) {
      return { success: false, error: 'لم يتم تحديد حجاج' };
    }
    if (!staffName || staffName.trim() === '') {
      return { success: false, error: 'اسم الموظف مطلوب' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var staff = String(staffName).trim();

    var confirmed = 0;
    var alreadyConfirmed = [];

    // قراءة الحالات الحالية دفعة واحدة
    var statusRange = sheet.getRange(2, COL.RECEPTION_STATUS + 1, sheet.getLastRow() - 1, 3).getValues();

    // تحضير الكتابة
    sheetRows.forEach(function(rowNum) {
      var dataIdx = rowNum - 2; // تحويل من رقم صف الشيت إلى index في المصفوفة
      if (dataIdx >= 0 && dataIdx < statusRange.length) {
        var currentStatus = String(statusRange[dataIdx][0]).trim();
        if (currentStatus === 'تم') {
          // حاج مُستقبَل مسبقاً
          alreadyConfirmed.push({
            row: rowNum,
            staff: String(statusRange[dataIdx][2]).trim(),
            time: String(statusRange[dataIdx][1]).trim()
          });
        } else {
          // كتابة التأكيد
          sheet.getRange(rowNum, COL.RECEPTION_STATUS + 1).setValue('تم');
          sheet.getRange(rowNum, COL.RECEPTION_TIME + 1).setValue(now);
          sheet.getRange(rowNum, COL.RECEPTION_STAFF + 1).setValue(staff);
          confirmed++;
        }
      }
    });

    // مسح الكاش بعد التأكيد
    clearCache();

    return {
      success: true,
      confirmed: confirmed,
      alreadyConfirmed: alreadyConfirmed,
      message: 'تم تأكيد استقبال ' + confirmed + ' حاج'
    };

  } catch (e) {
    Logger.log('confirmReception Error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ─────────────────────────────────────────────────────
//  متابعة الرحلة — Flight Tracking
// ─────────────────────────────────────────────────────

/**
 * getFlightsList — قائمة الرحلات لتاريخ معين
 * @param {string} airport — المطار
 * @param {string} hall — الصالة
 * @param {string} date — التاريخ (اختياري — الافتراضي: اليوم)
 * @returns {Object} — قائمة الرحلات مع إحصائيات الاستقبال
 */
function getFlightsList(airport, hall, date) {
  try {
    var allData = getAllData();
    var filtered = filterByAirportAndHall_(allData, airport, hall);

    // فلتر التاريخ
    if (date && date !== '') {
      filtered = filtered.filter(function(row) {
        return String(row[COL.ARRIVE_DATE]).substring(0, 10) === date;
      });
    }

    // تجميع حسب الرحلة
    var flights = {};
    filtered.forEach(function(row) {
      var key = String(row[COL.FLIGHT_NUMBER]).trim() + '|' + String(row[COL.ARRIVE_DATE]).substring(0, 10);
      if (!flights[key]) {
        flights[key] = {
          flightNumber: String(row[COL.FLIGHT_NUMBER]).trim(),
          airline: String(row[COL.AIRLINE_EN]).trim(),
          airlineAr: String(row[COL.AIRLINE_AR]).trim(),
          date: String(row[COL.ARRIVE_DATE]).substring(0, 10),
          time: formatTime_(row[COL.ARRIVAL_TIME]),
          timeRaw: getRawTime_(row[COL.ARRIVAL_TIME]),
          total: 0,
          received: 0,
          remaining: 0
        };
      }
      flights[key].total++;
      if (String(row[COL.RECEPTION_STATUS]).trim() === 'تم') {
        flights[key].received++;
      }
    });

    // حساب المتبقي + ترتيب
    var result = Object.keys(flights).map(function(key) {
      var f = flights[key];
      f.remaining = f.total - f.received;
      f.percentage = Math.round((f.received / f.total) * 100);
      return f;
    });

    result.sort(function(a, b) {
      if (a.date !== b.date) return a.date < b.date ? -1 : 1;
      return a.timeRaw < b.timeRaw ? -1 : (a.timeRaw > b.timeRaw ? 1 : 0);
    });

    return { success: true, data: result };

  } catch (e) {
    Logger.log('getFlightsList Error: ' + e.message);
    return { success: false, error: e.message };
  }
}


/**
 * getFlightPilgrims — قائمة حجاج رحلة محددة
 * @param {string} airport — المطار
 * @param {string} hall — الصالة
 * @param {string} date — التاريخ
 * @param {string} flightNumber — رقم الرحلة
 * @returns {Object} — قائمة الحجاج مع حالة الاستقبال
 */
function getFlightPilgrims(airport, hall, date, flightNumber) {
  try {
    var allData = getAllData();
    var filtered = filterByAirportAndHall_(allData, airport, hall);

    // فلتر التاريخ + الرحلة
    var pilgrims = filtered.filter(function(row) {
      var dateMatch = String(row[COL.ARRIVE_DATE]).substring(0, 10) === date;
      var flightMatch = String(row[COL.FLIGHT_NUMBER]).trim() === flightNumber;
      return dateMatch && flightMatch;
    });

    var result = pilgrims.map(function(row) {
      return buildPilgrimObject_(row);
    });

    // ترتيب: غير المستقبَلين أولاً
    result.sort(function(a, b) {
      if (a.isReceived !== b.isReceived) return a.isReceived ? 1 : -1;
      return a.name.localeCompare(b.name);
    });

    // إحصائيات
    var received = result.filter(function(p) { return p.isReceived; }).length;
    var firstReceptionTime = '';
    pilgrims.forEach(function(row) {
      var t = String(row[COL.RECEPTION_TIME]).trim();
      if (t && (firstReceptionTime === '' || t < firstReceptionTime)) {
        firstReceptionTime = t;
      }
    });

    return {
      success: true,
      data: result,
      stats: {
        total: result.length,
        received: received,
        remaining: result.length - received,
        percentage: result.length > 0 ? Math.round((received / result.length) * 100) : 0,
        firstReceptionTime: firstReceptionTime
      }
    };

  } catch (e) {
    Logger.log('getFlightPilgrims Error: ' + e.message);
    return { success: false, error: e.message };
  }
}


// ─────────────────────────────────────────────────────
//  إحصائيات اليوم
// ─────────────────────────────────────────────────────

/**
 * getTodayStats — ملخص استقبال اليوم
 * @param {string} airport — المطار
 * @param {string} hall — الصالة
 * @returns {Object} — إحصائيات شاملة
 */
function getTodayStats(airport, hall) {
  try {
    var allData = getAllData();
    var filtered = filterByAirportAndHall_(allData, airport, hall);
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var todayData = filtered.filter(function(row) {
      return String(row[COL.ARRIVE_DATE]).substring(0, 10) === today;
    });

    var totalToday = todayData.length;
    var receivedToday = todayData.filter(function(row) {
      return String(row[COL.RECEPTION_STATUS]).trim() === 'تم';
    }).length;

    // عدد الرحلات
    var flights = {};
    todayData.forEach(function(row) {
      flights[String(row[COL.FLIGHT_NUMBER]).trim()] = true;
    });

    return {
      success: true,
      today: today,
      totalPilgrims: totalToday,
      received: receivedToday,
      remaining: totalToday - receivedToday,
      percentage: totalToday > 0 ? Math.round((receivedToday / totalToday) * 100) : 0,
      totalFlights: Object.keys(flights).length
    };

  } catch (e) {
    Logger.log('getTodayStats Error: ' + e.message);
    return { success: false, error: e.message };
  }
}


/**
 * getAvailableDates — التواريخ المتاحة (لقائمة الفلتر)
 * @param {string} airport — المطار
 * @param {string} hall — الصالة
 * @returns {Object} — قائمة التواريخ
 */
function getAvailableDates(airport, hall) {
  try {
    var allData = getAllData();
    var filtered = filterByAirportAndHall_(allData, airport, hall);

    var dates = {};
    filtered.forEach(function(row) {
      var d = String(row[COL.ARRIVE_DATE]).substring(0, 10);
      if (d && d !== '' && d !== 'undefined') {
        dates[d] = true;
      }
    });

    var sorted = Object.keys(dates).sort();
    return { success: true, data: sorted };

  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ─────────────────────────────────────────────────────
//  MRZ OCR — استخراج رقم الجواز من صورة
// ─────────────────────────────────────────────────────

/**
 * extractPassportFromImage — رفع صورة → OCR → استخراج رقم الجواز
 * @param {string} base64Image — الصورة بصيغة Base64
 * @returns {Object} — رقم الجواز المُستخرج أو خطأ
 */
/**
 * extractPassportFromImage — استخراج رقم الجواز عبر Cloud Vision API
 * طلب واحد مباشر — بدون رفع ملفات — سرعة 1-2 ثانية
 * @param {string} base64Image — الصورة بصيغة Base64
 * @returns {Object} — رقم الجواز أو خطأ
 */
function extractPassportFromImage(base64Image) {
  try {
    if (!base64Image) {
      return { success: false, error: 'لم يتم إرسال صورة' };
    }

    // ── طلب Cloud Vision API ──
    var apiUrl = 'https://vision.googleapis.com/v1/images:annotate';
    var token = ScriptApp.getOAuthToken();

    var requestBody = {
      requests: [{
        image: { content: base64Image },
        features: [{ type: 'TEXT_DETECTION', maxResults: 1 }]
      }]
    };

    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'POST',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());

    // ── التحقق من الاستجابة ──
    if (json.responses && json.responses[0] && json.responses[0].textAnnotations) {
      var fullText = json.responses[0].textAnnotations[0].description;

      // استخراج رقم الجواز
      var passport = extractPassportNumber_(fullText);

      if (passport) {
        return { success: true, passport: passport };
      } else {
        return { success: false, error: 'لم يتم التعرف على رقم الجواز — حاول تصوير أوضح' };
      }
    }

    // ── خطأ من API ──
    var errorMsg = 'لم يتم قراءة النص من الصورة';
    if (json.responses && json.responses[0] && json.responses[0].error) {
      errorMsg = json.responses[0].error.message;
    }
    return { success: false, error: errorMsg };

  } catch (e) {
    Logger.log('extractPassportFromImage Error: ' + e.message);
    return { success: false, error: 'خطأ في معالجة الصورة: ' + e.message };
  }
}


/**
 * extractPassportNumber_ — استخراج رقم الجواز من نص OCR
 * يبحث بـ 3 طرق مرتبة حسب الدقة:
 * 1. سطر MRZ الثاني (44 حرف، يحتوي رقم الجواز في المواقع 0-8)
 * 2. نمط رقم جواز قريب من كلمة "Passport" في النص
 * 3. أي نمط يشبه رقم جواز في كامل النص
 */
function extractPassportNumber_(text) {
  if (!text) return null;

  // تنظيف النص
  var cleanText = text.replace(/\r/g, '');
  var lines = cleanText.split('\n').map(function(l) { return l.trim(); });

  // ── المحاولة 1: البحث عن سطر MRZ الثاني ──
  // النمط: حرف/رقم + أحرف/أرقام/< بطول 28+ بدون مسافات
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/\s/g, '');
    // سطر MRZ الثاني: يبدأ بحرف أو رقم، طوله 28-44، يحتوي < 
    if (line.length >= 28 && line.indexOf('<') > -1 && /^[A-Z0-9]/.test(line)) {
      // التأكد أنه سطر MRZ وليس سطر الاسم (سطر الاسم يبدأ بـ P<)
      if (line.substring(0, 2) !== 'P<' && line.substring(0, 2) !== 'P«') {
        // رقم الجواز: من البداية حتى أول 
        var passportPart = line.split('<')[0];
        if (passportPart.length >= 5 && passportPart.length <= 15) {
          return passportPart;
        }
      }
    }
  }

  // ── المحاولة 2: البحث بجوار كلمة Passport No ──
  var allText = cleanText.replace(/\n/g, ' ');
  var passportLabelRegex = /Passport\s*(?:No|Number|#|Num)?\.?\s*[:/]?\s*([A-Z0-9]{5,15})/i;
  var labelMatch = allText.match(passportLabelRegex);
  if (labelMatch) {
    return labelMatch[1];
  }

  // ── المحاولة 3: أنماط أرقام الجوازات الشائعة ──
  // حرف + 6-8 أرقام (مثل T1287131) أو حرفين + 5-7 أرقام (مثل AB1234567)
  var patterns = [
    /\b([A-Z][0-9]{6,8})\b/g,        // T1287131, A12345678
    /\b([A-Z]{2}[0-9]{5,7})\b/g,     // AB1234567
    /\b([0-9]{7,9})\b/g              // 123456789 (أرقام فقط)
  ];

  for (var p = 0; p < patterns.length; p++) {
    var matches = allText.match(patterns[p]);
    if (matches && matches.length > 0) {
      // فلتر: استبعاد التواريخ والأرقام غير المناسبة
      for (var m = 0; m < matches.length; m++) {
        var candidate = matches[m];
        // استبعاد التواريخ (مثل 19770101, 20280205)
        if (/^(19|20)\d{6}$/.test(candidate)) continue;
        // استبعاد الأرقام القصيرة جداً
        if (candidate.replace(/[^0-9A-Z]/g, '').length < 6) continue;
        return candidate;
      }
    }
  }

  return null;
}


// ─────────────────────────────────────────────────────
//  دوال مساعدة
// ─────────────────────────────────────────────────────

/** فلترة حسب المطار والصالة */
function filterByAirportAndHall_(data, airport, hall) {
  var filtered = data;

  if (airport && airport !== '') {
    filtered = filtered.filter(function(row) {
      return String(row[COL.ARRIVE_CITY]).trim().toLowerCase() === String(airport).trim().toLowerCase();
    });
  }

  if (hall && hall !== '' && hall !== 'all') {
    filtered = filtered.filter(function(row) {
      return String(row[COL.HALL]).trim() === String(hall).trim();
    });
  }

  return filtered;
}


/** بناء كائن الحاج */
function buildPilgrimObject_(row) {
  var isReceived = String(row[COL.RECEPTION_STATUS]).trim() === 'تم';
  var ROW_NUM_INDEX = row.length - 1; // آخر عنصر = رقم الصف

  return {
    sheetRow:      row[ROW_NUM_INDEX],
    bookingId:     String(row[COL.BOOKING_ID]).trim(),
    groupNumber:   String(row[COL.GROUP_NUMBER]).trim(),
    name:          String(row[COL.NAME]).trim(),
    passport:      String(row[COL.PASSPORT]).trim(),
    gender:        String(row[COL.GENDER]).trim(),
    nationalityAr: String(row[COL.NATIONALITY_AR]).trim(),
    nationalityEn: String(row[COL.NATIONALITY_EN]).trim(),
    airline:       String(row[COL.AIRLINE_EN]).trim(),
    airlineAr:     String(row[COL.AIRLINE_AR]).trim(),
    arrivalTime:   formatTime_(row[COL.ARRIVAL_TIME]),
    arrivalDate:   String(row[COL.ARRIVE_DATE]).substring(0, 10),
    flightNumber:  String(row[COL.FLIGHT_NUMBER]).trim(),
    flightType:    String(row[COL.FLIGHT_TYPE]).trim(),
    transport:     getTransportLabel_(row[COL.FLIGHT_TYPE]),
    destination:   getDestinationLabel_(row[COL.FIRST_HOUSE]),
    hotel:         getDestinationHotel_(row),
    // حالة الاستقبال
    isReceived:    isReceived,
    receptionTime: isReceived ? String(row[COL.RECEPTION_TIME]).trim() : '',
    receptionStaff: isReceived ? String(row[COL.RECEPTION_STAFF]).trim() : ''
  };
}


/** تحديد الفندق حسب الوجهة الأولى */
function getDestinationHotel_(row) {
  var firstHouse = String(row[COL.FIRST_HOUSE]).trim().toLowerCase();
  if (firstHouse === 'madina') return row[COL.MADINAH_EN] || 'غير محدد';
  if (firstHouse === 'makkah shifting') return row[COL.MAKKAH_SHIFT_EN] || 'غير محدد';
  if (firstHouse === 'makkah') return row[COL.MAKKAH_EN] || 'غير محدد';
  return 'غير محدد';
}


/** تسمية المواصلات */
function getTransportLabel_(flightType) {
  var type = String(flightType).trim().toUpperCase();
  if (type === 'B2B') return 'إكرام الضيف';
  if (type === 'B2C') return 'تحالف النقل';
  return 'غير محدد';
}


/** تسمية الوجهة */
function getDestinationLabel_(firstHouse) {
  var fh = String(firstHouse).trim().toLowerCase();
  if (fh === 'madina') return 'المدينة المنورة';
  if (fh === 'makkah') return 'مكة المكرمة';
  if (fh === 'makkah shifting') return 'مكة المكرمة (تحويلي)';
  return firstHouse;
}


/** تنسيق الوقت (12 ساعة) */
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


/** وقت خام للترتيب */
function getRawTime_(timeStr) {
  if (!timeStr) return '99:99';
  var cleaned = String(timeStr).replace(/\.0+$/, '').trim();
  var parts = cleaned.split(':');
  if (parts.length < 2) return '99:99';
  return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
}


/** مسح الكاش */
function clearCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE_KEY);
  cache.remove(CONFIG.CACHE_KEY + '_CHUNKS');
  for (var i = 0; i < 20; i++) {
    cache.remove(CONFIG.CACHE_KEY + '_' + i);
  }
  return 'Cache cleared';
}

/**
 * testVisionAPI — اختبار تشخيصي لـ Cloud Vision API
 * شغّلها يدوياً من Apps Script → Run → اقرأ الـ Log
 */
