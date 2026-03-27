// ============================================
// جلب البيانات — الحاج + التنقل + خريطة الفندق
// ============================================

function findPilgrimByPassport_(passportNo) {
  var inputPassport = String(passportNo).toUpperCase().trim();
  var cacheKey = 'pilgrim_' + inputPassport;

  var cached = getCache_(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(JOURNEY_SHEET);
  if (!sheet) {
    Logger.log('Sheet not found: ' + JOURNEY_SHEET);
    return null;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var colPassport = 8;
  var colName = 7;

  for (var i = 2; i < data.length; i++) {
    var rowPassport = String(data[i][colPassport]).toUpperCase().trim();

    if (rowPassport === inputPassport) {
      var name = colName !== -1 ? String(data[i][colName]).trim() : 'حاج';
      var result = { name: name, row: i, rowData: data[i] };
      setCache_(cacheKey, result);
      return result;
    }
  }

  return null;
}

// ============================================
// جلب نوع التنقل من شيت الباقات
// الربط: عمود B في الباقات = عمود B في رحلة الحاج (PackageId)
// المخرج: "حافلة" / "قطار" أو "-"
// ============================================
function getTransportType_(packageId) {
  if (!packageId || packageId === '-') return '-';
  var inputId = String(packageId).trim();
  var cacheKey = 'transport_' + inputId;

  var cached = getCache_(cacheKey);
  if (cached !== null) return cached;

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('الباقات');
  if (!sheet) {
    Logger.log('Sheet not found: الباقات');
    return '-';
  }

  var data = sheet.getDataRange().getValues();
  var colB = 1;
  var colTransport = 65;

  for (var i = 2; i < data.length; i++) {
    if (String(data[i][colB]).trim() === inputId) {
      var val = String(data[i][colTransport]).trim() || '-';
      setCache_(cacheKey, val);
      return val;
    }
  }
  return '-';
}

// ============================================
// جلب رابط خريطة الفندق من شيت الفنادق
// ============================================
function getHotelMapLink_(hotelName) {
  if (!hotelName || hotelName === '-') return null;
  var cacheKey = 'hmap2_' + hotelName;

  var cached = getCache_(cacheKey);
  if (cached !== null) return cached || null;

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('الفنادق');
  if (!sheet) return null;

  var data = sheet.getDataRange().getValues();
  var input = normalizeHotelName_(hotelName);

  for (var i = 1; i < data.length; i++) {
    var nameAr = normalizeHotelName_(String(data[i][0] || ''));
    var nameEn = normalizeHotelName_(String(data[i][2] || ''));

    if (!nameAr && !nameEn) continue;

    if (nameAr === input || nameEn === input ||
        (nameAr && input && (nameAr.indexOf(input) !== -1 || input.indexOf(nameAr) !== -1)) ||
        (nameEn && input && (nameEn.indexOf(input) !== -1 || input.indexOf(nameEn) !== -1))) {
      var link = String(data[i][3] || '').trim();
      setCache_(cacheKey, link || '');
      return link || null;
    }
  }

  setCache_(cacheKey, '');
  return null;
}

function normalizeHotelName_(name) {
  if (!name) return '';
  return name.trim().toLowerCase()
    .replace(/^(فندق|شركة|مؤسسة)\s+/g, '')
    .replace(/(الفندقية|المحدودة|للفنادق|العالمية)/g, '')
    .replace(/[أإآ]/g, 'ا')
    .replace(/[ةً]/g, 'ه')
    .replace(/ى/g, 'ي')
    .replace(/\s+/g, ' ')
    .trim();
}

// ============================================
// جلب تفاصيل الرحلة من شيت الطيران (B2B)
// ============================================
function getB2BFlightDetails_(flightNumber, isArrival) {
  try {
    if (!flightNumber || flightNumber === '-') return null;
    var inputFN = String(flightNumber).trim().toUpperCase();
    var cacheKey = 'b2bflight_' + inputFN + '_' + (isArrival ? 'arr' : 'ret');

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('الطيران');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    // Arrival: leg1 indices 21-27, leg2 indices 28-34
    // Return: leg1 indices 35-41, leg2 indices 42-48
    var leg1Start = isArrival ? 21 : 35;
    var leg2Start = isArrival ? 28 : 42;

    for (var i = 2; i < data.length; i++) {
      var fn1 = String(data[i][leg1Start] || '').trim().toUpperCase();
      var fn2 = String(data[i][leg2Start] || '').trim().toUpperCase();

      if (fn1 === inputFN || fn2 === inputFN) {
        var result = {};
        var hasLeg1 = String(data[i][leg1Start] || '').trim() !== '';
        var hasLeg2 = String(data[i][leg2Start] || '').trim() !== '';

        if (hasLeg1) {
          result.leg1 = {
            flightNo: String(data[i][leg1Start] || '-'),
            dateDepart: formatDate_(data[i][leg1Start + 1]),
            timeDepart: formatTime_(data[i][leg1Start + 2]),
            from: String(data[i][leg1Start + 3] || '-'),
            to: String(data[i][leg1Start + 4] || '-'),
            dateLand: formatDate_(data[i][leg1Start + 5]),
            timeLand: formatTime_(data[i][leg1Start + 6])
          };
        }
        if (hasLeg2) {
          result.leg2 = {
            flightNo: String(data[i][leg2Start] || '-'),
            dateDepart: formatDate_(data[i][leg2Start + 1]),
            timeDepart: formatTime_(data[i][leg2Start + 2]),
            from: String(data[i][leg2Start + 3] || '-'),
            to: String(data[i][leg2Start + 4] || '-'),
            dateLand: formatDate_(data[i][leg2Start + 5]),
            timeLand: formatTime_(data[i][leg2Start + 6])
          };
        }

        setCache_(cacheKey, result);
        return result;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getB2BFlightDetails_ error: ' + e);
    return null;
  }
}

// ============================================
// جلب تفاصيل الرحلة من شيت B2C
// ============================================
function getB2CFlightDetails_(passport) {
  try {
    if (!passport) return null;
    var inputPassport = String(passport).toUpperCase().trim();
    var cacheKey = 'b2cflight_' + inputPassport;

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('B2C');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowPassport = String(data[i][5] || '').toUpperCase().trim();
      if (rowPassport !== inputPassport) continue;

      var result = { arrival: {}, return: {} };

      // Arrival leg1 (transit): indices 32-38, leg2 (or direct): indices 39-45
      var arrLeg1FN = String(data[i][32] || '').trim();
      var arrLeg2FN = String(data[i][39] || '').trim();

      if (arrLeg1FN) {
        result.arrival.leg1 = {
          flightNo: arrLeg1FN,
          dateDepart: formatDate_(data[i][33]),
          timeDepart: formatTime_(data[i][34]),
          from: String(data[i][35] || '-'),
          to: String(data[i][36] || '-'),
          dateLand: formatDate_(data[i][37]),
          timeLand: formatTime_(data[i][38])
        };
      }
      if (arrLeg2FN) {
        result.arrival.leg2 = {
          flightNo: arrLeg2FN,
          dateDepart: formatDate_(data[i][40]),
          timeDepart: formatTime_(data[i][41]),
          from: String(data[i][42] || '-'),
          to: String(data[i][43] || '-'),
          dateLand: formatDate_(data[i][44]),
          timeLand: formatTime_(data[i][45])
        };
      }

      // Return leg1 (transit): indices 46-52, leg2 (or direct): indices 53-59
      var retLeg1FN = String(data[i][46] || '').trim();
      var retLeg2FN = String(data[i][53] || '').trim();

      if (retLeg1FN) {
        result.return.leg1 = {
          flightNo: retLeg1FN,
          dateDepart: formatDate_(data[i][47]),
          timeDepart: formatTime_(data[i][48]),
          from: String(data[i][49] || '-'),
          to: String(data[i][50] || '-'),
          dateLand: formatDate_(data[i][51]),
          timeLand: formatTime_(data[i][52])
        };
      }
      if (retLeg2FN) {
        result.return.leg2 = {
          flightNo: retLeg2FN,
          dateDepart: formatDate_(data[i][54]),
          timeDepart: formatTime_(data[i][55]),
          from: String(data[i][56] || '-'),
          to: String(data[i][57] || '-'),
          dateLand: formatDate_(data[i][58]),
          timeLand: formatTime_(data[i][59])
        };
      }

      setCache_(cacheKey, result);
      return result;
    }
    return null;
  } catch (e) {
    Logger.log('getB2CFlightDetails_ error: ' + e);
    return null;
  }
}

// ============================================
// جلب اسم الباقة من شيت الباقات
// ============================================
function getPackageName_(packageId) {
  try {
    if (!packageId || packageId === '-') return null;
    var inputId = String(packageId).trim();
    var cacheKey = 'pkgname_' + inputId;

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('الباقات');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    for (var i = 2; i < data.length; i++) {
      if (String(data[i][1]).trim() === inputId) {
        var result = {
          nameAr: String(data[i][2] || '-'),
          nameEn: String(data[i][60] || '-')
        };
        setCache_(cacheKey, result);
        return result;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getPackageName_ error: ' + e);
    return null;
  }
}

// ============================================
// جلب بيانات المرشد السياحي
// ============================================
function getTourGuide_(packageId) {
  try {
    if (!packageId || packageId === '-') return null;
    var inputId = String(packageId).trim();
    var cacheKey = 'guide_' + inputId;

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Tour Guide');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][8]).trim() === inputId) {
        var result = {
          nameAr: String(data[i][0] || '-'),
          nameEn: (String(data[i][1] || '') + ' ' + String(data[i][2] || '')).trim() || '-',
          phone: String(data[i][5] || '-')
        };
        setCache_(cacheKey, result);
        return result;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getTourGuide_ error: ' + e);
    return null;
  }
}
