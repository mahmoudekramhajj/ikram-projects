// ============================================
// جلب البيانات — من الشيتات المصدرية مباشرة
// بدل "رحلة الحاج" المدمج: Presonal Details + الباقات + الطيران + B2C
// ============================================

// ============================================
// البحث عن حاج بالجواز — يبني Virtual Row بنفس هيكل رحلة الحاج القديم
// ============================================
function findPilgrimByPassport_(passportNo, skipCache) {
  var inputPassport = String(passportNo).toUpperCase().trim();
  var cacheKey = 'pilgrim_' + inputPassport;

  if (!skipCache) {
    var cached = getCache_(cacheKey);
    if (cached) return cached;
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // === 1. البحث في Presonal Details ===
  var pdSheet = ss.getSheetByName(PERSONAL_SHEET);
  if (!pdSheet) { Logger.log('Sheet not found: ' + PERSONAL_SHEET); return null; }

  var pdData = pdSheet.getDataRange().getValues();
  var pd = null;
  for (var i = 1; i < pdData.length; i++) {
    if (String(pdData[i][PD.PASSPORT]).toUpperCase().trim() === inputPassport) {
      pd = pdData[i];
      break;
    }
  }
  if (!pd) return null;

  var packageId = String(pd[PD.PACKAGE_NO] || '').trim();
  var flightType = String(pd[PD.FLIGHT_TYPE] || '').trim().toUpperCase();

  // === 2. البحث في الباقات ===
  var pkg = null;
  if (packageId) {
    var pkgSheet = ss.getSheetByName(PACKAGES_SHEET);
    if (pkgSheet) {
      var pkgData = pkgSheet.getDataRange().getValues();
      for (var p = 2; p < pkgData.length; p++) {
        if (String(pkgData[p][PKG.NUSK_NO]).trim() === packageId) {
          pkg = pkgData[p];
          break;
        }
      }
    }
  }

  // === 3. البحث في الطيران (B2B) أو B2C ===
  var flt = null;
  var b2c = null;

  if (flightType === 'B2C') {
    var b2cSheet = ss.getSheetByName(B2C_SHEET);
    if (b2cSheet) {
      var b2cData = b2cSheet.getDataRange().getValues();
      for (var c = 1; c < b2cData.length; c++) {
        if (String(b2cData[c][B2CI.PASSPORT]).toUpperCase().trim() === inputPassport) {
          b2c = b2cData[c];
          break;
        }
      }
    }
  } else {
    // B2B — البحث باسم العقد (CONTRACT_NAME) في عمود CM من الطيران
    var contractName = String(pd[PD.CONTRACT_NAME] || '').trim();
    if (contractName) {
      var fltSheet = ss.getSheetByName(FLIGHTS_SHEET);
      if (fltSheet) {
        var fltData = fltSheet.getDataRange().getValues();
        for (var f = 2; f < fltData.length; f++) {
          if (String(fltData[f][FLT.CONTRACT_NAME] || '').trim() === contractName) {
            flt = fltData[f];
            break;
          }
        }
      }
    }
  }

  // === 4. بناء Virtual Row بنفس هيكل رحلة الحاج (48 عنصر) ===
  var row = buildVirtualRow_(pd, pkg, flt, b2c, flightType);

  var name = (String(pd[PD.FIRST_NAME_AR] || '') + ' ' + String(pd[PD.LAST_NAME_AR] || '')).trim();
  if (!name || name === '') {
    name = (String(pd[PD.FIRST_NAME_EN] || '') + ' ' + String(pd[PD.LAST_NAME_EN] || '')).trim();
  }

  var result = { name: name, row: -1, rowData: row };
  setCache_(cacheKey, result);
  return result;
}

// ============================================
// بناء المصفوفة الافتراضية (Virtual Row) — نفس ترتيب أعمدة رحلة الحاج
// ============================================
function buildVirtualRow_(pd, pkg, flt, b2c, flightType) {
  var row = [];
  for (var i = 0; i < 48; i++) row.push('');

  // --- من Presonal Details ---
  row[1]  = String(pd[PD.PACKAGE_NO] || '');           // PackageId
  row[4]  = String(pd[PD.CAMP] || '');                  // CampName
  row[6]  = String(pd[PD.GROUP] || '');                  // GroupNo
  row[7]  = (String(pd[PD.FIRST_NAME_AR] || '') + ' ' + String(pd[PD.LAST_NAME_AR] || '')).trim(); // Name
  row[3]  = String(pd[PD.CONTRACT_NAME] || '');          // ContractName (for B2B flight lookup)
  row[8]  = String(pd[PD.PASSPORT] || '');               // Passport

  var natEN = String(pd[PD.NATIONALITY] || '').trim();
  var countryEN = String(pd[PD.COUNTRY_RESIDENCE] || '').trim();
  row[12] = natEN;                                       // Nationality EN
  row[13] = NATIONALITY_AR_MAP[natEN] || natEN;          // Nationality AR
  row[14] = countryEN;                                   // Country EN
  row[15] = COUNTRY_AR_MAP[countryEN] || countryEN;      // Country AR

  row[25] = flightType || String(pd[PD.FLIGHT_TYPE] || ''); // Flight Type (arrival)
  row[35] = row[25];                                     // Flight Type (return — same)

  // --- من الطيران (B2B) ---
  if (flt) {
    var airlineEN = String(flt[FLT.AIRLINE] || '').trim();
    row[16] = AIRLINE_AR_MAP[airlineEN] || airlineEN;    // Airline AR (arrival)
    row[17] = airlineEN;                                 // Airline EN (arrival)
    row[26] = AIRLINE_AR_MAP[airlineEN] || airlineEN;    // Airline AR (return)
    row[27] = airlineEN;                                 // Airline EN (return)

    // الوصول — الوجهة النهائية
    var hasGO2 = String(flt[FLT.GO2_FLIGHT] || '').trim() !== '';
    if (hasGO2) {
      row[19] = flt[FLT.GO2_TO];                        // Arrive City (final)
      row[20] = flt[FLT.GO2_LAND_DATE];                 // Arrive Date (final)
      row[18] = flt[FLT.GO2_LAND_TIME];                 // Arrive Time (final)
    } else {
      row[19] = flt[FLT.GO1_TO];
      row[20] = flt[FLT.GO1_LAND_DATE];
      row[18] = flt[FLT.GO1_LAND_TIME];
    }
    row[21] = flt[FLT.GO1_FROM];                        // Depart City
    row[23] = flt[FLT.GO1_TAKEOFF_TIME];                // Depart Time
    row[24] = flt[FLT.GO1_FLIGHT];                      // Flight Number

    // العودة — الوجهة النهائية
    var hasRET2 = String(flt[FLT.RET2_FLIGHT] || '').trim() !== '';
    if (hasRET2) {
      row[29] = flt[FLT.RET2_TO];                       // Return Arrive City (final)
      row[28] = flt[FLT.RET2_LAND_TIME];                // Return Arrive Time (final)
    } else {
      row[29] = flt[FLT.RET1_TO];
      row[28] = flt[FLT.RET1_LAND_TIME];
    }
    row[31] = flt[FLT.RET1_FROM];                       // Return Depart City
    row[32] = flt[FLT.RET1_TAKEOFF_DATE];               // Return Depart Date
    row[33] = flt[FLT.RET1_TAKEOFF_TIME];               // Return Depart Time
    row[34] = flt[FLT.RET1_FLIGHT];                     // Return Flight No
  }

  // --- من B2C ---
  if (b2c) {
    // الوصول (حتى 3 قطع — DEP0 إن وُجد، وإلا ARR1 كأول قطعة)
    var hasDEP0 = String(b2c[B2CI.DEP0_FLIGHT] || '').trim() !== '';
    var hasARR2 = String(b2c[B2CI.ARR2_FLIGHT] || '').trim() !== '';

    if (hasDEP0) {
      // 3 قطع: مدينة المغادرة الفعلية = DEP0_FROM
      row[24] = b2c[B2CI.DEP0_FLIGHT];                  // Flight Number (first leg)
      row[21] = b2c[B2CI.DEP0_FROM];                    // Depart City (origin)
      row[23] = b2c[B2CI.DEP0_TIME];                    // Depart Time
    } else {
      row[24] = b2c[B2CI.ARR1_FLIGHT];                  // Flight Number (first leg)
      row[21] = b2c[B2CI.ARR1_FROM];                    // Depart City
      row[23] = b2c[B2CI.ARR1_TIME];                    // Depart Time
    }

    if (hasARR2) {
      row[19] = b2c[B2CI.ARR2_TO];                      // Arrive City (final = JED/MED)
      row[20] = b2c[B2CI.ARR2_LAND_DATE];               // Arrive Date (final)
      row[18] = b2c[B2CI.ARR2_LAND_TIME];               // Arrive Time (final)
    } else {
      row[19] = b2c[B2CI.ARR1_TO];
      row[20] = b2c[B2CI.ARR1_LAND_DATE];
      row[18] = b2c[B2CI.ARR1_LAND_TIME];
    }

    // العودة (حتى 3 قطع — RET3 إن وُجد، وإلا RET2 كآخر قطعة)
    var hasRET2b = String(b2c[B2CI.RET2_FLIGHT] || '').trim() !== '';
    var hasRET3 = String(b2c[B2CI.RET3_FLIGHT] || '').trim() !== '';

    row[34] = b2c[B2CI.RET1_FLIGHT];                    // Return Flight No (first leg)
    row[31] = b2c[B2CI.RET1_FROM];                      // Return Depart City (JED/MED)
    row[32] = b2c[B2CI.RET1_DATE];                      // Return Depart Date
    row[33] = b2c[B2CI.RET1_TIME];                      // Return Depart Time

    if (hasRET3) {
      // 3 قطع: مدينة الوصول النهائية = RET3_TO
      row[29] = b2c[B2CI.RET3_TO];                      // Return Arrive City (final destination)
      row[28] = b2c[B2CI.RET3_LAND_TIME];               // Return Arrive Time (final)
    } else if (hasRET2b) {
      row[29] = b2c[B2CI.RET2_TO];                      // Return Arrive City (final)
      row[28] = b2c[B2CI.RET2_LAND_TIME];               // Return Arrive Time (final)
    } else {
      row[29] = b2c[B2CI.RET1_TO];
      row[28] = b2c[B2CI.RET1_LAND_TIME];
    }

    // شركة الطيران — نحاول استنتاجها من رقم الرحلة
    var arrFN = String(b2c[B2CI.ARR1_FLIGHT] || '').trim();
    var airlineGuess = guessAirlineFromFlight_(arrFN);
    row[17] = airlineGuess.en || '';                     // Airline EN (arrival)
    row[16] = airlineGuess.ar || '';                     // Airline AR (arrival)

    var retFN = String(b2c[B2CI.RET1_FLIGHT] || '').trim();
    var retAirlineGuess = guessAirlineFromFlight_(retFN);
    row[27] = retAirlineGuess.en || '';                  // Airline EN (return)
    row[26] = retAirlineGuess.ar || '';                  // Airline AR (return)
  }

  // --- من الباقات (فنادق + تنقل) ---
  if (pkg) {
    var h1City = String(pkg[PKG.H1_CITY] || '').trim();
    var h2City = String(pkg[PKG.H2_CITY] || '').trim();

    // السكن الأول والثاني (بالترتيب الزمني — H1 أولاً ثم H2)
    row[36] = h1City;                                    // FirstHouse
    row[37] = pkg[PKG.H1_CHECKIN];                       // FirstHouseStart
    row[38] = pkg[PKG.H1_CHECKOUT];                      // FirstHouseEnd
    row[39] = h2City;                                    // LastHouse
    row[40] = pkg[PKG.H2_CHECKIN];                       // LastHouseStart
    row[41] = pkg[PKG.H2_CHECKOUT];                      // LastHouseEnd

    // تحديد فندق مكة وفندق المدينة حسب المدينة
    var h1IsMadinah = isMadinahCity_(h1City);

    if (h1IsMadinah) {
      // H1 = المدينة, H2 = مكة
      row[46] = pkg[PKG.H1_NAME_AR];                    // Madinah Hotel AR
      row[47] = pkg[PKG.H1_NAME_EN];                    // Madinah Hotel EN
      row[42] = pkg[PKG.H2_NAME_AR];                    // Makkah Hotel AR
      row[43] = pkg[PKG.H2_NAME_EN];                    // Makkah Hotel EN
    } else {
      // H1 = مكة, H2 = المدينة
      row[42] = pkg[PKG.H1_NAME_AR];                    // Makkah Hotel AR
      row[43] = pkg[PKG.H1_NAME_EN];                    // Makkah Hotel EN
      row[46] = pkg[PKG.H2_NAME_AR];                    // Madinah Hotel AR
      row[47] = pkg[PKG.H2_NAME_EN];                    // Madinah Hotel EN
    }

    // فندق Shift (الثالث)
    var h3NameAr = String(pkg[PKG.H3_NAME_AR] || '').trim();
    var h3NameEn = String(pkg[PKG.H3_NAME_EN] || '').trim();
    row[44] = h3NameAr;                                  // Shift Hotel AR
    row[45] = h3NameEn;                                  // Shift Hotel EN

    // خانات صريحة للفندقين H1 و H2 (ثابتة بالترتيب الزمني، بغض النظر عن المدينة)
    // تُحلّ مشكلة عرض الفنادق لباقات Shifting داخل مكة (H1 + H2 كلاهما في مكة)
    row[48] = String(pkg[PKG.H1_NAME_AR] || '').trim();  // H1 Hotel AR
    row[49] = String(pkg[PKG.H1_NAME_EN] || '').trim();  // H1 Hotel EN
    row[50] = String(pkg[PKG.H2_NAME_AR] || '').trim();  // H2 Hotel AR
    row[51] = String(pkg[PKG.H2_NAME_EN] || '').trim();  // H2 Hotel EN
  }

  return row;
}

// ============================================
// تحديد هل المدينة = المدينة المنورة
// ============================================
function isMadinahCity_(city) {
  if (!city) return false;
  var c = city.toLowerCase().trim();
  return c === 'med' || c === 'madi' || c === 'madina' || c === 'madinah' || c === 'medina' ||
    c.indexOf('madi') !== -1 || c.indexOf('مدين') !== -1 || c.indexOf('medina') !== -1;
}

// ============================================
// تحويل كود المدينة لاسم كامل
// ============================================
function getCityDisplayName_(cityCode, lang) {
  if (!cityCode) return '-';
  var c = String(cityCode).toLowerCase().trim();
  var isAr = (lang === 'ar');

  if (isMadinahCity_(c)) {
    return isAr ? 'المدينة المنورة' : 'Madinah';
  }
  if (c.indexOf('shift') !== -1 || c.indexOf('تحويل') !== -1) {
    return isAr ? 'مكة (تحويل)' : 'Makkah (Shifting)';
  }
  if (c === 'mak' || c === 'makkah' || c === 'mecca' || c.indexOf('مك') !== -1 || c.indexOf('makk') !== -1) {
    return isAr ? 'مكة المكرمة' : 'Makkah';
  }

  return cityCode;
}

// ============================================
// استنتاج شركة الطيران من رقم الرحلة (B2C)
// ============================================
function guessAirlineFromFlight_(flightNo) {
  if (!flightNo) return { en: '', ar: '' };
  var code = flightNo.replace(/[0-9\s]/g, '').toUpperCase();

  var codeMap = {
    'SV': 'Saudia', 'XY': 'Flyadeal', 'TK': 'Turkish Airlines',
    'EK': 'Emirates', 'EY': 'Etihad', 'QR': 'Qatar Airways',
    'GF': 'Gulf Air', 'MS': 'EgyptAir', 'RJ': 'Royal Jordanian',
    'ET': 'Ethiopian Airlines', 'F3': 'Flydubai', 'PC': 'Pegasus Airlines',
    'A3': 'Aegean Airlines', 'SM': 'Air Cairo', 'TF': 'AJet',
    'W6': 'Wizz Air', 'AJ': 'AnadoluJet'
  };

  var airlineEN = codeMap[code] || '';
  var airlineAR = airlineEN ? (AIRLINE_AR_MAP[airlineEN] || airlineEN) : '';
  return { en: airlineEN, ar: airlineAR };
}

// ============================================
// تأكيد وصول الحاج — الكتابة في شيت "Pilgrim Data"
// ============================================
function confirmArrivalInSheet_(passport, source) {
  try {
    var key = String(passport).toUpperCase().trim();
    var sheet = getPilgrimDataSheet_();
    var data = sheet.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');

    // البحث عن الحاج
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase().trim() === key) {
        var row = i + 1;
        sheet.getRange(row, PDATA_RECEPTION_STATUS + 1).setValue('تم');
        sheet.getRange(row, PDATA_RECEPTION_TIME + 1).setValue(now);
        sheet.getRange(row, PDATA_RECEPTION_STAFF + 1).setValue(source);
        return true;
      }
    }

    // لم يوجد صف — إنشاء صف جديد
    var pilgrim = findPilgrimByPassport_(passport);
    var name = pilgrim ? pilgrim.name : '';
    var newRow = [key, name, '', '', '', '', '', new Date().toISOString(), 'تم', now, source];
    sheet.appendRow(newRow);
    return true;

  } catch (e) {
    Logger.log('confirmArrivalInSheet_ error: ' + e.message);
    return false;
  }
}

// ============================================
// قراءة حالة الاستقبال من Pilgrim Data
// ============================================
function getReceptionStatus_(passport) {
  try {
    if (!passport) return null;
    var key = String(passport).toUpperCase().trim();
    var sheet = getPilgrimDataSheet_();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase().trim() === key) {
        var status = String(data[i][PDATA_RECEPTION_STATUS] || '').trim();
        var time = String(data[i][PDATA_RECEPTION_TIME] || '').trim();
        return { status: status, time: time };
      }
    }
    return null;
  } catch (e) {
    Logger.log('getReceptionStatus_ error: ' + e.message);
    return null;
  }
}

// ============================================
// تحديد مجموعة العمليات المناسبة حسب المطار وشركة الطيران
// ============================================
function getOpsGroupChatId_(arriveCity, airlineEn) {
  var city = String(arriveCity).trim().toLowerCase();

  // مطار المدينة
  if (city === 'madinah' || city === 'madina' || city.indexOf('مدين') !== -1) {
    return OPS_GROUPS.madinah;
  }

  // مطار جدة — تحديد الصالة حسب شركة الطيران
  if (city === 'jeddah' || city === 'jed' || city.indexOf('جد') !== -1) {
    var airline = String(airlineEn).trim();
    var terminal = AIRLINE_TERMINAL[airline] || 'T1'; // افتراضي: صالة 1
    return terminal === 'N' ? OPS_GROUPS.jeddah_north : OPS_GROUPS.jeddah_t1;
  }

  // افتراضي: جدة صالة 1
  return OPS_GROUPS.jeddah_t1;
}

// ============================================
// جلب نوع التنقل من شيت الباقات
// ============================================
function getTransportType_(packageId) {
  if (!packageId || packageId === '-') return '-';
  var inputId = String(packageId).trim();
  var cacheKey = 'transport_' + inputId;

  var cached = getCache_(cacheKey);
  if (cached !== null) return cached;

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(PACKAGES_SHEET);
  if (!sheet) {
    Logger.log('Sheet not found: ' + PACKAGES_SHEET);
    return '-';
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    if (String(data[i][PKG.NUSK_NO]).trim() === inputId) {
      var val = String(data[i][PKG.TRANSPORT]).trim() || '-';
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

    var nameArNoSpc = nameAr.replace(/\s/g, '');
    var nameEnNoSpc = nameEn.replace(/\s/g, '');
    var inputNoSpc = input.replace(/\s/g, '');

    if (nameAr === input || nameEn === input ||
        (nameAr && input && (nameAr.indexOf(input) !== -1 || input.indexOf(nameAr) !== -1)) ||
        (nameEn && input && (nameEn.indexOf(input) !== -1 || input.indexOf(nameEn) !== -1)) ||
        (nameArNoSpc && inputNoSpc && (nameArNoSpc.indexOf(inputNoSpc) !== -1 || inputNoSpc.indexOf(nameArNoSpc) !== -1)) ||
        (nameEnNoSpc && inputNoSpc && (nameEnNoSpc.indexOf(inputNoSpc) !== -1 || inputNoSpc.indexOf(nameEnNoSpc) !== -1))) {
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
function getB2BFlightDetails_(flightNumber, isArrival, contractName) {
  try {
    if (!flightNumber || flightNumber === '-') return null;
    var inputFN = String(flightNumber).trim().toUpperCase();
    var inputCN = contractName ? String(contractName).trim() : '';
    var cacheKey = 'b2bflight_' + inputFN + '_' + (isArrival ? 'arr' : 'ret') + (inputCN ? '_' + inputCN : '');

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(FLIGHTS_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    var leg1Start = isArrival ? 21 : 35;
    var leg2Start = isArrival ? 28 : 42;

    for (var i = 2; i < data.length; i++) {
      // إذا متوفر اسم العقد — نبحث به أولاً (أدق)
      if (inputCN) {
        if (String(data[i][FLT.CONTRACT_NAME] || '').trim() !== inputCN) continue;
      } else {
        var fn1 = String(data[i][leg1Start] || '').trim().toUpperCase();
        var fn2 = String(data[i][leg2Start] || '').trim().toUpperCase();
        if (fn1 !== inputFN && fn2 !== inputFN) continue;
      }
      {
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
    var sheet = ss.getSheetByName(B2C_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowPassport = String(data[i][B2CI.PASSPORT] || '').toUpperCase().trim();
      if (rowPassport !== inputPassport) continue;

      var result = { arrival: {}, return: {} };

      var arrLeg1FN = String(data[i][B2CI.ARR1_FLIGHT] || '').trim();
      var arrLeg2FN = String(data[i][B2CI.ARR2_FLIGHT] || '').trim();

      if (arrLeg1FN) {
        result.arrival.leg1 = {
          flightNo: arrLeg1FN,
          dateDepart: formatDate_(data[i][B2CI.ARR1_DATE]),
          timeDepart: formatTime_(data[i][B2CI.ARR1_TIME]),
          from: String(data[i][B2CI.ARR1_FROM] || '-'),
          to: String(data[i][B2CI.ARR1_TO] || '-'),
          dateLand: formatDate_(data[i][B2CI.ARR1_LAND_DATE]),
          timeLand: formatTime_(data[i][B2CI.ARR1_LAND_TIME])
        };
      }
      if (arrLeg2FN) {
        result.arrival.leg2 = {
          flightNo: arrLeg2FN,
          dateDepart: formatDate_(data[i][B2CI.ARR2_DATE]),
          timeDepart: formatTime_(data[i][B2CI.ARR2_TIME]),
          from: String(data[i][B2CI.ARR2_FROM] || '-'),
          to: String(data[i][B2CI.ARR2_TO] || '-'),
          dateLand: formatDate_(data[i][B2CI.ARR2_LAND_DATE]),
          timeLand: formatTime_(data[i][B2CI.ARR2_LAND_TIME])
        };
      }

      var retLeg1FN = String(data[i][B2CI.RET1_FLIGHT] || '').trim();
      var retLeg2FN = String(data[i][B2CI.RET2_FLIGHT] || '').trim();

      if (retLeg1FN) {
        result.return.leg1 = {
          flightNo: retLeg1FN,
          dateDepart: formatDate_(data[i][B2CI.RET1_DATE]),
          timeDepart: formatTime_(data[i][B2CI.RET1_TIME]),
          from: String(data[i][B2CI.RET1_FROM] || '-'),
          to: String(data[i][B2CI.RET1_TO] || '-'),
          dateLand: formatDate_(data[i][B2CI.RET1_LAND_DATE]),
          timeLand: formatTime_(data[i][B2CI.RET1_LAND_TIME])
        };
      }
      if (retLeg2FN) {
        result.return.leg2 = {
          flightNo: retLeg2FN,
          dateDepart: formatDate_(data[i][B2CI.RET2_DATE]),
          timeDepart: formatTime_(data[i][B2CI.RET2_TIME]),
          from: String(data[i][B2CI.RET2_FROM] || '-'),
          to: String(data[i][B2CI.RET2_TO] || '-'),
          dateLand: formatDate_(data[i][B2CI.RET2_LAND_DATE]),
          timeLand: formatTime_(data[i][B2CI.RET2_LAND_TIME])
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
// جلب تواريخ الفندق الثالث من شيت الباقات
// ============================================
function getThirdHotelDates_(packageId) {
  try {
    if (!packageId || packageId === '-') return null;
    var inputId = String(packageId).trim();
    var cacheKey = 'hotel3_' + inputId;

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(PACKAGES_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    for (var i = 2; i < data.length; i++) {
      if (String(data[i][PKG.NUSK_NO]).trim() === inputId) {
        var result = {
          city: String(data[i][PKG.H3_CITY] || '').trim(),
          checkIn: data[i][PKG.H3_CHECKIN],
          checkOut: data[i][PKG.H3_CHECKOUT],
          packageEnd: data[i][PKG.DATE_END]
        };
        setCache_(cacheKey, result);
        return result;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getThirdHotelDates_ error: ' + e);
    return null;
  }
}

// ============================================
// 🆕 جلب جميع فنادق الباقة (H1+H2+H3) من شيت الباقات — مصدر الحقيقة الموحَّد
// يُرجع مصفوفة مرتَّبة زمنياً: [{city, nameAr, nameEn, checkIn, checkOut}, ...]
// ============================================
function getPackageHotels_(packageId) {
  try {
    if (!packageId || packageId === '-') return [];
    var inputId = String(packageId).trim();
    var cacheKey = 'pkghotels_' + inputId;

    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(PACKAGES_SHEET);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    var hotels = [];

    function isValidName_(s) {
      if (!s) return false;
      var v = String(s).trim();
      if (!v || v === '-' || v.toLowerCase() === 'null') return false;
      return true;
    }

    for (var i = 2; i < data.length; i++) {
      if (String(data[i][PKG.NUSK_NO]).trim() !== inputId) continue;
      var row = data[i];

      // H1
      var h1Name = String(row[PKG.H1_NAME_AR] || '').trim();
      if (isValidName_(h1Name)) {
        hotels.push({
          city: String(row[PKG.H1_CITY] || '').trim(),
          nameAr: h1Name,
          nameEn: String(row[PKG.H1_NAME_EN] || '').trim(),
          checkIn: row[PKG.H1_CHECKIN],
          checkOut: row[PKG.H1_CHECKOUT]
        });
      }

      // H2
      var h2Name = String(row[PKG.H2_NAME_AR] || '').trim();
      if (isValidName_(h2Name)) {
        hotels.push({
          city: String(row[PKG.H2_CITY] || '').trim(),
          nameAr: h2Name,
          nameEn: String(row[PKG.H2_NAME_EN] || '').trim(),
          checkIn: row[PKG.H2_CHECKIN],
          checkOut: row[PKG.H2_CHECKOUT]
        });
      }

      // H3
      var h3Name = String(row[PKG.H3_NAME_AR] || '').trim();
      if (isValidName_(h3Name)) {
        hotels.push({
          city: String(row[PKG.H3_CITY] || '').trim(),
          nameAr: h3Name,
          nameEn: String(row[PKG.H3_NAME_EN] || '').trim(),
          checkIn: row[PKG.H3_CHECKIN],
          checkOut: row[PKG.H3_CHECKOUT]
        });
      }

      break;
    }

    setCache_(cacheKey, hotels);
    return hotels;
  } catch (e) {
    Logger.log('getPackageHotels_ error: ' + e);
    return [];
  }
}

// ============================================
// 🆕 جلب بيانات وصول الحاج (آخر leg + اسم المطار + IATA)
// يعتمد على getPilgrimFlightData_ ثم يُثري بـ getAirportDisplay_
// ============================================
function getPilgrimArrival_(passport, lang) {
  try {
    var data = getPilgrimFlightData_(passport);
    if (!data || !data.arrival || !data.arrival.legs || !data.arrival.legs.length) return null;
    var legs = data.arrival.legs;
    var last = legs[legs.length - 1];
    return {
      date: last.landDate,
      time: last.landTime,
      iata: last.to,
      airportAr: getAirportDisplay_(last.to, 'ar'),
      airportEn: getAirportDisplay_(last.to, 'en'),
      flightNo: last.flightNo,
      flightType: data.flightType || ''
    };
  } catch (e) {
    Logger.log('getPilgrimArrival_ error: ' + e);
    return null;
  }
}

// ============================================
// 🆕 جلب بيانات مغادرة الحاج (أول leg في العودة + اسم المطار)
// ============================================
function getPilgrimDeparture_(passport, lang) {
  try {
    var data = getPilgrimFlightData_(passport);
    if (!data || !data.return || !data.return.legs || !data.return.legs.length) return null;
    var first = data.return.legs[0];
    return {
      date: first.takeoffDate,
      time: first.takeoffTime,
      iata: first.from,
      airportAr: getAirportDisplay_(first.from, 'ar'),
      airportEn: getAirportDisplay_(first.from, 'en'),
      flightNo: first.flightNo
    };
  } catch (e) {
    Logger.log('getPilgrimDeparture_ error: ' + e);
    return null;
  }
}

// ============================================
// 🆕 منطق الوصول: تطبيق تاريخ وصول الحاج على فنادق الباقة
// قواعد:
//  ١) إذا أول فندق في "Med" + تاريخ وصول الحاج >= تاريخ خروج فندق المدينة → يُحذف فندق المدينة
//  ٢) وإلا: يُعدَّل تاريخ دخول أول فندق ليطابق تاريخ وصول الحاج (إن كان الوصول لاحقاً)
// يُرجع نسخة جديدة من المصفوفة (لا يُعدّل الأصل) — الفنادق المعدَّلة تُؤشَّر بـ `adjusted: true`
// ============================================
function applyArrivalOverride_(hotels, arrivalDate) {
  if (!hotels || !hotels.length) return hotels || [];
  if (!arrivalDate) return hotels;

  function toTs_(d) {
    if (!d) return NaN;
    if (d instanceof Date) return d.getTime();
    return new Date(d).getTime();
  }
  function copyHotel_(h) {
    return {
      city: h.city, nameAr: h.nameAr, nameEn: h.nameEn,
      checkIn: h.checkIn, checkOut: h.checkOut,
      adjusted: !!h.adjusted, dropped: !!h.dropped
    };
  }

  var arrTs = toTs_(arrivalDate);
  if (isNaN(arrTs)) return hotels;

  var first = hotels[0];
  var firstCity = String(first.city || '').toLowerCase().trim();
  var isFirstMadinah = isMadinahCity_(firstCity);
  var firstOutTs = toTs_(first.checkOut);

  // قاعدة ١: حذف فندق المدينة إن وصل الحاج بعد انتهائه
  if (isFirstMadinah && !isNaN(firstOutTs) && arrTs >= firstOutTs) {
    var rest = hotels.slice(1);
    if (rest.length > 0) {
      var newFirst = copyHotel_(rest[0]);
      var newFirstInTs = toTs_(newFirst.checkIn);
      if (!isNaN(newFirstInTs) && arrTs > newFirstInTs) {
        newFirst.checkIn = arrivalDate;
        newFirst.adjusted = true;
      }
      rest[0] = newFirst;
    }
    return rest;
  }

  // قاعدة ٢: تعديل دخول أول فندق ليطابق وصول الحاج (إن كان الوصول لاحقاً)
  var firstInTs = toTs_(first.checkIn);
  if (!isNaN(firstInTs) && arrTs > firstInTs) {
    var modified = hotels.slice();
    modified[0] = copyHotel_(first);
    modified[0].checkIn = arrivalDate;
    modified[0].adjusted = true;
    return modified;
  }

  return hotels;
}

// ============================================
// حساب تاريخ خروج الفندق الثالث
// ============================================
function calcThirdHotelCheckout_(returnDate, returnTime, packageEnd, fallbackCheckout) {
  var depDate = normDate_(returnDate);
  var depTime = formatTime_(returnTime);
  if (!depDate || depDate === '-') return formatDate_(fallbackCheckout);

  var timeParts = depTime.split(':');
  var hour = parseInt(timeParts[0], 10);
  if (isNaN(hour)) return formatDate_(fallbackCheckout);

  if (hour < 18) return depDate;

  var pkgEnd = normDate_(packageEnd);
  if (!pkgEnd || pkgEnd === '-') return depDate;

  if (pkgEnd <= depDate) return depDate;

  return getDateOffset_(depDate, 1);
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
    var sheet = ss.getSheetByName(PACKAGES_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();

    for (var i = 2; i < data.length; i++) {
      if (String(data[i][PKG.NUSK_NO]).trim() === inputId) {
        var result = {
          nameAr: String(data[i][PKG.NAME_AR] || '-'),
          nameEn: String(data[i][PKG.NAME_EN] || '-')
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

// ============================================
// جلب حالة التأشيرة + بيانات التذكرة من Presonal Details
// ============================================
function getVisaAndTicket_(passport) {
  try {
    if (!passport) return null;
    var passportNo = String(passport).trim().toUpperCase();
    var cacheKey = 'visa_ticket_' + passportNo;
    var cached = getCache_(cacheKey);
    if (cached !== null) return cached;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(PERSONAL_SHEET);
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][PD.PASSPORT]).trim().toUpperCase() === passportNo) {
        var result = {
          visaStatus: String(data[i][PD.VISA_STATUS] || '-'),
          ticketNo: String(data[i][PD.TICKET_NO] || ''),
          ticketLink: String(data[i][PD.TICKET_URL] || '')
        };
        setCache_(cacheKey, result);
        return result;
      }
    }
    return null;
  } catch (e) {
    Logger.log('getVisaAndTicket_ error: ' + e.message);
    return null;
  }
}

// ============================================
// جلب ملف التأشيرة من Google Drive
// ============================================
function getVisaImage_(passport) {
  try {
    if (!passport || !VISA_FOLDER_ID) return null;
    var passportNo = String(passport).trim().toUpperCase();
    var cacheKey = 'visa_img_' + passportNo;
    var cached = getCache_(cacheKey);
    if (cached !== null) return cached || null;

    var folder = DriveApp.getFolderById(VISA_FOLDER_ID);
    var files = folder.getFilesByName(passportNo + '.pdf');
    if (files.hasNext()) {
      var file = files.next();
      var result = { fileId: file.getId(), fileName: passportNo + '.pdf' };
      setCache_(cacheKey, result);
      return result;
    }
    setCache_(cacheKey, '');
    return null;
  } catch (e) {
    Logger.log('getVisaImage_ error: ' + e);
    return null;
  }
}

// إرسال ملف التأشيرة عبر Telegram كمستند
function sendVisaDocument_(chatId, visaInfo, caption, replyMarkup) {
  try {
    var file = DriveApp.getFileById(visaInfo.fileId);
    var blob = file.getBlob();
    var options = {
      method: 'post',
      payload: {
        chat_id: String(chatId),
        document: blob,
        caption: caption || '',
        parse_mode: 'HTML'
      },
      muteHttpExceptions: true
    };
    if (replyMarkup) options.payload.reply_markup = JSON.stringify(replyMarkup);
    UrlFetchApp.fetch(TELEGRAM_API + '/sendDocument', options);
  } catch (e) {
    Logger.log('sendVisaDocument_ error: ' + e);
  }
}

// ============================================
// جلب بيانات الطيران للحاج — هيكل نظيف ومنظّم
// يدعم: مباشر / ترانزيت ذهاب / ترانزيت عودة / ترانزيت الاثنين
// يقرأ من: "الطيران" (B2B) أو "B2C" (B2C)
// ============================================
function getPilgrimFlightData_(passport) {
  if (!passport) return null;
  var inputPassport = String(passport).toUpperCase().trim();
  var cacheKey = 'flightdata_' + inputPassport;

  var cached = getCache_(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // === 1. قراءة Presonal Details ===
  var pdSheet = ss.getSheetByName(PERSONAL_SHEET);
  if (!pdSheet) return null;
  var pdData = pdSheet.getDataRange().getValues();
  var pd = null;
  for (var i = 1; i < pdData.length; i++) {
    if (String(pdData[i][PD.PASSPORT]).toUpperCase().trim() === inputPassport) {
      pd = pdData[i];
      break;
    }
  }
  if (!pd) return null;

  var flightType = String(pd[PD.FLIGHT_TYPE] || '').trim().toUpperCase();
  var contractName = String(pd[PD.CONTRACT_NAME] || '').trim();

  var result = {
    flightType: flightType,
    airline: { code: '', nameAr: '', nameEn: '' },
    arrival: { direct: true, legs: [] },
    return:  { direct: true, legs: [] }
  };

  if (flightType === 'B2B') {
    // === B2B: قراءة من شيت "الطيران" ===
    var fltSheet = ss.getSheetByName(FLIGHTS_SHEET);
    if (!fltSheet || !contractName) return result;
    var fltData = fltSheet.getDataRange().getValues();
    var flt = null;
    for (var f = 2; f < fltData.length; f++) {
      if (String(fltData[f][FLT.CONTRACT_NAME] || '').trim() === contractName) {
        flt = fltData[f];
        break;
      }
    }
    if (!flt) return result;

    var airlineEN = String(flt[FLT.AIRLINE] || '').trim();
    result.airline.code = airlineEN;
    result.airline.nameEn = airlineEN;
    result.airline.nameAr = AIRLINE_AR_MAP[airlineEN] || airlineEN;

    // === رحلة الوصول ===
    var go1Filled = String(flt[FLT.GO1_FLIGHT] || '').trim() !== '';
    var go2Filled = String(flt[FLT.GO2_FLIGHT] || '').trim() !== '';

    if (go1Filled) {
      result.arrival.legs.push(buildLeg_(flt, FLT.GO1_FLIGHT));
    }
    if (go2Filled) {
      result.arrival.legs.push(buildLeg_(flt, FLT.GO2_FLIGHT));
    }
    result.arrival.direct = (result.arrival.legs.length <= 1);

    // === رحلة العودة ===
    var ret1Filled = String(flt[FLT.RET1_FLIGHT] || '').trim() !== '';
    var ret2Filled = String(flt[FLT.RET2_FLIGHT] || '').trim() !== '';

    if (ret1Filled) {
      result.return.legs.push(buildLeg_(flt, FLT.RET1_FLIGHT));
    }
    if (ret2Filled) {
      result.return.legs.push(buildLeg_(flt, FLT.RET2_FLIGHT));
    }
    result.return.direct = (result.return.legs.length <= 1);

  } else if (flightType === 'B2C') {
    // === B2C: قراءة من شيت "B2C" ===
    var b2cSheet = ss.getSheetByName(B2C_SHEET);
    if (!b2cSheet) return result;
    var b2cData = b2cSheet.getDataRange().getValues();
    var b2c = null;
    for (var c = 1; c < b2cData.length; c++) {
      if (String(b2cData[c][B2CI.PASSPORT]).toUpperCase().trim() === inputPassport) {
        b2c = b2cData[c];
        break;
      }
    }
    if (!b2c) return result;

    // === رحلة الوصول (حتى 3 قطع: DEP0 + ARR1 + ARR2) ===
    var dep0Filled = String(b2c[B2CI.DEP0_FLIGHT] || '').trim() !== '';
    var arr1Filled = String(b2c[B2CI.ARR1_FLIGHT] || '').trim() !== '';
    var arr2Filled = String(b2c[B2CI.ARR2_FLIGHT] || '').trim() !== '';

    // الترتيب الزمني: DEP0 (الأولى) → ARR1 (ترانزيت) → ARR2 (الوصول)
    if (dep0Filled) {
      result.arrival.legs.push(buildB2CLeg_(b2c, B2CI.DEP0_FLIGHT));
    }
    if (arr1Filled) {
      result.arrival.legs.push(buildB2CLeg_(b2c, B2CI.ARR1_FLIGHT));
    }
    if (arr2Filled) {
      result.arrival.legs.push(buildB2CLeg_(b2c, B2CI.ARR2_FLIGHT));
    }
    result.arrival.direct = (result.arrival.legs.length <= 1);

    // === رحلة العودة (حتى 3 قطع: RET1 + RET2 + RET3) ===
    var retB1Filled = String(b2c[B2CI.RET1_FLIGHT] || '').trim() !== '';
    var retB2Filled = String(b2c[B2CI.RET2_FLIGHT] || '').trim() !== '';
    var ret3Filled = String(b2c[B2CI.RET3_FLIGHT] || '').trim() !== '';

    // الترتيب الزمني: RET1 (الانطلاق) → RET2 (ترانزيت) → RET3 (الأخيرة)
    if (retB1Filled) {
      result.return.legs.push(buildB2CLeg_(b2c, B2CI.RET1_FLIGHT));
    }
    if (retB2Filled) {
      result.return.legs.push(buildB2CLeg_(b2c, B2CI.RET2_FLIGHT));
    }
    if (ret3Filled) {
      result.return.legs.push(buildB2CLeg_(b2c, B2CI.RET3_FLIGHT));
    }
    result.return.direct = (result.return.legs.length <= 1);

    // استنتاج شركة الطيران من رقم الرحلة الأولى
    var firstFlight = result.arrival.legs.length > 0 ? result.arrival.legs[0].flightNo : '';
    var guess = guessAirlineFromFlight_(firstFlight);
    result.airline.code = guess.en || '';
    result.airline.nameEn = guess.en || '';
    result.airline.nameAr = guess.ar || '';
  }

  setCache_(cacheKey, result);
  return result;
}

// بناء كائن مرحلة واحدة من صف شيت "الطيران" (B2B)
// الأعمدة متتابعة: Flight, TakeoffDate, TakeoffTime, From, To, LandDate, LandTime
function buildLeg_(flt, startIdx) {
  return {
    flightNo:     String(flt[startIdx]     || '-').trim(),
    takeoffDate:  flt[startIdx + 1],
    takeoffTime:  flt[startIdx + 2],
    from:         String(flt[startIdx + 3] || '-').trim(),
    to:           String(flt[startIdx + 4] || '-').trim(),
    landDate:     flt[startIdx + 5],
    landTime:     flt[startIdx + 6]
  };
}

// بناء كائن مرحلة واحدة من صف شيت B2C
// الأعمدة متتابعة: Flight, Date, Time, From, To, LandDate, LandTime
function buildB2CLeg_(b2c, startIdx) {
  return {
    flightNo:     String(b2c[startIdx]     || '-').trim(),
    takeoffDate:  b2c[startIdx + 1],
    takeoffTime:  b2c[startIdx + 2],
    from:         String(b2c[startIdx + 3] || '-').trim(),
    to:           String(b2c[startIdx + 4] || '-').trim(),
    landDate:     b2c[startIdx + 5],
    landTime:     b2c[startIdx + 6]
  };
}
