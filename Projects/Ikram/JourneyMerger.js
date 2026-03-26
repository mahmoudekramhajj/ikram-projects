/**
 * JourneyMerger.js
 * يدمج بيانات "رحلة الحاج 2" + "Presonal Details" في شيت "رحلة الحاج "
 * بنفس الشكل القديم (48 عمود) حتى لا تتأثر التطبيقات الأخرى
 *
 * مفتاح الربط: ApplicationId (رحلة الحاج 2) = رقم المجموعة (Presonal Details)
 * مطابقة الأفراد: Gender + IsMain ↔ الجنس + نوع الحاج
 */

// ===================== الإعدادات =====================
var MERGER_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SOURCE_JOURNEY: 'رحلة الحاج 2',
  SOURCE_PD: 'Presonal Details',
  TARGET_SHEET: 'رحلة الحاج ',  // مع المسافة في النهاية
  HEADER_ROW: [
    'BookingId', 'PackageId', 'ServiceProviderId', 'PackageYear',
    'CampName', 'ApplicantId', 'Group Number', 'Name',
    'Passport', 'Email', 'IsMain', 'Gender',
    'NationalityEn', 'NationalityAr', 'CountryResidenceEn', 'CountryResidenceAr',
    'ArrivalAirlineNameAr', 'ArrivalAirlineNameEn', 'ArrivalArrivalTime', 'ArrivalArriveCity',
    'ArrivalArriveDate', 'ArrivalDepartureCity', 'ArrivalDepartureDate', 'ArrivalDepartureTime',
    'ArrivalFlightNumber', 'ArrivalFlightType', 'ReturnAirlineNameAr', 'ReturnAirlineNameEn',
    'ReturnArrivalTime', 'ReturnArriveCity', 'ReturnArriveDate', 'ReturnDepartureCity',
    'ReturnDepartureDate', 'ReturnDepartureTime', 'ReturnFlightNumber', 'ReturnFlightType',
    'FirstHouse', 'FirstHouseStartDate', 'FirstHouseEndDate', 'LastHouse',
    'LastHouseStartDate', 'LastHouseEndDate', 'MakkahAr', 'MakkahEn',
    'MakkahShiftingAr', 'MakkahShiftingEn', 'MadinahAr', 'MadinahEn'
  ]
};

// أسماء أعمدة الرحلة والفنادق (16-47) بنفس الترتيب في كلا الشيتين
var FLIGHT_HOTEL_KEYS = [
  'ArrivalAirlineNameAr', 'ArrivalAirlineNameEn', 'ArrivalArrivalTime', 'ArrivalArriveCity',
  'ArrivalArriveDate', 'ArrivalDepartureCity', 'ArrivalDepartureDate', 'ArrivalDepartureTime',
  'ArrivalFlightNumber', 'ArrivalFlightType', 'ReturnAirlineNameAr', 'ReturnAirlineNameEn',
  'ReturnArrivalTime', 'ReturnArriveCity', 'ReturnArriveDate', 'ReturnDepartureCity',
  'ReturnDepartureDate', 'ReturnDepartureTime', 'ReturnFlightNumber', 'ReturnFlightType',
  'FirstHouse', 'FirstHouseStartDate', 'FirstHouseEndDate', 'LastHouse',
  'LastHouseStartDate', 'LastHouseEndDate', 'MakkahAr', 'MakkahEn',
  'MakkahShiftingAr', 'MakkahShiftingEn', 'MadinahAr', 'MadinahEn'
];

// خريطة الجنس عربي ← إنجليزي
var GENDER_MAP = { 'ذكر': 'Male', 'انثى': 'Female' };
// خريطة نوع الحاج ← IsMain
var ISMAIN_MAP = { 'رئيسي': '1', 'عضو عائلة': '0' };

// ===================== الدالة الرئيسية =====================

/**
 * دمج البيانات من رحلة الحاج 2 + Presonal Details → رحلة الحاج
 */
function mergeJourneyData() {
  var ss = SpreadsheetApp.openById(MERGER_CONFIG.SPREADSHEET_ID);
  var startTime = new Date();
  Logger.log('⏳ بدء عملية الدمج: ' + startTime.toLocaleString());

  // 1. قراءة البيانات المصدرية
  var journeySheet = ss.getSheetByName(MERGER_CONFIG.SOURCE_JOURNEY);
  var pdSheet = ss.getSheetByName(MERGER_CONFIG.SOURCE_PD);
  var targetSheet = ss.getSheetByName(MERGER_CONFIG.TARGET_SHEET);

  if (!journeySheet || !pdSheet || !targetSheet) {
    Logger.log('❌ خطأ: شيت مفقود — Journey2: ' + !!journeySheet + ', PD: ' + !!pdSheet + ', Target: ' + !!targetSheet);
    return;
  }

  var journeyData = journeySheet.getDataRange().getValues();
  var pdData = pdSheet.getDataRange().getValues();

  // تحديد الهيدرات
  var jHeaders = journeyData[0];
  var pdHeaders = pdData[0];

  // فهارس أعمدة رحلة الحاج 2
  var jIdx = buildHeaderIndex_(jHeaders);
  // فهارس أعمدة Presonal Details
  var pdIdx = buildHeaderIndex_(pdHeaders);

  Logger.log('📊 رحلة الحاج 2: ' + (journeyData.length - 1) + ' صف');
  Logger.log('📊 Presonal Details: ' + (pdData.length - 1) + ' صف');

  // 2. حفظ بيانات Reception Airport (أعمدة 48+) قبل الكتابة
  var receptionData = preserveReceptionData_(targetSheet);
  Logger.log('📋 بيانات Reception محفوظة: ' + Object.keys(receptionData).length + ' سجل');

  // 3. بناء Map للـ Presonal Details مجمّعة بـ رقم المجموعة
  var pdByGroup = {};
  for (var p = 1; p < pdData.length; p++) {
    var groupNo = String(pdData[p][pdIdx['رقم المجموعة']] || '').trim();
    if (!groupNo) continue;
    if (!pdByGroup[groupNo]) pdByGroup[groupNo] = [];
    pdByGroup[groupNo].push(pdData[p]);
  }

  // 4. دمج الصفوف
  var mergedRows = [];
  var unmatchedCount = 0;

  for (var j = 1; j < journeyData.length; j++) {
    var jRow = journeyData[j];
    var appId = String(jRow[jIdx['ApplicationId']] || '').trim();
    var gender = String(jRow[jIdx['Gender']] || '').trim();
    var isMain = String(jRow[jIdx['IsMain']] || '').trim();

    // البحث عن صف PD المطابق
    var pdRow = findMatchingPD_(pdByGroup[appId], gender, isMain, pdIdx);

    // بناء الصف بالشكل القديم (48 عمود)
    var merged = buildMergedRow_(jRow, jIdx, pdRow, pdIdx);
    mergedRows.push(merged);

    if (!pdRow) unmatchedCount++;
  }

  Logger.log('✅ صفوف مدموجة: ' + mergedRows.length);
  if (unmatchedCount > 0) {
    Logger.log('⚠️ صفوف بدون تطابق PD: ' + unmatchedCount);
  }

  // 5. كتابة النتيجة في الشيت الهدف
  writeToTarget_(targetSheet, mergedRows, receptionData);

  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('🎉 اكتمل الدمج في ' + elapsed + ' ثانية — ' + mergedRows.length + ' صف');
}

// ===================== الدوال المساعدة =====================

/**
 * بناء فهرس الهيدرات: اسم العمود → رقم الفهرس
 */
function buildHeaderIndex_(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) {
    idx[String(headers[i]).trim()] = i;
  }
  return idx;
}

/**
 * حفظ بيانات Reception Airport (الأعمدة 48+) من الشيت الحالي
 * المفتاح: BookingId|Passport
 */
function preserveReceptionData_(sheet) {
  var data = {};
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol <= 48) return data;

  var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var extraCols = lastCol - 48;

  for (var i = 0; i < allData.length; i++) {
    var bookingId = String(allData[i][0] || '').trim();
    var passport = String(allData[i][8] || '').trim().toUpperCase();
    if (!bookingId || !passport) continue;

    // التحقق من وجود بيانات في الأعمدة الإضافية
    var extraData = allData[i].slice(48);
    var hasData = extraData.some(function(v) { return v !== '' && v !== null && v !== undefined; });
    if (hasData) {
      data[bookingId + '|' + passport] = extraData;
    }
  }
  return data;
}

/**
 * إيجاد صف PD المطابق داخل المجموعة بناءً على Gender + IsMain
 * يزيل الصف المطابق من المصفوفة لتجنب التكرار
 */
function findMatchingPD_(pdRows, gender, isMain, pdIdx) {
  if (!pdRows || pdRows.length === 0) return null;

  // تحويل Gender الإنجليزي إلى العربي
  var genderAr = (gender === 'Male') ? 'ذكر' : 'انثى';
  // تحويل IsMain إلى نوع الحاج
  var pilgrimType = (isMain === '1') ? 'رئيسي' : 'عضو عائلة';

  for (var i = 0; i < pdRows.length; i++) {
    var pdGender = String(pdRows[i][pdIdx['الجنس']] || '').trim();
    var pdType = String(pdRows[i][pdIdx['نوع الحاج']] || '').trim();

    if (pdGender === genderAr && pdType === pilgrimType) {
      return pdRows.splice(i, 1)[0]; // إزالة وإرجاع
    }
  }

  // إذا لم يتطابق بالضبط، حاول بالجنس فقط
  for (var k = 0; k < pdRows.length; k++) {
    var pdGender2 = String(pdRows[k][pdIdx['الجنس']] || '').trim();
    if (pdGender2 === genderAr) {
      return pdRows.splice(k, 1)[0];
    }
  }

  // آخر محاولة: أي صف متبقي
  return pdRows.splice(0, 1)[0];
}

/**
 * بناء صف مدموج بالشكل القديم (48 عمود)
 */
function buildMergedRow_(jRow, jIdx, pdRow, pdIdx) {
  var row = new Array(48);

  // من رحلة الحاج 2
  row[0]  = jRow[jIdx['BookingId']] || '';           // BookingId
  row[1]  = jRow[jIdx['PackageId']] || '';           // PackageId
  row[2]  = jRow[jIdx['Service Provider']] || '';    // ServiceProviderId
  row[3]  = jRow[jIdx['PackageYear']] || '';         // PackageYear
  row[4]  = jRow[jIdx['CampName']] || '';            // CampName
  row[10] = jRow[jIdx['IsMain']] || '';              // IsMain
  row[11] = jRow[jIdx['Gender']] || '';              // Gender
  row[12] = jRow[jIdx['NationalityEn']] || '';       // NationalityEn
  row[14] = jRow[jIdx['CountryResidenceEn']] || '';  // CountryResidenceEn

  // Group Number = ApplicationId
  row[6] = jRow[jIdx['ApplicationId']] || '';

  // من Presonal Details
  if (pdRow) {
    row[5]  = pdRow[pdIdx['الرقم التسلسلي']] || '';                    // ApplicantId
    row[7]  = buildFullName_(pdRow, pdIdx);                             // Name
    row[8]  = pdRow[pdIdx['رقم جواز السفر']] || '';                    // Passport
    row[9]  = pdRow[pdIdx['البريد الإلكتروني']] || '';                 // Email
    row[13] = pdRow[pdIdx['الجنسية']] || '';                           // NationalityAr
    row[15] = pdRow[pdIdx['بلد الإقامة']] || '';                       // CountryResidenceAr
  } else {
    row[5] = row[7] = row[8] = row[9] = row[13] = row[15] = '';
  }

  // أعمدة الطيران والفنادق (16-47) — نفس الترتيب
  for (var f = 0; f < FLIGHT_HOTEL_KEYS.length; f++) {
    var key = FLIGHT_HOTEL_KEYS[f];
    row[16 + f] = (jIdx[key] !== undefined) ? (jRow[jIdx[key]] || '') : '';
  }

  return row;
}

/**
 * بناء الاسم الكامل من الاسم الأول + اسم العائلة (إنجليزي)
 */
function buildFullName_(pdRow, pdIdx) {
  var first = String(pdRow[pdIdx['الاسم الأول (الإنجليزية)']] || '').trim();
  var last = String(pdRow[pdIdx['اسم العائلة (الإنجليزية)']] || '').trim();
  if (first && last) return first + ' ' + last;
  if (first) return first;
  if (last) return last;

  // إذا لم يوجد إنجليزي، استخدم العربي
  var firstAr = String(pdRow[pdIdx['الاسم الأول (العربية)']] || '').trim();
  var lastAr = String(pdRow[pdIdx['اسم العائلة (العربية)']] || '').trim();
  if (firstAr && lastAr) return firstAr + ' ' + lastAr;
  return firstAr || lastAr || '';
}

/**
 * كتابة البيانات المدموجة في الشيت الهدف مع استعادة بيانات Reception
 */
function writeToTarget_(sheet, mergedRows, receptionData) {
  if (mergedRows.length === 0) {
    Logger.log('⚠️ لا توجد صفوف للكتابة');
    return;
  }

  // مسح المحتوى القديم (الأعمدة 1-48 فقط) مع الإبقاء على الهيدر
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, Math.max(lastCol, 48)).clearContent();
  }

  // كتابة الهيدر
  sheet.getRange(1, 1, 1, 48).setValues([MERGER_CONFIG.HEADER_ROW]);

  // كتابة البيانات (48 عمود)
  sheet.getRange(2, 1, mergedRows.length, 48).setValues(mergedRows);

  // استعادة بيانات Reception Airport
  if (Object.keys(receptionData).length > 0) {
    var maxExtraCols = 0;
    for (var key in receptionData) {
      maxExtraCols = Math.max(maxExtraCols, receptionData[key].length);
    }

    if (maxExtraCols > 0) {
      var receptionMatrix = [];
      var restored = 0;

      for (var r = 0; r < mergedRows.length; r++) {
        var bookingId = String(mergedRows[r][0] || '').trim();
        var passport = String(mergedRows[r][8] || '').trim().toUpperCase();
        var rKey = bookingId + '|' + passport;
        var extra = receptionData[rKey];

        if (extra) {
          // تأكد من أن المصفوفة بالطول الصحيح
          while (extra.length < maxExtraCols) extra.push('');
          receptionMatrix.push(extra);
          restored++;
        } else {
          receptionMatrix.push(new Array(maxExtraCols).fill(''));
        }
      }

      sheet.getRange(2, 49, mergedRows.length, maxExtraCols).setValues(receptionMatrix);
      Logger.log('📋 تم استعادة ' + restored + ' سجل Reception');
    }
  }
}

// ===================== القائمة المخصصة =====================

/**
 * إضافة زر الدمج في القائمة
 */
function onOpen_JourneyMerger() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 دمج البيانات')
    .addItem('دمج رحلة الحاج الآن', 'mergeJourneyData')
    .addToUi();
}
