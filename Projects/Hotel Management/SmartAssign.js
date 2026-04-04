// ============================================================
// دوال الاختبار — تُحذف بعد التأكد من عمل النظام
// ============================================================

function testSmartAssign() {
  var result = smartAssignHotel('Jabal Omar Hyatt Regency Hotel', 'Makkah');
  Logger.log(JSON.stringify(result, null, 2));
}

// اختبار مقسّم — كل دالة تعالج مجموعة فنادق لتجنب timeout
function testBatch1() { return testHotelBatch_(0, 9); }   // أول 9 فنادق
function testBatch2() { return testHotelBatch_(9, 18); }  // 9 التالية
function testBatch3() { return testHotelBatch_(18, 30); } // الباقي

function testHotelBatch_(startIdx, endIdx) {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var J = HOTEL_CONFIG.JOURNEY_COLS;

  var hotelCityMap = {};
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var med = String(row[J.MADINAH_EN] || '').trim();
    if (med && med !== 'NULL') hotelCityMap[med] = 'Madina';
    var mak = String(row[J.MAKKAH_EN] || '').trim();
    if (mak && mak !== 'NULL') hotelCityMap[mak] = 'Makkah';
    var shift = String(row[J.MAKKAH_SHIFT_EN] || '').trim();
    if (shift && shift !== 'NULL') hotelCityMap[shift] = 'Makkah Shifting';
  }

  var hotelNames = Object.keys(hotelCityMap);
  var end = Math.min(endIdx, hotelNames.length);
  Logger.log('═══ Batch ' + startIdx + '-' + end + ' من ' + hotelNames.length + ' فندق ═══');

  var totalWritten = 0, totalPilgrims = 0;
  for (var h = startIdx; h < end; h++) {
    var name = hotelNames[h];
    var city = hotelCityMap[name];
    var r = smartAssignHotel(name, city);
    totalWritten += r.totalWritten || 0;
    totalPilgrims += r.totalPilgrims || 0;
    Logger.log((r.success ? '✅' : '🚨') + ' ' + city + ' | ' + name + ': ' +
      (r.totalWritten||0) + '/' + (r.totalPilgrims||0) + ' | عائلات:' + (r.familiesAssigned||0) +
      ' مشاركين:' + (r.sharedRoomsAssigned||0) + ' جزئية:' + (r.partialRooms||0) +
      ' مرشد:' + (r.guideSharersExcluded||0) + ' بدون غرفة:' + (r.noRoomFamilies||0));
  }
  Logger.log('المجموع: ' + totalWritten + '/' + totalPilgrims);
}

function testAllHotels() {
  var results = smartAssignAll();
  Logger.log('═══ نتائج التوزيع لجميع الفنادق ═══');
  Logger.log('عدد الفنادق: ' + results.totalHotels);

  var totalPilgrims = 0, totalWritten = 0, totalGuideExcluded = 0;
  var totalFamilies = 0, totalShared = 0, totalPartial = 0, totalNoRoom = 0;
  var errors = [];
  var partialDetails = [];

  for (var i = 0; i < results.results.length; i++) {
    var r = results.results[i];
    if (!r.success) {
      errors.push(r.hotel + ': ' + r.error);
      continue;
    }
    totalPilgrims += r.totalPilgrims;
    totalWritten += r.totalWritten;
    totalGuideExcluded += r.guideSharersExcluded || 0;
    totalFamilies += r.familiesAssigned;
    totalShared += r.sharedRoomsAssigned;
    totalPartial += r.partialRooms;
    totalNoRoom += r.noRoomFamilies;

    Logger.log(r.city + ' | ' + r.hotel + ': ' + r.totalWritten + '/' + r.totalPilgrims +
      ' حاج | عائلات: ' + r.familiesAssigned + ' | مشاركين: ' + r.sharedRoomsAssigned +
      ' | جزئية: ' + r.partialRooms + ' | مرشد: ' + (r.guideSharersExcluded || 0) +
      ' | بدون غرفة: ' + r.noRoomFamilies);

    if (r.partialDetails && r.partialDetails.length > 0) {
      r.partialDetails.forEach(function(p) {
        partialDetails.push(r.hotel + ' → ' + p.roomId + ': ' + p.occupants + '/' + p.capacity + ' ' + p.gender);
      });
    }
  }

  Logger.log('\n═══ الملخص الشامل ═══');
  Logger.log('إجمالي الحجاج: ' + totalPilgrims);
  Logger.log('تم تسكينهم: ' + totalWritten);
  Logger.log('مشاركين مرشد (يدوي): ' + totalGuideExcluded);
  Logger.log('غرف عائلات: ' + totalFamilies);
  Logger.log('غرف مشاركين (ممتلئة): ' + totalShared);
  Logger.log('غرف جزئية: ' + totalPartial);
  Logger.log('عائلات بدون غرف: ' + totalNoRoom);
  Logger.log('غير مسكّنين: ' + (totalPilgrims - totalWritten - totalGuideExcluded));

  if (errors.length > 0) {
    Logger.log('\n🚨 أخطاء:');
    errors.forEach(function(e) { Logger.log('  ' + e); });
  }

  if (partialDetails.length > 0) {
    Logger.log('\n⚠️ غرف جزئية:');
    partialDetails.forEach(function(d) { Logger.log('  ' + d); });
  }
}

function testFullAudit() {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var J = HOTEL_CONFIG.JOURNEY_COLS;

  // جمع كل الفنادق
  var hotelCityMap = {};
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var med = String(row[J.MADINAH_EN] || '').trim();
    if (med && med !== 'NULL') hotelCityMap[med] = 'Madina';
    var mak = String(row[J.MAKKAH_EN] || '').trim();
    if (mak && mak !== 'NULL') hotelCityMap[mak] = 'Makkah';
    var shift = String(row[J.MAKKAH_SHIFT_EN] || '').trim();
    if (shift && shift !== 'NULL') hotelCityMap[shift] = 'Makkah Shifting';
  }

  Logger.log('═══ تدقيق جميع الفنادق ═══');
  var totalViolations = 0;
  var totalOverCapacity = 0;
  var grandAssigned = 0, grandUnassigned = 0;

  for (var hotelName in hotelCityMap) {
    var hotelSheet = getHotelSheet_(ss, hotelName);
    if (!hotelSheet) {
      Logger.log('⚠️ ' + hotelName + ': شيت غير موجود');
      continue;
    }

    var data = hotelSheet.getDataRange().getValues();
    var assigned = 0, unassigned = 0;
    var roomGroups = {};

    for (var i = 1; i < data.length; i++) {
      var rgId = String(data[i][16] || '');
      var gender = String(data[i][3]);
      var groupNum = String(data[i][5]);

      if (rgId) {
        assigned++;
        if (!roomGroups[rgId]) roomGroups[rgId] = { males: 0, females: 0, groups: {} };
        if (gender === 'Male' || gender === 'ذكر') roomGroups[rgId].males++;
        else roomGroups[rgId].females++;
        roomGroups[rgId].groups[groupNum] = true;
      } else {
        unassigned++;
      }
    }

    grandAssigned += assigned;
    grandUnassigned += unassigned;

    var violations = 0;
    var overCap = 0;

    for (var rgId in roomGroups) {
      var rg = roomGroups[rgId];
      var groupCount = Object.keys(rg.groups).length;
      var total = rg.males + rg.females;
      var isMixed = rg.males > 0 && rg.females > 0;
      var isMultiGroup = groupCount > 1;

      // مخالفة: مختلط + مجموعات مختلفة
      if (isMixed && isMultiGroup) {
        violations++;
        Logger.log('🚨 ' + hotelName + ' → ' + rgId + ': مختلط (' + rg.males + 'ذ+' + rg.females + 'أ) من ' + groupCount + ' مجموعات');
      }

      // تجاوز السعة
      var typeChar = rgId.replace(/\d+$/, '').slice(-1);
      var cap = typeChar === 'T' ? 3 : typeChar === 'D' ? 2 : 4;
      if (total > cap) {
        overCap++;
        Logger.log('🚨 ' + hotelName + ' → ' + rgId + ': ' + total + ' أشخاص في غرفة سعتها ' + cap);
      }
    }

    var roomCount = Object.keys(roomGroups).length;
    var status = violations === 0 && overCap === 0 ? '✅' : '🚨';
    Logger.log(status + ' ' + hotelCityMap[hotelName] + ' | ' + hotelName +
      ': ' + assigned + '/' + (data.length - 1) + ' مسكّنين | ' + roomCount + ' غرفة' +
      (violations > 0 ? ' | ' + violations + ' مخالفة جنس' : '') +
      (overCap > 0 ? ' | ' + overCap + ' تجاوز سعة' : ''));

    totalViolations += violations;
    totalOverCapacity += overCap;
  }

  Logger.log('\n═══ الملخص الشامل ═══');
  Logger.log('إجمالي مسكّنين: ' + grandAssigned);
  Logger.log('إجمالي غير مسكّنين: ' + grandUnassigned);
  Logger.log('مخالفات جنس: ' + totalViolations);
  Logger.log('تجاوزات سعة: ' + totalOverCapacity);
  Logger.log(totalViolations === 0 && totalOverCapacity === 0 ? '✅ النظام سليم 100%' : '🚨 يوجد مشاكل!');
}

function testAudit() {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var hotelSheet = getHotelSheet_(ss, 'Jabal Omar Hyatt Regency Hotel');
  if (!hotelSheet) { Logger.log('شيت الفندق غير موجود!'); return; }

  var data = hotelSheet.getDataRange().getValues();
  var headers = data[0];
  Logger.log('أعمدة الشيت: ' + headers.length);

  // إحصائيات التسكين
  var assigned = 0, unassigned = 0, arrived = 0;
  var roomGroups = {}; // roomGroupId → [members]
  var genderMix = {}; // roomGroupId → { Male: n, Female: n, groups: {} }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var applicantId = String(row[0]);
    var name = String(row[1]);
    var gender = String(row[3]);
    var groupNum = String(row[5]);
    var roomType = String(row[15]);
    var roomGroupId = String(row[16] || '');
    var checkinStatus = String(row[18] || '');

    if (checkinStatus === 'arrived') arrived++;

    if (roomGroupId) {
      assigned++;
      if (!roomGroups[roomGroupId]) roomGroups[roomGroupId] = [];
      roomGroups[roomGroupId].push({ name: name, gender: gender, group: groupNum, roomType: roomType });

      if (!genderMix[roomGroupId]) genderMix[roomGroupId] = { Male: 0, Female: 0, groups: {} };
      genderMix[roomGroupId][gender] = (genderMix[roomGroupId][gender] || 0) + 1;
      genderMix[roomGroupId].groups[groupNum] = true;
    } else {
      unassigned++;
    }
  }

  Logger.log('\n═══ تدقيق فندق HYT ═══');
  Logger.log('إجمالي الحجاج: ' + (data.length - 1));
  Logger.log('مسكّنين: ' + assigned);
  Logger.log('غير مسكّنين: ' + unassigned);
  Logger.log('وصلوا (arrived): ' + arrived);
  Logger.log('عدد الغرف المستخدمة: ' + Object.keys(roomGroups).length);

  // فحص القواعد
  var violations = [];

  for (var rgId in roomGroups) {
    var members = roomGroups[rgId];
    var mix = genderMix[rgId];
    var groupCount = Object.keys(mix.groups).length;
    var isMixedGender = mix.Male > 0 && mix.Female > 0;
    var isMultiGroup = groupCount > 1;

    // القاعدة: لو مختلط الجنس + مجموعات متعددة = مخالفة
    if (isMixedGender && isMultiGroup) {
      violations.push('🚨 غرفة ' + rgId + ': مختلطة الجنس + مجموعات مختلفة! (' +
        mix.Male + ' ذكور + ' + mix.Female + ' إناث من ' + groupCount + ' مجموعات)');
      members.forEach(function(m) {
        violations.push('   → ' + m.name + ' | ' + m.gender + ' | مجموعة ' + m.group);
      });
    }

    // فحص السعة: هل تجاوزت سعة الغرفة؟
    var capacity = 4; // default
    if (rgId.indexOf('T') !== -1 && rgId.match(/[A-Z]T\d/)) capacity = 3;
    if (rgId.indexOf('D') !== -1 && rgId.match(/[A-Z]D\d/)) capacity = 2;
    // استخدم الحرف قبل الأرقام
    var typeChar = rgId.replace(/\d+$/, '').slice(-1);
    if (typeChar === 'T') capacity = 3;
    else if (typeChar === 'D') capacity = 2;
    else capacity = 4;

    if (members.length > capacity) {
      violations.push('🚨 غرفة ' + rgId + ': ' + members.length + ' أشخاص في غرفة سعتها ' + capacity + '!');
    }
  }

  // فحص العائلات (نفس المجموعة = مختلط مسموح)
  Logger.log('\n═══ فحص القواعد ═══');
  if (violations.length === 0) {
    Logger.log('✅ لا مخالفات — كل الغرف تلتزم بالقواعد');
  } else {
    violations.forEach(function(v) { Logger.log(v); });
  }

  // عرض تفاصيل الغرف (أول 10)
  Logger.log('\n═══ عينة من الغرف ═══');
  var count = 0;
  for (var rgId in roomGroups) {
    if (count >= 10) break;
    var members = roomGroups[rgId];
    var mix = genderMix[rgId];
    var groupCount = Object.keys(mix.groups).length;
    var typeLabel = groupCount === 1 ? 'عائلة' : 'مشاركين';
    var genderLabel = (mix.Male > 0 && mix.Female > 0) ? 'مختلط' : (mix.Male > 0 ? 'ذكور' : 'إناث');

    Logger.log(rgId + ' (' + members.length + ' أشخاص | ' + typeLabel + ' | ' + genderLabel + '):');
    members.forEach(function(m) {
      Logger.log('  ' + m.name + ' | ' + m.gender + ' | مجموعة ' + m.group);
    });
    count++;
  }

  // الـ 9 غير المسكّنين (مشاركين مرشد)
  Logger.log('\n═══ غير المسكّنين ═══');
  for (var i = 1; i < data.length; i++) {
    if (!data[i][16]) {
      Logger.log('  ' + data[i][1] + ' | ' + data[i][3] + ' | مجموعة ' + data[i][5]);
    }
  }
}

function testWriteDebug() {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var hotelName = 'Jabal Omar Hyatt Regency Hotel';

  // 1. تحقق من شيت الفندق
  var hotelSheet = getHotelSheet_(ss, hotelName);
  Logger.log('hotelSheet found: ' + !!hotelSheet);

  if (!hotelSheet) {
    // حاول بالاختصار
    var abbr = sanitizeSheetName_(hotelName);
    Logger.log('abbreviation: ' + abbr);
    hotelSheet = ss.getSheetByName(abbr);
    Logger.log('sheet by abbr: ' + !!hotelSheet);
  }

  if (hotelSheet) {
    var lastRow = hotelSheet.getLastRow();
    var lastCol = hotelSheet.getLastColumn();
    Logger.log('شيت الفندق: ' + hotelSheet.getName() + ' | صفوف: ' + lastRow + ' | أعمدة: ' + lastCol);

    if (lastRow > 1) {
      // أول 3 صفوف
      var sample = hotelSheet.getRange(1, 1, Math.min(4, lastRow), Math.min(5, lastCol)).getValues();
      Logger.log('Headers + first rows:');
      sample.forEach(function(r, i) { Logger.log('  [' + i + '] ' + JSON.stringify(r)); });

      // عمود BookingId (الأول)
      var bid1 = String(hotelSheet.getRange(2, 1).getValue());
      Logger.log('First BookingId in hotel sheet: "' + bid1 + '"');
    }
  }

  // 2. تحقق من bookingId في رحلة الحاج
  var journeySheet = ss.getSheetByName('رحلة الحاج ');
  var journeyData = journeySheet.getDataRange().getValues();
  var hyattBids = [];
  for (var i = 1; i < journeyData.length; i++) {
    if (journeyData[i][43] === hotelName) {
      hyattBids.push(String(journeyData[i][5])); // ApplicantId = col 5 (= BookingId in hotel sheets)
    }
  }
  Logger.log('BookingIds from journey (first 5): ' + JSON.stringify(hyattBids.slice(0, 5)));

  // 3. مقارنة
  if (hotelSheet && hotelSheet.getLastRow() > 1) {
    var hotelBids = hotelSheet.getRange(2, 1, hotelSheet.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]); });
    Logger.log('BookingIds from hotel sheet (first 5): ' + JSON.stringify(hotelBids.slice(0, 5)));

    // هل يتطابقون؟
    var matches = 0;
    for (var i = 0; i < hyattBids.length; i++) {
      if (hotelBids.indexOf(hyattBids[i]) !== -1) matches++;
    }
    Logger.log('تطابقات: ' + matches + ' من ' + hyattBids.length);
  }
}

function testPrepare() {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var hotelName = 'Jabal Omar Hyatt Regency Hotel';
  var hotelCity = 'Makkah';

  // 1. تحقق من وجود الشيتات
  var journeySheet = ss.getSheetByName('رحلة الحاج ');
  Logger.log('Journey sheet: ' + (journeySheet ? 'موجود (' + journeySheet.getLastRow() + ' صف)' : 'غير موجود!'));

  var roomTypeSheet = ss.getSheetByName('Room Type');
  Logger.log('Room Type sheet: ' + (roomTypeSheet ? 'موجود (' + roomTypeSheet.getLastRow() + ' صف)' : 'غير موجود!'));

  var guideSheet = ss.getSheetByName('Guide Rabih');
  Logger.log('Guide Rabih sheet: ' + (guideSheet ? 'موجود (' + guideSheet.getLastRow() + ' صف)' : 'غير موجود!'));

  var mapSheet = ss.getSheetByName('Room Mapping');
  Logger.log('Room Mapping sheet: ' + (mapSheet ? 'موجود (' + mapSheet.getLastRow() + ' صف)' : 'غير موجود!'));

  // 2. عدّ الحجاج لهذا الفندق
  if (journeySheet) {
    var data = journeySheet.getDataRange().getValues();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var makkahEn = data[i][43]; // MAKKAH_EN
      if (makkahEn === hotelName) count++;
    }
    Logger.log('حجاج ' + hotelName + ': ' + count);
  }

  // 3. عدّ الغرف في Room Mapping
  if (mapSheet) {
    var mapData = mapSheet.getDataRange().getValues();
    var roomCount = 0;
    for (var i = 1; i < mapData.length; i++) {
      if (String(mapData[i][0]).indexOf('Hyatt') !== -1) roomCount++;
    }
    Logger.log('غرف Hyatt في Room Mapping: ' + roomCount);
  }
}

/**
 * Smart Room Assignment — التسكين الذكي
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.1 — يستبدل النظام القديم (4-phase AutoAssign + PreAssign)
 *
 * نقاط الدخول:
 *   smartAssignHotel(hotelName, hotelCity)  ← توزيع ذكي — فندق واحد
 *   smartAssignAll()                        ← توزيع ذكي — كل الفنادق
 *   updateAssignment(hotelName, hotelCity)  ← اكتمل العدد — تحديث التغييرات
 *   getGuideStats(hotelName, hotelCity)     ← إحصائية المرشدين
 */

// ============================================================
// GUIDE RABIH: بناء خريطة المرشدين من شيت Guide Rabih
// Returns: { passport → { guide, country, nationality } }
// ============================================================

function buildGuideRabihMap_(ss) {
  var cached = getCachedData_('guideRabihMap');
  if (cached) return cached;

  var map = {};
  try {
    var sheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.GUIDE_RABIH);
    if (!sheet) return map;

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    var data = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
    var G = HOTEL_CONFIG.GUIDE_RABIH_COLS;

    for (var i = 0; i < data.length; i++) {
      var passport = String(data[i][G.PASSPORT]).trim().toUpperCase();
      if (!passport) continue;

      map[passport] = {
        guide: String(data[i][G.GUIDE_NAME] || '').trim(),
        country: String(data[i][G.COUNTRY] || '').trim(),
        nationality: String(data[i][G.NATIONALITY] || '').trim()
      };
    }
  } catch (e) {
    Logger.log('buildGuideRabihMap_ ERROR: ' + e.toString());
  }
  setCachedData_('guideRabihMap', map);
  return map;
}

// ============================================================
// المرحلة 0: التجهيز — جمع البيانات + التصنيف
// ============================================================

function prepareAssignment_(ss, hotelName, hotelCity) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var R = HOTEL_CONFIG.ROOM_COLS;

  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();

  var roomTypeSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_TYPE);
  var roomTypeData = roomTypeSheet.getDataRange().getValues();

  var guideMap = buildGuideRabihMap_(ss);

  // بناء خريطة Room Type: groupNumber → info
  var groupInfoMap = {};
  for (var i = 1; i < roomTypeData.length; i++) {
    var row = roomTypeData[i];
    var gn = String(row[R.GROUP_NUMBER]).trim();
    if (!gn) continue;

    groupInfoMap[gn] = {
      groupNumber: gn,
      packageId: String(row[R.PACKAGE_ID] || ''),
      packageName: String(row[R.PACKAGE_NAME] || ''),
      pilgrims: Number(row[R.PILGRIMS]) || 0,
      male: Number(row[R.MALE]) || 0,
      female: Number(row[R.FEMALE]) || 0,
      nationality: String(row[R.NATIONALITY] || ''),
      tourGuide: String(row[R.TOUR_GUIDE] || ''),
      roomTypeMed: String(row[R.ROOM_TYPE_MED] || ''),
      fullRoomsMed: Number(row[R.FULL_ROOMS_MED]) || 0,
      sharedBedsMed: Number(row[R.SHARED_BEDS_MED]) || 0,
      roomTypeMak1: String(row[R.ROOM_TYPE_MAK1] || ''),
      fullRoomsMak1: Number(row[R.FULL_ROOMS_MAK1]) || 0,
      sharedBedsMak1: Number(row[R.SHARED_BEDS_MAK1]) || 0,
      roomTypeMak2: String(row[R.ROOM_TYPE_MAK2] || ''),
      fullRoomsMak2: Number(row[R.FULL_ROOMS_MAK2]) || 0,
      sharedBedsMak2: Number(row[R.SHARED_BEDS_MAK2]) || 0
    };
  }

  // جمع حجاج هذا الفندق
  var pilgrims = [];
  var normHotel = normalizeHotelName_(hotelName);

  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var pilgrimHotel = getHotelName_(row, hotelCity);
    if (!pilgrimHotel) continue;
    if (normalizeHotelName_(pilgrimHotel) !== normHotel) continue;

    var passport = String(row[J.PASSPORT]).trim().toUpperCase();
    var groupNumber = String(row[J.GROUP_NUMBER]).trim();
    var gender = String(row[J.GENDER]).trim();
    var guideInfo = guideMap[passport] || {};

    // ApplicantId (عمود 5) = المعرّف المستخدم في شيتات الفنادق
    pilgrims.push({
      bookingId: String(row[J.BOOKING_ID]),
      passport: passport,
      name: String(row[J.NAME]),
      gender: gender,
      groupNumber: groupNumber,
      packageId: String(row[J.PACKAGE_ID]),
      nationalityEn: String(row[J.NATIONALITY_EN] || ''),
      country: guideInfo.country || String(row[J.COUNTRY_RESIDENCE] || ''),
      guide: guideInfo.guide || ''
    });
  }

  // تجميع الحجاج حسب المجموعة
  var groups = {};
  for (var i = 0; i < pilgrims.length; i++) {
    var p = pilgrims[i];
    if (!groups[p.groupNumber]) groups[p.groupNumber] = [];
    groups[p.groupNumber].push(p);
  }

  // تصنيف المجموعات: عائلة أو مشاركين
  var families = [];
  var sharers = [];
  var guideSharers = 0; // عداد المستبعدين

  for (var gn in groups) {
    var members = groups[gn];
    var info = groupInfoMap[gn];
    if (!info) continue;

    var roomType = getRoomType_(info, hotelCity);
    var capacity = getRoomCapacity_(roomType);
    if (!capacity) continue;

    // المرشد: من Guide Rabih أو من Room Type
    var guide = '';
    for (var m = 0; m < members.length; m++) {
      if (members[m].guide) { guide = members[m].guide; break; }
    }
    if (!guide) guide = info.tourGuide || '';

    var classification = {
      groupNumber: gn,
      members: members,
      roomType: roomType,
      capacity: capacity,
      guide: guide,
      country: members[0].country || info.nationality || ''
    };

    if (members.length === capacity) {
      // عائلة: عدد الأفراد = سعة الغرفة — تُسكَّن تلقائياً (حتى لو عندهم مرشد)
      families.push(classification);
    } else {
      // مشاركين
      if (guide) {
        // مشاركين مع مرشد — المرشد يوزّعهم يدوياً
        guideSharers += members.length;
        continue;
      }
      sharers.push(classification);
    }
  }

  // قراءة الغرف المتاحة من Room Mapping
  var roomMapping = getAvailableRooms_(ss, hotelName, hotelCity);

  return {
    families: families,
    sharers: sharers,
    roomMapping: roomMapping,
    guideMap: guideMap,
    groupInfoMap: groupInfoMap,
    pilgrimCount: pilgrims.length,
    guideSharersExcluded: guideSharers
  };
}

// ============================================================
// جلب الغرف المتاحة من Room Mapping
// ============================================================

function getAvailableRooms_(ss, hotelName, hotelCity) {
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  if (!mapSheet) return { quad: [], triple: [], double: [] };

  var data = mapSheet.getDataRange().getValues();
  var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;
  var normHotel = normalizeHotelName_(hotelName);

  var rooms = { quad: [], triple: [], double: [] };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (normalizeHotelName_(String(row[MC.HOTEL_NAME])) !== normHotel) continue;
    if (String(row[MC.STATUS]) === 'Occupied') continue;

    var cap = Number(row[MC.CAPACITY]) || 4;
    var room = {
      internalId: String(row[MC.INTERNAL_ID]),
      roomType: String(row[MC.ROOM_TYPE]),
      capacity: cap,
      actualRoomNo: String(row[MC.ACTUAL_ROOM_NO] || ''),
      rowIndex: i,
      occupants: []
    };

    if (cap === 4) rooms.quad.push(room);
    else if (cap === 3) rooms.triple.push(room);
    else if (cap === 2) rooms.double.push(room);
  }

  return rooms;
}

// ============================================================
// المرحلة 1: تسكين العائلات
// ============================================================

function assignFamilies_(families, rooms) {
  var assigned = [];
  var noRoom = [];

  for (var i = 0; i < families.length; i++) {
    var fam = families[i];
    var pool = getPoolByCapacity_(rooms, fam.capacity);

    if (pool.length === 0) {
      var upgraded = upgradeRoom_(rooms, fam.capacity);
      if (upgraded) {
        pool.push(upgraded);
      } else {
        noRoom.push(fam);
        continue;
      }
    }

    var room = pool.shift();
    room.occupants = fam.members.slice();
    assigned.push({ room: room, members: fam.members, type: 'family' });
  }

  return { assigned: assigned, noRoom: noRoom };
}

// ============================================================
// المرحلة 2: تسكين المشاركين (بدون مرشد)
// ============================================================

function assignSharers_(sharers, rooms) {
  // فصل بالجنس
  var males = [];
  var females = [];

  for (var i = 0; i < sharers.length; i++) {
    var group = sharers[i];
    for (var j = 0; j < group.members.length; j++) {
      var m = group.members[j];
      m.roomType = group.roomType;
      m.requestedCapacity = group.capacity;
      m.groupSize = group.members.length;

      if (m.gender === 'Male' || m.gender === 'ذكر') {
        males.push(m);
      } else {
        females.push(m);
      }
    }
  }

  // ترتيب: مجموعات أكبر أولاً
  males.sort(sortSharers_);
  females.sort(sortSharers_);

  var maleResult = fillRoomsForGender_(males, rooms);
  var femaleResult = fillRoomsForGender_(females, rooms);

  return {
    assigned: maleResult.assigned.concat(femaleResult.assigned),
    partialRooms: maleResult.partialRooms.concat(femaleResult.partialRooms)
  };
}

function sortSharers_(a, b) {
  if (b.groupSize !== a.groupSize) return b.groupSize - a.groupSize;
  if (a.country !== b.country) return a.country.localeCompare(b.country);
  if (a.guide !== b.guide) return (a.guide || '').localeCompare(b.guide || '');
  return 0;
}

function fillRoomsForGender_(pilgrims, rooms) {
  var assigned = [];
  var partialRooms = [];
  var remaining = pilgrims.slice();

  while (remaining.length > 0) {
    var current = remaining.shift();
    var sameGroup = [current];

    // اجمع باقي أفراد نفس المجموعة
    for (var i = remaining.length - 1; i >= 0; i--) {
      if (remaining[i].groupNumber === current.groupNumber) {
        sameGroup.push(remaining.splice(i, 1)[0]);
      }
    }

    var room = findBestRoom_(rooms, sameGroup.length);
    if (!room) {
      room = findAnyRoom_(rooms);
      if (!room) continue; // لا غرف — يُسجّل في التقرير
    }

    // أدخل المجموعة
    for (var i = 0; i < sameGroup.length && room.occupants.length < room.capacity; i++) {
      room.occupants.push(sameGroup[i]);
    }

    // أكمل بالتفضيلات لو ناقصة
    if (room.occupants.length < room.capacity) {
      var needed = room.capacity - room.occupants.length;
      var fillers = findFillers_(remaining, needed, room.occupants[0]);
      for (var i = 0; i < fillers.length; i++) {
        room.occupants.push(fillers[i]);
      }
    }

    if (room.occupants.length >= room.capacity) {
      assigned.push({ room: room, members: room.occupants.slice(), type: 'shared' });
    } else {
      partialRooms.push({ room: room, members: room.occupants.slice(), type: 'partial' });
    }
  }

  return { assigned: assigned, partialRooms: partialRooms };
}

// ============================================================
// البحث بالتفضيلات — إكمال الغرفة
// لا يُعدّل candidates مباشرة — يُرجع النتائج ويزيلها بأمان
// ============================================================

function findFillers_(candidates, needed, reference) {
  if (!candidates || candidates.length === 0 || needed <= 0 || !reference) return [];

  var result = [];
  var usedIndices = {};

  // تفضيل 1: نفس المرشد + نفس البلد + منفرد
  for (var i = 0; i < candidates.length && result.length < needed; i++) {
    if (usedIndices[i]) continue;
    var c = candidates[i];
    if (c.guide && c.guide === reference.guide &&
        c.country && c.country === reference.country &&
        c.groupSize === 1) {
      result.push(c);
      usedIndices[i] = true;
    }
  }

  // تفضيل 2: نفس المرشد
  for (var i = 0; i < candidates.length && result.length < needed; i++) {
    if (usedIndices[i]) continue;
    var c = candidates[i];
    if (c.guide && c.guide === reference.guide) {
      result.push(c);
      usedIndices[i] = true;
    }
  }

  // تفضيل 3: نفس البلد
  for (var i = 0; i < candidates.length && result.length < needed; i++) {
    if (usedIndices[i]) continue;
    var c = candidates[i];
    if (c.country && c.country === reference.country) {
      result.push(c);
      usedIndices[i] = true;
    }
  }

  // تفضيل 4: أي شخص (منفردين أولاً)
  var remainingCandidates = [];
  for (var i = 0; i < candidates.length; i++) {
    if (!usedIndices[i]) remainingCandidates.push({ index: i, pilgrim: candidates[i] });
  }
  remainingCandidates.sort(function(a, b) { return a.pilgrim.groupSize - b.pilgrim.groupSize; });

  for (var i = 0; i < remainingCandidates.length && result.length < needed; i++) {
    result.push(remainingCandidates[i].pilgrim);
    usedIndices[remainingCandidates[i].index] = true;
  }

  // إزالة من candidates بترتيب عكسي (آمن)
  var indicesToRemove = Object.keys(usedIndices).map(Number).sort(function(a, b) { return b - a; });
  for (var i = 0; i < indicesToRemove.length; i++) {
    candidates.splice(indicesToRemove[i], 1);
  }

  return result;
}

// ============================================================
// المرحلة 3: التحسين — ملء الغرف الناقصة
// ============================================================

function optimizeRooms_(partialRooms) {
  if (partialRooms.length < 2) return partialRooms;

  // رتّب: الأكثر امتلاءً أولاً
  partialRooms.sort(function(a, b) {
    return b.room.occupants.length - a.room.occupants.length;
  });

  for (var i = 0; i < partialRooms.length; i++) {
    var target = partialRooms[i];
    var needed = target.room.capacity - target.room.occupants.length;
    if (needed <= 0) continue;

    var targetGender = target.room.occupants.length > 0 ?
      (target.room.occupants[0].gender || '') : '';

    for (var j = partialRooms.length - 1; j > i && needed > 0; j--) {
      var source = partialRooms[j];
      if (source.room.occupants.length === 0) continue;

      var sourceGender = source.room.occupants.length > 0 ?
        (source.room.occupants[0].gender || '') : '';

      // نفس الجنس فقط
      if (targetGender && sourceGender && targetGender !== sourceGender) continue;

      // لا نفصل مجموعة — ننقل مجموعات كاملة فقط
      var sourceGroups = groupByGroupNumber_(source.room.occupants);
      for (var gn in sourceGroups) {
        if (needed <= 0) break;
        var groupMembers = sourceGroups[gn];
        if (groupMembers.length <= needed) {
          for (var k = 0; k < groupMembers.length; k++) {
            target.room.occupants.push(groupMembers[k]);
            var idx = source.room.occupants.indexOf(groupMembers[k]);
            if (idx !== -1) source.room.occupants.splice(idx, 1);
          }
          needed -= groupMembers.length;
        }
      }
    }
  }

  // حذف الغرف الفاضية
  var result = [];
  for (var i = 0; i < partialRooms.length; i++) {
    if (partialRooms[i].room.occupants.length > 0) {
      result.push(partialRooms[i]);
    }
  }
  return result;
}

function groupByGroupNumber_(occupants) {
  var groups = {};
  for (var i = 0; i < occupants.length; i++) {
    var gn = occupants[i].groupNumber;
    if (!groups[gn]) groups[gn] = [];
    groups[gn].push(occupants[i]);
  }
  return groups;
}

// ============================================================
// دوال مساعدة للغرف
// ============================================================

function getPoolByCapacity_(rooms, capacity) {
  if (capacity === 4) return rooms.quad;
  if (capacity === 3) return rooms.triple;
  if (capacity === 2) return rooms.double;
  return rooms.quad;
}

function findBestRoom_(rooms, groupSize) {
  var pools = [rooms.quad, rooms.triple, rooms.double];
  for (var p = 0; p < pools.length; p++) {
    for (var i = 0; i < pools[p].length; i++) {
      if (pools[p][i].occupants.length === 0 && pools[p][i].capacity >= groupSize) {
        return pools[p].splice(i, 1)[0];
      }
    }
  }
  return null;
}

function findAnyRoom_(rooms) {
  if (rooms.quad.length > 0) return rooms.quad.shift();
  if (rooms.triple.length > 0) return rooms.triple.shift();
  if (rooms.double.length > 0) return rooms.double.shift();
  return null;
}

function upgradeRoom_(rooms, targetCapacity) {
  if (targetCapacity === 4 && rooms.triple.length > 0) {
    var room = rooms.triple.shift();
    room.capacity = 4;
    return room;
  }
  if (targetCapacity === 4 && rooms.double.length > 0) {
    var room = rooms.double.shift();
    room.capacity = 4;
    return room;
  }
  if (targetCapacity === 3 && rooms.double.length > 0) {
    var room = rooms.double.shift();
    room.capacity = 3;
    return room;
  }
  return null;
}

// ============================================================
// المرحلة 4: الكتابة في الشيت — BATCH (أداء عالي)
// ============================================================

function writeAssignments_(ss, hotelName, allAssigned, partialRooms) {
  var hotelSheet = getOrCreateHotelSheet_(ss, hotelName);
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;

  var lastRow = hotelSheet.getLastRow();
  if (lastRow <= 1) return 0;
  var numRows = lastRow - 1;

  // قراءة دفعة واحدة
  var bookingIds = hotelSheet.getRange(2, 1, numRows, 1).getValues();
  var roomGroupCol = hotelSheet.getRange(2, 17, numRows, 1).getValues();
  var roomNumberCol = hotelSheet.getRange(2, 18, numRows, 1).getValues();

  // شركاء الغرفة — عمود 24 (أو أنشئه)
  var hasCol24 = hotelSheet.getLastColumn() >= 24;
  var roommateCol;
  if (hasCol24) {
    roommateCol = hotelSheet.getRange(2, 24, numRows, 1).getValues();
  } else {
    roommateCol = [];
    for (var i = 0; i < numRows; i++) roommateCol.push(['']);
  }

  // بناء خريطة bookingId → rowIndex
  var bidToRow = {};
  for (var i = 0; i < bookingIds.length; i++) {
    bidToRow[String(bookingIds[i][0])] = i;
  }

  // بناء خريطة Room Mapping للتحديث
  var mapUpdates = {}; // rowIndex → { status, occupantIds }

  var allRooms = allAssigned.concat(partialRooms);
  var totalAssigned = 0;

  for (var r = 0; r < allRooms.length; r++) {
    var room = allRooms[r].room;
    var members = room.occupants;
    var roommateNames = members.map(function(m) { return m.name; }).join(' | ');
    var occupantIds = members.map(function(m) { return m.bookingId; }).join('|');

    for (var m = 0; m < members.length; m++) {
      var rowIdx = bidToRow[members[m].bookingId];
      if (rowIdx !== undefined) {
        roomGroupCol[rowIdx][0] = room.internalId;
        roomNumberCol[rowIdx][0] = room.actualRoomNo || '';
        roommateCol[rowIdx][0] = roommateNames;
        totalAssigned++;
      }
    }

    // تحديث Room Mapping
    if (room.rowIndex) {
      mapUpdates[room.rowIndex] = { status: 'Occupied', occupantIds: occupantIds };
    }
  }

  // كتابة دفعة واحدة
  if (totalAssigned > 0) {
    hotelSheet.getRange(2, 17, numRows, 1).setValues(roomGroupCol);
    hotelSheet.getRange(2, 18, numRows, 1).setValues(roomNumberCol);

    if (hasCol24) {
      hotelSheet.getRange(2, 24, numRows, 1).setValues(roommateCol);
    }
  }

  // تحديث Room Mapping دفعة واحدة
  if (mapSheet && Object.keys(mapUpdates).length > 0) {
    var mapData = mapSheet.getDataRange().getValues();
    var mapStatusCol = [];
    var mapOccCol = [];
    for (var i = 1; i < mapData.length; i++) {
      if (mapUpdates[i]) {
        mapStatusCol.push([mapUpdates[i].status]);
        mapOccCol.push([mapUpdates[i].occupantIds]);
      } else {
        mapStatusCol.push([mapData[i][MC.STATUS]]);
        mapOccCol.push([mapData[i][MC.OCCUPANT_IDS]]);
      }
    }
    if (mapStatusCol.length > 0) {
      mapSheet.getRange(2, MC.STATUS + 1, mapStatusCol.length, 1).setValues(mapStatusCol);
      mapSheet.getRange(2, MC.OCCUPANT_IDS + 1, mapOccCol.length, 1).setValues(mapOccCol);
    }
  }

  SpreadsheetApp.flush();
  return totalAssigned;
}

// ============================================================
// مسح التسكين الحالي لفندق — يحمي المسجّلين (arrived)
// ============================================================

function clearAssignment_(ss, hotelName) {
  var hotelSheet = getHotelSheet_(ss, hotelName);
  if (!hotelSheet) return;

  var lastRow = hotelSheet.getLastRow();
  if (lastRow <= 1) return;
  var numRows = lastRow - 1;

  // قراءة دفعة واحدة
  var rgCol = hotelSheet.getRange(2, 17, numRows, 1).getValues();
  var rnCol = hotelSheet.getRange(2, 18, numRows, 1).getValues();
  var checkinCol = hotelSheet.getRange(2, 19, numRows, 1).getValues(); // حالة Check-in

  var hasCol24 = hotelSheet.getLastColumn() >= 24;
  var rmCol = hasCol24 ? hotelSheet.getRange(2, 24, numRows, 1).getValues() : null;

  var cleared = 0;
  for (var i = 0; i < numRows; i++) {
    // حماية: لا تمسح حاج سجّل وصوله
    if (String(checkinCol[i][0]) === 'arrived') continue;

    if (rgCol[i][0]) {
      rgCol[i][0] = '';
      rnCol[i][0] = '';
      if (rmCol) rmCol[i][0] = '';
      cleared++;
    }
  }

  if (cleared > 0) {
    hotelSheet.getRange(2, 17, numRows, 1).setValues(rgCol);
    hotelSheet.getRange(2, 18, numRows, 1).setValues(rnCol);
    if (rmCol) hotelSheet.getRange(2, 24, numRows, 1).setValues(rmCol);
  }

  // مسح Room Mapping لهذا الفندق (batch)
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  if (mapSheet) {
    var mapData = mapSheet.getDataRange().getValues();
    var normHotel = normalizeHotelName_(hotelName);
    var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;

    var statusCol = [];
    var occCol = [];
    var changed = false;

    for (var i = 1; i < mapData.length; i++) {
      if (normalizeHotelName_(String(mapData[i][MC.HOTEL_NAME])) === normHotel &&
          String(mapData[i][MC.STATUS]) === 'Occupied') {
        statusCol.push(['Available']);
        occCol.push(['']);
        changed = true;
      } else {
        statusCol.push([mapData[i][MC.STATUS]]);
        occCol.push([mapData[i][MC.OCCUPANT_IDS]]);
      }
    }

    if (changed && statusCol.length > 0) {
      mapSheet.getRange(2, MC.STATUS + 1, statusCol.length, 1).setValues(statusCol);
      mapSheet.getRange(2, MC.OCCUPANT_IDS + 1, occCol.length, 1).setValues(occCol);
    }
  }

  SpreadsheetApp.flush();
}

// ============================================================
// نقطة الدخول: توزيع ذكي — فندق واحد
// ============================================================

function smartAssignHotel(hotelName, hotelCity) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) return { success: false, error: 'النظام مشغول' };

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);

    clearAssignment_(ss, hotelName);

    var prep = prepareAssignment_(ss, hotelName, hotelCity);
    var familyResult = assignFamilies_(prep.families, prep.roomMapping);
    var sharerResult = assignSharers_(prep.sharers, prep.roomMapping);
    var optimized = optimizeRooms_(sharerResult.partialRooms);

    var allAssigned = familyResult.assigned.concat(sharerResult.assigned);
    var totalWritten = writeAssignments_(ss, hotelName, allAssigned, optimized);

    var report = {
      success: true,
      hotel: hotelName,
      city: hotelCity,
      totalPilgrims: prep.pilgrimCount,
      guideSharersExcluded: prep.guideSharersExcluded,
      familiesAssigned: familyResult.assigned.length,
      sharedRoomsAssigned: sharerResult.assigned.length,
      partialRooms: optimized.length,
      totalWritten: totalWritten,
      noRoomFamilies: familyResult.noRoom.length,
      partialDetails: optimized.map(function(p) {
        return {
          roomId: p.room.internalId,
          occupants: p.room.occupants.length,
          capacity: p.room.capacity,
          gender: p.room.occupants.length > 0 ? p.room.occupants[0].gender : ''
        };
      })
    };

    Logger.log('SmartAssign: ' + hotelName + ' — ' + totalWritten + ' حاج في ' +
      (allAssigned.length + optimized.length) + ' غرفة');
    return report;
  } catch (e) {
    Logger.log('smartAssignHotel ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// نقطة الدخول: توزيع ذكي — كل الفنادق
// ============================================================

function smartAssignAll() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var J = HOTEL_CONFIG.JOURNEY_COLS;

  // جمع الفنادق مع المدينة الصحيحة
  var hotelCityMap = {}; // hotelName → hotelCity

  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];

    // المدينة
    var medHotel = String(row[J.MADINAH_EN] || '').trim();
    if (medHotel && medHotel !== 'NULL') hotelCityMap[medHotel] = 'Madina';

    // مكة
    var makHotel = String(row[J.MAKKAH_EN] || '').trim();
    if (makHotel && makHotel !== 'NULL') hotelCityMap[makHotel] = 'Makkah';

    // مكة التحويلية
    var shiftHotel = String(row[J.MAKKAH_SHIFT_EN] || '').trim();
    if (shiftHotel && shiftHotel !== 'NULL') hotelCityMap[shiftHotel] = 'Makkah Shifting';
  }

  var results = [];
  for (var name in hotelCityMap) {
    results.push(smartAssignHotel(name, hotelCityMap[name]));
  }

  return {
    success: true,
    totalHotels: results.length,
    results: results
  };
}

// ============================================================
// زر "اكتمل العدد" — تحديث التغييرات فقط
// ============================================================

function updateAssignment(hotelName, hotelCity) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) return { success: false, error: 'النظام مشغول' };

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var hotelSheet = getHotelSheet_(ss, hotelName);
    if (!hotelSheet) return { success: false, error: 'شيت الفندق غير موجود' };

    var lastRow = hotelSheet.getLastRow();
    if (lastRow <= 1) return { success: true, newPilgrims: 0, cancelled: 0, added: 0 };

    var numRows = lastRow - 1;
    var bookingIds = hotelSheet.getRange(2, 1, numRows, 1).getValues();
    var roomGroupCol = hotelSheet.getRange(2, 17, numRows, 1).getValues();

    // التسكين الحالي
    var currentAssigned = {};
    for (var i = 0; i < numRows; i++) {
      var bid = String(bookingIds[i][0]);
      var rgid = String(roomGroupCol[i][0] || '');
      if (rgid) currentAssigned[bid] = rgid;
    }

    // الحجاج من رحلة الحاج
    var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
    var journeyData = journeySheet.getDataRange().getValues();
    var J = HOTEL_CONFIG.JOURNEY_COLS;
    var normHotel = normalizeHotelName_(hotelName);

    var currentPilgrims = {};
    for (var i = 1; i < journeyData.length; i++) {
      var row = journeyData[i];
      var pilgrimHotel = getHotelName_(row, hotelCity);
      if (!pilgrimHotel) continue;
      if (normalizeHotelName_(pilgrimHotel) !== normHotel) continue;
      currentPilgrims[String(row[J.BOOKING_ID])] = true;
    }

    // كشف التغييرات
    var newPilgrims = [];
    var cancelled = [];

    for (var bid in currentPilgrims) {
      if (!currentAssigned[bid]) newPilgrims.push(bid);
    }
    for (var bid in currentAssigned) {
      if (!currentPilgrims[bid]) cancelled.push(bid);
    }

    // إخراج الملغيين (batch)
    if (cancelled.length > 0) {
      var cancelSet = {};
      for (var i = 0; i < cancelled.length; i++) cancelSet[cancelled[i]] = true;

      for (var i = 0; i < numRows; i++) {
        if (cancelSet[String(bookingIds[i][0])]) {
          roomGroupCol[i][0] = '';
        }
      }
      hotelSheet.getRange(2, 17, numRows, 1).setValues(roomGroupCol);
      SpreadsheetApp.flush();
    }

    // تسكين الجدد
    var addedCount = 0;
    if (newPilgrims.length > 0) {
      var prep = prepareAssignment_(ss, hotelName, hotelCity);

      var newFamilies = prep.families.filter(function(f) {
        return f.members.some(function(m) { return newPilgrims.indexOf(m.bookingId) !== -1; });
      });
      var newSharers = prep.sharers.filter(function(s) {
        return s.members.some(function(m) { return newPilgrims.indexOf(m.bookingId) !== -1; });
      });

      if (newFamilies.length > 0) {
        var famResult = assignFamilies_(newFamilies, prep.roomMapping);
        addedCount += writeAssignments_(ss, hotelName, famResult.assigned, []);
      }
      if (newSharers.length > 0) {
        var shrResult = assignSharers_(newSharers, prep.roomMapping);
        var optimized = optimizeRooms_(shrResult.partialRooms);
        addedCount += writeAssignments_(ss, hotelName, shrResult.assigned, optimized);
      }
    }

    return {
      success: true,
      hotel: hotelName,
      newPilgrims: newPilgrims.length,
      cancelled: cancelled.length,
      added: addedCount
    };
  } catch (e) {
    Logger.log('updateAssignment ERROR: ' + e.toString());
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// إحصائية المرشدين
// ============================================================

function getGuideStats(hotelName, hotelCity) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var guideMap = buildGuideRabihMap_(ss);

  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var normHotel = normalizeHotelName_(hotelName);

  var guideStats = {};

  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var pilgrimHotel = getHotelName_(row, hotelCity);
    if (!pilgrimHotel) continue;
    if (normalizeHotelName_(pilgrimHotel) !== normHotel) continue;

    var passport = String(row[J.PASSPORT]).trim().toUpperCase();
    var info = guideMap[passport];
    if (!info || !info.guide) continue;

    if (!guideStats[info.guide]) {
      guideStats[info.guide] = { total: 0 };
    }
    guideStats[info.guide].total++;
  }

  // الغرف الفاضية
  var roomMapping = getAvailableRooms_(ss, hotelName, hotelCity);

  return {
    success: true,
    hotel: hotelName,
    guides: guideStats,
    availableRooms: {
      quad: roomMapping.quad.length,
      triple: roomMapping.triple.length,
      double: roomMapping.double.length
    }
  };
}
