// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║              Room Type Builder — بناء الجدول التشغيلي للغرف                ║
// ╠══════════════════════════════════════════════════════════════════════════════╣
// ║  يقرأ من: NUSK Room Type + Personal Details + الباقات                      ║
// ║  يكتب في: Room Type                                                       ║
// ║  يُعاد بناؤه بالكامل عند كل تحديث لـ NUSK Room Type                      ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

var RT_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  
  // أسماء الشيتات
  NUSK_SHEET: 'NUSK Room Type',
  PD_SHEET: 'Presonal Details',
  PKG_SHEET: 'الباقات',
  B2C_SHEET: 'B2C',
OUTPUT_SHEET: 'Room Type Preview',
  
  // صفوف البداية
  NUSK_START: 2,    // الصف الأول بيانات
  PD_START: 2,
  PKG_START: 3,     // الباقات الهيدر بصف 2
  B2C_START: 2,
  
  // أعمدة NUSK Room Type
  NUSK: {
    GROUP: 1,        // A - GroupNumber
    BOOKING: 2,      // B - BookingId
    PKG_ID: 3,       // C - PackageId
    PKG_NAME: 4,     // D - Package
    APPLICANTS: 6,   // F - NumberOfApplicants
    ROOM_TYPE: 7,    // G - RoomType (2/3/4)
    NUM_ROOMS: 8,    // H - NumberOfRooms (= عدد الأسرّة)
    HOTEL_TYPE: 9    // I - HotelType
  },
  
  // أعمدة Personal Details
  PD: {
    GROUP: 2,        // B - Group Number
    GENDER: 5,       // E - Gender
    PASSPORT: 6,     // F - Passport Number
    FIRST_AR: 9,     // I - First Name Arabic
    LAST_AR: 10,     // J - Last Name Arabic
    FIRST_EN: 11,    // K - First Name English
    LAST_EN: 12,     // L - Last Name English
    PHONE: 15,       // O - Mobile Number
    GUIDE: 16,       // P - Tour Guide Name
    NATIONALITY: 18, // R - Nationality
    PKG_NUM: 19,     // S - Package Number
    FLIGHT_TYPE: 21, // U - Flight Contract Type
    ARRIVAL: 30,     // AD - Arrival Time Transportation
    DEPARTURE: 32    // AF - Departure Time Transportation
  },
  
  // أعمدة الباقات
  PKG: {
    NUSK_NO: 2,      // B - Nusk No.
    CITY_START: 10,  // J - City Of Start
    H1_CITY: 12,     // L - City (Hotel 1)
    H1_NAME_EN: 14,  // N - Name of Hotel English
    H2_CITY: 27,     // AA - City (Hotel 2)
    H2_NAME_EN: 29,  // AC - Name of Hotel English
    H3_CITY: 42,     // AP - City (Hotel 3)
    H3_NAME_EN: 44   // AR - Name of Hotel English
  },
  
  // أعمدة B2C (للرحلات)
  B2C: {
    GROUP: 2,        // B - رقم المجموعة
    PASSPORT: 6,     // F - رقم الجواز
    // ذهاب - الرحلة الأخيرة (وصول السعودية)
    ARR_TO: 44,      // AR - To 2 (JED/MED)
    ARR_DATE: 45,    // AS - DATE LANDING 2
    // عودة - الرحلة الأولى (مغادرة السعودية)
    DEP_FROM: 50,    // AX - From1 (Return)
    DEP_DATE: 48     // AV - Date TAKEOFF 1 (Return)
  },
  
  // خريطة سعة الغرف
  ROOM_CAPACITY: { '2': 2, '3': 3, '4': 4 },
  ROOM_NAMES: { '2': 'Double', '3': 'Triple', '4': 'Quad' }
};


// ==================== القائمة ====================

/**
 * إضافة قائمة Room Type في شريط الأدوات
 * يُستدعى من onOpen() في CompleteScript
 */
function addRoomTypeMenu() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('🏨 Room Type', [
    { name: '🔄 إعادة بناء Room Type', functionName: 'buildRoomType' },
    { name: '📊 إحصائيات سريعة', functionName: 'showRoomTypeStats' },
    { name: '─────────────', functionName: 'buildRoomType' },
    { name: 'ℹ️ آخر تحديث', functionName: 'showLastUpdate' }
  ]);
}


// ==================== الدالة الرئيسية ====================

/**
 * بناء جدول Room Type بالكامل من الصفر
 * يُستدعى يدوياً من القائمة أو تلقائياً عند تحديث NUSK
 */
function buildRoomType() {
  var startTime = new Date();
  Logger.log('╔══════════════════════════════════════╗');
  Logger.log('║     بدء بناء Room Type               ║');
  Logger.log('╚══════════════════════════════════════╝');
  
  var ss = SpreadsheetApp.openById(RT_CONFIG.SPREADSHEET_ID);
  
  // ═══ 1. قراءة البيانات المصدرية ═══
  Logger.log('📖 قراءة NUSK Room Type...');
  var nuskData = _readNUSK(ss);
  Logger.log('   → ' + Object.keys(nuskData).length + ' مجموعة');
  
  Logger.log('📖 قراءة Personal Details...');
  var pdData = _readPD(ss);
  Logger.log('   → ' + Object.keys(pdData).length + ' مجموعة');
  
  Logger.log('📖 قراءة الباقات...');
  var pkgData = _readPackages(ss);
  Logger.log('   → ' + Object.keys(pkgData).length + ' باقة');
  
  Logger.log('📖 قراءة B2C (مطارات)...');
  var b2cAirports = _readB2CAirports(ss);
  Logger.log('   → ' + Object.keys(b2cAirports).length + ' مجموعة بمطار');
  
  // ═══ 2. حفظ الملاحظات القديمة ═══
  var oldNotes = _saveOldNotes(ss);
  Logger.log('💾 ملاحظات محفوظة: ' + Object.keys(oldNotes).length);
  
  // ═══ 3. بناء الصفوف ═══
  Logger.log('🔧 بناء الصفوف...');
  var rows = _buildRows(nuskData, pdData, pkgData, b2cAirports, oldNotes);
  Logger.log('   → ' + rows.length + ' صف');
  
  // ═══ 4. كتابة النتائج ═══
  Logger.log('📝 كتابة Room Type...');
  _writeRoomType(ss, rows);
  
  // ═══ 5. تسجيل وقت التحديث ═══
  PropertiesService.getScriptProperties().setProperty('RT_LAST_UPDATE', new Date().toISOString());
  
  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('✅ اكتمل في ' + elapsed + ' ثانية');
  
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'تم بناء Room Type: ' + rows.length + ' مجموعة في ' + elapsed + ' ثانية',
      '✅ Room Type', 5
    );
  } catch(e) {}
}


// ==================== قراءة البيانات ==================== 

/**
 * قراءة NUSK Room Type → map بـ GroupNumber
 * كل مجموعة: { booking, pkgId, pkgName, applicants, hotels: { Madinah/Makkah/Makkah-Shifting: [{rt, beds}] } }
 */
function _readNUSK(ss) {
  var sheet = ss.getSheetByName(RT_CONFIG.NUSK_SHEET);
  if (!sheet || sheet.getLastRow() < RT_CONFIG.NUSK_START) return {};
  
  var data = sheet.getRange(RT_CONFIG.NUSK_START, 1, sheet.getLastRow() - RT_CONFIG.NUSK_START + 1, 9).getValues();
  var result = {};
  var C = RT_CONFIG.NUSK;
  
  for (var i = 0; i < data.length; i++) {
    var gn = String(data[i][C.GROUP - 1] || '').replace('.0', '').trim();
    if (!gn) continue;
    
    var ht = String(data[i][C.HOTEL_TYPE - 1] || '').trim();
    var rt = String(data[i][C.ROOM_TYPE - 1] || '').replace('.0', '').trim();
    var beds = parseInt(data[i][C.NUM_ROOMS - 1]) || 0;
    
    if (!result[gn]) {
      result[gn] = {
        booking: String(data[i][C.BOOKING - 1] || '').replace('.0', '').trim(),
        pkgId: String(data[i][C.PKG_ID - 1] || '').replace('.0', '').trim(),
        pkgName: String(data[i][C.PKG_NAME - 1] || '').trim(),
        applicants: parseInt(data[i][C.APPLICANTS - 1]) || 0,
        hotels: {}
      };
    }
    
    if (!result[gn].hotels[ht]) {
      result[gn].hotels[ht] = [];
    }
    result[gn].hotels[ht].push({ rt: rt, beds: beds });
  }
  
  return result;
}

/**
 * قراءة Personal Details → map بـ GroupNumber
 * كل مجموعة: { male, female, nationality, leader, phone, guide, arrival, departure }
 */
function _readPD(ss) {
  var sheet = ss.getSheetByName(RT_CONFIG.PD_SHEET);
  if (!sheet || sheet.getLastRow() < RT_CONFIG.PD_START) return {};
  
  var data = sheet.getRange(RT_CONFIG.PD_START, 1, sheet.getLastRow() - RT_CONFIG.PD_START + 1, 32).getValues();
  var result = {};
  var C = RT_CONFIG.PD;
  
  for (var i = 0; i < data.length; i++) {
    var gn = String(data[i][C.GROUP - 1] || '').replace('.0', '').trim();
    if (!gn) continue;
    
    var gender = String(data[i][C.GENDER - 1] || '').trim().toLowerCase();
    var nat = String(data[i][C.NATIONALITY - 1] || '').trim();
    var firstName = String(data[i][C.FIRST_EN - 1] || '').trim();
    var lastName = String(data[i][C.LAST_EN - 1] || '').trim();
    var phone = String(data[i][C.PHONE - 1] || '').trim();
    var guide = String(data[i][C.GUIDE - 1] || '').trim();
    var arrival = data[i][C.ARRIVAL - 1];
    var departure = data[i][C.DEPARTURE - 1];
    
    if (!result[gn]) {
      result[gn] = {
        male: 0,
        female: 0,
        nationalities: {},
        leader: firstName + ' ' + lastName,
        phone: phone,
        guide: guide,
        arrival: arrival,
        departure: departure
      };
    }
    
    if (gender === 'male') result[gn].male++;
    else if (gender === 'female') result[gn].female++;
    
    if (nat) {
      result[gn].nationalities[nat] = (result[gn].nationalities[nat] || 0) + 1;
    }
  }
  
  // تحويل الجنسيات إلى نص
  for (var gn in result) {
    var nats = result[gn].nationalities;
    var natArr = Object.keys(nats).sort(function(a, b) { return nats[b] - nats[a]; });
    result[gn].nationality = natArr.join(', ');
    delete result[gn].nationalities;
  }
  
  return result;
}

/**
 * قراءة الباقات → map بـ Nusk No
 * كل باقة: { cityStart, hotels: { Madinah/Makkah/Makkah-Shifting: hotelName } }
 */
function _readPackages(ss) {
  var sheet = ss.getSheetByName(RT_CONFIG.PKG_SHEET);
  if (!sheet || sheet.getLastRow() < RT_CONFIG.PKG_START) return {};
  
  var data = sheet.getRange(RT_CONFIG.PKG_START, 1, sheet.getLastRow() - RT_CONFIG.PKG_START + 1, 54).getValues();
  var result = {};
  var C = RT_CONFIG.PKG;
  
  for (var i = 0; i < data.length; i++) {
    var nusk = String(data[i][C.NUSK_NO - 1] || '').replace('.0', '').trim();
    if (!nusk) continue;
    
    var cityStart = String(data[i][C.CITY_START - 1] || '').trim();
    var hotels = {};
    
    // فندق 1
    var h1City = String(data[i][C.H1_CITY - 1] || '').trim();
    var h1Name = String(data[i][C.H1_NAME_EN - 1] || '').trim();
    if (h1City && h1Name) {
      if (h1City.indexOf('Med') >= 0) hotels['Madinah'] = h1Name;
      else if (h1City.indexOf('Mak') >= 0) hotels['Makkah'] = h1Name;
    }
    
    // فندق 2
    var h2City = String(data[i][C.H2_CITY - 1] || '').trim();
    var h2Name = String(data[i][C.H2_NAME_EN - 1] || '').trim();
    if (h2City && h2Name) {
      if (h2City.indexOf('Med') >= 0) hotels['Madinah'] = hotels['Madinah'] || h2Name;
      else if (h2City.indexOf('Mak') >= 0) {
        if (hotels['Makkah']) hotels['Makkah-Shifting'] = h2Name;
        else hotels['Makkah'] = h2Name;
      }
    }
    
    // فندق 3
    var h3City = String(data[i][C.H3_CITY - 1] || '').trim();
    var h3Name = String(data[i][C.H3_NAME_EN - 1] || '').trim();
    if (h3City && h3Name) {
      if (h3City.indexOf('Med') >= 0) hotels['Madinah'] = hotels['Madinah'] || h3Name;
      else if (h3City.indexOf('Mak') >= 0) hotels['Makkah-Shifting'] = hotels['Makkah-Shifting'] || h3Name;
    }
    
    result[nusk] = {
      cityStart: cityStart,
      hotels: hotels
    };
  }
  
  return result;
}

/**
 * قراءة B2C → map بـ GroupNumber → { arrAirport, depAirport }
 */
function _readB2CAirports(ss) {
  var sheet = ss.getSheetByName(RT_CONFIG.B2C_SHEET);
  if (!sheet || sheet.getLastRow() < RT_CONFIG.B2C_START) return {};
  
  var data = sheet.getRange(RT_CONFIG.B2C_START, 1, sheet.getLastRow() - RT_CONFIG.B2C_START + 1, 53).getValues();
  var result = {};
  var C = RT_CONFIG.B2C;
  
  for (var i = 0; i < data.length; i++) {
    var gn = String(data[i][C.GROUP - 1] || '').replace('.0', '').trim();
    if (!gn || result[gn]) continue; // أول حاج بالمجموعة يكفي
    
    var arrTo = String(data[i][C.ARR_TO - 1] || '').trim().toUpperCase();
    var depFrom = String(data[i][C.DEP_FROM - 1] || '').trim().toUpperCase();
    
    if (arrTo || depFrom) {
      result[gn] = {
        arrAirport: arrTo,
        depAirport: depFrom
      };
    }
  }
  
  return result;
}


// ==================== حفظ الملاحظات ==================== 

/**
 * حفظ الملاحظات القديمة قبل إعادة البناء
 * يُرجع map: GroupNumber → ملاحظة
 */
function _saveOldNotes(ss) {
  var sheet = ss.getSheetByName(RT_CONFIG.OUTPUT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 2) return {};
  
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var result = {};
  
  // العمود الأخير = الملاحظات (AF = 32)
  var noteCol = 31; // index 0-based for col 32
  
  for (var i = 0; i < data.length; i++) {
    var gn = String(data[i][0] || '').replace('.0', '').trim();
    var note = String(data[i][noteCol] || '').trim();
    if (gn && note) {
      result[gn] = note;
    }
  }
  
  return result;
}


// ==================== بناء الصفوف ==================== 

/**
 * بناء صفوف Room Type من البيانات المدمجة
 * يُرجع مصفوفة ثنائية جاهزة للكتابة
 */
function _buildRows(nuskData, pdData, pkgData, b2cAirports, oldNotes) {
  var rows = [];
  var RT_CAP = RT_CONFIG.ROOM_CAPACITY;
  var RT_NAME = RT_CONFIG.ROOM_NAMES;
  
  // ترتيب حسب ظهور المجموعات في NUSK
  var groupKeys = Object.keys(nuskData);
  
  for (var g = 0; g < groupKeys.length; g++) {
    var gn = groupKeys[g];
    var nusk = nuskData[gn];
    var pd = pdData[gn] || {};
    var pkg = pkgData[nusk.pkgId] || {};
    var pkgHotels = pkg.hotels || {};
    var airports = b2cAirports[gn] || {};
    
    // ─── تاريخ الوصول والمغادرة ───
    var arrival = pd.arrival || '';
    var departure = pd.departure || '';
    
    // ─── مطار الوصول: من B2C أو من مدينة بداية الباقة ───
    var arrAirport = airports.arrAirport || '';
    var depAirport = airports.depAirport || '';
    if (!arrAirport && pkg.cityStart) {
      arrAirport = pkg.cityStart.indexOf('Mak') >= 0 ? 'JED' : 'MED';
    }
    
    // ─── بيانات الفنادق والغرف ───
    var hotelData = _buildHotelCols(nusk.hotels, pkgHotels, RT_CAP, RT_NAME);
    
    // ─── الحالة ───
    var status = 'مؤكد';
    
    // ─── الملاحظات القديمة ───
    var note = oldNotes[gn] || '';
    
    // ─── بناء الصف ───
    var row = [
      parseInt(gn),                           // A: GroupNumber
      parseInt(nusk.booking),                  // B: BookingId
      parseInt(nusk.pkgId),                    // C: PackageId
      nusk.pkgName,                            // D: Package Name
      nusk.applicants,                         // E: عدد الأفراد
      pd.male || 0,                            // F: رجال
      pd.female || 0,                          // G: نساء
      pd.nationality || '',                    // H: الجنسية
      pd.leader || '',                         // I: رئيس المجموعة
      pd.phone || '',                          // J: رقم التواصل
      pd.guide || '',                          // K: Tour Guide
      arrival,                                 // L: تاريخ الوصول
      arrAirport,                              // M: مطار الوصول
      departure,                               // N: تاريخ المغادرة
      depAirport,                              // O: مطار المغادرة
      
      // المدينة
      hotelData.med.hotel,                     // P: Madinah Hotel
      hotelData.med.roomType,                  // Q: Room Type Madinah
      hotelData.med.beds,                      // R: No. Beds Madinah
      hotelData.med.fullRooms,                 // S: غرف كاملة Madinah
      hotelData.med.sharedBeds,                // T: أسرّة مشتركة Madinah
      
      // مكة 1
      hotelData.mak1.hotel,                    // U: Makkah 1 Hotel
      hotelData.mak1.roomType,                 // V: Room Type Makkah 1
      hotelData.mak1.beds,                     // W: No. Beds Makkah 1
      hotelData.mak1.fullRooms,                // X: غرف كاملة Makkah 1
      hotelData.mak1.sharedBeds,               // Y: أسرّة مشتركة Makkah 1
      
      // مكة 2 (انتقالية)
      hotelData.mak2.hotel,                    // Z: Makkah 2 Hotel
      hotelData.mak2.roomType,                 // AA: Room Type Makkah 2
      hotelData.mak2.beds,                     // AB: No. Beds Makkah 2
      hotelData.mak2.fullRooms,                // AC: غرف كاملة Makkah 2
      hotelData.mak2.sharedBeds,               // AD: أسرّة مشتركة Makkah 2
      
      status,                                  // AE: Status
      note                                     // AF: ملاحظات
    ];
    
    rows.push(row);
  }
  
  return rows;
}

/**
 * بناء أعمدة الفنادق لمجموعة واحدة
 * يتعامل مع الحالات الشاذة (أنواع غرف مختلطة في نفس الفندق)
 */
function _buildHotelCols(nuskHotels, pkgHotels, RT_CAP, RT_NAME) {
  var empty = { hotel: '', roomType: '', beds: '', fullRooms: '', sharedBeds: '' };
  
  var result = {
    med: _calcHotel(nuskHotels['Madinah'], pkgHotels['Madinah'], RT_CAP, RT_NAME),
    mak1: _calcHotel(nuskHotels['Makkah'], pkgHotels['Makkah'], RT_CAP, RT_NAME),
    mak2: _calcHotel(nuskHotels['Makkah-Shifting'], pkgHotels['Makkah-Shifting'], RT_CAP, RT_NAME)
  };
  
  return result;
}

/**
 * حساب بيانات فندق واحد
 * entries = [{rt, beds}, ...] — قد يكون أكثر من نوع غرفة (حالة شاذة)
 */
function _calcHotel(entries, hotelName, RT_CAP, RT_NAME) {
  if (!entries || entries.length === 0) {
    return { hotel: '', roomType: '', beds: '', fullRooms: '', sharedBeds: '' };
  }
  
  hotelName = hotelName || '';
  
  if (entries.length === 1) {
    // حالة عادية: نوع غرفة واحد
    var e = entries[0];
    var cap = RT_CAP[e.rt] || 4;
    var fullRooms = Math.floor(e.beds / cap);
    var shared = e.beds % cap;
    
    return {
      hotel: hotelName,
      roomType: RT_NAME[e.rt] || e.rt,
      beds: e.beds,
      fullRooms: fullRooms,
      sharedBeds: shared
    };
  }
  
  // حالة شاذة: أنواع مختلطة
  var rtParts = [];
  var bedsParts = [];
  var totalFull = 0;
  var totalShared = 0;
  
  for (var i = 0; i < entries.length; i++) {
    var e = entries[i];
    var cap = RT_CAP[e.rt] || 4;
    rtParts.push(RT_NAME[e.rt] || e.rt);
    bedsParts.push(e.beds);
    totalFull += Math.floor(e.beds / cap);
    totalShared += e.beds % cap;
  }
  
  return {
    hotel: hotelName,
    roomType: rtParts.join(' + '),
    beds: bedsParts.join(' + '),
    fullRooms: totalFull,
    sharedBeds: totalShared
  };
}


// ==================== كتابة النتائج ==================== 

/**
 * كتابة جدول Room Type بالكامل مع التنسيق
 */
function _writeRoomType(ss, rows) {
  var sheet = ss.getSheetByName(RT_CONFIG.OUTPUT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RT_CONFIG.OUTPUT_SHEET);
  }
  
  // مسح كامل
  sheet.clear();
  
  // ═══ الهيدر ═══
  var headers = [
    'GroupNumber', 'BookingId', 'PackageId', 'Package Name',
    'Pilgrims', 'Male', 'Female',
    'Nationality', 'Group Leader', 'Phone', 'Tour Guide',
    'Arrival Date', 'Arrival Airport', 'Departure Date', 'Departure Airport',
    'Madinah Hotel', 'Room Type Med', 'Beds Med', 'Full Rooms Med', 'Shared Beds Med',
    'Makkah 1 Hotel', 'Room Type Mak1', 'Beds Mak1', 'Full Rooms Mak1', 'Shared Beds Mak1',
    'Makkah 2 Hotel', 'Room Type Mak2', 'Beds Mak2', 'Full Rooms Mak2', 'Shared Beds Mak2',
    'Status', 'Notes'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ═══ البيانات ═══
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // ═══ التنسيق ═══
  _formatRoomType(sheet, headers.length, rows.length);
}

/**
 * تنسيق الشيت: ألوان الهيدر + عرض الأعمدة + تجميد
 */
function _formatRoomType(sheet, numCols, numRows) {
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  
  // ألوان الأقسام
  var sections = [
    { start: 1, end: 4, color: '#1F4E79', label: 'بيانات المجموعة' },
    { start: 5, end: 7, color: '#4A148C', label: 'الأفراد' },
    { start: 8, end: 11, color: '#004D40', label: 'معلومات التواصل' },
    { start: 12, end: 15, color: '#BF360C', label: 'الرحلات' },
    { start: 16, end: 20, color: '#2E7D32', label: 'المدينة' },
    { start: 21, end: 25, color: '#1565C0', label: 'مكة 1' },
    { start: 26, end: 30, color: '#C62828', label: 'مكة 2' },
    { start: 31, end: 32, color: '#37474F', label: 'الحالة' }
  ];
  
  for (var s = 0; s < sections.length; s++) {
    var sec = sections[s];
    var range = sheet.getRange(1, sec.start, 1, sec.end - sec.start + 1);
    range.setBackground(sec.color)
         .setFontColor('#FFFFFF')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  }
  
  // تجميد الصف الأول + أول 4 أعمدة
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  
  // عرض الأعمدة
  var widths = {
    1: 100, 2: 100, 3: 80, 4: 280,     // A-D
    5: 65, 6: 50, 7: 55,                 // E-G
    8: 120, 9: 150, 10: 120, 11: 120,    // H-K
    12: 110, 13: 75, 14: 110, 15: 75,    // L-O
    16: 180, 17: 100, 18: 65, 19: 85, 20: 90,  // P-T
    21: 180, 22: 100, 23: 65, 24: 85, 25: 90,  // U-Y
    26: 180, 27: 100, 28: 65, 29: 85, 30: 90,  // Z-AD
    31: 75, 32: 150                       // AE-AF
  };
  
  for (var col in widths) {
    sheet.setColumnWidth(parseInt(col), widths[col]);
  }
  
  // تنسيق الأعمدة المحسوبة (غرف كاملة / أسرّة مشتركة)
  if (numRows > 0) {
    // الأسرّة المشتركة — تلوين الخلايا > 0 بالأصفر
    var sharedCols = [20, 25, 30]; // T, Y, AD
    for (var c = 0; c < sharedCols.length; c++) {
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#FFF9C4')
        .setFontColor('#E65100')
        .setRanges([sheet.getRange(2, sharedCols[c], numRows, 1)])
        .build();
      var rules = sheet.getConditionalFormatRules();
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);
    }
  }
}


// ==================== إحصائيات ==================== 

/**
 * عرض إحصائيات سريعة عن Room Type
 */
function showRoomTypeStats() {
  var ss = SpreadsheetApp.openById(RT_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(RT_CONFIG.OUTPUT_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('⚠️ جدول Room Type فارغ — شغّل إعادة البناء أولاً');
    return;
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 32).getValues();
  
  var totalGroups = data.length;
  var totalPilgrims = 0;
  var totalMale = 0;
  var totalFemale = 0;
  var totalSharedMed = 0;
  var totalSharedMak1 = 0;
  var totalSharedMak2 = 0;
  
  for (var i = 0; i < data.length; i++) {
    totalPilgrims += (parseInt(data[i][4]) || 0);
    totalMale += (parseInt(data[i][5]) || 0);
    totalFemale += (parseInt(data[i][6]) || 0);
    totalSharedMed += (parseInt(data[i][19]) || 0);   // T
    totalSharedMak1 += (parseInt(data[i][24]) || 0);   // Y
    totalSharedMak2 += (parseInt(data[i][29]) || 0);   // AD
  }
  
  var msg = '📊 إحصائيات Room Type\n' +
    '━━━━━━━━━━━━━━━━━━━━━\n' +
    '👥 المجموعات: ' + totalGroups + '\n' +
    '🧑 الحجاج: ' + totalPilgrims + ' (♂' + totalMale + ' ♀' + totalFemale + ')\n' +
    '━━━━━━━━━━━━━━━━━━━━━\n' +
    '🛏️ أسرّة مشتركة تحتاج تجميع:\n' +
    '  🕌 المدينة: ' + totalSharedMed + ' سرير\n' +
    '  🕋 مكة 1: ' + totalSharedMak1 + ' سرير\n' +
    '  🕋 مكة 2: ' + totalSharedMak2 + ' سرير\n' +
    '━━━━━━━━━━━━━━━━━━━━━\n' +
    '📅 آخر تحديث: ' + (PropertiesService.getScriptProperties().getProperty('RT_LAST_UPDATE') || 'لم يُحدّث');
  
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * عرض وقت آخر تحديث
 */
function showLastUpdate() {
  var last = PropertiesService.getScriptProperties().getProperty('RT_LAST_UPDATE');
  SpreadsheetApp.getUi().alert('⏰ آخر تحديث لـ Room Type:\n' + (last || 'لم يُحدّث بعد'));
}