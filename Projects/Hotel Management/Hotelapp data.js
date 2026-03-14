/**
 * Hotel Management App — Data Layer
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v3.1 — Data retrieval functions + Tour Guide direct lookup
 */

// ============================================================
// DATA: GET PILGRIMS FOR HOTEL
// ============================================================

function getPilgrimsForHotel(hotelName, hotelCity) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var roomSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_TYPE);
  var roomData = roomSheet.getDataRange().getValues();
  var checkinData = getCheckinData_(ss, hotelName);
  var transportMap = getTransportMap_(ss);
  var guideMap = buildGuideMap_(ss);
  
  // Build group lookup
  var groupInfo = {};
  var R = HOTEL_CONFIG.ROOM_COLS;
  for (var i = 1; i < roomData.length; i++) {
    var rRow = roomData[i];
    var gn = rRow[R.GROUP_NUMBER];
    if (gn) {
      groupInfo[gn] = {
        packageName: rRow[R.PACKAGE_NAME],
        groupLeader: rRow[R.GROUP_LEADER],
        phone: rRow[R.PHONE],
        tourGuide: rRow[R.TOUR_GUIDE],
        roomTypeMed: rRow[R.ROOM_TYPE_MED],
        bedsMed: Number(rRow[R.BEDS_MED]) || 0,
        fullRoomsMed: Number(rRow[R.FULL_ROOMS_MED]) || 0,
        sharedBedsMed: Number(rRow[R.SHARED_BEDS_MED]) || 0,
        roomTypeMak1: rRow[R.ROOM_TYPE_MAK1],
        bedsMak1: Number(rRow[R.BEDS_MAK1]) || 0,
        fullRoomsMak1: Number(rRow[R.FULL_ROOMS_MAK1]) || 0,
        sharedBedsMak1: Number(rRow[R.SHARED_BEDS_MAK1]) || 0,
        roomTypeMak2: rRow[R.ROOM_TYPE_MAK2],
        bedsMak2: Number(rRow[R.BEDS_MAK2]) || 0,
        fullRoomsMak2: Number(rRow[R.FULL_ROOMS_MAK2]) || 0,
        sharedBedsMak2: Number(rRow[R.SHARED_BEDS_MAK2]) || 0
      };
    }
  }
  
  var pilgrims = [];
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var actualCity = getActualHotelCity_(row, hotelName, hotelCity);
    
    if (actualCity) {
      var bookingId = row[J.BOOKING_ID];
      var groupNum = row[J.GROUP_NUMBER];
      var gi = groupInfo[groupNum] || {};
      var packageId = row[J.PACKAGE_ID];
      var transport = transportMap[packageId] || 'حافلة';
      
      var hotelDates = getHotelDates_(row, actualCity);
      var expectedArrival = calculateExpectedArrival_(row, actualCity);
      var arrivalStatus = checkArrivalStatus_(row, actualCity, expectedArrival);
      
      var roomType = getRoomType_(gi, actualCity);
      var sharedBeds = getSharedBeds_(gi, actualCity);
      var fullRooms = getFullRooms_(gi, actualCity);
      
      var checkin = checkinData[bookingId] || {};
      
      // Determine next hotel name for departures
      var nextHotelName = getNextHotelName_(row, actualCity);
      
      pilgrims.push({
        bookingId: bookingId,
        groupNumber: groupNum,
        name: row[J.NAME],
        passport: row[J.PASSPORT],
        gender: row[J.GENDER],
        nationality: row[J.NATIONALITY_EN],
        countryResidence: row[J.COUNTRY_RESIDENCE] || '',
        
        arrivalCity: row[J.ARRIVAL_CITY],
        arrivalDate: formatDate_(row[J.ARRIVAL_DATE]),
        arrivalTime: formatTime_(row[J.ARRIVAL_TIME]),
        arrivalFlight: row[J.ARRIVAL_FLIGHT],
        
        returnDate: formatDate_(row[J.RETURN_DEPT_DATE]),
        returnTime: formatTime_(row[J.RETURN_DEPT_TIME]),
        returnFlight: row[J.RETURN_FLIGHT],
        returnCity: row[J.RETURN_DEPT_CITY],
        
        hotelCity: actualCity,
        hotelCheckIn: hotelDates.checkIn,
        hotelCheckOut: hotelDates.checkOut,
        phase: hotelDates.phase,
        
        expectedArrival: expectedArrival,
        arrivalStatus: arrivalStatus.status,
        earlyHours: arrivalStatus.earlyHours,
        lateDays: arrivalStatus.lateDays,
        
        roomType: roomType,
        sharedBeds: sharedBeds,
        fullRooms: fullRooms,
        capacity: getRoomCapacity_(roomType),
        
        packageName: gi.packageName || '',
        groupLeader: gi.groupLeader || '',
        leaderPhone: gi.phone || '',
        tourGuide: getGuideForPilgrim_(guideMap, row[J.PASSPORT]) || gi.tourGuide || '',
        transport: transport,
        nextHotelName: nextHotelName,
        
        checkInStatus: checkin.status || 'pending',
        roomNumber: checkin.roomNumber || '',
        checkInTime: checkin.checkInTime || '',
        roomGroupId: checkin.roomGroupId || ''
      });
    }
  }
  
  return pilgrims;
}

// ============================================================
// HELPER: Get Next Hotel Name (for departure destination)
// ============================================================

function getNextHotelName_(row, currentHotelCity) {
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var lastHouse  = row[J.LAST_HOUSE];

  // من المدينة → مكة (الأولى أو التحويل)
  if (currentHotelCity === 'Madina') {
    var makkahEn = row[J.MAKKAH_EN];
    var shiftEn  = row[J.MAKKAH_SHIFT_EN];
    if (makkahEn && makkahEn !== 'NULL') return makkahEn;
    if (shiftEn  && shiftEn  !== 'NULL') return shiftEn;
    return '';
  }

  // من مكة الرئيسية
  if (currentHotelCity === 'Makkah') {
    // تحويل فندقي: مكة → مكة Shifting
    var shiftEn = row[J.MAKKAH_SHIFT_EN];
    if (shiftEn && shiftEn !== 'NULL') return shiftEn;
    // مكة → المدينة (lastHouse)
    if (lastHouse === 'Madina') {
      var madEn = row[J.MADINAH_EN];
      if (madEn && madEn !== 'NULL') return madEn;
    }
    return '';
  }

  // من مكة تحويل → المدينة
  if (currentHotelCity === 'Makkah Shifting') {
    if (lastHouse === 'Madina') {
      var madEn = row[J.MADINAH_EN];
      if (madEn && madEn !== 'NULL') return madEn;
    }
    return '';
  }

  return '';
}

// ============================================================
// DATA: GET DEPARTURES FOR HOTEL
// ============================================================

function getDeparturesForHotel(hotelName, hotelCity) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var transportMap = getTransportMap_(ss);
  
  var departures = [];
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var actualCity = getActualHotelCity_(row, hotelName, hotelCity);
    
    if (actualCity) {
      var hotelDates = getHotelDates_(row, actualCity);
      
      // Skip No Show for Madinah
      if (actualCity === 'Madina' && hotelDates.phase === 'first') {
        var arrDate = row[J.ARRIVAL_DATE];
        var contractEnd = row[J.FIRST_HOUSE_END];
        if (arrDate && contractEnd) {
          var ad = new Date(arrDate); ad.setHours(0,0,0,0);
          var ce = new Date(contractEnd); ce.setHours(0,0,0,0);
          if (ad > ce) continue;
        }
      }
      
      var packageId = row[J.PACKAGE_ID];
      var transport = transportMap[packageId] || 'حافلة';
      var departure = calculateDeparture_(row, actualCity, hotelDates, transport);

      if (departure) {
        departures.push({
          bookingId: row[J.BOOKING_ID],
          name: row[J.NAME],
          groupNumber: row[J.GROUP_NUMBER],
          passport: row[J.PASSPORT],
          departureDate: hotelDates.checkOut,
          departureTime: departure.time,
          destination: departure.destination,
          destinationHotel: departure.destinationHotel,
          transport: departure.transport,
          type: departure.type,
          linkedFlight: departure.linkedFlight,
          linkedTime: departure.linkedTime
        });}
    }
  }
  
  return departures;
}
// ============================================================
// DATA: GET ROOM INVENTORY
// ============================================================

function getRoomInventory(hotelNameEN, hotelCity) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var pkgSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.PACKAGES);
  var pkgData = pkgSheet.getDataRange().getValues();
  
  var inventory = {};
  
  for (var i = 2; i < pkgData.length; i++) {
    var row = pkgData[i];
    if (!row[0]) continue;
    
    var hotelSlots = [
      { city: row[11], hotelEN: row[13], checkIn: row[14], checkOut: row[15], dbl: row[16], tri: row[18], quad: row[20] },
      { city: row[26], hotelEN: row[28], checkIn: row[29], checkOut: row[30], dbl: row[31], tri: row[33], quad: row[35] },
      { city: row[41], hotelEN: row[43], checkIn: row[44], checkOut: row[45], dbl: row[46], tri: row[48], quad: row[50] }
    ];
    
    for (var s = 0; s < hotelSlots.length; s++) {
      var slot = hotelSlots[s];
      if (!slot.hotelEN || !slot.checkIn || !slot.checkOut) continue;
      
      if (slot.hotelEN === hotelNameEN || normalizeHotelName_(slot.hotelEN) === normalizeHotelName_(hotelNameEN)) {
        var startDate = new Date(slot.checkIn);
        var endDate = new Date(slot.checkOut);
        var dbl = Number(slot.dbl) || 0;
        var tri = Number(slot.tri) || 0;
        var quad = Number(slot.quad) || 0;
        var rooms = dbl + tri + quad;
        var beds = (dbl * 2) + (tri * 3) + (quad * 4);
        
        for (var d = new Date(startDate); d < endDate; d.setDate(d.getDate() + 1)) {
          var key = Utilities.formatDate(d, 'Asia/Riyadh', 'yyyy-MM-dd');
          if (!inventory[key]) inventory[key] = { dbl: 0, tri: 0, quad: 0, rooms: 0, beds: 0 };
          inventory[key].dbl += dbl;
          inventory[key].tri += tri;
          inventory[key].quad += quad;
          inventory[key].rooms += rooms;
          inventory[key].beds += beds;
        }
      }
    }
  }
  
  // ── إضافة حالة الإشغال من Room Mapping ──
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  var roomStats = { total: 0, occupied: 0, available: 0, byType: {}, byPeriod: [] };
  
  if (mapSheet) {
    var mapData = mapSheet.getDataRange().getValues();
    var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;
    
    // تجميع حسب الفترة + النوع
    var periodMap = {};
    
    for (var i = 1; i < mapData.length; i++) {
      var hn = String(mapData[i][MC.HOTEL_NAME] || '');
      if (hn !== hotelNameEN && normalizeHotelName_(hn) !== normalizeHotelName_(hotelNameEN)) continue;
      
      var rt = String(mapData[i][MC.ROOM_TYPE] || '');
      var status = String(mapData[i][MC.STATUS] || 'Available');
      var ci = formatDate_(mapData[i][MC.CHECK_IN]);
      var co = formatDate_(mapData[i][MC.CHECK_OUT]);
      
      roomStats.total++;
      if (status === 'Occupied') roomStats.occupied++;
      else roomStats.available++;
      
      if (!roomStats.byType[rt]) roomStats.byType[rt] = { total: 0, occupied: 0, available: 0 };
      roomStats.byType[rt].total++;
      if (status === 'Occupied') roomStats.byType[rt].occupied++;
      else roomStats.byType[rt].available++;
      
      // تجميع حسب الفترة
      var periodKey = ci + '|' + co;
      if (!periodMap[periodKey]) periodMap[periodKey] = { checkIn: ci, checkOut: co, types: {} };
      if (!periodMap[periodKey].types[rt]) periodMap[periodKey].types[rt] = { total: 0, occupied: 0 };
      periodMap[periodKey].types[rt].total++;
      if (status === 'Occupied') periodMap[periodKey].types[rt].occupied++;
    }
    
    // تحويل إلى مصفوفة مرتّبة
    var periods = [];
    for (var pk in periodMap) {
      var p = periodMap[pk];
      var pTotal = 0, pOccupied = 0;
      var typesArr = [];
      for (var rt in p.types) {
        var t = p.types[rt];
        pTotal += t.total;
        pOccupied += t.occupied;
        typesArr.push({ type: rt, total: t.total, occupied: t.occupied, available: t.total - t.occupied });
      }
      periods.push({
        checkIn: p.checkIn, checkOut: p.checkOut,
        total: pTotal, occupied: pOccupied, available: pTotal - pOccupied,
        types: typesArr
      });
    }
    periods.sort(function(a, b) { return a.checkIn < b.checkIn ? -1 : a.checkIn > b.checkIn ? 1 : 0; });
    roomStats.byPeriod = periods;
  }
  
  return { inventory: inventory, roomStats: roomStats };
}


// ============================================================
// DATA: DASHBOARD STATS — مع كاش 30 دقيقة
// ============================================================

function getDashboardStats() {
  // محاولة قراءة من الكاش أولاً
  var cached = getCachedData_('dashboardStats');
  if (cached) return cached;

  // حساب مباشر إذا لا يوجد كاش
  var stats = computeDashboardStats_();
  setCachedData_('dashboardStats', stats);
  return stats;
}

// يُستدعى من trigger كل 30 دقيقة لتحديث الكاش
function updateDashboardCache() {
  var stats = computeDashboardStats_();
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(stats);
    cache.put('dashboardStats', json, 1800); // 30 دقيقة
  } catch (e) {
    Logger.log('updateDashboardCache ERROR: ' + e.toString());
  }
}

// إعداد trigger تلقائي — يُشغّل مرة واحدة من Script Editor
function setupDashboardTrigger() {
  // حذف triggers قديمة
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateDashboardCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // إنشاء trigger جديد كل 30 دقيقة
  ScriptApp.newTrigger('updateDashboardCache')
    .timeBased()
    .everyMinutes(30)
    .create();
  Logger.log('✅ تم إعداد trigger الداشبورد — كل 30 دقيقة');
}

function computeDashboardStats_() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();

  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var stats = {};
  var allCheckin = {};

  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var firstHouse = row[J.FIRST_HOUSE];
    var hotelName = getHotelName_(row, firstHouse);
    var actualCity = (firstHouse === 'Makkah' || firstHouse === 'Makkah Shifting') ? 'مكة' : firstHouse === 'Madina' ? 'المدينة' : firstHouse;

    if (!hotelName || hotelName === 'NULL') continue;

    // Check No Show for Madinah
    if (firstHouse === 'Madina') {
      var arrDate = row[J.ARRIVAL_DATE];
      var contractEnd = row[J.FIRST_HOUSE_END];
      if (arrDate && contractEnd) {
        var ad = new Date(arrDate); ad.setHours(0,0,0,0);
        var ce = new Date(contractEnd); ce.setHours(0,0,0,0);
        if (ad > ce) {
          if (!stats[hotelName]) {
            stats[hotelName] = { hotel: hotelName, city: actualCity, total: 0, arrived: 0, pending: 0, delayed: 0, noShow: 0, checkinPct: 0 };
          }
          stats[hotelName].total++;
          stats[hotelName].noShow++;
          continue;
        }
      }
    }

    if (!stats[hotelName]) {
      stats[hotelName] = { hotel: hotelName, city: actualCity, total: 0, arrived: 0, pending: 0, delayed: 0, noShow: 0, checkinPct: 0 };
    }

    stats[hotelName].total++;

    if (!allCheckin[hotelName]) {
      allCheckin[hotelName] = getCheckinData_(ss, hotelName);
    }

    var bookingId = row[J.BOOKING_ID];
    var checkin = allCheckin[hotelName][bookingId];

    if (checkin && checkin.status === 'arrived') {
      stats[hotelName].arrived++;
    } else if (checkin && checkin.status === 'delayed') {
      stats[hotelName].delayed++;
    } else {
      stats[hotelName].pending++;
    }
  }

  for (var key in stats) {
    var s = stats[key];
    var effective = s.total - s.noShow;
    s.checkinPct = effective > 0 ? Math.round((s.arrived / effective) * 100) : 0;
  }

  return stats;
}

// ============================================================
// ROOM MAPPING: توليد أرقام الغرف من الباقات
// يُشغّل: زر في واجهة المخزون
// يقرأ: الباقات + Room Mapping (الموجود)
// يكتب: صفوف جديدة فقط في Room Mapping — لا يحذف القديم
// ============================================================

function generateRoomMapping(hotelName, hotelCity) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, error: 'النظام مشغول — حاول مرة أخرى' };

  try {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var pkgSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.PACKAGES);
  var pkgData = pkgSheet.getDataRange().getValues();
  
  // === 1. إنشاء شيت Room Mapping إن لم يكن موجوداً ===
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  if (!mapSheet) {
    mapSheet = ss.insertSheet(HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
    mapSheet.appendRow(HOTEL_CONFIG.ROOM_MAPPING_HEADERS);
    mapSheet.getRange(1, 1, 1, HOTEL_CONFIG.ROOM_MAPPING_HEADERS.length).setFontWeight('bold');
    mapSheet.setFrozenRows(1);
  }
  
  // === 2. قراءة الغرف الموجودة — لتجنب التكرار ===
  var mapData = mapSheet.getDataRange().getValues();
  var existingCount = {};
  var allExistingIds = [];
  var hotelExisting = 0;
  
  for (var i = 1; i < mapData.length; i++) {
    if (mapData[i][1]) allExistingIds.push(String(mapData[i][1]));
    if (mapData[i][0] !== hotelName) continue;
    hotelExisting++;
    // مفتاح فريد: نوع|دخول|خروج|باقة
    var key = mapData[i][2] + '|' + mapData[i][4] + '|' + mapData[i][5] + '|' + mapData[i][6];
    existingCount[key] = (existingCount[key] || 0) + 1;
  }
  
  // === 3. تحليل الباقات — حساب الغرف المطلوبة ===
  var needed = [];
  
  for (var i = 2; i < pkgData.length; i++) {
    var row = pkgData[i];
    if (!row[0]) continue;
    
    var hotelSlots = [
      { hotelEN: row[13], checkIn: row[14], checkOut: row[15], dbl: row[16], tri: row[18], quad: row[20] },
      { hotelEN: row[28], checkIn: row[29], checkOut: row[30], dbl: row[31], tri: row[33], quad: row[35] },
      { hotelEN: row[43], checkIn: row[44], checkOut: row[45], dbl: row[46], tri: row[48], quad: row[50] }
    ];
    
    for (var s = 0; s < hotelSlots.length; s++) {
      var slot = hotelSlots[s];
      if (!slot.hotelEN || !slot.checkIn || !slot.checkOut) continue;
      if (slot.hotelEN !== hotelName && normalizeHotelName_(slot.hotelEN) !== normalizeHotelName_(hotelName)) continue;
      
      var ci = formatDate_(slot.checkIn);
      var co = formatDate_(slot.checkOut);
      var pkgRef = String(row[0]);
      
      var types = [
        { type: 'Quad', capacity: 4, count: Number(slot.quad) || 0 },
        { type: 'Triple', capacity: 3, count: Number(slot.tri) || 0 },
        { type: 'Double', capacity: 2, count: Number(slot.dbl) || 0 }
      ];
      
      for (var t = 0; t < types.length; t++) {
        if (types[t].count <= 0) continue;
        needed.push({
          type: types[t].type, capacity: types[t].capacity,
          checkIn: ci, checkOut: co, packageRef: pkgRef,
          count: types[t].count
        });
      }
    }
  }
  
  // === 4. حساب الفارق وتوليد الجديد فقط ===
  var newRows = [];
  var trackIds = allExistingIds.slice();
  
  for (var n = 0; n < needed.length; n++) {
    var item = needed[n];
    var key = item.type + '|' + item.checkIn + '|' + item.checkOut + '|' + item.packageRef;
    var existing = existingCount[key] || 0;
    var toGenerate = item.count - existing;
    
    for (var r = 0; r < toGenerate; r++) {
      var roomId = generateInternalRoomId_(hotelName, hotelCity, item.type, trackIds);
      trackIds.push(roomId);
      newRows.push([
        hotelName, roomId, item.type, item.capacity,
        item.checkIn, item.checkOut, item.packageRef, ''
      ]);
    }
  }
  
  // === 5. كتابة بالجملة ===
  if (newRows.length > 0) {
    mapSheet.getRange(mapSheet.getLastRow() + 1, 1, newRows.length, 8).setValues(newRows);
    SpreadsheetApp.flush();
  }
  
  return {
    success: true,
    generated: newRows.length,
    existing: hotelExisting,
    total: hotelExisting + newRows.length
  };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ROOM MAPPING: قراءة جدول الربط لفندق محدد
// يُشغّل: عند فتح تبويب المخزون
// يقرأ: شيت Room Mapping
// يُرجع: مصفوفة الغرف مع الأرقام الفعلية
// ============================================================

function getRoomMappingForHotel(hotelName) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  if (!mapSheet) return [];
  
  var data = mapSheet.getDataRange().getValues();
  var rooms = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== hotelName) continue;
    rooms.push({
      internalId: data[i][1],
      roomType: data[i][2],
      capacity: data[i][3],
      checkIn: formatDate_(data[i][4]),
      checkOut: formatDate_(data[i][5]),
      status: data[i][8] || 'Available',
      occupantIds: data[i][9] || '',
      actualRoomNo: data[i][7] || ''
    });
  }
  
  // ترتيب: نوع ← فترة ← تسلسل
  rooms.sort(function(a, b) {
    if (a.roomType !== b.roomType) return a.roomType > b.roomType ? 1 : -1;
    if (a.checkIn !== b.checkIn) return a.checkIn > b.checkIn ? 1 : -1;
    return a.internalId > b.internalId ? 1 : -1;
  });
  
  return rooms;
}
function debugRoomMapping() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var pkgSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.PACKAGES);
  var pkgData = pkgSheet.getDataRange().getValues();
  
  var hotels = {};
  for (var i = 2; i < pkgData.length; i++) {
    var row = pkgData[i];
    if (!row[0]) continue;
    var slots = [
      { col: 13, name: row[13] },
      { col: 28, name: row[28] },
      { col: 43, name: row[43] }
    ];
    slots.forEach(function(s) {
      if (s.name) hotels[s.name] = (hotels[s.name] || 0) + 1;
    });
  }
  
  Logger.log('=== فنادق الباقات (أعمدة 13/28/43) ===');
  for (var h in hotels) {
    Logger.log('"' + h + '" → ' + hotels[h] + ' باقة');
  }
}
// ============================================================
// ROOM MAPPING: توليد لجميع الفنادق دفعة واحدة
// يُشغّل: يدوياً من Script Editor
// يقرأ: الباقات + رحلة الحاج
// يكتب: شيت Room Mapping (مسح + إعادة بناء كامل)
// ============================================================

function generateAllRoomMappings() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var pkgSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.PACKAGES);
  var pkgData = pkgSheet.getDataRange().getValues();
  
  // ── 1. بناء خريطة فندق → مدينة من رحلة الحاج ──
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  if (!journeySheet) {
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (sheets[s].getName().trim() === HOTEL_CONFIG.SHEETS.JOURNEY.trim()) {
        journeySheet = sheets[s]; break;
      }
    }
  }
  var journeyData = journeySheet.getDataRange().getValues();
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  
  var hotelCityMap = {};
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var medH = String(row[J.MADINAH_EN] || '').trim();
    if (medH && medH !== 'NULL') hotelCityMap[medH] = 'Madina';
    var makH = String(row[J.MAKKAH_EN] || '').trim();
    if (makH && makH !== 'NULL') hotelCityMap[makH] = 'Makkah';
    var shiftH = String(row[J.MAKKAH_SHIFT_EN] || '').trim();
    if (shiftH && shiftH !== 'NULL') hotelCityMap[shiftH] = 'Makkah';
  }
  
  // ── 2. مسح Room Mapping وإعادة إنشاء ──
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  if (mapSheet) {
    mapSheet.clear();
  } else {
    mapSheet = ss.insertSheet(HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  }
  var headers = HOTEL_CONFIG.ROOM_MAPPING_HEADERS;
  mapSheet.appendRow(headers);
  mapSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  mapSheet.setFrozenRows(1);
  
  // ── 3. قراءة الباقات وتوليد كل الغرف ──
  var allRooms = [];
  var trackIds = [];
  
  for (var i = 2; i < pkgData.length; i++) {
    var row = pkgData[i];
    if (!row[0]) continue;
    var pkgRef = String(row[0]);
    
    var slots = [
      { hotelEN: row[13], checkIn: row[14], checkOut: row[15], dbl: row[16], tri: row[18], quad: row[20] },
      { hotelEN: row[28], checkIn: row[29], checkOut: row[30], dbl: row[31], tri: row[33], quad: row[35] },
      { hotelEN: row[43], checkIn: row[44], checkOut: row[45], dbl: row[46], tri: row[48], quad: row[50] }
    ];
    
    for (var s = 0; s < slots.length; s++) {
      var slot = slots[s];
      if (!slot.hotelEN || !slot.checkIn || !slot.checkOut) continue;
      
      var hotelName = String(slot.hotelEN).trim();
      var city = hotelCityMap[hotelName] || 'Makkah';
      var ci = formatDate_(slot.checkIn);
      var co = formatDate_(slot.checkOut);
      
      var types = [
        { type: 'Quad',   capacity: 4, count: Number(slot.quad) || 0 },
        { type: 'Triple', capacity: 3, count: Number(slot.tri)  || 0 },
        { type: 'Double', capacity: 2, count: Number(slot.dbl)  || 0 }
      ];
      
      for (var t = 0; t < types.length; t++) {
        if (types[t].count <= 0) continue;
        for (var r = 0; r < types[t].count; r++) {
          var roomId = generateInternalRoomId_(hotelName, city, types[t].type, trackIds);
          trackIds.push(roomId);
          allRooms.push([
            hotelName, roomId, types[t].type, types[t].capacity,
            ci, co, pkgRef, '', 'Available', ''
          ]);
        }
      }
    }
  }
  
  // ── 4. كتابة بالجملة ──
  if (allRooms.length > 0) {
    mapSheet.getRange(2, 1, allRooms.length, allRooms[0].length).setValues(allRooms);
    SpreadsheetApp.flush();
  }
  
  // ── 5. ملخص ──
  var summary = {};
  for (var i = 0; i < allRooms.length; i++) {
    var h = allRooms[i][0];
    if (!summary[h]) summary[h] = { count: 0, city: hotelCityMap[h] || '?' };
    summary[h].count++;
  }
  
  Logger.log('═══ Room Mapping Generated ═══');
  Logger.log('إجمالي الغرف: ' + allRooms.length);
  Logger.log('الفنادق: ' + Object.keys(summary).length);
  for (var h in summary) {
    Logger.log('  ' + h + ' (' + summary[h].city + '): ' + summary[h].count + ' غرفة');
  }
  
  return { success: true, total: allRooms.length, hotels: Object.keys(summary).length };
}

// ============================================================
// USERS: تسجيل الدخول بالبريد الإلكتروني
// شيت Users: البريد (A) | الاسم (B) | الدور (C) | الفندق (D) | فعّال (E)
// ============================================================

function getLoginUser() {
  var email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  if (!email) return { success: false, error: 'لا يمكن تحديد المستخدم — تأكد من تسجيل الدخول بحساب Google' };

  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, 'Users');

  if (!sheet) {
    // إذا لم يوجد شيت Users — السماح للجميع (وضع مفتوح)
    return { success: true, email: email, name: email, role: 'ops', hotel: '', openMode: true };
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase()) {
      if (String(data[i][4]).trim() !== '✅') {
        return { success: false, error: 'حسابك معطّل — تواصل مع المشرف' };
      }
      return {
        success: true,
        email: email,
        name: data[i][1] || email,
        role: data[i][2] || 'hotel',
        hotel: data[i][3] || ''
      };
    }
  }

  return { success: false, error: 'البريد ' + email + ' غير مسجل في النظام' };
}

// ============================================================
// AUDIT LOG: تسجيل العمليات
// شيت Audit_Log: التاريخ (A) | البريد (B) | العملية (C) | التفاصيل (D) | الفندق (E)
// ============================================================

function logAudit_(action, details, hotelName) {
  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var sheet = findSheet_(ss, 'Audit_Log');

    if (!sheet) {
      sheet = ss.insertSheet('Audit_Log');
      sheet.appendRow(['التاريخ', 'البريد', 'العملية', 'التفاصيل', 'الفندق']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var email = Session.getActiveUser().getEmail() || 'unknown';
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([now, email, action, details, hotelName || '']);
  } catch (e) {
    Logger.log('logAudit_ ERROR: ' + e.toString());
  }
}

// ============================================================
// ROOMMATES: الحصول على زملاء الغرفة
// ============================================================

function getRoommates(roomGroupId, hotelName) {
  if (!roomGroupId || !hotelName) return [];

  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = getHotelSheet_(ss, hotelName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var roommates = [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][16]) === String(roomGroupId)) {
      roommates.push({
        name: data[i][1],
        passport: data[i][2],
        gender: data[i][3],
        status: data[i][18] || 'pending'
      });
    }
  }

  return roommates;
}