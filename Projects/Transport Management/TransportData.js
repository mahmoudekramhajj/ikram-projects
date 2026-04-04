/**
 * Transport Management App — Data Layer
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0 — قراءة بيانات الرحلات والحجاج والمركبات
 */

// ============================================================
// DATA: الحصول على رحلات النقل حسب الموقع
// ============================================================

function getTripsForLocation(locationId, dateFilter) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRIPS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var TC = TRANSPORT_CONFIG.TRIP_COLS;
  var trips = [];

  for (var i = 1; i < data.length; i++) {
    var row = cleanRow_(data[i]);
    var origin = row[TC.ORIGIN];
    var dest = row[TC.DESTINATION];
    var date = row[TC.DATE];
    var status = row[TC.STATUS];

    // فلتر التاريخ
    if (dateFilter && date !== dateFilter) continue;

    // فلتر الموقع — نعرض الرحلات التي تنطلق من أو تصل إلى هذا الموقع
    var matchesLocation = matchesLocationFilter_(origin, dest, locationId);
    if (!matchesLocation) continue;

    trips.push({
      tripId: row[TC.TRIP_ID],
      operationType: Number(row[TC.OPERATION_TYPE]),
      operationName: row[TC.OPERATION_NAME],
      date: date,
      scheduledTime: row[TC.SCHEDULED_TIME],
      origin: origin,
      originType: row[TC.ORIGIN_TYPE],
      destination: dest,
      destinationType: row[TC.DESTINATION_TYPE],
      vehiclePlate: row[TC.VEHICLE_PLATE],
      driverName: row[TC.DRIVER_NAME],
      driverPhone: row[TC.DRIVER_PHONE],
      capacity: Number(row[TC.CAPACITY]) || 0,
      assignedCount: Number(row[TC.ASSIGNED_COUNT]) || 0,
      boardedCount: Number(row[TC.BOARDED_COUNT]) || 0,
      status: status || 'scheduled',
      linkedFlight: row[TC.LINKED_FLIGHT],
      linkedFlightTime: row[TC.LINKED_FLIGHT_TIME],
      packageIds: row[TC.PACKAGE_IDS],
      guideNames: row[TC.GUIDE_NAMES],
      notes: row[TC.NOTES],
      sheetRow: i + 1
    });
  }

  // ترتيب بالوقت
  trips.sort(function(a, b) {
    return (a.scheduledTime || '99:99').localeCompare(b.scheduledTime || '99:99');
  });

  return trips;
}

// ============================================================
// DATA: الحصول على جميع الرحلات (لوحة التحكم)
// ============================================================

function getAllTrips(dateFilter) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRIPS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var TC = TRANSPORT_CONFIG.TRIP_COLS;
  var trips = [];

  for (var i = 1; i < data.length; i++) {
    var row = cleanRow_(data[i]);
    if (dateFilter && row[TC.DATE] !== dateFilter) continue;

    trips.push({
      tripId: row[TC.TRIP_ID],
      operationType: Number(row[TC.OPERATION_TYPE]),
      operationName: row[TC.OPERATION_NAME],
      date: row[TC.DATE],
      scheduledTime: row[TC.SCHEDULED_TIME],
      origin: row[TC.ORIGIN],
      destination: row[TC.DESTINATION],
      assignedCount: Number(row[TC.ASSIGNED_COUNT]) || 0,
      boardedCount: Number(row[TC.BOARDED_COUNT]) || 0,
      status: row[TC.STATUS] || 'scheduled',
      linkedFlight: row[TC.LINKED_FLIGHT],
      sheetRow: i + 1
    });
  }

  return trips;
}

// ============================================================
// DATA: الحصول على سجل ركوب رحلة معينة
// ============================================================

function getBoardingForTrip(tripId) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.BOARDING);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var BC = TRANSPORT_CONFIG.BOARDING_COLS;
  var records = [];

  for (var i = 1; i < data.length; i++) {
    var row = cleanRow_(data[i]);
    if (row[BC.TRIP_ID] !== tripId) continue;

    records.push({
      boardingId: row[BC.BOARDING_ID],
      tripId: row[BC.TRIP_ID],
      bookingId: row[BC.BOOKING_ID],
      passport: row[BC.PASSPORT],
      pilgrimName: row[BC.PILGRIM_NAME],
      groupNumber: row[BC.GROUP_NUMBER],
      packageId: row[BC.PACKAGE_ID],
      scanTime: row[BC.SCAN_TIME],
      scanBy: row[BC.SCAN_BY],
      scanMethod: row[BC.SCAN_METHOD],
      declaredStatus: row[BC.DECLARED_STATUS],
      boardingStatus: row[BC.BOARDING_STATUS],
      notes: row[BC.NOTES]
    });
  }

  return records;
}

// ============================================================
// DATA: البحث عن حاج بالجواز أو BookingId (لمسح QR)
// ============================================================

function lookupPilgrim(searchValue) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.JOURNEY);
  if (!sheet) return { success: false, error: 'شيت رحلة الحاج غير موجود' };

  var data = sheet.getDataRange().getValues();
  var J = TRANSPORT_CONFIG.JOURNEY_COLS;
  var search = String(searchValue).trim();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var bookingId = String(row[J.BOOKING_ID] || '').trim();
    var passport = String(row[J.PASSPORT] || '').trim();

    if (bookingId === search || passport === search) {
      return {
        success: true,
        pilgrim: {
          bookingId: bookingId,
          packageId: String(row[J.PACKAGE_ID] || '').trim(),
          groupNumber: String(row[J.GROUP_NUMBER] || '').trim(),
          name: String(row[J.NAME] || '').trim(),
          passport: passport,
          gender: String(row[J.GENDER] || '').trim(),
          nationalityEn: String(row[J.NATIONALITY_EN] || '').trim(),
          nationalityAr: String(row[J.NATIONALITY_AR] || '').trim(),
          campName: String(row[J.CAMP_NAME] || '').trim(),
          arrivalCity: String(row[J.ARRIVAL_CITY] || '').trim(),
          arrivalDate: formatDate_(row[J.ARRIVAL_DATE]),
          arrivalTime: formatTime_(row[J.ARRIVAL_TIME]),
          arrivalFlight: String(row[J.ARRIVAL_FLIGHT] || '').trim(),
          returnDeptCity: String(row[J.RETURN_DEPT_CITY] || '').trim(),
          returnDeptDate: formatDate_(row[J.RETURN_DEPT_DATE]),
          returnDeptTime: formatTime_(row[J.RETURN_DEPT_TIME]),
          returnFlight: String(row[J.RETURN_FLIGHT] || '').trim(),
          firstHouse: String(row[J.FIRST_HOUSE] || '').trim(),
          firstHouseStart: formatDate_(row[J.FIRST_HOUSE_START]),
          firstHouseEnd: formatDate_(row[J.FIRST_HOUSE_END]),
          lastHouse: String(row[J.LAST_HOUSE] || '').trim(),
          lastHouseStart: formatDate_(row[J.LAST_HOUSE_START]),
          lastHouseEnd: formatDate_(row[J.LAST_HOUSE_END]),
          makkahEn: String(row[J.MAKKAH_EN] || '').trim(),
          makkahShiftEn: String(row[J.MAKKAH_SHIFT_EN] || '').trim(),
          madinahEn: String(row[J.MADINAH_EN] || '').trim()
        }
      };
    }
  }

  return { success: false, error: 'لم يتم العثور على الحاج: ' + search };
}

// ============================================================
// DATA: الحصول على المركبات
// ============================================================

function getVehicles() {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.VEHICLES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var vehicles = [];

  for (var i = 1; i < data.length; i++) {
    var row = cleanRow_(data[i]);
    if (!row[0]) continue;
    vehicles.push({
      vehicleId: row[0],
      contractId: row[1],
      type: row[2],
      plateNumber: row[3],
      capacity: Number(row[4]) || 0,
      driverName: row[5],
      driverPhone: row[6],
      companyName: row[7],
      status: row[8] || 'active'
    });
  }

  return vehicles;
}

// ============================================================
// DATA: الحصول على الحوادث
// ============================================================

function getIncidents(statusFilter) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.INCIDENTS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var incidents = [];

  for (var i = 1; i < data.length; i++) {
    var row = cleanRow_(data[i]);
    if (!row[0]) continue;
    if (statusFilter && row[11] !== statusFilter) continue;

    incidents.push({
      incidentId: row[0],
      tripId: row[1],
      type: row[2],
      severity: row[3],
      description: row[4],
      affectedPilgrims: Number(row[5]) || 0,
      resolution: row[6],
      resolvedBy: row[7],
      reportedBy: row[8],
      reportedAt: row[9],
      resolvedAt: row[10],
      status: row[11] || 'open'
    });
  }

  return incidents;
}

// ============================================================
// DATA: لوحة التحكم — إحصائيات اليوم
// ============================================================

function getDashboardStats(dateFilter) {
  var trips = getAllTrips(dateFilter);

  var stats = {
    totalTrips: trips.length,
    totalAssigned: 0,
    totalBoarded: 0,
    activeTrips: 0,
    completedTrips: 0,
    cancelledTrips: 0,
    incidents: 0,
    operationStats: {}
  };

  for (var i = 0; i < trips.length; i++) {
    var t = trips[i];
    stats.totalAssigned += t.assignedCount;
    stats.totalBoarded += t.boardedCount;

    if (t.status === 'arrived') stats.completedTrips++;
    else if (t.status === 'cancelled') stats.cancelledTrips++;
    else stats.activeTrips++;

    // تجميع حسب نوع العملية
    var opKey = String(t.operationType);
    if (!stats.operationStats[opKey]) {
      stats.operationStats[opKey] = { trips: 0, assigned: 0, boarded: 0, status: 'gray' };
    }
    stats.operationStats[opKey].trips++;
    stats.operationStats[opKey].assigned += t.assignedCount;
    stats.operationStats[opKey].boarded += t.boardedCount;

    // تحديد لون العملية
    if (t.status === 'cancelled') {
      stats.operationStats[opKey].status = 'red';
    } else if (t.status === 'arrived') {
      if (stats.operationStats[opKey].status !== 'red') {
        stats.operationStats[opKey].status = 'green';
      }
    } else {
      if (stats.operationStats[opKey].status === 'gray') {
        stats.operationStats[opKey].status = 'yellow';
      }
    }
  }

  // عدد الحوادث
  var incidents = getIncidents('open');
  stats.incidents = incidents.length;

  return stats;
}

// ============================================================
// HELPER: مطابقة الموقع
// ============================================================

function matchesLocationFilter_(origin, dest, locationId) {
  if (!locationId || locationId === 'all') return true;

  var locationNames = {
    'madinah_airport': 'مطار المدينة',
    'jeddah_t1': 'مطار جدة صالة 1',
    'jeddah_north': 'مطار جدة الشمالية',
    'madinah_station': 'محطة قطار المدينة',
    'makkah_station': 'محطة قطار مكة',
    'madinah_hotels': 'فنادق المدينة',
    'makkah_hotels': 'فنادق مكة',
    'camp_maisam': 'مخيم المعيصم',
    'camp_mujar': 'مخيم مجر الكبش',
    'camp_72': 'مخيم 72',
    'arafat': 'عرفة',
    'muzdalifah': 'مزدلفة'
  };

  var name = locationNames[locationId] || locationId;
  return origin.indexOf(name) >= 0 || dest.indexOf(name) >= 0;
}
