/**
 * Transport Management App — Actions Layer
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0 — إنشاء رحلات، تسجيل ركوب، إبلاغ عن حوادث
 */

// ============================================================
// ACTION: إنشاء رحلة نقل جديدة
// ============================================================

function createTrip(tripData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'النظام مشغول، حاول مرة أخرى' };
  }

  try {
    var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
    var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRIPS);
    if (!sheet) return { success: false, error: 'شيت الرحلات غير موجود' };

    var op = OPERATIONS[tripData.operationType];
    if (!op) return { success: false, error: 'نوع العملية غير صالح: ' + tripData.operationType };

    var now = formatDateTime_(new Date());
    var tripId = generateTripId_();

    var newRow = [
      tripId,
      tripData.operationType,
      op.name,
      tripData.date || '',
      tripData.scheduledTime || '',
      tripData.origin || op.origin,
      tripData.originType || op.originType,
      tripData.destination || op.dest,
      tripData.destinationType || op.destType,
      tripData.vehicleId || '',
      tripData.vehiclePlate || '',
      tripData.driverName || '',
      tripData.driverPhone || '',
      tripData.capacity || 0,
      tripData.assignedCount || 0,
      0, // boardedCount
      'scheduled',
      tripData.linkedFlight || '',
      tripData.linkedFlightTime || '',
      tripData.packageIds || '',
      tripData.guideNames || '',
      tripData.notes || '',
      tripData.createdBy || '',
      now,
      now
    ];

    sheet.appendRow(newRow);
    clearCache_('trips_');

    return { success: true, tripId: tripId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTION: تحديث حالة الرحلة
// ============================================================

function updateTripStatus(tripId, newStatus) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'النظام مشغول' };
  }

  try {
    var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
    var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRIPS);
    var data = sheet.getDataRange().getValues();
    var TC = TRANSPORT_CONFIG.TRIP_COLS;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][TC.TRIP_ID]).trim() === tripId) {
        var row = i + 1;
        sheet.getRange(row, TC.STATUS + 1).setValue(newStatus);
        sheet.getRange(row, TC.UPDATED_AT + 1).setValue(formatDateTime_(new Date()));
        clearCache_('trips_');
        return { success: true };
      }
    }

    return { success: false, error: 'الرحلة غير موجودة: ' + tripId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTION: تسجيل ركوب حاج (QR Scan)
// ============================================================

function recordBoarding(tripId, scannedValue, scanMethod, scanBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'النظام مشغول' };
  }

  try {
    var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);

    // 1. البحث عن الحاج
    var pilgrimResult = lookupPilgrim(scannedValue);
    if (!pilgrimResult.success) {
      return { success: false, error: pilgrimResult.error };
    }
    var pilgrim = pilgrimResult.pilgrim;

    // 2. التحقق من عدم التسجيل المسبق
    var boardingSheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.BOARDING);
    if (boardingSheet && boardingSheet.getLastRow() > 1) {
      var boardingData = boardingSheet.getDataRange().getValues();
      var BC = TRANSPORT_CONFIG.BOARDING_COLS;
      for (var i = 1; i < boardingData.length; i++) {
        if (String(boardingData[i][BC.TRIP_ID]).trim() === tripId &&
            String(boardingData[i][BC.PASSPORT]).trim() === pilgrim.passport &&
            String(boardingData[i][BC.BOARDING_STATUS]).trim() === 'boarded') {
          return {
            success: false,
            error: 'الحاج مسجّل مسبقاً في هذه الرحلة',
            pilgrim: pilgrim,
            alreadyBoarded: true
          };
        }
      }
    }

    // 3. تسجيل الركوب
    var boardingId = generateBoardingId_();
    var now = formatDateTime_(new Date());

    var newRow = [
      boardingId,
      tripId,
      pilgrim.bookingId,
      pilgrim.passport,
      pilgrim.name,
      pilgrim.groupNumber,
      pilgrim.packageId,
      now,
      scanBy || '',
      scanMethod || 'qr',
      'undeclared',
      'boarded',
      ''
    ];

    boardingSheet.appendRow(newRow);

    // 4. تحديث عدد الركاب في الرحلة
    var tripSheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRIPS);
    var tripData = tripSheet.getDataRange().getValues();
    var TC = TRANSPORT_CONFIG.TRIP_COLS;
    for (var i = 1; i < tripData.length; i++) {
      if (String(tripData[i][TC.TRIP_ID]).trim() === tripId) {
        var currentBoarded = Number(tripData[i][TC.BOARDED_COUNT]) || 0;
        tripSheet.getRange(i + 1, TC.BOARDED_COUNT + 1).setValue(currentBoarded + 1);
        tripSheet.getRange(i + 1, TC.UPDATED_AT + 1).setValue(now);
        break;
      }
    }

    clearCache_('trips_');

    return {
      success: true,
      boardingId: boardingId,
      pilgrim: pilgrim,
      message: 'تم تسجيل ركوب: ' + pilgrim.name
    };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTION: الإبلاغ عن حادث
// ============================================================

function reportIncident(incidentData) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'النظام مشغول' };
  }

  try {
    var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
    var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.INCIDENTS);
    if (!sheet) return { success: false, error: 'شيت الحوادث غير موجود' };

    var incidentId = generateIncidentId_();
    var now = formatDateTime_(new Date());

    var newRow = [
      incidentId,
      incidentData.tripId || '',
      incidentData.type || 'other',
      incidentData.severity || 'medium',
      incidentData.description || '',
      incidentData.affectedPilgrims || 0,
      '', // resolution
      '', // resolvedBy
      incidentData.reportedBy || '',
      now,
      '', // resolvedAt
      'open'
    ];

    sheet.appendRow(newRow);

    return { success: true, incidentId: incidentId };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTION: حل حادث
// ============================================================

function resolveIncident(incidentId, resolution, resolvedBy) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'النظام مشغول' };
  }

  try {
    var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
    var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.INCIDENTS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === incidentId) {
        var row = i + 1;
        sheet.getRange(row, 7).setValue(resolution || '');
        sheet.getRange(row, 8).setValue(resolvedBy || '');
        sheet.getRange(row, 11).setValue(formatDateTime_(new Date()));
        sheet.getRange(row, 12).setValue('resolved');
        return { success: true };
      }
    }

    return { success: false, error: 'الحادث غير موجود' };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTION: تحديث رحلة
// ============================================================

function updateTrip(tripId, updates) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, error: 'النظام مشغول' };
  }

  try {
    var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
    var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRIPS);
    var data = sheet.getDataRange().getValues();
    var TC = TRANSPORT_CONFIG.TRIP_COLS;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][TC.TRIP_ID]).trim() === tripId) {
        var row = i + 1;

        if (updates.scheduledTime !== undefined) sheet.getRange(row, TC.SCHEDULED_TIME + 1).setValue(updates.scheduledTime);
        if (updates.vehiclePlate !== undefined) sheet.getRange(row, TC.VEHICLE_PLATE + 1).setValue(updates.vehiclePlate);
        if (updates.driverName !== undefined) sheet.getRange(row, TC.DRIVER_NAME + 1).setValue(updates.driverName);
        if (updates.driverPhone !== undefined) sheet.getRange(row, TC.DRIVER_PHONE + 1).setValue(updates.driverPhone);
        if (updates.capacity !== undefined) sheet.getRange(row, TC.CAPACITY + 1).setValue(updates.capacity);
        if (updates.assignedCount !== undefined) sheet.getRange(row, TC.ASSIGNED_COUNT + 1).setValue(updates.assignedCount);
        if (updates.notes !== undefined) sheet.getRange(row, TC.NOTES + 1).setValue(updates.notes);
        if (updates.status !== undefined) sheet.getRange(row, TC.STATUS + 1).setValue(updates.status);

        sheet.getRange(row, TC.UPDATED_AT + 1).setValue(formatDateTime_(new Date()));
        clearCache_('trips_');

        return { success: true };
      }
    }

    return { success: false, error: 'الرحلة غير موجودة' };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
