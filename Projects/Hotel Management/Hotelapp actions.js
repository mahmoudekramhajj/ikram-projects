/**
 * Hotel Management App — Actions Layer
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v4.0 — CheckIn + Manual RoomGroup + RoomMapping
 * (AutoAssign moved to SmartAssign.js)
 */

// ============================================================
// ACTIONS: PERFORM CHECK-IN
// ============================================================

function performCheckIn(bookingId, hotelName, hotelCity, roomNumber) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'النظام مشغول — حاول مرة أخرى' };

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateHotelSheet_(ss, hotelName);
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      populateHotelSheet(hotelName, hotelCity);
      data = sheet.getDataRange().getValues();
    }

    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingId)) {
        sheet.getRange(i + 1, 18, 1, 3).setValues([[roomNumber, 'arrived', timestamp]]);
        SpreadsheetApp.flush();
        logAudit_('check-in', 'Booking: ' + bookingId + (roomNumber ? ' Room: ' + roomNumber : ''), hotelName);
        return { success: true, timestamp: timestamp };
      }
    }

    return { success: false, error: 'Booking not found in hotel sheet' };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTIONS: BULK CHECK-IN
// ============================================================

function bulkCheckIn(bookingIds, hotelName, hotelCity, roomNumbers) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return [{ success: false, error: 'النظام مشغول — حاول مرة أخرى' }];

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateHotelSheet_(ss, hotelName);
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      populateHotelSheet(hotelName, hotelCity);
      data = sheet.getDataRange().getValues();
    }

    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');

    var idSet = {};
    for (var b = 0; b < bookingIds.length; b++) {
      idSet[String(bookingIds[b])] = {
        room: (roomNumbers && roomNumbers[b]) ? roomNumbers[b] : '',
        found: false
      };
    }

    for (var i = 1; i < data.length; i++) {
      var bid = String(data[i][0]);
      if (idSet[bid] && !idSet[bid].found) {
        sheet.getRange(i + 1, 18, 1, 3).setValues([[idSet[bid].room, 'arrived', timestamp]]);
        idSet[bid].found = true;
      }
    }

    SpreadsheetApp.flush();
    logAudit_('bulk-check-in', bookingIds.length + ' حاج', hotelName);

    var results = [];
    for (var b = 0; b < bookingIds.length; b++) {
      var entry = idSet[String(bookingIds[b])];
      results.push(entry && entry.found
        ? { success: true, timestamp: timestamp }
        : { success: false, error: 'Booking not found in hotel sheet' }
      );
    }
    return results;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// ACTIONS: SAVE ROOM GROUP (Manual)
// ============================================================

function saveRoomGroup(bookingIds, hotelName, hotelCity, roomGroupId) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = getOrCreateHotelSheet_(ss, hotelName);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    populateHotelSheet(hotelName, hotelCity);
    data = sheet.getDataRange().getValues();
  }

  var gid = roomGroupId;
  if (!gid) {
    var roomType = 'Quad';
    for (var b = 0; b < bookingIds.length; b++) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(bookingIds[b])) {
          roomType = data[i][15] || 'Quad';
          break;
        }
      }
      if (roomType) break;
    }

    var existingIds = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][16]) existingIds.push(String(data[i][16]));
    }

    gid = generateInternalRoomId_(hotelName, hotelCity, roomType, existingIds);
  }

  var count = 0;
  for (var b = 0; b < bookingIds.length; b++) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(bookingIds[b])) {
        sheet.getRange(i + 1, 17).setValue(gid);
        count++;
        break;
      }
    }
  }

  SpreadsheetApp.flush();
  return { success: true, groupId: gid, count: count };
}

// ============================================================
// ACTIONS: UNASSIGN ROOM GROUP
// ============================================================

function unassignRoomGroup(roomGroupId, hotelName) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = getHotelSheet_(ss, hotelName);
  if (!sheet) return { success: false, error: 'Hotel sheet not found' };

  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][16] === roomGroupId) {
      sheet.getRange(i + 1, 17).setValue('');
      count++;
    }
  }

  if (count > 0) SpreadsheetApp.flush();
  return { success: true, removed: count };
}

// ============================================================
// ACTIONS: MOVE PILGRIM TO DIFFERENT ROOM GROUP
// ============================================================

function moveToRoomGroup(bookingId, hotelName, newRoomGroupId) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = getHotelSheet_(ss, hotelName);
  if (!sheet) return { success: false, error: 'Hotel sheet not found' };

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(bookingId)) {
      sheet.getRange(i + 1, 17).setValue(newRoomGroupId);
      return { success: true };
    }
  }

  return { success: false, error: 'Pilgrim not found' };
}

// ============================================================
// ROOM MAPPING: حفظ رقم الغرفة الفعلي
// ============================================================

function saveActualRoomNo(internalRoomId, actualRoomNo) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  if (!mapSheet) return { success: false, error: 'شيت Room Mapping غير موجود' };

  var data = mapSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(internalRoomId)) {
      mapSheet.getRange(i + 1, 8).setValue(actualRoomNo);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }

  return { success: false, error: 'رقم الغرفة الداخلي غير موجود: ' + internalRoomId };
}
