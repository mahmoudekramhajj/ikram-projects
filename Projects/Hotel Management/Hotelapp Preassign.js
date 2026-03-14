/**
 * Hotel Management App — Pre-Assignment Layer v2
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v3.6 — يستهلك غرف من Room Mapping + كتابة شركاء الغرفة
 */

// ── TRIGGER ──

function setupPreAssignTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runPreAssignment') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runPreAssignment').timeBased().everyHours(6).create();
  Logger.log('✅ Trigger: runPreAssignment كل 6 ساعات');
}

function clearPreAssignTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runPreAssignment') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('تم حذف ' + removed + ' trigger(s)');
}

// ── findSheet_ مع trim ──

function findSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim() === name.trim()) return sheets[i];
  }
  return null;
}

// ══════════════════════════════════════════════════════════════
// MAIN: تشغيل التعيين المبكر
// ══════════════════════════════════════════════════════════════

function runPreAssignment() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⚠️ Pre-Assignment: لم يتم الحصول على القفل — تخطي');
    return { error: 'Lock timeout' };
  }

  try {
  var startTime = new Date();
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var R = HOTEL_CONFIG.ROOM_COLS;
  
  // ── 1. قراءة البيانات المصدرية ──
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var roomSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_TYPE);
  
  if (!journeySheet || !roomSheet) {
    Logger.log('❌ شيت مصدري مفقود');
    return { error: 'Missing source sheets' };
  }
  
  var journeyData = journeySheet.getDataRange().getValues();
  var roomData = roomSheet.getDataRange().getValues();
  var guideMap = buildGuideMap_(ss);
  
  // ── 2. قراءة Room Mapping كاملاً ──
  var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
  var mapData = mapSheet ? mapSheet.getDataRange().getValues() : [];
  var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;
  
  // فهرس الغرف المتاحة: hotelName|roomType → [{rowIndex, roomId}, ...]
  var availableIndex = {};
  for (var i = 1; i < mapData.length; i++) {
    var status = String(mapData[i][MC.STATUS] || 'Available');
    if (status !== 'Available') continue;
    var hn = String(mapData[i][MC.HOTEL_NAME]);
    var rt = String(mapData[i][MC.ROOM_TYPE]);
    var rid = String(mapData[i][MC.INTERNAL_ID]);
    var key = hn + '|' + rt;
    if (!availableIndex[key]) availableIndex[key] = [];
    availableIndex[key].push({ rowIndex: i + 1, roomId: rid });
  }
  
  // ── 3. بناء معلومات المجموعات ──
  var groupInfo = {};
  for (var i = 1; i < roomData.length; i++) {
    var rRow = roomData[i];
    var gn = rRow[R.GROUP_NUMBER];
    if (!gn) continue;
    groupInfo[gn] = {
      roomTypeMed: rRow[R.ROOM_TYPE_MED],
      fullRoomsMed: Number(rRow[R.FULL_ROOMS_MED]) || 0,
      roomTypeMak1: rRow[R.ROOM_TYPE_MAK1],
      fullRoomsMak1: Number(rRow[R.FULL_ROOMS_MAK1]) || 0,
      roomTypeMak2: rRow[R.ROOM_TYPE_MAK2],
      fullRoomsMak2: Number(rRow[R.FULL_ROOMS_MAK2]) || 0,
      tourGuide: rRow[R.TOUR_GUIDE] || ''
    };
  }
  
  // ── 4. بناء خريطة الحجاج حسب الفندق ──
  var hotelMap = {};
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var bid = row[J.BOOKING_ID];
    var gn = row[J.GROUP_NUMBER];
    if (!bid || !gn) continue;
    
    var pilgrim = {
      bookingId: String(bid),
      name: String(row[J.NAME] || ''),
      passport: String(row[J.PASSPORT] || '').trim().toUpperCase(),
      gender: row[J.GENDER],
      groupNumber: gn
    };
    
    var medH = row[J.MADINAH_EN];
    if (medH && medH !== 'NULL' && String(medH).trim()) addToHotelMap_(hotelMap, medH, 'Madina', gn, pilgrim);
    var makH = row[J.MAKKAH_EN];
    if (makH && makH !== 'NULL' && String(makH).trim()) addToHotelMap_(hotelMap, makH, 'Makkah', gn, pilgrim);
    var shiftH = row[J.MAKKAH_SHIFT_EN];
    if (shiftH && shiftH !== 'NULL' && String(shiftH).trim()) addToHotelMap_(hotelMap, shiftH, 'Makkah Shifting', gn, pilgrim);
  }
  
  // ── 5. بناء التعيينات المتوقعة ──
  var expected = buildExpectedFullRooms_(hotelMap, groupInfo, guideMap);
  
  // ── 6. قراءة PA_Log ──
  var logSheet = getOrCreatePreAssignLog_(ss);
  var active = readActivePreAssignments_(logSheet);
  
  // ── 7. المقارنة ──
  var toAdd = [], toRemove = [];
  var stats = { assigned: 0, removed: 0, unchanged: 0, skipped: 0, noRoom: 0 };
  var processedKeys = {};
  
  for (var key in expected) {
    processedKeys[key] = true;
    var exp = expected[key];
    var act = active[key];
    if (!act) { toAdd.push(exp); }
    else if (act.dataHash === exp.dataHash) { stats.unchanged++; }
    else { toRemove.push(act); toAdd.push(exp); }
  }
  for (var key in active) {
    if (!processedKeys[key]) toRemove.push(active[key]);
  }
  
  // ── 8. إزالة أولاً ──
  for (var r = 0; r < toRemove.length; r++) {
    executeRemoval_(ss, logSheet, mapSheet, toRemove[r]);
    stats.removed++;
  }
  
  // ── 9. التعيينات الجديدة ──
  var addByHotel = {};
  for (var a = 0; a < toAdd.length; a++) {
    var entry = toAdd[a];
    if (!addByHotel[entry.hotelName]) addByHotel[entry.hotelName] = [];
    addByHotel[entry.hotelName].push(entry);
  }
  
  var mapUpdates = [];
  
  for (var hotelName in addByHotel) {
    var result = executeAssignments_(ss, logSheet, hotelName, addByHotel[hotelName], availableIndex, mapUpdates);
    stats.assigned += result.assigned;
    stats.noRoom += result.noRoom;
    stats.skipped += result.skipped;
  }
  
  // ── 10. كتابة تحديثات Room Mapping بالجملة ──
  if (mapSheet && mapUpdates.length > 0) {
    for (var u = 0; u < mapUpdates.length; u++) {
      var upd = mapUpdates[u];
      mapSheet.getRange(upd.rowIndex, MC.STATUS + 1).setValue('Occupied');
      mapSheet.getRange(upd.rowIndex, MC.OCCUPANT_IDS + 1).setValue(upd.occupantIds);
    }
  }
  
  // ── 11. النتائج ──
  SpreadsheetApp.flush();
  var elapsed = ((new Date()) - startTime) / 1000;
  var summary = {
    success: true, assigned: stats.assigned, removed: stats.removed,
    unchanged: stats.unchanged, noRoom: stats.noRoom, elapsed: elapsed.toFixed(1) + 's'
  };
  Logger.log('Pre-Assignment Summary: ' + JSON.stringify(summary));
  return summary;
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════
// EXPECTED: المجموعات الكاملة فقط
// ══════════════════════════════════════════════════════════════

function buildExpectedFullRooms_(hotelMap, groupInfo, guideMap) {
  var expected = {};
  for (var hotelKey in hotelMap) {
    var parts = hotelKey.split('|');
    var hotelName = parts[0], hotelCity = parts[1];
    var groups = hotelMap[hotelKey];
    
    for (var gn in groups) {
      var gi = groupInfo[gn];
      if (!gi) continue;
      var roomType = getRoomType_(gi, hotelCity);
      var fullRooms = getFullRooms_(gi, hotelCity);
      if (fullRooms <= 0 || !roomType) continue;
      
      var members = groups[gn];
      var capacity = getRoomCapacity_(roomType);
      
      var females = members.filter(function(p) { return p.gender === 'Female'; });
      var males = members.filter(function(p) { return p.gender === 'Male'; });
      var sorted = females.concat(males);
      
      var maxForFull = fullRooms * capacity;
      var forFull = sorted.slice(0, Math.min(maxForFull, members.length));
      if (forFull.length < capacity) continue;
      
      var actualFullRooms = Math.floor(forFull.length / capacity);
      forFull = forFull.slice(0, actualFullRooms * capacity);
      
      var bids = forFull.map(function(p) { return p.bookingId; }).sort().join(',');
      var dataHash = bids + '|' + roomType;
      var assignKey = gn + '|' + hotelName + '|' + hotelCity;
      
      expected[assignKey] = {
        assignKey: assignKey, groupNumber: gn,
        hotelName: hotelName, hotelCity: hotelCity,
        roomType: roomType, capacity: capacity,
        fullRooms: actualFullRooms, pilgrims: forFull,
        dataHash: dataHash
      };
    }
  }
  return expected;
}

// ══════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════

function addToHotelMap_(map, hotelName, hotelCity, groupNumber, pilgrim) {
  var key = hotelName + '|' + hotelCity;
  if (!map[key]) map[key] = {};
  if (!map[key][groupNumber]) map[key][groupNumber] = [];
  map[key][groupNumber].push(pilgrim);
}

function isPreAssignGuideGroup_(members, groupInfo, guideMap) {
  if (groupInfo.tourGuide && String(groupInfo.tourGuide).trim()) return true;
  for (var p = 0; p < members.length; p++) {
    if (members[p].passport && getGuideForPilgrim_(guideMap, members[p].passport)) return true;
  }
  return false;
}

// ══════════════════════════════════════════════════════════════
// LOG: إنشاء وقراءة PA_Log
// ══════════════════════════════════════════════════════════════

function getOrCreatePreAssignLog_(ss) {
  var sheetName = HOTEL_CONFIG.SHEETS.PRE_ASSIGN_LOG || 'PA_Log';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['RoomId','GroupNumber','HotelName','HotelCity','RoomType','Capacity','PilgrimCount','BookingIds','DataHash','Status','CreatedAt','UpdatedAt']);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readActivePreAssignments_(logSheet) {
  var data = logSheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][9]) !== 'ACTIVE') continue;
    var key = String(data[i][1]) + '|' + String(data[i][2]) + '|' + String(data[i][3]);
    if (!map[key]) {
      map[key] = { assignKey: key, roomIds: [], dataHash: String(data[i][8]), logRows: [], hotelName: String(data[i][2]) };
    }
    map[key].roomIds.push(String(data[i][0]));
    map[key].logRows.push(i + 1);
  }
  return map;
}

// ══════════════════════════════════════════════════════════════
// REMOVAL: إزالة تعيين + إعادة Room Mapping إلى Available
// ══════════════════════════════════════════════════════════════

function executeRemoval_(ss, logSheet, mapSheet, entry) {
  var hotelSheet = getHotelSheet_(ss, entry.hotelName);
  if (hotelSheet) clearPreAssignedRooms_(hotelSheet, entry.roomIds);
  
  if (mapSheet) {
    var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;
    var mapData = mapSheet.getDataRange().getValues();
    var idSet = {};
    for (var r = 0; r < entry.roomIds.length; r++) idSet[entry.roomIds[r]] = true;
    for (var i = 1; i < mapData.length; i++) {
      if (idSet[String(mapData[i][MC.INTERNAL_ID])]) {
        mapSheet.getRange(i + 1, MC.STATUS + 1).setValue('Available');
        mapSheet.getRange(i + 1, MC.OCCUPANT_IDS + 1).setValue('');
      }
    }
  }
  
  var now = new Date();
  for (var r = 0; r < entry.logRows.length; r++) {
    logSheet.getRange(entry.logRows[r], 10).setValue('REMOVED');
    logSheet.getRange(entry.logRows[r], 12).setValue(now);
  }
}

function clearPreAssignedRooms_(sheet, roomIds) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  var numRows = lastRow - 1;
  var rgCol = sheet.getRange(2, 17, numRows, 1).getValues();
  var rnCol = sheet.getRange(2, 18, numRows, 1).getValues();
  var stCol = sheet.getRange(2, 19, numRows, 1).getValues();
  
  var idSet = {};
  for (var r = 0; r < roomIds.length; r++) idSet[roomIds[r]] = true;
  
  var changed = false;
  for (var i = 0; i < rgCol.length; i++) {
    if (idSet[String(rgCol[i][0] || '')] && String(stCol[i][0] || '') !== 'arrived') {
      rgCol[i][0] = '';
      rnCol[i][0] = '';
      changed = true;
    }
  }
  if (changed) {
    sheet.getRange(2, 17, numRows, 1).setValues(rgCol);
    sheet.getRange(2, 18, numRows, 1).setValues(rnCol);
  }
}

// ══════════════════════════════════════════════════════════════
// EXECUTE: تعيين جديد — يستهلك من Room Mapping
// ══════════════════════════════════════════════════════════════

function executeAssignments_(ss, logSheet, hotelName, entries, availableIndex, mapUpdates) {
  if (!entries || entries.length === 0) return { assigned: 0, noRoom: 0, skipped: 0 };
  
  var hotelSheet = getOrCreateHotelSheet_(ss, hotelName);
  var sheetData = hotelSheet.getDataRange().getValues();
  
  if (sheetData.length <= 1) {
    var populateCity = (entries[0].hotelCity === 'Makkah Shifting') ? 'Makkah' : entries[0].hotelCity;
    populateHotelSheet(hotelName, populateCity);
  }
  
  var checkinData = getCheckinData_(ss, hotelName);
  
  var allAssignments = {};
  var roomOccupants = {};
  var logRows = [];
  var assigned = 0, noRoom = 0, skipped = 0;
  var now = new Date();
  
  for (var e = 0; e < entries.length; e++) {
    var entry = entries[e];
    
    var blocked = false;
    for (var p = 0; p < entry.pilgrims.length; p++) {
      var ci = checkinData[entry.pilgrims[p].bookingId];
      if (ci && (ci.status === 'arrived' || (ci.roomGroupId && String(ci.roomGroupId).trim()))) {
        blocked = true; break;
      }
    }
    if (blocked) { skipped++; continue; }
    
    var tempMembers = entry.pilgrims.slice();
    var poolKey = hotelName + '|' + entry.roomType;
    var pool = availableIndex[poolKey] || [];
    
    for (var r = 0; r < entry.fullRooms && tempMembers.length >= entry.capacity; r++) {
      if (pool.length === 0) {
        Logger.log('⚠️ لا غرف متاحة: ' + hotelName + ' ' + entry.roomType + ' — مجموعة ' + entry.groupNumber);
        noRoom++;
        break;
      }
      
      var picked = pool.shift();
      var roomId = picked.roomId;
      var roomMembers = tempMembers.splice(0, entry.capacity);
      
      var roomBids = [];
      for (var p = 0; p < roomMembers.length; p++) {
        allAssignments[roomMembers[p].bookingId] = roomId;
        roomBids.push(roomMembers[p].bookingId);
      }
      
      roomOccupants[roomId] = roomBids;
      
      mapUpdates.push({
        rowIndex: picked.rowIndex,
        occupantIds: roomBids.join(' | ')
      });
      
      logRows.push([
        roomId, entry.groupNumber, hotelName, entry.hotelCity,
        entry.roomType, entry.capacity, roomMembers.length,
        roomBids.join(' | '), entry.dataHash, 'ACTIVE', now, now
      ]);
      
      assigned++;
    }
  }
  
  if (Object.keys(allAssignments).length > 0) {
    batchWriteRoomAssignments_(hotelSheet, allAssignments);
  }
  
  writeRoommates_(hotelSheet, roomOccupants);
  
  if (logRows.length > 0) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }
  
  return { assigned: assigned, noRoom: noRoom, skipped: skipped };
}

// ══════════════════════════════════════════════════════════════
// شركاء الغرفة: كتابة عمود 24 في شيت الفندق
// ══════════════════════════════════════════════════════════════

function writeRoommates_(hotelSheet, roomOccupants) {
  var lastRow = hotelSheet.getLastRow();
  if (lastRow <= 1) return;
  var numRows = lastRow - 1;
  
  var bidCol = hotelSheet.getRange(2, 1, numRows, 1).getValues();
  var rgCol = hotelSheet.getRange(2, 17, numRows, 1).getValues();
  var nameCol = hotelSheet.getRange(2, 2, numRows, 1).getValues();
  
  // التأكد من وجود العنوان
  var header24 = hotelSheet.getRange(1, 24).getValue();
  if (!header24 || header24 !== 'شركاء الغرفة') {
    hotelSheet.getRange(1, 24).setValue('شركاء الغرفة').setFontWeight('bold');
  }
  
  // خريطة bid → name
  var bidToName = {};
  for (var i = 0; i < bidCol.length; i++) {
    bidToName[String(bidCol[i][0])] = String(nameCol[i][0] || '');
  }
  
  // بناء عمود الشركاء
  var roommateCol = [];
  for (var i = 0; i < rgCol.length; i++) {
    var roomId = String(rgCol[i][0] || '');
    var myBid = String(bidCol[i][0]);
    
    if (roomId && roomOccupants[roomId]) {
      var mates = roomOccupants[roomId].filter(function(b) { return b !== myBid; });
      var mateNames = mates.map(function(b) { return bidToName[b] || b; });
      roommateCol.push([mateNames.join(' | ')]);
    } else {
      roommateCol.push(['']);
    }
  }
  
  hotelSheet.getRange(2, 24, numRows, 1).setValues(roommateCol);
}

// ══════════════════════════════════════════════════════════════
// RESET: إعادة ضبط كاملة — يُشغّل مرة واحدة
// ══════════════════════════════════════════════════════════════

function resetAndRebuild() {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  Logger.log('═══ بدء إعادة الضبط ═══');
  
  // ── 1. مسح PA_Log ──
  var logSheet = ss.getSheetByName('PA_Log');
  if (logSheet) {
    logSheet.clear();
    logSheet.appendRow(['RoomId','GroupNumber','HotelName','HotelCity','RoomType','Capacity','PilgrimCount','BookingIds','DataHash','Status','CreatedAt','UpdatedAt']);
    logSheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    logSheet.setFrozenRows(1);
    Logger.log('✅ PA_Log — مُسح');
  }
  
  // ── 2. مسح PA_Tracker ──
  var trackerSheet = ss.getSheetByName('PA_Tracker');
  if (trackerSheet) { ss.deleteSheet(trackerSheet); Logger.log('✅ PA_Tracker — محذوف'); }
  
  // ── 3. مسح RoomGroup_ID من كل شيت فندق ──
  var abbr = HOTEL_CONFIG.HOTEL_ABBR;
  for (var fullName in abbr) {
    var sheetName = sanitizeSheetName_(fullName);
    var sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      var numRows = sheet.getLastRow() - 1;
      var emptyRG = Array(numRows).fill(['']);
      sheet.getRange(2, 17, numRows, 1).setValues(emptyRG);
      sheet.getRange(2, 18, numRows, 1).setValues(emptyRG);
      try { sheet.getRange(2, 24, numRows, 1).setValues(emptyRG); } catch(e) {}
      Logger.log('  🗑 ' + sheetName + ': مسح التعيينات');
    }
  }
  
  // ── 4. توليد Room Mapping لكل الفنادق ──
  Logger.log('═══ توليد Room Mapping ═══');
  var rmResult = generateAllRoomMappings();
  Logger.log('✅ Room Mapping: ' + rmResult.total + ' غرفة عبر ' + rmResult.hotels + ' فندق');
  
  // ── 5. تشغيل Pre-Assignment ──
  Logger.log('═══ تشغيل Pre-Assignment ═══');
  var paResult = runPreAssignment();
  Logger.log('✅ Pre-Assignment: ' + JSON.stringify(paResult));
  
  Logger.log('═══ اكتمل ═══');
  return { roomMapping: rmResult, preAssignment: paResult };
}