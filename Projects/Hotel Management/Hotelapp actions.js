/**
 * Hotel Management App — Actions Layer
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v3.4 — CheckIn + AutoAssign 4-phase + Optimization + RoomMapping + diagnose
 */

// ============================================================
// ACTIONS: PERFORM CHECK-IN
// ============================================================

function performCheckIn(bookingId, hotelName, hotelCity, roomNumber) {
  // قفل لمنع التعارض عند تسجيل دخول متزامن
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, error: 'النظام مشغول — حاول مرة أخرى' };

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateHotelSheet_(ss, hotelName);
    var data = sheet.getDataRange().getValues();

    // تعبئة تلقائية إذا الشيت فارغ
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
  
  // Auto-populate if sheet is empty
  if (data.length <= 1) {
    populateHotelSheet(hotelName, hotelCity);
    data = sheet.getDataRange().getValues();
  }
  
  // Generate internal room ID if not provided
  var gid = roomGroupId;
  if (!gid) {
    // Get room type from first pilgrim's data
    var roomType = 'Quad';
    for (var b = 0; b < bookingIds.length; b++) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(bookingIds[b])) {
          roomType = data[i][15] || 'Quad'; // col P = room type
          break;
        }
      }
      if (roomType) break;
    }
    
    // Collect existing IDs for sequence
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
        sheet.getRange(i + 1, 17).setValue(gid); // RoomGroup_ID (col Q = 17)
        count++;
        break;
      }
    }
  }
  
  SpreadsheetApp.flush();
  return { success: true, groupId: gid, count: count };
}

// ============================================================
// ACTIONS: AUTO ASSIGN ROOMS (v3.1 — Fixed)
// Phase 0: ALL groups with fullRooms > 0 → assign full rooms
//          (remaining pilgrims stay for Phase 2/3)
// Phase 1: Guide groups → mark as "awaiting guide" (manual)
// Phase 2: Same group shared beds → fill rooms (mixed gender OK)
// Phase 3: Cross-group shared beds → same gender only
// ============================================================

function autoAssignRooms(hotelName, hotelCity) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, error: 'النظام مشغول — حاول مرة أخرى' };

  try {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  
  var journeySheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var roomSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_TYPE);
  var roomData = roomSheet.getDataRange().getValues();
  var checkinData = getCheckinData_(ss, hotelName);
  
  // Build guide map from Tour Guide sheet (passport → guideName)
  var guideMap = buildGuideMap_(ss);
  
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var R = HOTEL_CONFIG.ROOM_COLS;
  
  // 1. Build group info from Room Type Preview
  var groupInfo = {};
  for (var i = 1; i < roomData.length; i++) {
    var rRow = roomData[i];
    var gn = rRow[R.GROUP_NUMBER];
    if (gn) {
      groupInfo[gn] = {
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
        sharedBedsMak2: Number(rRow[R.SHARED_BEDS_MAK2]) || 0,
        tourGuide: rRow[R.TOUR_GUIDE] || ''
      };
    }
  }
  
  // 2. Get unassigned pilgrims for this hotel
  var groupPilgrims = {};
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var actualCity = getActualHotelCity_(row, hotelName, hotelCity);
    if (!actualCity) continue;
    
    // Skip No Show (Madinah only)
    if (actualCity === 'Madina') {
      var hotelDates = getHotelDates_(row, actualCity);
      if (hotelDates.phase === 'first') {
        var arrDate = row[J.ARRIVAL_DATE];
        var contractEnd = row[J.FIRST_HOUSE_END];
        if (arrDate && contractEnd) {
          var ad = new Date(arrDate); ad.setHours(0,0,0,0);
          var ce = new Date(contractEnd); ce.setHours(0,0,0,0);
          if (ad > ce) continue;
        }
      }
    }
    
    var bookingId = row[J.BOOKING_ID];
    var checkin = checkinData[bookingId];
    if (checkin && checkin.roomGroupId) continue; // Already assigned
    
    var gn = row[J.GROUP_NUMBER];
    if (!groupPilgrims[gn]) groupPilgrims[gn] = [];
    groupPilgrims[gn].push({
      bookingId: bookingId,
      name: row[J.NAME],
      passport: row[J.PASSPORT],
      gender: row[J.GENDER],
      groupNumber: gn,
      nationality: row[J.NATIONALITY_EN],
      countryResidence: row[J.COUNTRY_RESIDENCE] || '',
      arrivalDate: formatDate_(row[J.ARRIVAL_DATE])
    });
  }

  // 3. Ensure hotel sheet is populated
  var sheet = getOrCreateHotelSheet_(ss, hotelName);
  var sheetData = sheet.getDataRange().getValues();
  if (sheetData.length <= 1) {
    populateHotelSheet(hotelName, hotelCity);
    sheet = ss.getSheetByName(sanitizeSheetName_(hotelName));
  }
  
  // Collect existing room IDs for sequence tracking
  var existingIds = [];
  var existingData = sheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    var rid = existingData[i][16]; // RoomGroup_ID column Q
    if (rid) existingIds.push(String(rid));
  }
  
  // 4. Execute 4 phases
  var assignmentMap = {}; // bookingId → roomId
  var assignments = []; // room info for result
  var guideGroups = []; // groups awaiting guide
  var phase0Count = 0;
  var phase1Count = 0;
  
  // ═══════════════════════════════════════════════════════════
  // PHASE 0: Groups with full rooms → assign complete rooms
  // *** FIX v3.1: Process ALL groups with fullRooms > 0,
  //     even if they ALSO have sharedBeds > 0.
  //     Remaining pilgrims stay in groupPilgrims for Phase 2/3.
  // ═══════════════════════════════════════════════════════════
  for (var gn in groupPilgrims) {
    var gi = groupInfo[gn];
    if (!gi) continue;
    
    var roomType = getRoomType_(gi, hotelCity);
    var fullRooms = getFullRooms_(gi, hotelCity);
    var sharedBeds = getSharedBeds_(gi, hotelCity);
    
    if (fullRooms <= 0) continue; // Skip groups with no full rooms
    
    var members = groupPilgrims[gn];
    var capacity = getRoomCapacity_(roomType);
    var fullRoomPilgrims = fullRooms * capacity;
    
    // *** FIX v3.1: Select WHO goes to full rooms vs shared beds
    // Females first to full rooms → males left for shared beds
    var females = members.filter(function(p) { return p.gender === 'Female'; });
    var males = members.filter(function(p) { return p.gender === 'Male'; });
    var sortedForFullRooms = females.concat(males);
    
    // Take first N for full rooms, leave rest for shared pool
    var forFullRooms = sortedForFullRooms.slice(0, Math.min(fullRoomPilgrims, members.length));
    var forSharedPool = sortedForFullRooms.slice(forFullRooms.length);
    
    // Assign full rooms
    var tempMembers = forFullRooms.slice();
    for (var r = 0; r < fullRooms && tempMembers.length > 0; r++) {
      var roomMembers = tempMembers.splice(0, capacity);
      if (roomMembers.length === 0) break;
      
      var roomId = generateInternalRoomId_(hotelName, hotelCity, roomType, existingIds);
      existingIds.push(roomId);
      
      for (var p = 0; p < roomMembers.length; p++) {
        assignmentMap[String(roomMembers[p].bookingId)] = roomId;
      }
      
      assignments.push({
        groupId: roomId,
        count: roomMembers.length,
        capacity: capacity,
        incomplete: roomMembers.length < capacity,
        phase: 0,
        groupNumber: gn
      });
      phase0Count++;
    }
    
    // Update groupPilgrims: keep only shared-pool pilgrims
    if (forSharedPool.length > 0) {
      groupPilgrims[gn] = forSharedPool;
    } else {
      delete groupPilgrims[gn]; // Fully assigned
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // PHASE 1: Guide groups — skip for manual assignment
  // *** FIX v3.1: Read guide from Tour Guide sheet by passport
  //     (Room Type Preview column K is empty)
  // ═══════════════════════════════════════════════════════════
  for (var gn in groupPilgrims) {
    var pilgrims = groupPilgrims[gn];
    
    // Check if any pilgrim in this group has a guide
    var groupGuide = '';
    for (var p = 0; p < pilgrims.length; p++) {
      var guide = getGuideForPilgrim_(guideMap, pilgrims[p].passport);
      if (guide) { groupGuide = guide; break; }
    }
    
    // Also check Room Type Preview as fallback
    var gi = groupInfo[gn];
    if (!groupGuide && gi && gi.tourGuide && String(gi.tourGuide).trim() !== '') {
      groupGuide = gi.tourGuide;
    }
    
    if (groupGuide) {
      guideGroups.push({
        groupNumber: gn,
        tourGuide: groupGuide,
        count: pilgrims.length
      });
      phase1Count += pilgrims.length;
      delete groupPilgrims[gn];
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // PHASE 2 + 3: Shared beds grouping
  // ═══════════════════════════════════════════════════════════
  var sharedPool = {}; // key = capacity, value = array of pilgrims
  
  for (var gn in groupPilgrims) {
    var gi = groupInfo[gn];
    if (!gi) continue;
    
    var roomType = getRoomType_(gi, hotelCity);
    var capacity = getRoomCapacity_(roomType);
    var typeKey = String(capacity);
    
    // All remaining pilgrims in groupPilgrims are shared-bed pilgrims
    var pilgrims = groupPilgrims[gn];
    if (pilgrims.length === 0) continue;
    
    for (var p = 0; p < pilgrims.length; p++) {
      pilgrims[p].roomType = roomType;
      pilgrims[p].capacity = capacity;
      if (!sharedPool[typeKey]) sharedPool[typeKey] = [];
      sharedPool[typeKey].push(pilgrims[p]);
    }
  }
  
  var used = {};
  
  for (var typeKey in sharedPool) {
    var pool = sharedPool[typeKey];
    var capacity = parseInt(typeKey);
    var rooms = [];
    
    // ═══════════════════════════════════════════════════════
    // PHASE 2: Same group — mixed gender OK (relatives)
    // ═══════════════════════════════════════════════════════
    var byGroup = {};
    for (var i = 0; i < pool.length; i++) {
      var gn = pool[i].groupNumber;
      if (!byGroup[gn]) byGroup[gn] = [];
      byGroup[gn].push(pool[i]);
    }
    
    for (var gn in byGroup) {
      var remaining = byGroup[gn].slice();
      
      while (remaining.length >= capacity) {
        var fInRemaining = remaining.filter(function(p) { return p.gender === 'Female'; });
        var mInRemaining = remaining.filter(function(p) { return p.gender === 'Male'; });
        
        // Try single-gender rooms first
        if (fInRemaining.length >= capacity) {
          var room = fInRemaining.slice(0, capacity);
          rooms.push(room);
          room.forEach(function(p) { used[p.bookingId] = true; });
          remaining = remaining.filter(function(p) { return !used[p.bookingId]; });
        } else if (mInRemaining.length >= capacity) {
          var room = mInRemaining.slice(0, capacity);
          rooms.push(room);
          room.forEach(function(p) { used[p.bookingId] = true; });
          remaining = remaining.filter(function(p) { return !used[p.bookingId]; });
        } else {
          // Mixed gender allowed (same group = relatives)
          var room = remaining.slice(0, capacity);
          rooms.push(room);
          room.forEach(function(p) { used[p.bookingId] = true; });
          remaining = remaining.filter(function(p) { return !used[p.bookingId]; });
        }
      }
      
      // Remaining from this group (< capacity) → leave for Phase 3
    }
    
    // ═══════════════════════════════════════════════════════
    // PHASE 3: Cross-group — same gender ONLY
    // Sort by: arrival date → country → nationality
    // ═══════════════════════════════════════════════════════
    var remainingPool = pool.filter(function(p) { return !used[p.bookingId]; });
    
    ['Female', 'Male'].forEach(function(gender) {
      var genderPool = remainingPool.filter(function(p) { return p.gender === gender; });
      
      genderPool.sort(function(a, b) {
        if (a.arrivalDate !== b.arrivalDate) return (a.arrivalDate || '').localeCompare(b.arrivalDate || '');
        if (a.countryResidence !== b.countryResidence) return (a.countryResidence || '').localeCompare(b.countryResidence || '');
        return (a.nationality || '').localeCompare(b.nationality || '');
      });
      
      var room = [];
      for (var i = 0; i < genderPool.length; i++) {
        if (used[genderPool[i].bookingId]) continue;
        room.push(genderPool[i]);
        used[genderPool[i].bookingId] = true;
        if (room.length >= capacity) {
          rooms.push(room);
          room = [];
        }
      }
      if (room.length > 0) rooms.push(room); // Incomplete room
    });
    
    // === Save with internal room IDs ===
    for (var r = 0; r < rooms.length; r++) {
      var room = rooms[r];
      var roomType = room[0].roomType || 'Quad';
      var roomId = generateInternalRoomId_(hotelName, hotelCity, roomType, existingIds);
      existingIds.push(roomId);
      
      for (var p = 0; p < room.length; p++) {
        assignmentMap[String(room[p].bookingId)] = roomId;
      }
      
      // Determine phase
      var roomGroups = {};
      room.forEach(function(m) { roomGroups[m.groupNumber] = true; });
      var phase = Object.keys(roomGroups).length === 1 ? 2 : 3;
      
      assignments.push({
        groupId: roomId,
        count: room.length,
        capacity: capacity,
        incomplete: room.length < capacity,
        phase: phase
      });
    }
  }
  
  // === BATCH WRITE all assignments at once ===
  var written = 0;
  if (Object.keys(assignmentMap).length > 0) {
    written = batchWriteRoomAssignments_(sheet, assignmentMap);
  }
  
  var poolTotal = 0;
  for (var k in sharedPool) { poolTotal += sharedPool[k].length; }
  
  return { 
    success: true, 
    roomsCreated: assignments.length,
    totalAssigned: written,
    incompleteRooms: assignments.filter(function(a) { return a.incomplete; }).length,
    assignments: assignments,
    guideGroups: guideGroups,
    debug: {
      groupsInHotel: Object.keys(groupPilgrims).length + guideGroups.length + phase0Count,
      phase0_complete: phase0Count,
      phase1_guide: phase1Count,
      sharedPoolTotal: poolTotal,
      assignmentMapSize: Object.keys(assignmentMap).length
    }
  };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// OPTIMIZATION: Post-assignment room optimization
// Moves pilgrims from larger incomplete rooms to smaller ones
// Rule: upgrade only (Quad→Triple/Double OK, Double→Triple FORBIDDEN)
// Rule: never separate same-group members
// ============================================================

function getOptimizationSuggestions(hotelName, hotelCity) {
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var sheet = getHotelSheet_(ss, hotelName);
  if (!sheet) return { suggestions: [] };
  
  var data = sheet.getDataRange().getValues();
  
  // Build rooms from sheet data
  var rooms = {}; // roomId → { members: [], capacity, roomType }
  for (var i = 1; i < data.length; i++) {
    var roomId = String(data[i][16] || '');
    if (!roomId) continue;
    
    var roomType = data[i][15] || 'Quad';
    var capacity = getRoomCapacity_(roomType);
    
    if (!rooms[roomId]) rooms[roomId] = { members: [], capacity: capacity, roomType: roomType };
    rooms[roomId].members.push({
      bookingId: String(data[i][0]),
      name: data[i][1],
      gender: data[i][3],
      groupNumber: data[i][5]
    });
  }
  
  // Find incomplete rooms (count < capacity)
  var incompleteRooms = [];
  for (var id in rooms) {
    var room = rooms[id];
    room.id = id;
    room.count = room.members.length;
    if (room.count < room.capacity) {
      incompleteRooms.push(room);
    }
  }
  
  // Sort: smallest count first (most wasteful)
  incompleteRooms.sort(function(a, b) { return a.count - b.count; });
  
  var suggestions = [];
  var processed = {};
  
  for (var i = 0; i < incompleteRooms.length; i++) {
    var source = incompleteRooms[i];
    if (processed[source.id]) continue;
    
    var sourceMembers = source.members;
    var sourceGenders = {};
    var sourceGroups = {};
    sourceMembers.forEach(function(m) { sourceGenders[m.gender] = true; sourceGroups[m.groupNumber] = true; });
    
    var sameGroup = Object.keys(sourceGroups).length === 1;
    
    // Try 1: Move entire source into one target
    for (var j = incompleteRooms.length - 1; j >= 0; j--) {
      if (i === j || processed[incompleteRooms[j].id]) continue;
      var target = incompleteRooms[j];
      
      // Target must be same or smaller type (upgrade)
      if (target.capacity > source.capacity) continue;
      
      // Check if target has space for source members
      var spaceInTarget = target.capacity - target.count;
      if (spaceInTarget < sourceMembers.length) continue;
      
      // Gender check — if different groups, must be same gender
      var targetGenders = {};
      target.members.forEach(function(m) { targetGenders[m.gender] = true; });
      var mixedGender = false;
      for (var g in sourceGenders) {
        if (targetGenders[g]) continue;
        mixedGender = true;
      }
      
      var targetGroups = {};
      target.members.forEach(function(m) { targetGroups[m.groupNumber] = true; });
      var crossGroup = false;
      for (var gn in sourceGroups) {
        if (!targetGroups[gn]) { crossGroup = true; break; }
      }
      
      if (mixedGender && crossGroup) continue; // No cross-group gender mixing
      
      suggestions.push({
        sourceRoom: source.id,
        sourceType: source.roomType,
        sourceCount: source.count,
        sourceCapacity: source.capacity,
        sourceMembers: sourceMembers.map(function(m) { return { bookingId: m.bookingId, name: m.name, gender: m.gender, groupNumber: m.groupNumber }; }),
        targetRoom: target.id,
        targetType: target.roomType,
        targetCount: target.count,
        targetCapacity: target.capacity,
        result: 'تحرير غرفة ' + source.roomType + ' (' + source.id + ')'
      });
      
      processed[source.id] = true;
      break;
    }
    
    // Try 2: Split source members across multiple targets (only if different groups)
    if (!processed[source.id] && !sameGroup && sourceMembers.length > 1) {
      var splitPlan = [];
      var remainingMembers = sourceMembers.slice();
      
      for (var j = 0; j < incompleteRooms.length && remainingMembers.length > 0; j++) {
        if (i === j || processed[incompleteRooms[j].id]) continue;
        var target = incompleteRooms[j];
        if (target.capacity > source.capacity) continue;
        
        var spaceInTarget = target.capacity - target.count;
        if (spaceInTarget <= 0) continue;
        
        var targetGenders = {};
        target.members.forEach(function(m) { targetGenders[m.gender] = true; });
        
        // Find compatible members to move
        var toMove = [];
        for (var m = remainingMembers.length - 1; m >= 0 && toMove.length < spaceInTarget; m--) {
          var member = remainingMembers[m];
          // Check gender compatibility
          var genderOk = Object.keys(targetGenders).length === 0 || targetGenders[member.gender];
          if (genderOk) {
            toMove.push(member);
            remainingMembers.splice(m, 1);
          }
        }
        
        if (toMove.length > 0) {
          splitPlan.push({ target: target.id, members: toMove });
        }
      }
      
      if (remainingMembers.length === 0 && splitPlan.length > 0) {
        suggestions.push({
          sourceRoom: source.id,
          sourceType: source.roomType,
          sourceCount: source.count,
          sourceCapacity: source.capacity,
          sourceMembers: sourceMembers.map(function(m) { return { bookingId: m.bookingId, name: m.name, gender: m.gender, groupNumber: m.groupNumber }; }),
          splitPlan: splitPlan,
          result: 'تحرير غرفة ' + source.roomType + ' (' + source.id + ') — توزيع على ' + splitPlan.length + ' غرف'
        });
        processed[source.id] = true;
      }
    }
  }
  
  return { 
    suggestions: suggestions,
    incompleteCount: incompleteRooms.length,
    potentialSavings: suggestions.length
  };
}

function executeOptimization(hotelName, suggestion) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, error: 'النظام مشغول — حاول مرة أخرى' };

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var sheet = getHotelSheet_(ss, hotelName);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    var data = sheet.getDataRange().getValues();
    var changes = 0;

    if (suggestion.targetRoom) {
      for (var i = 1; i < data.length; i++) {
        var bid = String(data[i][0]);
        var currentRoom = String(data[i][16] || '');
        if (currentRoom === suggestion.sourceRoom) {
          sheet.getRange(i + 1, 17).setValue(suggestion.targetRoom);
          changes++;
        }
      }
    } else if (suggestion.splitPlan) {
      var memberTargets = {};
      suggestion.splitPlan.forEach(function(plan) {
        plan.members.forEach(function(m) {
          memberTargets[m.bookingId] = plan.target;
        });
      });

      for (var i = 1; i < data.length; i++) {
        var bid = String(data[i][0]);
        if (memberTargets[bid]) {
          sheet.getRange(i + 1, 17).setValue(memberTargets[bid]);
          changes++;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, changes: changes };
  } finally {
    lock.releaseLock();
  }
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
    if (data[i][16] === roomGroupId) { // RoomGroup_ID = column Q (index 16)
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
// HELPER: Write Room Assignment to Hotel Sheet
// ============================================================

function writeRoomAssignment_(sheet, bookingId, roomGroupId) {
  // DEPRECATED — use batchWriteRoomAssignments_ instead
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(bookingId)) {
      sheet.getRange(i + 1, 17).setValue(roomGroupId);
      return;
    }
  }
}

/**
 * Batch write room assignments — reads sheet ONCE, writes ALL at once
 * @param {Sheet} sheet - Hotel sheet
 * @param {Object} assignmentMap - { bookingId: roomGroupId, ... }
 */
function batchWriteRoomAssignments_(sheet, assignmentMap) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  
  var bookingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var roomGroupCol = sheet.getRange(2, 17, lastRow - 1, 1).getValues();
  var roomNumberCol = sheet.getRange(2, 18, lastRow - 1, 1).getValues();
  
  // ActualRoomNo من Room Mapping
  var actualMap = {};
  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var mapSheet = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_MAPPING);
    if (mapSheet) {
      var mapData = mapSheet.getDataRange().getValues();
      var MC = HOTEL_CONFIG.ROOM_MAPPING_COLS;
      for (var m = 1; m < mapData.length; m++) {
        var rid = mapData[m][MC.INTERNAL_ID];
        var actual = mapData[m][MC.ACTUAL_ROOM_NO];
        if (rid && actual) actualMap[String(rid)] = String(actual);
      }
    }
  } catch (e) { /* skip */ }
  
  var count = 0;
  for (var i = 0; i < bookingIds.length; i++) {
    var bid = String(bookingIds[i][0]);
    if (assignmentMap[bid]) {
      var roomId = assignmentMap[bid];
      roomGroupCol[i][0] = roomId;
      // حماية: لا تكتب فوق رقم غرفة موجود يدوياً
      if (!roomNumberCol[i][0] && actualMap[roomId]) roomNumberCol[i][0] = actualMap[roomId];
      count++;
    }
  }
  
  if (count > 0) {
    sheet.getRange(2, 17, lastRow - 1, 1).setValues(roomGroupCol);
    sheet.getRange(2, 18, lastRow - 1, 1).setValues(roomNumberCol);
    SpreadsheetApp.flush();
  }
  
  return count;
}
function diagnoseAutoAssign() {
  var hotelName = 'Swissotel Al Maqam Makkah';
  var hotelCity = 'Makkah';
  
  var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var journeyData = findSheet_(ss, HOTEL_CONFIG.SHEETS.JOURNEY).getDataRange().getValues();
  var roomData = findSheet_(ss, HOTEL_CONFIG.SHEETS.ROOM_TYPE).getDataRange().getValues();
  var guideMap = buildGuideMap_(ss);
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  var R = HOTEL_CONFIG.ROOM_COLS;

  var groupInfo = {};
  for (var i = 1; i < roomData.length; i++) {
    var gn = roomData[i][R.GROUP_NUMBER];
    if (gn) groupInfo[gn] = {
      roomTypeMed: roomData[i][R.ROOM_TYPE_MED], fullRoomsMed: Number(roomData[i][R.FULL_ROOMS_MED])||0, sharedBedsMed: Number(roomData[i][R.SHARED_BEDS_MED])||0,
      roomTypeMak1: roomData[i][R.ROOM_TYPE_MAK1], fullRoomsMak1: Number(roomData[i][R.FULL_ROOMS_MAK1])||0, sharedBedsMak1: Number(roomData[i][R.SHARED_BEDS_MAK1])||0,
      roomTypeMak2: roomData[i][R.ROOM_TYPE_MAK2], fullRoomsMak2: Number(roomData[i][R.FULL_ROOMS_MAK2])||0, sharedBedsMak2: Number(roomData[i][R.SHARED_BEDS_MAK2])||0,
      tourGuide: roomData[i][R.TOUR_GUIDE]||''
    };
  }

  var groupPilgrims = {};
  var dupCheck = {}, dups = [];

  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    if (getHotelName_(row, hotelCity) !== hotelName) continue;
    var bid = String(row[J.BOOKING_ID]);
    if (dupCheck[bid]) dups.push(bid + ' صف ' + (i+1) + ' و ' + dupCheck[bid]);
    dupCheck[bid] = i + 1;
    var gn = row[J.GROUP_NUMBER];
    if (!groupPilgrims[gn]) groupPilgrims[gn] = [];
    groupPilgrims[gn].push({ bookingId: bid, name: row[J.NAME], gender: row[J.GENDER], groupNumber: gn, nationality: row[J.NATIONALITY_EN], arrivalDate: formatDate_(row[J.ARRIVAL_DATE]) });
  }

  var log = [];
  log.push('═══ تشخيص: ' + hotelName + ' (' + hotelCity + ') ═══');
  
  // فحص التكرار
  if (dups.length > 0) { log.push('🚨 تكرار BookingId: ' + dups.length); dups.forEach(function(d){ log.push('  ' + d); }); }
  else log.push('✅ لا تكرار');

  // فحص قيم الجنس
  var gv = {};
  for (var gn in groupPilgrims) groupPilgrims[gn].forEach(function(p){ var g=String(p.gender); gv[g]=(gv[g]||0)+1; });
  log.push('قيم الجنس: ' + JSON.stringify(gv));

  var total = 0;
  for (var gn in groupPilgrims) total += groupPilgrims[gn].length;
  log.push('إجمالي: ' + total + ' حاج | ' + Object.keys(groupPilgrims).length + ' مجموعة');

  // Phase 0
  log.push('\n══ PHASE 0 ══');
  for (var gn in groupPilgrims) {
    var gi = groupInfo[gn];
    if (!gi) { log.push('  grp ' + gn + ': ❌ لا يوجد في Room Type Preview'); continue; }
    var roomType = getRoomType_(gi, hotelCity);
    var fullRooms = getFullRooms_(gi, hotelCity);
    if (fullRooms <= 0) continue;
    var members = groupPilgrims[gn], cap = getRoomCapacity_(roomType);
    var forFull = Math.min(fullRooms * cap, members.length);
    log.push('  grp ' + gn + ': ' + members.length + ' pax | ' + roomType + '(cap=' + cap + ') | full=' + fullRooms + ' → assigned=' + forFull + ' shared=' + (members.length - forFull));
    if (members.length > forFull) groupPilgrims[gn] = members.slice(forFull);
    else delete groupPilgrims[gn];
  }

  // Phase 1
  log.push('\n══ PHASE 1 ══');
  for (var gn in groupPilgrims) {
    var ps = groupPilgrims[gn], guide = '';
    for (var p = 0; p < ps.length; p++) { var g = getGuideForPilgrim_(guideMap, ps[p].passport); if (g) { guide = g; break; } }
    var gi = groupInfo[gn];
    if (!guide && gi && gi.tourGuide && String(gi.tourGuide).trim()) guide = gi.tourGuide;
    if (guide) { log.push('  grp ' + gn + ': guide=' + guide + ' | ' + ps.length + ' pax → يدوي'); delete groupPilgrims[gn]; }
  }

  // Shared Pool
  log.push('\n══ PHASE 2+3 ══');
  var sharedPool = {};
  for (var gn in groupPilgrims) {
    var gi = groupInfo[gn];
    if (!gi) { log.push('  grp ' + gn + ': ❌ لا يوجد في Room Type Preview → ضائعة!'); continue; }
    var roomType = getRoomType_(gi, hotelCity), cap = getRoomCapacity_(roomType), tk = String(cap);
    var ps = groupPilgrims[gn];
    if (ps.length === 0) continue;
    log.push('  grp ' + gn + ': ' + ps.length + ' pax → pool[' + tk + '] | roomType="' + roomType + '"');
    for (var p = 0; p < ps.length; p++) { ps[p].roomType = roomType; ps[p].capacity = cap; if (!sharedPool[tk]) sharedPool[tk]=[]; sharedPool[tk].push(ps[p]); }
  }

  var used = {};
  for (var tk in sharedPool) {
    var pool = sharedPool[tk], cap = parseInt(tk), rooms = [];
    log.push('\n── Pool[' + cap + ']: ' + pool.length + ' pax ──');

    var byGrp = {};
    for (var i = 0; i < pool.length; i++) { var gn = pool[i].groupNumber; if (!byGrp[gn]) byGrp[gn]=[]; byGrp[gn].push(pool[i]); }

    for (var gn in byGrp) {
      var rem = byGrp[gn].slice();
      while (rem.length >= cap) {
        var room = rem.slice(0, cap);
        rooms.push(room);
        room.forEach(function(p){ used[p.bookingId]=true; });
        rem = rem.filter(function(p){ return !used[p.bookingId]; });
        log.push('  P2 grp ' + gn + ': room ' + room.length + '/' + cap);
      }
      if (rem.length > 0) log.push('  P2 grp ' + gn + ': ' + rem.length + ' → P3');
    }

    var remPool = pool.filter(function(p){ return !used[p.bookingId]; });
    log.push('  P3 total: ' + remPool.length);

    ['Female','Male'].forEach(function(gender) {
      var gp = remPool.filter(function(p){ return p.gender === gender; });
      if (gp.length === 0) return;
      log.push('  P3 ' + gender + ': ' + gp.length);
      var room = [];
      for (var i = 0; i < gp.length; i++) {
        if (used[gp[i].bookingId]) continue;
        room.push(gp[i]); used[gp[i].bookingId] = true;
        if (room.length >= cap) { rooms.push(room); log.push('    → room ' + room.length + '/' + cap); room = []; }
      }
      if (room.length > 0) { rooms.push(room); log.push('    → room ' + room.length + '/' + cap + ' (incomplete)'); }
    });

    // ★ فحص النتائج
    for (var r = 0; r < rooms.length; r++) {
      var room = rooms[r], gs={}, grs={};
      room.forEach(function(m){ gs[m.gender]=true; grs[m.groupNumber]=true; });
      if (room.length > cap || (Object.keys(gs).length > 1 && Object.keys(grs).length > 1)) {
        log.push('  🚨 Room ' + r + ': ' + room.length + '/' + cap + ' groups=' + JSON.stringify(Object.keys(grs)) + ' genders=' + JSON.stringify(Object.keys(gs)));
        room.forEach(function(m){ log.push('    ' + m.bookingId + ' | ' + m.name + ' | ' + m.gender + ' | grp=' + m.groupNumber); });
      }
    }
  }

  // فحص حجاج لم يدخلوا أي pool
  var notInPool = 0;
  for (var gn in groupPilgrims) {
    if (!groupInfo[gn]) notInPool += groupPilgrims[gn].length;
  }
  if (notInPool > 0) log.push('\n🚨 ' + notInPool + ' حاج بدون Room Type Preview — لم يُعالجوا!');

  Logger.log(log.join('\n'));
}

// ============================================================
// ROOM MAPPING: حفظ رقم الغرفة الفعلي
// يُشغّل: من واجهة المخزون — حقل إدخال لكل غرفة
// يكتب: عمود ActualRoomNo (H) في شيت Room Mapping
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
// ============================================================
// إعادة تعيين الأسرّة المشتركة فقط — حماية Pre-Assignment
// ============================================================

function clearSharedRoomGroups(hotelName) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, error: 'النظام مشغول — حاول مرة أخرى' };

  try {
    var ss = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
    var hotelSheet = getHotelSheet_(ss, hotelName);
    if (!hotelSheet) return { success: false, error: 'شيت الفندق غير موجود' };

    var protectedIds = {};
    var logSheet = ss.getSheetByName(HOTEL_CONFIG.SHEETS.PRE_ASSIGN_LOG || 'PA_Log');
    if (logSheet) {
      var logData = logSheet.getDataRange().getValues();
      for (var i = 1; i < logData.length; i++) {
        if (String(logData[i][9]) === 'ACTIVE') {
          protectedIds[String(logData[i][0])] = true;
        }
      }
    }

    var lastRow = hotelSheet.getLastRow();
    if (lastRow <= 1) return { success: true, cleared: 0 };
    var numRows = lastRow - 1;

    var rgCol = hotelSheet.getRange(2, 17, numRows, 1).getValues();
    var rnCol = hotelSheet.getRange(2, 18, numRows, 1).getValues();
    var stCol = hotelSheet.getRange(2, 19, numRows, 1).getValues();

    var cleared = 0;
    for (var i = 0; i < rgCol.length; i++) {
      var roomId = String(rgCol[i][0] || '');
      if (!roomId) continue;
      if (protectedIds[roomId]) continue;
      if (String(stCol[i][0] || '') === 'arrived') continue;

      rgCol[i][0] = '';
      rnCol[i][0] = '';
      cleared++;
    }

    if (cleared > 0) {
      hotelSheet.getRange(2, 17, numRows, 1).setValues(rgCol);
      hotelSheet.getRange(2, 18, numRows, 1).setValues(rnCol);
      SpreadsheetApp.flush();
    }

    return { success: true, cleared: cleared, protected: Object.keys(protectedIds).length };
  } finally {
    lock.releaseLock();
  }
}