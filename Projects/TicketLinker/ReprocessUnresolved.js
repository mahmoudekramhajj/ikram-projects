/**
 * ReprocessUnresolved.js — محاولة مطابقة إضافية لـ 548 صف PENDING_MANUAL
 *
 * الاستراتيجيات (بالترتيب):
 *   C) استخراج جواز من نص النوتس "Found in PD: ... (passport XXXXX)" — يحل ~438 صف
 *   A) ClaudePnr (عمود H) → شيت الطيران → pool العقد → مطابقة اسم
 *   B) AllPassengers (عمود J) → كل اسم راكب → بحث في PD الكامل
 *
 * التحسينات عن النسخة الأولى:
 *   - PD و Flights تُحمَّل مرة واحدة في الذاكرة (cache)
 *   - writeToPdCached_() يكتب مباشرة بدون مسح PD كل مرة (أسرع بكثير)
 *   - Strategy C تُشغَّل أولاً لأنها الأضمن والأسرع
 */

// ==================== أعمدة TL_UnresolvedTickets (0-based index) ====================
var UNRES_COL = {
  ORIGINAL_FILE_ID: 1,   // B
  COPIED_FILE_ID: 2,     // C
  ORIGINAL_FILE_NAME: 3, // D
  FOLDER_NAME: 4,        // E
  EXTRACTED_NAME: 5,     // F
  CLAUDE_FULL_NAME: 6,   // G
  CLAUDE_PNR: 7,         // H
  CLAUDE_TICKET_NO: 8,   // I
  ALL_PASSENGERS: 9,     // J
  FLIGHTS: 10,           // K
  NATIONALITY: 11,       // L
  PASSPORT: 12,          // M
  STATUS: 13,            // N
  LINKED_AT: 14,         // O
  NOTES: 15              // P
};

// ==================== بناء الـ Cache ====================

/**
 * يُحمّل كل بيانات PD مرة واحدة في الذاكرة
 * يقرأ كل الأعمدة ويبني fields map بالتسميات (لـ Strategy E)
 * @return {Array<Object>} [{sheetRow, passport, firstName, lastName, contractName, ticketUrl, fields}]
 */
function loadPdCache_(ss) {
  var sheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (!sheet) throw new Error('PD sheet not found: ' + TL.Config.SHEET_PD);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.min(sheet.getLastColumn(), 40);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var cache = [];
  for (var i = 0; i < data.length; i++) {
    // بناء خريطة label→value لكل الأعمدة غير الفارغة
    var fields = {};
    for (var c = 0; c < lastCol; c++) {
      var h = String(headers[c] || '').trim();
      var v = String(data[i][c] || '').trim();
      if (h && v) fields[h] = v;
    }
    cache.push({
      sheetRow: i + 2,
      passport: String(data[i][5] || '').trim(),      // F
      firstName: String(data[i][10] || '').trim(),     // K
      lastName: String(data[i][11] || '').trim(),      // L
      contractName: String(data[i][21] || '').trim(),  // V
      ticketUrl: String(data[i][25] || '').trim(),     // Z
      fields: fields
    });
  }
  return cache;
}

/**
 * يبني خريطة PNR → contractName من شيت الطيران
 * @return {Object} {PNR: contractName}
 */
function loadFlightMap_(ss) {
  var sheet = ss.getSheetByName(TL.Config.SHEET_FLIGHTS);
  if (!sheet) throw new Error('Flights sheet not found: ' + TL.Config.SHEET_FLIGHTS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var pnrData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var cmData  = sheet.getRange(2, 91, lastRow - 1, 1).getValues();
  var map = {};
  for (var i = 0; i < pnrData.length; i++) {
    var cellPnr = String(pnrData[i][0] || '').trim();
    var contractName = String(cmData[i][0] || '').trim();
    if (!cellPnr || !contractName) continue;
    var pnrs = cellPnr.split(/\s*-\s*/);
    for (var j = 0; j < pnrs.length; j++) {
      var pnr = pnrs[j].trim().toUpperCase();
      if (pnr) map[pnr] = contractName;
    }
  }
  return map;
}

/**
 * يبني خريطة passport → rowIndex من pdCache (لكتابة سريعة)
 * @return {Object} {passport: sheetRow}
 */
function buildPassportIndex_(pdCache) {
  var index = {};
  for (var i = 0; i < pdCache.length; i++) {
    var p = pdCache[i].passport;
    if (p) index[p] = i;  // index في pdCache، وليس sheetRow مباشرة
  }
  return index;
}

// ==================== كتابة سريعة باستخدام Cache ====================

/**
 * يكتب رابط Drive في PD مباشرة باستخدام pdCache (بدون مسح PD كل مرة)
 * @param {string} passport
 * @param {string} fileId
 * @param {Object} passportIndex - {passport: cacheIndex}
 * @param {Array} pdCache
 * @param {Sheet} pdSheet
 * @param {number} ticketUrlCol - العمود 1-based
 * @return {Object} {status, oldZ, newZ, rowIndex}
 */
function writeToPdCached_(passport, fileId, passportIndex, pdCache, pdSheet, ticketUrlCol) {
  if (!passport || !fileId) return { status: 'INVALID_INPUT' };

  var cacheIdx = passportIndex[passport];
  if (cacheIdx === undefined) return { status: 'PASSPORT_NOT_FOUND_IN_PD', passport: passport };

  var pilgrim = pdCache[cacheIdx];
  var rowIndex = pilgrim.sheetRow;
  var newUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
  var currentZ = pilgrim.ticketUrl;

  if (currentZ === newUrl) {
    return { status: 'ALREADY_WRITTEN', rowIndex: rowIndex, oldZ: currentZ, newZ: newUrl };
  }

  pdSheet.getRange(rowIndex, ticketUrlCol).setValue(newUrl);
  // حدّث الـ cache لتعكس التعديل
  pdCache[cacheIdx].ticketUrl = newUrl;

  return {
    status: currentZ ? 'REPLACED' : 'WRITTEN',
    rowIndex: rowIndex,
    oldZ: currentZ,
    newZ: newUrl
  };
}

// ==================== الاستراتيجيات ====================

/**
 * Strategy C: استخراج جواز من نص عمود P (Notes)
 * نمط: "Found in PD: ... (passport XXXXXX) but NOT IN B2C"
 * @return {Object} {linked, passport, notes} أو {noMatch, notes}
 */
function tryStrategyC_(existingNotes) {
  if (!existingNotes) return { noMatch: true, notes: 'C: no existing notes' };
  var m = String(existingNotes).match(/\(passport\s+([A-Z0-9]+)\)/i);
  if (!m) return { noMatch: true, notes: 'C: no passport pattern in notes' };
  return { linked: true, passport: m[1], notes: 'C: passport from notes=' + m[1] };
}

/**
 * Strategy A: ClaudePnr → شيت الطيران → pool عقد → مطابقة اسم
 */
function tryStrategyA_(claudePnr, nameToMatch, flightMap, pdCache) {
  if (!claudePnr || !nameToMatch) return { noMatch: true, notes: 'A: no PNR or name' };

  var pnr = String(claudePnr).trim().toUpperCase();
  var contractName = flightMap[pnr];
  if (!contractName) return { noMatch: true, notes: 'A: PNR ' + pnr + ' not in flights' };

  var pool = pdCache.filter(function(p) { return p.contractName === contractName; });
  if (pool.length === 0) return { noMatch: true, notes: 'A: empty pool for ' + contractName };

  var match = matchNameInCachePool_(nameToMatch, pool);
  if (!match) return { noMatch: true, notes: 'A: name not in pool (' + contractName + ', pool=' + pool.length + ')' };
  if (match.collision) return { collision: true, notes: 'A: collision in ' + contractName + ' for ' + nameToMatch };

  return { linked: true, passport: match.passport, notes: 'A: PNR=' + pnr + ' key=' + match.matchedKey };
}

/**
 * Strategy B: AllPassengers JSON → كل اسم → بحث في PD الكامل
 */
function tryStrategyB_(allPassengersJson, pdCache) {
  if (!allPassengersJson) return { noMatch: true, notes: 'B: no AllPassengers data' };

  var passengers = [];
  try {
    var parsed = JSON.parse(String(allPassengersJson));
    if (Array.isArray(parsed)) {
      parsed.forEach(function(p) {
        var name = (typeof p === 'string') ? p : (p.name || p.fullName || p.full_name || '');
        if (name) passengers.push(String(name).trim());
      });
    }
  } catch (e) {
    var raw = String(allPassengersJson).trim();
    if (raw) passengers.push(raw);
  }

  if (passengers.length === 0) return { noMatch: true, notes: 'B: empty passengers list' };

  var allPassports = {};
  var collisionNames = [];

  for (var i = 0; i < passengers.length; i++) {
    var name = passengers[i];
    if (!name) continue;
    var match = searchNameInPdCache_(name, pdCache);
    if (!match) continue;
    if (match.collision) { collisionNames.push(name); continue; }
    if (match.passport) allPassports[match.passport] = name + '(' + match.matchedKey + ')';
  }

  var passportKeys = Object.keys(allPassports);
  if (passportKeys.length === 1) {
    return { linked: true, passport: passportKeys[0], notes: 'B: matched=' + allPassports[passportKeys[0]] };
  }
  if (passportKeys.length > 1) {
    return { collision: true, notes: 'B: ' + passportKeys.length + ' passports from ' + passengers.length + ' passengers' };
  }
  if (collisionNames.length > 0) {
    return { collision: true, notes: 'B: collision on: ' + collisionNames.slice(0, 2).join(', ') };
  }
  return { noMatch: true, notes: 'B: ' + passengers.length + ' passengers, none matched' };
}

// ==================== Strategy D: إعادة تشغيل Claude على PDF ====================

/**
 * Strategy D: يُعيد تشغيل Claude على PDF مباشرة (للصفوف التي فشل فيها استخراج النص سابقاً)
 * يُستخدم لصفوف: PDF_TEXT_EMPTY، CLAUDE_ERROR، CLAUDE_NO_NAME
 * @return {Object} {linked, passport, notes} أو {collision} أو {noMatch}
 */
function tryStrategyD_(copiedFileId, originalFileId, pdCache) {
  var fileId = copiedFileId || originalFileId;
  if (!fileId) return { noMatch: true, notes: 'D: no fileId' };

  var claudeResult;
  try {
    claudeResult = matchViaPdfClaudeExtraction_({
      fileId: fileId,
      fileName: '',
      folderName: ''
    });
  } catch (e) {
    return { noMatch: true, notes: 'D: exception: ' + e.message.substring(0, 80) };
  }

  if (!claudeResult) return { noMatch: true, notes: 'D: null result' };

  if (claudeResult.status === 'PDF_TEXT_EMPTY') {
    return { noMatch: true, notes: 'D: PDF_TEXT_EMPTY still' };
  }
  if (claudeResult.status === 'CLAUDE_ERROR') {
    return { noMatch: true, notes: 'D: CLAUDE_ERROR: ' + (claudeResult.notes || []).join(' ') };
  }
  if (claudeResult.status === 'CLAUDE_NO_NAME') {
    return { noMatch: true, notes: 'D: CLAUDE_NO_NAME' };
  }
  if (claudeResult.status === 'COLLISION') {
    return { collision: true, notes: 'D: COLLISION' };
  }
  if (claudeResult.status === 'OK' && claudeResult.passport) {
    return {
      linked: true,
      passport: claudeResult.passport,
      notes: 'D: Claude re-extracted name, matched in PD (passport=' + claudeResult.passport + ')'
    };
  }

  return { noMatch: true, notes: 'D: status=' + claudeResult.status };
}

// ==================== Strategy E: Claude يحل تعارض الأسماء ذكاءً ====================

/**
 * يبني prompt لـ Claude يحتوي بيانات التذكرة + بيانات المرشحين الكاملة من PD
 */
function buildCollisionDisambiguationPrompt_(nameOnTicket, flightsJson, candidates) {
  var flightsText = '';
  try {
    var flights = JSON.parse(String(flightsJson || '[]'));
    if (Array.isArray(flights)) {
      for (var fi = 0; fi < flights.length; fi++) {
        var f = flights[fi];
        flightsText += '  - ' + (f.flightNo || '') + ': ' + (f.from || '') + ' → ' + (f.to || '') +
          ' (' + (f.depDate || '') + ' ' + (f.depTime || '') + ')\n';
      }
    }
  } catch (e) {}
  if (!flightsText) flightsText = '  (no flight data)\n';

  var candidatesText = '';
  // أعمدة تُعرض بشكل منفصل — لا تتكرر في fields
  var skipKeys = ['رقم جواز السفر', 'الاسم الأول (إنجليزي)', 'اسم العائلة (إنجليزي)', 'رابط التذكرة'];
  for (var ci = 0; ci < candidates.length; ci++) {
    var c = candidates[ci];
    candidatesText += '\nCandidate ' + (ci + 1) + ':\n';
    candidatesText += '  Passport: ' + c.passport + '\n';
    candidatesText += '  Name: ' + c.firstName + ' ' + c.lastName + '\n';
    candidatesText += '  Contract: ' + (c.contractName || 'unknown') + '\n';
    if (c.fields) {
      var keys = Object.keys(c.fields);
      for (var ki = 0; ki < keys.length; ki++) {
        var key = keys[ki];
        if (skipKeys.indexOf(key) === -1) {
          candidatesText += '  ' + key + ': ' + c.fields[key] + '\n';
        }
      }
    }
  }

  return 'A flight ticket needs to be matched to a pilgrim. Two people share the same name.\n\n' +
    'TICKET:\n' +
    '  Name on ticket: ' + nameOnTicket + '\n' +
    '  Flights:\n' + flightsText + '\n' +
    'DATABASE CANDIDATES:' + candidatesText + '\n' +
    'RULES:\n' +
    '- Flight departure city/country strongly indicates where the pilgrim lives\n' +
    '- Passport format: letter+8digits = British/EU, 9 digits = South African, 8 digits = Indian/Pakistani\n' +
    '- Contract name, nationality, country fields may confirm the match\n' +
    '- Return ONLY valid JSON:\n' +
    '  {"passport":"XXXXX","confidence":"high|medium|low","reason":"brief explanation"}\n' +
    '  OR if genuinely uncertain: {"passport":null,"reason":"explanation"}\n' +
    '- Only commit to a match if you are at least medium confidence\n';
}

/**
 * استدعاء Claude لحل تعارض الأسماء
 */
function callClaudeForDisambiguation_(prompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(TL.Config.PROP.ANTHROPIC_API_KEY);
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set');

  var payload = {
    model: TL.Config.CLAUDE_MODEL,
    max_tokens: 350,
    messages: [{ role: 'user', content: prompt }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': TL.Config.CLAUDE_ANTHROPIC_VERSION },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(TL.Config.CLAUDE_API_URL, options);
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) throw new Error('Claude API ' + code + ': ' + body.substring(0, 200));

  var data = JSON.parse(body);
  var text = (data.content && data.content[0] && data.content[0].text) || '';
  var jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('No JSON in response: ' + text.substring(0, 100));
  return JSON.parse(jsonMatch[0]);
}

/**
 * Strategy E: Claude يحلل بيانات المرشحين الكاملة من PD + بيانات الرحلة ويقرر
 * يُستدعى فقط عند وجود تعارض أسماء (A أو B أرجعا collision)
 */
function tryStrategyE_(claudeFullName, claudeFlightsJson, pdCache) {
  if (!claudeFullName) return { noMatch: true, notes: 'E: no name for disambiguation' };

  // جرّب البحث من جديد للحصول على المرشحين
  var searchResult = searchNameInPdCache_(claudeFullName, pdCache);
  if (!searchResult || !searchResult.collision || !searchResult.matches || searchResult.matches.length < 2) {
    return { noMatch: true, notes: 'E: collision not reproducible in fresh search' };
  }

  // جلب بيانات كل مرشح من الـ cache
  var candidates = [];
  for (var mi = 0; mi < searchResult.matches.length; mi++) {
    var passport = searchResult.matches[mi].passport;
    for (var ci = 0; ci < pdCache.length; ci++) {
      if (pdCache[ci].passport === passport) {
        candidates.push(pdCache[ci]);
        break;
      }
    }
  }
  if (candidates.length < 2) return { noMatch: true, notes: 'E: candidates not in cache' };

  // بناء الـ prompt وإرساله لـ Claude
  var prompt = buildCollisionDisambiguationPrompt_(claudeFullName, claudeFlightsJson, candidates);
  var decision;
  try {
    decision = callClaudeForDisambiguation_(prompt);
  } catch (e) {
    return { noMatch: true, notes: 'E: Claude error: ' + e.message.substring(0, 80) };
  }

  if (!decision || !decision.passport) {
    return { noMatch: true, notes: 'E: Claude uncertain. ' + (decision && decision.reason ? String(decision.reason).substring(0, 100) : '') };
  }
  if (decision.confidence === 'low') {
    return { noMatch: true, notes: 'E: low confidence. ' + String(decision.reason || '').substring(0, 100) };
  }

  return {
    linked: true,
    passport: String(decision.passport),
    notes: 'E: collision resolved (confidence=' + decision.confidence + '). ' + String(decision.reason || '').substring(0, 120)
  };
}

// ==================== بحث في الـ Cache ====================

function matchNameInCachePool_(pdfName, pool) {
  if (!pdfName || !pool || pool.length === 0) return null;
  var split = splitNameByLastWord_(pdfName);
  var pdfKeys = buildAllTokenKeys_(split.firstName, split.lastName);
  pdfKeys.push(normalizeName_(pdfName));
  var uniqPdfKeys = [];
  pdfKeys.forEach(function(k) { if (k && uniqPdfKeys.indexOf(k) === -1) uniqPdfKeys.push(k); });

  var matches = [];
  for (var i = 0; i < pool.length; i++) {
    var p = pool[i];
    var pdKeys = buildAllTokenKeys_(p.firstName, p.lastName);
    for (var j = 0; j < uniqPdfKeys.length; j++) {
      if (pdKeys.indexOf(uniqPdfKeys[j]) !== -1) {
        matches.push({ passport: p.passport, firstName: p.firstName, lastName: p.lastName, matchedKey: uniqPdfKeys[j] });
        break;
      }
    }
  }
  if (matches.length === 1) return matches[0];
  if (matches.length > 1) return { collision: true, matches: matches };
  return null;
}

function searchNameInPdCache_(pdfName, pdCache) {
  if (!pdfName || !pdCache || pdCache.length === 0) return null;
  var split = splitNameByLastWord_(pdfName);
  var pdfKeys = buildAllTokenKeys_(split.firstName, split.lastName);
  pdfKeys.push(normalizeName_(pdfName));
  var uniqPdfKeys = [];
  pdfKeys.forEach(function(k) { if (k && uniqPdfKeys.indexOf(k) === -1) uniqPdfKeys.push(k); });

  var matches = [];
  for (var i = 0; i < pdCache.length; i++) {
    var p = pdCache[i];
    if (!p.firstName && !p.lastName) continue;
    var pdKeys = buildAllTokenKeys_(p.firstName, p.lastName);
    for (var j = 0; j < uniqPdfKeys.length; j++) {
      if (pdKeys.indexOf(uniqPdfKeys[j]) !== -1) {
        matches.push({ passport: p.passport, firstName: p.firstName, lastName: p.lastName, matchedKey: uniqPdfKeys[j] });
        break;
      }
    }
  }
  if (matches.length === 1) return matches[0];
  if (matches.length > 1) return { collision: true, matches: matches };
  return null;
}

// ==================== مساعدات الكتابة والتسجيل ====================

function logAndWrite_(passport, originalFileId, fileName, stage, notes,
                       passportIndex, pdCache, pdSheet, ticketUrlCol) {
  var result = writeToPdCached_(passport, originalFileId, passportIndex, pdCache, pdSheet, ticketUrlCol);
  logWriteOperation_({
    passport: passport,
    fileId: originalFileId,
    fileName: fileName,
    stage: stage,
    status: result.status,
    oldZ: result.oldZ || '',
    newZ: result.newZ || '',
    rowIndex: result.rowIndex || '',
    notes: notes
  });
  return result;
}

/**
 * يحدّث 3 خلايا في صف واحد دفعةً (N، O، P)
 * يستخدم setValues على range متصلة لتقليل API calls
 */
function updateUnresolvedRow_(sheet, rowNum, status, notes, existingNotes) {
  var linkedAt = (status === 'LINKED') ? new Date() : '';
  var newNotes = existingNotes ? existingNotes + ' | ' + notes : notes;
  // N=14, O=15, P=16 — 3 خلايا متصلة → setValues مرة واحدة
  sheet.getRange(rowNum, 14, 1, 3).setValues([[status, linkedAt, newNotes]]);
}

// ==================== الدالة الرئيسية ====================

/**
 * شغّل reprocessUnresolvedTickets() من محرر GAS
 * الترتيب: C (جواز من النوتس) → A (ClaudePnr) → B (AllPassengers)
 */
function reprocessUnresolvedTickets() {
  var startTime = Date.now();
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);

  Logger.log('[Reprocess] تحميل البيانات...');
  var pdCache = loadPdCache_(ss);
  var flightMap = loadFlightMap_(ss);
  var passportIndex = buildPassportIndex_(pdCache);
  Logger.log('[Reprocess] PD=' + pdCache.length + ' passportIndex=' + Object.keys(passportIndex).length + ' FlightPNRs=' + Object.keys(flightMap).length);

  var pdSheet = ss.getSheetByName(TL.Config.SHEET_PD);
  // اجلب رقم عمود TICKET_URL مرة واحدة
  var ticketUrlCol = (function() {
    var lastCol = pdSheet.getLastColumn();
    var headers = pdSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === TL.Config.PD_HEADERS.TICKET_URL) return i + 1;
    }
    return TL.Config.PD_COL.TICKET_URL;
  })();

  var unresolvedSheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!unresolvedSheet) throw new Error('TL_UnresolvedTickets not found');

  var rows = unresolvedSheet.getDataRange().getValues();
  Logger.log('[Reprocess] TL_UnresolvedTickets rows: ' + (rows.length - 1) + ' | ticketUrlCol=' + ticketUrlCol);

  var stats = { total: 0, linkedC: 0, linkedA: 0, linkedB: 0, linkedD: 0, linkedE: 0, collision: 0, noMatch: 0, writeError: 0 };

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (String(row[UNRES_COL.STATUS] || '').trim() !== 'PENDING_MANUAL') continue;

    stats.total++;

    var originalFileId = String(row[UNRES_COL.ORIGINAL_FILE_ID] || '').trim();
    var fileName       = String(row[UNRES_COL.ORIGINAL_FILE_NAME] || '').trim();
    var extractedName  = String(row[UNRES_COL.EXTRACTED_NAME] || '').trim();
    var claudeFullName = String(row[UNRES_COL.CLAUDE_FULL_NAME] || '').trim();
    var claudePnr      = String(row[UNRES_COL.CLAUDE_PNR] || '').trim();
    var allPassengers     = row[UNRES_COL.ALL_PASSENGERS];
    var claudeFlightsJson = String(row[UNRES_COL.FLIGHTS] || '').trim();
    var existingNotes     = String(row[UNRES_COL.NOTES] || '').trim();
    var nameToMatch       = claudeFullName || extractedName;

    var resolved = false;
    var collisionNotes = [];

    // ===== Strategy C: جواز من النوتس =====
    var resultC = tryStrategyC_(existingNotes);
    if (resultC.linked) {
      var writeC = logAndWrite_(resultC.passport, originalFileId, fileName, 'Reprocess-C', resultC.notes,
                                passportIndex, pdCache, pdSheet, ticketUrlCol);
      if (writeC.status === 'WRITTEN' || writeC.status === 'REPLACED' || writeC.status === 'ALREADY_WRITTEN') {
        updateUnresolvedRow_(unresolvedSheet, i + 1, 'LINKED', resultC.notes, existingNotes);
        stats.linkedC++;
        resolved = true;
      } else {
        updateUnresolvedRow_(unresolvedSheet, i + 1, 'PENDING_MANUAL',
          'Reprocess-C write error: ' + writeC.status + ' passport=' + resultC.passport, existingNotes);
        stats.writeError++;
        resolved = true;
      }
    }

    // ===== Strategy A: ClaudePnr =====
    if (!resolved && claudePnr) {
      var resultA = tryStrategyA_(claudePnr, nameToMatch, flightMap, pdCache);
      if (resultA.linked) {
        var writeA = logAndWrite_(resultA.passport, originalFileId, fileName, 'Reprocess-A', resultA.notes,
                                   passportIndex, pdCache, pdSheet, ticketUrlCol);
        if (writeA.status === 'WRITTEN' || writeA.status === 'REPLACED' || writeA.status === 'ALREADY_WRITTEN') {
          updateUnresolvedRow_(unresolvedSheet, i + 1, 'LINKED', resultA.notes, existingNotes);
          stats.linkedA++;
          resolved = true;
        } else {
          updateUnresolvedRow_(unresolvedSheet, i + 1, 'PENDING_MANUAL',
            'Reprocess-A write error: ' + writeA.status, existingNotes);
          stats.writeError++;
          resolved = true;
        }
      } else if (resultA.collision) {
        collisionNotes.push(resultA.notes);
      }
    }

    // ===== Strategy B: AllPassengers =====
    if (!resolved && allPassengers) {
      var resultB = tryStrategyB_(String(allPassengers), pdCache);
      if (resultB.linked) {
        var writeB = logAndWrite_(resultB.passport, originalFileId, fileName, 'Reprocess-B', resultB.notes,
                                   passportIndex, pdCache, pdSheet, ticketUrlCol);
        if (writeB.status === 'WRITTEN' || writeB.status === 'REPLACED' || writeB.status === 'ALREADY_WRITTEN') {
          updateUnresolvedRow_(unresolvedSheet, i + 1, 'LINKED', resultB.notes, existingNotes);
          stats.linkedB++;
          resolved = true;
        } else {
          updateUnresolvedRow_(unresolvedSheet, i + 1, 'PENDING_MANUAL',
            'Reprocess-B write error: ' + writeB.status, existingNotes);
          stats.writeError++;
          resolved = true;
        }
      } else if (resultB.collision) {
        collisionNotes.push(resultB.notes);
      }
    }

    // ===== Strategy E: Claude يحل تعارض الأسماء بتحليل بيانات PD كاملة =====
    // يُشغَّل إذا وجد collision في هذه الجولة، أو إذا كانت النوتس القديمة تحتوي COLLISION
    var hasExistingCollision = existingNotes.indexOf('COLLISION') !== -1;
    if (!resolved && (collisionNotes.length > 0 || hasExistingCollision) && nameToMatch) {
      var resultE = tryStrategyE_(nameToMatch, claudeFlightsJson, pdCache);
      if (resultE.linked) {
        var writeE = logAndWrite_(resultE.passport, originalFileId, fileName, 'Reprocess-E', resultE.notes,
                                   passportIndex, pdCache, pdSheet, ticketUrlCol);
        if (writeE.status === 'WRITTEN' || writeE.status === 'REPLACED' || writeE.status === 'ALREADY_WRITTEN') {
          updateUnresolvedRow_(unresolvedSheet, i + 1, 'LINKED', resultE.notes, existingNotes);
          stats.linkedE++;
          resolved = true;
        } else {
          updateUnresolvedRow_(unresolvedSheet, i + 1, 'PENDING_MANUAL',
            'Reprocess-E write error: ' + writeE.status, existingNotes);
          stats.writeError++;
          resolved = true;
        }
      }
      // E يضيف للـ notes التفسير حتى لو فشل — لا نضيف لـ collisionNotes
      Utilities.sleep(500); // rate limit خفيف بين استدعاءات Claude
    }

    // ===== Strategy D: إعادة تشغيل Claude (للصفوف التي فشل فيها استخراج النص) =====
    if (!resolved) {
      var needsClaudeRetry = existingNotes.indexOf('PDF_TEXT_EMPTY') !== -1 ||
                             existingNotes.indexOf('CLAUDE_ERROR') !== -1 ||
                             existingNotes.indexOf('CLAUDE_NO_NAME') !== -1 ||
                             existingNotes.indexOf('GO_TO_STAGE_5') !== -1;
      if (needsClaudeRetry) {
        var copiedFileId = String(row[UNRES_COL.COPIED_FILE_ID] || '').trim();
        var resultD = tryStrategyD_(copiedFileId, originalFileId, pdCache);
        if (resultD.linked) {
          var writeD = logAndWrite_(resultD.passport, originalFileId, fileName, 'Reprocess-D', resultD.notes,
                                     passportIndex, pdCache, pdSheet, ticketUrlCol);
          if (writeD.status === 'WRITTEN' || writeD.status === 'REPLACED' || writeD.status === 'ALREADY_WRITTEN') {
            updateUnresolvedRow_(unresolvedSheet, i + 1, 'LINKED', resultD.notes, existingNotes);
            stats.linkedD++;
            resolved = true;
          } else {
            updateUnresolvedRow_(unresolvedSheet, i + 1, 'PENDING_MANUAL',
              'Reprocess-D write error: ' + writeD.status, existingNotes);
            stats.writeError++;
            resolved = true;
          }
        } else if (resultD.collision) {
          collisionNotes.push(resultD.notes);
        }
        // rate limit بسيط بين استدعاءات Claude
        Utilities.sleep(1200);
      }
    }

    if (!resolved) {
      if (collisionNotes.length > 0) {
        updateUnresolvedRow_(unresolvedSheet, i + 1, 'PENDING_MANUAL', collisionNotes.join(' | '), existingNotes);
        stats.collision++;
      } else {
        stats.noMatch++;
      }
    }

    // حماية من timeout
    if (Date.now() - startTime > 5 * 60 * 1000) {
      Logger.log('[Reprocess] تحذير: timeout بعد ' + stats.total + ' صف');
      break;
    }
  }

  var elapsed = Math.round((Date.now() - startTime) / 1000);
  var summary = 'Reprocess ' + elapsed + 'ث | C=' + stats.linkedC +
    ' A=' + stats.linkedA + ' B=' + stats.linkedB + ' D=' + stats.linkedD +
    ' E=' + stats.linkedE +
    ' collision=' + stats.collision + ' noMatch=' + stats.noMatch +
    ' writeError=' + stats.writeError + ' / total=' + stats.total;

  Logger.log('[Reprocess] ' + summary);
  return stats;
}
