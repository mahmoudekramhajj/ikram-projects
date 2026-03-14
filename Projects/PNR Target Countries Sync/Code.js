/**
 * ══════════════════════════════════════════════════════════════════
 * مشروع مستقل: مزامنة الدول المستهدفة v5
 * PNR Target Countries → الطيران (CN-CY)
 * ══════════════════════════════════════════════════════════════════
 * 
 * التحديث: توسيع من 9 أعمدة (CN-CV) إلى 12 عمود (CN-CY)
 * المصدر: 13 عمود (A=PNR + B-M=12 دولة)
 * الهدف: أعمدة 92-103 (CN-CY)
 * 
 * الدوال:
 *   previewSync()           → معاينة بدون كتابة
 *   syncTargetCountries()   → تنفيذ المزامنة
 *   coverageReport()        → تقرير التغطية
 *   setupAutoSync()         → مزامنة يومية
 *   removeAutoSync()        → إلغاء المزامنة
 * ══════════════════════════════════════════════════════════════════
 */

var SYNC_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  
  SOURCE_SHEET: 'PNR Target Countries',
  SOURCE_START_ROW: 2,
  SOURCE_COLS: 13,          // A-M (PNR + 12 دولة)
  SOURCE_COUNTRY_COLS: 12,  // B-M (12 دولة)
  
  TARGET_SHEET: 'الطيران',
  TARGET_START_ROW: 3,
  TARGET_PNR_COL: 2,
  TARGET_COUNTRY_START: 92, // CN
  TARGET_COUNTRY_END: 103   // CY (كان 100=CV، أصبح 103=CY)
};


// ==================== المزامنة الرئيسية ====================

function syncTargetCountries() {
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SPREADSHEET_ID);
  
  var sourceData = readSourceData(ss);
  if (!sourceData || Object.keys(sourceData).length === 0) {
    Logger.log('❌ شيت المصدر فارغ أو غير موجود');
    return;
  }
  
  var targetSheet = ss.getSheetByName(SYNC_CONFIG.TARGET_SHEET);
  if (!targetSheet) {
    Logger.log('❌ شيت الطيران غير موجود');
    return;
  }
  
  var lastRow = targetSheet.getLastRow();
  if (lastRow < SYNC_CONFIG.TARGET_START_ROW) {
    Logger.log('⚠️ لا توجد بيانات في شيت الطيران');
    return;
  }
  
  var numRows = lastRow - SYNC_CONFIG.TARGET_START_ROW + 1;
  var pnrValues = targetSheet.getRange(SYNC_CONFIG.TARGET_START_ROW, SYNC_CONFIG.TARGET_PNR_COL, numRows, 1).getValues();
  
  var numCols = SYNC_CONFIG.TARGET_COUNTRY_END - SYNC_CONFIG.TARGET_COUNTRY_START + 1; // 12
  var writeData = [];
  var matchCount = 0;
  var noMatchCount = 0;
  var emptyCount = 0;
  
  for (var i = 0; i < pnrValues.length; i++) {
    var pnrCell = String(pnrValues[i][0] || '').trim();
    
    if (!pnrCell) {
      writeData.push(createEmptyRow(numCols));
      emptyCount++;
      continue;
    }
    
    var countries = findCountries(pnrCell, sourceData);
    
    if (countries) {
      writeData.push(countries);
      matchCount++;
    } else {
      writeData.push(createEmptyRow(numCols));
      noMatchCount++;
    }
  }
  
  // مسح Data Validation ثم الكتابة
  var writeRange = targetSheet.getRange(
    SYNC_CONFIG.TARGET_START_ROW,
    SYNC_CONFIG.TARGET_COUNTRY_START,
    numRows,
    numCols
  );
  writeRange.clearDataValidations();
  writeRange.setValues(writeData);
  
  Logger.log('✅ اكتملت المزامنة');
  Logger.log('━━━━━━━━━━━━━━━━━');
  Logger.log('✅ تطابق: ' + matchCount);
  Logger.log('❌ بدون تطابق: ' + noMatchCount);
  Logger.log('⬜ فارغ: ' + emptyCount);
  Logger.log('📋 الإجمالي: ' + numRows);
  Logger.log('📦 المصدر: ' + Object.keys(sourceData).length + ' PNR');
  Logger.log('📊 الأعمدة: CN-CY (' + numCols + ' عمود)');
}


// ==================== المعاينة ====================

function previewSync() {
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SPREADSHEET_ID);
  
  var sourceData = readSourceData(ss);
  if (!sourceData) {
    Logger.log('❌ شيت المصدر غير موجود');
    return;
  }
  
  var targetSheet = ss.getSheetByName(SYNC_CONFIG.TARGET_SHEET);
  if (!targetSheet) {
    Logger.log('❌ شيت الطيران غير موجود');
    return;
  }
  
  var lastRow = targetSheet.getLastRow();
  var numRows = lastRow - SYNC_CONFIG.TARGET_START_ROW + 1;
  var pnrValues = targetSheet.getRange(SYNC_CONFIG.TARGET_START_ROW, SYNC_CONFIG.TARGET_PNR_COL, numRows, 1).getValues();
  
  var matched = 0, unmatched = 0, empty = 0;
  var examples = [];
  var unmatchedPNRs = [];
  
  for (var i = 0; i < pnrValues.length; i++) {
    var pnrCell = String(pnrValues[i][0] || '').trim();
    if (!pnrCell) { empty++; continue; }
    
    var countries = findCountries(pnrCell, sourceData);
    if (countries) {
      matched++;
      if (examples.length < 5) {
        var filtered = countries.filter(function(c) { return c; });
        examples.push('  ✅ صف ' + (SYNC_CONFIG.TARGET_START_ROW + i) + ': ' + pnrCell + ' → ' + filtered.join(', '));
      }
    } else {
      unmatched++;
      if (unmatchedPNRs.length < 10) {
        unmatchedPNRs.push('  ❌ صف ' + (SYNC_CONFIG.TARGET_START_ROW + i) + ': ' + pnrCell);
      }
    }
  }
  
  Logger.log('👁️ معاينة المزامنة (بدون كتابة)');
  Logger.log('━━━━━━━━━━━━━━━━━');
  Logger.log('✅ سيتطابق: ' + matched);
  Logger.log('❌ بدون تطابق: ' + unmatched);
  Logger.log('⬜ فارغ: ' + empty);
  Logger.log('📊 الأعمدة: CN-CY (12 عمود)');
  
  if (examples.length > 0) {
    Logger.log('\n🔍 أمثلة تطابق:');
    for (var e = 0; e < examples.length; e++) {
      Logger.log(examples[e]);
    }
  }
  
  if (unmatchedPNRs.length > 0) {
    Logger.log('\n⚠️ بدون تطابق:');
    for (var u = 0; u < unmatchedPNRs.length; u++) {
      Logger.log(unmatchedPNRs[u]);
    }
    if (unmatched > 10) {
      Logger.log('  ... و ' + (unmatched - 10) + ' آخرين');
    }
  }
}


// ==================== تقرير التغطية ====================

function coverageReport() {
  var ss = SpreadsheetApp.openById(SYNC_CONFIG.SPREADSHEET_ID);
  
  var sourceData = readSourceData(ss);
  if (!sourceData) {
    Logger.log('❌ شيت المصدر غير موجود');
    return;
  }
  
  var totalSource = Object.keys(sourceData).length;
  var allCountries = {};
  
  for (var pnr in sourceData) {
    var countries = sourceData[pnr];
    for (var i = 0; i < countries.length; i++) {
      var c = String(countries[i] || '').trim();
      if (c) allCountries[c] = (allCountries[c] || 0) + 1;
    }
  }
  
  var sortedCountries = Object.keys(allCountries).sort(function(a, b) {
    return allCountries[b] - allCountries[a];
  });
  
  Logger.log('📊 تقرير تغطية الدول المستهدفة');
  Logger.log('━━━━━━━━━━━━━━━━━');
  Logger.log('📦 PNRs في المصدر: ' + totalSource);
  Logger.log('🌍 دول فريدة: ' + sortedCountries.length);
  Logger.log('');
  
  for (var j = 0; j < sortedCountries.length; j++) {
    Logger.log('  ' + sortedCountries[j] + ': ' + allCountries[sortedCountries[j]] + ' PNR');
  }
}


// ==================== الدوال المساعدة ====================

function readSourceData(ss) {
  var sheet = ss.getSheetByName(SYNC_CONFIG.SOURCE_SHEET);
  if (!sheet) return null;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < SYNC_CONFIG.SOURCE_START_ROW) return null;
  
  var numRows = lastRow - SYNC_CONFIG.SOURCE_START_ROW + 1;
  // قراءة 13 عمود (A-M)
  var data = sheet.getRange(SYNC_CONFIG.SOURCE_START_ROW, 1, numRows, SYNC_CONFIG.SOURCE_COLS).getValues();
  
  var sourceMap = {};
  
  for (var i = 0; i < data.length; i++) {
    var pnr = String(data[i][0] || '').trim();
    if (!pnr) continue;
    
    // B-M = أعمدة 1-12 = 12 دولة → CN-CY
    var countries = [];
    for (var col = 1; col <= SYNC_CONFIG.SOURCE_COUNTRY_COLS; col++) {
      countries.push(String(data[i][col] || '').trim());
    }
    
    // PNR مركب → خزّن كل جزء
    var parts = pnr.split(/[-\s]+/);
    for (var p = 0; p < parts.length; p++) {
      var part = parts[p].trim().toUpperCase();
      if (part) sourceMap[part] = countries;
    }
    
    sourceMap[pnr.toUpperCase()] = countries;
  }
  
  return sourceMap;
}

function findCountries(pnrCell, sourceData) {
  var fullKey = pnrCell.toUpperCase();
  if (sourceData[fullKey]) return sourceData[fullKey];
  
  var parts = pnrCell.split(/[\s\-]+/);
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim().toUpperCase();
    if (part && sourceData[part]) return sourceData[part];
  }
  
  for (var key in sourceData) {
    if (fullKey.indexOf(key) !== -1 || key.indexOf(fullKey) !== -1) {
      return sourceData[key];
    }
  }
  
  return null;
}

function createEmptyRow(numCols) {
  var row = [];
  for (var i = 0; i < numCols; i++) row.push('');
  return row;
}


// ==================== المزامنة التلقائية ====================

function setupAutoSync() {
  removeAutoSync();
  
  ScriptApp.newTrigger('syncTargetCountries')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  
  Logger.log('✅ المزامنة التلقائية مفعّلة — يومياً الساعة 6 صباحاً');
}

function removeAutoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncTargetCountries') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  Logger.log('🛑 تم إلغاء المزامنة التلقائية');
}