// ╔══════════════════════════════════════════════════════════════╗
// ║              B2C Sync — مزامنة بيانات B2C + الرحلات         ║
// ╠══════════════════════════════════════════════════════════════╣
// ║  الخطوة 1: إضافة حجاج B2C الجدد من PD إلى شيت B2C         ║
// ║  الخطوة 2: سحب بيانات الرحلات من GDS إلى B2C               ║
// ╚══════════════════════════════════════════════════════════════╝

// ==================== الإعدادات ====================

var B2C_CONFIG = {
  // الشيت الرئيسي (PD + B2C)
  MAIN_SS_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  PD_SHEET: 'Presonal Details',
  B2C_SHEET: 'B2C',
  
  // شيت GDS (خارجي)
  GDS_SS_ID: '1poy9gVjicyIF4LZyTsaKrOG4u-w4KiO54HOrWYXjwaM',
  GDS_SHEET: 'B2C',
  
  // أعمدة مرجعية
  PASSPORT_COL: 6,         // F — رقم الجواز (مفتاح الربط)
  CONTRACT_TYPE_COL: 21,   // U — نوع عقد الطيران
  
  // أعمدة PD المطلوب نسخها
  PD_DIRECT_COLS: 27,      // A-AA (1-27) نسخ مباشر
  PD_CAMP_COL: 28,         // AB — المخيم (يُتجاوز)
  PD_SHIFT_START: 29,      // AC — بداية الأعمدة المُزاحة
  PD_SHIFT_END: 33,        // AG — نهاية الأعمدة المُزاحة
  
  // أعمدة B2C
  B2C_PERSONAL_END: 32,    // AF — آخر عمود بيانات شخصية
  B2C_FLIGHT_START: 33,    // AG — بداية بيانات الرحلات
  B2C_FLIGHT_END: 61,      // BI — نهاية بيانات الرحلات (PNR)
  
  // أعمدة GDS
  GDS_FLIGHT_START: 33,
  GDS_FLIGHT_END: 61
};


// ==================== الدالة الرئيسية ====================

/**
 * تشغيل كامل: مزامنة الأسماء + سحب الرحلات
 */
function syncB2C() {
  var startTime = new Date();
  Logger.log('╔══════════════════════════════════════╗');
  Logger.log('║     بدء مزامنة B2C الكاملة          ║');
  Logger.log('╚══════════════════════════════════════╝');
  Logger.log('⏰ ' + startTime.toLocaleString());
  
  // الخطوة 1
  var namesResult = _syncNames();
  
  // الخطوة 2
  var flightsResult = _syncFlights();
  
  // التقرير النهائي
  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📊 التقرير النهائي');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('👤 أسماء جديدة مُضافة: ' + namesResult.added);
  Logger.log('👤 موجودون مسبقاً: ' + namesResult.existing);
  Logger.log('✈️ رحلات مسحوبة: ' + flightsResult.filled);
  Logger.log('⏳ بدون رحلة في GDS: ' + flightsResult.notInGds);
  Logger.log('⏭️ رحلات موجودة مسبقاً: ' + flightsResult.alreadyFilled);
  Logger.log('⏱️ الزمن: ' + elapsed + ' ثانية');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}


/**
 * الخطوة 1 فقط: إضافة الأسماء الجديدة
 */
function syncNamesOnly() {
  Logger.log('▶️ تشغيل مزامنة الأسماء فقط...');
  var result = _syncNames();
  Logger.log('✅ انتهى — مُضاف: ' + result.added + ' | موجود: ' + result.existing);
}


/**
 * الخطوة 2 فقط: سحب الرحلات
 */
function syncFlightsOnly() {
  Logger.log('▶️ تشغيل سحب الرحلات فقط...');
  var result = _syncFlights();
  Logger.log('✅ انتهى — مسحوب: ' + result.filled + ' | بدون GDS: ' + result.notInGds + ' | موجود: ' + result.alreadyFilled);
}


// ==================== الخطوة 1: مزامنة الأسماء ====================

function _syncNames() {
  Logger.log('');
  Logger.log('━━━ الخطوة 1: مزامنة الأسماء ━━━');
  
  var ss = SpreadsheetApp.openById(B2C_CONFIG.MAIN_SS_ID);
  var pdSheet = ss.getSheetByName(B2C_CONFIG.PD_SHEET);
  var b2cSheet = ss.getSheetByName(B2C_CONFIG.B2C_SHEET);
  
  if (!pdSheet || !b2cSheet) {
    Logger.log('❌ شيت غير موجود!');
    return { added: 0, existing: 0 };
  }
  
  // قراءة جوازات B2C الحالية
  var b2cLastRow = b2cSheet.getLastRow();
  var b2cPassports = {};
  
  if (b2cLastRow > 1) {
    var b2cData = b2cSheet.getRange(2, B2C_CONFIG.PASSPORT_COL, b2cLastRow - 1, 1).getValues();
    for (var i = 0; i < b2cData.length; i++) {
      var p = String(b2cData[i][0] || '').trim();
      if (p) b2cPassports[p] = true;
    }
  }
  Logger.log('📋 جوازات B2C الحالية: ' + Object.keys(b2cPassports).length);
  
  // قراءة PD — فلترة B2C فقط
  var pdLastRow = pdSheet.getLastRow();
  var pdData = pdSheet.getRange(2, 1, pdLastRow - 1, B2C_CONFIG.PD_SHIFT_END).getValues();
  
  var newRows = [];
  var existing = 0;
  
  for (var j = 0; j < pdData.length; j++) {
    var row = pdData[j];
    var contractType = String(row[B2C_CONFIG.CONTRACT_TYPE_COL - 1] || '').trim();
    
    // فلتر: B2C فقط
    if (contractType !== 'B2C') continue;
    
    var passport = String(row[B2C_CONFIG.PASSPORT_COL - 1] || '').trim();
    if (!passport) continue;
    
    // تحقق: موجود مسبقاً؟
    if (b2cPassports[passport]) {
      existing++;
      continue;
    }
    
    // بناء صف B2C الجديد
    var b2cRow = _mapPdToB2c(row);
    newRows.push(b2cRow);
    
    // تسجيل الجواز لتفادي التكرار في نفس الدفعة
    b2cPassports[passport] = true;
  }
  
  Logger.log('🆕 أسماء جديدة للإضافة: ' + newRows.length);
  Logger.log('⏭️ موجودون مسبقاً: ' + existing);
  
  // الكتابة
  if (newRows.length > 0) {
    var writeRow = b2cLastRow + 1;
    // نكتب 32 عمود (بيانات شخصية) + 29 عمود فارغ (رحلات)
    var fullRows = [];
    for (var k = 0; k < newRows.length; k++) {
      var personalData = newRows[k]; // 32 عمود
      var emptyFlights = new Array(B2C_CONFIG.B2C_FLIGHT_END - B2C_CONFIG.B2C_FLIGHT_START + 1).fill('');
      fullRows.push(personalData.concat(emptyFlights));
    }
    
    b2cSheet.getRange(writeRow, 1, fullRows.length, fullRows[0].length).setValues(fullRows);
    Logger.log('✅ تمت إضافة ' + newRows.length + ' صف في B2C (بداية من صف ' + writeRow + ')');
  }
  
  return { added: newRows.length, existing: existing };
}


/**
 * تحويل صف PD إلى صف B2C (32 عمود)
 * PD: A-AA(1-27) نسخ مباشر → B2C A-AA(1-27)
 * PD: AB(28) المخيم → يُتجاوز
 * PD: AC-AG(29-33) → B2C AB-AF(28-32)
 */
function _mapPdToB2c(pdRow) {
  var b2cRow = [];
  
  // نسخ مباشر: أعمدة 1-27 (index 0-26)
  for (var i = 0; i < B2C_CONFIG.PD_DIRECT_COLS; i++) {
    b2cRow.push(pdRow[i] || '');
  }
  
  // تجاوز عمود 28 (المخيم) — index 27
  
  // نسخ أعمدة 29-33 (index 28-32) → B2C أعمدة 28-32
  for (var j = B2C_CONFIG.PD_SHIFT_START - 1; j < B2C_CONFIG.PD_SHIFT_END; j++) {
    b2cRow.push(pdRow[j] || '');
  }
  
  // المجموع: 27 + 5 = 32 عمود
  return b2cRow;
}


// ==================== الخطوة 2: سحب الرحلات ====================

function _syncFlights() {
  Logger.log('');
  Logger.log('━━━ الخطوة 2: سحب الرحلات ━━━');
  
  var mainSS = SpreadsheetApp.openById(B2C_CONFIG.MAIN_SS_ID);
  var b2cSheet = mainSS.getSheetByName(B2C_CONFIG.B2C_SHEET);
  
  if (!b2cSheet) {
    Logger.log('❌ شيت B2C غير موجود!');
    return { filled: 0, notInGds: 0, alreadyFilled: 0 };
  }
  
  // قراءة GDS
  var gdsSS = SpreadsheetApp.openById(B2C_CONFIG.GDS_SS_ID);
  var gdsSheet = gdsSS.getSheetByName(B2C_CONFIG.GDS_SHEET);
  
  if (!gdsSheet) {
    Logger.log('❌ شيت GDS غير موجود!');
    return { filled: 0, notInGds: 0, alreadyFilled: 0 };
  }
  
  // بناء فهرس GDS: جواز → بيانات الرحلة
  var gdsIndex = _buildGdsIndex(gdsSheet);
  Logger.log('📋 فهرس GDS: ' + Object.keys(gdsIndex).length + ' جواز');
  
  // قراءة B2C
  var b2cLastRow = b2cSheet.getLastRow();
  if (b2cLastRow < 2) {
    Logger.log('⚠️ شيت B2C فارغ');
    return { filled: 0, notInGds: 0, alreadyFilled: 0 };
  }
  
  var numRows = b2cLastRow - 1;
  
  // قراءة جوازات B2C
  var b2cPassports = b2cSheet.getRange(2, B2C_CONFIG.PASSPORT_COL, numRows, 1).getValues();
  
  // قراءة بيانات الرحلات الحالية (لتحديد المملوء)
  var flightCols = B2C_CONFIG.B2C_FLIGHT_END - B2C_CONFIG.B2C_FLIGHT_START + 1; // 29 عمود
  var currentFlights = b2cSheet.getRange(2, B2C_CONFIG.B2C_FLIGHT_START, numRows, flightCols).getValues();
  
  var filled = 0;
  var notInGds = 0;
  var alreadyFilled = 0;
  var updatedRows = []; // [{rowIndex, data}]
  
  for (var i = 0; i < numRows; i++) {
    var passport = String(b2cPassports[i][0] || '').trim();
    if (!passport) continue;
    
    // تحقق: هل بيانات الرحلة مملوءة مسبقاً؟
    var hasFlightData = _hasFlightData(currentFlights[i]);
    if (hasFlightData) {
      alreadyFilled++;
      continue;
    }
    
    // البحث في GDS
    var gdsFlights = gdsIndex[passport];
    if (!gdsFlights) {
      notInGds++;
      continue;
    }
    
    // تسجيل الصف للتحديث
    updatedRows.push({
      rowIndex: i + 2, // الصف الفعلي في الشيت
      data: gdsFlights
    });
    filled++;
  }
  
  Logger.log('✈️ رحلات جاهزة للسحب: ' + filled);
  Logger.log('⏳ بدون رحلة في GDS: ' + notInGds);
  Logger.log('⏭️ مملوءة مسبقاً: ' + alreadyFilled);
  
  // الكتابة — صف بصف لتجنب الكتابة على صفوف غير مستهدفة
  if (updatedRows.length > 0) {
    for (var j = 0; j < updatedRows.length; j++) {
      var entry = updatedRows[j];
      b2cSheet.getRange(entry.rowIndex, B2C_CONFIG.B2C_FLIGHT_START, 1, flightCols).setValues([entry.data]);
    }
    Logger.log('✅ تم تحديث ' + updatedRows.length + ' صف بالرحلات');
  }
  
  return { filled: filled, notInGds: notInGds, alreadyFilled: alreadyFilled };
}


/**
 * بناء فهرس GDS: { جواز: [29 قيمة رحلة] }
 */
function _buildGdsIndex(gdsSheet) {
  var lastRow = gdsSheet.getLastRow();
  if (lastRow < 2) return {};
  
  var numRows = lastRow - 1;
  
  // قراءة أعمدة الجواز
  var passports = gdsSheet.getRange(2, B2C_CONFIG.PASSPORT_COL, numRows, 1).getValues();
  
  // قراءة أعمدة الرحلات (33-61)
  var flightCols = B2C_CONFIG.GDS_FLIGHT_END - B2C_CONFIG.GDS_FLIGHT_START + 1;
  var flights = gdsSheet.getRange(2, B2C_CONFIG.GDS_FLIGHT_START, numRows, flightCols).getValues();
  
  var index = {};
  for (var i = 0; i < numRows; i++) {
    var passport = String(passports[i][0] || '').trim();
    if (!passport) continue;
    
    // تحقق أن الصف فيه بيانات رحلة فعلية
    if (!_hasFlightData(flights[i])) continue;
    
    index[passport] = flights[i];
  }
  
  return index;
}


/**
 * تحقق: هل صف الرحلة يحتوي بيانات؟
 * يكفي وجود أي قيمة في أول 7 خانات (FlightNo أو From أو PNR)
 */
function _hasFlightData(flightRow) {
  if (!flightRow) return false;
  // تحقق من أعمدة رئيسية: FlightNo1(0), From1(3), FlightNo2(7), PNR(28)
  var checkIndices = [0, 3, 7, 10, 14, 21, 28];
  for (var i = 0; i < checkIndices.length; i++) {
    var idx = checkIndices[i];
    if (idx < flightRow.length && flightRow[idx] && String(flightRow[idx]).trim() !== '') {
      return true;
    }
  }
  return false;
}


// ==================== المعاينة ====================

/**
 * معاينة بدون كتابة — لمراجعة ما سيحدث
 */
function previewSync() {
  Logger.log('╔══════════════════════════════════════╗');
  Logger.log('║     👁️ معاينة المزامنة (بدون كتابة)   ║');
  Logger.log('╚══════════════════════════════════════╝');
  
  var ss = SpreadsheetApp.openById(B2C_CONFIG.MAIN_SS_ID);
  var pdSheet = ss.getSheetByName(B2C_CONFIG.PD_SHEET);
  var b2cSheet = ss.getSheetByName(B2C_CONFIG.B2C_SHEET);
  
  // --- معاينة الأسماء ---
  var b2cLastRow = b2cSheet.getLastRow();
  var b2cPassports = {};
  
  if (b2cLastRow > 1) {
    var b2cData = b2cSheet.getRange(2, B2C_CONFIG.PASSPORT_COL, b2cLastRow - 1, 1).getValues();
    for (var i = 0; i < b2cData.length; i++) {
      var p = String(b2cData[i][0] || '').trim();
      if (p) b2cPassports[p] = true;
    }
  }
  
  var pdLastRow = pdSheet.getLastRow();
  var pdData = pdSheet.getRange(2, 1, pdLastRow - 1, B2C_CONFIG.CONTRACT_TYPE_COL).getValues();
  
  var newNames = [];
  for (var j = 0; j < pdData.length; j++) {
    var contractType = String(pdData[j][B2C_CONFIG.CONTRACT_TYPE_COL - 1] || '').trim();
    if (contractType !== 'B2C') continue;
    
    var passport = String(pdData[j][B2C_CONFIG.PASSPORT_COL - 1] || '').trim();
    if (!passport || b2cPassports[passport]) continue;
    
    // اسم للعرض
    var nameEn = String(pdData[j][10] || '') + ' ' + String(pdData[j][11] || '');
    newNames.push({ passport: passport, name: nameEn.trim() });
    b2cPassports[passport] = true;
  }
  
  Logger.log('');
  Logger.log('━━━ الأسماء الجديدة: ' + newNames.length + ' ━━━');
  for (var k = 0; k < Math.min(newNames.length, 15); k++) {
    Logger.log('  🆕 ' + newNames[k].name + ' (' + newNames[k].passport + ')');
  }
  if (newNames.length > 15) {
    Logger.log('  ... و ' + (newNames.length - 15) + ' آخرين');
  }
  
  // --- معاينة الرحلات ---
  var gdsSS = SpreadsheetApp.openById(B2C_CONFIG.GDS_SS_ID);
  var gdsSheet = gdsSS.getSheetByName(B2C_CONFIG.GDS_SHEET);
  var gdsIndex = _buildGdsIndex(gdsSheet);
  
  var numB2c = b2cLastRow > 1 ? b2cLastRow - 1 : 0;
  var flightCols = B2C_CONFIG.B2C_FLIGHT_END - B2C_CONFIG.B2C_FLIGHT_START + 1;
  
  var missingFlights = 0;
  var canFill = 0;
  var noGds = 0;
  var alreadyFilled = 0;
  
  if (numB2c > 0) {
    var b2cPassportVals = b2cSheet.getRange(2, B2C_CONFIG.PASSPORT_COL, numB2c, 1).getValues();
    var currentFlights = b2cSheet.getRange(2, B2C_CONFIG.B2C_FLIGHT_START, numB2c, flightCols).getValues();
    
    for (var m = 0; m < numB2c; m++) {
      var pp = String(b2cPassportVals[m][0] || '').trim();
      if (!pp) continue;
      
      if (_hasFlightData(currentFlights[m])) {
        alreadyFilled++;
        continue;
      }
      
      missingFlights++;
      if (gdsIndex[pp]) {
        canFill++;
      } else {
        noGds++;
      }
    }
  }
  
  Logger.log('');
  Logger.log('━━━ الرحلات ━━━');
  Logger.log('📋 إجمالي B2C: ' + numB2c);
  Logger.log('✅ رحلات مملوءة: ' + alreadyFilled);
  Logger.log('❓ بدون رحلة: ' + missingFlights);
  Logger.log('  ✈️ يمكن سحبها من GDS: ' + canFill);
  Logger.log('  ⏳ غير متوفرة في GDS: ' + noGds);
  Logger.log('📦 فهرس GDS: ' + Object.keys(gdsIndex).length + ' جواز');
}


// ==================== تقرير الحالة ====================

/**
 * تقرير شامل عن حالة B2C
 */
function syncStatus() {
  Logger.log('╔══════════════════════════════════════╗');
  Logger.log('║         📊 حالة شيت B2C              ║');
  Logger.log('╚══════════════════════════════════════╝');
  
  var ss = SpreadsheetApp.openById(B2C_CONFIG.MAIN_SS_ID);
  var pdSheet = ss.getSheetByName(B2C_CONFIG.PD_SHEET);
  var b2cSheet = ss.getSheetByName(B2C_CONFIG.B2C_SHEET);
  
  // إحصائيات PD
  var pdLastRow = pdSheet.getLastRow();
  var pdTypes = pdSheet.getRange(2, B2C_CONFIG.CONTRACT_TYPE_COL, pdLastRow - 1, 1).getValues();
  var pdB2cCount = 0;
  for (var i = 0; i < pdTypes.length; i++) {
    if (String(pdTypes[i][0] || '').trim() === 'B2C') pdB2cCount++;
  }
  
  // إحصائيات B2C
  var b2cLastRow = b2cSheet.getLastRow();
  var b2cCount = b2cLastRow > 1 ? b2cLastRow - 1 : 0;
  
  var withFlights = 0;
  var withoutFlights = 0;
  var withPNR = 0;
  
  if (b2cCount > 0) {
    var flightCols = B2C_CONFIG.B2C_FLIGHT_END - B2C_CONFIG.B2C_FLIGHT_START + 1;
    var flightData = b2cSheet.getRange(2, B2C_CONFIG.B2C_FLIGHT_START, b2cCount, flightCols).getValues();
    
    for (var j = 0; j < b2cCount; j++) {
      if (_hasFlightData(flightData[j])) {
        withFlights++;
        // PNR — آخر عمود (index 28)
        if (flightData[j][28] && String(flightData[j][28]).trim() !== '') {
          withPNR++;
        }
      } else {
        withoutFlights++;
      }
    }
  }
  
  Logger.log('');
  Logger.log('━━━ الأرقام ━━━');
  Logger.log('👤 B2C في PD: ' + pdB2cCount);
  Logger.log('👤 صفوف B2C: ' + b2cCount);
  Logger.log('🔄 الفرق (يحتاج مزامنة أسماء): ' + (pdB2cCount - b2cCount));
  Logger.log('');
  Logger.log('━━━ الرحلات ━━━');
  Logger.log('✅ مع بيانات رحلة: ' + withFlights + ' (' + (b2cCount > 0 ? Math.round(withFlights/b2cCount*100) : 0) + '%)');
  Logger.log('❌ بدون رحلة: ' + withoutFlights);
  Logger.log('🎫 مع PNR: ' + withPNR);
}


// ==================== القائمة في الشيت ====================

/**
 * إضافة قائمة "B2C Sync" عند فتح الشيت
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 B2C Sync')
    .addItem('▶️ مزامنة كاملة (أسماء + رحلات)', 'syncB2C')
    .addSeparator()
    .addItem('👤 مزامنة أسماء فقط', 'syncNamesOnly')
    .addItem('✈️ سحب رحلات فقط', 'syncFlightsOnly')
    .addSeparator()
    .addItem('👁️ معاينة (بدون كتابة)', 'previewSync')
    .addItem('📊 تقرير الحالة', 'syncStatus')
    .addSeparator()
    .addItem('⏰ تفعيل التشغيل اليومي', 'setupDailyTrigger')
    .addItem('🛑 إيقاف التشغيل اليومي', 'removeTriggers')
    .addToUi();
}


// ==================== Triggers ====================

/**
 * شغّلها مرة واحدة فقط — تُفعّل القائمة + التشغيل اليومي
 */
function initialSetup() {
  var triggers = ScriptApp.getProjectTriggers();
  
  // حذف triggers سابقة
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // 1. Trigger لإظهار القائمة عند فتح الشيت
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(B2C_CONFIG.MAIN_SS_ID)
    .onOpen()
    .create();
  
  // 2. Trigger يومي الساعة 6 صباحاً
  ScriptApp.newTrigger('syncB2C')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  
  Logger.log('✅ تم الإعداد:');
  Logger.log('  📋 القائمة ستظهر عند فتح الشيت');
  Logger.log('  ⏰ المزامنة اليومية الساعة 6 صباحاً');
}

/**
 * تفعيل التشغيل اليومي فقط
 */
function setupDailyTrigger() {
  // حذف trigger يومي سابق فقط (إبقاء onOpen)
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncB2C') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('syncB2C')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  
  Logger.log('✅ تم تفعيل التشغيل اليومي الساعة 6 صباحاً');
  SpreadsheetApp.getUi().alert('✅ تم تفعيل المزامنة اليومية\n\nسيتم تشغيل syncB2C تلقائياً كل يوم الساعة 6 صباحاً.');
}

/**
 * إيقاف التشغيل اليومي فقط (إبقاء القائمة)
 */
function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncB2C') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  Logger.log('🛑 تم إيقاف التشغيل اليومي (' + removed + ' trigger)');
  SpreadsheetApp.getUi().alert('🛑 تم إيقاف التشغيل اليومي');
}