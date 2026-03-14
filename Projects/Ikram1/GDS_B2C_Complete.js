/**
 * ══════════════════════════════════════════════════════════════════
 * GDS B2C Complete v1.0 — ملف شامل لاستخراج وإكمال بيانات الرحلات
 * ══════════════════════════════════════════════════════════════════
 * 
 * يعمل على شيت B2C داخل ملف Ikram Abuown
 * يجمع الوظائف الثلاثة في ملف واحد:
 *   1. استخراج بيانات الرحلات من تذاكر PDF
 *   2. إكمال بيانات الوصول عبر AeroDataBox API
 *   3. إصلاح تصنيف الرحلات المباشرة
 * 
 * التصنيف الصحيح:
 *   مباشر  → AN-AT (ذهاب) + AU-BA (عودة) — AG-AM فارغ
 *   ترانزيت → AG-AM (رحلة 1) + AN-AT (رحلة 2) ذهاباً
 *              AU-BA (رحلة 1) + BB-BH (رحلة 2) عودةً
 * 
 * الأوامر:
 *   من القائمة: ✈️ بيانات رحلات B2C
 *   أو تلقائياً كل 10 دقائق
 * ══════════════════════════════════════════════════════════════════
 */


// ==================== الإعدادات ====================

var CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'B2C',
  DATA_START: 2,
  
  // أعمدة مرجعية
  COL_PASSPORT: 6,     // F — رقم الجواز
  COL_TICKET: 25,      // Y — رقم التذكرة
  COL_LINK: 26,        // Z — رابط التذكرة PDF
  COL_PNR: 61,         // BI — Airline PNR
  
  // ═══ أعمدة الرحلات (AG-BI) ═══
  // ذهاب - الرحلة 1 (AG-AM) = أعمدة 33-39
  DEP1: { fn: 33, dd: 34, dt: 35, fr: 36, to: 37, ad: 38, at: 39 },
  // ذهاب - الرحلة 2 (AN-AT) = أعمدة 40-46
  DEP2: { fn: 40, dd: 41, dt: 42, fr: 43, to: 44, ad: 45, at: 46 },
  // عودة - الرحلة 1 (AU-BA) = أعمدة 47-53
  RET1: { fn: 47, dd: 48, dt: 49, fr: 50, to: 51, ad: 52, at: 53 },
  // عودة - الرحلة 2 (BB-BH) = أعمدة 54-60
  RET2: { fn: 54, dd: 55, dt: 56, fr: 57, to: 58, ad: 59, at: 60 },
  
  // إعدادات المعالجة
  BATCH_SIZE_PDF: 50,
  BATCH_SIZE_API: 25,
  DELAY_API: 2000,
  MAX_TIME: 4.5 * 60 * 1000,
  FOLDER_NAME: 'Ikram_Tickets'
};

var API = {
  KEY: '1d82a75bd2msh6da5259e10fbb77p1229ebjsne04353e515dd',
  HOST: 'aerodatabox.p.rapidapi.com'
};


// ==================== القائمة ====================

/**
 * يُستدعى من onOpen() في CompleteScript.gs
 * أضف هذا السطر داخل onOpen() هناك:
 *   addB2CFlightMenu();
 */
function addB2CFlightMenu() {
  SpreadsheetApp.getUi().createMenu('✈️ رحلات B2C')
    .addItem('🚀 تحديث شامل (PDF + API)', 'runFullUpdate')
    .addSeparator()
    .addItem('📄 استخراج PDF فقط', 'runPDFOnly')
    .addItem('🌐 إكمال الوصول API فقط', 'runAPIOnly')
    .addItem('🔧 إصلاح التصنيف', 'fixClassification')
    .addSeparator()
    .addItem('📊 حالة التقدم', 'showStatus')
    .addItem('⏹️ إيقاف الكل', 'stopAll')
    .addSeparator()
    .addItem('⏰ تفعيل التحديث التلقائي (10 دقائق)', 'enableAutoTrigger')
    .addItem('❌ إلغاء التحديث التلقائي', 'disableAutoTrigger')
    .addSeparator()
    .addItem('📧 إرسال تقرير الآن', 'sendDailyReport')
    .addItem('📧 تفعيل التقرير اليومي (7 صباحاً)', 'enableDailyReport')
    .addItem('🚫 إلغاء التقرير اليومي', 'disableDailyReport')
    .addToUi();
}


// ==================== بديل getUi الآمن ====================

function safeUI_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return {
      alert: function(msg) { Logger.log(msg); },
      ButtonSet: { YES_NO: 'YES_NO' },
      Button: { YES: 'YES' }
    };
  }
}

function toast_(msg, title, sec) {
  try {
    SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).toast(msg, title || '', sec || 5);
  } catch (e) {
    Logger.log((title || '') + ': ' + msg);
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🚀 الأمر الرئيسي — تحديث شامل                   ║
// ╚═══════════════════════════════════════════════════════════════╝

function runFullUpdate() {
  var ui = safeUI_();
  
  // إعداد PDF
  var newCount = setupPDFExtraction_();
  
  if (newCount > 0) {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('B2C_AUTO_API', 'true');
    
    ui.alert(
      '🚀 التحديث الشامل\n\n' +
      '📄 المرحلة 1: استخراج ' + newCount + ' تذكرة PDF\n' +
      '🌐 المرحلة 2: إكمال الوصول عبر API (تلقائياً)\n\n' +
      '⏳ جاري البدء...'
    );
    startPDFAuto_();
  } else {
    ui.alert('✅ لا توجد تذاكر جديدة\n\n🔄 بدء إكمال الوصول عبر API...');
    runAPIOnly();
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              📄 المرحلة 1: استخراج PDF                        ║
// ╚═══════════════════════════════════════════════════════════════╝

function runPDFOnly() {
  var newCount = setupPDFExtraction_();
  if (newCount > 0) {
    safeUI_().alert('📄 تذاكر جديدة: ' + newCount + '\n\n⏳ جاري البدء...');
    startPDFAuto_();
  } else {
    safeUI_().alert('✅ لا توجد تذاكر جديدة للاستخراج');
  }
}

/**
 * إعداد — يكتشف التذاكر الجديدة (التي ليس لها بيانات في AN = العمود 40)
 * @return {number} عدد التذاكر الجديدة
 */
function setupPDFExtraction_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) return 0;
  
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var ticketData = sheet.getRange(CONFIG.DATA_START, CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(CONFIG.DATA_START, CONFIG.DEP2.fn, numRows, 1).getValues();
  
  var tickets = {};
  var alreadyDone = 0;
  
  for (var i = 0; i < ticketData.length; i++) {
    var ticketNo = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    
    if (!ticketNo || link.indexOf('http') !== 0) continue;
    
    if (hasDep2) {
      alreadyDone++;
      continue;
    }
    
    if (!tickets[ticketNo]) {
      tickets[ticketNo] = link;
    }
  }
  
  var folder = getOrCreateFolder_();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('B2C_TICKETS', JSON.stringify(tickets));
  props.setProperty('B2C_PDF_DONE', '[]');
  props.setProperty('B2C_PDF_FAILED', '[]');
  props.setProperty('B2C_FOLDER', folder.getId());
  
  var newCount = Object.keys(tickets).length;
  Logger.log('📄 إعداد: جديد=' + newCount + ' | معبأ=' + alreadyDone);
  return newCount;
}

function startPDFAuto_() {
  stopPDFTrigger_();
  ScriptApp.newTrigger('processPDFBatch').timeBased().everyMinutes(1).create();
  toast_('🚀 PDF تلقائي — ' + CONFIG.BATCH_SIZE_PDF + ' تذكرة/دقيقة', '▶️', 10);
  processPDFBatch();
}

function processPDFBatch() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  var allTickets = JSON.parse(props.getProperty('B2C_TICKETS') || '{}');
  var done = JSON.parse(props.getProperty('B2C_PDF_DONE') || '[]');
  var failed = JSON.parse(props.getProperty('B2C_PDF_FAILED') || '[]');
  var folderId = props.getProperty('B2C_FOLDER');
  
  if (!folderId) return;
  
  var allKeys = Object.keys(allTickets);
  var remaining = allKeys.filter(function(k) { return done.indexOf(k) === -1 && failed.indexOf(k) === -1; });
  
  if (remaining.length === 0) {
    toast_('🎉 PDF اكتمل! نجح: ' + done.length + ' | فشل: ' + failed.length, '✅', 10);
    stopPDFTrigger_();
    
    if (props.getProperty('B2C_AUTO_API') === 'true') {
      props.deleteProperty('B2C_AUTO_API');
      toast_('🔄 المرحلة 2: إكمال الوصول عبر API...', '✈️', 10);
      Utilities.sleep(3000);
      runAPIOnly();
    }
    return;
  }
  
  // بناء خريطة: تذكرة → صفوف
  var lastRow = sheet.getLastRow();
  var ticketCol = sheet.getRange(CONFIG.DATA_START, CONFIG.COL_TICKET, lastRow - 1, 1).getValues();
  var ticketRowMap = {};
  for (var i = 0; i < ticketCol.length; i++) {
    var t = String(ticketCol[i][0] || '').trim();
    if (t) {
      if (!ticketRowMap[t]) ticketRowMap[t] = [];
      ticketRowMap[t].push(i + CONFIG.DATA_START);
    }
  }
  
  var folder = DriveApp.getFolderById(folderId);
  var batch = remaining.slice(0, CONFIG.BATCH_SIZE_PDF);
  var startTime = Date.now();
  var batchOK = 0, batchFail = 0;
  
  toast_('⏳ PDF: ' + batch.length + ' تذكرة (متبقي: ' + remaining.length + ')', '🔄', 30);
  
  for (var b = 0; b < batch.length; b++) {
    if (Date.now() - startTime > CONFIG.MAX_TIME) break;
    
    var ticketNo = batch[b];
    var pdfUrl = allTickets[ticketNo];
    
    try {
      // ═══ البحث في المجلد أولاً قبل التحميل ═══
      var blob = null;
      var existingFiles = folder.getFilesByName(ticketNo + '.pdf');
      
      if (existingFiles.hasNext()) {
        // موجود — استخدمه مباشرة
        blob = existingFiles.next().getBlob();
        Logger.log('📁 موجود: ' + ticketNo);
      } else {
        // غير موجود — حمّله
        blob = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true, followRedirects: true }).getBlob();
        blob.setName(ticketNo + '.pdf').setContentType('application/pdf');
        folder.createFile(blob);
        Logger.log('⬇️ تم تحميل: ' + ticketNo);
      }
      
      var text = extractTextFromPDF_(blob);
      if (!text) { failed.push(ticketNo); batchFail++; continue; }
      
      var flightData = parseTicketText_(text);
      if (!flightData || (!flightData.segments.departure.length && !flightData.segments.return.length)) {
        failed.push(ticketNo); batchFail++; continue;
      }
      
      var rows = ticketRowMap[ticketNo] || [];
      for (var r = 0; r < rows.length; r++) {
        writeFlightData_(sheet, rows[r], flightData);
      }
      
      done.push(ticketNo);
      batchOK++;
    } catch (e) {
      Logger.log('PDF error ' + ticketNo + ': ' + e.message);
      failed.push(ticketNo);
      batchFail++;
    }
    
    Utilities.sleep(500);
  }
  
  props.setProperty('B2C_PDF_DONE', JSON.stringify(done));
  props.setProperty('B2C_PDF_FAILED', JSON.stringify(failed));
  
  toast_('PDF: نجح ' + batchOK + ' | فشل ' + batchFail + ' | متبقي ' + (remaining.length - batch.length), 'نتيجة', 10);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              ✍️ كتابة بيانات الرحلات (المنطق الصحيح)         ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * كتابة بيانات الرحلات بالتصنيف الصحيح:
 *   مباشر  → DEP2 (AN-AT) + RET1 (AU-BA)
 *   ترانزيت → DEP1 (AG-AM) + DEP2 (AN-AT) + RET1 (AU-BA) + RET2 (BB-BH)
 */
function writeFlightData_(sheet, row, flightData) {
  var dep = flightData.segments.departure || [];
  var ret = flightData.segments.return || [];
  
  // ═══ PNR ═══
  if (flightData.pnr) {
    sheet.getRange(row, CONFIG.COL_PNR).setValue(flightData.pnr);
  }
  
  // ═══ ذهاب ═══
  if (dep.length === 1) {
    // مباشر → AN-AT فقط (DEP2)
    writeSegment_(sheet, row, CONFIG.DEP2, dep[0]);
    // تأكد أن AG-AM فارغ
    clearSegment_(sheet, row, CONFIG.DEP1);
  } else if (dep.length >= 2) {
    // ترانزيت → AG-AM (رحلة 1) + AN-AT (رحلة 2)
    writeSegment_(sheet, row, CONFIG.DEP1, dep[0]);
    writeSegment_(sheet, row, CONFIG.DEP2, dep[1]);
  }
  
  // ═══ عودة ═══
  if (ret.length === 1) {
    // مباشر → AU-BA فقط (RET1)
    writeSegment_(sheet, row, CONFIG.RET1, ret[0]);
    // تأكد أن BB-BH فارغ
    clearSegment_(sheet, row, CONFIG.RET2);
  } else if (ret.length >= 2) {
    // ترانزيت → AU-BA (رحلة 1) + BB-BH (رحلة 2)
    writeSegment_(sheet, row, CONFIG.RET1, ret[0]);
    writeSegment_(sheet, row, CONFIG.RET2, ret[1]);
  }
}

function writeSegment_(sheet, row, cols, seg) {
  sheet.getRange(row, cols.fn, 1, 7).setValues([[
    seg.flightNo || '', seg.depDate || '', seg.depTime || '',
    seg.from || '', seg.to || '', seg.arrDate || '', seg.arrTime || ''
  ]]);
}

function clearSegment_(sheet, row, cols) {
  var range = sheet.getRange(row, cols.fn, 1, 7);
  range.clearContent();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🌐 المرحلة 2: إكمال الوصول عبر API               ║
// ╚═══════════════════════════════════════════════════════════════╝

function runAPIOnly() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var ui = safeUI_();
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) { ui.alert('لا توجد بيانات'); return; }
  
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var allData = sheet.getRange(CONFIG.DATA_START, 1, numRows, 61).getValues();
  
  var segments = [
    { name: 'DEP1', cols: CONFIG.DEP1 },
    { name: 'DEP2', cols: CONFIG.DEP2 },
    { name: 'RET1', cols: CONFIG.RET1 },
    { name: 'RET2', cols: CONFIG.RET2 }
  ];
  
  var missing = [];
  var seen = {};
  
  for (var i = 0; i < numRows; i++) {
    for (var s = 0; s < segments.length; s++) {
      var seg = segments[s];
      var fn = cleanVal_(allData[i][seg.cols.fn - 1]);
      var dd = cleanDate_(allData[i][seg.cols.dd - 1]);
      var ad = cleanDate_(allData[i][seg.cols.ad - 1]);
      var at = cleanVal_(allData[i][seg.cols.at - 1]);
      
      if (!fn || !dd) continue;
      if (ad && at) continue; // مكتمل
      
      var apiNo = fn.replace(/-/g, '').toUpperCase();
      var key = apiNo + '|' + dd;
      
      if (!seen[key]) {
        seen[key] = { flightNo: apiNo, date: dd, rows: [] };
        missing.push(seen[key]);
      }
      seen[key].rows.push({ rowIdx: i, seg: seg });
    }
  }
  
  if (missing.length === 0) {
    ui.alert('✅ كل بيانات الرحلات مكتملة!');
    return;
  }
  
  var props = PropertiesService.getScriptProperties();
  props.setProperty('B2C_API_LIST', JSON.stringify(missing));
  props.setProperty('B2C_API_IDX', '0');
  props.setProperty('B2C_API_OK', '0');
  props.setProperty('B2C_API_FAIL', '0');
  props.setProperty('B2C_API_FILLED', '0');
  
  ui.alert(
    '🌐 إكمال الوصول\n\n' +
    '✈️ رحلات ناقصة: ' + missing.length + '\n' +
    '📦 دفعات: ' + Math.ceil(missing.length / CONFIG.BATCH_SIZE_API) + '\n\n' +
    '⏳ جاري البدء...'
  );
  
  stopAPITrigger_();
  ScriptApp.newTrigger('processAPIBatch').timeBased().everyMinutes(1).create();
  processAPIBatch();
}

function processAPIBatch() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  var missing = JSON.parse(props.getProperty('B2C_API_LIST') || '[]');
  var idx = parseInt(props.getProperty('B2C_API_IDX') || '0');
  var okCount = parseInt(props.getProperty('B2C_API_OK') || '0');
  var failCount = parseInt(props.getProperty('B2C_API_FAIL') || '0');
  var filledCount = parseInt(props.getProperty('B2C_API_FILLED') || '0');
  
  if (idx >= missing.length) {
    toast_('🎉 API اكتمل! نجح: ' + okCount + ' | خلايا: ' + filledCount, '✅', 10);
    stopAPITrigger_();
    return;
  }
  
  var endIdx = Math.min(idx + CONFIG.BATCH_SIZE_API, missing.length);
  var startTime = Date.now();
  
  // قراءة كل البيانات مرة واحدة
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var allData = sheet.getRange(CONFIG.DATA_START, 1, numRows, 61).getValues();
  
  toast_('🌐 API: ' + (idx + 1) + '-' + endIdx + '/' + missing.length, '✈️', 30);
  
  for (var b = idx; b < endIdx; b++) {
    if (Date.now() - startTime > CONFIG.MAX_TIME) { idx = b; break; }
    
    var item = missing[b];
    
    try {
      var data = fetchFlightAPI_(item.flightNo, item.date);
      
      if (data) {
        okCount++;
        for (var r = 0; r < item.rows.length; r++) {
          var ri = item.rows[r];
          var row = ri.rowIdx + CONFIG.DATA_START;
          var cols = ri.seg.cols;
          var updated = false;
          
          // قراءة الـ 7 خلايا مرة واحدة
          var current = sheet.getRange(row, cols.fn, 1, 7).getValues()[0];
          var newVals = current.slice();
          
          // ad = index 5, at = index 6, fr = index 3, to = index 4
          if (!cleanDate_(current[5]) && data.arrDate) { newVals[5] = data.arrDate; updated = true; }
          if (!cleanVal_(current[6]) && data.arrTime) { newVals[6] = data.arrTime; updated = true; }
          if (!cleanVal_(current[3]) && data.from) { newVals[3] = data.from; updated = true; }
          if (!cleanVal_(current[4]) && data.to) { newVals[4] = data.to; updated = true; }
          
          if (updated) {
            sheet.getRange(row, cols.fn, 1, 7).setValues([newVals]);
            filledCount++;
          }
        }
      } else {
        failCount++;
      }
    } catch (e) {
      Logger.log('API error ' + item.flightNo + ': ' + e.message);
      failCount++;
    }
    
    idx = b + 1;
    Utilities.sleep(CONFIG.DELAY_API);
  }
  
  props.setProperty('B2C_API_IDX', String(idx));
  props.setProperty('B2C_API_OK', String(okCount));
  props.setProperty('B2C_API_FAIL', String(failCount));
  props.setProperty('B2C_API_FILLED', String(filledCount));
  
  if (idx >= missing.length) {
    toast_('🎉 API اكتمل! نجح: ' + okCount + ' | خلايا: ' + filledCount, '✅', 10);
    stopAPITrigger_();
  }
}

function fetchFlightAPI_(flightNo, dateStr) {
  try {
    var url = 'https://aerodatabox.p.rapidapi.com/flights/number/' +
              flightNo + '/' + dateStr +
              '?withAircraftImage=false&withLocation=false&dateLocalRole=Both';
    
    var resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'X-RapidAPI-Key': API.KEY, 'X-RapidAPI-Host': API.HOST },
      muteHttpExceptions: true
    });
    
    var code = resp.getResponseCode();
    if (code === 429) {
      Utilities.sleep(5000);
      resp = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: { 'X-RapidAPI-Key': API.KEY, 'X-RapidAPI-Host': API.HOST },
        muteHttpExceptions: true
      });
      code = resp.getResponseCode();
    }
    if (code !== 200) return null;
    
    var data = JSON.parse(resp.getContentText());
    if (!data || data.length === 0) return null;
    
    var fl = data[0];
    var dep = fl.departure || {};
    var arr = fl.arrival || {};
    var depDT = (dep.scheduledTime || {}).local || (dep.scheduledTime || {}).utc || '';
    var arrDT = (arr.scheduledTime || {}).local || (arr.scheduledTime || {}).utc || '';
    
    return {
      from: (dep.airport || {}).iata || '',
      to: (arr.airport || {}).iata || '',
      depTime: extractTime_(depDT),
      arrTime: extractTime_(arrDT),
      depDate: extractDatePart_(depDT),
      arrDate: extractDatePart_(arrDT)
    };
  } catch (e) {
    Logger.log('fetchFlightAPI_ error: ' + e.message);
    return null;
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🔧 إصلاح التصنيف                                ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * إصلاح الصفوف المصنفة خطأ:
 *   1. مباشرة مكررة (AG = AN نفس الرحلة) → مسح AG، إبقاء AN
 *   2. مباشرة في AG بدون AN → نقل AG → AN
 */
function fixClassification() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var ui = safeUI_();
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) return;
  
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var allData = sheet.getRange(CONFIG.DATA_START, 1, numRows, 61).getValues();
  
  var fixedDup = 0;    // مكررة → مسح AG
  var fixedMove = 0;   // AG ممتلئ بدون AN → نقل AG→AN
  var fixedRetDup = 0; // عودة مكررة
  var fixedRetMove = 0;
  
  for (var i = 0; i < numRows; i++) {
    var row = i + CONFIG.DATA_START;
    
    // ═══ ذهاب ═══
    var dep1Fn = cleanVal_(allData[i][CONFIG.DEP1.fn - 1]);
    var dep2Fn = cleanVal_(allData[i][CONFIG.DEP2.fn - 1]);
    
    if (dep1Fn && dep2Fn) {
      // كلاهما ممتلئ — تحقق إذا مكرر
      if (dep1Fn.replace(/-/g, '') === dep2Fn.replace(/-/g, '')) {
        // مكرر → مسح DEP1 (AG-AM)
        clearSegment_(sheet, row, CONFIG.DEP1);
        fixedDup++;
      }
    } else if (dep1Fn && !dep2Fn) {
      // AG ممتلئ بدون AN → نقل AG→AN (كتلة واحدة)
      var dep1Vals = sheet.getRange(row, CONFIG.DEP1.fn, 1, 7).getValues()[0];
      sheet.getRange(row, CONFIG.DEP2.fn, 1, 7).setValues([dep1Vals]);
      clearSegment_(sheet, row, CONFIG.DEP1);
      fixedMove++;
    }
    
    // ═══ عودة ═══
    var ret1Fn = cleanVal_(allData[i][CONFIG.RET1.fn - 1]);
    var ret2Fn = cleanVal_(allData[i][CONFIG.RET2.fn - 1]);
    
    if (ret1Fn && ret2Fn) {
      if (ret1Fn.replace(/-/g, '') === ret2Fn.replace(/-/g, '')) {
        clearSegment_(sheet, row, CONFIG.RET2);
        fixedRetDup++;
      }
    }
    // العودة المباشرة في RET1 — هذا صحيح، لا تحتاج إصلاح
  }
  
  var total = fixedDup + fixedMove + fixedRetDup + fixedRetMove;
  
  ui.alert(
    '🔧 نتيجة الإصلاح\n━━━━━━━━━━━━━━━\n\n' +
    '📦 ذهاب — مكرر مُمسح (AG): ' + fixedDup + '\n' +
    '📦 ذهاب — منقول (AG→AN): ' + fixedMove + '\n' +
    '📦 عودة — مكرر مُمسح (BB): ' + fixedRetDup + '\n\n' +
    '✅ إجمالي الإصلاحات: ' + total
  );
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              ⏰ التحديث التلقائي كل 10 دقائق                  ║
// ╚═══════════════════════════════════════════════════════════════╝

function enableAutoTrigger() {
  disableAutoTrigger();
  ScriptApp.newTrigger('autoCheckNewRows')
    .timeBased()
    .everyMinutes(10)
    .create();
  safeUI_().alert('✅ تم تفعيل التحديث التلقائي\n\nسيفحص كل 10 دقائق إذا أُضيفت أسطر جديدة');
}

function disableAutoTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoCheckNewRows') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * يُشغّل تلقائياً كل 10 دقائق
 * يبحث عن أسطر جديدة (فيها تذكرة + رابط بدون بيانات رحلة)
 */
function autoCheckNewRows() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) return;
  
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var ticketData = sheet.getRange(CONFIG.DATA_START, CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(CONFIG.DATA_START, CONFIG.DEP2.fn, numRows, 1).getValues();
  
  var newTickets = {};
  
  for (var i = 0; i < ticketData.length; i++) {
    var ticketNo = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    
    if (!ticketNo || link.indexOf('http') !== 0) continue;
    if (hasDep2) continue;
    if (newTickets[ticketNo]) continue;
    
    newTickets[ticketNo] = link;
  }
  
  var count = Object.keys(newTickets).length;
  
  if (count === 0) {
    Logger.log('⏰ فحص تلقائي: لا توجد أسطر جديدة');
    return;
  }
  
  Logger.log('⏰ فحص تلقائي: وُجدت ' + count + ' تذكرة جديدة — بدء المعالجة');
  
  // إعداد وتشغيل
  var props = PropertiesService.getScriptProperties();
  var folder = getOrCreateFolder_();
  props.setProperty('B2C_TICKETS', JSON.stringify(newTickets));
  props.setProperty('B2C_PDF_DONE', '[]');
  props.setProperty('B2C_PDF_FAILED', '[]');
  props.setProperty('B2C_FOLDER', folder.getId());
  props.setProperty('B2C_AUTO_API', 'true');
  
  startPDFAuto_();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              📊 التقارير والتحكم                              ║
// ╚═══════════════════════════════════════════════════════════════╝

function showStatus() {
  var props = PropertiesService.getScriptProperties();
  var ui = safeUI_();
  
  // PDF
  var tickets = JSON.parse(props.getProperty('B2C_TICKETS') || '{}');
  var pdfDone = JSON.parse(props.getProperty('B2C_PDF_DONE') || '[]');
  var pdfFailed = JSON.parse(props.getProperty('B2C_PDF_FAILED') || '[]');
  var pdfTotal = Object.keys(tickets).length;
  
  // API
  var apiList = JSON.parse(props.getProperty('B2C_API_LIST') || '[]');
  var apiIdx = parseInt(props.getProperty('B2C_API_IDX') || '0');
  var apiOk = parseInt(props.getProperty('B2C_API_OK') || '0');
  var apiFail = parseInt(props.getProperty('B2C_API_FAIL') || '0');
  var apiFilled = parseInt(props.getProperty('B2C_API_FILLED') || '0');
  
  // إحصائيات الشيت
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var numRows = lastRow > 1 ? lastRow - 1 : 0;
  
  var withFlight = 0, noFlight = 0;
  if (numRows > 0) {
    var dep2 = sheet.getRange(2, CONFIG.DEP2.fn, numRows, 1).getValues();
    for (var i = 0; i < numRows; i++) {
      if (String(dep2[i][0] || '').trim()) withFlight++;
      else noFlight++;
    }
  }
  
  var report = '📊 حالة شيت B2C\n═══════════════════\n\n';
  report += '👤 إجمالي الصفوف: ' + numRows + '\n';
  report += '✅ مع بيانات رحلة: ' + withFlight + '\n';
  report += '❓ بدون رحلة: ' + noFlight + '\n\n';
  
  report += '📄 PDF:\n';
  report += '  نجح: ' + pdfDone.length + ' | فشل: ' + pdfFailed.length + ' | إجمالي: ' + pdfTotal + '\n\n';
  
  report += '🌐 API:\n';
  report += '  نجح: ' + apiOk + ' | فشل: ' + apiFail + '\n';
  report += '  خلايا مُكتملة: ' + apiFilled + '\n';
  report += '  تقدم: ' + apiIdx + '/' + apiList.length + '\n';
  
  // Triggers
  var triggers = ScriptApp.getProjectTriggers();
  var activeT = [];
  for (var t = 0; t < triggers.length; t++) {
    activeT.push(triggers[t].getHandlerFunction());
  }
  report += '\n⏰ Triggers نشطة: ' + (activeT.length > 0 ? activeT.join(', ') : 'لا يوجد');
  
  ui.alert(report);
}

function stopAll() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'processPDFBatch' || fn === 'processAPIBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  toast_('⏹️ تم إيقاف جميع العمليات', 'إيقاف', 5);
}

function stopPDFTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processPDFBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function stopAPITrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processAPIBatch') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              📧 تقرير يومي بالإيميل                           ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * إرسال تقرير يومي صباحي بالرحلات الناقصة والقادمة
 */
function sendDailyReport() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START) return;
  
  var numRows = lastRow - CONFIG.DATA_START + 1;
  var allData = sheet.getRange(CONFIG.DATA_START, 1, numRows, 61).getValues();
  
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Riyadh', 'yyyy-MM-dd');
  
  // إحصائيات عامة
  var total = 0, withFlight = 0, noFlight = 0, withPNR = 0;
  
  // رحلات ناقصة الوصول (AS أو AT فارغ)
  var missingArrival = [];
  
  // رحلات ناقصة المغادرة (AU أو AV أو AW فارغ)
  var missingDeparture = [];
  
  // رحلات قادمة خلال 3 أيام
  var upcoming3Days = [];
  
  // مغادرة خلال 3 أيام
  var departing3Days = [];
  
  var threeDaysLater = new Date(today.getTime() + 3 * 24 * 60 * 60 * 1000);
  var threeDaysStr = Utilities.formatDate(threeDaysLater, 'Asia/Riyadh', 'yyyy-MM-dd');
  
  for (var i = 0; i < numRows; i++) {
    var passport = cleanVal_(allData[i][CONFIG.COL_PASSPORT - 1]);
    if (!passport) continue;
    total++;
    
    var dep2Fn = cleanVal_(allData[i][CONFIG.DEP2.fn - 1]);  // AN
    var arrDate = cleanDate_(allData[i][CONFIG.DEP2.ad - 1]); // AS
    var arrTime = cleanVal_(allData[i][CONFIG.DEP2.at - 1]);  // AT
    
    var ret1Fn = cleanVal_(allData[i][CONFIG.RET1.fn - 1]);   // AU
    var retDate = cleanDate_(allData[i][CONFIG.RET1.dd - 1]);  // AV
    var retTime = cleanVal_(allData[i][CONFIG.RET1.dt - 1]);   // AW
    
    var pnr = cleanVal_(allData[i][CONFIG.COL_PNR - 1]);      // BI
    var name = cleanVal_(allData[i][10]) + ' ' + cleanVal_(allData[i][11]); // اسم إنجليزي K+L
    
    if (dep2Fn) {
      withFlight++;
    } else {
      noFlight++;
      continue;
    }
    
    if (pnr) withPNR++;
    
    // فحص الناقص — وصول
    if (dep2Fn && (!arrDate || !arrTime)) {
      missingArrival.push({
        name: name,
        flight: dep2Fn,
        date: cleanDate_(allData[i][CONFIG.DEP2.dd - 1]),
        arrDate: arrDate || '❌',
        arrTime: arrTime || '❌'
      });
    }
    
    // فحص الناقص — مغادرة
    if (ret1Fn && (!retDate || !retTime)) {
      missingDeparture.push({
        name: name,
        flight: ret1Fn,
        retDate: retDate || '❌',
        retTime: retTime || '❌'
      });
    }
    
    // رحلات قادمة — وصول خلال 3 أيام
    if (arrDate && arrDate >= todayStr && arrDate <= threeDaysStr) {
      upcoming3Days.push({
        name: name,
        flight: dep2Fn,
        date: arrDate,
        time: arrTime || '⏳ غير محدد',
        pnr: pnr
      });
    }
    
    // رحلات قادمة — مغادرة خلال 3 أيام
    if (retDate && retDate >= todayStr && retDate <= threeDaysStr) {
      departing3Days.push({
        name: name,
        flight: ret1Fn,
        date: retDate,
        time: retTime || '⏳ غير محدد',
        pnr: pnr
      });
    }
  }
  
  // ═══ بناء الإيميل ═══
  var subject = '📊 تقرير رحلات B2C — ' + todayStr;
  
  var html = '<div dir="rtl" style="font-family: Arial, sans-serif; max-width: 700px;">';
  
  // الملخص
  html += '<h2 style="color: #1a73e8;">📊 تقرير رحلات B2C اليومي</h2>';
  html += '<p style="color: #666;">التاريخ: ' + todayStr + ' | التوقيت: توقيت مكة المكرمة</p>';
  
  html += '<table style="border-collapse: collapse; width: 100%; margin: 10px 0;">';
  html += '<tr style="background: #1a73e8; color: white;">';
  html += '<td style="padding: 8px;">إجمالي الحجاج</td>';
  html += '<td style="padding: 8px;">مع رحلة</td>';
  html += '<td style="padding: 8px;">بدون رحلة</td>';
  html += '<td style="padding: 8px;">مع PNR</td></tr>';
  html += '<tr style="background: #f8f9fa;">';
  html += '<td style="padding: 8px; font-weight: bold;">' + total + '</td>';
  html += '<td style="padding: 8px; color: green;">' + withFlight + '</td>';
  html += '<td style="padding: 8px; color: red;">' + noFlight + '</td>';
  html += '<td style="padding: 8px;">' + withPNR + '</td></tr>';
  html += '</table>';
  
  // وصول خلال 3 أيام
  if (upcoming3Days.length > 0) {
    html += '<h3 style="color: #0d652d;">🛬 وصول خلال 3 أيام (' + upcoming3Days.length + ' حاج)</h3>';
    html += _buildFlightTable(upcoming3Days, 'وصول');
  }
  
  // مغادرة خلال 3 أيام
  if (departing3Days.length > 0) {
    html += '<h3 style="color: #b45309;">🛫 مغادرة خلال 3 أيام (' + departing3Days.length + ' حاج)</h3>';
    html += _buildFlightTable(departing3Days, 'مغادرة');
  }
  
  // ناقص الوصول — مجمّع حسب الرحلة
  if (missingArrival.length > 0) {
    html += '<h3 style="color: #c62828;">⚠️ ناقص بيانات الوصول (' + missingArrival.length + ' حاج)</h3>';
    var arrGroups = _groupByFlight(missingArrival, 'flight', 'date');
    html += '<table style="border-collapse: collapse; width: 100%; font-size: 13px;">';
    html += '<tr style="background: #ffebee;"><th style="padding: 6px; border: 1px solid #ddd;">الرحلة</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">تاريخ الإقلاع</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">تاريخ الوصول</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">وقت الوصول</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">عدد الحجاج</th></tr>';
    var arrKeys = Object.keys(arrGroups).sort();
    for (var a = 0; a < arrKeys.length; a++) {
      var ag = arrGroups[arrKeys[a]];
      html += '<tr>';
      html += '<td style="padding: 5px; border: 1px solid #eee; font-weight: bold;">' + ag.flight + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + ag.date + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + ag.arrDate + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + ag.arrTime + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + ag.count + '</td></tr>';
    }
    html += '</table>';
  }
  
  // ناقص المغادرة — مجمّع حسب الرحلة
  if (missingDeparture.length > 0) {
    html += '<h3 style="color: #e65100;">⚠️ ناقص بيانات المغادرة (' + missingDeparture.length + ' حاج)</h3>';
    var depGroups = _groupByFlight(missingDeparture, 'flight', 'retDate');
    html += '<table style="border-collapse: collapse; width: 100%; font-size: 13px;">';
    html += '<tr style="background: #fff3e0;"><th style="padding: 6px; border: 1px solid #ddd;">الرحلة</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">التاريخ</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">الوقت</th>';
    html += '<th style="padding: 6px; border: 1px solid #ddd;">عدد الحجاج</th></tr>';
    var depKeys = Object.keys(depGroups).sort();
    for (var d = 0; d < depKeys.length; d++) {
      var dg = depGroups[depKeys[d]];
      html += '<tr>';
      html += '<td style="padding: 5px; border: 1px solid #eee; font-weight: bold;">' + dg.flight + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + dg.retDate + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + dg.retTime + '</td>';
      html += '<td style="padding: 5px; border: 1px solid #eee;">' + dg.count + '</td></tr>';
    }
    html += '</table>';
  }
  
  // ذيل
  html += '<hr style="margin: 20px 0; border-color: #eee;">';
  html += '<p style="color: #999; font-size: 12px;">تقرير تلقائي من نظام إكرام الضيف — شيت B2C</p>';
  html += '</div>';
  
  // إرسال
  var email = Session.getActiveUser().getEmail();
  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: html
  });
  
  Logger.log('📧 تم إرسال التقرير إلى: ' + email);
}

/**
 * بناء جدول رحلات قادمة
 */
function _buildFlightTable(flights, type) {
  // تجميع حسب الرحلة + التاريخ
  var groups = {};
  for (var i = 0; i < flights.length; i++) {
    var f = flights[i];
    var key = f.flight + '|' + f.date + '|' + f.time;
    if (!groups[key]) {
      groups[key] = { flight: f.flight, date: f.date, time: f.time, count: 0, names: [] };
    }
    groups[key].count++;
    if (groups[key].names.length < 3) groups[key].names.push(f.name);
  }
  
  var html = '<table style="border-collapse: collapse; width: 100%; font-size: 13px;">';
  html += '<tr style="background: #e8f5e9;"><th style="padding: 6px; border: 1px solid #ddd;">الرحلة</th>';
  html += '<th style="padding: 6px; border: 1px solid #ddd;">التاريخ</th>';
  html += '<th style="padding: 6px; border: 1px solid #ddd;">الوقت</th>';
  html += '<th style="padding: 6px; border: 1px solid #ddd;">عدد الحجاج</th></tr>';
  
  var keys = Object.keys(groups).sort();
  for (var k = 0; k < keys.length; k++) {
    var g = groups[keys[k]];
    html += '<tr>';
    html += '<td style="padding: 5px; border: 1px solid #eee; font-weight: bold;">' + g.flight + '</td>';
    html += '<td style="padding: 5px; border: 1px solid #eee;">' + g.date + '</td>';
    html += '<td style="padding: 5px; border: 1px solid #eee;">' + g.time + '</td>';
    html += '<td style="padding: 5px; border: 1px solid #eee;">' + g.count + ' حاج</td>';
    html += '</tr>';
  }
  html += '</table>';
  return html;
}

/**
 * تجميع حسب الرحلة + التاريخ — يحتفظ بأول قيمة لكل حقل
 */
function _groupByFlight(items, flightField, dateField) {
  var groups = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var key = item[flightField] + '|' + (item[dateField] || '');
    if (!groups[key]) {
      groups[key] = {};
      // نسخ كل الحقول من أول عنصر
      for (var prop in item) {
        if (item.hasOwnProperty(prop) && prop !== 'name') {
          groups[key][prop] = item[prop];
        }
      }
      groups[key].count = 0;
    }
    groups[key].count++;
  }
  return groups;
}
/**
 * تفعيل التقرير اليومي — كل يوم الساعة 7 صباحاً بتوقيت مكة
 */
function enableDailyReport() {
  disableDailyReport();
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .inTimezone('Asia/Riyadh')
    .create();
  safeUI_().alert('✅ تم تفعيل التقرير اليومي\n\nسيُرسل كل يوم الساعة 7 صباحاً بتوقيت مكة');
}

function disableDailyReport() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🔨 دوال مساعدة                                   ║
// ╚═══════════════════════════════════════════════════════════════╝

function extractTextFromPDF_(blob) {
  var docId = null;
  try {
    var metadata = { name: 'temp_b2c_' + Date.now(), mimeType: 'application/vnd.google-apps.document' };
    var boundary = '-----b2c' + Date.now();
    var payload = Utilities.newBlob(
      '--' + boundary + '\r\n' +
      'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) + '\r\n' +
      '--' + boundary + '\r\n' +
      'Content-Type: application/pdf\r\n' +
      'Content-Transfer-Encoding: base64\r\n\r\n' +
      Utilities.base64Encode(blob.getBytes()) + '\r\n' +
      '--' + boundary + '--'
    ).getBytes();
    
    var response = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
      {
        method: 'post',
        contentType: 'multipart/related; boundary=' + boundary,
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: payload,
        muteHttpExceptions: true
      }
    );
    
    if (response.getResponseCode() !== 200) return null;
    
    var file = JSON.parse(response.getContentText());
    docId = file.id;
    var text = DocumentApp.openById(docId).getBody().getText();
    DriveApp.getFileById(docId).setTrashed(true);
    return text;
  } catch (e) {
    if (docId) try { DriveApp.getFileById(docId).setTrashed(true); } catch (x) {}
    return null;
  }
}

function parseTicketText_(text) {
  var result = { pnr: '', segments: { departure: [], return: [] } };
  
  // PNR
  var pnrMatch = text.match(/(?:Airline Reference|PNR|Airline PNR)\s*[\(\):\n\r\s]*([A-Z0-9]{5,7})/i) ||
                 text.match(/Airline Reference\s*\(PNR\)\s*[\n\r]+\s*([A-Z0-9]{5,7})/i) ||
                 text.match(/PNR[:\s]*[\n\r]+\s*([A-Z0-9]{5,7})/i);
  if (pnrMatch) result.pnr = pnrMatch[1].trim();
  
  // تقسيم ذهاب/عودة
  var retIndex = -1;
  var retPatterns = [/Return from/i, /العودة من/i];
  for (var p = 0; p < retPatterns.length; p++) {
    var rm = text.search(retPatterns[p]);
    if (rm > 0) { retIndex = rm; break; }
  }
  
  var depText, retText;
  if (retIndex > 0) {
    depText = text.substring(0, retIndex);
    retText = text.substring(retIndex);
  } else {
    depText = text;
    retText = '';
  }
  
  var depSegs = extractFlightSegments_(depText);
  var retSegs = extractFlightSegments_(retText);
  
  if (retSegs.length === 0 && depSegs.length >= 2) {
    var half = Math.floor(depSegs.length / 2);
    retSegs = depSegs.slice(half);
    depSegs = depSegs.slice(0, half);
  }
  
  result.segments.departure = dedup_(depSegs).slice(0, 2);
  result.segments.return = dedup_(retSegs).slice(0, 2);
  
  return result;
}

function dedup_(segments) {
  var seen = {};
  var unique = [];
  for (var i = 0; i < segments.length; i++) {
    var key = segments[i].flightNo + '|' + segments[i].from + '|' + segments[i].to;
    if (!seen[key]) { seen[key] = true; unique.push(segments[i]); }
  }
  return unique;
}

function extractFlightSegments_(sectionText) {
  if (!sectionText || sectionText.length < 20) return [];
  
  var segments = [];
  var flightRegex = /\b([A-Z]{2})[-\s]?(\d{2,4})\b/g;
  var skipPrefixes = ['AM','PM','ID','OK','NO','OF','IN','TO','UP','DO','IF','OR','AN','AT','ON','BY'];
  var flightMatches = [];
  var m;
  
  while ((m = flightRegex.exec(sectionText)) !== null) {
    if (skipPrefixes.indexOf(m[1]) !== -1) continue;
    flightMatches.push({ number: m[1] + '-' + m[2], index: m.index, end: m.index + m[0].length });
  }
  
  for (var i = 0; i < flightMatches.length; i++) {
    var flight = flightMatches[i];
    var prevStart = i === 0 ? 0 : flightMatches[i - 1].end;
    var beforeText = sectionText.substring(prevStart, flight.index);
    var nextEnd = i < flightMatches.length - 1 ? flightMatches[i + 1].index : sectionText.length;
    var afterText = sectionText.substring(flight.end, Math.min(flight.end + 500, nextEnd));
    
    var depInfo = extractDateTimeAirport_(beforeText, true);
    var arrInfo = extractDateTimeAirport_(afterText, false);
    
    if (!arrInfo.time && afterText.length > 0) {
      arrInfo = extractDateTimeAirport_(sectionText.substring(flight.end, Math.min(flight.end + 200, sectionText.length)), false);
    }
    
    segments.push({
      flightNo: flight.number,
      depDate: depInfo.date || '',
      depTime: depInfo.time || '',
      from: depInfo.airport || '',
      arrDate: arrInfo.date || '',
      arrTime: arrInfo.time || '',
      to: arrInfo.airport || ''
    });
  }
  return segments;
}

function extractDateTimeAirport_(text, isBefore) {
  var result = { date: '', time: '', airport: '' };
  
  var dates = [], times = [], airports = [];
  var dm, tm, am;
  
  var dateRegex = /((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\.?\s+\d{1,2},?\s+\d{4})/gi;
  while ((dm = dateRegex.exec(text)) !== null) dates.push(dm[1].trim());
  
  var timeRegex = /(\d{1,2}:\d{2}\s*(?:AM|PM))/gi;
  while ((tm = timeRegex.exec(text)) !== null) times.push(tm[1].trim());
  
  var airportRegex = /\b([A-Z]{3})\b/g;
  var skipWords = ['THE','AND','FOR','NOT','ARE','WAS','HAS','HAD','BUT','HIS','HER','ITS','OUR',
    'MAY','JAN','FEB','MAR','APR','JUN','JUL','AUG','SEP','OCT','NOV','DEC','YOU','ALL','CAN',
    'NEW','ONE','TWO','DAY','GET','USE','PER','ANY','HOW','BIN','PDF','SAR','USD','EUR','GBP',
    'CHF','KGM','BAG','MIN','MAX','REF','VIA','FEE','TAX','ADT','CHD','INF','PNR','MRS','MIS',
    'OWN','SET','ADD','END','WAY','AIR','FLY','RUN','SIT','ROW','NET','PAX','QTY','NOS','STD',
    'ETD','ETA','STA'];
  while ((am = airportRegex.exec(text)) !== null) {
    if (skipWords.indexOf(am[1]) === -1) airports.push(am[1]);
  }
  
  if (isBefore) {
    result.date = dates.length > 0 ? dates[dates.length - 1] : '';
    result.time = times.length > 0 ? times[times.length - 1] : '';
    result.airport = airports.length > 0 ? airports[airports.length - 1] : '';
  } else {
    result.date = dates.length > 0 ? dates[0] : '';
    result.time = times.length > 0 ? times[0] : '';
    result.airport = airports.length > 0 ? airports[0] : '';
  }
  
  if (result.date) result.date = formatParsedDate_(result.date);
  if (result.time) result.time = convertTo24h_(result.time);
  
  return result;
}

function formatParsedDate_(dateStr) {
  try {
    var months = { 'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
                   'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12' };
    var parts = dateStr.replace(',', '').trim().split(/\s+/);
    if (parts.length >= 3) {
      return parts[2] + '-' + (months[parts[0].substring(0, 3).toLowerCase()] || '01') + '-' + ('0' + parseInt(parts[1])).slice(-2);
    }
    return dateStr;
  } catch (e) { return dateStr; }
}

function convertTo24h_(timeStr) {
  try {
    var match = timeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    if (!match) return timeStr;
    var h = parseInt(match[1]);
    if (match[3].toUpperCase() === 'PM' && h !== 12) h += 12;
    if (match[3].toUpperCase() === 'AM' && h === 12) h = 0;
    return ('0' + h).slice(-2) + ':' + match[2];
  } catch (e) { return timeStr; }
}

function extractTime_(dt) {
  if (!dt) return '';
  var m = String(dt).match(/[\sT](\d{2}:\d{2})/);
  return m ? m[1] : '';
}

function extractDatePart_(dt) {
  if (!dt) return '';
  var m = String(dt).match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : '';
}

function cleanVal_(val) {
  if (!val && val !== 0) return '';
  return String(val).trim();
}

function cleanDate_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (val.getFullYear() < 2000) return '';
    return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2) + '-' + ('0' + val.getDate()).slice(-2);
  }
  var str = String(val).trim();
  var m = str.match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : str;
}

function getOrCreateFolder_() {
  var folders = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
}