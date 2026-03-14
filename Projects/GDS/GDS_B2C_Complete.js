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
 * هيكل أعمدة B2C (4 رحلات + PNR):
 *   DEP1 (AG-AM) أعمدة 33-39 = قدوم رحلة 1
 *   DEP2 (AN-AT) أعمدة 40-46 = قدوم رحلة 2
 *   RET1 (AU-BA) أعمدة 47-53 = عودة رحلة 1
 *   RET2 (BB-BH) أعمدة 54-60 = عودة رحلة 2
 *   PNR  (BI)    عمود 61     = رقم حجز الطيران
 *
 * التصنيف:
 *   مباشر  → DEP2 + RET1 (DEP1 و RET2 فارغين)
 *   ترانزيت → DEP1 + DEP2 + RET1 + RET2
 * 
 * الأوامر:
 *   من القائمة: ✈️ بيانات رحلات B2C
 *   أو تلقائياً كل 10 دقائق
 * ══════════════════════════════════════════════════════════════════
 */


// ==================== الإعدادات ====================

var B2C_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'B2C',
  DATA_START: 2,

  // أعمدة مرجعية
  COL_PASSPORT: 6,     // F — رقم الجواز
  COL_TICKET: 25,      // Y — رقم التذكرة
  COL_LINK: 26,        // Z — رابط التذكرة PDF
  COL_PNR: 61,         // BI — Airline PNR

  // ═══ أعمدة الرحلات (AG-BH) — 4 رحلات ═══
  DEP1: { fn: 33, dd: 34, dt: 35, fr: 36, to: 37, ad: 38, at: 39 },
  DEP2: { fn: 40, dd: 41, dt: 42, fr: 43, to: 44, ad: 45, at: 46 },
  RET1: { fn: 47, dd: 48, dt: 49, fr: 50, to: 51, ad: 52, at: 53 },
  RET2: { fn: 54, dd: 55, dt: 56, fr: 57, to: 58, ad: 59, at: 60 },

  // إعدادات المعالجة
  BATCH_SIZE_PDF: 25,
  BATCH_SIZE_API: 20,
  DELAY_API: 2000,
  MAX_TIME: 4.5 * 60 * 1000,
  FOLDER_NAME: 'Ikram_Tickets'
};

var API = {
  KEY: '1d82a75bd2msh6da5259e10fbb77p1229ebjsne04353e515dd',
  HOST: 'aerodatabox.p.rapidapi.com'
};


// ==================== القائمة ====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('✈️ بيانات رحلات B2C')
    .addItem('🚀 تحديث شامل (PDF + API)', 'runFullUpdate')
    .addSeparator()
    .addItem('📄 استخراج PDF فقط', 'runPDFOnly')
    .addItem('🌐 إكمال الوصول API فقط', 'runAPIOnly')
    .addItem('🔧 إصلاح التصنيف', 'fixClassification')
    .addSeparator()
    .addItem('📊 حالة التقدم', 'showStatus')
    .addItem('🔍 تشخيص', 'diagnosePDFSetup')
    .addItem('⏹️ إيقاف الكل', 'stopAll')
    .addSeparator()
    .addItem('⏰ تفعيل التحديث التلقائي (10 دقائق)', 'enableAutoTrigger')
    .addItem('❌ إلغاء التحديث التلقائي', 'disableAutoTrigger')
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
    SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID).toast(msg, title || '', sec || 5);
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
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log('شيت ' + B2C_CONFIG.SHEET_NAME + ' غير موجود'); return 0; }
  var lastRow = sheet.getLastRow();
  if (lastRow < B2C_CONFIG.DATA_START) return 0;
  
  var numRows = lastRow - B2C_CONFIG.DATA_START + 1;
  var ticketData = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.DEP2.fn, numRows, 1).getValues();

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
  toast_('🚀 PDF تلقائي — ' + B2C_CONFIG.BATCH_SIZE_PDF + ' تذكرة/دقيقة', '▶️', 10);
  processPDFBatch();
}

function processPDFBatch() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log('شيت ' + B2C_CONFIG.SHEET_NAME + ' غير موجود'); return; }

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
  var ticketCol = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.COL_TICKET, lastRow - 1, 1).getValues();
  var ticketRowMap = {};
  for (var i = 0; i < ticketCol.length; i++) {
    var t = String(ticketCol[i][0] || '').trim();
    if (t) {
      if (!ticketRowMap[t]) ticketRowMap[t] = [];
      ticketRowMap[t].push(i + B2C_CONFIG.DATA_START);
    }
  }
  
  var folder = DriveApp.getFolderById(folderId);
  var batch = remaining.slice(0, B2C_CONFIG.BATCH_SIZE_PDF);
  var startTime = Date.now();
  var batchOK = 0, batchFail = 0;
  
  toast_('⏳ PDF: ' + batch.length + ' تذكرة (متبقي: ' + remaining.length + ')', '🔄', 30);
  
  for (var b = 0; b < batch.length; b++) {
    if (Date.now() - startTime > B2C_CONFIG.MAX_TIME) break;
    
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
 *   مباشر  → DEP2 (AN-AT) + RET1 (AU-BA) — DEP1 و RET2 فارغين
 *   ترانزيت → DEP1 (AG-AM) + DEP2 (AN-AT) + RET1 (AU-BA) + RET2 (BB-BH)
 */
function writeFlightData_(sheet, row, flightData) {
  var dep = flightData.segments.departure || [];
  var ret = flightData.segments.return || [];

  // ═══ PNR ═══
  if (flightData.pnr) {
    sheet.getRange(row, B2C_CONFIG.COL_PNR).setValue(flightData.pnr);
  }

  // ═══ قدوم ═══
  if (dep.length === 1) {
    writeSegment_(sheet, row, B2C_CONFIG.DEP2, dep[0]);
    clearSegment_(sheet, row, B2C_CONFIG.DEP1);
  } else if (dep.length >= 2) {
    writeSegment_(sheet, row, B2C_CONFIG.DEP1, dep[0]);
    writeSegment_(sheet, row, B2C_CONFIG.DEP2, dep[1]);
  }

  // ═══ عودة ═══
  if (ret.length === 1) {
    writeSegment_(sheet, row, B2C_CONFIG.RET1, ret[0]);
    clearSegment_(sheet, row, B2C_CONFIG.RET2);
  } else if (ret.length >= 2) {
    writeSegment_(sheet, row, B2C_CONFIG.RET1, ret[0]);
    writeSegment_(sheet, row, B2C_CONFIG.RET2, ret[1]);
  }
}

function writeSegment_(sheet, row, cols, seg) {
  sheet.getRange(row, cols.fn).setValue(seg.flightNo || '');
  sheet.getRange(row, cols.dd).setValue(seg.depDate || '');
  sheet.getRange(row, cols.dt).setValue(seg.depTime || '');
  sheet.getRange(row, cols.fr).setValue(seg.from || '');
  sheet.getRange(row, cols.to).setValue(seg.to || '');
  sheet.getRange(row, cols.ad).setValue(seg.arrDate || '');
  sheet.getRange(row, cols.at).setValue(seg.arrTime || '');
}

function clearSegment_(sheet, row, cols) {
  var range = sheet.getRange(row, cols.fn, 1, 7);
  range.clearContent();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🌐 المرحلة 2: إكمال الوصول عبر API               ║
// ╚═══════════════════════════════════════════════════════════════╝

function runAPIOnly() {
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  var ui = safeUI_();
  if (!sheet) { ui.alert('شيت ' + B2C_CONFIG.SHEET_NAME + ' غير موجود'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < B2C_CONFIG.DATA_START) { ui.alert('لا توجد بيانات'); return; }
  
  var numRows = lastRow - B2C_CONFIG.DATA_START + 1;
  var allData = sheet.getRange(B2C_CONFIG.DATA_START, 1, numRows, 61).getValues();
  
  var segments = [
    { name: 'DEP1', cols: B2C_CONFIG.DEP1 },
    { name: 'DEP2', cols: B2C_CONFIG.DEP2 },
    { name: 'RET1', cols: B2C_CONFIG.RET1 },
    { name: 'RET2', cols: B2C_CONFIG.RET2 }
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
    '📦 دفعات: ' + Math.ceil(missing.length / B2C_CONFIG.BATCH_SIZE_API) + '\n\n' +
    '⏳ جاري البدء...'
  );
  
  stopAPITrigger_();
  ScriptApp.newTrigger('processAPIBatch').timeBased().everyMinutes(1).create();
  processAPIBatch();
}

function processAPIBatch() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log('شيت ' + B2C_CONFIG.SHEET_NAME + ' غير موجود'); return; }

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
  
  var endIdx = Math.min(idx + B2C_CONFIG.BATCH_SIZE_API, missing.length);
  var startTime = Date.now();
  
  // قراءة كل البيانات مرة واحدة
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - B2C_CONFIG.DATA_START + 1;
  var allData = sheet.getRange(B2C_CONFIG.DATA_START, 1, numRows, 61).getValues();
  
  toast_('🌐 API: ' + (idx + 1) + '-' + endIdx + '/' + missing.length, '✈️', 30);
  
  for (var b = idx; b < endIdx; b++) {
    if (Date.now() - startTime > B2C_CONFIG.MAX_TIME) { idx = b; break; }
    
    var item = missing[b];
    
    try {
      var data = fetchFlightAPI_(item.flightNo, item.date);
      
      if (data) {
        okCount++;
        // ملء كل الصفوف المرتبطة
        for (var r = 0; r < item.rows.length; r++) {
          var ri = item.rows[r];
          var row = ri.rowIdx + B2C_CONFIG.DATA_START;
          var cols = ri.seg.cols;
          var updated = false;
          
          if (!cleanDate_(allData[ri.rowIdx][cols.ad - 1]) && data.arrDate) {
            sheet.getRange(row, cols.ad).setValue(data.arrDate);
            updated = true;
          }
          if (!cleanVal_(allData[ri.rowIdx][cols.at - 1]) && data.arrTime) {
            sheet.getRange(row, cols.at).setValue(data.arrTime);
            updated = true;
          }
          if (!cleanVal_(allData[ri.rowIdx][cols.fr - 1]) && data.from) {
            sheet.getRange(row, cols.fr).setValue(data.from);
            updated = true;
          }
          if (!cleanVal_(allData[ri.rowIdx][cols.to - 1]) && data.to) {
            sheet.getRange(row, cols.to).setValue(data.to);
            updated = true;
          }
          if (updated) filledCount++;
        }
      } else {
        failCount++;
      }
    } catch (e) {
      Logger.log('API error ' + item.flightNo + ': ' + e.message);
      failCount++;
    }
    
    idx = b + 1;
    Utilities.sleep(B2C_CONFIG.DELAY_API);
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
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  var ui = safeUI_();
  if (!sheet) { ui.alert('شيت B2C غير موجود'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < B2C_CONFIG.DATA_START) return;
  
  var numRows = lastRow - B2C_CONFIG.DATA_START + 1;
  var allData = sheet.getRange(B2C_CONFIG.DATA_START, 1, numRows, 61).getValues();

  var fixedDup = 0;
  var fixedMove = 0;
  var fixedRetDup = 0;

  for (var i = 0; i < numRows; i++) {
    var row = i + B2C_CONFIG.DATA_START;

    var dep1Fn = cleanVal_(allData[i][B2C_CONFIG.DEP1.fn - 1]);
    var dep2Fn = cleanVal_(allData[i][B2C_CONFIG.DEP2.fn - 1]);

    if (dep1Fn && dep2Fn) {
      if (dep1Fn.replace(/-/g, '') === dep2Fn.replace(/-/g, '')) {
        clearSegment_(sheet, row, B2C_CONFIG.DEP1);
        fixedDup++;
      }
    } else if (dep1Fn && !dep2Fn) {
      var dep1Vals = sheet.getRange(row, B2C_CONFIG.DEP1.fn, 1, 7).getValues()[0];
      sheet.getRange(row, B2C_CONFIG.DEP2.fn, 1, 7).setValues([dep1Vals]);
      clearSegment_(sheet, row, B2C_CONFIG.DEP1);
      fixedMove++;
    }

    var ret1Fn = cleanVal_(allData[i][B2C_CONFIG.RET1.fn - 1]);
    var ret2Fn = cleanVal_(allData[i][B2C_CONFIG.RET2.fn - 1]);

    if (ret1Fn && ret2Fn) {
      if (ret1Fn.replace(/-/g, '') === ret2Fn.replace(/-/g, '')) {
        clearSegment_(sheet, row, B2C_CONFIG.RET2);
        fixedRetDup++;
      }
    }
  }

  var total = fixedDup + fixedMove + fixedRetDup;
  ui.alert(
    '🔧 نتيجة الإصلاح\n━━━━━━━━━━━━━━━\n\n' +
    '📦 قدوم — مكرر مُمسح (DEP1): ' + fixedDup + '\n' +
    '📦 قدوم — منقول (DEP1→DEP2): ' + fixedMove + '\n' +
    '📦 عودة — مكرر مُمسح (RET2): ' + fixedRetDup + '\n\n' +
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
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  if (!sheet) { Logger.log('شيت ' + B2C_CONFIG.SHEET_NAME + ' غير موجود'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < B2C_CONFIG.DATA_START) return;
  
  var numRows = lastRow - B2C_CONFIG.DATA_START + 1;
  var ticketData = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.DEP2.fn, numRows, 1).getValues();

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
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  if (!sheet) { ui.alert('شيت ' + B2C_CONFIG.SHEET_NAME + ' غير موجود'); return; }
  var lastRow = sheet.getLastRow();
  var numRows = lastRow > 1 ? lastRow - 1 : 0;
  
  var withFlight = 0, noFlight = 0, partial = 0;
  if (numRows > 0) {
    var dep2Col = sheet.getRange(2, B2C_CONFIG.DEP2.fn, numRows, 1).getValues();
    var ret1Col = sheet.getRange(2, B2C_CONFIG.RET1.fn, numRows, 1).getValues();
    for (var i = 0; i < numRows; i++) {
      var hasDep2 = String(dep2Col[i][0] || '').trim();
      var hasRet1 = String(ret1Col[i][0] || '').trim();
      if (hasDep2 && hasRet1) withFlight++;
      else if (hasDep2 || hasRet1) partial++;
      else noFlight++;
    }
  }
  
  var report = '📊 حالة شيت B2C\n═══════════════════\n\n';
  report += '👤 إجمالي الصفوف: ' + numRows + '\n';
  report += '✅ مكتمل (ذهاب+عودة): ' + withFlight + '\n';
  report += '⚠️ جزئي (ذهاب أو عودة فقط): ' + partial + '\n';
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
// ║              🔍 تشخيص                                         ║
// ╚═══════════════════════════════════════════════════════════════╝

function diagnosePDFSetup() {
  var ui = safeUI_();
  var ss = SpreadsheetApp.openById(B2C_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(B2C_CONFIG.SHEET_NAME);
  if (!sheet) {
    var names = ss.getSheets().map(function(s) { return s.getName(); });
    ui.alert('شيت "' + B2C_CONFIG.SHEET_NAME + '" غير موجود\n\nالشيتات: ' + names.join(', '));
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < B2C_CONFIG.DATA_START) { ui.alert('لا توجد بيانات'); return; }

  var numRows = lastRow - B2C_CONFIG.DATA_START + 1;
  var ticketData = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.DEP2.fn, numRows, 1).getValues();
  var ret1Data = sheet.getRange(B2C_CONFIG.DATA_START, B2C_CONFIG.RET1.fn, numRows, 1).getValues();

  var withTicket = 0, withLink = 0, complete = 0, partial = 0, ready = 0;

  for (var i = 0; i < numRows; i++) {
    var ticket = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    var hasRet1 = String(ret1Data[i][0] || '').trim();

    if (ticket) withTicket++;
    if (link.indexOf('http') === 0) withLink++;

    if (hasDep2 && hasRet1) { complete++; }
    else if (hasDep2 || hasRet1) { partial++; }

    if (ticket && link.indexOf('http') === 0 && !hasDep2) ready++;
  }

  ui.alert(
    '🔍 تشخيص شيت B2C\n═══════════════════\n\n' +
    '📋 إجمالي الصفوف: ' + numRows + '\n' +
    '🎫 مع رقم تذكرة: ' + withTicket + '\n' +
    '🔗 مع رابط PDF: ' + withLink + '\n\n' +
    '✅ مكتمل (DEP2+RET1): ' + complete + '\n' +
    '⚠️ جزئي: ' + partial + '\n' +
    '🆕 جاهز للاستخراج: ' + ready
  );
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
  var folders = DriveApp.getFoldersByName(B2C_CONFIG.FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(B2C_CONFIG.FOLDER_NAME);
}