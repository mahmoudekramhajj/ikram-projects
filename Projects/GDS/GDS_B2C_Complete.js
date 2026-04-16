/**
 * ══════════════════════════════════════════════════════════════════
 * B2C Complete — سكربت موحّد لإدارة شيت B2C
 * ══════════════════════════════════════════════════════════════════
 *
 * المرحلة 1: مزامنة أسماء حجاج B2C من Presonal Details
 * المرحلة 2: استخراج بيانات الرحلات من تذاكر PDF
 * المرحلة 3: إصلاح تصنيف الرحلات المباشرة
 *
 * هيكل أعمدة B2C:
 *   A-AF  (1-32)  = بيانات شخصية (من PD)
 *   AG-AM (33-39) = DEP1 قدوم رحلة 1
 *   AN-AT (40-46) = DEP2 قدوم رحلة 2
 *   AU-BA (47-53) = RET1 عودة رحلة 1
 *   BB-BH (54-60) = RET2 عودة رحلة 2
 *   BI    (61)    = PNR
 *
 * التصنيف:
 *   مباشر  → DEP2 + RET1
 *   ترانزيت → DEP1 + DEP2 + RET1 + RET2
 * ══════════════════════════════════════════════════════════════════
 */


// ==================== الإعدادات ====================

var CONFIG = {
  SS_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  PD_SHEET: 'Presonal Details',
  B2C_SHEET: 'B2C',

  // أعمدة مرجعية
  COL_PASSPORT: 6,        // F
  COL_CONTRACT: 21,       // U — نوع عقد الطيران
  COL_TICKET: 25,         // Y — رقم التذكرة
  COL_LINK: 26,           // Z — رابط التذكرة PDF
  COL_PNR: 61,            // BI

  // أعمدة PD → B2C
  PD_DIRECT_COLS: 27,     // A-AA نسخ مباشر
  PD_SHIFT_START: 29,     // AC
  PD_SHIFT_END: 33,       // AG

  // أعمدة الرحلات
  DEP1: { fn: 33, dd: 34, dt: 35, fr: 36, to: 37, ad: 38, at: 39 },
  DEP2: { fn: 40, dd: 41, dt: 42, fr: 43, to: 44, ad: 45, at: 46 },
  RET1: { fn: 47, dd: 48, dt: 49, fr: 50, to: 51, ad: 52, at: 53 },
  RET2: { fn: 54, dd: 55, dt: 56, fr: 57, to: 58, ad: 59, at: 60 },

  FLIGHT_START: 33,
  FLIGHT_END: 61,

  // PDF
  BATCH_SIZE: 25,
  MAX_TIME: 4.5 * 60 * 1000,
  FOLDER_NAME: 'Ikram_Tickets'
};


// ==================== القائمة ====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔄 B2C Complete')
    .addItem('▶️ مزامنة شاملة (أسماء + PDF)', 'runFullSync')
    .addSeparator()
    .addItem('👤 مزامنة أسماء فقط', 'syncB2C')
    .addItem('📄 استخراج PDF فقط', 'runPDFOnly')
    .addItem('🔧 إصلاح التصنيف', 'fixClassification')
    .addSeparator()
    .addItem('📊 تقرير الحالة', 'showStatus')
    .addItem('🔍 تشخيص PDF', 'diagnosePDFSetup')
    .addItem('👁️ معاينة الأسماء', 'previewSync')
    .addSeparator()
    .addItem('⏰ تفعيل التشغيل التلقائي', 'setupTriggers')
    .addItem('⏹️ إيقاف الكل', 'stopAll')
    .addItem('❌ إلغاء التشغيل التلقائي', 'removeTriggers')
    .addToUi();
}

function safeUI_() {
  try { return SpreadsheetApp.getUi(); }
  catch (e) { return { alert: function(msg) { Logger.log(msg); } }; }
}

function toast_(msg, title, sec) {
  try { SpreadsheetApp.openById(CONFIG.SS_ID).toast(msg, title || '', sec || 5); }
  catch (e) { Logger.log((title || '') + ': ' + msg); }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🚀 مزامنة شاملة                                  ║
// ╚═══════════════════════════════════════════════════════════════╝

function runFullSync() {
  var ui = safeUI_();
  var namesResult = _syncNames();
  var newCount = setupPDFExtraction_();

  if (newCount > 0) {
    ui.alert(
      '🚀 المزامنة الشاملة\n\n' +
      '👤 أسماء جديدة: ' + namesResult.added + '\n' +
      '📄 تذاكر للاستخراج: ' + newCount + '\n\n' +
      '⏳ جاري البدء...'
    );
    startPDFAuto_();
  } else {
    ui.alert(
      '✅ المزامنة الشاملة\n\n' +
      '👤 أسماء جديدة: ' + namesResult.added + '\n' +
      '📄 لا توجد تذاكر جديدة'
    );
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              👤 المرحلة 1: مزامنة الأسماء                      ║
// ╚═══════════════════════════════════════════════════════════════╝

function syncB2C() {
  var startTime = new Date();
  Logger.log('━━━ مزامنة أسماء B2C ━━━');
  var result = _syncNames();
  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  Logger.log('👤 مُضاف: ' + result.added + ' | موجود: ' + result.existing + ' | ⏱️ ' + elapsed + 'ث');
}

function _syncNames() {
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var pdSheet = ss.getSheetByName(CONFIG.PD_SHEET);
  var b2cSheet = ss.getSheetByName(CONFIG.B2C_SHEET);

  if (!pdSheet || !b2cSheet) {
    Logger.log('❌ شيت غير موجود!');
    return { added: 0, existing: 0 };
  }

  var b2cLastRow = b2cSheet.getLastRow();
  var b2cPassports = {};

  if (b2cLastRow > 1) {
    var b2cData = b2cSheet.getRange(2, CONFIG.COL_PASSPORT, b2cLastRow - 1, 1).getValues();
    for (var i = 0; i < b2cData.length; i++) {
      var p = String(b2cData[i][0] || '').trim();
      if (p) b2cPassports[p] = true;
    }
  }

  var pdLastRow = pdSheet.getLastRow();
  var pdData = pdSheet.getRange(2, 1, pdLastRow - 1, CONFIG.PD_SHIFT_END).getValues();

  var newRows = [];
  var existing = 0;

  for (var j = 0; j < pdData.length; j++) {
    var row = pdData[j];
    var contractType = String(row[CONFIG.COL_CONTRACT - 1] || '').trim();
    if (contractType !== 'B2C') continue;

    var passport = String(row[CONFIG.COL_PASSPORT - 1] || '').trim();
    if (!passport) continue;

    if (b2cPassports[passport]) { existing++; continue; }

    newRows.push(_mapPdToB2c(row));
    b2cPassports[passport] = true;
  }

  if (newRows.length > 0) {
    var emptyFlights = new Array(CONFIG.FLIGHT_END - CONFIG.FLIGHT_START + 1).fill('');
    var fullRows = newRows.map(function(r) { return r.concat(emptyFlights); });
    b2cSheet.getRange(b2cLastRow + 1, 1, fullRows.length, fullRows[0].length).setValues(fullRows);
    Logger.log('✅ أُضيف ' + newRows.length + ' حاج جديد');
  }

  return { added: newRows.length, existing: existing };
}

function _mapPdToB2c(pdRow) {
  var b2cRow = [];
  for (var i = 0; i < CONFIG.PD_DIRECT_COLS; i++) b2cRow.push(pdRow[i] || '');
  // تجاوز عمود 28 (المخيم)
  for (var j = CONFIG.PD_SHIFT_START - 1; j < CONFIG.PD_SHIFT_END; j++) b2cRow.push(pdRow[j] || '');
  return b2cRow;
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              📄 المرحلة 2: استخراج PDF                         ║
// ╚═══════════════════════════════════════════════════════════════╝

function runPDFOnly() {
  var newCount = setupPDFExtraction_();
  if (newCount > 0) {
    safeUI_().alert('📄 تذاكر جديدة: ' + newCount + '\n\n⏳ جاري البدء...');
    startPDFAuto_();
  } else {
    safeUI_().alert('✅ لا توجد تذاكر جديدة');
  }
}

function setupPDFExtraction_() {
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var numRows = lastRow - 1;
  var ticketData = sheet.getRange(2, CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(2, CONFIG.DEP2.fn, numRows, 1).getValues();

  var tickets = {};
  for (var i = 0; i < ticketData.length; i++) {
    var ticketNo = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    if (!ticketNo || link.indexOf('http') !== 0 || hasDep2) continue;
    if (!tickets[ticketNo]) tickets[ticketNo] = link;
  }

  var folder = getOrCreateFolder_();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('B2C_TICKETS', JSON.stringify(tickets));
  props.setProperty('B2C_PDF_DONE', '[]');
  props.setProperty('B2C_PDF_FAILED', '[]');
  props.setProperty('B2C_FOLDER', folder.getId());

  return Object.keys(tickets).length;
}

function startPDFAuto_() {
  stopPDFTrigger_();
  ScriptApp.newTrigger('processPDFBatch').timeBased().everyMinutes(1).create();
  toast_('🚀 PDF — ' + CONFIG.BATCH_SIZE + ' تذكرة/دقيقة', '▶️', 10);
  processPDFBatch();
}

function processPDFBatch() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  if (!sheet) return;

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
    return;
  }

  var lastRow = sheet.getLastRow();
  var ticketCol = sheet.getRange(2, CONFIG.COL_TICKET, lastRow - 1, 1).getValues();
  var ticketRowMap = {};
  for (var i = 0; i < ticketCol.length; i++) {
    var t = String(ticketCol[i][0] || '').trim();
    if (t) {
      if (!ticketRowMap[t]) ticketRowMap[t] = [];
      ticketRowMap[t].push(i + 2);
    }
  }

  var folder = DriveApp.getFolderById(folderId);
  var batch = remaining.slice(0, CONFIG.BATCH_SIZE);
  var startTime = Date.now();
  var batchOK = 0, batchFail = 0;

  for (var b = 0; b < batch.length; b++) {
    if (Date.now() - startTime > CONFIG.MAX_TIME) break;

    var ticketNo = batch[b];
    var pdfUrl = allTickets[ticketNo];

    try {
      var blob = null;
      var existingFiles = folder.getFilesByName(ticketNo + '.pdf');

      if (existingFiles.hasNext()) {
        blob = existingFiles.next().getBlob();
      } else {
        blob = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true, followRedirects: true }).getBlob();
        blob.setName(ticketNo + '.pdf').setContentType('application/pdf');
        folder.createFile(blob);
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
// ║              ✍️ كتابة بيانات الرحلات                          ║
// ╚═══════════════════════════════════════════════════════════════╝

function writeFlightData_(sheet, row, flightData) {
  var dep = flightData.segments.departure || [];
  var ret = flightData.segments.return || [];

  if (flightData.pnr) sheet.getRange(row, CONFIG.COL_PNR).setValue(flightData.pnr);

  if (dep.length === 1) {
    writeSegment_(sheet, row, CONFIG.DEP2, dep[0]);
    clearSegment_(sheet, row, CONFIG.DEP1);
  } else if (dep.length >= 2) {
    writeSegment_(sheet, row, CONFIG.DEP1, dep[0]);
    writeSegment_(sheet, row, CONFIG.DEP2, dep[1]);
  }

  if (ret.length === 1) {
    writeSegment_(sheet, row, CONFIG.RET1, ret[0]);
    clearSegment_(sheet, row, CONFIG.RET2);
  } else if (ret.length >= 2) {
    writeSegment_(sheet, row, CONFIG.RET1, ret[0]);
    writeSegment_(sheet, row, CONFIG.RET2, ret[1]);
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
  sheet.getRange(row, cols.fn, 1, 7).clearContent();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🔧 إصلاح التصنيف                                 ║
// ╚═══════════════════════════════════════════════════════════════╝

function fixClassification() {
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  var ui = safeUI_();
  if (!sheet) { ui.alert('شيت B2C غير موجود'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var numRows = lastRow - 1;
  var allData = sheet.getRange(2, 1, numRows, 61).getValues();
  var fixedDup = 0, fixedMove = 0, fixedRetDup = 0;

  for (var i = 0; i < numRows; i++) {
    var row = i + 2;
    var dep1Fn = cleanVal_(allData[i][CONFIG.DEP1.fn - 1]);
    var dep2Fn = cleanVal_(allData[i][CONFIG.DEP2.fn - 1]);

    if (dep1Fn && dep2Fn && dep1Fn.replace(/-/g, '') === dep2Fn.replace(/-/g, '')) {
      clearSegment_(sheet, row, CONFIG.DEP1); fixedDup++;
    } else if (dep1Fn && !dep2Fn) {
      sheet.getRange(row, CONFIG.DEP2.fn, 1, 7).setValues(sheet.getRange(row, CONFIG.DEP1.fn, 1, 7).getValues());
      clearSegment_(sheet, row, CONFIG.DEP1); fixedMove++;
    }

    var ret1Fn = cleanVal_(allData[i][CONFIG.RET1.fn - 1]);
    var ret2Fn = cleanVal_(allData[i][CONFIG.RET2.fn - 1]);
    if (ret1Fn && ret2Fn && ret1Fn.replace(/-/g, '') === ret2Fn.replace(/-/g, '')) {
      clearSegment_(sheet, row, CONFIG.RET2); fixedRetDup++;
    }
  }

  ui.alert('🔧 إصلاحات: مكرر=' + fixedDup + ' | نقل=' + fixedMove + ' | عودة=' + fixedRetDup);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              ⏰ التشغيل التلقائي                               ║
// ╚═══════════════════════════════════════════════════════════════╝

function autoCheckNewRows() {
  var namesResult = _syncNames();
  if (namesResult.added > 0) Logger.log('⏰ أسماء جديدة: ' + namesResult.added);

  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var numRows = lastRow - 1;
  var ticketData = sheet.getRange(2, CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(2, CONFIG.DEP2.fn, numRows, 1).getValues();

  var newTickets = {};
  for (var i = 0; i < ticketData.length; i++) {
    var ticketNo = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    if (!ticketNo || link.indexOf('http') !== 0 || hasDep2 || newTickets[ticketNo]) continue;
    newTickets[ticketNo] = link;
  }

  var count = Object.keys(newTickets).length;
  if (count === 0) { Logger.log('⏰ لا توجد تذاكر جديدة'); return; }

  Logger.log('⏰ ' + count + ' تذكرة جديدة — بدء المعالجة');
  var props = PropertiesService.getScriptProperties();
  var folder = getOrCreateFolder_();
  props.setProperty('B2C_TICKETS', JSON.stringify(newTickets));
  props.setProperty('B2C_PDF_DONE', '[]');
  props.setProperty('B2C_PDF_FAILED', '[]');
  props.setProperty('B2C_FOLDER', folder.getId());
  startPDFAuto_();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              📊 التقارير                                       ║
// ╚═══════════════════════════════════════════════════════════════╝

function showStatus() {
  var ui = safeUI_();
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var pdSheet = ss.getSheetByName(CONFIG.PD_SHEET);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  if (!sheet) { ui.alert('شيت B2C غير موجود'); return; }

  // PD count
  var pdLastRow = pdSheet.getLastRow();
  var pdTypes = pdSheet.getRange(2, CONFIG.COL_CONTRACT, pdLastRow - 1, 1).getValues();
  var pdB2c = 0;
  for (var i = 0; i < pdTypes.length; i++) {
    if (String(pdTypes[i][0] || '').trim() === 'B2C') pdB2c++;
  }

  // B2C stats
  var lastRow = sheet.getLastRow();
  var numRows = lastRow > 1 ? lastRow - 1 : 0;
  var withFlight = 0, noFlight = 0;

  if (numRows > 0) {
    var dep2Col = sheet.getRange(2, CONFIG.DEP2.fn, numRows, 1).getValues();
    var ret1Col = sheet.getRange(2, CONFIG.RET1.fn, numRows, 1).getValues();
    for (var j = 0; j < numRows; j++) {
      if (String(dep2Col[j][0] || '').trim() && String(ret1Col[j][0] || '').trim()) withFlight++;
      else noFlight++;
    }
  }

  // PDF stats
  var props = PropertiesService.getScriptProperties();
  var pdfDone = JSON.parse(props.getProperty('B2C_PDF_DONE') || '[]');
  var pdfFailed = JSON.parse(props.getProperty('B2C_PDF_FAILED') || '[]');

  var triggers = ScriptApp.getProjectTriggers().map(function(t) { return t.getHandlerFunction(); });

  ui.alert(
    '📊 حالة B2C\n═══════════════════\n\n' +
    '👤 PD: ' + pdB2c + ' | B2C: ' + numRows + ' | فرق: ' + (pdB2c - numRows) + '\n' +
    '✅ مع رحلة: ' + withFlight + ' | ❓ بدون: ' + noFlight + '\n\n' +
    '📄 PDF نجح: ' + pdfDone.length + ' | فشل: ' + pdfFailed.length + '\n\n' +
    '⏰ Triggers: ' + (triggers.length > 0 ? triggers.join(', ') : 'لا يوجد')
  );
}

function previewSync() {
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var pdSheet = ss.getSheetByName(CONFIG.PD_SHEET);
  var b2cSheet = ss.getSheetByName(CONFIG.B2C_SHEET);

  var b2cLastRow = b2cSheet.getLastRow();
  var b2cPassports = {};
  if (b2cLastRow > 1) {
    var b2cData = b2cSheet.getRange(2, CONFIG.COL_PASSPORT, b2cLastRow - 1, 1).getValues();
    for (var i = 0; i < b2cData.length; i++) {
      var p = String(b2cData[i][0] || '').trim();
      if (p) b2cPassports[p] = true;
    }
  }

  var pdData = pdSheet.getRange(2, 1, pdSheet.getLastRow() - 1, CONFIG.COL_CONTRACT).getValues();
  var newNames = [];
  for (var j = 0; j < pdData.length; j++) {
    if (String(pdData[j][CONFIG.COL_CONTRACT - 1] || '').trim() !== 'B2C') continue;
    var passport = String(pdData[j][CONFIG.COL_PASSPORT - 1] || '').trim();
    if (!passport || b2cPassports[passport]) continue;
    newNames.push(String(pdData[j][10] || '') + ' ' + String(pdData[j][11] || '').trim());
    b2cPassports[passport] = true;
  }

  Logger.log('أسماء جديدة: ' + newNames.length);
  for (var k = 0; k < Math.min(newNames.length, 15); k++) Logger.log('  🆕 ' + newNames[k]);
}

function diagnosePDFSetup() {
  var ui = safeUI_();
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  if (!sheet) { ui.alert('شيت B2C غير موجود'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('لا توجد بيانات'); return; }

  var numRows = lastRow - 1;
  var ticketData = sheet.getRange(2, CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(2, CONFIG.DEP2.fn, numRows, 1).getValues();
  var ret1Data = sheet.getRange(2, CONFIG.RET1.fn, numRows, 1).getValues();

  var withTicket = 0, withLink = 0, complete = 0, partial = 0, ready = 0;
  for (var i = 0; i < numRows; i++) {
    var ticket = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    var hasRet1 = String(ret1Data[i][0] || '').trim();
    if (ticket) withTicket++;
    if (link.indexOf('http') === 0) withLink++;
    if (hasDep2 && hasRet1) complete++;
    else if (hasDep2 || hasRet1) partial++;
    if (ticket && link.indexOf('http') === 0 && !hasDep2) ready++;
  }

  ui.alert(
    '🔍 تشخيص PDF\n═══════════════════\n\n' +
    '📋 صفوف: ' + numRows + '\n' +
    '🎫 تذكرة: ' + withTicket + ' | 🔗 رابط: ' + withLink + '\n\n' +
    '✅ مكتمل: ' + complete + ' | ⚠️ جزئي: ' + partial + '\n' +
    '🆕 جاهز: ' + ready
  );
}

function debugFailedPDFs() {
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.B2C_SHEET);
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1;
  var ticketData = sheet.getRange(2, CONFIG.COL_TICKET, numRows, 2).getValues();
  var dep2Data = sheet.getRange(2, CONFIG.DEP2.fn, numRows, 1).getValues();

  var failed = [], seen = {};
  for (var i = 0; i < numRows; i++) {
    var ticket = String(ticketData[i][0] || '').trim();
    var link = String(ticketData[i][1] || '').trim();
    var hasDep2 = String(dep2Data[i][0] || '').trim();
    if (ticket && link.indexOf('http') === 0 && !hasDep2 && !seen[ticket]) {
      seen[ticket] = true;
      failed.push({ ticket: ticket, link: link });
    }
  }

  Logger.log('تذاكر بدون بيانات (مع رابط): ' + failed.length);
  var folder = getOrCreateFolder_();

  for (var j = 0; j < Math.min(3, failed.length); j++) {
    var t = failed[j];
    Logger.log('\n━━━ تذكرة: ' + t.ticket + ' ━━━');
    try {
      var blob, ef = folder.getFilesByName(t.ticket + '.pdf');
      blob = ef.hasNext() ? ef.next().getBlob() : UrlFetchApp.fetch(t.link, {muteHttpExceptions:true,followRedirects:true}).getBlob();
      var text = extractTextFromPDF_(blob);
      if (!text) { Logger.log('OCR فارغ'); continue; }
      Logger.log(text.substring(0, 2000));
      var res = parseTicketText_(text);
      Logger.log('PNR: ' + (res.pnr||'-') + ' | ذهاب: ' + res.segments.departure.length + ' | عودة: ' + res.segments.return.length);
      res.segments.departure.forEach(function(s){Logger.log('  DEP: ' + s.flightNo + ' ' + s.depDate + ' ' + s.depTime + ' ' + s.from + '>' + s.arrTime + ' ' + s.to);});
      res.segments.return.forEach(function(s){Logger.log('  RET: ' + s.flightNo + ' ' + s.depDate + ' ' + s.depTime + ' ' + s.from + '>' + s.arrTime + ' ' + s.to);});
    } catch(e) { Logger.log('خطأ: ' + e.message); }
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              ⏰ Triggers                                       ║
// ╚═══════════════════════════════════════════════════════════════╝

function initialSetup() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);

  ScriptApp.newTrigger('onOpen').forSpreadsheet(CONFIG.SS_ID).onOpen().create();
  ScriptApp.newTrigger('autoCheckNewRows').timeBased().everyMinutes(10).create();

  Logger.log('✅ القائمة + فحص تلقائي كل 10 دقائق');
}

function setupTriggers() {
  removeTriggers();
  ScriptApp.newTrigger('autoCheckNewRows').timeBased().everyMinutes(10).create();
  safeUI_().alert('✅ تم التفعيل\n\nفحص تلقائي كل 10 دقائق (أسماء + تذاكر)');
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'autoCheckNewRows' || fn === 'syncB2C') ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('❌ تم إلغاء التشغيل التلقائي');
}

function stopAll() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'processPDFBatch' || fn === 'autoCheckNewRows' || fn === 'syncB2C') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  toast_('⏹️ تم إيقاف الكل', 'إيقاف', 5);
}

function stopPDFTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processPDFBatch') ScriptApp.deleteTrigger(triggers[i]);
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              🔨 دوال مساعدة                                    ║
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
      { method: 'post', contentType: 'multipart/related; boundary=' + boundary,
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: payload, muteHttpExceptions: true }
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

  var pnrMatch = text.match(/(?:Airline Reference|PNR|Airline PNR)\s*[\(\):\n\r\s]*([A-Z0-9]{5,7})/i) ||
                 text.match(/Airline Reference\s*\(PNR\)\s*[\n\r]+\s*([A-Z0-9]{5,7})/i) ||
                 text.match(/PNR[:\s]*[\n\r]+\s*([A-Z0-9]{5,7})/i);
  if (pnrMatch) result.pnr = pnrMatch[1].trim();

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

  // Trim at non-flight sections
  var stopMarkers = ['Traveller Route', 'Cancel and Date change', 'Travel checklist'];
  for (var sm = 0; sm < stopMarkers.length; sm++) {
    var di = depText.indexOf(stopMarkers[sm]);
    if (di > 0) depText = depText.substring(0, di);
    var ri = retText.indexOf(stopMarkers[sm]);
    if (ri > 0) retText = retText.substring(0, ri);
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
  var seen = {}, unique = [];
  for (var i = 0; i < segments.length; i++) {
    var key = segments[i].flightNo + '|' + segments[i].from + '|' + segments[i].to;
    if (!seen[key]) { seen[key] = true; unique.push(segments[i]); }
  }
  return unique;
}

function extractFlightSegments_(sectionText) {
  if (!sectionText || sectionText.length < 20) return [];

  var segments = [];
  var flightRegex = /\b([A-Z][A-Z0-9]|[0-9][A-Z])[-\s]?(\d{2,4})\b/g;
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

    if (!arrInfo.time) {
      var firstBefore = extractDateTimeAirport_(beforeText, false);
      if (firstBefore.airport && firstBefore.airport !== depInfo.airport) arrInfo = firstBefore;
    }

    segments.push({
      flightNo: flight.number,
      depDate: depInfo.date || '', depTime: depInfo.time || '', from: depInfo.airport || '',
      arrDate: arrInfo.date || '', arrTime: arrInfo.time || '', to: arrInfo.airport || ''
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
    var months = {'jan':'01','feb':'02','mar':'03','apr':'04','may':'05','jun':'06',
                  'jul':'07','aug':'08','sep':'09','oct':'10','nov':'11','dec':'12'};
    var parts = dateStr.replace(',', '').trim().split(/\s+/);
    if (parts.length >= 3) return parts[2] + '-' + (months[parts[0].substring(0,3).toLowerCase()] || '01') + '-' + ('0' + parseInt(parts[1])).slice(-2);
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
