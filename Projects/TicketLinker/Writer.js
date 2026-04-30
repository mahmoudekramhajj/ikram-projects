/**
 * Writer.js — كتابة رابط Drive في عمود Z في شيت PD (Presonal Details)
 *
 * القواعد الذهبية:
 *   1. passport-keyed lookup (لا row index ثابت)
 *   2. header-based column detection (يحمي من column shift)
 *   3. Idempotent: تجاوز لو نفس URL، استبدل لو مختلف
 *   4. الكتابة في PD فقط — سكريبت آخر يتولى نقلها لـ B2C/GDS
 *   5. Lock للحماية من تشغيلين متزامنين
 */

// ==================== Helpers ====================

/**
 * يجد رقم العمود في صف الـ headers
 * @return {number} 1-based, أو 0 لو لم يوجد
 */
function findColumnByHeader_(sheet, headerName, fallbackCol) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerName) return i + 1;
  }
  // fallback لو الـ header غير موجود (احتياط)
  return fallbackCol || 0;
}

/**
 * يجد صف B2C بـ passport
 * @return {number} 1-based, أو 0 لو لم يوجد
 */
function findRowByPassport_(sheet, passport, passportCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var data = sheet.getRange(2, passportCol, lastRow - 1, 1).getValues();
  var target = String(passport).trim();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === target) return i + 2;
  }
  return 0;
}

/**
 * يبني رابط Drive من fileId
 */
function buildDriveUrl_(fileId) {
  return 'https://drive.google.com/file/d/' + fileId + '/view';
}

// ==================== Lock ====================

function acquireLock_() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(TL.Config.PROP.LOCK_FLAG);
  if (existing) {
    var ts = parseInt(existing, 10);
    var ageMs = Date.now() - ts;
    // قفل عمره أكثر من 10د → اعتبره معطّلاً
    if (ageMs < 10 * 60 * 1000) return false;
  }
  props.setProperty(TL.Config.PROP.LOCK_FLAG, String(Date.now()));
  return true;
}

function releaseLock_() {
  PropertiesService.getScriptProperties().deleteProperty(TL.Config.PROP.LOCK_FLAG);
}

// ==================== Writer الرئيسي ====================

/**
 * يكتب رابط Drive في عمود Z في شيت PD لحاج بـ passport
 * @param {string} passport
 * @param {string} fileId
 * @return {Object} {status, oldZ, newZ, rowIndex}
 */
function writeTicketUrlForPassport_(passport, fileId) {
  if (!passport || !fileId) {
    return { status: 'INVALID_INPUT', error: 'passport or fileId missing' };
  }

  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (!sheet) return { status: 'SHEET_NOT_FOUND' };

  // Header-based detection (يستخدم PD_HEADERS)
  var passportCol = findColumnByHeader_(sheet, TL.Config.PD_HEADERS.PASSPORT, TL.Config.PD_COL.PASSPORT);
  var ticketUrlCol = findColumnByHeader_(sheet, TL.Config.PD_HEADERS.TICKET_URL, TL.Config.PD_COL.TICKET_URL);
  if (!passportCol || !ticketUrlCol) {
    return { status: 'HEADER_MISMATCH', passportCol: passportCol, ticketUrlCol: ticketUrlCol };
  }

  // ابحث عن الصف بـ passport
  var rowIndex = findRowByPassport_(sheet, passport, passportCol);
  if (!rowIndex) {
    return { status: 'PASSPORT_NOT_FOUND_IN_PD', passport: passport };
  }

  // اقرأ Z الحالي
  var newUrl = buildDriveUrl_(fileId);
  var currentZ = String(sheet.getRange(rowIndex, ticketUrlCol).getValue() || '').trim();

  if (currentZ === newUrl) {
    return { status: 'ALREADY_WRITTEN', rowIndex: rowIndex, oldZ: currentZ, newZ: newUrl };
  }

  // Idempotent: استبدل إن مختلف
  sheet.getRange(rowIndex, ticketUrlCol).setValue(newUrl);

  return {
    status: currentZ ? 'REPLACED' : 'WRITTEN',
    rowIndex: rowIndex,
    oldZ: currentZ,
    newZ: newUrl
  };
}

// ==================== سجل التشغيل (RunLog) ====================

/**
 * يسجّل عملية كتابة في تبويب TL_RunLog
 * (ينشئ التبويب لو غير موجود)
 */
function logWriteOperation_(entry) {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_RunLog');
  if (!sheet) {
    sheet = ss.insertSheet('TL_RunLog');
    sheet.appendRow([
      'Timestamp', 'Passport', 'FileId', 'FileName', 'Stage',
      'Status', 'OldZ', 'NewZ', 'RowIndex', 'Notes'
    ]);
  }
  sheet.appendRow([
    new Date(),
    entry.passport || '',
    entry.fileId || '',
    entry.fileName || '',
    entry.stage || '',
    entry.status || '',
    entry.oldZ || '',
    entry.newZ || '',
    entry.rowIndex || '',
    entry.notes || ''
  ]);
}

// ==================== اختبار ====================

function testWriteForPassport(passport, fileId) {
  return writeTicketUrlForPassport_(passport, fileId);
}
