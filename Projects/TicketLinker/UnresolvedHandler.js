/**
 * UnresolvedHandler.js — المرحلة 5: تذاكر لم نجد لها صاحباً
 *
 * المنطق:
 *   1. ننشئ مجلد فرعي "Unresolved_Tickets" في Drive (لو غير موجود)
 *   2. ننسخ PDF إلى هذا المجلد (لا ننقله، نُبقي الأصل)
 *   3. ننشئ صفّاً في تبويب TL_UnresolvedTickets يحوي:
 *      - رابط PDF، اسم الملف، اسم المجلد
 *      - بيانات Claude المستخرَجة (الأسماء، الرحلات)
 *      - خانتان فارغتان: الجنسية، رقم الجواز
 *      - حالة: PENDING_MANUAL
 *   4. الموظف يملأ الخانتين يدوياً
 *   5. HourlyTrigger يفحص كل ساعة ويُكمل المعالجة
 */

var TL_UNRESOLVED_FOLDER_NAME = 'Unresolved_Tickets';
var TL_UNRESOLVED_TAB = 'TL_UnresolvedTickets';

// ==================== مجلد Drive ====================

function ensureUnresolvedFolder_() {
  var parent = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var subs = parent.getFoldersByName(TL_UNRESOLVED_FOLDER_NAME);
  if (subs.hasNext()) return subs.next();
  return parent.createFolder(TL_UNRESOLVED_FOLDER_NAME);
}

// ==================== تبويب الشيت ====================

function ensureUnresolvedTab_() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL_UNRESOLVED_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(TL_UNRESOLVED_TAB);
    sheet.appendRow([
      'Timestamp',           // A
      'OriginalFileId',      // B
      'CopiedFileId',        // C
      'OriginalFileName',    // D
      'FolderName',          // E
      'ExtractedName',       // F (من اسم الملف)
      'ClaudeFullName',      // G
      'ClaudePnr',           // H
      'ClaudeTicketNo',      // I
      'AllPassengers',       // J (JSON)
      'Flights',             // K (JSON)
      'Nationality',         // L ⚠️ خانة فارغة - يملأها الموظف
      'Passport',            // M ⚠️ خانة فارغة - يملأها الموظف
      'Status',              // N (PENDING_MANUAL / LINKED / NO_MATCH_AFTER_FILL)
      'LinkedAt',            // O
      'Notes'                // P
    ]);
    sheet.setFrozenRows(1);
    // تظليل الخانتين الفارغتين بأصفر فاتح للفت النظر
    sheet.getRange('L:M').setBackground('#fff2cc');
  }
  return sheet;
}

// ==================== التسجيل ====================

/**
 * يضيف PDF غير مُطابَق للمراجعة اليدوية (مع dedup عبر fileId)
 * - لو fileId موجود سابقاً → تحديث Timestamp + Notes فقط (لا يلمس L/M/N).
 * - لو جديد → appendRow كاملاً.
 * @param {Object} pdfInfo {fileId, fileName, folderName}
 * @param {Object} extracted {extractedName, claudeData (optional)}
 * @return {Object} {status, copiedFileId, rowAdded}
 */
function registerUnresolved_(pdfInfo, extracted) {
  var sheet = ensureUnresolvedTab_();
  var claudeData = (extracted && extracted.claudeData) || {};
  var notesStr = (extracted && extracted.notes) ? extracted.notes.join(' | ') : '';

  // ===== فحص dedup: هل fileId موجود في عمود B؟ =====
  var existingRow = 0;
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var fileIdCol = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    var target = String(pdfInfo.fileId).trim();
    for (var i = 0; i < fileIdCol.length; i++) {
      if (String(fileIdCol[i][0]).trim() === target) {
        existingRow = i + 2;
        break;
      }
    }
  }

  if (existingRow) {
    // التكرار: حدّث Timestamp + Notes فقط (احفظ L/M/N من تدخّل الموظف)
    var oldNotes = String(sheet.getRange(existingRow, 16).getValue() || '').trim();
    var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd HH:mm');
    var mergedNotes = oldNotes
      ? (oldNotes + ' || retry@' + ts + ': ' + notesStr)
      : notesStr;
    sheet.getRange(existingRow, 1).setValue(new Date());      // A - Timestamp
    sheet.getRange(existingRow, 16).setValue(mergedNotes);    // P - Notes

    return {
      status: 'UNRESOLVED_UPDATED',
      copiedFileId: null,
      fileName: pdfInfo.fileName,
      rowIndex: existingRow
    };
  }

  // ===== جديد: انسخ PDF + أضف صفاً جديداً =====
  var unresolvedFolder = ensureUnresolvedFolder_();
  var originalFile = DriveApp.getFileById(pdfInfo.fileId);
  var copiedFile;
  try {
    copiedFile = originalFile.makeCopy(pdfInfo.fileName, unresolvedFolder);
  } catch (e) {
    copiedFile = null;
  }

  sheet.appendRow([
    new Date(),
    pdfInfo.fileId,
    copiedFile ? copiedFile.getId() : '',
    pdfInfo.fileName,
    pdfInfo.folderName || '',
    (extracted && extracted.extractedName) || '',
    claudeData.fullName || '',
    claudeData.pnr || '',
    claudeData.ticketNumber || '',
    JSON.stringify(claudeData.allPassengers || []),
    JSON.stringify(claudeData.flights || []),
    '',  // Nationality - فارغ
    '',  // Passport - فارغ
    'PENDING_MANUAL',
    '',  // LinkedAt
    notesStr
  ]);

  return {
    status: 'REGISTERED_UNRESOLVED',
    copiedFileId: copiedFile ? copiedFile.getId() : null,
    fileName: pdfInfo.fileName
  };
}

// ==================== فحص الإكمال (يُستدعى من HourlyTrigger) ====================

/**
 * يفحص تبويب Unresolved، ويكمل معالجة الصفوف التي امتلأت خاناتها
 */
function processFilledUnresolvedRows_() {
  var sheet = ensureUnresolvedTab_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { processed: 0 };

  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  var processed = 0;
  var updates = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var status = String(row[13] || '').trim();        // N - Status
    var nationality = String(row[11] || '').trim();    // L
    var passport = String(row[12] || '').trim();       // M

    if (status !== 'PENDING_MANUAL') continue;
    if (!nationality || !passport) continue;

    var originalFileId = row[1];
    var rowIndex = i + 2;

    // اكتب الرابط في B2C
    var writeResult = writeTicketUrlForPassport_(passport, originalFileId);

    if (writeResult.status === 'WRITTEN' || writeResult.status === 'REPLACED' || writeResult.status === 'ALREADY_WRITTEN') {
      // نجح → حدّث Status لـ LINKED
      updates.push({
        row: rowIndex,
        status: 'LINKED',
        linkedAt: new Date(),
        notes: writeResult.status + ' (B2C row ' + writeResult.rowIndex + ')'
      });
      processed++;
    } else {
      // فشل → حدّث للتنبيه
      updates.push({
        row: rowIndex,
        status: 'NO_MATCH_AFTER_FILL',
        linkedAt: '',
        notes: writeResult.status + ': ' + (writeResult.error || '')
      });
    }
  }

  // طبّق التحديثات
  for (var j = 0; j < updates.length; j++) {
    var u = updates[j];
    sheet.getRange(u.row, 14).setValue(u.status);   // N
    sheet.getRange(u.row, 15).setValue(u.linkedAt); // O
    sheet.getRange(u.row, 16).setValue(u.notes);    // P
  }

  return { processed: processed, totalUpdates: updates.length };
}

// ==================== اختبار ====================

/**
 * إعادة معالجة Unresolved بـ Claude (لـ COLLISION/NoPdMatch/Other)
 * تتجاوز الصفوف التي notes فيها "NOT IN B2C" لأن Claude لا يحلّها.
 *
 * @return {Object} stats
 */
function retryUnresolvedWithClaude() {
  var sheet = ensureUnresolvedTab_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { processed: 0 };

  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;
  var results = {
    examined: 0,
    skippedNoB2cBnotIssue: 0,
    skippedAlreadyLinked: 0,
    claudeRan: 0,
    nowLinked: 0,
    stillUnresolved: 0,
    claudeErrors: 0,
    rowsUpdated: []
  };

  for (var i = 0; i < data.length; i++) {
    if (Date.now() >= deadline) break;
    results.examined++;

    var row = data[i];
    var rowIndex = i + 2;
    var status = String(row[13] || '').trim();    // N
    var notes = String(row[15] || '').trim();      // P
    var originalFileId = row[1];
    var fileName = String(row[3] || '');
    var folderName = String(row[4] || '');

    // تجاوز
    if (status !== 'PENDING_MANUAL') {
      results.skippedAlreadyLinked++;
      continue;
    }
    if (notes.indexOf('NOT IN B2C') !== -1) {
      results.skippedNoB2cBnotIssue++;
      continue;
    }
    // تجاوز الصفوف التي جُرِّبت بـ Claude سابقاً وفشلت (لا فائدة من إعادة المحاولة)
    if (notes.indexOf('Claude:') !== -1 || notes.indexOf('Claude error') !== -1 ||
        notes.indexOf('Claude found in PD but') !== -1) {
      results.skippedAlreadyLinked++;
      continue;
    }
    if (!originalFileId) continue;

    // جرّب Claude
    results.claudeRan++;
    try {
      var claudeMatch = matchViaPdfClaudeExtraction_({
        fileId: originalFileId,
        fileName: fileName,
        folderName: folderName
      });

      // حفظ Claude data في الصف (G, H, I, J, K)
      if (claudeMatch.claudeData) {
        var cd = claudeMatch.claudeData;
        sheet.getRange(rowIndex, 7).setValue(cd.fullName || '');
        sheet.getRange(rowIndex, 8).setValue(cd.pnr || '');
        sheet.getRange(rowIndex, 9).setValue(cd.ticketNumber || '');
        sheet.getRange(rowIndex, 10).setValue(JSON.stringify(cd.allPassengers || []));
        sheet.getRange(rowIndex, 11).setValue(JSON.stringify(cd.flights || []));
      }

      if (claudeMatch.status === 'OK' && claudeMatch.passport) {
        // اكتب في B2C
        var writeResult = writeTicketUrlForPassport_(claudeMatch.passport, originalFileId);

        if (writeResult.status === 'WRITTEN' || writeResult.status === 'REPLACED' || writeResult.status === 'ALREADY_WRITTEN') {
          sheet.getRange(rowIndex, 13).setValue(claudeMatch.passport);  // M (passport)
          sheet.getRange(rowIndex, 14).setValue('LINKED');               // N
          sheet.getRange(rowIndex, 15).setValue(new Date());             // O
          sheet.getRange(rowIndex, 16).setValue('Claude: ' + (claudeMatch.notes || []).join(' | '));
          results.nowLinked++;
          results.rowsUpdated.push({ row: rowIndex, status: 'LINKED', passport: claudeMatch.passport });
          continue;
        } else {
          sheet.getRange(rowIndex, 13).setValue(claudeMatch.passport);
          sheet.getRange(rowIndex, 16).setValue('Claude found in PD but ' + writeResult.status);
        }
      } else {
        sheet.getRange(rowIndex, 16).setValue('Claude: ' + claudeMatch.status + ' | ' + (claudeMatch.notes || []).join(' | '));
      }
      results.stillUnresolved++;
    } catch (e) {
      results.claudeErrors++;
      sheet.getRange(rowIndex, 16).setValue('Claude error: ' + e.message);
    }
  }

  return results;
}

function testRegisterUnresolved(fileId, fileName, folderName) {
  return registerUnresolved_(
    { fileId: fileId, fileName: fileName, folderName: folderName },
    { extractedName: 'TEST', notes: ['manual test'] }
  );
}

function testProcessFilled() {
  return processFilledUnresolvedRows_();
}
