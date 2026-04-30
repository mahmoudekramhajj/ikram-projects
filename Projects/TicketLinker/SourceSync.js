/**
 * SourceSync.js — مزامنة من المجلدات الأصلية A/B (read-only) إلى C (TICKETS_FOLDER_ID)
 *
 * الفلسفة: كل التذاكر يجب أن تكون في ذاكرتي الخاصة بالمشروع، لا لدى شركة أخرى.
 *
 * التدفق:
 *   1. اقرأ TL_CopyLog مرة واحدة → set من sourceFileIds المنسوخة سابقاً.
 *   2. لكل مصدر في SOURCE_FOLDER_IDS:
 *      - أنشئ/اعثر على مجلد فرعي في C باسم <name>.
 *      - امشِ recursive على الشجرة، اجمع كل PDFs.
 *      - لكل PDF غير منسوخ سابقاً: makeCopy إلى C/<name>/ بـ prefix من اسم الأب.
 *      - سجّل في TL_CopyLog.
 *   3. يحترم MAX_RUNTIME_MS — يترك البقية للجولة التالية.
 *
 * idempotent: تشغيله مرتين متتاليتين لا ينسخ شيئاً جديداً.
 */

var TL_COPY_LOG_TAB = 'TL_CopyLog';
var TL_SYNC_MAX_DEPTH = 5;  // حماية من أشجار عميقة

// ==================== شيت السجل ====================

function ensureCopyLogTab_() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL_COPY_LOG_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(TL_COPY_LOG_TAB);
    sheet.appendRow([
      'Timestamp',           // A
      'SourceFileId',        // B - مفتاح idempotency
      'CopiedFileId',        // C
      'FileName',            // D - اسم الملف بعد النسخ (مع prefix)
      'SourceRoot',          // E - A_inbox / B_inbox
      'SourceParentFolder',  // F - اسم المجلد الأب في A/B
      'Status',              // G - OK / SKIP_DUP / ERROR
      'Error'                // H
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * يقرأ كل sourceFileIds المنسوخة سابقاً (Status=OK) من TL_CopyLog
 * @return {Object} {sourceFileId: true}
 */
function loadCopiedSourceIds_() {
  var sheet = ensureCopyLogTab_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var data = sheet.getRange(2, 2, lastRow - 1, 6).getValues(); // B..G
  var copied = {};
  for (var i = 0; i < data.length; i++) {
    var sid = String(data[i][0] || '').trim();      // B
    var status = String(data[i][5] || '').trim();   // G
    if (!sid) continue;
    if (status === 'OK' || status === 'SKIP_DUP') copied[sid] = true;
  }
  return copied;
}

// ==================== مجلد الوجهة ====================

/**
 * يضمن وجود مجلد فرعي باسم <name> داخل C، ويعيده.
 */
function ensureTargetSubfolder_(targetName) {
  var parent = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var subs = parent.getFoldersByName(targetName);
  if (subs.hasNext()) return subs.next();
  return parent.createFolder(targetName);
}

// ==================== جمع PDFs بشكل recursive ====================

/**
 * يجمع كل PDFs من شجرة مجلد بشكل recursive، مع اسم الأب المباشر.
 * @param {Folder} folder
 * @param {string} parentPath - مسار الأب (لـ prefix)
 * @param {number} depth
 * @param {Array} acc - مجمع النتائج
 * @param {number} deadline - timestamp ms
 */
function collectPdfsRecursive_(folder, parentPath, depth, acc, deadline) {
  if (Date.now() >= deadline) return;
  if (depth > TL_SYNC_MAX_DEPTH) return;

  // PDFs في هذا المجلد
  var pdfs = folder.getFilesByType(MimeType.PDF);
  while (pdfs.hasNext() && Date.now() < deadline) {
    var pdf = pdfs.next();
    acc.push({
      file: pdf,
      fileId: pdf.getId(),
      fileName: pdf.getName(),
      parentName: folder.getName(),
      parentPath: parentPath
    });
  }

  // مجلدات فرعية
  var subs = folder.getFolders();
  while (subs.hasNext() && Date.now() < deadline) {
    var sub = subs.next();
    var newPath = parentPath ? (parentPath + '/' + sub.getName()) : sub.getName();
    collectPdfsRecursive_(sub, newPath, depth + 1, acc, deadline);
  }
}

// ==================== الدالة الرئيسية ====================

/**
 * المزامنة الرئيسية — تُستدعى من scheduledProcessNewTickets قبل المعالجة.
 * @return {Object} ملخّص
 */
function syncSourceFoldersToTickets_() {
  var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;
  var summary = {
    sources: [],
    totalScanned: 0,
    totalCopied: 0,
    totalSkipped: 0,
    totalErrors: 0,
    timedOut: false
  };

  var alreadyCopied = loadCopiedSourceIds_();
  var logSheet = ensureCopyLogTab_();
  var newRows = [];

  for (var s = 0; s < TL.Config.SOURCE_FOLDER_IDS.length; s++) {
    if (Date.now() >= deadline) {
      summary.timedOut = true;
      break;
    }
    var src = TL.Config.SOURCE_FOLDER_IDS[s];
    var srcSummary = {
      name: src.name,
      scanned: 0,
      copied: 0,
      skipped: 0,
      errors: 0
    };

    var sourceFolder, targetFolder;
    try {
      sourceFolder = DriveApp.getFolderById(src.id);
      targetFolder = ensureTargetSubfolder_(src.name);
    } catch (e) {
      srcSummary.fatalError = e.message;
      summary.sources.push(srcSummary);
      continue;
    }

    // اجمع كل PDFs (recursive)
    var collected = [];
    collectPdfsRecursive_(sourceFolder, '', 0, collected, deadline);
    srcSummary.scanned = collected.length;
    summary.totalScanned += collected.length;

    // انسخ ما لم يُنسخ سابقاً
    for (var i = 0; i < collected.length; i++) {
      if (Date.now() >= deadline) {
        summary.timedOut = true;
        break;
      }
      var item = collected[i];

      if (alreadyCopied[item.fileId]) {
        srcSummary.skipped++;
        summary.totalSkipped++;
        continue;
      }

      // اسم الملف بعد النسخ: <parentName>__<originalName> لتجنّب التصادم
      var newName = item.parentName && item.parentName !== sourceFolder.getName()
        ? (item.parentName + '__' + item.fileName)
        : item.fileName;

      try {
        var copy = item.file.makeCopy(newName, targetFolder);
        newRows.push([
          new Date(),
          item.fileId,
          copy.getId(),
          newName,
          src.name,
          item.parentPath || item.parentName,
          'OK',
          ''
        ]);
        alreadyCopied[item.fileId] = true; // حماية intra-run
        srcSummary.copied++;
        summary.totalCopied++;
      } catch (e) {
        newRows.push([
          new Date(),
          item.fileId,
          '',
          newName,
          src.name,
          item.parentPath || item.parentName,
          'ERROR',
          e.message
        ]);
        srcSummary.errors++;
        summary.totalErrors++;
      }
    }

    summary.sources.push(srcSummary);
  }

  // كتابة دفعة واحدة في TL_CopyLog (أكفأ من appendRow متكرر)
  if (newRows.length > 0) {
    var startRow = logSheet.getLastRow() + 1;
    logSheet.getRange(startRow, 1, newRows.length, 8).setValues(newRows);
  }

  Logger.log('[Sync] ' + JSON.stringify(summary));
  return summary;
}

// ==================== اختبار يدوي ====================

/**
 * تشغيل يدوي للمزامنة (يُستدعى عبر Web App للتحقق)
 */
function runSyncManual() {
  return syncSourceFoldersToTickets_();
}
