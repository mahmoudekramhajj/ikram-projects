/**
 * Pipeline.js — السلّم الكامل من PDF إلى Z في B2C
 *
 * يربط المراحل:
 *   PnrResolver → Matcher (1+3) → ClaudeMatcher (4) → UnresolvedHandler (5) → Writer
 */

/**
 * يعالج PDF واحداً من البداية للنهاية
 */
function processSinglePdf_(pdfFile, folderName) {
  var pdfInfo = {
    fileId: pdfFile.getId(),
    fileName: pdfFile.getName(),
    folderName: folderName
  };

  // M1: استخراج PNR
  var pnrInfo = resolvePnr_(pdfInfo);
  pdfInfo.pnr = pnrInfo.pnr;

  // M2: السلّم الأساسي (المرحلة 1 + المرحلة 3)
  var match = matchPilgrimForPdf_(pdfInfo);

  // لو نجح في 1 أو 3 → اكتب
  if (match.status === 'OK' && match.passport) {
    var writeResult = writeTicketUrlForPassport_(match.passport, pdfInfo.fileId);
    logWriteOperation_({
      passport: match.passport,
      fileId: pdfInfo.fileId,
      fileName: pdfInfo.fileName,
      stage: match.stage,
      status: writeResult.status,
      oldZ: writeResult.oldZ,
      newZ: writeResult.newZ,
      rowIndex: writeResult.rowIndex,
      notes: 'matched: ' + match.matchedFirstName + ' ' + match.matchedLastName + ' (key: ' + match.matchedKey + ')'
    });

    // حالة نادرة: passport من Matcher غير موجود في PD (race condition)
    if (writeResult.status === 'PASSPORT_NOT_FOUND_IN_PD') {
      var unresolvedPd = registerUnresolved_(pdfInfo, {
        extractedName: match.extractedName,
        notes: ['RACE CONDITION: passport ' + match.passport + ' was found by Matcher but NOT IN PD at write time.']
      });
      return {
        fileName: pdfInfo.fileName,
        finalStage: match.stage,
        finalStatus: 'UNRESOLVED_RACE_CONDITION',
        passport: match.passport
      };
    }

    return {
      fileName: pdfInfo.fileName,
      finalStage: match.stage,
      finalStatus: writeResult.status,
      passport: match.passport
    };
  }

  // لو COLLISION → سجّل في Unresolved مباشرة (لا تخمّن)
  if (match.status === 'COLLISION') {
    var collisionResult = registerUnresolved_(pdfInfo, {
      extractedName: match.extractedName,
      notes: ['COLLISION at stage ' + match.stage + ': ' +
              JSON.stringify(match.collision.map(function(c) { return c.passport + '/' + c.firstName + ' ' + c.lastName; }))]
    });
    logWriteOperation_({
      fileId: pdfInfo.fileId,
      fileName: pdfInfo.fileName,
      stage: match.stage,
      status: 'UNRESOLVED_COLLISION',
      notes: collisionResult.status
    });
    return { fileName: pdfInfo.fileName, finalStage: match.stage, finalStatus: 'UNRESOLVED_COLLISION' };
  }

  // المرحلة 4: Claude
  if (match.status === 'GO_TO_STAGE_4') {
    var claudeMatch;
    try {
      claudeMatch = matchViaPdfClaudeExtraction_(pdfInfo);
    } catch (e) {
      claudeMatch = { status: 'CLAUDE_ERROR', notes: ['Exception: ' + e.message] };
    }

    if (claudeMatch.status === 'OK' && claudeMatch.passport) {
      var writeResult2 = writeTicketUrlForPassport_(claudeMatch.passport, pdfInfo.fileId);
      logWriteOperation_({
        passport: claudeMatch.passport,
        fileId: pdfInfo.fileId,
        fileName: pdfInfo.fileName,
        stage: 4,
        status: writeResult2.status,
        oldZ: writeResult2.oldZ,
        newZ: writeResult2.newZ,
        rowIndex: writeResult2.rowIndex,
        notes: 'matched via Claude: ' + (claudeMatch.notes || []).join(' | ')
      });
      return {
        fileName: pdfInfo.fileName,
        finalStage: 4,
        finalStatus: writeResult2.status,
        passport: claudeMatch.passport
      };
    }

    // المرحلة 5: Unresolved
    var unresolvedResult = registerUnresolved_(pdfInfo, {
      extractedName: match.extractedName,
      claudeData: claudeMatch.claudeData,
      notes: (match.notes || []).concat(claudeMatch.notes || [])
    });
    logWriteOperation_({
      fileId: pdfInfo.fileId,
      fileName: pdfInfo.fileName,
      stage: 5,
      status: 'UNRESOLVED',
      notes: unresolvedResult.status + ' | claude: ' + (claudeMatch.status || 'N/A')
    });
    return { fileName: pdfInfo.fileName, finalStage: 5, finalStatus: 'UNRESOLVED' };
  }

  // حالة غير متوقعة
  logWriteOperation_({
    fileId: pdfInfo.fileId,
    fileName: pdfInfo.fileName,
    stage: match.stage,
    status: match.status,
    notes: 'unexpected status'
  });
  return { fileName: pdfInfo.fileName, finalStage: match.stage, finalStatus: match.status };
}

// ==================== التشغيل ====================

/**
 * يعالج كل PDFs في مجلد فرعي
 */
function processFolder(folderName) {
  if (!acquireLock_()) return { error: 'Another run is in progress' };

  try {
    var folder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
    var subFolders = folder.getFolders();
    var targetFolder = null;

    while (subFolders.hasNext()) {
      var sub = subFolders.next();
      if (sub.getName().toUpperCase() === folderName.toUpperCase() ||
          sub.getName().indexOf(folderName) !== -1) {
        targetFolder = sub;
        break;
      }
    }
    if (!targetFolder) return { error: 'Folder not found: ' + folderName };

    var pdfs = targetFolder.getFilesByType(MimeType.PDF);
    var results = [];
    var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;

    while (pdfs.hasNext() && Date.now() < deadline) {
      var pdf = pdfs.next();
      try {
        var r = processSinglePdf_(pdf, targetFolder.getName());
        results.push(r);
      } catch (e) {
        results.push({
          fileName: pdf.getName(),
          finalStatus: 'EXCEPTION',
          error: e.message
        });
      }
    }

    var stats = computeStats_(results);

    return {
      folder: targetFolder.getName(),
      stats: stats,
      results: results
    };
  } finally {
    releaseLock_();
  }
}

function computeStats_(results) {
  return {
    total: results.length,
    written: results.filter(function(r) { return r.finalStatus === 'WRITTEN'; }).length,
    replaced: results.filter(function(r) { return r.finalStatus === 'REPLACED'; }).length,
    already: results.filter(function(r) { return r.finalStatus === 'ALREADY_WRITTEN'; }).length,
    unresolved: results.filter(function(r) { return r.finalStatus === 'UNRESOLVED'; }).length,
    raceCondition: results.filter(function(r) { return r.finalStatus === 'UNRESOLVED_RACE_CONDITION'; }).length,
    collision: results.filter(function(r) { return r.finalStatus === 'UNRESOLVED_COLLISION'; }).length,
    other: results.filter(function(r) {
      return ['WRITTEN', 'REPLACED', 'ALREADY_WRITTEN', 'UNRESOLVED',
              'UNRESOLVED_COLLISION', 'UNRESOLVED_RACE_CONDITION'].indexOf(r.finalStatus) === -1;
    }).length
  };
}

/**
 * يعالج كل المجلدات الفرعية في المجلد الرئيسي
 * مع timeout safety + cursor للاستمرار
 *
 * استراتيجية: نعالج مجلداً مجلداً، ونحفظ في cursor آخر مجلد مكتمل.
 * في الاستدعاء التالي نتجاوزه. لو انتهت المجلدات نمسح cursor.
 */
function processAllFolders() {
  if (!acquireLock_()) return { error: 'Another run is in progress' };

  var props = PropertiesService.getScriptProperties();
  var processedSoFar = JSON.parse(props.getProperty('TL_PROCESSED_FOLDERS') || '[]');
  var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;
  var summary = {
    foldersProcessedThisRun: [],
    foldersSkipped: [],
    totalPdfs: 0,
    totalReplaced: 0,
    totalWritten: 0,
    totalAlready: 0,
    totalUnresolved: 0,
    totalRaceCondition: 0,
    totalCollisions: 0,
    totalErrors: 0,
    completed: false
  };

  try {
    var rootFolder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
    var subFolders = rootFolder.getFolders();
    var allFolders = [];
    while (subFolders.hasNext()) {
      var f = subFolders.next();
      // تجاوز مجلد Unresolved_Tickets
      if (f.getName() === 'Unresolved_Tickets') continue;
      allFolders.push({ id: f.getId(), name: f.getName() });
    }

    summary.totalFolders = allFolders.length;
    summary.alreadyDone = processedSoFar.length;

    for (var i = 0; i < allFolders.length; i++) {
      var folder = allFolders[i];

      // تجاوز المُعالَج
      if (processedSoFar.indexOf(folder.id) !== -1) {
        summary.foldersSkipped.push(folder.name);
        continue;
      }

      // فحص الوقت
      if (Date.now() >= deadline) {
        Logger.log('Deadline reached, saving cursor');
        break;
      }

      // عالج المجلد (نسخة inline من processFolder بدون lock مزدوج)
      try {
        var folderResult = processFolderInternal_(folder.id, folder.name, deadline);
        summary.foldersProcessedThisRun.push({
          name: folder.name,
          stats: folderResult.stats
        });
        summary.totalPdfs += folderResult.stats.total;
        summary.totalReplaced += folderResult.stats.replaced;
        summary.totalWritten += folderResult.stats.written;
        summary.totalAlready += folderResult.stats.already;
        summary.totalUnresolved += folderResult.stats.unresolved;
        summary.totalRaceCondition += folderResult.stats.raceCondition || 0;
        summary.totalCollisions += folderResult.stats.collision;

        // أضف للـ processed
        processedSoFar.push(folder.id);
      } catch (e) {
        summary.totalErrors++;
        Logger.log('Error processing folder ' + folder.name + ': ' + e.message);
      }
    }

    // لو أكملنا كل المجلدات → امسح cursor
    if (processedSoFar.length >= allFolders.length) {
      props.deleteProperty('TL_PROCESSED_FOLDERS');
      summary.completed = true;
    } else {
      props.setProperty('TL_PROCESSED_FOLDERS', JSON.stringify(processedSoFar));
    }

    summary.totalProcessedSoFar = processedSoFar.length;
    return summary;
  } finally {
    releaseLock_();
  }
}

/**
 * نسخة داخلية من processFolder بدون lock (للاستخدام داخل processAllFolders)
 * @param {string} folderId
 * @param {string} folderName
 * @param {number} deadline
 * @param {Object} [skipFileIds] - خريطة {fileId: true} للملفات المرتبطة مسبقاً — تُتجاوز
 */
function processFolderInternal_(folderId, folderName, deadline, skipFileIds) {
  var folder = DriveApp.getFolderById(folderId);
  var pdfs = folder.getFilesByType(MimeType.PDF);
  var results = [];

  while (pdfs.hasNext() && Date.now() < deadline) {
    var pdf = pdfs.next();

    // تجاوز الملفات المرتبطة مسبقاً (بناءً على fileId)
    if (skipFileIds && skipFileIds[pdf.getId()]) {
      results.push({
        fileName: pdf.getName(),
        finalStatus: 'ALREADY_WRITTEN',
        skipped: true
      });
      continue;
    }

    try {
      var r = processSinglePdf_(pdf, folderName);
      results.push(r);
    } catch (e) {
      results.push({
        fileName: pdf.getName(),
        finalStatus: 'EXCEPTION',
        error: e.message
      });
    }
  }

  return {
    folder: folderName,
    stats: computeStats_(results)
  };
}

/**
 * يقرأ fileIds من مصدرين ويبني خريطة {fileId: true}:
 *   1. عمود Z في PD (نجاحات الكتابة) — Drive URLs فقط
 *   2. TL_UnresolvedTickets عمود B (Tombstone للفاشلين الدائمين)
 * يُستخدم لتجاوز الملفات المعالَجة (نجاحاً أو فشلاً) في الجولات التدريجية،
 * لمنع نزيف Claude API على الفاشلين الدائمين.
 * @return {Object} {fileId: true}
 */
function loadProcessedFileIds_() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var processed = {};

  // ===== المصدر 1: PD/Z (نجاحات Drive) =====
  var pdSheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (pdSheet) {
    var ticketUrlCol = (function() {
      var lastCol = pdSheet.getLastColumn();
      var headers = pdSheet.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i]).trim() === TL.Config.PD_HEADERS.TICKET_URL) return i + 1;
      }
      return TL.Config.PD_COL.TICKET_URL;
    })();

    var pdLastRow = pdSheet.getLastRow();
    if (pdLastRow >= 2) {
      var urlData = pdSheet.getRange(2, ticketUrlCol, pdLastRow - 1, 1).getValues();
      var drivePattern = /\/d\/([a-zA-Z0-9_-]+)\//;
      for (var i = 0; i < urlData.length; i++) {
        var url = String(urlData[i][0] || '').trim();
        if (!url) continue;
        var m = url.match(drivePattern);
        if (m && m[1]) processed[m[1]] = true;
      }
    }
  }

  // ===== المصدر 2: TL_UnresolvedTickets عمود B (Tombstone) =====
  // يحمي من إعادة معالجة fileIds فشلت سابقاً (NAME_NOT_IN_PD/COLLISION/CLAUDE_ERROR/...).
  // أي صف status فيه (PENDING_MANUAL/LINKED/NO_MATCH_AFTER_FILL) → tombstone.
  var unresolvedSheet = ss.getSheetByName(TL_UNRESOLVED_TAB);
  if (unresolvedSheet) {
    var uLastRow = unresolvedSheet.getLastRow();
    if (uLastRow >= 2) {
      // اقرأ B (fileId) و N (status) معاً
      var rng = unresolvedSheet.getRange(2, 2, uLastRow - 1, 13).getValues();
      for (var j = 0; j < rng.length; j++) {
        var fid = String(rng[j][0] || '').trim();    // عمود B
        var st = String(rng[j][12] || '').trim();    // عمود N (offset 12 من B)
        if (!fid) continue;
        // أي status معروف يُعتبر tombstone؛ فارغ = صف معطوب نتجاهله
        if (st === 'PENDING_MANUAL' || st === 'LINKED' || st === 'NO_MATCH_AFTER_FILL') {
          processed[fid] = true;
        }
      }
    }
  }

  return processed;
}

/**
 * معالجة تدريجية — يُعيد الـ cursor ويتجاوز الملفات المرتبطة مسبقاً
 * الاستخدام: للتشغيل الدوري عند إضافة تذاكر جديدة للـ Drive
 */
function processAllFoldersIncremental() {
  if (!acquireLock_()) return { error: 'Another run is in progress' };

  try {
    // أعد الـ cursor لتفحص كل المجلدات من جديد
    PropertiesService.getScriptProperties().deleteProperty('TL_PROCESSED_FOLDERS');

    // حمّل fileIds المرتبطة مسبقاً (من عمود Z في PD)
    var skipFileIds = loadProcessedFileIds_();
    var skippedTotal = Object.keys(skipFileIds).length;
    Logger.log('[Incremental] Loaded ' + skippedTotal + ' already-linked fileIds to skip');

    var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;
    var summary = {
      mode: 'incremental',
      alreadyLinked: skippedTotal,
      foldersProcessedThisRun: [],
      totalNew: 0,
      totalSkipped: 0,
      totalUnresolved: 0,
      totalErrors: 0,
      completed: false
    };

    var rootFolder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
    var subFolders = rootFolder.getFolders();
    var allFolders = [];
    var excluded = TL.Config.EXCLUDED_FOLDERS || ['Unresolved_Tickets'];
    while (subFolders.hasNext()) {
      var f = subFolders.next();
      if (excluded.indexOf(f.getName()) !== -1) continue;
      allFolders.push({ id: f.getId(), name: f.getName() });
    }

    for (var i = 0; i < allFolders.length; i++) {
      if (Date.now() >= deadline) break;
      var folder = allFolders[i];

      try {
        var folderResult = processFolderInternal_(folder.id, folder.name, deadline, skipFileIds);
        var st = folderResult.stats;
        var newInFolder = st.written + st.replaced;
        summary.totalNew += newInFolder;
        summary.totalSkipped += st.already;
        summary.totalUnresolved += st.unresolved;

        if (newInFolder > 0) {
          summary.foldersProcessedThisRun.push({ name: folder.name, new: newInFolder });
        }
      } catch (e) {
        summary.totalErrors++;
        Logger.log('[Incremental] Error in folder ' + folder.name + ': ' + e.message);
      }
    }

    summary.completed = true;
    Logger.log('[Incremental] Done — new=' + summary.totalNew + ' skipped=' + summary.totalSkipped);
    return summary;

  } finally {
    releaseLock_();
  }
}

/**
 * دالة التشغيل الدوري — تُستدعى من التريجر
 * التدفق: SourceSync (A/B → C) ثم Pipeline incremental على C.
 */
function scheduledProcessNewTickets() {
  if (!acquireLock_()) {
    Logger.log('[Scheduled] Another run in progress, skipping');
    return;
  }
  releaseLock_(); // releaseLock_ لأن syncSourceFoldersToTickets_ + processAllFoldersIncremental كلاهما سيحصل على lock داخلياً

  // 1. مزامنة من A/B إلى C
  var syncResult;
  try {
    syncResult = syncSourceFoldersToTickets_();
  } catch (e) {
    syncResult = { error: e.message };
    Logger.log('[Scheduled] Sync error: ' + e.message);
  }

  // 2. معالجة C تدريجياً (مع tombstone)
  var processResult = processAllFoldersIncremental();

  var combined = { sync: syncResult, process: processResult };
  Logger.log('[Scheduled] ' + JSON.stringify(combined));
  return combined;
}

/**
 * إعادة تعيين cursor (للاستخدام لو أردت إعادة المعالجة من الصفر)
 */
function resetProcessedCursor() {
  PropertiesService.getScriptProperties().deleteProperty('TL_PROCESSED_FOLDERS');
  return { ok: true };
}

/**
 * Worker: يعالج مجلداً واحداً بالاسم (للتشغيل المتوازي)
 * يستخدم LockService لحماية الـ Script Properties من race conditions.
 * @param {string} folderName
 * @return {Object}
 */
function workerProcessFolder(folderName) {
  var rootFolder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var subFolders = rootFolder.getFolders();
  var targetFolder = null;
  while (subFolders.hasNext()) {
    var sub = subFolders.next();
    if (sub.getName() === folderName) { targetFolder = sub; break; }
  }
  if (!targetFolder) return { error: 'Folder not found: ' + folderName };

  var folderId = targetFolder.getId();
  var lock = LockService.getScriptLock();

  // === Phase 1: claim (فحص + حجز بـ lock) ===
  if (!lock.tryLock(30000)) {
    return { folder: folderName, status: 'LOCK_TIMEOUT' };
  }
  try {
    var props = PropertiesService.getScriptProperties();
    var processed = JSON.parse(props.getProperty('TL_PROCESSED_FOLDERS') || '[]');
    var inProgress = JSON.parse(props.getProperty('TL_IN_PROGRESS') || '[]');

    if (processed.indexOf(folderId) !== -1) {
      return { folder: folderName, status: 'ALREADY_PROCESSED' };
    }
    if (inProgress.indexOf(folderId) !== -1) {
      return { folder: folderName, status: 'IN_PROGRESS_BY_ANOTHER_WORKER' };
    }
    inProgress.push(folderId);
    props.setProperty('TL_IN_PROGRESS', JSON.stringify(inProgress));
  } finally {
    lock.releaseLock();
  }

  // === Phase 2: المعالجة (طويلة، بدون lock) ===
  var processResult;
  try {
    var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;
    processResult = processFolderInternal_(folderId, folderName, deadline);
  } catch (e) {
    processResult = { error: e.message, stats: {} };
  }

  // === Phase 3: mark as processed + release in-progress (atomic بـ lock) ===
  if (!lock.tryLock(30000)) {
    return { folder: folderName, status: 'LOCK_TIMEOUT_FINAL', stats: processResult.stats };
  }
  try {
    var props2 = PropertiesService.getScriptProperties();

    // أضف للـ processed
    var processed2 = JSON.parse(props2.getProperty('TL_PROCESSED_FOLDERS') || '[]');
    if (processed2.indexOf(folderId) === -1) processed2.push(folderId);
    props2.setProperty('TL_PROCESSED_FOLDERS', JSON.stringify(processed2));

    // احذف من in-progress
    var ip = JSON.parse(props2.getProperty('TL_IN_PROGRESS') || '[]');
    var idx = ip.indexOf(folderId);
    if (idx !== -1) {
      ip.splice(idx, 1);
      props2.setProperty('TL_IN_PROGRESS', JSON.stringify(ip));
    }
  } finally {
    lock.releaseLock();
  }

  return {
    folder: folderName,
    status: processResult.error ? 'ERROR' : 'DONE',
    stats: processResult.stats || {},
    error: processResult.error
  };
}

/**
 * يُرجع قائمة المجلدات المتبقية للمعالجة (غير معالَجة + غير قيد معالجة)
 */
function getPendingFolderNames() {
  var props = PropertiesService.getScriptProperties();
  var processed = JSON.parse(props.getProperty('TL_PROCESSED_FOLDERS') || '[]');
  var inProgress = JSON.parse(props.getProperty('TL_IN_PROGRESS') || '[]');

  var rootFolder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var subFolders = rootFolder.getFolders();
  var pending = [];
  while (subFolders.hasNext()) {
    var f = subFolders.next();
    if (f.getName() === 'Unresolved_Tickets') continue;
    var id = f.getId();
    if (processed.indexOf(id) === -1 && inProgress.indexOf(id) === -1) {
      pending.push(f.getName());
    }
  }
  return { count: pending.length, pending: pending };
}

/**
 * يمسح قائمة in-progress (إذا تعطل worker)
 */
function clearInProgress() {
  PropertiesService.getScriptProperties().deleteProperty('TL_IN_PROGRESS');
  return { ok: true };
}

/**
 * عرض حالة الـ cursor
 */
function getProcessingCursor() {
  var props = PropertiesService.getScriptProperties();
  var processed = JSON.parse(props.getProperty('TL_PROCESSED_FOLDERS') || '[]');
  var rootFolder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var subFolders = rootFolder.getFolders();
  var allFolders = [];
  while (subFolders.hasNext()) {
    var f = subFolders.next();
    if (f.getName() === 'Unresolved_Tickets') continue;
    allFolders.push({ id: f.getId(), name: f.getName() });
  }

  var processedNames = allFolders
    .filter(function(f) { return processed.indexOf(f.id) !== -1; })
    .map(function(f) { return f.name; });
  var pendingNames = allFolders
    .filter(function(f) { return processed.indexOf(f.id) === -1; })
    .map(function(f) { return f.name; });

  return {
    totalFolders: allFolders.length,
    processedCount: processedNames.length,
    pendingCount: pendingNames.length,
    processedNames: processedNames,
    pendingNames: pendingNames
  };
}

/**
 * Dry-run: تشغيل المراحل بدون كتابة فعلية في B2C
 */
function dryRunFolder(folderName) {
  var folder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var subFolders = folder.getFolders();
  var targetFolder = null;

  while (subFolders.hasNext()) {
    var sub = subFolders.next();
    if (sub.getName().toUpperCase() === folderName.toUpperCase() ||
        sub.getName().indexOf(folderName) !== -1) {
      targetFolder = sub;
      break;
    }
  }
  if (!targetFolder) return { error: 'Folder not found: ' + folderName };

  var pnrCandidates = extractPnrsFromFolderName_(targetFolder.getName());
  var pnr = pnrCandidates.length > 0 ? pnrCandidates[0] : null;

  var pdfs = targetFolder.getFilesByType(MimeType.PDF);
  var results = [];
  while (pdfs.hasNext()) {
    var pdf = pdfs.next();
    var match = matchPilgrimForPdf_({
      fileId: pdf.getId(),
      fileName: pdf.getName(),
      folderName: targetFolder.getName(),
      pnr: pnr
    });
    results.push({
      fileName: match.fileName,
      stage: match.stage,
      status: match.status,
      passport: match.passport,
      matched: match.matchedFirstName ? (match.matchedFirstName + ' ' + match.matchedLastName) : null
    });
  }

  var stats = {
    total: results.length,
    stage1OK: results.filter(function(r) { return r.stage === 1 && r.status === 'OK'; }).length,
    stage3OK: results.filter(function(r) { return r.stage === 3 && r.status === 'OK'; }).length,
    needStage4: results.filter(function(r) { return r.status === 'GO_TO_STAGE_4'; }).length,
    collisions: results.filter(function(r) { return r.status === 'COLLISION'; }).length
  };

  return { folder: targetFolder.getName(), pnr: pnr, stats: stats };
}
