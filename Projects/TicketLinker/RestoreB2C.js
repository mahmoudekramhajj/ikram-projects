/**
 * RestoreB2C.js — استرجاع عمود Z في B2C لحالته قبل تشغيل TicketLinker
 *
 * المنطق:
 *   1. اقرأ كل صف في TL_RunLog حيث Status ∈ {REPLACED, WRITTEN}
 *   2. لكل (passport, oldZ) → اكتب oldZ مرة أخرى في B2C (passport-keyed)
 *   3. سجّل النتائج في TL_RestoreLog
 *
 * Idempotent: لو B2C الآن = oldZ، تجاوز
 */

// ==================== Helpers ====================

function _getRunLogRows() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_RunLog');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // الأعمدة: Timestamp, Passport, FileId, FileName, Stage, Status, OldZ, NewZ, RowIndex, Notes
  var data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  return data.map(function(r) {
    return {
      timestamp: r[0],
      passport: String(r[1] || '').trim(),
      fileId: String(r[2] || ''),
      fileName: String(r[3] || ''),
      stage: r[4],
      status: String(r[5] || ''),
      oldZ: String(r[6] || ''),
      newZ: String(r[7] || ''),
      rowIndex: r[8],
      notes: String(r[9] || '')
    };
  });
}

function _restoreOneRow(sheet, passportCol, ticketUrlCol, passport, oldZ, currentNewZ) {
  var rowIndex = findRowByPassport_(sheet, passport, passportCol);
  if (!rowIndex) return { status: 'PASSPORT_NOT_FOUND', passport: passport };

  var currentZ = String(sheet.getRange(rowIndex, ticketUrlCol).getValue() || '').trim();

  // لو الـ Z الحالي ليس قيمتنا (تغيّر بعد كتابتنا) → لا نسترجع
  if (currentZ !== currentNewZ) {
    return {
      status: 'SKIPPED_CHANGED_AFTER',
      passport: passport,
      currentZ: currentZ,
      expectedNewZ: currentNewZ
    };
  }

  // Idempotent: لو oldZ نفس currentZ، تجاوز
  if (currentZ === oldZ) {
    return { status: 'ALREADY_RESTORED', passport: passport };
  }

  sheet.getRange(rowIndex, ticketUrlCol).setValue(oldZ);
  return {
    status: 'RESTORED',
    passport: passport,
    rowIndex: rowIndex,
    restoredTo: oldZ
  };
}

// ==================== Main ====================

/**
 * استرجاع B2C كاملاً من TL_RunLog
 */
function restoreB2cFromRunLog() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var b2cSheet = ss.getSheetByName(TL.Config.SHEET_B2C);
  if (!b2cSheet) throw new Error('B2C sheet not found');

  var passportCol = findColumnByHeader_(b2cSheet, TL.Config.B2C_HEADERS.PASSPORT, TL.Config.B2C_COL.PASSPORT);
  var ticketUrlCol = findColumnByHeader_(b2cSheet, TL.Config.B2C_HEADERS.TICKET_URL, TL.Config.B2C_COL.TICKET_URL);
  if (!passportCol || !ticketUrlCol) throw new Error('Headers not found in B2C');

  var rows = _getRunLogRows();
  // فقط الصفوف التي كتبنا فيها فعلاً (لها fileId/newZ)
  var writtenRows = rows.filter(function(r) {
    return r.passport && r.newZ &&
           (r.status === 'REPLACED' || r.status === 'WRITTEN' || r.status === 'OK');
  });

  // dedupe بـ passport (آخر كتابة لكل حاج)
  var byPassport = {};
  writtenRows.forEach(function(r) {
    byPassport[r.passport] = r;  // آخر صف يفوز
  });
  var uniqueRows = Object.keys(byPassport).map(function(k) { return byPassport[k]; });

  var stats = { total: uniqueRows.length, restored: 0, alreadyRestored: 0, skippedChanged: 0, notFound: 0, errors: 0 };
  var errors = [];

  // إنشاء/فتح RestoreLog
  var restoreLog = ss.getSheetByName('TL_RestoreLog');
  if (!restoreLog) {
    restoreLog = ss.insertSheet('TL_RestoreLog');
    restoreLog.appendRow(['Timestamp', 'Passport', 'Status', 'RowIndex', 'OldZ_Restored', 'CurrentZ_Before', 'Notes']);
  }

  var batch = [];
  for (var i = 0; i < uniqueRows.length; i++) {
    var row = uniqueRows[i];
    try {
      var result = _restoreOneRow(b2cSheet, passportCol, ticketUrlCol, row.passport, row.oldZ, row.newZ);

      if (result.status === 'RESTORED') stats.restored++;
      else if (result.status === 'ALREADY_RESTORED') stats.alreadyRestored++;
      else if (result.status === 'SKIPPED_CHANGED_AFTER') stats.skippedChanged++;
      else if (result.status === 'PASSPORT_NOT_FOUND') stats.notFound++;

      batch.push([
        new Date(),
        row.passport,
        result.status,
        result.rowIndex || '',
        row.oldZ,
        result.currentZ || '',
        result.expectedNewZ ? ('expected:' + result.expectedNewZ) : ''
      ]);
    } catch (e) {
      stats.errors++;
      errors.push({ passport: row.passport, error: String(e) });
      batch.push([new Date(), row.passport, 'ERROR', '', row.oldZ, '', String(e)]);
    }

    // كتابة على دفعات لتقليل API calls
    if (batch.length >= 50) {
      restoreLog.getRange(restoreLog.getLastRow() + 1, 1, batch.length, 7).setValues(batch);
      batch = [];
    }
  }

  if (batch.length) {
    restoreLog.getRange(restoreLog.getLastRow() + 1, 1, batch.length, 7).setValues(batch);
  }

  return {
    stats: stats,
    errors: errors.slice(0, 10),
    timestamp: new Date().toISOString()
  };
}

/**
 * dry-run: يظهر ماذا سيُغيَّر بدون تنفيذ
 */
function previewRestore() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var b2cSheet = ss.getSheetByName(TL.Config.SHEET_B2C);
  var passportCol = findColumnByHeader_(b2cSheet, TL.Config.B2C_HEADERS.PASSPORT, TL.Config.B2C_COL.PASSPORT);
  var ticketUrlCol = findColumnByHeader_(b2cSheet, TL.Config.B2C_HEADERS.TICKET_URL, TL.Config.B2C_COL.TICKET_URL);

  var rows = _getRunLogRows();
  var writtenRows = rows.filter(function(r) {
    return r.passport && r.newZ &&
           (r.status === 'REPLACED' || r.status === 'WRITTEN' || r.status === 'OK');
  });

  var byPassport = {};
  writtenRows.forEach(function(r) { byPassport[r.passport] = r; });
  var uniqueRows = Object.keys(byPassport).map(function(k) { return byPassport[k]; });

  // عيّنة 5 صفوف لمعاينة
  var sample = uniqueRows.slice(0, 5).map(function(r) {
    var rowIndex = findRowByPassport_(b2cSheet, r.passport, passportCol);
    var currentZ = rowIndex ? String(b2cSheet.getRange(rowIndex, ticketUrlCol).getValue() || '').trim() : '';
    return {
      passport: r.passport,
      rowIndex: rowIndex,
      currentZ: currentZ,
      willRestoreTo: r.oldZ,
      changed: currentZ !== r.newZ
    };
  });

  return {
    totalToProcess: uniqueRows.length,
    sample: sample
  };
}
