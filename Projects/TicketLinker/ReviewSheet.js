/**
 * ReviewSheet.js — شيت مراجعة بشرية للـ 96 صف المتبقي
 *
 * الدوال:
 *   buildReviewSheet()     — ينشئ/يحدّث شيت "مراجعة التذاكر" بأعمدة عربية واضحة
 *   syncReviewPassports()  — يقرأ الجوازات المُدخَلة ويكتبها في PD مباشرة
 */

var REVIEW_SHEET_NAME = 'مراجعة التذاكر';

// أعمدة شيت المراجعة (1-based)
var RV = {
  PDF_LINK:       1,   // A — رابط التذكرة
  FOLDER:         2,   // B — المجلد / العقد
  PDF_NAME:       3,   // C — اسم ملف PDF
  NAME_FROM_PDF:  4,   // D — الاسم المستخرج
  REASON:         5,   // E — السبب (Collision / لا تطابق / ...)
  CANDIDATES:     6,   // F — المرشحون (للـ collision)
  PASSPORT_INPUT: 7,   // G — الجواز الصحيح ← يملأه الموظف
  STATUS:         8,   // H — الحالة (فارغ / تم الربط / خطأ)
  UNRES_ROW:      9    // I — رقم صف TL_UnresolvedTickets (مخفي)
};

// ==================== بناء شيت المراجعة ====================

function buildReviewSheet() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);

  // تحميل بيانات للمطابقة
  var pdCache = loadPdCache_(ss);
  var flightMap = loadFlightMap_(ss);

  var unresolvedSheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!unresolvedSheet) throw new Error('TL_UnresolvedTickets not found');

  var rows = unresolvedSheet.getDataRange().getValues();

  // جمع الصفوف PENDING_MANUAL
  var pending = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][UNRES_COL.STATUS] || '').trim() === 'PENDING_MANUAL') {
      pending.push({ rowData: rows[i], sheetRow: i + 1 });
    }
  }

  if (pending.length === 0) {
    SpreadsheetApp.getUi().alert('لا توجد صفوف PENDING_MANUAL — انتهت المراجعة!');
    return { count: 0 };
  }

  Logger.log('[Review] عدد الصفوف المعلّقة: ' + pending.length);

  // إنشاء أو مسح شيت المراجعة
  var reviewSheet = ss.getSheetByName(REVIEW_SHEET_NAME);
  if (reviewSheet) {
    reviewSheet.clearContents();
    reviewSheet.clearFormats();
  } else {
    reviewSheet = ss.insertSheet(REVIEW_SHEET_NAME);
  }

  // رأس الجدول
  var header = [
    'التذكرة (فتح)',
    'المجلد / العقد',
    'اسم ملف PDF',
    'الاسم المستخرج',
    'السبب',
    'المرشحون',
    'الجواز الصحيح ✏️',
    'الحالة',
    'صف المصدر'
  ];
  reviewSheet.getRange(1, 1, 1, header.length).setValues([header]);

  // تنسيق الرأس
  var headerRange = reviewSheet.getRange(1, 1, 1, header.length);
  headerRange.setBackground('#232E64').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // بناء بيانات الصفوف
  var dataRows = [];
  for (var j = 0; j < pending.length; j++) {
    var row = pending[j].rowData;
    var sheetRow = pending[j].sheetRow;

    var originalFileId = String(row[UNRES_COL.ORIGINAL_FILE_ID] || '').trim();
    var copiedFileId   = String(row[UNRES_COL.COPIED_FILE_ID] || '').trim();
    var folderName     = String(row[UNRES_COL.FOLDER_NAME] || '').trim();
    var pdfName        = String(row[UNRES_COL.ORIGINAL_FILE_NAME] || '').trim();
    var extractedName  = String(row[UNRES_COL.EXTRACTED_NAME] || '').trim();
    var claudeName     = String(row[UNRES_COL.CLAUDE_FULL_NAME] || '').trim();
    var claudePnr      = String(row[UNRES_COL.CLAUDE_PNR] || '').trim();
    var allPassengers  = row[UNRES_COL.ALL_PASSENGERS];
    var existingNotes  = String(row[UNRES_COL.NOTES] || '').trim();

    // اسم للعرض
    var displayName = claudeName || extractedName || pdfName;

    // رابط PDF (يفضّل نسخة Unresolved، يُرجع للأصلي إن لم توجد)
    var linkId = copiedFileId || originalFileId;
    var pdfUrl = linkId ? 'https://drive.google.com/file/d/' + linkId + '/view' : '';

    // تصنيف السبب
    var reason = classifyReason_(existingNotes, allPassengers);

    // المرشحون (للـ collision)
    var candidates = '';
    if (reason.isCollision) {
      candidates = getCollisionCandidates_(claudePnr, displayName, allPassengers, flightMap, pdCache);
    }

    dataRows.push([
      pdfUrl ? '=HYPERLINK("' + pdfUrl + '","فتح PDF")' : 'لا رابط',
      folderName,
      pdfName,
      displayName,
      reason.label,
      candidates,
      '',        // الجواز الصحيح — فارغ للموظف
      '',        // الحالة
      sheetRow   // رقم صف TL_UnresolvedTickets
    ]);
  }

  // كتابة البيانات (باستثناء عمود A الذي يحتوي فورمولا)
  // نكتب كل عمود على حدة لأن عمود A فورمولا
  for (var k = 0; k < dataRows.length; k++) {
    var r = k + 2; // صف البيانات (يبدأ من 2)
    var rowArr = dataRows[k];

    // عمود A: فورمولا
    if (rowArr[0] && String(rowArr[0]).indexOf('HYPERLINK') !== -1) {
      reviewSheet.getRange(r, 1).setFormula(rowArr[0].replace(/^=/, ''));
    } else {
      reviewSheet.getRange(r, 1).setValue(rowArr[0]);
    }

    // الأعمدة B-I: قيم عادية
    reviewSheet.getRange(r, 2, 1, rowArr.length - 1).setValues([rowArr.slice(1)]);
  }

  // تنسيق عمود الجواز (G) — تظليل أصفر للفت الانتباه
  if (dataRows.length > 0) {
    reviewSheet.getRange(2, RV.PASSPORT_INPUT, dataRows.length, 1)
      .setBackground('#fff2cc').setFontWeight('bold');

    // إخفاء عمود المصدر (I)
    reviewSheet.hideColumns(RV.UNRES_ROW);

    // عرض الأعمدة
    reviewSheet.setColumnWidth(RV.PDF_LINK, 90);
    reviewSheet.setColumnWidth(RV.FOLDER, 140);
    reviewSheet.setColumnWidth(RV.PDF_NAME, 200);
    reviewSheet.setColumnWidth(RV.NAME_FROM_PDF, 180);
    reviewSheet.setColumnWidth(RV.REASON, 120);
    reviewSheet.setColumnWidth(RV.CANDIDATES, 320);
    reviewSheet.setColumnWidth(RV.PASSPORT_INPUT, 150);
    reviewSheet.setColumnWidth(RV.STATUS, 110);

    // تجميد الرأس
    reviewSheet.setFrozenRows(1);

    // RTL
    reviewSheet.setRightToLeft(true);
  }

  // ملاحظة تعليمية في أعلى الشيت
  var note = 'ملء عمود "الجواز الصحيح" ثم تشغيل syncReviewPassports() لتطبيق التغييرات';
  reviewSheet.getRange(1, 1).setNote(note);

  SpreadsheetApp.flush();
  Logger.log('[Review] تم إنشاء شيت المراجعة: ' + dataRows.length + ' صف');

  return { count: dataRows.length, sheetName: REVIEW_SHEET_NAME };
}

// ==================== تصنيف السبب ====================

function classifyReason_(notes, allPassengers) {
  var n = String(notes || '').toLowerCase();

  if (n.indexOf('collision') !== -1) {
    return { label: '⚠️ تعارض أسماء', isCollision: true };
  }
  if (n.indexOf('pdf_text_empty') !== -1) {
    return { label: '📄 PDF فارغ', isCollision: false };
  }
  if (n.indexOf('claude_no_name') !== -1) {
    var passengerCount = 0;
    if (allPassengers) {
      try { passengerCount = JSON.parse(String(allPassengers)).length; } catch(e) {}
    }
    return { label: passengerCount > 1 ? '👥 تذكرة جماعية (' + passengerCount + ')' : '🔤 لم يُستخرج اسم', isCollision: false };
  }
  if (n.indexOf('claude_error') !== -1) {
    return { label: '🔴 خطأ Claude API', isCollision: false };
  }
  if (n.indexOf('not found in pd') !== -1 || n.indexOf('go_to_stage_5') !== -1) {
    return { label: '🔍 اسم غير موجود في PD', isCollision: false };
  }
  if (n.indexOf('no_match') !== -1 || n.indexOf('not matched') !== -1) {
    return { label: '❓ لا تطابق', isCollision: false };
  }
  return { label: '❓ غير معروف', isCollision: false };
}

// ==================== جلب المرشحين للـ Collision ====================

/**
 * يُعيد نص قابل للقراءة بقائمة المرشحين المحتملين
 */
function getCollisionCandidates_(claudePnr, nameToMatch, allPassengersJson, flightMap, pdCache) {
  var results = [];

  // Strategy A: عبر PNR
  if (claudePnr) {
    var pnr = String(claudePnr).trim().toUpperCase();
    var contractName = flightMap[pnr];
    if (contractName && nameToMatch) {
      var pool = pdCache.filter(function(p) { return p.contractName === contractName; });
      var match = matchNameInCachePool_(nameToMatch, pool);
      if (match && match.collision && match.matches) {
        match.matches.forEach(function(m) {
          results.push(m.firstName + ' ' + m.lastName + ' | جواز: ' + m.passport);
        });
      }
    }
  }

  // Strategy B: عبر AllPassengers
  if (results.length === 0 && allPassengersJson) {
    var passengers = [];
    try {
      var parsed = JSON.parse(String(allPassengersJson));
      if (Array.isArray(parsed)) {
        parsed.forEach(function(p) {
          var name = (typeof p === 'string') ? p : (p.name || p.fullName || '');
          if (name) passengers.push(String(name).trim());
        });
      }
    } catch(e) {}

    var seen = {};
    for (var i = 0; i < passengers.length; i++) {
      var match2 = searchNameInPdCache_(passengers[i], pdCache);
      if (match2 && match2.collision && match2.matches) {
        match2.matches.forEach(function(m) {
          var key = m.passport;
          if (!seen[key]) {
            seen[key] = true;
            results.push(m.firstName + ' ' + m.lastName + ' | جواز: ' + m.passport);
          }
        });
      } else if (match2 && match2.passport && !seen[match2.passport]) {
        seen[match2.passport] = true;
        results.push(match2.firstName + ' ' + match2.lastName + ' | جواز: ' + match2.passport);
      }
    }
  }

  if (results.length === 0) return '(يرجى البحث يدوياً في PD)';
  return results.join('\n');
}

// ==================== مزامنة الجوازات المُدخَلة ====================

/**
 * يقرأ عمود G (الجواز الصحيح) ويكتب في PD مباشرة
 * شغّل هذه الدالة بعد ملء الجوازات
 */
function syncReviewPassports() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var reviewSheet = ss.getSheetByName(REVIEW_SHEET_NAME);
  if (!reviewSheet) throw new Error('شيت المراجعة غير موجود — شغّل buildReviewSheet() أولاً');

  var pdCache = loadPdCache_(ss);
  var passportIndex = buildPassportIndex_(pdCache);
  var pdSheet = ss.getSheetByName(TL.Config.SHEET_PD);

  // رقم عمود TICKET_URL في PD
  var ticketUrlCol = (function() {
    var headers = pdSheet.getRange(1, 1, 1, pdSheet.getLastColumn()).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === TL.Config.PD_HEADERS.TICKET_URL) return i + 1;
    }
    return TL.Config.PD_COL.TICKET_URL;
  })();

  var unresolvedSheet = ss.getSheetByName('TL_UnresolvedTickets');
  var unresolvedData = unresolvedSheet.getDataRange().getValues();

  var lastRow = reviewSheet.getLastRow();
  if (lastRow < 2) { Logger.log('[Sync] لا بيانات في شيت المراجعة'); return { synced: 0 }; }

  var data = reviewSheet.getRange(2, 1, lastRow - 1, RV.UNRES_ROW).getValues();

  var stats = { synced: 0, alreadyDone: 0, error: 0 };

  for (var i = 0; i < data.length; i++) {
    var passport     = String(data[i][RV.PASSPORT_INPUT - 1] || '').trim();
    var currentStatus = String(data[i][RV.STATUS - 1] || '').trim();
    var unresRow     = parseInt(data[i][RV.UNRES_ROW - 1], 10);

    if (!passport) continue;
    if (currentStatus === 'تم الربط ✅') { stats.alreadyDone++; continue; }

    // اجلب fileId من TL_UnresolvedTickets
    var originalFileId = String(unresolvedData[unresRow - 1][UNRES_COL.ORIGINAL_FILE_ID] || '').trim();
    var fileName       = String(unresolvedData[unresRow - 1][UNRES_COL.ORIGINAL_FILE_NAME] || '').trim();
    var existingNotes  = String(unresolvedData[unresRow - 1][UNRES_COL.NOTES] || '').trim();

    if (!originalFileId) {
      reviewSheet.getRange(i + 2, RV.STATUS).setValue('❌ fileId مفقود');
      stats.error++;
      continue;
    }

    // اكتب في PD
    var writeResult = writeToPdCached_(passport, originalFileId, passportIndex, pdCache, pdSheet, ticketUrlCol);

    if (writeResult.status === 'WRITTEN' || writeResult.status === 'REPLACED' || writeResult.status === 'ALREADY_WRITTEN') {
      // سجّل في RunLog
      logWriteOperation_({
        passport: passport,
        fileId: originalFileId,
        fileName: fileName,
        stage: 'Review-Manual',
        status: writeResult.status,
        oldZ: writeResult.oldZ || '',
        newZ: writeResult.newZ || '',
        rowIndex: writeResult.rowIndex || '',
        notes: 'Manual review input'
      });

      // حدّث TL_UnresolvedTickets
      var newNotes = existingNotes + ' | Review: manual passport=' + passport;
      unresolvedSheet.getRange(unresRow, 14, 1, 3).setValues([['LINKED', new Date(), newNotes]]);

      // حدّث شيت المراجعة
      reviewSheet.getRange(i + 2, RV.STATUS).setValue('تم الربط ✅');
      stats.synced++;

    } else if (writeResult.status === 'PASSPORT_NOT_FOUND_IN_PD') {
      reviewSheet.getRange(i + 2, RV.STATUS).setValue('❌ جواز غير موجود في PD');
      stats.error++;
    } else {
      reviewSheet.getRange(i + 2, RV.STATUS).setValue('❌ ' + writeResult.status);
      stats.error++;
    }
  }

  SpreadsheetApp.flush();
  Logger.log('[Sync] synced=' + stats.synced + ' alreadyDone=' + stats.alreadyDone + ' error=' + stats.error);
  return stats;
}
