/**
 * Diagnose.js — أدوات تشخيص
 */

function diagnoseB2C() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_B2C);
  if (!sheet) return { error: 'Sheet not found' };
  return {
    sheet: TL.Config.SHEET_B2C,
    lastRow: sheet.getLastRow(),
    lastCol: sheet.getLastColumn()
  };
}

/**
 * يعرض آخر N صف من TL_RunLog
 */
function showRecentRunLog(n) {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_RunLog');
  if (!sheet) return { error: 'TL_RunLog not found' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0, rows: [] };
  var count = n || 20;
  var startRow = Math.max(2, lastRow - count + 1);
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
  return {
    count: data.length,
    headers: sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0],
    rows: data
  };
}

/**
 * يعرض إحصائيات RunLog (groupby status)
 */
function runLogStats() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_RunLog');
  if (!sheet) return { error: 'TL_RunLog not found' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  // أعمدة: Timestamp(0), Passport(1), FileId(2), FileName(3), Stage(4), Status(5), OldZ(6), NewZ(7), RowIndex(8), Notes(9)
  var byStatus = {};
  var byStage = {};
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][5] || 'EMPTY');
    var stage = String(data[i][4] || 'EMPTY');
    byStatus[status] = (byStatus[status] || 0) + 1;
    byStage[stage] = (byStage[stage] || 0) + 1;
  }
  return {
    total: data.length,
    byStatus: byStatus,
    byStage: byStage
  };
}

/**
 * تحليل مفصّل لـ TL_UnresolvedTickets
 */
function analyzeUnresolved() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!sheet) return { error: 'No tab' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };
  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

  var byNotes = {};      // groupby reason
  var hasClaudeData = 0; // كم صف فيه Claude data
  var noClaudeData = 0;
  var sampleNoB2C = [];
  var sampleCollision = [];
  var sampleNoMatch = [];

  for (var i = 0; i < data.length; i++) {
    var notes = String(data[i][15] || '').substring(0, 100);
    var claudeName = String(data[i][6] || '');
    var folder = String(data[i][4] || '');
    var pdfName = String(data[i][3] || '');

    // تصنيف بسيط
    var category = 'OTHER';
    if (notes.indexOf('NOT IN B2C') !== -1) category = 'NO_B2C_ROW';
    else if (notes.indexOf('COLLISION') !== -1) category = 'COLLISION';
    else if (notes.indexOf('Stage 3:') !== -1 || notes.indexOf('Claude') !== -1) category = 'NO_PD_MATCH';

    byNotes[category] = (byNotes[category] || 0) + 1;

    if (claudeName) hasClaudeData++; else noClaudeData++;

    // عينات (5 من كل نوع)
    if (category === 'NO_B2C_ROW' && sampleNoB2C.length < 5) {
      sampleNoB2C.push({ folder: folder, pdf: pdfName, notes: notes });
    } else if (category === 'COLLISION' && sampleCollision.length < 5) {
      sampleCollision.push({ folder: folder, pdf: pdfName, notes: notes });
    } else if (category === 'NO_PD_MATCH' && sampleNoMatch.length < 5) {
      sampleNoMatch.push({ folder: folder, pdf: pdfName, notes: notes, claudeName: claudeName });
    }
  }

  return {
    total: data.length,
    byCategory: byNotes,
    claudeRan: hasClaudeData,
    claudeNotRun: noClaudeData,
    samples: {
      noB2cRow: sampleNoB2C,
      collision: sampleCollision,
      noPdMatch: sampleNoMatch
    }
  };
}

/**
 * عرض بيانات PD كاملة لجواز محدد
 */
function getPdRowByPassport(passport) {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (!sheet) return { error: 'PD sheet not found' };
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][5] || '').trim() === String(passport).trim()) {
      var result = {};
      for (var c = 0; c < lastCol; c++) {
        var h = String(headers[c] || '').trim();
        var v = String(data[i][c] || '').trim();
        if (h && v) result[h] = v;
      }
      result['_sheetRow'] = i + 2;
      return result;
    }
  }
  return { error: 'passport not found: ' + passport };
}

/**
 * اختبار Strategy E مباشرة على اسم PATEL/MUHAMMAD MR
 * يُظهر ما سيرسله لـ Claude وما سيردّه
 */
function testStrategyEOnPatel() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var pdCache = loadPdCache_(ss);

  // بيانات رحلة PATEL من الصف المعروف
  var claudeFlightsJson = '[{"flightNo":"SV 124","depDate":"2026-05-14","depTime":"14:45","from":"Manchester (MAN)","to":"Jeddah (JED)","arrDate":"2026-05-14","arrTime":"23:15"},{"flightNo":"SV 1420","depDate":"2026-05-15","depTime":"01:30","from":"Jeddah (JED)","to":"Madinah (MED)","arrDate":"2026-05-15","arrTime":"02:30"}]';
  var claudeFullName = 'PATEL/MUHAMMAD MR';

  var result = tryStrategyE_(claudeFullName, claudeFlightsJson, pdCache);
  return result;
}

/**
 * تشخيص صفوف PATEL المكررة في TL_UnresolvedTickets
 * يقارن كل صف يحمل PATELMUHAMMAD ويُظهر الفروق بين الصفوف
 */
function diagnosePatels() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!sheet) return { error: 'TL_UnresolvedTickets not found' };

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

  var results = [];
  for (var i = 0; i < data.length; i++) {
    var fname = String(data[i][3] || '');
    if (fname.toUpperCase().indexOf('PATELMUHAMMAD') !== -1 || fname.toUpperCase().indexOf('PATEL') !== -1) {
      results.push({
        sheetRow: i + 2,
        timestamp: data[i][0],
        originalFileId: data[i][1],
        copiedFileId: data[i][2],
        fileName: fname,
        folderName: data[i][4],
        extractedName: data[i][5],
        claudeFullName: data[i][6],
        claudePnr: data[i][7],
        claudeFirstName: data[i][8],
        allPassengers: data[i][9],
        stage: data[i][10],
        substage: data[i][11],
        matchedBy: data[i][12],
        status: data[i][13],
        resolvedAt: data[i][14],
        notes: data[i][15]
      });
    }
  }

  // تحليل: هل الملفين لهم نفس originalFileId؟
  var analysis = {};
  if (results.length >= 2) {
    analysis.sameOriginalFileId = results[0].originalFileId === results[1].originalFileId;
    analysis.sameCopiedFileId = results[0].copiedFileId === results[1].copiedFileId;
    analysis.sameTimestamp = String(results[0].timestamp) === String(results[1].timestamp);
    analysis.sameClaudePnr = results[0].claudePnr === results[1].claudePnr;
    analysis.sameAllPassengers = results[0].allPassengers === results[1].allPassengers;
    analysis.differentFields = [];
    var fields = ['originalFileId','copiedFileId','timestamp','extractedName','claudeFullName','claudePnr','allPassengers','stage','notes'];
    for (var f = 0; f < fields.length; f++) {
      var key = fields[f];
      if (String(results[0][key]) !== String(results[1][key])) {
        analysis.differentFields.push({
          field: key,
          row1: String(results[0][key]).substring(0, 100),
          row2: String(results[1][key]).substring(0, 100)
        });
      }
    }
  }

  return { count: results.length, rows: results, analysis: analysis };
}

/**
 * تشخيص شامل للصفوف PENDING_MANUAL المتبقية
 * سريع: لا يستدعي extractPdfText_ (بطيء) — يعتمد على Notes + فحص وصول الملف
 *
 * الفئات:
 *   FILE_NOT_ACCESSIBLE  — DriveApp لا يستطيع الوصول لأي من ملفَي الصف
 *   PDF_TEXT_EMPTY       — ذُكر في Notes أن النص فارغ سابقاً
 *   COLLISION_UNRESOLVED — تعارض أسماء لم يُحَل
 *   NAME_NOT_IN_PD       — Claude استخرج الاسم لكنه غير موجود في PD
 *   CLAUDE_ERROR         — Claude API فشل
 *   NO_CLAUDE_DATA       — Claude لم يُشغَّل (عمود G فارغ + لا سبب محدد في Notes)
 *   OTHER                — غير مصنَّف
 */
function diagnoseRemainingUnresolved() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!sheet) return { error: 'TL_UnresolvedTickets not found' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, byCategory: {} };

  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();

  var byCategory = {
    FILE_NOT_ACCESSIBLE: 0,
    PDF_TEXT_EMPTY: 0,
    COLLISION_UNRESOLVED: 0,
    NAME_NOT_IN_PD: 0,
    CLAUDE_ERROR: 0,
    NO_CLAUDE_DATA: 0,
    OTHER: 0
  };

  var samples = {
    FILE_NOT_ACCESSIBLE: [],
    PDF_TEXT_EMPTY: [],
    COLLISION_UNRESOLVED: [],
    NAME_NOT_IN_PD: [],
    CLAUDE_ERROR: [],
    NO_CLAUDE_DATA: [],
    OTHER: []
  };

  var MAX_SAMPLES = 3;
  var total = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var status = String(row[13] || '').trim();  // N
    if (status !== 'PENDING_MANUAL') continue;
    total++;

    var originalFileId = String(row[1] || '').trim();   // B
    var copiedFileId   = String(row[2] || '').trim();   // C
    var pdfName        = String(row[3] || '').trim();   // D
    var folderName     = String(row[4] || '').trim();   // E
    var claudeName     = String(row[6] || '').trim();   // G
    var notes          = String(row[15] || '').trim();  // P
    var notesLower     = notes.toLowerCase();

    var category = 'OTHER';

    // 1. فحص وصول الملف (copiedFileId أولاً، ثم originalFileId)
    var fileId = copiedFileId || originalFileId;
    var fileAccessible = false;
    if (fileId) {
      try {
        DriveApp.getFileById(fileId);
        fileAccessible = true;
      } catch (e) {
        // جرّب الآخر إن وُجد
        if (copiedFileId && originalFileId && fileId === copiedFileId) {
          try {
            DriveApp.getFileById(originalFileId);
            fileAccessible = true;
          } catch (e2) {
            fileAccessible = false;
          }
        }
      }
    }

    if (!fileAccessible) {
      category = 'FILE_NOT_ACCESSIBLE';
    } else if (notesLower.indexOf('pdf_text_empty') !== -1) {
      category = 'PDF_TEXT_EMPTY';
    } else if (notesLower.indexOf('collision') !== -1) {
      category = 'COLLISION_UNRESOLVED';
    } else if (notesLower.indexOf('not found in pd') !== -1 ||
               notesLower.indexOf('go_to_stage_5') !== -1 ||
               notesLower.indexOf('claude extracted name but not found') !== -1) {
      category = 'NAME_NOT_IN_PD';
    } else if (notesLower.indexOf('claude_error') !== -1 ||
               notesLower.indexOf('claude error') !== -1 ||
               notesLower.indexOf('claude api error') !== -1) {
      category = 'CLAUDE_ERROR';
    } else if (!claudeName) {
      category = 'NO_CLAUDE_DATA';
    } else {
      category = 'OTHER';
    }

    byCategory[category]++;

    if (samples[category].length < MAX_SAMPLES) {
      samples[category].push({
        sheetRow: i + 2,
        pdfName: pdfName,
        folder: folderName,
        claudeName: claudeName,
        notes: notes.substring(0, 150)
      });
    }
  }

  return {
    total: total,
    byCategory: byCategory,
    samples: samples
  };
}

/**
 * يُعيد قائمة الملفات غير الموصولة (FILE_NOT_ACCESSIBLE) للمعالجة اليدوية
 */
function listInaccessibleFiles() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!sheet) return { error: 'TL_UnresolvedTickets not found' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0, files: [] };

  var data = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
  var files = [];

  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][13] || '').trim();
    if (status !== 'PENDING_MANUAL') continue;

    var originalFileId = String(data[i][1] || '').trim();
    var copiedFileId   = String(data[i][2] || '').trim();
    var pdfName        = String(data[i][3] || '').trim();
    var folderName     = String(data[i][4] || '').trim();

    var originalOk = false, copiedOk = false;

    if (originalFileId) {
      try { DriveApp.getFileById(originalFileId); originalOk = true; } catch(e) {}
    }
    if (copiedFileId) {
      try { DriveApp.getFileById(copiedFileId); copiedOk = true; } catch(e) {}
    }

    if (!originalOk && !copiedOk) {
      files.push({
        sheetRow: i + 2,
        pdfName: pdfName,
        folder: folderName,
        originalFileId: originalFileId,
        copiedFileId: copiedFileId
      });
    }
  }

  return { count: files.length, files: files };
}

/**
 * يعرض إحصائيات Unresolved tab
 */
function unresolvedStats() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!sheet) return { error: 'TL_UnresolvedTickets not found' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };
  var data = sheet.getRange(2, 14, lastRow - 1, 1).getValues();  // Status column N
  var byStatus = {};
  for (var i = 0; i < data.length; i++) {
    var s = String(data[i][0] || 'EMPTY');
    byStatus[s] = (byStatus[s] || 0) + 1;
  }
  return { total: data.length, byStatus: byStatus };
}

/**
 * يفحص آخر N صف من TL_UnresolvedTickets:
 *   - يبيّن مصدر المجلد (A_inbox / B_inbox / غيره)
 *   - يبحث عن claudeName في PD ليُحدّد هل الحاج موجود
 *   - groupby بالمجلد للسبب الجذري
 */
function diagnoseLatestUnresolved(n) {
  var count = n || 30;
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_UnresolvedTickets');
  if (!sheet) return { error: 'TL_UnresolvedTickets not found' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };

  var startRow = Math.max(2, lastRow - count + 1);
  var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 16).getValues();

  // PD index بـ first/last name
  var pd = ss.getSheetByName(TL.Config.SHEET_PD);
  var pdLast = pd.getLastRow();
  var pdNames = {};  // upper(firstName + ' ' + lastName) → row
  if (pdLast >= 2) {
    var pdData = pd.getRange(2, 1, pdLast - 1, 13).getValues();
    for (var p = 0; p < pdData.length; p++) {
      var fn = String(pdData[p][10] || '').trim().toUpperCase();   // K
      var ln = String(pdData[p][11] || '').trim().toUpperCase();   // L
      if (!fn && !ln) continue;
      pdNames[fn + ' ' + ln] = p + 2;
      pdNames[ln + ' ' + fn] = p + 2;
      pdNames[fn] = p + 2;
      pdNames[ln] = p + 2;
    }
  }

  var byFolder = {};
  var inPd = 0;
  var notInPd = 0;
  var samples = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var folder = String(row[4] || '').trim();          // E
    var pdfName = String(row[3] || '').trim();         // D
    var claudeName = String(row[6] || '').trim();      // G
    var notes = String(row[15] || '').substring(0, 150);

    byFolder[folder] = (byFolder[folder] || 0) + 1;

    // ابحث في PD
    var pdMatch = null;
    if (claudeName) {
      var key = claudeName.toUpperCase().replace(/[\/,]/g, ' ').replace(/\s+/g, ' ').trim();
      var parts = key.split(' ').filter(function(s) { return s.length > 1; });
      if (pdNames[key]) pdMatch = pdNames[key];
      if (!pdMatch && parts.length >= 2) {
        if (pdNames[parts[0] + ' ' + parts[parts.length - 1]]) pdMatch = pdNames[parts[0] + ' ' + parts[parts.length - 1]];
      }
      if (!pdMatch && parts.length >= 1) {
        for (var k = 0; k < parts.length; k++) {
          if (pdNames[parts[k]]) { pdMatch = pdNames[parts[k]]; break; }
        }
      }
    }

    if (pdMatch) inPd++;
    else notInPd++;

    if (samples.length < 10) {
      samples.push({
        sheetRow: startRow + i,
        folder: folder,
        pdfName: pdfName,
        claudeName: claudeName,
        pdMatch: pdMatch ? ('PD row ' + pdMatch) : 'NOT_IN_PD',
        notes: notes
      });
    }
  }

  return {
    examined: data.length,
    fromRow: startRow,
    toRow: lastRow,
    byFolder: byFolder,
    inPd: inPd,
    notInPd: notInPd,
    samples: samples
  };
}

/**
 * يبني شيت TL_BInboxReview للمراجعة اليدوية لتذاكر B_inbox.
 *
 * يمشي على C/B_inbox/ recursive، يستخرج PNR من اسم كل ملف،
 * يبحث في شيت الطيران عن العقد، ويعدّ الحجاج في PD ضمن نفس العقد.
 *
 * أعمدة الشيت:
 *   A  Timestamp
 *   B  CopiedFileId
 *   C  FileName
 *   D  ParentFolder        (مسار داخل B_inbox)
 *   E  ExtractedPNR
 *   F  FlightFoundInSheet  (YES/NO)
 *   G  ContractName        (من شيت الطيران)
 *   H  PilgrimsInPDForContract (عدد)
 *   I  ExtractedNameFromFile   (محاولة استخراج الاسم)
 *   J  DriveLink
 *   K  ManualDecision      (يملؤها الموظف: LINK_TO_PASSPORT_xxxxx / IGNORE / SKIP)
 *   L  PassportToLinkTo    (لو الموظف اختار ربطها بحاج محدد)
 *   M  Notes               (للموظف)
 */
function buildBInboxReviewSheet() {
  var REVIEW_TAB = 'TL_BInboxReview';
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);

  // 1) اعثر على مجلد B_inbox داخل C
  var rootFolder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var bSubs = rootFolder.getFoldersByName('B_inbox');
  if (!bSubs.hasNext()) return { error: 'B_inbox subfolder not found in TICKETS_FOLDER' };
  var bFolder = bSubs.next();

  // 2) اجمع كل PDFs (recursive، نفس نمط SourceSync)
  var collected = [];
  var deadline = Date.now() + TL.Config.MAX_RUNTIME_MS;
  collectPdfsRecursive_(bFolder, '', 0, collected, deadline);

  // 3) فهرس الطيران: PNR → contract name
  var flightsSheet = ss.getSheetByName(TL.Config.SHEET_FLIGHTS);
  var pnrToContract = {};
  if (flightsSheet) {
    var fLast = flightsSheet.getLastRow();
    var fLastCol = flightsSheet.getLastColumn();
    if (fLast >= 2) {
      var fData = flightsSheet.getRange(2, 1, fLast - 1, Math.max(fLastCol, TL.Config.FLIGHTS_COL.CONTRACT_NAME)).getValues();
      for (var i = 0; i < fData.length; i++) {
        var pnrCell = String(fData[i][TL.Config.FLIGHTS_COL.PNR - 1] || '').trim();
        var contract = String(fData[i][TL.Config.FLIGHTS_COL.CONTRACT_NAME - 1] || '').trim();
        if (!pnrCell) continue;
        var pnrs = pnrCell.split(/\s*-\s*|\s*\/\s*|\s+/).filter(function(s) { return /^[A-Z0-9]{6}$/.test(s.toUpperCase()); });
        for (var j = 0; j < pnrs.length; j++) {
          var key = pnrs[j].toUpperCase();
          if (!pnrToContract[key]) pnrToContract[key] = contract;
        }
      }
    }
  }

  // 4) فهرس PD: contract → عدد الحجاج
  var pdSheet = ss.getSheetByName(TL.Config.SHEET_PD);
  var contractToCount = {};
  if (pdSheet) {
    var pdLast = pdSheet.getLastRow();
    if (pdLast >= 2) {
      var pdData = pdSheet.getRange(2, TL.Config.PD_COL.CONTRACT_NAME, pdLast - 1, 1).getValues();
      for (var p = 0; p < pdData.length; p++) {
        var c = String(pdData[p][0] || '').trim();
        if (!c) continue;
        contractToCount[c] = (contractToCount[c] || 0) + 1;
      }
    }
  }

  // 5) أنشئ/أعد بناء شيت المراجعة
  var sheet = ss.getSheetByName(REVIEW_TAB);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(REVIEW_TAB);
  }
  var headers = [
    'Timestamp', 'CopiedFileId', 'FileName', 'ParentFolder',
    'ExtractedPNR', 'FlightFoundInSheet', 'ContractName', 'PilgrimsInPDForContract',
    'ExtractedNameFromFile', 'DriveLink',
    'ManualDecision', 'PassportToLinkTo', 'Notes'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  // 6) جمّع الصفوف
  var rows = [];
  var stats = {
    totalPdfs: collected.length,
    withPnr: 0,
    withoutPnr: 0,
    pnrInFlights: 0,
    pnrNotInFlights: 0
  };
  var now = new Date();

  for (var k = 0; k < collected.length; k++) {
    var item = collected[k];
    var pnr = extractPnrFromFileName_(item.fileName) || '';
    var nameGuess = '';
    // محاولة استخراج اسم من اسم الملف بعد PNR
    // النمط: PNR_NAME.pdf أو PNR-NAME.pdf
    var afterPnr = item.fileName.replace(/^[A-Z0-9]{6}[._\s\-]+/, '').replace(/\.pdf$/i, '').trim();
    if (afterPnr && afterPnr !== item.fileName) {
      nameGuess = afterPnr.replace(/[_\-]+/g, ' ').trim();
    }
    var contract = pnr ? (pnrToContract[pnr] || '') : '';
    var pilgrimCount = contract ? (contractToCount[contract] || 0) : 0;
    var driveLink = 'https://drive.google.com/file/d/' + item.fileId + '/view';

    if (pnr) stats.withPnr++; else stats.withoutPnr++;
    if (contract) stats.pnrInFlights++;
    else if (pnr) stats.pnrNotInFlights++;

    rows.push([
      now,
      item.fileId,
      item.fileName,
      item.parentPath || item.parentName,
      pnr,
      contract ? 'YES' : (pnr ? 'NO' : ''),
      contract,
      pilgrimCount,
      nameGuess,
      driveLink,
      '', // ManualDecision (فارغ — للموظف)
      '', // PassportToLinkTo
      ''  // Notes
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // عرض القوائم بشكل مفيد
  sheet.autoResizeColumns(1, headers.length);

  return {
    sheetName: REVIEW_TAB,
    rowsWritten: rows.length,
    stats: stats
  };
}
