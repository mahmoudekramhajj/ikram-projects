/**
 * Matcher.js — السلّم الكامل لمطابقة PDF بحاج
 *
 * المراحل المتفق عليها:
 *   المرحلة 1: شيت الطيران (PNR في B → اسم العقد في CM) → PD (عمود V) → مطابقة الاسم في pool
 *   المرحلة 3: PD كاملاً بالاسم (4 تركيبات، exact فقط)
 *   المرحلة 4: Claude API (يُضاف لاحقاً)
 *   المرحلة 5: Unresolved (يُضاف لاحقاً)
 *
 * ملاحظة: لا fuzzy/Jaro-Winkler — exact فقط لتجنب false positives.
 */

// ==================== Helpers: Name Normalization ====================

/**
 * يستخرج اسم الحاج من اسم الملف
 *   "MOHAMED MUNEEB RAVAT.pdf" → "MOHAMED MUNEEB RAVAT"
 *   "TASNEEM KARA(1).pdf" → "TASNEEM KARA"
 *   "BEGUMRIMA MRS-7ERY6E.pdf" → "BEGUMRIMA MRS"
 */
function extractNameFromFileName_(fileName) {
  if (!fileName) return null;
  var noExt = fileName.replace(/\.pdf$/i, '');
  noExt = noExt.replace(/\(\d+\)$/, '').trim();
  // إن وُجد PNR بصيغة "-XXXXXX" في نهاية الاسم، احذفه
  noExt = noExt.replace(/-[A-Z0-9]{6}$/, '').trim();
  return noExt || null;
}

/**
 * تطبيع اسم: uppercase، بدون مسافات، بدون diacritics، بدون ألقاب
 */
function normalizeName_(name) {
  if (!name) return '';
  return String(name)
    .toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\b(MR|MRS|MS|MSTR|MISS|CHD|INF|MASTER)\b/g, '')
    .replace(/[^A-Z]/g, '')
    .trim();
}

/**
 * يقسّم اسم الملف إلى firstName + lastName حسب آخر كلمة
 *   "MOHAMED MUNEEB RAVAT" → {firstName: "MOHAMED MUNEEB", lastName: "RAVAT"}
 *   "TASNEEM KARA" → {firstName: "TASNEEM", lastName: "KARA"}
 *   "AHMED" → {firstName: "AHMED", lastName: ""}
 */
function splitNameByLastWord_(fullName) {
  if (!fullName) return { firstName: '', lastName: '' };
  var cleaned = String(fullName)
    .replace(/\b(MR|MRS|MS|MSTR|MISS|CHD|INF|MASTER)\b/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  var parts = cleaned.split(' ').filter(function(p) { return p.length > 0; });
  if (parts.length === 0) return { firstName: '', lastName: '' };
  if (parts.length === 1) return { firstName: parts[0], lastName: '' };
  return {
    firstName: parts.slice(0, -1).join(' '),
    lastName: parts[parts.length - 1]
  };
}

/**
 * يبني تركيبات tokenKey للمقارنة الدقيقة
 * يدعم تجاوز middle names (PDF: "MOHAMED MUNEEB RAVAT" vs PD: "MOHAMED RAVAT")
 *
 * التركيبات المُنتَجة:
 *   1. firstName كاملاً + lastName  (الأساسية)
 *   2. lastName + firstName كاملاً  (مقلوبة)
 *   3. أول كلمة من firstName + lastName  (تجاوز middle names في PDF)
 *   4. lastName + أول كلمة من firstName
 *   5. آخر كلمة من firstName + lastName  (لو middle name هو الـ "first")
 *   6. lastName + آخر كلمة من firstName
 *   7. mononym fallback لو واحد فقط
 */
function buildAllTokenKeys_(firstName, lastName) {
  var keys = [];
  var fNorm = normalizeName_(firstName);
  var lNorm = normalizeName_(lastName);

  if (fNorm && lNorm) {
    // 1+2: التركيبة الكاملة
    keys.push(fNorm + lNorm);
    keys.push(lNorm + fNorm);

    // 3-6: تجاوز middle names إن وُجدت كلمات متعددة في firstName
    var firstWords = String(firstName)
      .replace(/\b(MR|MRS|MS|MSTR|MISS|CHD|INF|MASTER)\b/gi, '')
      .replace(/\s+/g, ' ').trim().split(' ').filter(function(w) { return w.length > 0; });

    if (firstWords.length > 1) {
      var firstFirst = normalizeName_(firstWords[0]);
      var firstLast = normalizeName_(firstWords[firstWords.length - 1]);

      if (firstFirst && firstFirst !== fNorm) {
        keys.push(firstFirst + lNorm);
        keys.push(lNorm + firstFirst);
      }
      if (firstLast && firstLast !== fNorm && firstLast !== firstFirst) {
        keys.push(firstLast + lNorm);
        keys.push(lNorm + firstLast);
      }
    }
  } else if (fNorm) {
    keys.push(fNorm);
  } else if (lNorm) {
    keys.push(lNorm);
  }

  // أزل المكررات
  var uniq = [];
  keys.forEach(function(k) { if (uniq.indexOf(k) === -1) uniq.push(k); });
  return uniq;
}

// ==================== المرحلة 1: شيت الطيران ====================

/**
 * يبحث عن PNR في شيت الطيران، عمود B
 * عمود B قد يحوي عدة PNRs مفصولة بـ " - "
 * @return {Object|null} {row, contractName} أو null
 */
function findContractByPnrInFlights_(pnr) {
  if (!pnr) return null;
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_FLIGHTS);
  if (!sheet) throw new Error('Sheet not found: ' + TL.Config.SHEET_FLIGHTS);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  // عمود B (PNR) و عمود CM (اسم العقد) — نقرأ فقط هذين العمودين للأداء
  var pnrColData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();  // B
  var cmColData = sheet.getRange(2, 91, lastRow - 1, 1).getValues();  // CM = 91 (1-based)
  // ملاحظة: CM = العمود رقم 91. سنتحقق من هذا في diagnose.

  var pnrUpper = String(pnr).trim().toUpperCase();

  for (var i = 0; i < pnrColData.length; i++) {
    var cellPnr = String(pnrColData[i][0] || '').trim();
    if (!cellPnr) continue;

    // قسّم الخلية على " - " وابحث عن تطابق
    var pnrs = cellPnr.split(/\s*-\s*/).map(function(p) { return p.trim().toUpperCase(); });
    if (pnrs.indexOf(pnrUpper) !== -1) {
      return {
        row: i + 2,
        cellPnrRaw: cellPnr,
        contractName: String(cmColData[i][0] || '').trim()
      };
    }
  }
  return null;
}

/**
 * يجلب pool حجاج من PD حيث عمود V (اسم العقد) يطابق
 * @return {Array<Object>} [{rowIndex, passport, firstName, lastName, contractName}]
 */
function getPilgrimsPoolByContract_(contractName) {
  if (!contractName) return [];
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (!sheet) throw new Error('Sheet not found: ' + TL.Config.SHEET_PD);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // أعمدة PD: F=6 passport, K=11 first, L=12 last, V=22 contractName
  var data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
  var pool = [];
  var target = String(contractName).trim();

  for (var i = 0; i < data.length; i++) {
    var rowContract = String(data[i][21] || '').trim();  // V (index 21)
    if (rowContract === target) {
      pool.push({
        rowIndex: i + 2,
        passport: String(data[i][5] || '').trim(),     // F
        firstName: String(data[i][10] || '').trim(),    // K
        lastName: String(data[i][11] || '').trim(),     // L
        contractName: rowContract
      });
    }
  }
  return pool;
}

/**
 * يطابق اسم PDF مع pool عبر exact tokenKey (كل التركيبات)
 * @return {Object|null} {passport, firstName, lastName, matchedKey} أو null
 */
function matchNameInPool_(pdfName, pool) {
  if (!pdfName || !pool || pool.length === 0) return null;

  var split = splitNameByLastWord_(pdfName);
  var pdfKeys = buildAllTokenKeys_(split.firstName, split.lastName);
  // أيضاً ضع الاسم كاملاً ملصوقاً (لو الاسم في PDF مكتوب كله بدون مسافات)
  pdfKeys.push(normalizeName_(pdfName));

  // أزل المكررات
  var uniqKeys = [];
  pdfKeys.forEach(function(k) {
    if (k && uniqKeys.indexOf(k) === -1) uniqKeys.push(k);
  });

  var matches = [];
  for (var i = 0; i < pool.length; i++) {
    var p = pool[i];
    var pdKeys = buildAllTokenKeys_(p.firstName, p.lastName);

    // exact match: أي من مفاتيح PDF يطابق أي من مفاتيح PD
    for (var j = 0; j < uniqKeys.length; j++) {
      if (pdKeys.indexOf(uniqKeys[j]) !== -1) {
        matches.push({
          passport: p.passport,
          firstName: p.firstName,
          lastName: p.lastName,
          matchedKey: uniqKeys[j]
        });
        break;  // لا تكرّر نفس الحاج
      }
    }
  }

  if (matches.length === 1) return matches[0];
  if (matches.length > 1) return { collision: true, matches: matches };
  return null;
}

// ==================== المرحلة 3: PD كاملاً بالاسم ====================

/**
 * يبحث في PD كاملاً (6793 صف) عن اسم بمطابقة exact
 * @return {Object|null} {passport, firstName, lastName, matchedKey} أو null أو {collision}
 */
function searchNameInFullPd_(pdfName) {
  if (!pdfName) return null;
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (!sheet) throw new Error('Sheet not found: ' + TL.Config.SHEET_PD);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var split = splitNameByLastWord_(pdfName);
  var pdfKeys = buildAllTokenKeys_(split.firstName, split.lastName);
  pdfKeys.push(normalizeName_(pdfName));
  var uniqKeys = [];
  pdfKeys.forEach(function(k) {
    if (k && uniqKeys.indexOf(k) === -1) uniqKeys.push(k);
  });

  // قراءة F, K, L فقط (للأداء)
  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
  var matches = [];

  for (var i = 0; i < data.length; i++) {
    var passport = String(data[i][5] || '').trim();    // F
    var firstName = String(data[i][10] || '').trim();   // K
    var lastName = String(data[i][11] || '').trim();    // L
    if (!firstName && !lastName) continue;

    var pdKeys = buildAllTokenKeys_(firstName, lastName);
    for (var j = 0; j < uniqKeys.length; j++) {
      if (pdKeys.indexOf(uniqKeys[j]) !== -1) {
        matches.push({
          rowIndex: i + 2,
          passport: passport,
          firstName: firstName,
          lastName: lastName,
          matchedKey: uniqKeys[j]
        });
        break;
      }
    }
  }

  if (matches.length === 1) return matches[0];
  if (matches.length > 1) return { collision: true, matches: matches };
  return null;
}

// ==================== الدالة الرئيسية: السلّم الكامل ====================

/**
 * يطابق PDF بحاج عبر السلّم الكامل
 * @param {Object} pdfInfo {fileId, fileName, folderName, pnr}
 * @return {Object} {status, stage, passport, ...}
 */
function matchPilgrimForPdf_(pdfInfo) {
  var result = {
    fileId: pdfInfo.fileId,
    fileName: pdfInfo.fileName,
    folderName: pdfInfo.folderName,
    folderPnr: pdfInfo.pnr,
    extractedName: null,
    stage: null,                  // 1, 3, 4, 5
    status: null,                  // 'OK' | 'COLLISION' | 'GO_TO_STAGE_3' | 'GO_TO_STAGE_4' | 'UNRESOLVED'
    passport: null,
    matchedFirstName: null,
    matchedLastName: null,
    matchedKey: null,
    contractName: null,
    notes: []
  };

  result.extractedName = extractNameFromFileName_(pdfInfo.fileName);
  if (!result.extractedName) {
    result.status = 'NO_NAME_IN_FILENAME';
    result.notes.push('Filename has no extractable name; need stage 4 (Claude on PDF text)');
    return result;
  }

  // ===== المرحلة 1 =====
  if (pdfInfo.pnr) {
    var contractInfo = findContractByPnrInFlights_(pdfInfo.pnr);
    if (contractInfo && contractInfo.contractName) {
      result.contractName = contractInfo.contractName;
      result.notes.push('Stage 1: Found PNR in flights, contract = ' + contractInfo.contractName);

      var pool = getPilgrimsPoolByContract_(contractInfo.contractName);
      result.notes.push('Stage 1: PD pool size = ' + pool.length);

      if (pool.length > 0) {
        var match = matchNameInPool_(result.extractedName, pool);
        if (match && match.collision) {
          result.stage = 1;
          result.status = 'COLLISION';
          result.collision = match.matches;
          return result;
        }
        if (match) {
          result.stage = 1;
          result.status = 'OK';
          result.passport = match.passport;
          result.matchedFirstName = match.firstName;
          result.matchedLastName = match.lastName;
          result.matchedKey = match.matchedKey;
          return result;
        }
        result.notes.push('Stage 1: Name not found in pool, falling to stage 3');
      } else {
        result.notes.push('Stage 1: Empty pool for contract, falling to stage 3');
      }
    } else {
      result.notes.push('Stage 1: PNR not found in flights sheet, falling to stage 3');
    }
  } else {
    result.notes.push('Stage 1: No PNR resolved from folder, falling to stage 3');
  }

  // ===== المرحلة 3 =====
  var pdMatch = searchNameInFullPd_(result.extractedName);
  if (pdMatch && pdMatch.collision) {
    result.stage = 3;
    result.status = 'COLLISION';
    result.collision = pdMatch.matches;
    return result;
  }
  if (pdMatch) {
    result.stage = 3;
    result.status = 'OK';
    result.passport = pdMatch.passport;
    result.matchedFirstName = pdMatch.firstName;
    result.matchedLastName = pdMatch.lastName;
    result.matchedKey = pdMatch.matchedKey;
    return result;
  }

  // ===== فشلت 1 و 3 → سيُحال للمرحلة 4 (Claude) ثم 5 (Unresolved) لاحقاً =====
  result.stage = 3;
  result.status = 'GO_TO_STAGE_4';
  result.notes.push('Stage 3: Name not found in any combination across full PD');
  return result;
}

// ==================== اختبارات ====================

function testFlightSearch(pnr) {
  return findContractByPnrInFlights_(pnr);
}

function testPoolByContract(contractName) {
  var pool = getPilgrimsPoolByContract_(contractName);
  return {
    contractName: contractName,
    poolSize: pool.length,
    sample: pool.slice(0, 5)
  };
}

function testMatchSinglePdf(fileName, folderName) {
  var pnrInfo = resolvePnr_({
    fileId: '',
    fileName: fileName,
    folderName: folderName
  });
  return matchPilgrimForPdf_({
    fileId: '',
    fileName: fileName,
    folderName: folderName,
    pnr: pnrInfo.pnr
  });
}

/**
 * اختبار شامل على عقد:
 * - يجلب كل PDFs في مجلد
 * - يطبّق السلّم الكامل
 * - يُرجع تقرير مفصّل
 */
function testFullLadderForFolder(folderName) {
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
      extractedName: match.extractedName,
      stage: match.stage,
      status: match.status,
      passport: match.passport,
      matched: match.matchedFirstName ? (match.matchedFirstName + ' ' + match.matchedLastName) : null,
      matchedKey: match.matchedKey
    });
  }

  // إحصائيات
  var stats = {
    totalPdfs: results.length,
    stage1_OK: results.filter(function(r) { return r.stage === 1 && r.status === 'OK'; }).length,
    stage3_OK: results.filter(function(r) { return r.stage === 3 && r.status === 'OK'; }).length,
    needStage4: results.filter(function(r) { return r.status === 'GO_TO_STAGE_4'; }).length,
    collisions: results.filter(function(r) { return r.status === 'COLLISION'; }).length,
    other: results.filter(function(r) {
      return r.status !== 'OK' && r.status !== 'GO_TO_STAGE_4' && r.status !== 'COLLISION';
    }).length
  };

  return {
    folderName: targetFolder.getName(),
    pnr: pnr,
    stats: stats,
    results: results
  };
}
