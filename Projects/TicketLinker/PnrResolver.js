/**
 * PnrResolver.js — M1: استخراج PNR من PDF
 *
 * استراتيجية متدرّجة (الأرخص أولاً):
 *   1. اسم الملف (filename) → regex ^([A-Z0-9]{6})
 *   2. اسم المجلد (folder) → نفس regex بعد تنظيف
 *   3. نص PDF (OCR via Drive.Files.copy) → regex بعد كلمات مفتاحية
 *   4. Claude fallback (يُضاف لاحقاً عند الحاجة)
 *
 * Self-check: PNR المستخرَج يجب أن يوجد في شيت "الطيران" (يُتحقَّق منه عند الـ matching).
 *             هنا نخرج فقط النص المرشَّح، التحقق في M2.
 */

/**
 * استخراج PNR من اسم الملف
 * @param {string} fileName
 * @return {string|null} PNR (6 chars) أو null
 */
function extractPnrFromFileName_(fileName) {
  if (!fileName) return null;
  var m = fileName.match(TL.Config.PNR_FILENAME_REGEX);
  return m ? m[1].toUpperCase() : null;
}

/**
 * استخراج PNRs من اسم المجلد (قد يحوي عدة PNRs مثل "(T28XB3), (VENHEK)")
 * @param {string} folderName
 * @return {Array<string>} مصفوفة PNRs محتملة
 */
function extractPnrsFromFolderName_(folderName) {
  if (!folderName) return [];

  // إزالة الأحرف العربية والأقواس والشُرَط
  var cleaned = folderName
    .replace(/[\u0600-\u06FF]/g, ' ')        // إزالة العربية
    .replace(/[(),\-_]/g, ' ')                // إزالة فواصل
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();

  var matches = cleaned.match(TL.Config.PNR_REGEX) || [];
  // إزالة المكررات
  var uniq = [];
  for (var i = 0; i < matches.length; i++) {
    if (uniq.indexOf(matches[i]) === -1) uniq.push(matches[i]);
  }
  return uniq;
}

/**
 * استخراج نص PDF عبر Drive.Files.copy({ocr:true})
 * @param {string} fileId
 * @return {string} نص PDF (قد يكون فارغاً)
 */
function extractPdfText_(fileId) {
  var tmpDocId = null;
  try {
    var resource = {
      title: '_tl_tmp_' + fileId + '_' + new Date().getTime(),
      mimeType: 'application/vnd.google-apps.document'
    };

    // المحاولة الأولى: بدون OCR — يعمل أفضل مع PDF نصي (itinerary، تذاكر إلكترونية)
    // ملاحظة: لا يوجد رقم جواز في أي تذكرة — نستخرج الاسم فقط
    var copy = Drive.Files.copy(resource, fileId);
    tmpDocId = copy.id;
    var text = DocumentApp.openById(tmpDocId).getBody().getText();

    // المحاولة الثانية: لو النص قصير جداً → PDF ممسوح ضوئياً → جرّب OCR
    if (!text || text.length < 50) {
      try { Drive.Files.remove(tmpDocId); } catch(e) {}
      tmpDocId = null;
      resource.title = '_tl_tmp_ocr_' + fileId + '_' + new Date().getTime();
      copy = Drive.Files.copy(resource, fileId, { ocr: true, ocrLanguage: 'en' });
      tmpDocId = copy.id;
      text = DocumentApp.openById(tmpDocId).getBody().getText();
    }

    return text || '';
  } catch (e) {
    Logger.log('extractPdfText_ error for ' + fileId + ': ' + e.message);
    return '';
  } finally {
    if (tmpDocId) {
      try {
        Drive.Files.remove(tmpDocId);
      } catch (e) {
        // محاولة بديلة: نقل للسلة
        try { DriveApp.getFileById(tmpDocId).setTrashed(true); } catch (e2) {}
      }
    }
  }
}

/**
 * استخراج PNR من نص PDF عبر patterns معروفة
 * @param {string} text
 * @return {string|null}
 */
function extractPnrFromText_(text) {
  if (!text) return null;

  var patterns = TL.Config.PNR_TEXT_PATTERNS;
  for (var i = 0; i < patterns.length; i++) {
    var m = text.match(patterns[i]);
    if (m && m[1]) return m[1].toUpperCase();
  }
  return null;
}

/**
 * الدالة الرئيسية: استخراج PNR من PDF واحد
 * @param {Object} pdfInfo {fileId, fileName, folderName}
 * @return {Object} {pnr, source, candidates, extractedText}
 */
function resolvePnr_(pdfInfo) {
  var result = {
    fileId: pdfInfo.fileId,
    fileName: pdfInfo.fileName,
    folderName: pdfInfo.folderName,
    pnr: null,
    source: null,         // 'filename' | 'folder' | 'text' | null
    candidates: [],       // كل PNRs المحتملة (للـ matching متعدد PNRs)
    extractedText: ''     // نص PDF (لإعادة الاستخدام في M2)
  };

  // 1. اسم الملف
  var pnrFromFile = extractPnrFromFileName_(pdfInfo.fileName);
  if (pnrFromFile) {
    result.pnr = pnrFromFile;
    result.source = 'filename';
    result.candidates.push(pnrFromFile);
  }

  // 2. اسم المجلد (قد يضيف candidates إضافية)
  var pnrsFromFolder = extractPnrsFromFolderName_(pdfInfo.folderName);
  for (var i = 0; i < pnrsFromFolder.length; i++) {
    if (result.candidates.indexOf(pnrsFromFolder[i]) === -1) {
      result.candidates.push(pnrsFromFolder[i]);
    }
  }
  if (!result.pnr && pnrsFromFolder.length === 1) {
    result.pnr = pnrsFromFolder[0];
    result.source = 'folder';
  }

  // 3. لو لم يُحدَّد PNR قاطع بعد، نقرأ نص PDF
  if (!result.pnr || result.candidates.length > 1) {
    var text = extractPdfText_(pdfInfo.fileId);
    result.extractedText = text;

    var pnrFromText = extractPnrFromText_(text);
    if (pnrFromText) {
      // النص هو الأكثر موثوقية — يحسم الترشيح
      result.pnr = pnrFromText;
      result.source = 'text';
      if (result.candidates.indexOf(pnrFromText) === -1) {
        result.candidates.push(pnrFromText);
      }
    }
  }

  return result;
}

// ==================== اختبارات سريعة ====================

/**
 * اختبار يدوي: مرّر fileId و fileName و folderName كـ JSON
 * مثال: ?fn=testResolvePnr&args=["fileId","9HN9OB.pdf","Bulgaria"]
 */
function testResolvePnr(fileId, fileName, folderName) {
  var result = resolvePnr_({
    fileId: fileId,
    fileName: fileName,
    folderName: folderName
  });
  // قَصِّر النص في الإخراج لتجنب التضخم
  if (result.extractedText && result.extractedText.length > 500) {
    result.extractedTextSample = result.extractedText.substring(0, 500);
    delete result.extractedText;
  }
  return result;
}

/**
 * اختبار على عينة من ملفات Drive — يقرأ أول 5 ملفات من المجلد الرئيسي
 * ويُرجع النتائج بدون نص PDF كامل
 */
function testResolvePnrSample() {
  var folder = DriveApp.getFolderById(TL.Config.TICKETS_FOLDER_ID);
  var results = [];
  var count = 0;
  var maxSamples = 5;

  // امشِ على المجلدات الفرعية
  var subFolders = folder.getFolders();
  while (subFolders.hasNext() && count < maxSamples) {
    var subFolder = subFolders.next();
    var pdfs = subFolder.getFilesByType(MimeType.PDF);
    while (pdfs.hasNext() && count < maxSamples) {
      var pdf = pdfs.next();
      var res = resolvePnr_({
        fileId: pdf.getId(),
        fileName: pdf.getName(),
        folderName: subFolder.getName()
      });
      // لا نُرجع النص الكامل في النتيجة العامة
      results.push({
        fileName: res.fileName,
        folderName: res.folderName,
        pnr: res.pnr,
        source: res.source,
        candidates: res.candidates,
        textLength: res.extractedText ? res.extractedText.length : 0
      });
      count++;
    }
  }

  return {
    sampleSize: results.length,
    results: results
  };
}
