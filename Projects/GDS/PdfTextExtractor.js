/**
 * PdfTextExtractor.js — استخراج نص من PDF عبر OCR
 *
 * الطريقة: تحويل PDF → Google Doc مع OCR → قراءة النص → حذف الـ Doc المؤقت
 * يتطلب: Drive Advanced Service مُفعَّل في appsscript.json
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.PdfTextExtractor = {
  /**
   * @param {string} fileId - معرف ملف PDF في Drive
   * @param {string} language - 'ar' للعربية، 'en' للإنجليزية (افتراضي 'ar')
   * @return {Object} { status, text?, length?, reason? }
   */
  extract: function(fileId, language) {
    if (!fileId) return { status: 'error', reason: 'empty_fileId' };

    var startTime = new Date();
    var tempDocId = null;

    try {
      // إنشاء نسخة Google Doc مع OCR
      var resource = {
        title: 'temp_ocr_' + new Date().getTime()
      };
      var docFile = Drive.Files.copy(resource, fileId, {
        ocr: true,
        ocrLanguage: language || 'ar'
      });
      tempDocId = docFile.id;

      // استخراج النص
      var doc = DocumentApp.openById(tempDocId);
      var text = doc.getBody().getText();

      return {
        status: 'ok',
        text: text,
        length: text.length,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    } catch (e) {
      return {
        status: 'error',
        reason: 'exception',
        exception: e.message,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    } finally {
      if (tempDocId) {
        try {
          DriveApp.getFileById(tempDocId).setTrashed(true);
        } catch (cleanupErr) {
          GDS2.Log.warn('PdfTextExtractor cleanup failed', {
            tempDocId: tempDocId,
            error: cleanupErr.message
          });
        }
      }
    }
  }
};
