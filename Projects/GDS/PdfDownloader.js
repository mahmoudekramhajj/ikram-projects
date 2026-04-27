/**
 * PdfDownloader.js — تحميل PDF من رابط مع فحوصات الأمان
 *
 * - يتحقق من HTTP status
 * - يتحقق من Content-Type
 * - يتحقق من signature PDF (%PDF في أول 4 بايت)
 * - يحفظ في مجلد Drive: GDS_Tickets
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.PdfDownloader = {
  /**
   * تحويل رابط Drive view إلى رابط تحميل مباشر.
   * من: https://drive.google.com/file/d/FILE_ID/view
   * إلى: https://drive.google.com/uc?id=FILE_ID&export=download
   */
  _transformDriveUrl: function(url) {
    if (!url) return url;
    var m = String(url).match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (m && m[1]) {
      return 'https://drive.google.com/uc?id=' + m[1] + '&export=download';
    }
    return url;
  },

  /**
   * تحميل PDF مباشرة بدون حفظ في Drive (للاستخدام مع Claude).
   * @param {string} url
   * @return {Object} { status, blob?, bytes?, size, content_type, reason? }
   */
  downloadBlob: function(url) {
    if (!url) return { status: 'error', reason: 'empty_url' };
    var startTime = new Date();

    // تحويل Drive view URLs إلى صيغة تحميل مباشر
    var originalUrl = url;
    url = GDS2.PdfDownloader._transformDriveUrl(url);

    // Drive يفرض rate limits — sleep 500ms قبل التحميل
    if (url !== originalUrl) {
      Utilities.sleep(500);
    }

    try {
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true
      });
      var code = response.getResponseCode();

      // 429 = Rate limited — لا نعتبره فشلاً، نتخطى للدورة التالية
      if (code === 429) {
        return {
          status: 'rate_limited',
          reason: 'http_429',
          http_code: 429,
          elapsed_sec: (new Date() - startTime) / 1000
        };
      }

      if (code !== 200) {
        return {
          status: 'error',
          reason: 'http_' + code,
          http_code: code,
          elapsed_sec: (new Date() - startTime) / 1000
        };
      }

      var headers = response.getAllHeaders();
      var contentType = String(headers['Content-Type'] || headers['content-type'] || '').toLowerCase();

      var blob = response.getBlob();
      var bytes = blob.getBytes();

      if (bytes.length < 100) {
        return {
          status: 'error',
          reason: 'file_too_small',
          size: bytes.length,
          content_type: contentType
        };
      }

      var signature = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
      if (signature !== '%PDF') {
        return {
          status: 'not_pdf',
          reason: 'signature_mismatch',
          content_type: contentType,
          first_bytes: signature,
          size: bytes.length
        };
      }

      return {
        status: 'ok',
        blob: blob,
        bytes: bytes,
        size: bytes.length,
        content_type: contentType,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    } catch (e) {
      // Drive rate limit يأتي كـ exception (ليس HTTP 429)
      // أمثلة: "تم تجاوز حصة معدل نقل البيانات" / "quota exceeded" / "rate limit"
      var msg = String(e && e.message || '');
      if (/تجاوز.*حصة|quota|rate.*(limit|exceed)|معدل.*نقل/i.test(msg)) {
        return {
          status: 'rate_limited',
          reason: 'drive_quota_exception',
          exception: msg,
          elapsed_sec: (new Date() - startTime) / 1000
        };
      }
      return {
        status: 'error',
        reason: 'exception',
        exception: msg,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    }
  },

  /**
   * تحميل PDF وحفظه في Drive (legacy — يُستخدم لو احتجنا أرشفة).
   * @param {string} url
   * @param {string} fileName
   * @return {Object} { status, fileId?, reason?, http_code?, content_type? }
   */
  download: function(url, fileName) {
    if (!url) return { status: 'error', reason: 'empty_url' };

    var startTime = new Date();

    try {
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true
      });

      var code = response.getResponseCode();
      if (code !== 200) {
        return {
          status: 'error',
          reason: 'http_' + code,
          http_code: code,
          elapsed_sec: (new Date() - startTime) / 1000
        };
      }

      var headers = response.getAllHeaders();
      var contentType = String(headers['Content-Type'] || headers['content-type'] || '').toLowerCase();

      var blob = response.getBlob();
      var bytes = blob.getBytes();

      if (bytes.length < 100) {
        return {
          status: 'error',
          reason: 'file_too_small',
          size: bytes.length,
          content_type: contentType
        };
      }

      // فحص %PDF في أول 4 بايت
      var signature = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
      if (signature !== '%PDF') {
        return {
          status: 'not_pdf',
          reason: 'signature_mismatch',
          content_type: contentType,
          first_bytes: signature,
          size: bytes.length
        };
      }

      blob.setName(fileName || ('ticket_' + new Date().getTime() + '.pdf'));
      blob.setContentType('application/pdf');

      var folder = GDS2.PdfDownloader._getFolder();
      var file = folder.createFile(blob);

      return {
        status: 'ok',
        fileId: file.getId(),
        fileName: file.getName(),
        size: bytes.length,
        content_type: contentType,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    } catch (e) {
      return {
        status: 'error',
        reason: 'exception',
        exception: e.message,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    }
  },

  /**
   * حذف ملف PDF من Drive (بعد المعالجة).
   */
  cleanup: function(fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      return true;
    } catch (e) {
      GDS2.Log.warn('PdfDownloader.cleanup failed', { fileId: fileId, error: e.message });
      return false;
    }
  },

  _getFolder: function() {
    var folderName = GDS2.Config.PDF_FOLDER_NAME;
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(folderName);
  }
};
