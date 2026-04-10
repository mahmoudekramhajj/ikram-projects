/**
 * TicketProcessor.js — معالج تذاكر الطيران
 * يقرأ ملفات PDF من مجلد Drive، يستخرج بيانات الراكب،
 * ينشئ رابط مشاركة، ويضع البيانات في الصف المناسب بالشيت
 *
 * يدعم: طيران الخليج (Gulf Air) + الخطوط التركية (Turkish Airlines)
 */

// ==========================================
// الدالة الرئيسية
// ==========================================

function processTickets() {
  var startTime = new Date();
  var MAX_RUNTIME_MS = 5 * 60 * 1000; // 5 دقائق (حد GAS هو 6 دقائق)

  var folder = DriveApp.getFolderById(CONFIG.TICKETS_FOLDER_ID);
  var files = folder.getFilesByType('application/pdf');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    Logger.log('❌ الورقة غير موجودة: ' + CONFIG.SHEET_NAME);
    return { error: 'Sheet not found' };
  }

  // تحميل بيانات الأسماء والعقود من الشيت
  var lastRow = sheet.getLastRow();
  var namesData = sheet.getRange(2, CONFIG.COL_FIRST_NAME, lastRow - 1, 2).getValues(); // K & L
  var contractData = sheet.getRange(2, CONFIG.COL_CONTRACT, lastRow - 1, 1).getValues(); // V

  var processedFolder = getOrCreateProcessedFolder_(folder);
  var results = { processed: [], skipped: [], errors: [] };

  while (files.hasNext()) {
    // فحص المهلة الزمنية
    var elapsed = new Date() - startTime;
    if (elapsed > MAX_RUNTIME_MS) {
      Logger.log('⏰ توقف مؤقت — انتهت المهلة بعد ' + Math.round(elapsed / 1000) + ' ثانية');
      Logger.log('💡 شغّل processTickets مرة أخرى لإكمال الباقي');
      break;
    }

    var file = files.next();
    var fileName = file.getName();
    Logger.log('📄 معالجة: ' + fileName);

    try {
      // 1. استخراج النص من PDF
      var text = extractTextFromPDF_(file.getId());
      if (!text || text.trim().length === 0) {
        results.errors.push({ file: fileName, error: 'لا يمكن استخراج نص من الملف' });
        Logger.log('⚠️ لا نص في: ' + fileName);
        continue;
      }

      // 2. تحديد نوع التذكرة واستخراج البيانات
      var ticketData = parseTicketData_(text, fileName);
      if (!ticketData) {
        results.errors.push({ file: fileName, error: 'لا يمكن تحديد نوع التذكرة' });
        Logger.log('⚠️ نوع تذكرة غير معروف: ' + fileName);
        continue;
      }

      Logger.log('✈️ ' + ticketData.airline + ' — ' + ticketData.firstName + ' ' + ticketData.lastName);
      Logger.log('📋 Booking: ' + ticketData.bookingRef + ' | Ticket: ' + ticketData.ticketNumber);

      // 3. البحث عن الصف المطابق (بالاسم + العقد كاحتياط)
      var matchRow = findMatchingRow_(namesData, ticketData.firstName, ticketData.lastName, contractData, ticketData.bookingRef);
      if (matchRow === -1) {
        results.skipped.push({
          file: fileName,
          name: ticketData.firstName + ' ' + ticketData.lastName,
          reason: 'اسم غير موجود في الجدول'
        });
        Logger.log('⚠️ لم يتم العثور على: ' + ticketData.firstName + ' ' + ticketData.lastName);
        continue;
      }

      var sheetRow = matchRow + 2; // +2 لأن namesData تبدأ من صف 0 والبيانات من صف 2

      // 4. إنشاء رابط مشاركة
      var shareLink = createShareLink_(file);

      // 5. كتابة البيانات في الشيت
      sheet.getRange(sheetRow, CONFIG.COL_TICKET_LINK).setValue(shareLink);
      sheet.getRange(sheetRow, CONFIG.COL_BOOKING_REF).setValue(ticketData.bookingRef);
      sheet.getRange(sheetRow, CONFIG.COL_TICKET_NUM).setValue(ticketData.ticketNumber);

      // 6. نقل الملف لمجلد "تمت المعالجة"
      file.moveTo(processedFolder);

      results.processed.push({
        file: fileName,
        name: ticketData.firstName + ' ' + ticketData.lastName,
        row: sheetRow,
        bookingRef: ticketData.bookingRef,
        ticketNumber: ticketData.ticketNumber
      });

      Logger.log('✅ تم — صف ' + sheetRow);

    } catch (e) {
      results.errors.push({ file: fileName, error: e.message });
      Logger.log('❌ خطأ في ' + fileName + ': ' + e.message);
    }
  }

  // ملخص
  Logger.log('');
  Logger.log('========== الملخص ==========');
  Logger.log('✅ تمت معالجتها: ' + results.processed.length);
  Logger.log('⏭️ تم تخطيها: ' + results.skipped.length);
  Logger.log('❌ أخطاء: ' + results.errors.length);

  return results;
}


// ==========================================
// استخراج النص من PDF
// ==========================================

function extractTextFromPDF_(fileId) {
  // تحويل PDF إلى Google Doc عبر OCR ثم قراءة النص
  var resource = {
    title: 'temp_ocr_' + new Date().getTime(),
    mimeType: 'application/vnd.google-apps.document'
  };

  var options = {
    ocr: true,
    ocrLanguage: 'en'
  };

  var docFile = Drive.Files.copy(resource, fileId, options);
  var doc = DocumentApp.openById(docFile.id);
  var text = doc.getBody().getText();

  // حذف الملف المؤقت
  DriveApp.getFileById(docFile.id).setTrashed(true);

  return text;
}


// ==========================================
// تحليل بيانات التذكرة
// ==========================================

function parseTicketData_(text, fileName) {
  // محاولة Gulf Air أولاً
  var gulfAir = parseGulfAir_(text, fileName);
  if (gulfAir) return gulfAir;

  // محاولة Turkish Airlines
  var turkish = parseTurkish_(text, fileName);
  if (turkish) return turkish;

  // محاولة أخيرة: استخراج الاسم من اسم الملف (Electronic ticket receipt... for MR/MRS NAME)
  var fileMatch = fileName.match(/for\s+(?:MR|MRS|MS|MISS|MSTR)\s+(.+?)\.pdf$/i);
  if (fileMatch) {
    var fullName = fileMatch[1].trim().toUpperCase();
    var parts = fullName.split(/\s+/);
    if (parts.length >= 2) {
      var data = {
        airline: 'Unknown (from filename)',
        lastName: parts[parts.length - 1],
        firstName: parts.slice(0, -1).join(' '),
        bookingRef: '',
        ticketNumber: ''
      };
      // محاولة استخراج Booking Ref و Ticket Number من النص
      var bookingMatch = text.match(/RESERVATION\s*CODE\s*\n?\s*([A-Z0-9]{5,8})/i);
      if (bookingMatch) data.bookingRef = bookingMatch[1].trim();
      var ticketMatch = text.match(/TICKET\s*NUMBER\s*\n?\s*(\d{10,14})/i);
      if (ticketMatch) data.ticketNumber = ticketMatch[1].trim();
      return data;
    }
  }

  return null;
}

/**
 * تحليل تذكرة طيران الخليج
 * الاسم بصيغة: RAFIQ/ABDUR MR (بعد "Prepared For")
 * رقم الحجز: RESERVATION CODE → EIWXRA
 * رقم التذكرة: TICKET NUMBER → 0722136458341
 */
function parseGulfAir_(text, fileName) {
  if (text.indexOf('GULF AIR') === -1 && text.indexOf('Gulf Air') === -1) {
    return null;
  }

  var data = { airline: 'Gulf Air', firstName: '', lastName: '', bookingRef: '', ticketNumber: '' };

  // استخراج الاسم — "Prepared For" ثم LASTNAME/FIRSTNAME TITLE
  // أو بصيغة LASTNAME/FIRSTNAME
  var namePatterns = [
    /Prepared\s*For\s*\n?\s*([A-Z]+)\s*\/\s*([A-Z]+)\s*(?:MR|MRS|MS|MISS|MSTR|INF)?/i,
    /([A-Z]+)\s*\/\s*([A-Z]+)\s*(?:MR|MRS|MS|MISS|MSTR|INF)/i
  ];

  for (var i = 0; i < namePatterns.length; i++) {
    var nameMatch = text.match(namePatterns[i]);
    if (nameMatch) {
      data.lastName = nameMatch[1].trim();
      data.firstName = nameMatch[2].trim();
      break;
    }
  }

  // استخراج رقم الحجز
  var bookingMatch = text.match(/RESERVATION\s*CODE\s*\n?\s*([A-Z0-9]{5,8})/i);
  if (bookingMatch) {
    data.bookingRef = bookingMatch[1].trim();
  }

  // استخراج رقم التذكرة
  var ticketMatch = text.match(/TICKET\s*NUMBER\s*\n?\s*(\d{10,14})/i);
  if (ticketMatch) {
    data.ticketNumber = ticketMatch[1].trim();
  }

  if (data.firstName && data.lastName) {
    return data;
  }

  return null;
}

/**
 * تحليل تذكرة الخطوط التركية
 * الاسم بصيغة: Passenger Name : AYOUB ALI
 * رقم الحجز: Booking Ref : SQ73H4 أو Rezervasyon No
 * رقم التذكرة: Ticket Number : 2352297111540 أو Bilet No
 */
function parseTurkish_(text, fileName) {
  if (text.indexOf('TURKISH') === -1 && text.indexOf('Turkish') === -1 &&
      text.indexOf('THY') === -1 && text.indexOf('TK ') === -1) {
    return null;
  }

  var data = { airline: 'Turkish Airlines', firstName: '', lastName: '', bookingRef: '', ticketNumber: '' };

  // استخراج الاسم — "Passenger Name : FIRSTNAME LASTNAME"
  // نستبعد اختصارات شركات الطيران (THY, TK, GF) ونأخذ كلمتين فقط
  var namePatterns = [
    /Yolcu\s*ismi\s*\/?\s*Passenger\s*Name\s*:?\s*([A-Z]{2,})\s+([A-Z]{2,})/i,
    /Passenger\s*Name\s*:?\s*([A-Z]{2,})\s+([A-Z]{2,})/i
  ];

  // كلمات يجب استبعادها من الاسم
  var excludeWords = ['THY', 'TK', 'GF', 'TURKISH', 'AIRLINES', 'GULF', 'AIR'];

  for (var i = 0; i < namePatterns.length; i++) {
    var nameMatch = text.match(namePatterns[i]);
    if (nameMatch) {
      data.firstName = nameMatch[1].trim();
      data.lastName = nameMatch[2].trim();
      // إزالة أي كلمة مستبعدة
      if (excludeWords.indexOf(data.firstName.toUpperCase()) !== -1) data.firstName = '';
      if (excludeWords.indexOf(data.lastName.toUpperCase()) !== -1) data.lastName = '';
      if (data.firstName && data.lastName) break;
    }
  }

  // استخراج رقم الحجز
  var bookingPatterns = [
    /Booking\s*Ref\s*:?\s*([A-Z0-9]{5,8})/i,
    /Rezervasyon\s*No\s*\/?\s*Booking\s*Ref\s*:?\s*([A-Z0-9]{5,8})/i
  ];

  for (var i = 0; i < bookingPatterns.length; i++) {
    var bookingMatch = text.match(bookingPatterns[i]);
    if (bookingMatch) {
      data.bookingRef = bookingMatch[1].trim();
      break;
    }
  }

  // استخراج رقم التذكرة
  var ticketPatterns = [
    /Ticket\s*Number\s*:?\s*(\d{10,14})/i,
    /Bilet\s*No\s*\/?\s*Ticket\s*Number\s*:?\s*(\d{10,14})/i
  ];

  for (var i = 0; i < ticketPatterns.length; i++) {
    var ticketMatch = text.match(ticketPatterns[i]);
    if (ticketMatch) {
      data.ticketNumber = ticketMatch[1].trim();
      break;
    }
  }

  if (data.firstName && data.lastName) {
    return data;
  }

  // محاولة بديلة: استخراج الاسم من اسم الملف
  var fileNameMatch = fileName.match(/([A-Z]+)\s+([A-Z]+)/i);
  if (fileNameMatch) {
    data.firstName = fileNameMatch[1].trim().toUpperCase();
    data.lastName = fileNameMatch[2].trim().toUpperCase();
    if (data.firstName && data.lastName) {
      return data;
    }
  }

  return null;
}


// ==========================================
// البحث عن الصف المطابق
// ==========================================

function findMatchingRow_(namesData, firstName, lastName, contractData, bookingRef) {
  firstName = firstName.toUpperCase().trim();
  lastName = lastName.toUpperCase().trim();

  // دالة مساعدة: إزالة المسافات للمقارنة
  function noSpace(s) { return s.replace(/\s+/g, ''); }

  // الاسم الكامل من التذكرة (بدون مسافات)
  var ticketFullNoSpace = noSpace(firstName + lastName);

  for (var i = 0; i < namesData.length; i++) {
    var sheetFirst = String(namesData[i][0]).toUpperCase().trim();
    var sheetLast = String(namesData[i][1]).toUpperCase().trim();

    // 1. مطابقة تامة
    if (sheetFirst === firstName && sheetLast === lastName) {
      return i;
    }

    // 2. مطابقة معكوسة (الاسم والعائلة معكوسين)
    if (sheetFirst === lastName && sheetLast === firstName) {
      return i;
    }

    // 3. مطابقة بدون مسافات (MDABDUL = MD ABDUL، ALAREEF = AL AREEF)
    var sheetFullNoSpace = noSpace(sheetFirst + sheetLast);
    if (sheetFullNoSpace === ticketFullNoSpace) {
      return i;
    }

    // 4. مطابقة معكوسة بدون مسافات
    var ticketFullReversedNoSpace = noSpace(lastName + firstName);
    if (sheetFullNoSpace === ticketFullReversedNoSpace) {
      return i;
    }

    // 5. مطابقة جزئية — الاسم الأول من التذكرة يحتوي على اسم الشيت (أو العكس)
    if (sheetLast === lastName &&
        (noSpace(sheetFirst) === noSpace(firstName) ||
         sheetFirst.indexOf(firstName) !== -1 ||
         firstName.indexOf(sheetFirst) !== -1)) {
      return i;
    }

    // 6. نفس الشيء مع الأسماء المعكوسة
    if (sheetLast === firstName &&
        (noSpace(sheetFirst) === noSpace(lastName) ||
         sheetFirst.indexOf(lastName) !== -1 ||
         lastName.indexOf(sheetFirst) !== -1)) {
      return i;
    }
  }

  // 7. البحث عبر Booking Ref في عمود العقد (V) كملاذ أخير
  if (bookingRef && contractData) {
    var candidates = [];
    for (var i = 0; i < contractData.length; i++) {
      var contract = String(contractData[i][0]).toUpperCase().trim();
      if (contract.indexOf(bookingRef.toUpperCase()) !== -1) {
        // وجدنا تطابق في العقد — نتحقق من الاسم بشكل أوسع
        var sheetFirst = String(namesData[i][0]).toUpperCase().trim();
        var sheetLast = String(namesData[i][1]).toUpperCase().trim();
        var sheetFullNoSpace = noSpace(sheetFirst + sheetLast);

        // تحقق بسيط: هل أي جزء من الاسم يتطابق؟
        if (ticketFullNoSpace.indexOf(noSpace(sheetFirst)) !== -1 ||
            ticketFullNoSpace.indexOf(noSpace(sheetLast)) !== -1 ||
            noSpace(lastName + firstName).indexOf(noSpace(sheetFirst)) !== -1) {
          return i;
        }
        candidates.push(i);
      }
    }
  }

  return -1;
}


// ==========================================
// إنشاء رابط مشاركة
// ==========================================

function createShareLink_(file) {
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}


// ==========================================
// مجلد الملفات المعالجة
// ==========================================

function getOrCreateProcessedFolder_(parentFolder) {
  var folders = parentFolder.getFoldersByName(CONFIG.PROCESSED_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(CONFIG.PROCESSED_FOLDER_NAME);
}


// ==========================================
// أدوات مساعدة
// ==========================================

/**
 * اختبار على ملف واحد — لا يكتب في الشيت
 */
function testSingleFile() {
  var folder = DriveApp.getFolderById(CONFIG.TICKETS_FOLDER_ID);
  var files = folder.getFilesByType('application/pdf');

  if (!files.hasNext()) {
    Logger.log('لا توجد ملفات PDF في المجلد');
    return;
  }

  var file = files.next();
  Logger.log('📄 اختبار ملف: ' + file.getName());

  var text = extractTextFromPDF_(file.getId());
  Logger.log('--- النص المستخرج ---');
  Logger.log(text.substring(0, 2000));
  Logger.log('--- نهاية النص ---');

  var ticketData = parseTicketData_(text, file.getName());
  if (ticketData) {
    Logger.log('✈️ شركة الطيران: ' + ticketData.airline);
    Logger.log('👤 الاسم: ' + ticketData.firstName + ' ' + ticketData.lastName);
    Logger.log('📋 Booking Ref: ' + ticketData.bookingRef);
    Logger.log('🎫 Ticket Number: ' + ticketData.ticketNumber);
  } else {
    Logger.log('❌ لم يتم التعرف على التذكرة');
  }

  return ticketData;
}

/**
 * عرض قائمة الملفات في المجلد
 */
function listTicketFiles() {
  var folder = DriveApp.getFolderById(CONFIG.TICKETS_FOLDER_ID);
  var files = folder.getFilesByType('application/pdf');
  var list = [];

  while (files.hasNext()) {
    var file = files.next();
    list.push({
      name: file.getName(),
      id: file.getId(),
      size: file.getSize(),
      date: file.getDateCreated()
    });
  }

  Logger.log('عدد الملفات: ' + list.length);
  list.forEach(function(f) {
    Logger.log('  📄 ' + f.name + ' (' + Math.round(f.size / 1024) + ' KB)');
  });

  return list;
}

/**
 * عرض الأسماء في الشيت (للتحقق)
 */
function listSheetNames() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, CONFIG.COL_FIRST_NAME, lastRow - 1, 2).getValues();

  Logger.log('عدد الأسماء: ' + data.length);
  data.forEach(function(row, i) {
    if (row[0] || row[1]) {
      Logger.log('  صف ' + (i + 2) + ': ' + row[0] + ' ' + row[1]);
    }
  });

  return data.length;
}


// ==========================================
// Web App Entry Points
// ==========================================

function doGet(e) {
  return handleClaudeAPI_(e);
}

function doPost(e) {
  return handleClaudeAPI_(e);
}
