/**
 * NusukParser.js — تحليل تذاكر نسك الإلكترونية (PDF)
 *
 * يستخرج من PDF نسك:
 * - Booking ID (مثل A26041250043250)
 * - PNR (Airline Reference)
 * - اسم المسافر
 * - رحلات الذهاب والعودة (كل leg بتفاصيله)
 */


// كلمات وهمية — لو ظهرت في الاسم نعتبره فاشل
// أكواد مطارات IATA المستخدمة في رحلات الحج
var AIRPORT_CODES = [
  'jed', 'med', 'dxb', 'doh', 'ist', 'saw', 'cpt', 'dur', 'jnb',
  'man', 'lhr', 'lgw', 'stn', 'bhx', 'cdg', 'fra', 'ams', 'bru',
  'cai', 'auh', 'kul', 'sin', 'bom', 'del', 'lhe', 'isb', 'khi',
  'jfk', 'ord', 'iah', 'dmm', 'ruh', 'mct', 'bah', 'kwi', 'amm',
  'bei', 'cas', 'tun', 'alg', 'acc', 'los', 'abj', 'dkr', 'cmn',
  'add', 'nbo', 'dar', 'ebb', 'kgl', 'bkk', 'cgk', 'mnl', 'hnd',
  'icn', 'pek', 'pvg', 'syd', 'mel', 'yvr', 'yyz', 'mxp', 'fco',
  'mad', 'bcn', 'lis', 'ath', 'vie', 'zrh', 'gen', 'muc', 'ham',
  'dus', 'cph', 'osl', 'arn', 'hel', 'waw', 'prg', 'bud', 'svo',
  'led', 'tbs', 'bak', 'ika', 'ssh', 'hrg'
];

var FAKE_NAME_WORDS = [
  'your', 'fli', 'flight', 'itinerary', 'please', 'dear', 'ticket',
  'booking', 'passenger', 'travel', 'nusuk', 'hajj', 'schedule',
  'change', 'details', 'attached', 'confirmed', 'updated', 'new',
  'the', 'for', 'and', 'with', 'from', 'this', 'that', 'are',
  'economy', 'business', 'class', 'baggage', 'allowance'
];


/**
 * استخراج النص من PDF عبر OCR
 * نفس الآلية المستخدمة في TicketProcessor
 */
function extractTextFromPDF_(fileId) {
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


/**
 * تحليل نص تذكرة نسك واستخراج البيانات
 */
function parseNusukTicket_(text) {
  if (!text || text.trim().length === 0) return null;

  var data = {
    bookingId: '',
    pnr: '',
    passengerName: '',
    firstName: '',
    lastName: '',
    outboundLegs: [],
    returnLegs: [],
    arrivalCity: '',
    arrivalDate: '',
    arrivalTime: '',
    departureCity: '',
    departureDate: '',
    departureTime: ''
  };

  // === 1. استخراج Booking ID ===
  var bookingMatch = text.match(/(?:Booking\s*ID|Nusuk\s*Hajj\s*ID)[:\s]*\n?\s*(A\d{14,17})/i);
  if (bookingMatch) {
    data.bookingId = bookingMatch[1].trim();
  }

  // === 2. استخراج PNR ===
  var pnrPatterns = [
    /(?:Airline\s*Reference\s*\(?PNR\)?)[:\s]*\n?\s*([A-Z0-9]{5,8})/i,
    /PNR[:\s#]*([A-Z0-9]{5,8})/i,
    /(?:Reservation|Booking)\s*(?:Code|Ref(?:erence)?)[:\s]*\n?\s*([A-Z0-9]{5,8})/i
  ];
  for (var p = 0; p < pnrPatterns.length; p++) {
    var pnrMatch = text.match(pnrPatterns[p]);
    if (pnrMatch) {
      data.pnr = pnrMatch[1].trim();
      break;
    }
  }

  // === 3. استخراج جميع المسافرين ===
  data.passengers = extractAllPassengers_(text);

  // التوافق — أول مسافر يُستخدم كـ الاسم الرئيسي
  if (data.passengers.length > 0) {
    data.passengerName = data.passengers[0].fullName;
    data.firstName = data.passengers[0].firstName;
    data.lastName = data.passengers[0].lastName;
  }

  // === 4. استخراج الرحلات ===
  var legs = extractFlightLegs_(text);
  classifyLegs_(data, legs);

  // === 5. تحديد تاريخ الوصول والمغادرة للمملكة ===
  calculateKSADates_(data);

  return data;
}


/**
 * استخراج جميع أسماء المسافرين من النص
 * يدعم: "NAME NAME (Adult)" و "Mr NAME NAME (Adult)"
 */
function extractAllPassengers_(text) {
  var passengers = [];
  var seen = {};

  // نمط شامل: اسم (بلقب أو بدون) ثم (Adult/Child/Infant)
  var allNamesPattern = /(?:(?:Mr|Mrs|Ms|Miss|Mstr|Master|DR)\s+)?([A-Z][A-Z\s]{2,40}?)\s*\((?:Adult|Child|Infant|ADT|CHD|INF)\)/g;
  var match;

  while ((match = allNamesPattern.exec(text)) !== null) {
    var raw = match[1].trim();
    var cleaned = cleanName_(raw);

    if (!cleaned || !isValidName_(cleaned)) continue;

    // تجنب التكرار
    var key = cleaned.toUpperCase();
    if (seen[key]) continue;
    seen[key] = true;

    // تقسيم إلى أول + عائلة
    var parts = cleaned.split(/\s+/);
    var firstName, lastName;
    if (parts.length >= 2) {
      firstName = parts[0];
      lastName = parts.slice(1).join(' ');
    } else {
      firstName = cleaned;
      lastName = '';
    }

    passengers.push({
      fullName: cleaned,
      firstName: firstName,
      lastName: lastName
    });
  }

  // لو ما لقينا شيء — نجرب أنماط بديلة
  if (passengers.length === 0) {
    var fallbackPatterns = [
      /Passenger\s*Name[:\s]*([A-Z][A-Z\s\/]{2,40})/ig,
      /([A-Z]{2,20})\/([A-Z]{2,20})\s*(?:MR|MRS|MS|MISS|MSTR)/ig,
      /(?:Prepared\s*For|Issued\s*To)[:\s]*\n?\s*([A-Z][A-Z\s\/]{2,40})/ig
    ];

    for (var f = 0; f < fallbackPatterns.length; f++) {
      while ((match = fallbackPatterns[f].exec(text)) !== null) {
        var name;
        if (f === 1 && match[2]) {
          name = match[2] + ' ' + match[1]; // LAST/FIRST → FIRST LAST
        } else {
          name = match[1].trim();
        }
        name = cleanName_(name);
        if (!name || !isValidName_(name)) continue;

        var key2 = name.toUpperCase();
        if (seen[key2]) continue;
        seen[key2] = true;

        var p2 = name.split(/\s+/);
        passengers.push({
          fullName: name,
          firstName: p2[0],
          lastName: p2.length >= 2 ? p2.slice(1).join(' ') : ''
        });
      }
    }
  }

  return passengers;
}


/**
 * التحقق أن الاسم حقيقي وليس نص عشوائي من PDF
 */
function isValidName_(name) {
  if (!name || name.length < 3) return false;

  var lower = name.toLowerCase().trim();
  var words = lower.split(/\s+/);

  // كل كلمة في الاسم يجب أن لا تكون وهمية
  for (var i = 0; i < words.length; i++) {
    if (FAKE_NAME_WORDS.indexOf(words[i]) !== -1) {
      return false;
    }
  }

  // رفض إذا الاسم كله كود مطار فقط (كلمة واحدة = 3 أحرف)
  if (words.length === 1 && AIRPORT_CODES.indexOf(words[0]) !== -1) {
    return false;
  }

  // الاسم يجب أن يحتوي حرفين كبيرين على الأقل
  var upperCount = (name.match(/[A-Z]/g) || []).length;
  if (upperCount < 2) return false;

  // يجب أن لا يحتوي أرقام
  if (/\d/.test(name)) return false;

  return true;
}


/**
 * تنظيف الاسم — إزالة ألقاب وكلمات غير ضرورية
 */
/**
 * تنظيف خفيف — ألقاب فقط (للمحاولة الأولى بدون فلاتر)
 */
function lightClean_(name) {
  name = name.replace(/\b(?:MR|MRS|MS|MISS|MSTR|MASTER|DR|PROF|SIR)\b/gi, '').trim();
  name = name.replace(/\s{2,}/g, ' ').trim();
  name = name.replace(/^\/|\/$/g, '').trim();
  return name.length >= 3 ? name : '';
}


/**
 * تنظيف عميق — ألقاب + أكواد مطارات من البداية (للمحاولة الثانية)
 * مثال: "DXB DUR FATIMA KATHRADA" → "FATIMA KATHRADA"
 * لكن: "FAIQ DAR" يبقى كما هو (DAR اسم عائلة حقيقي)
 */
function deepClean_(name) {
  name = lightClean_(name);
  var nameWords = name.split(/\s+/);
  while (nameWords.length > 1 && AIRPORT_CODES.indexOf(nameWords[0].toLowerCase()) !== -1) {
    nameWords.shift();
  }
  name = nameWords.join(' ');
  return name.length >= 3 ? name : '';
}


// التوافق مع الاستدعاءات القديمة
function cleanName_(name) {
  return lightClean_(name);
}


/**
 * استخراج جميع أجزاء الرحلة من النص
 */
function extractFlightLegs_(text) {
  var legs = [];
  var match;

  // نمط 1: من الإيميل — "From *City (CODE)* to *City (CODE)*"
  var emailPattern = /From\s*\*?([^(]*?)\(([A-Z]{3})\)\*?\s*to\s*\*?([^(]*?)\(([A-Z]{3})\)\*?\s*Departure:\s*([\d:]+\s*[AP]M)\s*\*?([A-Za-z]+\s+\d{1,2}\s+[A-Za-z]+\s+\d{4})\*?\s*Arrival:\s*([\d:]+\s*[AP]M)/gi;
  while ((match = emailPattern.exec(text)) !== null) {
    legs.push({
      flightNumber: '',
      fromCity: match[2].trim(),
      toCity: match[4].trim(),
      depTime: match[5].trim(),
      depDate: match[6].trim(),
      arrTime: match[7].trim(),
      arrDate: ''
    });
  }
  if (legs.length > 0) return legs;

  // نمط 2: من PDF — منطق المرساة (anchor)
  // ====================================
  // كل رحلة في PDF نسك بنية ثابتة:
  //   [تاريخ + وقت + مطار]  ← إقلاع
  //   [رقم الرحلة + المدة]  ← المرساة
  //   [تاريخ + وقت + مطار]  ← وصول
  //
  // استثناء: آخر رحلة (قبل Traveller details) الوصول يظهر قبل الإقلاع:
  //   [تاريخ + وقت + مطار]  ← وصول (يظهر أولاً!)
  //   Traveller details
  //   [تاريخ + وقت + مطار]  ← إقلاع
  //   [رقم الرحلة + المدة]
  // ====================================

  // 1. قطع النص عند بداية القسم العربي (لتجنب التكرار)
  var arabicStart = text.search(/[\u0600-\u06FF]/);
  var workText = arabicStart > 0 ? text.substring(0, arabicStart) : text;

  // 2. جمع كل نقاط (تاريخ + وقت + مطار) مع مواقعها
  var dtPattern = /([A-Z][a-z]+\s+\d{1,2},?\s+\d{4})\s*\n\s*(\d{1,2}:\d{2}\s*[AP]M)\s*\n\s*([A-Z]{3})/g;
  var allDT = [];
  while ((match = dtPattern.exec(workText)) !== null) {
    allDT.push({
      date: match[1].trim(),
      time: match[2].trim(),
      airport: match[3].trim(),
      pos: match.index
    });
  }

  // 3. جمع أرقام الرحلات مع مواقعها ونهاياتها
  //    [-]? بدل [-\s]? لتجنب التقاط "XB 0" من "DXB 07h 40m"
  //    \s+ يشمل \n لالتقاط "EK-772\n09h 45m"
  var fnPattern = /\b([A-Z]{2}-?\d{1,5})\s+(\d{1,2}h\s*\d{1,2}m)/g;
  var allFN = [];
  while ((match = fnPattern.exec(workText)) !== null) {
    allFN.push({
      flightNum: match[1].replace(/\s+/g, ''),
      pos: match.index,
      endPos: match.index + match[0].length
    });
  }

  // 4. لكل رقم رحلة — إيجاد الإقلاع والوصول
  //
  // البنية في PDF نسك:
  //   حالة عادية:  تاريخ\nوقت\nمطار → [رحلة مدة] → تاريخ\nوقت\nمطار
  //   حالة ترانزيت: تاريخ\nوقت\nمطار → [رحلة مدة\nEconomy\nمطار] (الوصول بدون تاريخ)
  //   آخر رحلة:   الوصول يظهر قبل الإقلاع (Traveller details بينهم)
  //
  var usedDT = {};

  for (var fi = 0; fi < allFN.length; fi++) {
    var fnPos = allFN[fi].pos;
    var fnEnd = fnPos + allFN[fi].flightNum.length + 10; // نهاية تقريبية لنص الرحلة+المدة

    // أقرب (تاريخ+وقت+مطار) قبل رقم الرحلة = إقلاع
    var depIdx = -1;
    for (var j = allDT.length - 1; j >= 0; j--) {
      if (allDT[j].pos < fnPos && !usedDT[j]) {
        depIdx = j;
        break;
      }
    }

    if (depIdx === -1) continue;
    var dep = allDT[depIdx];

    // البحث عن الوصول — 3 طرق بالأولوية:
    var arrIdx = -1;
    var arrAirport = '';
    var arrDate = '';
    var arrTime = '';

    // طريقة 1 (أولوية عليا): كود مطار بعد "Economy" مباشرة (حالة الترانزيت)
    // في الترانزيت: رقم_رحلة مدة\nEconomy\nمطار_وصول (بدون تاريخ)
    var afterFn = workText.substring(allFN[fi].endPos, allFN[fi].endPos + 80);
    var transitMatch = afterFn.match(/Economy\s*\n\s*([A-Z]{3})\b/);
    if (transitMatch && transitMatch[1] !== dep.airport) {
      arrAirport = transitMatch[1];
    }

    // طريقة 2: أقرب (تاريخ+وقت+مطار) بعد رقم الرحلة وقبل الرحلة التالية
    if (!arrAirport) {
      var nextFnPos = (fi + 1 < allFN.length) ? allFN[fi + 1].pos : workText.length;
      for (var j = 0; j < allDT.length; j++) {
        if (allDT[j].pos > fnPos && !usedDT[j] && allDT[j].pos < nextFnPos) {
          arrIdx = j;
          arrAirport = allDT[j].airport;
          arrDate = allDT[j].date;
          arrTime = allDT[j].time;
          break;
        }
      }
    }

    // طريقة 3: آخر رحلة — الوصول ظهر قبل الإقلاع (قبل Traveller details)
    if (!arrAirport) {
      for (var j = depIdx - 1; j >= 0; j--) {
        if (!usedDT[j]) {
          arrAirport = allDT[j].airport;
          arrDate = allDT[j].date;
          arrTime = allDT[j].time;
          arrIdx = j;
          break;
        }
      }
    }

    if (!arrAirport) continue;

    // التأكد من منطقية المسار — الإقلاع والوصول مطارات مختلفة
    if (dep.airport === arrAirport) continue;

    // تسجيل كمستخدم
    usedDT[depIdx] = true;
    if (arrIdx !== -1) usedDT[arrIdx] = true;

    legs.push({
      flightNumber: allFN[fi].flightNum,
      fromCity: dep.airport,
      toCity: arrAirport,
      depDate: dep.date,
      depTime: dep.time,
      arrDate: arrDate,
      arrTime: arrTime
    });
  }

  return legs;
}


/**
 * تصنيف الرحلات إلى ذهاب وعودة
 */
function classifyLegs_(data, legs) {
  var ksaAirports = ['JED', 'MED', 'DMM', 'RUH'];
  var foundArrival = false;

  for (var i = 0; i < legs.length; i++) {
    var leg = legs[i];
    if (!foundArrival) {
      data.outboundLegs.push(leg);
      if (ksaAirports.indexOf(leg.toCity) !== -1) {
        foundArrival = true;
      }
    } else {
      data.returnLegs.push(leg);
    }
  }
}


/**
 * حساب تاريخ الوصول والمغادرة من/إلى المملكة
 */
function calculateKSADates_(data) {
  var ksaAirports = ['JED', 'MED', 'DMM', 'RUH'];

  if (data.outboundLegs.length > 0) {
    var lastOutbound = data.outboundLegs[data.outboundLegs.length - 1];
    if (ksaAirports.indexOf(lastOutbound.toCity) !== -1) {
      data.arrivalCity = lastOutbound.toCity;
      data.arrivalDate = lastOutbound.arrDate || lastOutbound.depDate;
      data.arrivalTime = lastOutbound.arrTime || '';
    }
  }

  if (data.returnLegs.length > 0) {
    var firstReturn = data.returnLegs[0];
    if (ksaAirports.indexOf(firstReturn.fromCity) !== -1) {
      data.departureCity = firstReturn.fromCity;
      data.departureDate = firstReturn.depDate;
      data.departureTime = firstReturn.depTime || '';
    }
  }
}


/**
 * حذف مجلد "PDFs بدون لقب"
 */
/**
 * فحص شيت المقارنة — البحث عن مسارات غير منطقية
 */
function auditRoutes() {
  var ssId = PropertiesService.getScriptProperties().getProperty('COMPARISON_SS_ID');
  if (!ssId) return { error: 'No sheet' };
  var sheet = SpreadsheetApp.openById(ssId).getSheetByName(CONFIG.COMPARISON_SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'Empty' };

  var data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
  var issues = [];

  // أعمدة المسارات: N(14), Q(17), T(20), W(23)
  var routeCols = [13, 16, 19, 22]; // 0-indexed
  var routeNames = ['ذهاب 1', 'ذهاب 2', 'عودة 1', 'عودة 2'];

  for (var i = 0; i < data.length; i++) {
    for (var c = 0; c < routeCols.length; c++) {
      var route = String(data[i][routeCols[c]]);
      if (!route || route === '') continue;

      var parts = route.split(' إلى ');
      if (parts.length === 2 && parts[0].trim() === parts[1].trim()) {
        issues.push({
          row: i + 2,
          pnr: String(data[i][1]),
          name: String(data[i][3]) + ' ' + String(data[i][4]),
          route: route,
          type: routeNames[c],
          problem: 'نفس المطار'
        });
      }
    }
  }

  return { total: data.length, issues: issues, issueCount: issues.length };
}


function deleteNoTitleFolder() {
  var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var folders = parentFolder.getFoldersByName('PDFs بدون لقب');
  if (folders.hasNext()) {
    var f = folders.next();
    f.setTrashed(true);
    return { deleted: true };
  }
  return { deleted: false, msg: 'not found' };
}

/**
 * نقل الـ PDFs التي لا تحتوي على لقب (Mr/Mrs) إلى مجلد مستقل — دفعات
 */
function separateNoTitlePDFs() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('CHANGES_SS_ID');
  if (!ssId) return { error: 'No changes sheet' };

  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(CONFIG.CHANGES_SHEET_NAME);
  if (!sheet) return { error: 'Sheet not found' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'No data' };

  // تتبع أين وصلنا
  var startIdx = parseInt(props.getProperty('SEPARATE_IDX') || '0');

  // نمط بلقب: "Mr/Mrs NAME (Adult)"
  var withTitlePattern = /(?:Mr|Mrs|Ms|Miss|Mstr|Master)\s+[A-Z][A-Z\s]{2,40}?\s*\((?:Adult|Child|Infant)\)/;
  // نمط بدون لقب: "NAME (Adult)" مباشرة
  var noTitlePattern = /\n[A-Z][A-Z\s]{2,40}?\s*\((?:Adult|Child|Infant)\)/;

  // إنشاء مجلد فرعي
  var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var folderName = 'PDFs بدون لقب';
  var folders = parentFolder.getFoldersByName(folderName);
  var targetFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);

  var data = sheet.getRange(2, 1, lastRow - 1, 41).getValues();
  var moved = 0;
  var withTitle = 0;
  var errors = 0;
  var startTime = new Date().getTime();

  for (var i = startIdx; i < data.length; i++) {
    // حد 5 دقائق
    if (new Date().getTime() - startTime > 5 * 60 * 1000) {
      props.setProperty('SEPARATE_IDX', String(i));
      return { status: 'partial', processed: i - startIdx, moved: moved, withTitle: withTitle, remaining: data.length - i, folderUrl: targetFolder.getUrl() };
    }

    var pdfLink = String(data[i][37]); // col AL
    var m = pdfLink.match(/\/d\/([^\/]+)/);
    if (!m) continue;

    var fileId = m[1];
    try {
      var text = extractTextFromPDF_(fileId);
      var hasTitle = withTitlePattern.test(text);
      var hasNoTitle = noTitlePattern.test(text);

      // PDF بدون لقب = فيه اسم بدون Mr/Mrs قبل (Adult) وما فيه نسخة بلقب
      if (hasNoTitle && !hasTitle) {
        var file = DriveApp.getFileById(fileId);
        targetFolder.addFile(file);
        moved++;
      } else {
        withTitle++;
      }
    } catch (e) {
      errors++;
    }
  }

  // انتهينا — مسح المؤشر
  props.deleteProperty('SEPARATE_IDX');
  return { status: 'done', total: data.length, moved: moved, withTitle: withTitle, errors: errors, folderUrl: targetFolder.getUrl() };
}


/**
 * تشخيص PDF من شيت المقارنة — يعرض بالضبط ما يلتقطه كل regex
 */
function debugPDFByRow(rowNum) {
  rowNum = parseInt(rowNum) || 2;
  var ssId = PropertiesService.getScriptProperties().getProperty('COMPARISON_SS_ID');
  if (!ssId) return { error: 'No comparison sheet' };

  var sheet = SpreadsheetApp.openById(ssId).getSheetByName(CONFIG.COMPARISON_SHEET_NAME);
  if (!sheet) return { error: 'Sheet not found' };

  var row = sheet.getRange(rowNum, 1, 1, 26).getValues()[0];
  var pdfLink = String(row[24]); // col Y
  var m = pdfLink.match(/\/d\/([^\/]+)/);
  if (!m) return { error: 'No PDF link in row ' + rowNum };

  var text = extractTextFromPDF_(m[1]);

  // التقاط dateTimePairs — نفس regex الكود
  var dateTimePattern = /([A-Z][a-z]+\s+\d{1,2},?\s+\d{4})\s*\n?\s*(\d{1,2}:\d{2}\s*[AP]M)\s*\n?\s*([A-Z]{3})/g;
  var pairs = [];
  var match;
  while ((match = dateTimePattern.exec(text)) !== null) {
    pairs.push({
      idx: match.index,
      date: match[1],
      time: match[2],
      airport: match[3],
      context: text.substring(Math.max(0, match.index - 80), Math.min(text.length, match.index + match[0].length + 20)).replace(/\n/g, '\\n')
    });
  }

  // التقاط أرقام الرحلات
  var fnPattern = /([A-Z]{2}[-\s]?\d{1,5})\s*\d+h\s*\d+m/g;
  var flightNums = [];
  while ((match = fnPattern.exec(text)) !== null) {
    flightNums.push({ fn: match[1], idx: match.index });
  }

  // النتائج كما يبنيها الكود الحالي
  var legs = extractFlightLegs_(text);

  return {
    pnr: String(row[1]),
    bookingId: String(row[2]),
    textLength: text.length,
    dateTimePairsCount: pairs.length,
    dateTimePairs: pairs,
    flightNumbers: flightNums,
    legsBuilt: legs,
    rawText: text
  };
}


/**
 * اختبار تحليل PDF
 */
function testParseNusukPDF(fileId) {
  if (!fileId) {
    Logger.log('استخدم: testParseNusukPDF("FILE_ID")');
    return;
  }

  var text = extractTextFromPDF_(fileId);
  Logger.log('=== النص المستخرج (أول 3000 حرف) ===');
  Logger.log(text.substring(0, 3000));
  Logger.log('=== نهاية النص ===\n');

  var data = parseNusukTicket_(text);
  if (data) {
    Logger.log('Booking ID: ' + data.bookingId);
    Logger.log('PNR: ' + data.pnr);
    Logger.log('اسم الحاج: ' + data.passengerName);
    Logger.log('الاسم الأول: ' + data.firstName);
    Logger.log('العائلة: ' + data.lastName);
    Logger.log('رحلات الذهاب: ' + data.outboundLegs.length);
    data.outboundLegs.forEach(function(leg, i) {
      Logger.log('  ذهاب ' + (i + 1) + ': ' + leg.flightNumber + ' ' + leg.fromCity + '→' + leg.toCity + ' ' + leg.depDate + ' ' + leg.depTime);
    });
    Logger.log('رحلات العودة: ' + data.returnLegs.length);
    data.returnLegs.forEach(function(leg, i) {
      Logger.log('  عودة ' + (i + 1) + ': ' + leg.flightNumber + ' ' + leg.fromCity + '→' + leg.toCity + ' ' + leg.depDate + ' ' + leg.depTime);
    });
    Logger.log('وصول المملكة: ' + data.arrivalCity + ' ' + data.arrivalDate + ' ' + data.arrivalTime);
    Logger.log('مغادرة المملكة: ' + data.departureCity + ' ' + data.departureDate + ' ' + data.departureTime);
  } else {
    Logger.log('❌ لم يتم التعرف على التذكرة');
  }

  return data;
}
