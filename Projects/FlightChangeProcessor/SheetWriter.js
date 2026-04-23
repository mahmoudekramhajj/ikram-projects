/**
 * SheetWriter.js — كتابة نتائج التغييرات في Spreadsheet منفصل
 * النسخة 2: مبسّط + مقارنة القديم والجديد
 */


/**
 * مسح الـ Spreadsheet القديم وإنشاء جديد
 */
function resetChangesSheet() {
  PropertiesService.getScriptProperties().deleteProperty('CHANGES_SS_ID');
  Logger.log('🗑️ تم مسح ID القديم — سيُنشأ Spreadsheet جديد عند التشغيل');
}


/**
 * الحصول على أو إنشاء Spreadsheet التغييرات
 */
function getOrCreateChangesSheet_() {
  var ss;

  var ssId = CONFIG.CHANGES_SPREADSHEET_ID;
  if (!ssId) {
    ssId = PropertiesService.getScriptProperties().getProperty('CHANGES_SS_ID') || '';
  }

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      Logger.log('⚠️ لم يُعثر على Spreadsheet — سيتم إنشاؤه');
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('تغييرات الطيران V2 — إكرام الضيف');

    var file = DriveApp.getFileById(ss.getId());
    var folder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    Logger.log('✅ Spreadsheet جديد: ' + ss.getId());
    PropertiesService.getScriptProperties().setProperty('CHANGES_SS_ID', ss.getId());
  }

  var sheet = ss.getSheetByName(CONFIG.CHANGES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.CHANGES_SHEET_NAME);
    createHeaders_(sheet);
    Logger.log('✅ تم إنشاء شيت: ' + CONFIG.CHANGES_SHEET_NAME);
  }

  return sheet;
}


/**
 * الهيدر الجديد — مبسّط مع مقارنة
 */
function createHeaders_(sheet) {
  var headers = [
    // هوية التغيير
    'رقم التغيير',         // A
    'PNR',                 // B
    'Booking ID',          // C
    'Incident #',          // D
    // بيانات الحاج من PDF
    'الاسم الأول (PDF)',    // E
    'اسم العائلة (PDF)',    // F
    // بيانات الحاج من النظام
    'الاسم الأول (النظام)', // G
    'اسم العائلة (النظام)', // H
    'الرقم التسلسلي',      // I
    'رقم الجواز',          // J
    'رقم الباقة',          // K
    'اسم الباقة',          // L
    'نوع العقد',           // M
    'نوع الحجز',           // (B2B/B2C)
    // الرحلة الجديدة (من PDF)
    'جديد: ذهاب رحلة 1',   // N
    'جديد: ذهاب المسار 1', // O
    'جديد: ذهاب تاريخ 1',  // P
    'جديد: ذهاب وقت 1',    // Q
    'جديد: ذهاب رحلة 2',   // R
    'جديد: ذهاب المسار 2', // S
    'جديد: ذهاب تاريخ 2',  // T
    'جديد: ذهاب وقت 2',    // U
    'جديد: عودة رحلة 1',   // V
    'جديد: عودة المسار 1', // W
    'جديد: عودة تاريخ 1',  // X
    'جديد: عودة وقت 1',    // Y
    'جديد: عودة رحلة 2',   // Z
    'جديد: عودة المسار 2', // AA
    'جديد: عودة تاريخ 2',  // AB
    'جديد: عودة وقت 2',    // AC
    // الرحلة الحالية (من شيت الطيران/B2C)
    'حالي: ذهاب رحلة 1',   // AD
    'حالي: ذهاب تاريخ 1',  // AE
    'حالي: ذهاب رحلة 2',   // AF
    'حالي: ذهاب تاريخ 2',  // AG
    'حالي: عودة رحلة 1',   // AH
    'حالي: عودة تاريخ 1',  // AI
    'حالي: عودة رحلة 2',   // AJ
    'حالي: عودة تاريخ 2',  // AK
    // معلومات إضافية
    'رابط PDF',            // AL
    'تاريخ الإيميل',       // AM
    'حالة المطابقة',       // AN
    'ملاحظات'              // AO
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // تلوين الأقسام
  sheet.getRange(1, 1, 1, 4).setBackground('#4a86c8').setFontColor('#fff');   // هوية
  sheet.getRange(1, 5, 1, 2).setBackground('#f0ad4e').setFontColor('#fff');   // PDF اسم
  sheet.getRange(1, 7, 1, 8).setBackground('#5cb85c').setFontColor('#fff');   // نظام + نوع الحجز
  sheet.getRange(1, 15, 1, 16).setBackground('#476831').setFontColor('#fff'); // جديد
  sheet.getRange(1, 31, 1, 8).setBackground('#232E64').setFontColor('#fff');  // حالي
  sheet.getRange(1, 39, 1, 4).setBackground('#777').setFontColor('#fff');     // إضافي
}


/**
 * كتابة صف تغيير — النسخة المبسّطة
 */
function writeChangeRow_(sheet, data) {
  var changeNum = data.changeNum || ('CHG-' + String(sheet.getLastRow()).padStart(4, '0'));

  var ob1 = data.outboundLegs && data.outboundLegs[0] ? data.outboundLegs[0] : {};
  var ob2 = data.outboundLegs && data.outboundLegs[1] ? data.outboundLegs[1] : {};
  var rt1 = data.returnLegs && data.returnLegs[0] ? data.returnLegs[0] : {};
  var rt2 = data.returnLegs && data.returnLegs[1] ? data.returnLegs[1] : {};

  var row = [
    changeNum,
    data.pnr || '',
    data.bookingId || '',
    data.incidentNum || '',
    // اسم من PDF
    data.pdfFirstName || data.firstName || '',
    data.pdfLastName || data.lastName || '',
    // اسم من النظام
    data.sysFirstName || '',
    data.sysLastName || '',
    data.serialNum || '',
    data.passport || '',
    data.pkgNum || '',
    data.pkgName || '',
    data.flightType || '',
    data.bookingType || '',
    // الرحلة الجديدة (من PDF)
    ob1.flightNumber || '',
    ob1.fromCity && ob1.toCity ? 'من ' + ob1.fromCity + ' إلى ' + ob1.toCity : '',
    ob1.depDate || '',
    ob1.depTime || '',
    ob2.flightNumber || '',
    ob2.fromCity && ob2.toCity ? 'من ' + ob2.fromCity + ' إلى ' + ob2.toCity : '',
    ob2.depDate || '',
    ob2.depTime || '',
    rt1.flightNumber || '',
    rt1.fromCity && rt1.toCity ? 'من ' + rt1.fromCity + ' إلى ' + rt1.toCity : '',
    rt1.depDate || '',
    rt1.depTime || '',
    rt2.flightNumber || '',
    rt2.fromCity && rt2.toCity ? 'من ' + rt2.fromCity + ' إلى ' + rt2.toCity : '',
    rt2.depDate || '',
    rt2.depTime || '',
    // الرحلة الحالية (من شيت الطيران/B2C)
    data.curOutFlight1 || '',
    data.curOutDate1 || '',
    data.curOutFlight2 || '',
    data.curOutDate2 || '',
    data.curRetFlight1 || '',
    data.curRetDate1 || '',
    data.curRetFlight2 || '',
    data.curRetDate2 || '',
    // إضافي
    data.pdfLink || '',
    data.emailDate || '',
    data.status || '',
    data.notes || ''
  ];

  // كتابة مباشرة (appendRow يفشل بصمت إذا كان هناك Filter نشط)
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  // تلوين حسب الحالة
  var statusCol = 41; // AO (بعد إضافة عمود نوع الحجز)
  var statusCell = sheet.getRange(newRow, statusCol);

  if (data.status === 'تم المطابقة') {
    statusCell.setBackground('#d4edda');
  } else if (data.status === 'لم يُطابَق') {
    statusCell.setBackground('#fff3cd');
  } else if (data.status === 'خطأ تحليل') {
    statusCell.setBackground('#f8d7da');
  }
}



/**
 * إحصائيات
 */
function getChangesStats() {
  var sheet = getOrCreateChangesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0 };

  var statuses = sheet.getRange(2, 41, lastRow - 1, 1).getValues();
  var stats = { total: lastRow - 1, matched: 0, unmatched: 0, errors: 0 };

  statuses.forEach(function(row) {
    var s = String(row[0]);
    if (s === 'تم المطابقة') stats.matched++;
    else if (s === 'لم يُطابَق') stats.unmatched++;
    else stats.errors++;
  });

  Logger.log('📊 الإجمالي: ' + stats.total + ' | ✅ مطابق: ' + stats.matched + ' | ⚠️ غير مطابق: ' + stats.unmatched);
  return stats;
}


// =====================================================
// === شيت المقارنة — Spreadsheet منفصل ===
// =====================================================


/**
 * عرض روابط الشيتات الحالية
 */
function getSheetUrls() {
  var changesId = PropertiesService.getScriptProperties().getProperty('CHANGES_SS_ID') || '';
  var comparisonId = PropertiesService.getScriptProperties().getProperty('COMPARISON_SS_ID') || '';
  return {
    changes: changesId ? 'https://docs.google.com/spreadsheets/d/' + changesId : '(غير موجود)',
    comparison: comparisonId ? 'https://docs.google.com/spreadsheets/d/' + comparisonId : '(غير موجود)'
  };
}


/**
 * حذف كل الـ Spreadsheets القديمة المشابهة (ما عدا النشطة حالياً)
 */
function deleteOldSimilarSheets() {
  var activeChangesId = PropertiesService.getScriptProperties().getProperty('CHANGES_SS_ID') || '';
  var activeComparisonId = PropertiesService.getScriptProperties().getProperty('COMPARISON_SS_ID') || '';

  var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var files = parentFolder.getFiles();
  var deleted = [];
  var kept = [];

  var patterns = [
    /تغييرات الطيران/,
    /مقارنة أسماء الحجاج/,
    /Flight.*Changes/i,
    /Pilgrim.*Comparison/i
  ];

  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName();
    var id = f.getId();
    var mime = f.getMimeType();

    if (mime !== 'application/vnd.google-apps.spreadsheet') continue;

    var matches = patterns.some(function(p) { return p.test(name); });
    if (!matches) continue;

    if (id === activeChangesId || id === activeComparisonId) {
      kept.push(name + ' (نشط)');
      continue;
    }

    f.setTrashed(true);
    deleted.push(name);
  }

  Logger.log('🗑️ حُذف: ' + deleted.length + ' ملف');
  Logger.log('✅ محفوظ: ' + kept.length + ' ملف نشط');
  return { deleted: deleted, kept: kept };
}


/**
 * مسح الـ Spreadsheet القديم للمقارنة
 */
function resetComparisonSheet() {
  PropertiesService.getScriptProperties().deleteProperty('COMPARISON_SS_ID');
  Logger.log('🗑️ تم مسح ID شيت المقارنة — سيُنشأ جديد عند التشغيل');
}


/**
 * الحصول على أو إنشاء Spreadsheet المقارنة
 */
function getOrCreateComparisonSheet_() {
  var ss;

  var ssId = CONFIG.COMPARISON_SPREADSHEET_ID;
  if (!ssId) {
    ssId = PropertiesService.getScriptProperties().getProperty('COMPARISON_SS_ID') || '';
  }

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      Logger.log('⚠️ لم يُعثر على Spreadsheet المقارنة — سيتم إنشاؤه');
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('مقارنة أسماء الحجاج — إكرام الضيف');

    var file = DriveApp.getFileById(ss.getId());
    var folder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    Logger.log('✅ Spreadsheet مقارنة جديد: ' + ss.getId());
    PropertiesService.getScriptProperties().setProperty('COMPARISON_SS_ID', ss.getId());
  }

  var sheet = ss.getSheetByName(CONFIG.COMPARISON_SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.COMPARISON_SHEET_NAME);
    createComparisonHeaders_(sheet);
    Logger.log('✅ تم إنشاء شيت: ' + CONFIG.COMPARISON_SHEET_NAME);
  }

  return sheet;
}


/**
 * هيدر شيت المقارنة
 */
function createComparisonHeaders_(sheet) {
  var headers = [
    '#',                      // A
    'PNR',                    // B
    'Booking ID',             // C
    // أسماء من PDF
    'الاسم الأول (PDF)',       // D
    'اسم العائلة (PDF)',       // E
    // أسماء من النظام
    'الاسم الأول (النظام)',    // F
    'اسم العائلة (النظام)',    // G
    // المقارنة
    'حالة المطابقة',          // H
    // بيانات الحاج
    'الرقم التسلسلي',         // I
    'رقم الجواز',             // J
    'رقم الباقة',             // K
    'اسم الباقة',             // L
    'نوع الحجز',              // M
    // رحلات الذهاب (من PDF)
    'ذهاب: رحلة 1',           // N
    'ذهاب: المسار 1',         // N
    'ذهاب: تاريخ 1',          // O
    'ذهاب: رحلة 2',           // P
    'ذهاب: المسار 2',         // Q
    'ذهاب: تاريخ 2',          // R
    // رحلات العودة (من PDF)
    'عودة: رحلة 1',           // S
    'عودة: المسار 1',         // T
    'عودة: تاريخ 1',          // U
    'عودة: رحلة 2',           // V
    'عودة: المسار 2',         // W
    'عودة: تاريخ 2',          // X
    // إضافي
    'رابط PDF',               // Y
    'تاريخ الإيميل'           // Z
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // تلوين الأقسام
  sheet.getRange(1, 1, 1, 3).setBackground('#4a86c8').setFontColor('#fff');    // هوية
  sheet.getRange(1, 4, 1, 2).setBackground('#f0ad4e').setFontColor('#fff');    // PDF اسم
  sheet.getRange(1, 6, 1, 2).setBackground('#5cb85c').setFontColor('#fff');    // نظام اسم
  sheet.getRange(1, 8, 1, 1).setBackground('#9B6FB1').setFontColor('#fff');    // مطابقة
  sheet.getRange(1, 9, 1, 5).setBackground('#476831').setFontColor('#fff');    // بيانات الحاج + نوع الحجز
  sheet.getRange(1, 14, 1, 6).setBackground('#232E64').setFontColor('#fff');   // ذهاب
  sheet.getRange(1, 20, 1, 6).setBackground('#1a2350').setFontColor('#fff');   // عودة
  sheet.getRange(1, 26, 1, 2).setBackground('#777').setFontColor('#fff');      // إضافي
}


/**
 * كتابة صف في شيت المقارنة + تلوين الصف كامل
 */
function writeComparisonRow_(sheet, data) {
  var rowNum = sheet.getLastRow();

  var ob1 = data.outboundLegs && data.outboundLegs[0] ? data.outboundLegs[0] : {};
  var ob2 = data.outboundLegs && data.outboundLegs[1] ? data.outboundLegs[1] : {};
  var rt1 = data.returnLegs && data.returnLegs[0] ? data.returnLegs[0] : {};
  var rt2 = data.returnLegs && data.returnLegs[1] ? data.returnLegs[1] : {};

  var row = [
    rowNum,                                                                    // #
    data.pnr || '',                                                            // PNR
    data.bookingId || '',                                                      // Booking ID
    // أسماء PDF
    data.pdfFirstName || '',                                                   // الاسم الأول (PDF)
    data.pdfLastName || '',                                                    // اسم العائلة (PDF)
    // أسماء النظام
    data.sysFirstName || '',                                                   // الاسم الأول (النظام)
    data.sysLastName || '',                                                    // اسم العائلة (النظام)
    // حالة المطابقة
    data.comparisonStatus || '',                                               // حالة المطابقة
    // بيانات الحاج
    data.serialNum || '',                                                      // الرقم التسلسلي
    data.passport || '',                                                       // رقم الجواز
    data.pkgNum || '',                                                         // رقم الباقة
    data.pkgName || '',                                                        // اسم الباقة
    data.bookingType || '',                                                    // نوع الحجز
    // ذهاب
    ob1.flightNumber || '',                                                    // ذهاب: رحلة 1
    ob1.fromCity && ob1.toCity ? 'من ' + ob1.fromCity + ' إلى ' + ob1.toCity : '',         // ذهاب: المسار 1
    ob1.depDate || '',                                                         // ذهاب: تاريخ 1
    ob2.flightNumber || '',                                                    // ذهاب: رحلة 2
    ob2.fromCity && ob2.toCity ? 'من ' + ob2.fromCity + ' إلى ' + ob2.toCity : '',         // ذهاب: المسار 2
    ob2.depDate || '',                                                         // ذهاب: تاريخ 2
    // عودة
    rt1.flightNumber || '',                                                    // عودة: رحلة 1
    rt1.fromCity && rt1.toCity ? 'من ' + rt1.fromCity + ' إلى ' + rt1.toCity : '',         // عودة: المسار 1
    rt1.depDate || '',                                                         // عودة: تاريخ 1
    rt2.flightNumber || '',                                                    // عودة: رحلة 2
    rt2.fromCity && rt2.toCity ? 'من ' + rt2.fromCity + ' إلى ' + rt2.toCity : '',         // عودة: المسار 2
    rt2.depDate || '',                                                         // عودة: تاريخ 2
    // إضافي
    data.pdfLink || '',                                                        // رابط PDF
    data.emailDate || ''                                                       // تاريخ الإيميل
  ];

  // كتابة مباشرة (appendRow يفشل بصمت إذا كان هناك Filter نشط)
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  // تلوين الصف كامل حسب الحالة
  var numCols = row.length;
  var rowRange = sheet.getRange(newRow, 1, 1, numCols);

  if (data.comparisonStatus === 'متطابق') {
    rowRange.setBackground('#d4edda');
  } else if (data.comparisonStatus === 'غير متطابق') {
    rowRange.setBackground('#fff3cd');
  } else if (data.comparisonStatus === 'لا يوجد اسم') {
    rowRange.setBackground('#f8d7da');
  }
}


/**
 * قراءة بيانات الطيران من شيت التغييرات للتحليل
 */
/**
 * البحث عن الصفوف غير المتطابقة
 */
function findUnmatched() {
  var sheet = getOrCreateComparisonSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();
  var results = [];

  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][7]); // H = حالة المطابقة
    if (status !== 'متطابق' && status !== '') {
      results.push({
        row: i + 2,
        pnr: String(data[i][1]),
        bookingId: String(data[i][2]),
        pdfFirst: String(data[i][3]),
        pdfLast: String(data[i][4]),
        sysFirst: String(data[i][5]),
        sysLast: String(data[i][6]),
        status: status
      });
    }
  }

  return results;
}


function readFlightData() {
  var sheet = getOrCreateChangesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'No data' };

  var data = sheet.getRange(2, 1, lastRow - 1, 41).getValues();
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    rows.push({
      row: i + 2,
      pnr: String(data[i][1]),           // B
      bookingId: String(data[i][2]),      // C
      pdfFirst: String(data[i][4]),       // E
      pdfLast: String(data[i][5]),        // F
      sysFirst: String(data[i][6]),       // G
      sysLast: String(data[i][7]),        // H
      serial: String(data[i][8]),         // I
      // رحلات جديدة (من PDF)
      newOb1Flight: String(data[i][13]),  // N
      newOb1Route: String(data[i][14]),   // O
      newOb1Date: String(data[i][15]),    // P
      newOb1Time: String(data[i][16]),    // Q
      newOb2Flight: String(data[i][17]),  // R
      newOb2Route: String(data[i][18]),   // S
      newOb2Date: String(data[i][19]),    // T
      newOb2Time: String(data[i][20]),    // U
      newRt1Flight: String(data[i][21]),  // V
      newRt1Route: String(data[i][22]),   // W
      newRt1Date: String(data[i][23]),    // X
      newRt1Time: String(data[i][24]),    // Y
      newRt2Flight: String(data[i][25]),  // Z
      newRt2Route: String(data[i][26]),   // AA
      newRt2Date: String(data[i][27]),    // AB
      newRt2Time: String(data[i][28]),    // AC
      // رحلات حالية (من النظام)
      curOb1Flight: String(data[i][29]), // AD
      curOb1Date: String(data[i][30]),   // AE
      curOb2Flight: String(data[i][31]), // AF
      curOb2Date: String(data[i][32]),   // AG
      curRt1Flight: String(data[i][33]), // AH
      curRt1Date: String(data[i][34]),   // AI
      curRt2Flight: String(data[i][35]), // AJ
      curRt2Date: String(data[i][36]),   // AK
      status: String(data[i][39]),       // AN
    });
  }

  return { total: rows.length, rows: rows };
}


/**
 * إرجاع روابط الشيتات
 */
function getSheetLinks() {
  var props = PropertiesService.getScriptProperties();
  var changesId = props.getProperty('CHANGES_SS_ID') || '';
  var compId = props.getProperty('COMPARISON_SS_ID') || '';
  return {
    changes: changesId ? 'https://docs.google.com/spreadsheets/d/' + changesId : '',
    comparison: compId ? 'https://docs.google.com/spreadsheets/d/' + compId : ''
  };
}


/**
 * تدقيق شامل — يفحص شيت المقارنة والبيانات المصدرية
 */
function fullAudit() {
  var ssId = PropertiesService.getScriptProperties().getProperty('COMPARISON_SS_ID');
  if (!ssId) return { error: 'No comparison sheet' };

  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(CONFIG.COMPARISON_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { error: 'No data' };

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();

  // تحميل بيانات PD للمقارنة
  var mainSS = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  var pdSheet = mainSS.getSheetByName(CONFIG.PD_SHEET);
  var pdData = pdSheet && pdSheet.getLastRow() > 1
    ? pdSheet.getRange(2, 1, pdSheet.getLastRow() - 1, 26).getValues()
    : [];

  var issues = [];

  // === 1. أسماء فارغة ===
  var emptyNames = [];
  for (var i = 0; i < data.length; i++) {
    var pdfFirst = String(data[i][3]).trim();
    var pdfLast = String(data[i][4]).trim();
    if (!pdfFirst && !pdfLast) {
      emptyNames.push({ row: i + 2, pnr: String(data[i][1]) });
    }
  }
  if (emptyNames.length > 0) issues.push({ type: 'أسماء PDF فارغة', count: emptyNames.length, samples: emptyNames.slice(0, 5) });

  // === 2. متطابق لكن اسم النظام فارغ ===
  var emptySysNames = [];
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][7]);
    var sysFirst = String(data[i][5]).trim();
    var sysLast = String(data[i][6]).trim();
    if (status === 'متطابق' && !sysFirst && !sysLast) {
      emptySysNames.push({ row: i + 2, pnr: String(data[i][1]), pdfName: String(data[i][3]) + ' ' + String(data[i][4]) });
    }
  }
  if (emptySysNames.length > 0) issues.push({ type: 'متطابق لكن اسم النظام فارغ', count: emptySysNames.length, samples: emptySysNames.slice(0, 5) });

  // === 3. صفوف مكررة (نفس PNR + اسم) ===
  var seen = {};
  var duplicates = [];
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][1]).trim() + '|' + String(data[i][3]).trim().toUpperCase() + '|' + String(data[i][4]).trim().toUpperCase();
    if (seen[key] !== undefined) {
      duplicates.push({ row: i + 2, firstRow: seen[key], pnr: String(data[i][1]), name: String(data[i][3]) + ' ' + String(data[i][4]) });
    } else {
      seen[key] = i + 2;
    }
  }
  if (duplicates.length > 0) issues.push({ type: 'صفوف مكررة (نفس PNR + اسم)', count: duplicates.length, samples: duplicates.slice(0, 5) });

  // === 4. رحلات ذهاب فارغة ===
  var noOutbound = [];
  for (var i = 0; i < data.length; i++) {
    var ob1 = String(data[i][12]).trim();
    if (!ob1) noOutbound.push({ row: i + 2, pnr: String(data[i][1]), name: String(data[i][3]) + ' ' + String(data[i][4]) });
  }
  if (noOutbound.length > 0) issues.push({ type: 'رحلة ذهاب فارغة', count: noOutbound.length, samples: noOutbound.slice(0, 5) });

  // === 5. رحلات عودة فارغة ===
  var noReturn = [];
  for (var i = 0; i < data.length; i++) {
    var rt1 = String(data[i][18]).trim();
    if (!rt1) noReturn.push({ row: i + 2, pnr: String(data[i][1]), name: String(data[i][3]) + ' ' + String(data[i][4]) });
  }
  if (noReturn.length > 0) issues.push({ type: 'رحلة عودة فارغة', count: noReturn.length, samples: noReturn.slice(0, 5) });

  // === 6. اسم PDF مختلف عن اسم النظام ===
  var nameMismatch = [];
  for (var i = 0; i < data.length; i++) {
    var status = String(data[i][7]);
    if (status !== 'متطابق') continue;
    var pdfFull = (String(data[i][3]) + String(data[i][4])).toUpperCase().replace(/[\s\-']/g, '');
    var sysFull = (String(data[i][5]) + String(data[i][6])).toUpperCase().replace(/[\s\-']/g, '');
    if (pdfFull && sysFull && pdfFull !== sysFull) {
      var sysReverse = (String(data[i][6]) + String(data[i][5])).toUpperCase().replace(/[\s\-']/g, '');
      if (pdfFull !== sysReverse) {
        nameMismatch.push({
          row: i + 2, pnr: String(data[i][1]),
          pdfName: String(data[i][3]) + ' ' + String(data[i][4]),
          sysName: String(data[i][5]) + ' ' + String(data[i][6])
        });
      }
    }
  }
  if (nameMismatch.length > 0) issues.push({ type: 'اسم PDF مختلف عن اسم النظام', count: nameMismatch.length, samples: nameMismatch.slice(0, 10) });

  // === 7. Booking ID فارغ ===
  var noBooking = [];
  for (var i = 0; i < data.length; i++) {
    if (!String(data[i][2]).trim()) noBooking.push({ row: i + 2, pnr: String(data[i][1]) });
  }
  if (noBooking.length > 0) issues.push({ type: 'Booking ID فارغ', count: noBooking.length, samples: noBooking.slice(0, 5) });

  // === 8. مسافرين ناقصين — PNR في PD لكن ما ظهر بالشيت ===
  var pnrSheetCount = {};
  for (var i = 0; i < data.length; i++) {
    var pnr = String(data[i][1]).trim().toUpperCase();
    if (pnr) pnrSheetCount[pnr] = (pnrSheetCount[pnr] || 0) + 1;
  }

  var pnrPdCount = {};
  for (var i = 0; i < pdData.length; i++) {
    var contract = String(pdData[i][CONFIG.PD.CONTRACT_NAME] || '').toUpperCase();
    for (var pnr in pnrSheetCount) {
      if (contract.indexOf(pnr) !== -1) {
        pnrPdCount[pnr] = (pnrPdCount[pnr] || 0) + 1;
      }
    }
  }

  var missingPax = [];
  for (var pnr in pnrSheetCount) {
    if (pnrPdCount[pnr] && pnrPdCount[pnr] > pnrSheetCount[pnr]) {
      missingPax.push({ pnr: pnr, inSheet: pnrSheetCount[pnr], inPD: pnrPdCount[pnr], missing: pnrPdCount[pnr] - pnrSheetCount[pnr] });
    }
  }
  if (missingPax.length > 0) issues.push({ type: 'مسافرين ناقصين (PD أكثر من الشيت)', count: missingPax.length, samples: missingPax.slice(0, 10) });

  // === 9. رقم تسلسلي فارغ رغم المطابقة ===
  var noSerial = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][7]) === 'متطابق' && !String(data[i][8]).trim()) {
      noSerial.push({ row: i + 2, pnr: String(data[i][1]), name: String(data[i][3]) + ' ' + String(data[i][4]) });
    }
  }
  if (noSerial.length > 0) issues.push({ type: 'متطابق لكن بدون رقم تسلسلي', count: noSerial.length, samples: noSerial.slice(0, 5) });

  return {
    totalRows: data.length,
    uniquePilgrims: Object.keys(seen).length,
    totalPD: pdData.length,
    issuesFound: issues.length,
    issues: issues
  };
}


/**
 * إحصائيات شيت المقارنة
 */
function getComparisonStats() {
  var sheet = getOrCreateComparisonSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0 };

  var statuses = sheet.getRange(2, 8, lastRow - 1, 1).getValues(); // عمود H = حالة المطابقة
  var stats = { total: lastRow - 1, matched: 0, unmatched: 0, noName: 0 };

  statuses.forEach(function(row) {
    var s = String(row[0]);
    if (s === 'متطابق') stats.matched++;
    else if (s === 'غير متطابق') stats.unmatched++;
    else stats.noName++;
  });

  Logger.log('📊 المقارنة — الإجمالي: ' + stats.total + ' | ✅ متطابق: ' + stats.matched + ' | ⚠️ غير متطابق: ' + stats.unmatched + ' | ❌ لا اسم: ' + stats.noName);
  return stats;
}


// =====================================================
// === V3: شيت موحّد (دمج التغييرات + المقارنة) ========
// =====================================================

/**
 * إنشاء أو الحصول على شيت V3 الموحّد
 */
function getOrCreateUnifiedSheet_() {
  var ss;
  var ssId = (CONFIG.UNIFIED_SPREADSHEET_ID || '') ||
             PropertiesService.getScriptProperties().getProperty('UNIFIED_SS_ID') || '';

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) {}
  }

  if (!ss) {
    ss = SpreadsheetApp.create('تغييرات الطيران V3 — إكرام الضيف');
    var file = DriveApp.getFileById(ss.getId());
    var folder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    PropertiesService.getScriptProperties().setProperty('UNIFIED_SS_ID', ss.getId());
    Logger.log('✅ Spreadsheet V3 جديد: ' + ss.getId());
  }

  var sheet = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.UNIFIED_SHEET_NAME);
    createUnifiedHeaders_(sheet);
    Logger.log('✅ تم إنشاء شيت: ' + CONFIG.UNIFIED_SHEET_NAME);
  }

  return sheet;
}


/**
 * هيدر V3 — 46 عموداً يدمج التغييرات + المقارنة + تحذيرات + إخطارات
 */
function createUnifiedHeaders_(sheet) {
  var headers = [
    // هوية (A-E)
    'رقم التغيير',                 // A
    'PNR',                         // B
    'Booking ID',                  // C
    'Incident #',                  // D
    'تاريخ الإيميل',               // E
    // أسماء (F-J)
    'الاسم الأول (PDF)',            // F
    'اسم العائلة (PDF)',            // G
    'الاسم الأول (النظام)',         // H
    'اسم العائلة (النظام)',         // I
    'حالة المطابقة',               // J
    // بيانات النظام (K-P)
    'الرقم التسلسلي',              // K
    'رقم الجواز',                  // L
    'رقم الباقة',                  // M
    'اسم الباقة',                  // N
    'نوع الحجز',                   // O
    'نوع الرحلة',                  // P
    // === ذهاب — مرحلة 1 قبل الترانزيت (Q-W) === 7 أعمدة
    'FlightNo1 (ذهاب)',           // Q
    'Date TAKEOFF 1',             // R
    'TAKEOFF TIME 1',             // S
    'From 1',                     // T
    'To 1',                       // U
    'DATE LANDING 1',             // V
    'LANDING TIME 1',             // W
    // === ذهاب — مرحلة 2 وصول السعودية (X-AD) === 7 أعمدة
    'FlightNo 2 (ذهاب)',          // X
    'Date TAKEOFF 2',             // Y
    'TAKEOFF TIME 2',             // Z
    'From 2',                     // AA
    'To 2',                       // AB
    'DATE LANDING 2',             // AC
    'LANDING TIME 2',             // AD
    // === عودة — مرحلة 1 مغادرة السعودية (AE-AK) === 7 أعمدة
    'FlightNo1 (عودة)',           // AE
    'Date TAKEOFF 1',             // AF
    'TAKEOFF TIME 1',             // AG
    'From 1',                     // AH
    'To 1',                       // AI
    'DATE LANDING 1',             // AJ
    'LANDING TIME 1',             // AK
    // === عودة — مرحلة 2 ترانزيت للبلد (AL-AR) === 7 أعمدة
    'FlightNo 2 (عودة)',          // AL
    'Date TAKEOFF 2',             // AM
    'TAKEOFF TIME 2',             // AN
    'From 2',                     // AO
    'To 2',                       // AP
    'DATE LANDING 2',             // AQ
    'LANDING TIME 2',             // AR
    // الرحلات الحالية (AS-AV)
    'حالي: ذهاب رحلة 1',           // AS
    'حالي: ذهاب تاريخ 1',          // AT
    'حالي: عودة رحلة 1',           // AU
    'حالي: عودة تاريخ 1',          // AV
    // إضافي (AW-AZ)
    'رابط التذكرة',                // AW
    'المصدر',                      // AX (PDF / نص)
    'ملاحظات',                     // AY
    '⚠️ تحذيرات',                   // AZ
    // إخطار البوت (BA-BB)
    'حالة إخطار الحاج',            // BA
    'تاريخ الإخطار'                // BB
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // تلوين الأقسام
  sheet.getRange(1, 1, 1, 5).setBackground('#4a86c8').setFontColor('#fff');    // هوية
  sheet.getRange(1, 6, 1, 5).setBackground('#f0ad4e').setFontColor('#fff');    // أسماء
  sheet.getRange(1, 11, 1, 6).setBackground('#5cb85c').setFontColor('#fff');   // النظام
  sheet.getRange(1, 17, 1, 14).setBackground('#476831').setFontColor('#fff');  // ذهاب (28 → actually 14 cols)
  sheet.getRange(1, 31, 1, 14).setBackground('#6c9a48').setFontColor('#fff');  // عودة
  sheet.getRange(1, 45, 1, 4).setBackground('#232E64').setFontColor('#fff');   // رحلات حالية
  sheet.getRange(1, 49, 1, 4).setBackground('#777').setFontColor('#fff');      // إضافي
  sheet.getRange(1, 53, 1, 2).setBackground('#9B6FB1').setFontColor('#fff');   // إخطار البوت
}


/**
 * كتابة صف موحّد V3 — يدمج بيانات التغيير + المطابقة + التحذيرات
 */
function writeUnifiedRow_(sheet, data) {
  var changeNum = data.changeNum || ('CHG-' + String(sheet.getLastRow()).padStart(4, '0'));

  // === قاعدة التفريغ (مطابقة شيت الطيران) ===
  // الذهاب: GO2 دائماً = رحلة وصول السعودية. GO1 = مرحلة الترانزيت قبلها (إن وُجدت).
  // العودة: RET1 دائماً = رحلة مغادرة السعودية. RET2 = مرحلة الترانزيت بعدها (إن وُجدت).
  var outbound = data.outboundLegs || [];
  var returnL  = data.returnLegs  || [];

  var ob1 = {}, ob2 = {};
  if (outbound.length === 1) {
    ob2 = outbound[0];  // مباشر → GO2 فقط
  } else if (outbound.length >= 2) {
    ob1 = outbound[0];
    ob2 = outbound[outbound.length - 1]; // آخر رحلة = وصول السعودية
  }

  var rt1 = {}, rt2 = {};
  if (returnL.length === 1) {
    rt1 = returnL[0];  // مباشر → RET1 فقط
  } else if (returnL.length >= 2) {
    rt1 = returnL[0];  // أول رحلة = مغادرة السعودية
    rt2 = returnL[returnL.length - 1];
  }

  // التحقق المنطقي — يولّد تحذيرات إن وُجدت
  var warnings = validateFlightLegs_(outbound, returnL);

  // حالة المطابقة — متطابق / غير متطابق / لا يوجد اسم
  var matchStatus = 'لا يوجد اسم';
  if (data.pdfFirstName || data.pdfLastName) {
    matchStatus = data.status === 'تم المطابقة' ? 'متطابق' : 'غير متطابق';
  }

  function legFields(leg) {
    return [
      leg.flightNumber || '',
      leg.depDate || '',
      leg.depTime || '',
      leg.fromCity || '',
      leg.toCity || '',
      leg.arrDate || leg.depDate || '',
      leg.arrTime || ''
    ];
  }

  var row = [
    // هوية (A-E)
    changeNum,
    data.pnr || '',
    data.bookingId || '',
    data.incidentNum || '',
    data.emailDate || '',
    // أسماء (F-J)
    data.pdfFirstName || '',
    data.pdfLastName || '',
    data.sysFirstName || '',
    data.sysLastName || '',
    matchStatus,
    // النظام (K-P)
    data.serialNum || '',
    data.passport || '',
    data.pkgNum || '',
    data.pkgName || '',
    data.bookingType || '',
    data.flightType || ''
  ]
  .concat(legFields(ob1))  // Q-W
  .concat(legFields(ob2))  // X-AD
  .concat(legFields(rt1))  // AE-AK
  .concat(legFields(rt2))  // AL-AR
  .concat([
    // الرحلات الحالية (AS-AV)
    data.curOutFlight1 || '',
    data.curOutDate1 || '',
    data.curRetFlight1 || '',
    data.curRetDate1 || '',
    // إضافي (AW-AZ)
    data.pdfLink || '',
    data.source || 'PDF',
    data.notes || '',
    warnings,
    // إخطار البوت (BA-BB)
    '-',
    ''
  ]);

  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  // تلوين صف التحذيرات إن وُجدت
  if (warnings) {
    sheet.getRange(newRow, 52).setBackground('#fff3cd'); // AZ = تحذيرات (index 52 in 1-based)
  }

  // تلوين حالة المطابقة
  var matchCell = sheet.getRange(newRow, 10); // J
  if (matchStatus === 'متطابق') matchCell.setBackground('#d4edda');
  else if (matchStatus === 'غير متطابق') matchCell.setBackground('#fff3cd');
}


/**
 * تحميل مفاتيح الصفوف الموجودة في V3 — للـ dedup
 * المفتاح الجديد: PNR + Incident# + firstName + lastName
 */
function loadUnifiedKeys_(sheet) {
  var keys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;

  // B=PNR(idx1), D=Incident(idx3), F=firstName(idx5), G=lastName(idx6)
  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = (
      String(data[i][1]) + '|' +  // PNR
      String(data[i][3]) + '|' +  // Incident
      String(data[i][5]) + '|' +  // firstName
      String(data[i][6])           // lastName
    ).toUpperCase();
    keys[key] = true;
  }
  return keys;
}


/**
 * إحصائيات V3
 */
function getUnifiedStats() {
  var sheet = getOrCreateUnifiedSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0 };

  var data = sheet.getRange(2, 10, lastRow - 1, 31).getValues(); // J (matchStatus) إلى AN (warnings)
  var stats = { total: lastRow - 1, matched: 0, unmatched: 0, noName: 0, withWarnings: 0 };

  data.forEach(function(row) {
    var status = String(row[0]); // J
    var warn = String(row[30]);   // AN (offset 30 from J)
    if (status === 'متطابق') stats.matched++;
    else if (status === 'غير متطابق') stats.unmatched++;
    else stats.noName++;
    if (warn) stats.withWarnings++;
  });

  Logger.log('📊 V3 — الإجمالي: ' + stats.total + ' | ✅ مطابق: ' + stats.matched + ' | ⚠️ غير مطابق: ' + stats.unmatched + ' | ❌ لا اسم: ' + stats.noName + ' | 🚨 بتحذيرات: ' + stats.withWarnings);
  return stats;
}


/**
 * رابط شيت V3
 */
function getUnifiedSheetUrl() {
  var id = PropertiesService.getScriptProperties().getProperty('UNIFIED_SS_ID') || '';
  return id ? 'https://docs.google.com/spreadsheets/d/' + id : '(غير موجود)';
}


/**
 * اختبار صف واحد: يُزيل labels من ثريد معيّن ويشغّل scanEmails
 */
function testOneThread(threadId) {
  threadId = threadId || '19db50525b7431b5';
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { error: 'thread_not_found', threadId: threadId };

  // إزالة labels المعالجة والمتخطاة
  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    var n = labels[i].getName();
    if (n === CONFIG.PROCESSED_LABEL || n === CONFIG.SKIPPED_LABEL) {
      thread.removeLabel(labels[i]);
    }
  }
  Logger.log('✅ تم إزالة labels من thread ' + threadId);

  // تشغيل scanEmails — سيعالج فقط هذا الثريد
  return scanEmails();
}


/**
 * Re-run validation on all V4 rows in place — no Gmail rescan.
 * Updates column AZ (warnings) only.
 */
function revalidateAllV3() {
  var sheet = getOrCreateUnifiedSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0 };

  // V4 schema: legs start at Q (col 17), 7 cols per leg × 4 legs = 28 cols (Q-AR)
  // GO1=Q-W(0-6), GO2=X-AD(7-13), RET1=AE-AK(14-20), RET2=AL-AR(21-27)
  var data = sheet.getRange(2, 17, lastRow - 1, 28).getValues();
  var updates = [];
  var cleared = 0, newCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    function mkLeg(offset) {
      return {
        flightNumber: row[offset],
        depDate: row[offset + 1],
        depTime: row[offset + 2],
        fromCity: row[offset + 3],
        toCity: row[offset + 4],
        arrDate: row[offset + 5],
        arrTime: row[offset + 6]
      };
    }
    var outLegs = [];
    if (row[0])  outLegs.push(mkLeg(0));  // GO1
    if (row[7])  outLegs.push(mkLeg(7));  // GO2
    var retLegs = [];
    if (row[14]) retLegs.push(mkLeg(14)); // RET1
    if (row[21]) retLegs.push(mkLeg(21)); // RET2

    var newWarn = validateFlightLegs_(outLegs, retLegs) || '';
    updates.push([newWarn]);
    if (newWarn) newCount++; else cleared++;
  }

  // Bulk write to column AZ (52)
  sheet.getRange(2, 52, updates.length, 1).setValues(updates);

  Logger.log('🔁 revalidateAllV3 | الإجمالي: ' + updates.length + ' | بتحذيرات: ' + newCount + ' | نظيف: ' + cleared);
  return { total: updates.length, withWarnings: newCount, clean: cleared };
}


/**
 * List rows with warnings (from unified sheet V4)
 */
function listWarningRows() {
  var sheet = getOrCreateUnifiedSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('لا بيانات'); return []; }

  // V4: AZ=warnings (col 52, index 51)
  var data = sheet.getRange(2, 1, lastRow - 1, 52).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var warn = String(data[i][51]); // AZ index=51
    if (warn && warn.trim() !== '') {
      out.push({
        row: i + 2,
        chg: String(data[i][0]),
        pnr: String(data[i][1]),
        name: String(data[i][5]) + ' ' + String(data[i][6]),
        warning: warn
      });
    }
  }
  Logger.log('🚨 عدد الصفوف بتحذيرات: ' + out.length);
  out.forEach(function(r) {
    Logger.log('صف ' + r.row + ' | ' + r.chg + ' | ' + r.pnr + ' | ' + r.name + ' → ' + r.warning);
  });
  return out;
}
