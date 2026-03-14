/**
 * ══════════════════════════════════════════════════════════════════
 * شيت توفر الغرف - Room Availability
 * ══════════════════════════════════════════════════════════════════
 * 
 * الدوال:
 *   createRoomAvailabilitySheet()  → إنشاء الشيت وملؤه تلقائياً
 *   refreshRoomAvailability()     → تحديث الأسماء بدون مسح التوفر
 *   onEditRoomAvailability(e)     → تحديث "آخر تحديث" تلقائياً
 * 
 * ══════════════════════════════════════════════════════════════════
 */

var ROOM_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'توفر_الغرف',
  PACKAGES_SHEET: 'الباقات',
  PACKAGES_START_ROW: 3,
  
  // القيم الثلاث للتوفر
  AVAILABLE: '✅ متوفر',
  NOT_AVAILABLE: '❌ غير متوفر',
  UNCONFIRMED: '❓ غير مؤكد',
  
  // أعمدة الباقات (المصدر)
  PKG_NUSK: 2,      // B
  PKG_NAME: 3,      // C
  PKG_NAME_EN: 61,  // BI
  PKG_H1_CITY: 12,  // L
  PKG_H1_NAME: 13,  // M
  PKG_H1_EN: 14,    // N
  PKG_H2_CITY: 27,  // AA
  PKG_H2_NAME: 28,  // AB
  PKG_H2_EN: 29,    // AC
  PKG_H3_CITY: 42,  // AP
  PKG_H3_NAME: 43,  // AQ
  PKG_H3_EN: 44     // AR
};


// ╔═══════════════════════════════════════════════════════════════╗
// ║              إنشاء الشيت وملؤه تلقائياً                      ║
// ╚═══════════════════════════════════════════════════════════════╝

function createRoomAvailabilitySheet() {
  var ss = SpreadsheetApp.openById(ROOM_CONFIG.SPREADSHEET_ID);
  
  // التحقق من وجود الشيت
  var existing = ss.getSheetByName(ROOM_CONFIG.SHEET_NAME);
  if (existing) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '⚠️ تنبيه',
      'شيت "' + ROOM_CONFIG.SHEET_NAME + '" موجود بالفعل.\nهل تريد حذفه وإنشاء واحد جديد؟\n\n⚠️ سيتم فقدان بيانات التوفر المدخلة!',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    ss.deleteSheet(existing);
  }
  
  // قراءة بيانات الباقات
  var pkgSheet = ss.getSheetByName(ROOM_CONFIG.PACKAGES_SHEET);
  if (!pkgSheet) {
    Logger.log('❌ شيت الباقات غير موجود');
    return;
  }
  
  var lastRow = pkgSheet.getLastRow();
  if (lastRow < ROOM_CONFIG.PACKAGES_START_ROW) {
    Logger.log('⚠️ لا توجد بيانات في الباقات');
    return;
  }
  
  var numRows = lastRow - ROOM_CONFIG.PACKAGES_START_ROW + 1;
  var pkgData = pkgSheet.getRange(ROOM_CONFIG.PACKAGES_START_ROW, 1, numRows, 66).getValues();
  
  // إنشاء الشيت الجديد
  var sheet = ss.insertSheet(ROOM_CONFIG.SHEET_NAME);
  
  // ───────── الهيدر (صف 1 — المجموعات) ─────────
  var headers1 = [
    '', '', '',
    'الفندق الأول', '', '',
    'الفندق الثاني', '', '',
    'الفندق الثالث', '', '',
    '', ''
  ];
  sheet.getRange(1, 1, 1, headers1.length).setValues([headers1]);
  
  // دمج خلايا المجموعات
  sheet.getRange(1, 4, 1, 3).merge();  // الفندق الأول D-F
  sheet.getRange(1, 7, 1, 3).merge();  // الفندق الثاني G-I
  sheet.getRange(1, 10, 1, 3).merge(); // الفندق الثالث J-L
  
  // ───────── الهيدر (صف 2 — التفاصيل) ─────────
  var headers2 = [
    'Nusk No', 'الباقة', 'Package',
    'Dbl', 'Tri', 'Quad',
    'Dbl', 'Tri', 'Quad',
    'Dbl', 'Tri', 'Quad',
    'آخر تحديث', 'ملاحظات'
  ];
  sheet.getRange(2, 1, 1, headers2.length).setValues([headers2]);
  
  // ───────── تنسيق الهيدر ─────────
  // صف 1
  var h1Range = sheet.getRange(1, 1, 1, headers2.length);
  h1Range.setBackground('#2d5a33');
  h1Range.setFontColor('#ffffff');
  h1Range.setFontWeight('bold');
  h1Range.setFontSize(11);
  h1Range.setHorizontalAlignment('center');
  
  // صف 2
  var h2Range = sheet.getRange(2, 1, 1, headers2.length);
  h2Range.setBackground('#3c7b44');
  h2Range.setFontColor('#ffffff');
  h2Range.setFontWeight('bold');
  h2Range.setFontSize(10);
  h2Range.setHorizontalAlignment('center');
  
  // تلوين مجموعات الفنادق في صف 1
  sheet.getRange(1, 4, 1, 3).setBackground('#e65100'); // الأول — برتقالي (مكة غالباً)
  sheet.getRange(1, 7, 1, 3).setBackground('#1565c0'); // الثاني — أزرق
  sheet.getRange(1, 10, 1, 3).setBackground('#6a1b9a'); // الثالث — بنفسجي
  
  // ───────── ملء البيانات ─────────
  var writeData = [];
  var validRows = 0;
  
  for (var i = 0; i < pkgData.length; i++) {
    var row = pkgData[i];
    var nusk = row[ROOM_CONFIG.PKG_NUSK - 1];
    var name = row[ROOM_CONFIG.PKG_NAME - 1];
    
    if (!name || String(name).trim() === '') continue;
    
    var nameEn = row[ROOM_CONFIG.PKG_NAME_EN - 1] || '';
    var h1 = String(row[ROOM_CONFIG.PKG_H1_NAME - 1] || '').trim();
    var h2 = String(row[ROOM_CONFIG.PKG_H2_NAME - 1] || '').trim();
    var h3 = String(row[ROOM_CONFIG.PKG_H3_NAME - 1] || '').trim();
    
    // تحديد القيمة الافتراضية: غير مؤكد لو الفندق موجود، فارغ لو ما فيه فندق
    var def = ROOM_CONFIG.UNCONFIRMED;
    
    writeData.push([
      nusk,
      String(name).trim(),
      String(nameEn).trim(),
      h1 ? def : '',   // H1 Dbl
      h1 ? def : '',   // H1 Tri
      h1 ? def : '',   // H1 Quad
      h2 ? def : '',   // H2 Dbl
      h2 ? def : '',   // H2 Tri
      h2 ? def : '',   // H2 Quad
      h3 ? def : '',   // H3 Dbl
      h3 ? def : '',   // H3 Tri
      h3 ? def : '',   // H3 Quad
      '',               // آخر تحديث
      ''                // ملاحظات
    ]);
    validRows++;
  }
  
  if (validRows === 0) {
    Logger.log('⚠️ لا توجد باقات صالحة');
    return;
  }
  
  // كتابة البيانات
  var dataRange = sheet.getRange(3, 1, validRows, 14);
  dataRange.setValues(writeData);
  
  // ───────── Data Validation (القوائم المنسدلة) ─────────
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      ROOM_CONFIG.AVAILABLE,
      ROOM_CONFIG.NOT_AVAILABLE,
      ROOM_CONFIG.UNCONFIRMED
    ], true)
    .setAllowInvalid(false)
    .build();
  
  // تطبيق على أعمدة التوفر (D-L) للصفوف المملوءة
  for (var col = 4; col <= 12; col++) {
    sheet.getRange(3, col, validRows, 1).setDataValidation(rule);
  }
  
  // ───────── التنسيق الشرطي ─────────
  var availRange = sheet.getRange(3, 4, validRows, 9); // D3:L
  
  // ✅ أخضر
  var greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(ROOM_CONFIG.AVAILABLE)
    .setBackground('#c8e6c9')
    .setFontColor('#2e7d32')
    .setRanges([availRange])
    .build();
  
  // ❌ أحمر
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(ROOM_CONFIG.NOT_AVAILABLE)
    .setBackground('#ffcdd2')
    .setFontColor('#c62828')
    .setRanges([availRange])
    .build();
  
  // ❓ أصفر
  var yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(ROOM_CONFIG.UNCONFIRMED)
    .setBackground('#fff9c4')
    .setFontColor('#f57f17')
    .setRanges([availRange])
    .build();
  
  sheet.setConditionalFormatRules([greenRule, redRule, yellowRule]);
  
  // ───────── تنسيقات إضافية ─────────
  // تجميد الصفين الأولين والأعمدة الثلاثة الأولى
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(3);
  
  // عرض الأعمدة
  sheet.setColumnWidth(1, 80);   // Nusk No
  sheet.setColumnWidth(2, 250);  // الباقة
  sheet.setColumnWidth(3, 250);  // Package EN
  for (var c = 4; c <= 12; c++) {
    sheet.setColumnWidth(c, 110); // أعمدة التوفر
  }
  sheet.setColumnWidth(13, 140);  // آخر تحديث
  sheet.setColumnWidth(14, 200);  // ملاحظات
  
  // محاذاة وسطية لأعمدة التوفر
  sheet.getRange(3, 4, validRows, 9).setHorizontalAlignment('center');
  sheet.getRange(3, 1, validRows, 1).setHorizontalAlignment('center');
  
  // حماية أعمدة القراءة فقط (A-C)
  var protection = sheet.getRange(3, 1, validRows, 3).protect();
  protection.setDescription('أعمدة تلقائية — لا تعدّل');
  protection.setWarningOnly(true);
  
  // ───────── تقرير النتيجة ─────────
  Logger.log('✅ تم إنشاء شيت "' + ROOM_CONFIG.SHEET_NAME + '"');
  Logger.log('━━━━━━━━━━━━━━━━━');
  Logger.log('📦 عدد الباقات: ' + validRows);
  Logger.log('📊 الأعمدة: 14 (3 معلومات + 9 توفر + تحديث + ملاحظات)');
  Logger.log('🎨 التنسيق الشرطي: ✅ أخضر | ❌ أحمر | ❓ أصفر');
  Logger.log('🔒 الأعمدة A-C محمية (قراءة فقط)');
  Logger.log('❄️ تجميد: صفين + 3 أعمدة');
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'تم إنشاء شيت توفر الغرف بنجاح!\n' + validRows + ' باقة',
    '✅ اكتمل', 5
  );
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║           تحديث آخر تعديل تلقائياً (onEdit)                   ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * يُضاف كـ Installable Trigger (onEdit)
 * أو يُدمج مع onEditInstallable الموجود
 */
function onEditRoomAvailability(e) {
  if (!e || !e.range) return;
  
  var sheet = e.range.getSheet();
  if (sheet.getName() !== ROOM_CONFIG.SHEET_NAME) return;
  
  var row = e.range.getRow();
  var col = e.range.getColumn();
  
  // فقط أعمدة التوفر (D=4 إلى L=12)
  if (row < 3 || col < 4 || col > 12) return;
  
  // تحديث عمود "آخر تحديث" (M=13)
  sheet.getRange(row, 13).setValue(new Date());
  sheet.getRange(row, 13).setNumberFormat('dd/MM HH:mm');
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║           تحديث الأسماء بدون مسح التوفر                       ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * لو أُضيفت باقات جديدة أو تغيّرت الأسماء
 * يحدّث A-C بدون لمس D-N
 */
function refreshRoomAvailability() {
  var ss = SpreadsheetApp.openById(ROOM_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ROOM_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    Logger.log('❌ شيت التوفر غير موجود — شغّل createRoomAvailabilitySheet() أولاً');
    return;
  }
  
  var pkgSheet = ss.getSheetByName(ROOM_CONFIG.PACKAGES_SHEET);
  if (!pkgSheet) return;
  
  var lastRow = pkgSheet.getLastRow();
  var numRows = lastRow - ROOM_CONFIG.PACKAGES_START_ROW + 1;
  var pkgData = pkgSheet.getRange(ROOM_CONFIG.PACKAGES_START_ROW, 1, numRows, 66).getValues();
  
  // قراءة Nusk الموجودة في شيت التوفر
  var existingLastRow = sheet.getLastRow();
  var existingNusk = {};
  if (existingLastRow >= 3) {
    var existingData = sheet.getRange(3, 1, existingLastRow - 2, 1).getValues();
    for (var e = 0; e < existingData.length; e++) {
      existingNusk[String(existingData[e][0])] = e + 3; // row number
    }
  }
  
  var newCount = 0;
  var updateCount = 0;
  
  for (var i = 0; i < pkgData.length; i++) {
    var row = pkgData[i];
    var nusk = row[ROOM_CONFIG.PKG_NUSK - 1];
    var name = row[ROOM_CONFIG.PKG_NAME - 1];
    if (!name || String(name).trim() === '') continue;
    
    var nameEn = row[ROOM_CONFIG.PKG_NAME_EN - 1] || '';
    var h1 = String(row[ROOM_CONFIG.PKG_H1_NAME - 1] || '').trim();
    var h2 = String(row[ROOM_CONFIG.PKG_H2_NAME - 1] || '').trim();
    var h3 = String(row[ROOM_CONFIG.PKG_H3_NAME - 1] || '').trim();
    
    var nuskStr = String(nusk);
    
    if (existingNusk[nuskStr]) {
      // تحديث الأسماء فقط (A-C)
      var targetRow = existingNusk[nuskStr];
      sheet.getRange(targetRow, 1, 1, 3).setValues([[
        nusk, String(name).trim(), String(nameEn).trim()
      ]]);
      updateCount++;
    } else {
      // باقة جديدة — إضافة صف
      var newRow = sheet.getLastRow() + 1;
      var def = ROOM_CONFIG.UNCONFIRMED;
      sheet.getRange(newRow, 1, 1, 14).setValues([[
        nusk,
        String(name).trim(),
        String(nameEn).trim(),
        h1 ? def : '', h1 ? def : '', h1 ? def : '',
        h2 ? def : '', h2 ? def : '', h2 ? def : '',
        h3 ? def : '', h3 ? def : '', h3 ? def : '',
        '', ''
      ]]);
      
      // Data Validation للصف الجديد
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList([
          ROOM_CONFIG.AVAILABLE,
          ROOM_CONFIG.NOT_AVAILABLE,
          ROOM_CONFIG.UNCONFIRMED
        ], true)
        .setAllowInvalid(false)
        .build();
      
      for (var c = 4; c <= 12; c++) {
        sheet.getRange(newRow, c).setDataValidation(rule);
      }
      sheet.getRange(newRow, 4, 1, 9).setHorizontalAlignment('center');
      
      newCount++;
    }
  }
  
  Logger.log('✅ تم التحديث');
  Logger.log('📝 تحديث أسماء: ' + updateCount);
  Logger.log('🆕 باقات جديدة: ' + newCount);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'تحديث: ' + updateCount + ' | جديد: ' + newCount,
    '✅ تم التحديث', 5
  );
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    القائمة                                    ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * أضف هذا في دالة onOpen() الموجودة أو شغّلها مستقلة
 */
function addRoomAvailabilityMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛏️ توفر الغرف')
    .addItem('🆕 إنشاء شيت التوفر', 'createRoomAvailabilitySheet')
    .addItem('🔄 تحديث الأسماء (بدون مسح التوفر)', 'refreshRoomAvailability')
    .addToUi();
}