/**
 * ══════════════════════════════════════════════════════════════════
 * نظام إدارة المرشدين — Tour Guide Manager v1.0
 * ══════════════════════════════════════════════════════════════════
 * 
 * المكونات:
 *   1. التحقق من تسجيل الحجاج (مطابقة مع Personal Details)
 *   2. كشف التكرار بين المرشدين
 *   3. بناء جدول الربط المؤكد (Tour Guide Confirmed)
 *   4. تحديث عمود Tour Guide في Personal Details (P) و Room Type (K)
 * 
 * الجداول المعنية:
 *   - Tour Guide: المدخلات من المرشدين (A-L)
 *   - Presonal Details: السجل الرسمي للحجاج
 *   - Tour Guide Confirmed: الربط النهائي المؤكد
 *   - Room Type: عمود K يُحدَّث تلقائياً
 * 
 * مفتاح الربط: Passport Number
 * ══════════════════════════════════════════════════════════════════
 */


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    الإعدادات                                   ║
// ╚═══════════════════════════════════════════════════════════════╝

const TG_CONFIG = {
  // أسماء الشيتات
  SHEET_TOUR_GUIDE: "Tour Guide",
  SHEET_PERSONAL_DETAILS: "Presonal Details",  // الاسم الحالي بالشيت
  SHEET_CONFIRMED: "Tour Guide Confirmed",
  SHEET_ROOM_TYPE: "Room Type",
  
  // أعمدة Tour Guide (الجدول الجديد)
  TG: {
    GUIDE_NAME: 1,      // A - اسم المرشد
    FIRST_NAME: 2,      // B - الاسم الأول
    LAST_NAME: 3,       // C - اسم العائلة
    GENDER: 4,          // D - الجنس
    PASSPORT: 5,        // E - رقم الجواز (مفتاح الربط)
    MOBILE: 6,          // F - الهاتف
    EMAIL: 7,           // G - البريد
    NATIONALITY: 8,     // H - الجنسية
    PACKAGE_NAME: 9,    // I - اسم الباقة
    REG_STATUS: 10,     // J - حالة التسجيل (تلقائي)
    DUPLICATE_CHECK: 11,// K - فحص التكرار (تلقائي)
    SUBMISSION_DATE: 12, // L - تاريخ الإدخال
    DATA_START_ROW: 2
  },
  
  // أعمدة Personal Details
  PD: {
    SERIAL: 1,          // A
    GROUP_NUMBER: 2,    // B
    PASSPORT: 6,        // F - مفتاح الربط
    FIRST_NAME: 11,     // K
    LAST_NAME: 12,      // L
    TOUR_GUIDE: 16,     // P - سيُملأ تلقائياً
    PACKAGE_NAME: 20,   // T
    DATA_START_ROW: 2
  },
  
  // أعمدة Tour Guide Confirmed
  CF: {
    GUIDE_NAME: 1,      // A
    PASSPORT: 2,        // B
    FIRST_NAME: 3,      // C
    LAST_NAME: 4,       // D
    GROUP_NUMBER: 5,    // E
    PACKAGE_NAME: 6,    // F
    NATIONALITY: 7,     // G
    CONFIRMED_DATE: 8,  // H
    DATA_START_ROW: 2
  },
  
  // أعمدة Room Type
  RT: {
    GROUP_NUMBER: 1,    // A
    TOUR_GUIDE: 11,     // K
    DATA_START_ROW: 2
  },
  
  // رموز الحالة
  STATUS: {
    REGISTERED: "✅ Registered",
    NOT_FOUND: "❌ Not Found",
    UNIQUE: "✅ Unique",
    DUPLICATE: "⚠️ Duplicate"
  }
};


/**
 * تطبيع رقم الجواز:
 *   1. إزالة الفراغات من البداية والنهاية
 *   2. إزالة أي رمز غير حرف أو رقم (مثل - * " . وغيرها)
 *   3. تحويل لأحرف كبيرة
 *   4. إزالة الأصفار البادئة إذا كان الجواز أرقام فقط
 */
function _normalizePassport(raw) {
  var s = String(raw || '').trim();
  s = s.replace(/[^A-Za-z0-9]/g, '');
  s = s.toUpperCase();
  if (/^\d+$/.test(s)) s = s.replace(/^0+/, '') || '0';
  return s;
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    القائمة                                     ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * إضافة قائمة Tour Guide — تُستدعى من onOpen في CompleteScript
 * أضف هذا السطر في onOpen():
 *   addTourGuideMenu();
 */
function addTourGuideMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('👥 Tour Guide')
    .addItem('🔍 فحص جميع السجلات', 'tgValidateAll')
    .addItem('🔍 فحص الصف الحالي', 'tgValidateCurrentRow')
    .addSeparator()
    .addItem('✅ بناء جدول الربط المؤكد', 'tgBuildConfirmed')
    .addItem('📤 تحديث Personal Details + Room Type', 'tgPushToSheets')
    .addSeparator()
    .addItem('📊 تقرير المرشدين', 'tgGuideReport')
    .addItem('🏗️ إنشاء هيكل الجداول', 'tgSetupSheets')
    .addToUi();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              إنشاء هيكل الجداول                                ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * إنشاء/تحديث هيكل Tour Guide + Tour Guide Confirmed
 */
function tgSetupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // ═══ Tour Guide Sheet ═══
  let tgSheet = ss.getSheetByName(TG_CONFIG.SHEET_TOUR_GUIDE);
  
  if (tgSheet) {
    // الشيت موجود — نسأل قبل إعادة الهيكلة
    const resp = ui.alert(
      'تنبيه',
      'جدول Tour Guide موجود. هل تريد إعادة كتابة الهيدر فقط (البيانات لن تُحذف)؟',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  } else {
    tgSheet = ss.insertSheet(TG_CONFIG.SHEET_TOUR_GUIDE);
  }
  
  // كتابة الهيدر
  const tgHeaders = [
    'Guide Name', 'First Name', 'Last Name', 'Gender',
    'Passport Number', 'Mobile Number', 'Email', 'Nationality',
    'Package Name', 'Registration Status', 'Duplicate Check', 'Submission Date'
  ];
  tgSheet.getRange(1, 1, 1, tgHeaders.length).setValues([tgHeaders]);
  
  // تنسيق الهيدر
  const tgHeaderRange = tgSheet.getRange(1, 1, 1, tgHeaders.length);
  tgHeaderRange.setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // تنسيق أعمدة الحالة (J, K) — خلفية فاتحة
  tgSheet.getRange(2, TG_CONFIG.TG.REG_STATUS, 1000, 1).setBackground('#e8f5e9');
  tgSheet.getRange(2, TG_CONFIG.TG.DUPLICATE_CHECK, 1000, 1).setBackground('#fff3e0');
  
  // تجميد الصف الأول
  tgSheet.setFrozenRows(1);
  
  // ═══ Tour Guide Confirmed Sheet ═══
  let cfSheet = ss.getSheetByName(TG_CONFIG.SHEET_CONFIRMED);
  if (!cfSheet) {
    cfSheet = ss.insertSheet(TG_CONFIG.SHEET_CONFIRMED);
  }
  
  const cfHeaders = [
    'Guide Name', 'Passport Number', 'First Name', 'Last Name',
    'Group Number', 'Package Name', 'Nationality', 'Confirmed Date'
  ];
  cfSheet.getRange(1, 1, 1, cfHeaders.length).setValues([cfHeaders]);
  
  const cfHeaderRange = cfSheet.getRange(1, 1, 1, cfHeaders.length);
  cfHeaderRange.setBackground('#0d652d')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  cfSheet.setFrozenRows(1);
  
  ui.alert('✅ تم إنشاء/تحديث هيكل الجداول بنجاح');
  SpreadsheetApp.getActiveSpreadsheet().toast('هيكل Tour Guide + Confirmed جاهز', '✅', 5);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              فحص جميع السجلات                                  ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * فحص كل صفوف Tour Guide: تسجيل + تكرار
 */
function tgValidateAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tgSheet = ss.getSheetByName(TG_CONFIG.SHEET_TOUR_GUIDE);
  const pdSheet = ss.getSheetByName(TG_CONFIG.SHEET_PERSONAL_DETAILS);
  
  if (!tgSheet || !pdSheet) {
    SpreadsheetApp.getUi().alert('❌ تأكد من وجود شيت Tour Guide و Presonal Details');
    return;
  }
  
  const lastRow = tgSheet.getLastRow();
  if (lastRow < TG_CONFIG.TG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('⚠️ لا توجد بيانات في Tour Guide');
    return;
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('جارٍ الفحص...', '🔍', -1);
  
  // ═══ تحميل بيانات Personal Details (جواز → بيانات) ═══
  const pdData = pdSheet.getDataRange().getValues();
  const pdMap = {};  // passport → {firstName, lastName, groupNumber, packageName}
  
  for (let i = 1; i < pdData.length; i++) {
    const passport = _normalizePassport(pdData[i][TG_CONFIG.PD.PASSPORT - 1]);
    if (passport) {
      pdMap[passport] = {
        firstName: pdData[i][TG_CONFIG.PD.FIRST_NAME - 1] || '',
        lastName: pdData[i][TG_CONFIG.PD.LAST_NAME - 1] || '',
        groupNumber: pdData[i][TG_CONFIG.PD.GROUP_NUMBER - 1] || '',
        packageName: pdData[i][TG_CONFIG.PD.PACKAGE_NAME - 1] || ''
      };
    }
  }

  // ═══ تحميل بيانات Tour Guide ═══
  const tgData = tgSheet.getRange(TG_CONFIG.TG.DATA_START_ROW, 1, lastRow - 1, TG_CONFIG.TG.SUBMISSION_DATE).getValues();

  // ═══ المرحلة 1: بناء خريطة التكرار الكاملة (جواز → كل المرشدين) ═══
  const passportGuidesMap = {};  // passport → [{guide, rowIndex}, ...]

  for (let i = 0; i < tgData.length; i++) {
    const passport = _normalizePassport(tgData[i][TG_CONFIG.TG.PASSPORT - 1]);
    const guideName = String(tgData[i][TG_CONFIG.TG.GUIDE_NAME - 1] || '').trim();
    
    if (!passport) continue;
    
    if (!passportGuidesMap[passport]) passportGuidesMap[passport] = [];
    passportGuidesMap[passport].push({ guide: guideName, rowIndex: i });
  }
  
  // ═══ المرحلة 2: فحص كل صف ═══
  const regResults = [];
  const dupResults = [];
  
  let registeredCount = 0;
  let notFoundCount = 0;
  let duplicateCount = 0;
  
  for (let i = 0; i < tgData.length; i++) {
    const passport = _normalizePassport(tgData[i][TG_CONFIG.TG.PASSPORT - 1]);
    const guideName = String(tgData[i][TG_CONFIG.TG.GUIDE_NAME - 1] || '').trim();

    // فحص فارغ
    if (!passport) {
      regResults.push(['⚠️ No Passport']);
      dupResults.push(['']);
      continue;
    }
    
    // ═══ 1. فحص التسجيل ═══
    if (pdMap[passport]) {
      regResults.push([TG_CONFIG.STATUS.REGISTERED]);
      registeredCount++;
    } else {
      regResults.push([TG_CONFIG.STATUS.NOT_FOUND]);
      notFoundCount++;
    }
    
    // ═══ 2. فحص التكرار — يعلّم الجميع ═══
    const allEntries = passportGuidesMap[passport] || [];
    
    // جمع أسماء المرشدين الآخرين (غير المرشد الحالي)
    const otherGuides = [...new Set(
      allEntries
        .filter(e => e.guide !== guideName)
        .map(e => e.guide)
    )];
    
    // فحص تكرار عند نفس المرشد
    const sameGuideCount = allEntries.filter(e => e.guide === guideName).length;
    
    if (otherGuides.length > 0) {
      // مكرر عند مرشدين آخرين — يُعلَّم الجميع
      dupResults.push([TG_CONFIG.STATUS.DUPLICATE + ' (Guides: ' + otherGuides.join(', ') + ')']);
      duplicateCount++;
    } else if (sameGuideCount > 1) {
      // مكرر عند نفس المرشد
      dupResults.push(['⚠️ Duplicate (same guide)']);
    } else {
      dupResults.push([TG_CONFIG.STATUS.UNIQUE]);
    }
  }
  
  // ═══ كتابة النتائج بالجملة ═══
  const startRow = TG_CONFIG.TG.DATA_START_ROW;
  tgSheet.getRange(startRow, TG_CONFIG.TG.REG_STATUS, regResults.length, 1).setValues(regResults);
  tgSheet.getRange(startRow, TG_CONFIG.TG.DUPLICATE_CHECK, dupResults.length, 1).setValues(dupResults);
  
  // ═══ تنسيق شرطي ═══
  _applyConditionalFormatting(tgSheet, startRow, regResults.length);
  
  // ═══ التقرير ═══
  const total = tgData.length;
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ تم الفحص: ${total} سجل\n` +
    `مسجل: ${registeredCount} | غير موجود: ${notFoundCount} | مكرر: ${duplicateCount}`,
    '📊 نتيجة الفحص', 10
  );
  
  // تسجيل وقت الفحص
  PropertiesService.getScriptProperties().setProperty('TG_LAST_VALIDATE', new Date().toISOString());
}


/**
 * فحص الصف الحالي فقط
 */
function tgValidateCurrentRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  if (sheet.getName() !== TG_CONFIG.SHEET_TOUR_GUIDE) {
    SpreadsheetApp.getUi().alert('⚠️ يجب أن تكون في شيت Tour Guide');
    return;
  }
  
  const row = ss.getActiveCell().getRow();
  if (row < TG_CONFIG.TG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('⚠️ اختر صفاً يحتوي بيانات');
    return;
  }
  
  const passport = _normalizePassport(sheet.getRange(row, TG_CONFIG.TG.PASSPORT).getValue());
  const guideName = String(sheet.getRange(row, TG_CONFIG.TG.GUIDE_NAME).getValue() || '').trim();

  if (!passport) {
    SpreadsheetApp.getUi().alert('⚠️ لا يوجد رقم جواز في هذا الصف');
    return;
  }

  // فحص التسجيل
  const pdSheet = ss.getSheetByName(TG_CONFIG.SHEET_PERSONAL_DETAILS);
  const pdData = pdSheet.getDataRange().getValues();
  let found = null;

  for (let i = 1; i < pdData.length; i++) {
    if (_normalizePassport(pdData[i][TG_CONFIG.PD.PASSPORT - 1]) === passport) {
      found = pdData[i];
      break;
    }
  }
  
  // فحص التكرار — يجمع كل المرشدين الآخرين
  const tgData = sheet.getDataRange().getValues();
  const otherGuides = [];
  
  for (let i = 1; i < tgData.length; i++) {
    const rowNum = i + 1;
    if (rowNum === row) continue;
    const pp = _normalizePassport(tgData[i][TG_CONFIG.TG.PASSPORT - 1]);
    const gn = String(tgData[i][TG_CONFIG.TG.GUIDE_NAME - 1] || '').trim();
    if (pp === passport && gn !== guideName && !otherGuides.includes(gn)) {
      otherGuides.push(gn);
    }
  }

  // كتابة النتائج
  if (found) {
    sheet.getRange(row, TG_CONFIG.TG.REG_STATUS).setValue(TG_CONFIG.STATUS.REGISTERED);
  } else {
    sheet.getRange(row, TG_CONFIG.TG.REG_STATUS).setValue(TG_CONFIG.STATUS.NOT_FOUND);
  }
  
  if (otherGuides.length > 0) {
    sheet.getRange(row, TG_CONFIG.TG.DUPLICATE_CHECK)
      .setValue(TG_CONFIG.STATUS.DUPLICATE + ' (Guides: ' + otherGuides.join(', ') + ')');
  } else {
    sheet.getRange(row, TG_CONFIG.TG.DUPLICATE_CHECK).setValue(TG_CONFIG.STATUS.UNIQUE);
  }
  
  // رسالة
  const regMsg = found 
    ? `✅ مسجل: ${found[TG_CONFIG.PD.FIRST_NAME - 1]} ${found[TG_CONFIG.PD.LAST_NAME - 1]} — ${found[TG_CONFIG.PD.PACKAGE_NAME - 1]}`
    : '❌ غير موجود في Personal Details';
  const dupMsg = otherGuides.length > 0
    ? `⚠️ مكرر عند المرشدين: ${otherGuides.join(', ')}`
    : '✅ فريد';
  
  SpreadsheetApp.getUi().alert(`📋 نتيجة فحص الجواز: ${passport}\n\n${regMsg}\n${dupMsg}`);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║              بناء جدول الربط المؤكد                            ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * ينشئ/يحدّث Tour Guide Confirmed من السجلات المؤهلة فقط:
 *   - ✅ مسجل في Personal Details
 *   - ✅ فريد (ليس مكرراً عند مرشد آخر)
 */
function tgBuildConfirmed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const tgSheet = ss.getSheetByName(TG_CONFIG.SHEET_TOUR_GUIDE);
  const pdSheet = ss.getSheetByName(TG_CONFIG.SHEET_PERSONAL_DETAILS);
  let cfSheet = ss.getSheetByName(TG_CONFIG.SHEET_CONFIRMED);
  
  if (!tgSheet || !pdSheet) {
    ui.alert('❌ تأكد من وجود Tour Guide و Presonal Details');
    return;
  }
  
  // فحص أولاً
  tgValidateAll();
  
  // قراءة بيانات Tour Guide
  const lastRow = tgSheet.getLastRow();
  if (lastRow < TG_CONFIG.TG.DATA_START_ROW) return;
  
  const tgData = tgSheet.getRange(TG_CONFIG.TG.DATA_START_ROW, 1, lastRow - 1, TG_CONFIG.TG.SUBMISSION_DATE).getValues();
  
  // بناء خريطة Personal Details
  const pdData = pdSheet.getDataRange().getValues();
  const pdMap = {};
  for (let i = 1; i < pdData.length; i++) {
    const passport = _normalizePassport(pdData[i][TG_CONFIG.PD.PASSPORT - 1]);
    if (passport) {
      pdMap[passport] = {
        firstName: pdData[i][TG_CONFIG.PD.FIRST_NAME - 1] || '',
        lastName: pdData[i][TG_CONFIG.PD.LAST_NAME - 1] || '',
        groupNumber: pdData[i][TG_CONFIG.PD.GROUP_NUMBER - 1] || '',
        packageName: pdData[i][TG_CONFIG.PD.PACKAGE_NAME - 1] || '',
        nationality: ''
      };
    }
  }

  // فلترة: مسجل + فريد
  const passportGuideMap = {};
  const confirmedRows = [];
  const now = new Date();

  for (let i = 0; i < tgData.length; i++) {
    const passport = _normalizePassport(tgData[i][TG_CONFIG.TG.PASSPORT - 1]);
    const guideName = String(tgData[i][TG_CONFIG.TG.GUIDE_NAME - 1] || '').trim();
    const nationality = String(tgData[i][TG_CONFIG.TG.NATIONALITY - 1] || '').trim();
    
    if (!passport || !guideName) continue;
    
    // فحص مسجل
    if (!pdMap[passport]) continue;
    
    // فحص فريد (أول مرشد يفوز)
    if (passportGuideMap[passport] && passportGuideMap[passport] !== guideName) continue;
    passportGuideMap[passport] = guideName;
    
    // مكرر عند نفس المرشد — نتجاهل الثاني
    if (confirmedRows.some(r => r[1] === passport)) continue;
    
    const pd = pdMap[passport];
    confirmedRows.push([
      guideName,
      passport,
      pd.firstName,
      pd.lastName,
      pd.groupNumber,
      pd.packageName,
      nationality || pd.nationality,
      now
    ]);
  }
  
  // ═══ كتابة النتائج ═══
  if (!cfSheet) {
    cfSheet = ss.insertSheet(TG_CONFIG.SHEET_CONFIRMED);
    const cfHeaders = [
      'Guide Name', 'Passport Number', 'First Name', 'Last Name',
      'Group Number', 'Package Name', 'Nationality', 'Confirmed Date'
    ];
    cfSheet.getRange(1, 1, 1, cfHeaders.length).setValues([cfHeaders]);
    cfSheet.getRange(1, 1, 1, cfHeaders.length)
      .setBackground('#0d652d')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    cfSheet.setFrozenRows(1);
  }
  
  // مسح البيانات القديمة (الهيدر يبقى)
  const cfLastRow = cfSheet.getLastRow();
  if (cfLastRow > 1) {
    cfSheet.getRange(2, 1, cfLastRow - 1, 8).clearContent();
  }
  
  // كتابة الجديدة
  if (confirmedRows.length > 0) {
    cfSheet.getRange(2, 1, confirmedRows.length, 8).setValues(confirmedRows);
    
    // تنسيق عمود التاريخ
    cfSheet.getRange(2, TG_CONFIG.CF.CONFIRMED_DATE, confirmedRows.length, 1)
      .setNumberFormat('yyyy-MM-dd HH:mm');
  }
  
  // تسجيل
  PropertiesService.getScriptProperties().setProperty('TG_LAST_BUILD', new Date().toISOString());
  
  ui.alert(
    `✅ تم بناء جدول الربط المؤكد\n\n` +
    `إجمالي السجلات المؤهلة: ${confirmedRows.length}\n` +
    `المرشدين الفريدين: ${[...new Set(confirmedRows.map(r => r[0]))].length}`
  );
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║         تحديث Personal Details + Room Type                    ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * يدفع اسم المرشد من Confirmed إلى:
 *   1. Personal Details → عمود P (Tour Guide Name)
 *   2. Room Type → عمود K (Tour Guide)
 */
function tgPushToSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const cfSheet = ss.getSheetByName(TG_CONFIG.SHEET_CONFIRMED);
  const pdSheet = ss.getSheetByName(TG_CONFIG.SHEET_PERSONAL_DETAILS);
  const rtSheet = ss.getSheetByName(TG_CONFIG.SHEET_ROOM_TYPE);
  
  if (!cfSheet) {
    ui.alert('❌ شغّل "بناء جدول الربط المؤكد" أولاً');
    return;
  }
  
  const resp = ui.alert(
    'تأكيد',
    'سيتم تحديث عمود Tour Guide في:\n' +
    '• Personal Details (عمود P)\n' +
    '• Room Type (عمود K)\n\n' +
    'متابعة؟',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  
  SpreadsheetApp.getActiveSpreadsheet().toast('جارٍ التحديث...', '📤', -1);
  
  // ═══ تحميل Confirmed ═══
  const cfLastRow = cfSheet.getLastRow();
  if (cfLastRow < 2) {
    ui.alert('⚠️ جدول الربط المؤكد فارغ');
    return;
  }
  
  const cfData = cfSheet.getRange(2, 1, cfLastRow - 1, 8).getValues();
  
  // بناء خرائط: passport → guideName, groupNumber → guideName
  const passportGuideMap = {};
  const groupGuideMap = {};  // group → Set of guide names
  
  for (const row of cfData) {
    const guideName = String(row[0] || '').trim();
    const passport = _normalizePassport(row[1]);
    const groupNumber = row[4];
    
    if (passport && guideName) {
      passportGuideMap[passport] = guideName;
    }
    if (groupNumber && guideName) {
      if (!groupGuideMap[groupNumber]) groupGuideMap[groupNumber] = new Set();
      groupGuideMap[groupNumber].add(guideName);
    }
  }
  
  // ═══ 1. تحديث Personal Details → عمود P ═══
  let pdUpdated = 0;
  if (pdSheet) {
    const pdLastRow = pdSheet.getLastRow();
    const pdPassports = pdSheet.getRange(2, TG_CONFIG.PD.PASSPORT, pdLastRow - 1, 1).getValues();
    const pdGuideCol = pdSheet.getRange(2, TG_CONFIG.PD.TOUR_GUIDE, pdLastRow - 1, 1).getValues();
    
    for (let i = 0; i < pdPassports.length; i++) {
      const pp = _normalizePassport(pdPassports[i][0]);
      if (pp && passportGuideMap[pp]) {
        pdGuideCol[i][0] = passportGuideMap[pp];
        pdUpdated++;
      }
    }
    
    pdSheet.getRange(2, TG_CONFIG.PD.TOUR_GUIDE, pdGuideCol.length, 1).setValues(pdGuideCol);
  }
  
  // ═══ 2. تحديث Room Type → عمود K ═══
  let rtUpdated = 0;
  if (rtSheet) {
    const rtLastRow = rtSheet.getLastRow();
    if (rtLastRow > 1) {
      const rtGroups = rtSheet.getRange(2, TG_CONFIG.RT.GROUP_NUMBER, rtLastRow - 1, 1).getValues();
      const rtGuideCol = rtSheet.getRange(2, TG_CONFIG.RT.TOUR_GUIDE, rtLastRow - 1, 1).getValues();
      
      for (let i = 0; i < rtGroups.length; i++) {
        const gn = rtGroups[i][0];
        if (gn && groupGuideMap[gn]) {
          const guides = [...groupGuideMap[gn]];
          rtGuideCol[i][0] = guides.join(', ');
          rtUpdated++;
        }
      }
      
      rtSheet.getRange(2, TG_CONFIG.RT.TOUR_GUIDE, rtGuideCol.length, 1).setValues(rtGuideCol);
    }
  }
  
  ui.alert(
    `✅ تم التحديث\n\n` +
    `Personal Details: ${pdUpdated} حاج تم ربطه بمرشد\n` +
    `Room Type: ${rtUpdated} مجموعة تم ربطها بمرشد`
  );
  
  SpreadsheetApp.getActiveSpreadsheet().toast('اكتمل التحديث', '✅', 5);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    تقرير المرشدين                              ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * تقرير ملخص لكل مرشد: عدد الحجاج، المسجلين، غير المسجلين، المكررين
 */
function tgGuideReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tgSheet = ss.getSheetByName(TG_CONFIG.SHEET_TOUR_GUIDE);
  
  if (!tgSheet || tgSheet.getLastRow() < TG_CONFIG.TG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('⚠️ لا توجد بيانات');
    return;
  }
  
  // شغّل الفحص أولاً
  tgValidateAll();
  
  const lastRow = tgSheet.getLastRow();
  const data = tgSheet.getRange(TG_CONFIG.TG.DATA_START_ROW, 1, lastRow - 1, TG_CONFIG.TG.SUBMISSION_DATE).getValues();
  
  // تجميع حسب المرشد
  const guideStats = {};
  
  for (const row of data) {
    const guide = String(row[TG_CONFIG.TG.GUIDE_NAME - 1] || '').trim();
    if (!guide) continue;
    
    if (!guideStats[guide]) {
      guideStats[guide] = { total: 0, registered: 0, notFound: 0, duplicate: 0 };
    }
    
    guideStats[guide].total++;
    
    const regStatus = String(row[TG_CONFIG.TG.REG_STATUS - 1] || '');
    const dupStatus = String(row[TG_CONFIG.TG.DUPLICATE_CHECK - 1] || '');
    
    if (regStatus.includes('Registered')) guideStats[guide].registered++;
    if (regStatus.includes('Not Found')) guideStats[guide].notFound++;
    if (dupStatus.includes('Duplicate') && !dupStatus.includes('same guide')) guideStats[guide].duplicate++;
  }
  
  // بناء التقرير
  let report = '📊 تقرير المرشدين\n';
  report += '══════════════════════════════\n\n';
  
  const sortedGuides = Object.entries(guideStats).sort((a, b) => b[1].total - a[1].total);
  
  let totalAll = 0, regAll = 0, nfAll = 0, dupAll = 0;
  
  for (const [guide, stats] of sortedGuides) {
    report += `👤 ${guide}\n`;
    report += `   📋 الإجمالي: ${stats.total} | ✅ مسجل: ${stats.registered} | ❌ غير موجود: ${stats.notFound} | ⚠️ مكرر: ${stats.duplicate}\n\n`;
    totalAll += stats.total;
    regAll += stats.registered;
    nfAll += stats.notFound;
    dupAll += stats.duplicate;
  }
  
  report += '══════════════════════════════\n';
  report += `المجموع: ${totalAll} حاج | ${Object.keys(guideStats).length} مرشد\n`;
  report += `✅ ${regAll} | ❌ ${nfAll} | ⚠️ ${dupAll}`;
  
  SpreadsheetApp.getUi().alert(report);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    التنسيق الشرطي                              ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * تلوين تلقائي لأعمدة الحالة
 */
function _applyConditionalFormatting(sheet, startRow, numRows) {
  // عمود Registration Status (J)
  const regRange = sheet.getRange(startRow, TG_CONFIG.TG.REG_STATUS, numRows, 1);
  const regValues = regRange.getValues();
  const regBgs = regValues.map(r => {
    const v = String(r[0]);
    if (v.includes('Registered')) return ['#c8e6c9'];  // أخضر فاتح
    if (v.includes('Not Found')) return ['#ffcdd2'];    // أحمر فاتح
    return ['#fff9c4'];  // أصفر فاتح
  });
  regRange.setBackgrounds(regBgs);
  
  // عمود Duplicate Check (K)
  const dupRange = sheet.getRange(startRow, TG_CONFIG.TG.DUPLICATE_CHECK, numRows, 1);
  const dupValues = dupRange.getValues();
  const dupBgs = dupValues.map(r => {
    const v = String(r[0]);
    if (v.includes('Unique')) return ['#c8e6c9'];      // أخضر
    if (v.includes('Duplicate')) return ['#ffe0b2'];    // برتقالي فاتح
    return ['#ffffff'];
  });
  dupRange.setBackgrounds(dupBgs);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    onEdit Trigger                              ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * فحص فوري عند إدخال رقم جواز في Tour Guide
 * يُضاف داخل onEditInstallable الموجود في CompleteScript:
 *   tgOnEdit(e);
 */
function tgOnEdit(e) {
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TG_CONFIG.SHEET_TOUR_GUIDE) return;
  
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // فحص فوري عند إدخال/تعديل رقم الجواز (عمود E)
  if (row >= TG_CONFIG.TG.DATA_START_ROW && col === TG_CONFIG.TG.PASSPORT) {
    const passport = _normalizePassport(e.value);
    if (!passport) return;

    // فحص التسجيل
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pdSheet = ss.getSheetByName(TG_CONFIG.SHEET_PERSONAL_DETAILS);
    const pdData = pdSheet.getDataRange().getValues();

    let found = false;
    for (let i = 1; i < pdData.length; i++) {
      if (_normalizePassport(pdData[i][TG_CONFIG.PD.PASSPORT - 1]) === passport) {
        found = true;
        break;
      }
    }
    
    sheet.getRange(row, TG_CONFIG.TG.REG_STATUS).setValue(
      found ? TG_CONFIG.STATUS.REGISTERED : TG_CONFIG.STATUS.NOT_FOUND
    );
    
    // فحص التكرار — يجمع كل المرشدين الآخرين
    const tgData = sheet.getDataRange().getValues();
    const guideName = String(sheet.getRange(row, TG_CONFIG.TG.GUIDE_NAME).getValue() || '').trim();
    const otherGuides = [];
    
    for (let i = 1; i < tgData.length; i++) {
      if (i + 1 === row) continue;
      const pp = _normalizePassport(tgData[i][TG_CONFIG.TG.PASSPORT - 1]);
      const gn = String(tgData[i][TG_CONFIG.TG.GUIDE_NAME - 1] || '').trim();
      if (pp === passport && gn !== guideName && !otherGuides.includes(gn)) {
        otherGuides.push(gn);
      }
    }
    
    sheet.getRange(row, TG_CONFIG.TG.DUPLICATE_CHECK).setValue(
      otherGuides.length > 0
        ? TG_CONFIG.STATUS.DUPLICATE + ' (Guides: ' + otherGuides.join(', ') + ')'
        : TG_CONFIG.STATUS.UNIQUE
    );
    
    // تلوين الصف
    _applyConditionalFormatting(sheet, row, 1);
  }
}