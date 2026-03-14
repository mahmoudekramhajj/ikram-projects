/**
 * ═══════════════════════════════════════════════════════════════
 * MinaCamp.gs — نظام توزيع خيام مِنى
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0
 * ───────────────────────────────────────────────────────────────
 * المشروع: Ikram (المرتبط بالشيت الرئيسي)
 * المشغّل: قائمة يدوية ⛺ مخيم مِنى (تُفعّل من onOpen)
 * يقرأ: شيت "رحلة الحاج " + شيت "Tour Guide"
 * يكتب: شيت "مخيم مني"
 * ═══════════════════════════════════════════════════════════════
 */

// ============================================================
// الإعدادات
// ============================================================

var MINA = {
  SS_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  
  SHEETS: {
    JOURNEY: 'رحلة الحاج ',   // ⚠ مسافة في النهاية
    GUIDE: 'Tour Guide',
    MINA: 'مخيم مني'
  },
  
  // أعمدة رحلة الحاج (0-based)
  J: {
    CAMP: 4,             // E: المخيم (المعيصم / مجر الكبش)
    APPLICANT_ID: 5,     // F: ApplicantId — المفتاح الفريد
    GROUP_NUMBER: 6,     // G: رقم المجموعة
    NAME: 7,             // H: اسم الحاج
    PASSPORT: 8,         // I: جواز السفر
    GENDER: 11,          // L: الجنس
    NATIONALITY: 12,     // M: الجنسية
    RESIDENCE: 14        // O: دولة الإقامة
  },
  
  // أعمدة Tour Guide (0-based)
  G: { NAME: 0, PASSPORT: 4, REG: 9, UNIQUE: 10 },
  GUIDE_START_ROW: 2,
  
  // إعدادات التوزيع
  CAPACITY: 80,           // عدد الحجاج لكل خيمة
  TENT_PREFIX: '49/',      // بادئة ترقيم الخيام
  TENT_START_NUM: 2,       // أول رقم خيمة → 49/2
  SOFA_START: 29710,       // أول رقم صوفا بيد
  
  // ترتيب الأقسام الأربعة
  SECTIONS: [
    { camp: 'المعيصم', gender: 'M', label: 'المعيصم رجال' },
    { camp: 'مجر الكبش', gender: 'M', label: 'مجر الكبش رجال' },
    { camp: 'المعيصم', gender: 'F', label: 'المعيصم نساء' },
    { camp: 'مجر الكبش', gender: 'F', label: 'مجر الكبش نساء' }
  ],
  
  // رؤوس أعمدة شيت المخرجات
  HEADERS: [
    'رقم الخيمة',       // A
    'رقم الصوفا بيد',   // B
    'ApplicantId',      // C
    'اسم الحاج',        // D
    'جواز السفر',       // E
    'الجنس',            // F
    'الجنسية',          // G
    'دولة الإقامة',     // H
    'المخيم',           // I
    'المرشد',           // J
    'رقم المجموعة'      // K
  ]
};


// ============================================================
// القائمة — تُستدعى من onOpen في CompleteScript.gs
// ============================================================

function addMinaCampMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⛺ مخيم مِنى')
    .addItem('🔄 توزيع الخيام', 'assignMinaTents')
    .addItem('📊 إحصائيات التوزيع', 'showMinaStats')
    .addToUi();
}


// ============================================================
// التوزيع الرئيسي
// ============================================================

/**
 * الدالة الرئيسية: توزيع الحجاج على خيام مِنى
 * المشغّل: قائمة يدوية
 */
function assignMinaTents() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '⛺ توزيع خيام مِنى',
    'سيتم مسح البيانات الحالية في شيت "مخيم مني" وإعادة التوزيع.\n\nمتابعة؟',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;
  
  var startTime = new Date();
  var ss = SpreadsheetApp.openById(MINA.SS_ID);
  
  // ── المرحلة 1: جلب البيانات ──
  SpreadsheetApp.getActiveSpreadsheet().toast('جارِ قراءة البيانات...', '⛺ مخيم مِنى', 30);
  
  var journeyData = readJourneyData_(ss);
  if (!journeyData.length) {
    ui.alert('⚠️ لا توجد بيانات في شيت رحلة الحاج');
    return;
  }
  
  var guideMap = buildMinaGuideMap_(ss);
  Logger.log('إجمالي الحجاج: ' + journeyData.length + ' | المرشدين: ' + Object.keys(guideMap).length);
  
  // ── المرحلة 2: تقسيم إلى 4 أقسام ──
  SpreadsheetApp.getActiveSpreadsheet().toast('جارِ تقسيم الحجاج...', '⛺ مخيم مِنى', 30);
  
  var sections = splitIntoSections_(journeyData, guideMap);
  
  // ── المرحلة 3: ترتيب كل قسم ──
  for (var s = 0; s < sections.length; s++) {
    sections[s].pilgrims = sortSection_(sections[s].pilgrims);
  }
  
  // ── المرحلة 4: توزيع الخيام + ترقيم الصوفا ──
  SpreadsheetApp.getActiveSpreadsheet().toast('جارِ توزيع الخيام...', '⛺ مخيم مِنى', 30);
  
  var result = assignTentsAndSofas_(sections);
  
  // ── المرحلة 5: كتابة النتائج ──
  SpreadsheetApp.getActiveSpreadsheet().toast('جارِ كتابة النتائج...', '⛺ مخيم مِنى', 30);
  
  writeToMinaSheet_(ss, result.rows);
  
  // ── ملخص ──
  var elapsed = ((new Date() - startTime) / 1000).toFixed(1);
  var summary = '✅ تم التوزيع بنجاح\n\n'
    + '📊 الإجمالي: ' + result.rows.length + ' حاج\n'
    + '🏕️ عدد الخيام: ' + result.totalTents + '\n'
    + '⏱️ الوقت: ' + elapsed + ' ثانية\n\n'
    + '── تفاصيل الأقسام ──\n';
  
  for (var i = 0; i < result.sectionStats.length; i++) {
    var stat = result.sectionStats[i];
    summary += stat.label + ': ' + stat.pilgrims + ' حاج → ' + stat.tents + ' خيمة'
      + ' (من ' + stat.firstTent + ' إلى ' + stat.lastTent + ')\n';
  }
  
  ui.alert('⛺ نتائج التوزيع', summary, ui.ButtonSet.OK);
}


// ============================================================
// المرحلة 1: قراءة البيانات
// ============================================================

function readJourneyData_(ss) {
  var sheet = ss.getSheetByName(MINA.SHEETS.JOURNEY);
  if (!sheet) {
    Logger.log('شيت رحلة الحاج غير موجود');
    return [];
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var J = MINA.J;
  var pilgrims = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var applicantId = String(row[J.APPLICANT_ID]).trim();
    var camp = String(row[J.CAMP]).trim();
    var gender = String(row[J.GENDER]).trim();
    
    // تخطي الصفوف بدون ApplicantId أو بدون مخيم
    if (!applicantId || !camp) continue;
    
    pilgrims.push({
      applicantId: applicantId,
      name: String(row[J.NAME]).trim(),
      passport: String(row[J.PASSPORT]).trim().toUpperCase(),
      gender: normalizeGender_(gender),
      genderOriginal: gender,
      nationality: String(row[J.NATIONALITY]).trim(),
      residence: String(row[J.RESIDENCE]).trim(),
      camp: normalizeCamp_(camp),
      campOriginal: camp,
      groupNumber: String(row[J.GROUP_NUMBER]).trim(),
      guide: ''
    });
  }
  
  return pilgrims;
}


function buildMinaGuideMap_(ss) {
  var map = {};
  try {
    var sheet = ss.getSheetByName(MINA.SHEETS.GUIDE);
    if (!sheet) return map;
    
    var lastRow = sheet.getLastRow();
    var startRow = MINA.GUIDE_START_ROW;
    if (lastRow < startRow) return map;
    
    var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 11).getValues();
    var G = MINA.G;
    
    for (var i = 0; i < data.length; i++) {
      var guideName = String(data[i][G.NAME]).trim();
      var passport = String(data[i][G.PASSPORT]).trim().toUpperCase();
      var regStatus = String(data[i][G.REG]).trim();
      var dupCheck = String(data[i][G.UNIQUE]).trim();
      
      if (guideName && passport
          && regStatus.indexOf('Registered') !== -1
          && dupCheck.indexOf('Unique') !== -1) {
        map[passport] = guideName;
      }
    }
  } catch (e) {
    Logger.log('buildMinaGuideMap_ ERROR: ' + e.toString());
  }
  return map;
}


// ============================================================
// المرحلة 2: تقسيم إلى أقسام
// ============================================================

function splitIntoSections_(pilgrims, guideMap) {
  // ربط المرشد
  for (var i = 0; i < pilgrims.length; i++) {
    var passport = pilgrims[i].passport;
    if (guideMap[passport]) {
      pilgrims[i].guide = guideMap[passport];
    }
  }
  
  // تهيئة الأقسام
  var sections = [];
  for (var s = 0; s < MINA.SECTIONS.length; s++) {
    sections.push({
      camp: MINA.SECTIONS[s].camp,
      gender: MINA.SECTIONS[s].gender,
      label: MINA.SECTIONS[s].label,
      pilgrims: []
    });
  }
  
  // توزيع الحجاج على الأقسام
  var unmatched = 0;
  for (var i = 0; i < pilgrims.length; i++) {
    var p = pilgrims[i];
    var placed = false;
    
    for (var s = 0; s < sections.length; s++) {
      if (p.camp === sections[s].camp && p.gender === sections[s].gender) {
        sections[s].pilgrims.push(p);
        placed = true;
        break;
      }
    }
    
    if (!placed) {
      unmatched++;
      Logger.log('⚠ حاج غير مُصنّف: ' + p.applicantId
        + ' | مخيم=' + p.campOriginal + ' | جنس=' + p.genderOriginal);
    }
  }
  
  if (unmatched > 0) {
    Logger.log('⚠ إجمالي الحجاج غير المُصنّفين: ' + unmatched);
  }
  
  for (var s = 0; s < sections.length; s++) {
    Logger.log(sections[s].label + ': ' + sections[s].pilgrims.length + ' حاج');
  }
  
  return sections;
}


// ============================================================
// المرحلة 3: ترتيب داخل كل قسم
// ============================================================

/**
 * ترتيب بأولوية: المرشد → دولة الإقامة → الجنسية
 */
function sortSection_(pilgrims) {
  return pilgrims.sort(function(a, b) {
    var guideA = a.guide || 'zzz_no_guide';
    var guideB = b.guide || 'zzz_no_guide';
    if (guideA !== guideB) return guideA.localeCompare(guideB);
    
    var resA = a.residence || 'zzz_no_residence';
    var resB = b.residence || 'zzz_no_residence';
    if (resA !== resB) return resA.localeCompare(resB);
    
    var natA = a.nationality || 'zzz_no_nationality';
    var natB = b.nationality || 'zzz_no_nationality';
    return natA.localeCompare(natB);
  });
}


// ============================================================
// المرحلة 4: توزيع الخيام + ترقيم الصوفا
// ============================================================

/**
 * 80 حاج/خيمة — كل قسم يبدأ خيمة جديدة
 * تسلسل مستمر للخيام والصوفا عبر الأقسام الأربعة
 */
function assignTentsAndSofas_(sections) {
  var currentTentNum = MINA.TENT_START_NUM;
  var currentSofa = MINA.SOFA_START;
  var allRows = [];
  var sectionStats = [];
  var totalTents = 0;
  
  for (var s = 0; s < sections.length; s++) {
    var sec = sections[s];
    var pilgrims = sec.pilgrims;
    
    if (pilgrims.length === 0) {
      sectionStats.push({
        label: sec.label, pilgrims: 0, tents: 0,
        firstTent: '-', lastTent: '-'
      });
      continue;
    }
    
    var sectionFirstTent = MINA.TENT_PREFIX + currentTentNum;
    var countInCurrentTent = 0;
    
    for (var i = 0; i < pilgrims.length; i++) {
      if (countInCurrentTent >= MINA.CAPACITY) {
        currentTentNum++;
        countInCurrentTent = 0;
      }
      
      var tentLabel = MINA.TENT_PREFIX + currentTentNum;
      var p = pilgrims[i];
      
      allRows.push([
        tentLabel,           // A: رقم الخيمة
        currentSofa,         // B: رقم الصوفا بيد
        p.applicantId,       // C: ApplicantId
        p.name,              // D: اسم الحاج
        p.passport,          // E: جواز السفر
        p.genderOriginal,    // F: الجنس
        p.nationality,       // G: الجنسية
        p.residence,         // H: دولة الإقامة
        p.campOriginal,      // I: المخيم
        p.guide,             // J: المرشد
        p.groupNumber        // K: رقم المجموعة
      ]);
      
      currentSofa++;
      countInCurrentTent++;
    }
    
    var sectionLastTent = MINA.TENT_PREFIX + currentTentNum;
    var sectionTents = currentTentNum - parseInt(sectionFirstTent.split('/')[1]) + 1;
    totalTents += sectionTents;
    
    sectionStats.push({
      label: sec.label,
      pilgrims: pilgrims.length,
      tents: sectionTents,
      firstTent: sectionFirstTent,
      lastTent: sectionLastTent
    });
    
    // القسم التالي يبدأ خيمة جديدة
    currentTentNum++;
  }
  
  return {
    rows: allRows,
    totalTents: totalTents,
    sectionStats: sectionStats
  };
}


// ============================================================
// المرحلة 5: كتابة النتائج
// ============================================================

function writeToMinaSheet_(ss, rows) {
  var sheet = ss.getSheetByName(MINA.SHEETS.MINA);
  
  if (!sheet) {
    sheet = ss.insertSheet(MINA.SHEETS.MINA);
  }
  
  sheet.clearContents();
  sheet.clearFormats();
  
  // الرؤوس
  var headers = MINA.HEADERS;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a237e')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  if (rows.length === 0) return;
  
  // البيانات بالجملة
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
  // تنسيقات
  sheet.setFrozenRows(1);
  sheet.setRightToLeft(true);
  
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 60);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 100);
  
  // تلوين متبادل لكل خيمة
  applyTentColoring_(sheet, rows);
  
  Logger.log('تم كتابة ' + rows.length + ' صف في شيت مخيم مني');
}


function applyTentColoring_(sheet, rows) {
  if (rows.length === 0) return;
  
  var colors = ['#e3f2fd', '#ffffff'];
  var currentTent = '';
  var colorIdx = 0;
  
  var backgrounds = [];
  for (var i = 0; i < rows.length; i++) {
    var tent = rows[i][0];
    if (tent !== currentTent) {
      currentTent = tent;
      colorIdx = 1 - colorIdx;
    }
    var rowColor = [];
    for (var c = 0; c < MINA.HEADERS.length; c++) {
      rowColor.push(colors[colorIdx]);
    }
    backgrounds.push(rowColor);
  }
  
  sheet.getRange(2, 1, rows.length, MINA.HEADERS.length).setBackgrounds(backgrounds);
}


// ============================================================
// الدوال المساعدة
// ============================================================

function normalizeGender_(gender) {
  if (!gender) return '';
  var g = gender.toLowerCase().trim();
  if (g === 'ذكر' || g === 'male' || g === 'm') return 'M';
  if (g === 'أنثى' || g === 'انثى' || g === 'female' || g === 'f') return 'F';
  return '';
}

function normalizeCamp_(camp) {
  if (!camp) return '';
  var c = camp.trim();
  if (c.indexOf('معيصم') !== -1 || c.indexOf('معيصيم') !== -1) return 'المعيصم';
  if (c.indexOf('مجر') !== -1 || c.indexOf('كبش') !== -1) return 'مجر الكبش';
  return c;
}


// ============================================================
// إحصائيات التوزيع
// ============================================================

function showMinaStats() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.openById(MINA.SS_ID);
  var sheet = ss.getSheetByName(MINA.SHEETS.MINA);
  
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('⚠️ شيت مخيم مني فارغ — شغّل التوزيع أولاً');
    return;
  }
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, MINA.HEADERS.length).getValues();
  
  var tents = {}, camps = {}, genders = {}, guides = {};
  
  for (var i = 0; i < data.length; i++) {
    tents[data[i][0]] = (tents[data[i][0]] || 0) + 1;
    camps[data[i][8]] = (camps[data[i][8]] || 0) + 1;
    genders[data[i][5]] = (genders[data[i][5]] || 0) + 1;
    if (data[i][9]) guides[data[i][9]] = (guides[data[i][9]] || 0) + 1;
  }
  
  var guidePilgrims = 0;
  for (var g in guides) guidePilgrims += guides[g];
  
  var msg = '📊 إحصائيات مخيم مِنى\n\n'
    + 'إجمالي الحجاج: ' + data.length + '\n'
    + 'عدد الخيام: ' + Object.keys(tents).length + '\n\n'
    + '── حسب المخيم ──\n';
  for (var c in camps) msg += c + ': ' + camps[c] + '\n';
  
  msg += '\n── حسب الجنس ──\n';
  for (var gn in genders) msg += gn + ': ' + genders[gn] + '\n';
  
  msg += '\n── المرشدون ──\n'
    + 'عدد المرشدين: ' + Object.keys(guides).length + '\n'
    + 'حجاج بدون مرشد: ' + (data.length - guidePilgrims);
  
  ui.alert('📊 إحصائيات', msg, ui.ButtonSet.OK);
}