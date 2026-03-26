/**
 * ══════════════════════════════════════════════════════════════════
 * نظام البحث عن الحجاج - شركة إكرام الضيف للسياحة
 * الإصدار 2.0
 * التعديلات:
 *   - مصدر المرشد من شيت Tour Guide (بشرط Registered + Unique)
 *   - مصدر المواصلات من شيت الباقات عمود BN
 * ══════════════════════════════════════════════════════════════════
 */

var CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  PILGRIMS_SHEET: 'Presonal Details',
  PACKAGES_SHEET: 'الباقات',
  FLIGHTS_SHEET: 'الطيران',
  SETTINGS_SHEET: 'الإعدادات',
  TOUR_GUIDE_SHEET: 'Tour Guide',
  PILGRIMS_START: 2,
  PACKAGES_START: 3,
  FLIGHTS_START: 3,
  TOUR_GUIDE_START: 2
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('PilgrimIndex')
    .evaluate()
    .setTitle('Ikram AlDyf - Pilgrim Search')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ══════════════════════════════════════════════════════════════
// إحصائية المرشدين لكل باقة
// ══════════════════════════════════════════════════════════════
function getGuidePackageStats() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var guideSheet = ss.getSheetByName(CONFIG.TOUR_GUIDE_SHEET);
    var pilgrimSheet = ss.getSheetByName(CONFIG.PILGRIMS_SHEET);
    if (!guideSheet || !pilgrimSheet) return JSON.stringify({ success: false, error: 'Sheet not found' });

    // بناء خريطة جواز → باقة من Presonal Details
    var pkgMap = {};
    var pLastRow = pilgrimSheet.getLastRow();
    if (pLastRow >= CONFIG.PILGRIMS_START) {
      var pData = pilgrimSheet.getRange(CONFIG.PILGRIMS_START, 1, pLastRow - CONFIG.PILGRIMS_START + 1, 20).getValues();
      for (var i = 0; i < pData.length; i++) {
        var passport = String(pData[i][5]).trim().toUpperCase();
        var pkgName  = String(pData[i][19]).trim();
        if (passport && pkgName) pkgMap[passport] = pkgName;
      }
    }

    // بناء الإحصائية من Tour Guide
    var gLastRow = guideSheet.getLastRow();
    var pkgStats = {};
    if (gLastRow >= CONFIG.TOUR_GUIDE_START) {
      var gData = guideSheet.getRange(CONFIG.TOUR_GUIDE_START, 1, gLastRow - CONFIG.TOUR_GUIDE_START + 1, 11).getValues();
      for (var j = 0; j < gData.length; j++) {
        var guideName = String(gData[j][0]).trim();
        var gPassport = String(gData[j][4]).trim().toUpperCase();
        var pkg       = pkgMap[gPassport] || String(gData[j][8]).trim();
        if (!guideName || !pkg) continue;
        if (!pkgStats[pkg]) pkgStats[pkg] = {};
        pkgStats[pkg][guideName] = (pkgStats[pkg][guideName] || 0) + 1;
      }
    }

    // تحويل إلى مصفوفة مرتبة
    var result = [];
    for (var pkg in pkgStats) {
      var guidesArr = [];
      for (var g in pkgStats[pkg]) {
        guidesArr.push({ name: g, count: pkgStats[pkg][g] });
      }
      guidesArr.sort(function(a, b) { return b.count - a.count; });
      var total = guidesArr.reduce(function(s, x) { return s + x.count; }, 0);
      result.push({ pkg: pkg, guides: guidesArr, total: total });
    }
    result.sort(function(a, b) { return b.total - a.total; });

    return JSON.stringify({ success: true, data: result });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPassword() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
    if (!sheet) return '1234';
    return sheet.getRange('B1').getValue().toString();
  } catch (e) {
    return '1234';
  }
}

// ══════════════════════════════════════════════════════════════
// بناء خريطة المرشدين من شيت Tour Guide
// الشرط: J = ✅ Registered و K = ✅ Unique
// المفتاح: رقم الجواز (E) → اسم المرشد (A)
// ══════════════════════════════════════════════════════════════

function buildGuideMap(ss) {
  var map = {};
  try {
    var sheet = ss.getSheetByName(CONFIG.TOUR_GUIDE_SHEET);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.TOUR_GUIDE_START) return map;

    var numRows = lastRow - CONFIG.TOUR_GUIDE_START + 1;
    var data = sheet.getRange(CONFIG.TOUR_GUIDE_START, 1, numRows, 11).getValues();

    for (var i = 0; i < data.length; i++) {
      var guideName = String(data[i][0]).trim();        // A: اسم المرشد
      var passport  = String(data[i][4]).trim().toUpperCase(); // E: رقم الجواز
      var regStatus = String(data[i][9]).trim();         // J: حالة التسجيل
      var dupCheck  = String(data[i][10]).trim();        // K: فحص التكرار

      // إذا كانت أعمدة J و K فارغة نقبل السجل، وإذا كانت ممتلئة نتحقق منها
      var regOk  = !regStatus  || regStatus.indexOf('Registered') !== -1;
      var uniqOk = !dupCheck   || dupCheck.indexOf('Unique')      !== -1;

      if (guideName && passport && regOk && uniqOk) {
        map[passport] = guideName;
      }
    }
  } catch (e) {
    Logger.log('buildGuideMap ERROR: ' + e.toString());
  }
  return map;
}

// ══════════════════════════════════════════════════════════════
// بناء خريطة المواصلات من شيت الباقات
// المفتاح: Nusk No (عمود B) → نوع التنقل (عمود BN = 66)
// ══════════════════════════════════════════════════════════════

function buildTransportMap(ss) {
  var map = {};
  try {
    var sheet = ss.getSheetByName(CONFIG.PACKAGES_SHEET);
    if (!sheet) return map;
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.PACKAGES_START) return map;

    var numRows = lastRow - CONFIG.PACKAGES_START + 1;
    // نقرأ عمود B (2) و عمود BN (66)
    var dataNusk = sheet.getRange(CONFIG.PACKAGES_START, 2, numRows, 1).getValues();
    var dataTransport = sheet.getRange(CONFIG.PACKAGES_START, 66, numRows, 1).getValues();

    for (var i = 0; i < dataNusk.length; i++) {
      var nusk = dataNusk[i][0];
      var transport = String(dataTransport[i][0]).trim();
      if (nusk && transport) {
        map[Number(nusk)] = transport;
      }
    }
  } catch (e) {
    Logger.log('buildTransportMap ERROR: ' + e.toString());
  }
  return map;
}

// ══════════════════════════════════════════════════════════════
// جلب بيانات الحجاج
// ══════════════════════════════════════════════════════════════

function getAllPilgrims() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.PILGRIMS_SHEET);
    if (!sheet) return JSON.stringify({ pilgrims: [], stats: {}, success: false, error: 'Sheet not found' });

    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.PILGRIMS_START) return JSON.stringify({ pilgrims: [], stats: {}, success: true });

    // بناء الخرائط أولاً
    var guideMap = buildGuideMap(ss);
    var transportMap = buildTransportMap(ss);

    var numRows = lastRow - CONFIG.PILGRIMS_START + 1;
    var data = sheet.getRange(CONFIG.PILGRIMS_START, 1, numRows, 32).getValues();
    var pilgrims = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var serial = row[0];
      if (!serial || String(serial).trim() === '') continue;

      var passport = cleanStr(row[5]).toUpperCase();
      var packageNo = toNum(row[18]);

      pilgrims.push({
        serial: toNum(serial),
        groupNo: toNum(row[1]),
        pilgrimType: cleanStr(row[2]),
        category: cleanStr(row[3]),
        gender: cleanStr(row[4]),
        passport: cleanStr(row[5]),
        passportExpiry: fmtDate(row[6]),
        passportIssue: fmtDate(row[7]),
        firstNameAr: cleanStr(row[8]),
        lastNameAr: cleanStr(row[9]),
        firstNameEn: cleanStr(row[10]),
        lastNameEn: cleanStr(row[11]),
        birthDate: fmtDate(row[12]),
        email: cleanStr(row[13]),
        phone: cleanStr(row[14]),
        guide: guideMap[passport] || '',
        country: cleanStr(row[16]),
        nationality: cleanStr(row[17]),
        packageNo: packageNo,
        packageName: cleanStr(row[19]),
        flightType: cleanStr(row[20]),
        contractName: cleanStr(row[21]),
        visaStatus: cleanStr(row[22]),
        isLocal: row[23] === true || row[23] === 'TRUE' || row[23] === 'True',
        ticketNo: cleanStr(row[24]),
        ticketLink: cleanStr(row[25]),
        invoiceNo: cleanStr(row[26]),
        camp: cleanStr(row[27]),
        transportType: transportMap[packageNo] || '',
        arrivalTime: fmtDateTime(row[28]),
        departureTime: fmtDateTime(row[30]),
        bookingDetails: cleanStr(row[31])
      });
    }

    var stats = calcStats(pilgrims);

    return JSON.stringify({
      pilgrims: pilgrims,
      stats: stats,
      success: true
    });

  } catch (error) {
    Logger.log('getAllPilgrims ERROR: ' + error.toString());
    return JSON.stringify({ pilgrims: [], stats: {}, success: false, error: error.toString() });
  }
}

// ══════════════════════════════════════════════════════════════
// حساب الإحصائيات
// ══════════════════════════════════════════════════════════════

function calcStats(pilgrims) {
  var total = pilgrims.length;
  var groups = {};
  var countries = {};
  var packages = {};
  var nationalities = {};
  var camps = {};
  var transports = {};
  var b2b = 0, b2c = 0, local = 0;
  var male = 0, female = 0;
  var guides = {};

  for (var i = 0; i < pilgrims.length; i++) {
    var p = pilgrims[i];
    groups[p.groupNo] = true;

    if (p.country) {
      countries[p.country] = (countries[p.country] || 0) + 1;
    }

    if (p.packageNo) {
      packages[p.packageNo] = (packages[p.packageNo] || 0) + 1;
    }

    if (p.nationality) {
      nationalities[p.nationality] = (nationalities[p.nationality] || 0) + 1;
    }

    if (p.camp) {
      camps[p.camp] = (camps[p.camp] || 0) + 1;
    }

    if (p.transportType) {
      transports[p.transportType] = (transports[p.transportType] || 0) + 1;
    }

    if (p.flightType === 'B2B') b2b++;
    else if (p.flightType === 'B2C') b2c++;

    if (p.isLocal) local++;

    if (p.gender === 'ذكر') male++;
    else if (p.gender === 'انثى' || p.gender === 'أنثى') female++;

    if (p.guide) {
      guides[p.guide] = (guides[p.guide] || 0) + 1;
    }
  }

  return {
    totalPilgrims: total,
    totalGroups: Object.keys(groups).length,
    totalCountries: Object.keys(countries).length,
    totalPackages: Object.keys(packages).length,
    totalNationalities: Object.keys(nationalities).length,
    totalCamps: Object.keys(camps).length,
    totalTransports: Object.keys(transports).length,
    b2b: b2b,
    b2c: b2c,
    local: local,
    male: male,
    female: female,
    totalGuides: Object.keys(guides).length
  };
}

// ══════════════════════════════════════════════════════════════
// دوال مساعدة
// ══════════════════════════════════════════════════════════════

function cleanStr(val) {
  if (!val && val !== 0) return '';
  return String(val).trim();
}

function toNum(val) {
  if (!val && val !== 0) return 0;
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

function fmtDate(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      if (val.getFullYear() < 2000 && val.getFullYear() > 1950) {
        // Could be a valid birth date
      } else if (val.getFullYear() < 1900) {
        return '';
      }
      return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2) + '-' + ('0' + val.getDate()).slice(-2);
    }
    var str = String(val).trim();
    if (str.indexOf('1899') !== -1 || str.indexOf('1900') !== -1) return '';
    return str;
  } catch (e) {
    return '';
  }
}

function fmtDateTime(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      if (val.getFullYear() < 2000) return '';
      return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2) + '-' + ('0' + val.getDate()).slice(-2) +
        ' ' + ('0' + val.getHours()).slice(-2) + ':' + ('0' + val.getMinutes()).slice(-2);
    }
    return String(val).trim();
  } catch (e) {
    return '';
  }
}