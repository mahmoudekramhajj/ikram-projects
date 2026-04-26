/**
 * PilgrimStats — لوحة إحصائيات الحجاج
 * شاشات عرض منفصلة للشاشات الكبيرة — فريق العمليات
 */

var CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'Presonal Details',
  CACHE_KEY: 'PILGRIM_STATS_V8',
  CACHE_DURATION: 300
};

// أسماء أعمدة Presonal Details
var COL = {
  CATEGORY: 'فئة الحجاج',   // عمود D — أفراد / مجاملة
  GENDER: 'الجنس',
  RESIDENCE: 'بلد الإقامة',
  NATIONALITY: 'الجنسية',
  FLIGHT_TYPE: 'نوع عقد الطيران',
  CAMP: 'المخيم'
};

function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';

  // ── منفذ تشخيصي — يعيد JSON ──
  if (view === 'diagnose') {
    return ContentService
      .createTextOutput(JSON.stringify(diagnosePilgrimStats_(), null, 2))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var viewMap = {
    'kpi': 'KpiView',
    'summary': 'SummaryView',
    'b2c-courtesy': 'B2cCourtesyView',
    'countries': 'CountriesView',
    'nationalities': 'NationalitiesView',
    'camps': 'CampsView'
  };
  var templateName = viewMap[view] || 'IndexView';
  var template = HtmlService.createTemplateFromFile(templateName);
  template.baseUrl = ScriptApp.getService().getUrl();

  var titles = {
    'kpi': 'إحصائيات الحجاج — نظرة عامة',
    'summary': 'ملخص العمليات — الباقات والتأشيرات والمرشدين',
    'b2c-courtesy': 'حجاج الأفراد والمجاملة — حسب الدولة',
    'countries': 'توزيع الحجاج حسب بلد الإقامة',
    'nationalities': 'توزيع الحجاج حسب الجنسية',
    'camps': 'المخيمات — B2B/B2C — الجنس'
  };
  return template.evaluate()
    .setTitle(titles[view] || 'إحصائيات الحجاج — إكرام الضيف')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPilgrimStats() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(CONFIG.CACHE_KEY);
    if (cached) { return cached; }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return JSON.stringify({ success: true, data: buildEmptyData_() });
    }
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // بناء فهارس الأعمدة — مع fallback بالفهرس المباشر
    var ci = {};
    for (var k in COL) { ci[k] = headers.indexOf(COL[k]); }
    // fallback إذا لم يُطابَق الاسم
    if (ci.CATEGORY < 0) ci.CATEGORY = 3;   // عمود D
    if (ci.GENDER < 0) ci.GENDER = 4;       // عمود E
    if (ci.RESIDENCE < 0) ci.RESIDENCE = 16; // عمود Q
    if (ci.NATIONALITY < 0) ci.NATIONALITY = 17; // عمود R
    if (ci.FLIGHT_TYPE < 0) ci.FLIGHT_TYPE = 20; // عمود U
    if (ci.CAMP < 0) ci.CAMP = 27;          // عمود AB

    // المتغيرات
    var total = 0;
    var gM = 0, gF = 0;
    var b2bM = 0, b2bF = 0, b2bT = 0;
    var b2cM = 0, b2cF = 0, b2cT = 0;
    var courM = 0, courF = 0, courT = 0;
    var indM = 0, indF = 0, indT = 0;  // أفراد (individuals)

    var byRes = {}, byNat = {}, byCamp = {};
    var b2bByCo = {}, b2cByCo = {};
    var campGender = {}; // { campName: { male, female, total } }
    var b2cByCoMF = {};  // B2C حسب الدولة مع ذكور/إناث
    var courByCo = {};   // المجاملة حسب الدولة مع ذكور/إناث
    var indByCo = {};    // الأفراد حسب الدولة مع ذكور/إناث

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      // تجاهل الصفوف الفارغة (لا اسم في عمود A)
      if (!row[0] || String(row[0]).trim() === '') continue;
      total++;
      var g = String(row[ci.GENDER] || '').trim();
      var isMale = (g === 'ذكر') ? 1 : 0;
      var isFemale = (g === 'انثى') ? 1 : 0;
      gM += isMale; gF += isFemale;

      // فئة الحجاج — أفراد أو مجاملة
      var cat = String(row[ci.CATEGORY] || '').trim();
      if (cat === 'مجاملة') {
        courT++; courM += isMale; courF += isFemale;
      } else {
        indT++; indM += isMale; indF += isFemale;
      }

      // نوع الطيران
      var ft = String(row[ci.FLIGHT_TYPE] || '').trim().toUpperCase();
      if (ft === 'B2B') { b2bT++; b2bM += isMale; b2bF += isFemale; }
      else if (ft === 'B2C') { b2cT++; b2cM += isMale; b2cF += isFemale; }

      // المخيمات
      var camp = String(row[ci.CAMP] || '').trim();
      if (camp) {
        byCamp[camp] = (byCamp[camp] || 0) + 1;
        if (!campGender[camp]) { campGender[camp] = { male: 0, female: 0, total: 0 }; }
        campGender[camp].male += isMale;
        campGender[camp].female += isFemale;
        campGender[camp].total++;
      }

      // الإقامة والجنسية
      var res = String(row[ci.RESIDENCE] || '').trim();
      if (res) {
        byRes[res] = (byRes[res] || 0) + 1;
        if (ft === 'B2B') { b2bByCo[res] = (b2bByCo[res] || 0) + 1; }
        else if (ft === 'B2C') {
          b2cByCo[res] = (b2cByCo[res] || 0) + 1;
          if (!b2cByCoMF[res]) b2cByCoMF[res] = { male: 0, female: 0, total: 0 };
          b2cByCoMF[res].male += isMale; b2cByCoMF[res].female += isFemale; b2cByCoMF[res].total++;
        }
        if (cat === 'مجاملة') {
          if (!courByCo[res]) courByCo[res] = { male: 0, female: 0, total: 0 };
          courByCo[res].male += isMale; courByCo[res].female += isFemale; courByCo[res].total++;
        } else {
          if (!indByCo[res]) indByCo[res] = { male: 0, female: 0, total: 0 };
          indByCo[res].male += isMale; indByCo[res].female += isFemale; indByCo[res].total++;
        }
      }

      var nat = String(row[ci.NATIONALITY] || '').trim();
      if (nat) { byNat[nat] = (byNat[nat] || 0) + 1; }
    }

    // تحويل campGender لمصفوفة مرتبة
    var campsMF = [];
    for (var cn in campGender) {
      campsMF.push({ name: cn, male: campGender[cn].male, female: campGender[cn].female, total: campGender[cn].total });
    }
    campsMF.sort(function(a, b) { return b.total - a.total; });

    var result = {
      success: true,
      data: {
        total: total,
        gender: { male: gM, female: gF },
        flightType: { b2b: b2bT, b2c: b2cT },
        individualsMF: { total: indT, male: indM, female: indF },
        courtesyMF: { total: courT, male: courM, female: courF },
        b2bMF: { total: b2bT, male: b2bM, female: b2bF },
        b2cMF: { total: b2cT, male: b2cM, female: b2cF },
        campsMF: campsMF,
        uniqueNationalities: Object.keys(byNat).length,
        uniqueResidences: Object.keys(byRes).length,
        byResidence: toSortedList_(byRes, total),
        byNationality: toSortedList_(byNat, total),
        byCamp: toSortedList_(byCamp, total),
        b2bByCountry: toSortedList_(b2bByCo, b2bT),
        b2cByCountry: toSortedList_(b2cByCo, b2cT),
        b2cByCountryMF: toSortedListMF_(b2cByCoMF),
        courtesyByCountryMF: toSortedListMF_(courByCo),
        individualsByCountryMF: toSortedListMF_(indByCo),
        timestamp: new Date().toISOString()
      }
    };

    var json = JSON.stringify(result);
    try { cache.put(CONFIG.CACHE_KEY, json, CONFIG.CACHE_DURATION); } catch(ce) {}
    return json;
  } catch (err) {
    return JSON.stringify({ success: false, error: err.message || err.toString() });
  }
}

function toSortedList_(obj, total) {
  var arr = [];
  for (var k in obj) {
    arr.push({ name: k, count: obj[k], pct: total > 0 ? Math.round(obj[k] / total * 1000) / 10 : 0 });
  }
  arr.sort(function(a, b) { return b.count - a.count; });
  return arr;
}

function toSortedListMF_(obj) {
  var arr = [];
  for (var k in obj) {
    arr.push({ name: k, total: obj[k].total, male: obj[k].male, female: obj[k].female });
  }
  arr.sort(function(a, b) { return b.total - a.total; });
  return arr;
}

function buildEmptyData_() {
  var z = { total: 0, male: 0, female: 0 };
  return {
    total: 0, gender: { male: 0, female: 0 }, flightType: { b2b: 0, b2c: 0 },
    individualsMF: z, courtesyMF: z, b2bMF: z, b2cMF: z, campsMF: [],
    uniqueNationalities: 0, uniqueResidences: 0,
    byResidence: [], byNationality: [], byCamp: [], b2bByCountry: [], b2cByCountry: [],
    b2cByCountryMF: [], courtesyByCountryMF: [], individualsByCountryMF: [],
    timestamp: new Date().toISOString()
  };
}

// ══════════════════════════════════════════════════════════════
// ملخص العمليات — الباقات + التأشيرات + المرشدين
// ══════════════════════════════════════════════════════════════

function getSummaryStats() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // ── الباقات ──
    var pkgSheet = ss.getSheetByName('الباقات');
    var pkgCats = {}, totalPkgs = 0, soldPkgs = 0;
    if (pkgSheet) {
      var pkgLast = pkgSheet.getLastRow();
      if (pkgLast >= 3) {
        var pkgData = pkgSheet.getRange(3, 1, pkgLast - 2, 68).getValues();
        for (var i = 0; i < pkgData.length; i++) {
          if (!pkgData[i][1]) continue;
          totalPkgs++;
          var cat = mapCat_(String(pkgData[i][3]).trim());
          var capacity = Number(pkgData[i][10]) || 0;
          var sales = Number(pkgData[i][57]) || 0;
          if (sales >= capacity && capacity > 0) soldPkgs++;
          if (!pkgCats[cat]) pkgCats[cat] = { count: 0, pilgrims: 0 };
          pkgCats[cat].count++;
          pkgCats[cat].pilgrims += sales;
        }
      }
    }
    var catOrder = ['اقتصادية', 'مميزة', 'مميزة انتقالية', 'فاخرة', 'فاخرة انتقالية'];
    var pkgCatArr = [];
    for (var ci2 = 0; ci2 < catOrder.length; ci2++) {
      if (pkgCats[catOrder[ci2]]) pkgCatArr.push({ name: catOrder[ci2], count: pkgCats[catOrder[ci2]].count, pilgrims: pkgCats[catOrder[ci2]].pilgrims });
    }

    // ── الحجاج — التأشيرات + الجنس ──
    var pSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var afrad = { total: 0, male: 0, female: 0 };
    var mojamala = { total: 0, male: 0, female: 0 };
    var visaStats = {}, campStats = {};
    var b2b = 0, b2c = 0;
    var countries = {};

    if (pSheet) {
      var pLast = pSheet.getLastRow();
      if (pLast >= 2) {
        var pLastCol = pSheet.getLastColumn();
        var pHeaders = pSheet.getRange(1, 1, 1, pLastCol).getValues()[0];
        var pRows = pSheet.getRange(2, 1, pLast - 1, pLastCol).getValues();

        var idx = {};
        idx.cat = pHeaders.indexOf('فئة الحجاج'); if (idx.cat < 0) idx.cat = 3;
        idx.gender = pHeaders.indexOf('الجنس'); if (idx.gender < 0) idx.gender = 4;
        idx.visa = pHeaders.indexOf('حالة التأشيرة'); if (idx.visa < 0) idx.visa = 22;
        idx.camp = pHeaders.indexOf('المخيم'); if (idx.camp < 0) idx.camp = 27;
        idx.flight = pHeaders.indexOf('نوع عقد الطيران'); if (idx.flight < 0) idx.flight = 20;
        idx.country = pHeaders.indexOf('بلد الإقامة'); if (idx.country < 0) idx.country = 16;

        for (var j = 0; j < pRows.length; j++) {
          var r = pRows[j];
          if (!r[0] || String(r[0]).trim() === '') continue;
          var g = String(r[idx.gender] || '').trim();
          var iM = g === 'ذكر' ? 1 : 0;
          var iF = g === 'انثى' ? 1 : 0;
          var ct = String(r[idx.cat] || '').trim();

          if (ct === 'مجاملة') { mojamala.total++; mojamala.male += iM; mojamala.female += iF; }
          else { afrad.total++; afrad.male += iM; afrad.female += iF; }

          var v = String(r[idx.visa] || '').trim();
          if (v) visaStats[v] = (visaStats[v] || 0) + 1;

          var cm = String(r[idx.camp] || '').trim();
          if (cm) campStats[cm] = (campStats[cm] || 0) + 1;

          var ft = String(r[idx.flight] || '').trim().toUpperCase();
          if (ft === 'B2B') b2b++;
          else if (ft === 'B2C') b2c++;

          var co = String(r[idx.country] || '').trim();
          if (co) countries[co] = (countries[co] || 0) + 1;
        }
      }
    }

    // ── ترتيب ──
    var visaArr = [];
    for (var vk in visaStats) visaArr.push({ status: vk, count: visaStats[vk] });
    visaArr.sort(function(a, b2) { return b2.count - a.count; });

    var campArr = [];
    for (var ck in campStats) campArr.push({ name: ck, count: campStats[ck] });
    campArr.sort(function(a, b3) { return b3.count - a.count; });

    // الدول — أعلى 10 مع أكواد الأعلام
    var coArr = [];
    for (var cok in countries) coArr.push({ name: cok, code: countryCode_(cok), count: countries[cok] });
    coArr.sort(function(a, b4) { return b4.count - a.count; });
    coArr = coArr.slice(0, 10);

    return JSON.stringify({
      success: true,
      data: {
        totalPilgrims: 6800,
        totalGuides: 200,
        totalPackages: totalPkgs,
        soldPackages: soldPkgs,
        remainingPackages: totalPkgs - soldPkgs,
        packageCategories: pkgCatArr,
        afrad: afrad,
        mojamala: mojamala,
        visaStats: visaArr,
        campStats: campArr,
        flightStats: { B2B: b2b, B2C: b2c },
        countries: coArr,
        totalCountries: coArr.length,
        timestamp: new Date().toISOString()
      }
    });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.message || err.toString() });
  }
}

function countryCode_(name) {
  var map = {
    'الولايات المتحدة الأمريكية': 'us', 'فرنسا': 'fr', 'ألمانيا': 'de',
    'المملكة المتحدة': 'gb', 'إيطاليا': 'it', 'أسبانيا': 'es',
    'كندا': 'ca', 'السويد': 'se', 'بلغاريا': 'bg', 'جنوب افريقيا': 'za',
    'هولندا': 'nl', 'النمسا': 'at', 'سويسرا': 'ch', 'الأردن': 'jo',
    'النرويج': 'no', 'صربيا': 'rs', 'بلجيكا': 'be', 'البرازيل': 'br',
    'أستراليا': 'au', 'اليونان': 'gr', 'فنلندا': 'fi', 'إندونيسيا': 'id',
    'تركيا': 'tr', 'مصر': 'eg', 'تونس': 'tu', 'المغرب': 'ma',
    'العراق': 'iq', 'الدنمارك': 'dk', 'رومانيا': 'ro', 'البرتغال': 'pt',
    'ايرلندا': 'ie', 'نيوزيلاندا': 'nz', 'بولندا': 'pl', 'كرواتيا': 'hr',
    'ماليزيا': 'my', 'لبنان': 'lb', 'ليبيا': 'ly', 'فلسطين': 'ps'
  };
  return map[name] || '';
}

function mapCat_(cat) {
  var m = { 'Standard': 'اقتصادية', 'Economy': 'اقتصادية', 'Premium': 'مميزة', 'Premium Shifting': 'مميزة انتقالية', 'Luxury': 'فاخرة', 'Luxury Shifting': 'فاخرة انتقالية' };
  return m[cat] || cat || 'أخرى';
}

// ══════════════════════════════════════════════════════════════
// 🔬 دالة تشخيصية مؤقتة — لكشف سبب تضخم الإجمالي
// ══════════════════════════════════════════════════════════════
function diagnosePilgrimStats_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var report = {
    sheetName: CONFIG.SHEET_NAME,
    sheetId: CONFIG.SPREADSHEET_ID,
    timestamp: new Date().toISOString()
  };

  // 1) أبعاد الشيت — مصادر مختلفة
  report.dimensions = {
    getLastRow: sheet.getLastRow(),
    getLastColumn: sheet.getLastColumn(),
    getMaxRows: sheet.getMaxRows(),
    getMaxColumns: sheet.getMaxColumns(),
    getDataRange_numRows: sheet.getDataRange().getNumRows(),
    getDataRange_numCols: sheet.getDataRange().getNumColumns()
  };

  // 2) الرؤوس (أول 30 عمود)
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  report.headers = {};
  for (var h = 0; h < Math.min(headers.length, 30); h++) {
    report.headers['col_' + (h+1) + '_' + columnLetter_(h+1)] = String(headers[h]);
  }

  // 3) قراءة كل البيانات
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  report.totalRowsRead = data.length;

  // 4) فحص العمود A — كم صف فيه قيمة، كم فارغ، كم formula
  var colA_filled = 0, colA_empty = 0;
  var colA_samples_filled = [], colA_samples_empty = [];
  for (var i = 0; i < data.length; i++) {
    var v = data[i][0];
    if (v === '' || v === null || v === undefined || String(v).trim() === '') {
      colA_empty++;
      if (colA_samples_empty.length < 3) colA_samples_empty.push({ row: i+2, allRowSample: sliceRow_(data[i], 6) });
    } else {
      colA_filled++;
      if (colA_samples_filled.length < 3) colA_samples_filled.push({ row: i+2, value: String(v) });
    }
  }
  report.columnA = {
    name: String(headers[0]),
    filled: colA_filled,
    empty: colA_empty,
    samplesFilled: colA_samples_filled,
    samplesEmpty: colA_samples_empty
  };

  // 5) محاولة إيجاد عمود رقم الجواز
  var passportCol = -1;
  var passportNames = ['رقم الجواز', 'جواز السفر', 'Passport', 'Passport No', 'رقم جواز السفر'];
  for (var pn = 0; pn < passportNames.length; pn++) {
    var idx = headers.indexOf(passportNames[pn]);
    if (idx >= 0) { passportCol = idx; break; }
  }
  report.passportColumn = passportCol >= 0
    ? { found: true, index: passportCol, header: String(headers[passportCol]) }
    : { found: false, hint: 'لم يُعثر على عمود الجواز بأحد الأسماء المتوقعة' };

  // 6) إن وُجد عمود جواز — تحليل التكرار
  if (passportCol >= 0) {
    var passportCount = {};
    for (var j = 0; j < data.length; j++) {
      var p = String(data[j][passportCol] || '').trim();
      if (!p) continue;
      passportCount[p] = (passportCount[p] || 0) + 1;
    }
    var distribution = { '1x': 0, '2x': 0, '3x': 0, '4x': 0, '5x+': 0 };
    var dupSamples = [];
    for (var pk in passportCount) {
      var c = passportCount[pk];
      if (c === 1) distribution['1x']++;
      else if (c === 2) distribution['2x']++;
      else if (c === 3) distribution['3x']++;
      else if (c === 4) distribution['4x']++;
      else distribution['5x+']++;
      if (c > 1 && dupSamples.length < 5) dupSamples.push({ passport: pk, occurrences: c });
    }
    report.duplication = {
      uniquePassports: Object.keys(passportCount).length,
      distribution: distribution,
      duplicateSamples: dupSamples
    };
  }

  // 7) فحص عمود فئة الحجاج (D) — قيم مميزة
  var catValues = {};
  for (var k = 0; k < data.length; k++) {
    var cv = String(data[k][3] || '').trim() || '(فارغ)';
    catValues[cv] = (catValues[cv] || 0) + 1;
  }
  report.categoryColumnD = catValues;

  // 8) فحص عمود نوع الطيران (U / index 20)
  var ftValues = {};
  for (var m = 0; m < data.length; m++) {
    var fv = String(data[m][20] || '').trim().toUpperCase() || '(فارغ)';
    ftValues[fv] = (ftValues[fv] || 0) + 1;
  }
  report.flightTypeColumnU = ftValues;

  // 9) فحص عمود المخيم (AB / index 27)
  var campValues = {};
  for (var n = 0; n < data.length; n++) {
    var cmv = String(data[n][27] || '').trim() || '(فارغ)';
    campValues[cmv] = (campValues[cmv] || 0) + 1;
  }
  report.campColumnAB = campValues;

  return report;
}

function sliceRow_(row, n) {
  var out = [];
  for (var i = 0; i < Math.min(row.length, n); i++) {
    out.push(String(row[i] || '').substring(0, 30));
  }
  return out;
}

function columnLetter_(colNum) {
  var s = '';
  while (colNum > 0) {
    var m = (colNum - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    colNum = Math.floor((colNum - m) / 26);
  }
  return s;
}
