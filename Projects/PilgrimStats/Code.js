/**
 * PilgrimStats — لوحة إحصائيات الحجاج
 * شاشات عرض منفصلة للشاشات الكبيرة — فريق العمليات
 */

var CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'Presonal Details',
  CACHE_KEY: 'PILGRIM_STATS_V3',
  CACHE_DURATION: 300
};

var COURTESY_CAMP = 'مخيم  72ب';

// أسماء أعمدة Presonal Details
var COL_NAME = {
  GENDER: 'الجنس',
  RESIDENCE: 'بلد الإقامة',
  NATIONALITY: 'الجنسية',
  FLIGHT_TYPE: 'نوع عقد الطيران',
  CAMP: 'المخيم'
};

function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';
  var viewMap = {
    'kpi': 'KpiView',
    'countries': 'CountriesView',
    'nationalities': 'NationalitiesView',
    'camps': 'CampsView'
  };
  var templateName = viewMap[view] || 'IndexView';
  var template = HtmlService.createTemplateFromFile(templateName);
  template.baseUrl = ScriptApp.getService().getUrl();

  var titles = {
    'kpi': 'إحصائيات الحجاج — نظرة عامة',
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

    // بناء فهارس الأعمدة
    var ci = {};
    for (var k in COL_NAME) { ci[k] = headers.indexOf(COL_NAME[k]); }

    // المتغيرات
    var total = data.length;
    var gM = 0, gF = 0;
    var b2bM = 0, b2bF = 0, b2bT = 0;
    var b2cM = 0, b2cF = 0, b2cT = 0;
    var courM = 0, courF = 0, courT = 0;
    var nonCM = 0, nonCF = 0, nonCT = 0;

    var byRes = {}, byNat = {}, byCamp = {};
    var b2bByCo = {}, b2cByCo = {};
    var campGender = {}; // { campName: [male, female, total] }

    for (var i = 0; i < total; i++) {
      var row = data[i];
      var g = String(row[ci.GENDER] || '').trim();
      var m = (g === 'ذكر') ? 1 : 0;
      var f = (g === 'انثى') ? 1 : 0;
      gM += m; gF += f;

      var ft = String(row[ci.FLIGHT_TYPE] || '').trim().toUpperCase();
      if (ft === 'B2B') { b2bT++; b2bM += m; b2bF += f; }
      else if (ft === 'B2C') { b2cT++; b2cM += m; b2cF += f; }

      var camp = String(row[ci.CAMP] || '').trim();
      if (camp) {
        byCamp[camp] = (byCamp[camp] || 0) + 1;
        if (!campGender[camp]) { campGender[camp] = [0, 0, 0]; }
        campGender[camp][0] += m;
        campGender[camp][1] += f;
        campGender[camp][2]++;
        if (camp === COURTESY_CAMP) { courT++; courM += m; courF += f; }
        else { nonCT++; nonCM += m; nonCF += f; }
      }

      var res = String(row[ci.RESIDENCE] || '').trim();
      if (res) {
        byRes[res] = (byRes[res] || 0) + 1;
        if (ft === 'B2B') { b2bByCo[res] = (b2bByCo[res] || 0) + 1; }
        else if (ft === 'B2C') { b2cByCo[res] = (b2cByCo[res] || 0) + 1; }
      }

      var nat = String(row[ci.NATIONALITY] || '').trim();
      if (nat) { byNat[nat] = (byNat[nat] || 0) + 1; }
    }

    // تحويل campGender لمصفوفة
    var campsMF = [];
    for (var cn in campGender) {
      campsMF.push({ name: cn, male: campGender[cn][0], female: campGender[cn][1], total: campGender[cn][2] });
    }
    campsMF.sort(function(a, b) { return b.total - a.total; });

    var result = {
      success: true,
      data: {
        total: total,
        gender: { male: gM, female: gF },
        flightType: { b2b: b2bT, b2c: b2cT },
        courtesy: courT,
        allMF: { total: total, male: gM, female: gF },
        nonCourtesyMF: { total: nonCT, male: nonCM, female: nonCF },
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

function buildEmptyData_() {
  var z = { total: 0, male: 0, female: 0 };
  return {
    total: 0, gender: { male: 0, female: 0 }, flightType: { b2b: 0, b2c: 0 }, courtesy: 0,
    allMF: z, nonCourtesyMF: z, courtesyMF: z, b2bMF: z, b2cMF: z, campsMF: [],
    uniqueNationalities: 0, uniqueResidences: 0,
    byResidence: [], byNationality: [], byCamp: [], b2bByCountry: [], b2cByCountry: [],
    timestamp: new Date().toISOString()
  };
}
