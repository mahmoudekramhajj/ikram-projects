/**
 * ═══════════════════════════════════════════════════════════════
 * MinaCampSearch.gs — بحث مخيم مِنى (Web App مستقل)
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0
 * ───────────────────────────────────────────────────────────────
 * المشروع: Mina Camp Search (مشروع Apps Script مستقل)
 * المشغّل: doGet → Web App URL
 * يقرأ: شيت "مخيم مني" من الجدول الرئيسي عبر openById
 * المخرج: صفحة بحث HTML للمشرفين الميدانيين
 * ═══════════════════════════════════════════════════════════════
 */

// ============================================================
// الإعدادات
// ============================================================

var CONFIG = {
  SS_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  SHEET_NAME: 'مخيم مني',
  CACHE_TTL: 300,    // 5 دقائق كاش
  MAX_RESULTS: 200,
  
  // أعمدة شيت مخيم مني (0-based)
  COLS: {
    TENT: 0,          // A: رقم الخيمة
    SOFA: 1,          // B: رقم الصوفا بيد
    APPLICANT_ID: 2,  // C: ApplicantId
    NAME: 3,          // D: اسم الحاج
    PASSPORT: 4,      // E: جواز السفر
    GENDER: 5,        // F: الجنس
    NATIONALITY: 6,   // G: الجنسية
    RESIDENCE: 7,     // H: دولة الإقامة
    CAMP: 8,          // I: المخيم
    GUIDE: 9,         // J: المرشد
    GROUP: 10         // K: رقم المجموعة
  }
};


// ============================================================
// Web App Entry Point
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('⛺ بحث مخيم مِنى')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}


// ============================================================
// جلب كل البيانات مرة واحدة (للفلترة المحلية في المتصفح)
// ============================================================

function getAllData() {
  var data = getMinaData_();
  return { success: data.length > 0, total: data.length, rows: data };
}


// ============================================================
// قراءة البيانات مع كاش
// ============================================================

function getMinaData_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('mina_data');
  
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  
  var ss = SpreadsheetApp.openById(CONFIG.SS_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  var numCols = Object.keys(CONFIG.COLS).length;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();
  
  var rows = [];
  var C = CONFIG.COLS;
  
  for (var i = 0; i < data.length; i++) {
    rows.push({
      t: String(data[i][C.TENT]),
      s: String(data[i][C.SOFA]),
      a: String(data[i][C.APPLICANT_ID]),
      n: String(data[i][C.NAME]),
      p: String(data[i][C.PASSPORT]),
      g: String(data[i][C.GENDER]),
      na: String(data[i][C.NATIONALITY]),
      r: String(data[i][C.RESIDENCE]),
      c: String(data[i][C.CAMP]),
      gu: String(data[i][C.GUIDE]),
      gr: String(data[i][C.GROUP])
    });
  }
  
  // محاولة تخزين كاش (حد 100KB)
  try {
    var json = JSON.stringify(rows);
    if (json.length < 95000) {
      cache.put('mina_data', json, CONFIG.CACHE_TTL);
    }
  } catch (e) {
    Logger.log('Cache too large, skipping');
  }
  
  return rows;
}


// ============================================================
// البحث — API للواجهة الأمامية
// ============================================================

function searchPilgrims(q) {
  try {
    var data = getMinaData_();
    
    if (!data.length) {
      return { success: false, message: 'شيت مخيم مني فارغ — شغّل التوزيع أولاً', results: [] };
    }
    
    var results = [];
    var fPassport = q.passport ? q.passport.toUpperCase() : '';
    var fName = q.name ? q.name.toLowerCase() : '';
    var fNationality = q.nationality || '';
    var fResidence = q.residence || '';
    var fCamp = q.camp || '';
    var fGuide = q.guide || '';
    var fTent = q.tent || '';
    var fApplicantId = q.applicantId || '';
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      if (fPassport && row.p.toUpperCase().indexOf(fPassport) === -1) continue;
      if (fName && row.n.toLowerCase().indexOf(fName) === -1) continue;
      if (fNationality && row.na !== fNationality) continue;
      if (fResidence && row.r !== fResidence) continue;
      if (fCamp && row.c !== fCamp) continue;
      if (fGuide && row.gu !== fGuide) continue;
      if (fTent && row.t.indexOf(fTent) === -1) continue;
      if (fApplicantId && row.a.indexOf(fApplicantId) === -1) continue;
      
      results.push({
        tent: row.t,
        sofa: row.s,
        applicantId: row.a,
        name: row.n,
        passport: row.p,
        gender: row.g,
        nationality: row.na,
        residence: row.r,
        camp: row.c,
        guide: row.gu,
        groupNumber: row.gr
      });
      
      if (results.length >= CONFIG.MAX_RESULTS) break;
    }
    
    // إحصائيات سريعة
    var tents = {}, camps = {};
    for (var j = 0; j < results.length; j++) {
      tents[results[j].tent] = true;
      camps[results[j].camp] = (camps[results[j].camp] || 0) + 1;
    }
    
    return {
      success: true,
      total: results.length,
      capped: results.length >= CONFIG.MAX_RESULTS,
      stats: { tents: Object.keys(tents).length, camps: camps },
      results: results
    };
    
  } catch (e) {
    Logger.log('searchPilgrims ERROR: ' + e.toString());
    return { success: false, message: 'خطأ في البحث: ' + e.message, results: [] };
  }
}


// ============================================================
// جلب القوائم المنسدلة
// ============================================================

function getFilterLists() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('mina_filters');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  
  var data = getMinaData_();
  var natSet = {}, resSet = {}, campSet = {}, guideSet = {}, tentSet = {};
  
  for (var i = 0; i < data.length; i++) {
    if (data[i].na) natSet[data[i].na] = true;
    if (data[i].r) resSet[data[i].r] = true;
    if (data[i].c) campSet[data[i].c] = true;
    if (data[i].gu) guideSet[data[i].gu] = true;
    if (data[i].t) tentSet[data[i].t] = true;
  }
  
  var result = {
    nationalities: Object.keys(natSet).sort(),
    residences: Object.keys(resSet).sort(),
    camps: Object.keys(campSet).sort(),
    guides: Object.keys(guideSet).sort(),
    totalPilgrims: data.length,
    totalTents: Object.keys(tentSet).length
  };
  
  try { cache.put('mina_filters', JSON.stringify(result), CONFIG.CACHE_TTL); } catch (e) {}
  return result;
}


// ============================================================
// تفاصيل خيمة كاملة
// ============================================================

function getTentDetails(tentNumber) {
  var data = getMinaData_();
  var pilgrims = [];
  var camp = '';
  
  for (var i = 0; i < data.length; i++) {
    if (data[i].t === tentNumber) {
      if (!camp) camp = data[i].c;
      pilgrims.push({
        sofa: data[i].s,
        name: data[i].n,
        passport: data[i].p,
        gender: data[i].g,
        nationality: data[i].na,
        residence: data[i].r,
        guide: data[i].gu,
        groupNumber: data[i].gr
      });
    }
  }
  
  return {
    success: pilgrims.length > 0,
    tent: tentNumber,
    camp: camp,
    count: pilgrims.length,
    pilgrims: pilgrims
  };
}