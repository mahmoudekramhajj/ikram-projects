/**
 * ══════════════════════════════════════════════════════════════
 * Ikram Hajj - Sales & Operations Report
 * الإصدار 1.1
 * ══════════════════════════════════════════════════════════════
 * 
 * Web App منفصل لتقارير المبيعات والإحصائيات
 * المصدر: شيت الطيران (Col 87: البيع, Col 88: المتبقي, Col 89: النسبة)
 * 
 * ══════════════════════════════════════════════════════════════
 */

var REPORT_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  FLIGHTS_SHEET: 'الطيران',
  SETTINGS_SHEET: 'الاعدادات',
  DATA_START: 3
};

// ══════════════════════════════════════
// Web App Entry Point
// ══════════════════════════════════════

function doGet(e) {
  return HtmlService.createTemplateFromFile('SalesIndex')
    .evaluate()
    .setTitle('Ikram Sales Report')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ══════════════════════════════════════
// Data Fetching
// ══════════════════════════════════════

function getSalesData() {
  try {
    var ss = SpreadsheetApp.openById(REPORT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(REPORT_CONFIG.FLIGHTS_SHEET);
    if (!sheet) return JSON.stringify({ success: false, error: 'Sheet not found' });
    
    var lastRow = sheet.getLastRow();
    if (lastRow < REPORT_CONFIG.DATA_START) return JSON.stringify({ success: false, error: 'No data' });
    
    var numRows = lastRow - REPORT_CONFIG.DATA_START + 1;
    var data = sheet.getRange(REPORT_CONFIG.DATA_START, 1, numRows, 92).getValues();
    
    var flights = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // قراءة كل صف يحتوي على أي بيانات مفيدة
      // نتحقق من PNR أو Country أو PAX
      var pnr = row[1];          // Col B
      var country = row[4];       // Col E
      var pax = toNum(row[7]);    // Col H
      
      var hasPNR = pnr && String(pnr).trim() !== '';
      var hasCountry = country && String(country).trim() !== '';
      var hasPax = pax > 0;
      
      // تخطي فقط الصفوف الفارغة تماماً
      if (!hasPNR && !hasCountry && !hasPax) continue;
      
      var sold = toNum(row[86]);       // Col 87 (index 86) - البيع
      var remaining = toNum(row[87]);  // Col 88 (index 87) - المتبقي
      var pctRaw = row[88];            // Col 89 (index 88) - النسبة المئوية
      
      // إذا المتبقي فارغ، نحسبه من PAX - Sold
      if (remaining === 0 && pax > 0 && sold >= 0) {
        remaining = pax - sold;
      }
      
      // Parse percentage
      var salesPct = 0;
      if (pctRaw) {
        salesPct = parseFloat(pctRaw);
        if (!isNaN(salesPct) && salesPct > 0 && salesPct <= 1) salesPct = salesPct * 100;
        if (isNaN(salesPct)) salesPct = 0;
      }
      if (salesPct === 0 && pax > 0 && sold > 0) {
        salesPct = (sold / pax) * 100;
      }
      
      flights.push({
        no: row[0],
        pnr: clean(pnr) || ('ROW-' + (i + REPORT_CONFIG.DATA_START)),
        supplier: clean(row[2]),
        status: clean(row[3]),
        country: clean(row[4]),
        city: clean(row[5]),
        airline: clean(row[6]),
        pax: pax,
        sold: sold,
        remaining: remaining,
        salesPct: Math.round(salesPct * 10) / 10,
        contractNo: clean(row[90])
      });
    }
    
    var report = generateReport(flights);
    
    return JSON.stringify({
      success: true,
      flights: flights,
      report: report,
      totalRowsRead: numRows,
      flightsFound: flights.length,
      timestamp: new Date().toISOString()
    });
    
  } catch (error) {
    Logger.log('getSalesData ERROR: ' + error.toString());
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// ══════════════════════════════════════
// Report Generation
// ══════════════════════════════════════

function generateReport(flights) {
  var byCountry = {};
  var byAirline = {};
  var bySupplier = {};
  var totals = { pnrs: 0, pax: 0, sold: 0, remaining: 0 };
  
  for (var i = 0; i < flights.length; i++) {
    var f = flights[i];
    totals.pnrs++;
    totals.pax += f.pax;
    totals.sold += f.sold;
    totals.remaining += f.remaining;
    
    var c = f.country || 'Unknown';
    if (!byCountry[c]) byCountry[c] = { name: c, pnrs: 0, pax: 0, sold: 0, remaining: 0 };
    byCountry[c].pnrs++;
    byCountry[c].pax += f.pax;
    byCountry[c].sold += f.sold;
    byCountry[c].remaining += f.remaining;
    
    var a = f.airline || 'Unknown';
    if (!byAirline[a]) byAirline[a] = { name: a, pnrs: 0, pax: 0, sold: 0, remaining: 0 };
    byAirline[a].pnrs++;
    byAirline[a].pax += f.pax;
    byAirline[a].sold += f.sold;
    byAirline[a].remaining += f.remaining;
    
    var s = f.supplier || 'Unknown';
    if (!bySupplier[s]) bySupplier[s] = { name: s, pnrs: 0, pax: 0, sold: 0, remaining: 0, countries: [] };
    bySupplier[s].pnrs++;
    bySupplier[s].pax += f.pax;
    bySupplier[s].sold += f.sold;
    bySupplier[s].remaining += f.remaining;
    if (f.country && bySupplier[s].countries.indexOf(f.country) === -1) {
      bySupplier[s].countries.push(f.country);
    }
  }
  
  var countries = objToArray(byCountry).sort(function(a, b) { return b.pax - a.pax; });
  var airlines = objToArray(byAirline).sort(function(a, b) { return b.pax - a.pax; });
  var suppliers = objToArray(bySupplier).sort(function(a, b) { return b.pax - a.pax; });
  
  totals.salesPct = totals.pax > 0 ? Math.round((totals.sold / totals.pax) * 1000) / 10 : 0;
  
  return {
    totals: totals,
    countries: countries,
    airlines: airlines,
    suppliers: suppliers
  };
}

function objToArray(obj) {
  var arr = [];
  for (var key in obj) {
    var item = obj[key];
    item.salesPct = item.pax > 0 ? Math.round((item.sold / item.pax) * 1000) / 10 : 0;
    arr.push(item);
  }
  return arr;
}

// ══════════════════════════════════════
// Helpers
// ══════════════════════════════════════

function clean(val) {
  if (!val) return '';
  return String(val).trim();
}

function toNum(val) {
  if (!val && val !== 0) return 0;
  var n = Number(val);
  return isNaN(n) ? 0 : Math.round(n);
}

// ══════════════════════════════════════
// دالة تشخيص - شغلها مرة ثم افتح View > Execution log
// يمكنك حذفها بعد التأكد
// ══════════════════════════════════════

function debugFlightsData() {
  var ss = SpreadsheetApp.openById(REPORT_CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(REPORT_CONFIG.FLIGHTS_SHEET);
  
  Logger.log('=== SHEET INFO ===');
  Logger.log('Sheet name: ' + sheet.getName());
  Logger.log('Last Row: ' + sheet.getLastRow());
  Logger.log('Last Column: ' + sheet.getLastColumn());
  
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - REPORT_CONFIG.DATA_START + 1;
  var data = sheet.getRange(REPORT_CONFIG.DATA_START, 1, numRows, 92).getValues();
  
  var withPNR = 0;
  var emptyPNR = 0;
  var withCountry = 0;
  var withPax = 0;
  
  for (var i = 0; i < data.length; i++) {
    var pnr = data[i][1];
    var country = data[i][4];
    var pax = data[i][7];
    
    if (pnr && String(pnr).trim() !== '') withPNR++;
    else emptyPNR++;
    if (country && String(country).trim() !== '') withCountry++;
    if (pax && Number(pax) > 0) withPax++;
    
    if (i < 3 || i >= data.length - 3) {
      Logger.log('Row ' + (i + REPORT_CONFIG.DATA_START) + 
        ': PNR=' + String(pnr).substring(0, 20) + 
        ' | Country=' + country + 
        ' | PAX=' + pax + 
        ' | Sold(87)=' + data[i][86] + 
        ' | Remain(88)=' + data[i][87]);
    }
  }
  
  Logger.log('=== RESULTS ===');
  Logger.log('Total rows: ' + data.length);
  Logger.log('With PNR: ' + withPNR);
  Logger.log('Empty PNR: ' + emptyPNR);
  Logger.log('With Country: ' + withCountry);
  Logger.log('With PAX>0: ' + withPax);
  
  var result = JSON.parse(getSalesData());
  Logger.log('=== getSalesData ===');
  Logger.log('Flights found: ' + (result.flights ? result.flights.length : 0));
  if (result.report) {
    Logger.log('Totals: PNRs=' + result.report.totals.pnrs + 
      ' PAX=' + result.report.totals.pax + 
      ' Sold=' + result.report.totals.sold +
      ' Remaining=' + result.report.totals.remaining);
  }
}