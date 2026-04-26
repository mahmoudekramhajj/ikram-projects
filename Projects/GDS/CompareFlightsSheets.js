/**
 * CompareFlightsSheets.js — مقارنة PNR بين شيت All الخارجي وشيت الطيران عندنا
 *
 * خطوة 1: قراءة PNRs من كلا الشيتين
 * خطوة 2: مطابقة + تقرير غير المتطابق
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.CompareFlights = {
  // الشيت الخارجي
  ALL_SHEET_ID: '1d-beCOEjXuJnSs6kPPuoTw3GgfWVgCDP-kfBBqApH6s',
  ALL_SHEET_NAME: 'All',

  // الأعمدة في شيت All (حسب الرؤوس المكتشفة)
  ALL_COL_PNR: 24,              // X — PNR
  ALL_COL_PNR_NAME: 23,         // W — PNR NAME IN NUSUK
  ALL_COL_CONTRACT_NAME: 57,    // "CONTRACT NAME" (أخير)

  // الأعمدة في شيت الطيران عندنا
  FLIGHTS_COL_PNR: 2,           // B — PNR
  FLIGHTS_COL_CONTRACT: 91,     // CM — اسم العقد (حسب إشارة المستخدم)

  /**
   * قراءة هيكل شيت All + عرض عينة
   */
  diagnoseAll: function() {
    var ss = SpreadsheetApp.openById(GDS2.CompareFlights.ALL_SHEET_ID);
    var sheet = ss.getSheetByName(GDS2.CompareFlights.ALL_SHEET_NAME);
    if (!sheet) return { error: 'All sheet not found' };

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    // قراءة صف 1 (header)
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // إحصاء صفوف البيانات (PNR غير فارغ)
    var pnrCol = GDS2.CompareFlights.ALL_COL_PNR;
    var pnrData = lastRow >= 2 ? sheet.getRange(2, pnrCol, lastRow - 1, 1).getValues() : [];
    var dataCount = 0;
    var pnrs = [];
    for (var i = 0; i < pnrData.length; i++) {
      var p = String(pnrData[i][0] || '').trim();
      if (p) { dataCount++; pnrs.push(p); }
    }

    return {
      sheet: GDS2.CompareFlights.ALL_SHEET_NAME,
      total_rows: lastRow,
      data_rows: dataCount,
      cols: lastCol,
      header_at_pnr_col: headers[pnrCol - 1],
      header_at_pnr_name_col: headers[GDS2.CompareFlights.ALL_COL_PNR_NAME - 1],
      sample_pnrs: pnrs.slice(0, 10),
      duplicates: GDS2.CompareFlights._findDuplicates(pnrs)
    };
  },

  /**
   * قراءة هيكل شيت الطيران + عينة
   */
  diagnoseFlights: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var sheet = ss.getSheetByName(GDS2.Config.SHEET_FLIGHTS);
    if (!sheet) return { error: 'Flights sheet not found' };

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    // اقرأ صف 1 و2 (قد يكون هيكل رأسين)
    var row1 = sheet.getRange(1, 1, 1, Math.min(lastCol, 95)).getValues()[0];
    var row2 = sheet.getRange(2, 1, 1, Math.min(lastCol, 95)).getValues()[0];

    var pnrCol = GDS2.CompareFlights.FLIGHTS_COL_PNR;
    var cmCol = GDS2.CompareFlights.FLIGHTS_COL_CONTRACT;

    // بحث عن صفوف بـ PNR صالح (تخطّي رؤوس)
    var dataRange = sheet.getRange(1, pnrCol, lastRow, 1).getValues();
    var cmRange = sheet.getRange(1, cmCol, lastRow, 1).getValues();
    var pnrs = [];
    var contractNames = [];
    for (var i = 0; i < dataRange.length; i++) {
      var p = String(dataRange[i][0] || '').trim();
      var c = String(cmRange[i][0] || '').trim();
      // نعتبر الصف بيانات لو PNR موجود + ليس "PNR" (الرأس)
      if (p && p.toUpperCase() !== 'PNR' && p.length >= 4) {
        pnrs.push({ row: i + 1, pnr: p, contract_name: c });
        contractNames.push(c);
      }
    }

    return {
      sheet: GDS2.Config.SHEET_FLIGHTS,
      total_rows: lastRow,
      cols: lastCol,
      row1_at_pnr_col: row1[pnrCol - 1],
      row2_at_pnr_col: row2[pnrCol - 1],
      row1_at_cm_col: row1[cmCol - 1],
      row2_at_cm_col: row2[cmCol - 1],
      data_rows_found: pnrs.length,
      sample_pnrs: pnrs.slice(0, 5),
      sample_contract_names: contractNames.slice(0, 5).filter(function(c){return c;})
    };
  },

  /**
   * المقارنة الفعلية بين الشيتين.
   */
  compare: function() {
    var all = GDS2.CompareFlights.diagnoseAll();
    var flights = GDS2.CompareFlights.diagnoseFlights();

    if (all.error || flights.error) {
      return { error: all.error || flights.error };
    }

    // خريطة PNRs في الطيران
    var flightsPNRMap = {};
    var flightsContractMap = {};
    var fsheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_FLIGHTS);
    var lastRow = fsheet.getLastRow();
    var fpnrs = fsheet.getRange(1, GDS2.CompareFlights.FLIGHTS_COL_PNR, lastRow, 1).getValues();
    var fcm = fsheet.getRange(1, GDS2.CompareFlights.FLIGHTS_COL_CONTRACT, lastRow, 1).getValues();
    for (var i = 0; i < fpnrs.length; i++) {
      var p = String(fpnrs[i][0] || '').trim();
      var c = String(fcm[i][0] || '').trim();
      if (p && p.toUpperCase() !== 'PNR' && p.length >= 4) {
        flightsPNRMap[p] = { row: i + 1, contract_name: c };
        if (c) flightsContractMap[c] = { row: i + 1, pnr: p };
      }
    }

    // خريطة PNRs في All
    var allsheet = SpreadsheetApp.openById(GDS2.CompareFlights.ALL_SHEET_ID).getSheetByName(GDS2.CompareFlights.ALL_SHEET_NAME);
    var allLastRow = allsheet.getLastRow();
    var apnrs = allsheet.getRange(1, GDS2.CompareFlights.ALL_COL_PNR, allLastRow, 1).getValues();
    var apnrNames = allsheet.getRange(1, GDS2.CompareFlights.ALL_COL_PNR_NAME, allLastRow, 1).getValues();

    var allPNRMap = {};
    for (var j = 0; j < apnrs.length; j++) {
      var ap = String(apnrs[j][0] || '').trim();
      var apName = String(apnrNames[j][0] || '').trim();
      if (ap && ap.toUpperCase() !== 'PNR' && ap.length >= 4) {
        allPNRMap[ap] = { row: j + 1, pnr_name: apName };
      }
    }

    // مطابقة
    var matchedByPNR = [];
    var inAllNotFlights = [];
    var inFlightsNotAll = [];

    for (var ap1 in allPNRMap) {
      if (flightsPNRMap[ap1]) {
        matchedByPNR.push({
          pnr: ap1,
          all_row: allPNRMap[ap1].row,
          flights_row: flightsPNRMap[ap1].row,
          pnr_name: allPNRMap[ap1].pnr_name,
          contract_name: flightsPNRMap[ap1].contract_name
        });
      } else {
        inAllNotFlights.push({
          pnr: ap1,
          all_row: allPNRMap[ap1].row,
          pnr_name: allPNRMap[ap1].pnr_name
        });
      }
    }

    for (var fp1 in flightsPNRMap) {
      if (!allPNRMap[fp1]) {
        inFlightsNotAll.push({
          pnr: fp1,
          flights_row: flightsPNRMap[fp1].row,
          contract_name: flightsPNRMap[fp1].contract_name
        });
      }
    }

    return {
      all_count: Object.keys(allPNRMap).length,
      flights_count: Object.keys(flightsPNRMap).length,
      matched_count: matchedByPNR.length,
      in_all_not_flights: inAllNotFlights.length,
      in_flights_not_all: inFlightsNotAll.length,
      sample_matched: matchedByPNR.slice(0, 5),
      sample_in_all_not_flights: inAllNotFlights.slice(0, 10),
      sample_in_flights_not_all: inFlightsNotAll.slice(0, 10)
    };
  },

  _findDuplicates: function(arr) {
    var seen = {};
    var dups = [];
    for (var i = 0; i < arr.length; i++) {
      var v = arr[i];
      if (seen[v]) {
        if (dups.indexOf(v) === -1) dups.push(v);
      } else {
        seen[v] = true;
      }
    }
    return dups;
  }
};

// ==================== Global ====================

function diagnoseAllSheet() { return GDS2.CompareFlights.diagnoseAll(); }
function diagnoseFlightsForCompare() { return GDS2.CompareFlights.diagnoseFlights(); }
function comparePNRs() { return GDS2.CompareFlights.compare(); }

/**
 * عرض صف كامل من شيت All حسب رقم الصف
 */
function inspectAllRow(rowNum) {
  var ss = SpreadsheetApp.openById(GDS2.CompareFlights.ALL_SHEET_ID);
  var sheet = ss.getSheetByName(GDS2.CompareFlights.ALL_SHEET_NAME);
  if (!sheet) return { error: 'All not found' };
  var lastCol = sheet.getLastColumn();
  var row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var result = {};
  for (var i = 0; i < row.length; i++) {
    var h = headers[i] || 'col_' + (i + 1);
    var v = row[i];
    if (v !== '' && v !== null && v !== undefined) {
      result[h] = v;
    }
  }
  return result;
}

/**
 * جلب قطاعات الرحلات من كل صفوف شيت All (قراءة فقط).
 * يُرجع مصفوفة: [{row, pnr, segments:[{n, flightNo, date, from, to, timeD, timeA}]}]
 */
function dumpAllSegments() {
  var ss = SpreadsheetApp.openById(GDS2.CompareFlights.ALL_SHEET_ID);
  var sheet = ss.getSheetByName(GDS2.CompareFlights.ALL_SHEET_NAME);
  if (!sheet) return { error: 'All not found' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // الأعمدة 24=PNR، 34-57 = FlightNo1..TimeA4 (6 أعمدة × 4 قطاعات)
  // getDisplayValues لتفادي اختلاف التوقيت في Date objects
  var data = sheet.getRange(2, 1, lastRow - 1, 57).getDisplayValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var pnr = String(row[23] || '').trim();
    if (!pnr || pnr.toUpperCase() === 'PNR' || pnr.length < 4) continue;
    var status = String(row[18] || '').trim(); // col 19 in All = Status
    var segs = [];
    for (var seg = 0; seg < 4; seg++) {
      var base = 33 + seg * 6; // 34-1 = col index of FlightNo(seg+1)
      var flightNo = String(row[base] || '').trim();
      var date = row[base + 1];
      var from = String(row[base + 2] || '').trim();
      var to = String(row[base + 3] || '').trim();
      var timeD = row[base + 4];
      var timeA = row[base + 5];
      // إدراج القطاع إذا فيه أي قيمة
      if (flightNo || date || from || to || timeD !== '' || timeA !== '') {
        segs.push({
          n: seg + 1, flightNo: flightNo, date: date,
          from: from, to: to, timeD: timeD, timeA: timeA
        });
      }
    }
    out.push({ row: i + 2, pnr: pnr, status: status, segments: segs });
  }
  return out;
}

/**
 * جلب قطاعات الرحلات من كل صفوف شيت الطيران (قراءة فقط).
 * يُرجع مصفوفة: [{row, pnr, departure:[{...}], return:[{...}]}]
 * ذهاب: cols 22-35 (2 legs)، عودة: cols 36-49 (2 legs)
 */
function dumpFlightsSegments() {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
  var sheet = ss.getSheetByName(GDS2.Config.SHEET_FLIGHTS);
  if (!sheet) return { error: 'Flights not found' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  // getDisplayValues لتفادي اختلاف التوقيت في Date objects
  var data = sheet.getRange(1, 1, lastRow, 49).getDisplayValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var pnr = String(row[1] || '').trim();
    if (!pnr || pnr.toUpperCase() === 'PNR' || pnr.length < 4) continue;
    var status = String(row[3] || '').trim(); // col 4 = Status

    function readLeg(startCol) {
      // startCol هو 0-based: FlightNo, DateTakeoff, TimeTakeoff, From, To, DateLanding, TimeLanding
      return {
        flightNo:     String(row[startCol]     || '').trim(),
        dateTakeoff:  row[startCol + 1],
        timeTakeoff:  row[startCol + 2],
        from:         String(row[startCol + 3] || '').trim(),
        to:           String(row[startCol + 4] || '').trim(),
        dateLanding:  row[startCol + 5],
        timeLanding:  row[startCol + 6]
      };
    }
    // إرجاع كل الـlegs بترتيب ثابت حتى الفارغة منها
    var departure = [readLeg(21), readLeg(28)];
    var ret = [readLeg(35), readLeg(42)];
    out.push({ row: i + 1, pnr: pnr, status: status, departure: departure, return: ret });
  }
  return out;
}

/**
 * تطبيق تحديثات على شيت الطيران (كتابة).
 * updates: [{row:int, col:int, value:string}]
 * يُرجع: {written: N, errors: [...]}
 */
function applyFlightUpdates(updates) {
  if (!updates || !updates.length) return { written: 0, errors: [] };
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
  var sheet = ss.getSheetByName(GDS2.Config.SHEET_FLIGHTS);
  if (!sheet) return { error: 'Flights not found' };
  var written = 0;
  var errors = [];
  for (var i = 0; i < updates.length; i++) {
    var u = updates[i];
    try {
      sheet.getRange(u.row, u.col).setValue(u.value);
      written++;
    } catch (e) {
      errors.push({ row: u.row, col: u.col, error: e.message });
    }
  }
  return { written: written, errors: errors, total: updates.length };
}

/**
 * عرض صف كامل من شيت الطيران
 */
function inspectFlightsRow(rowNum) {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
  var sheet = ss.getSheetByName(GDS2.Config.SHEET_FLIGHTS);
  if (!sheet) return { error: 'Flights not found' };
  var lastCol = sheet.getLastColumn();
  var row = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  // row1 + row2 headers (merged structure)
  var row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var row2 = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var result = {};
  for (var i = 0; i < row.length; i++) {
    var h1 = String(row1[i] || '').trim();
    var h2 = String(row2[i] || '').trim();
    var header = h2 || h1 || 'col_' + (i + 1);
    if (h1 && h2 && h1 !== h2) header = h1 + ' / ' + h2;
    var v = row[i];
    if (v !== '' && v !== null && v !== undefined) {
      result['col_' + (i + 1) + '_' + header] = v;
    }
  }
  return result;
}
