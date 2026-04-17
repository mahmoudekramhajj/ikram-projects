/**
 * PilgrimMatcher.js — مطابقة الحاج + جلب الرحلات الحالية
 *
 * المصادر: Presonal Details + B2C + شيت الطيران
 * المطابقة: PNR → اسم أول + عائلة
 */


/**
 * تحميل كل البيانات المطلوبة
 */
function loadAllData_() {
  var ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
  var data = {};

  // Presonal Details
  var pdSheet = ss.getSheetByName(CONFIG.PD_SHEET);
  if (pdSheet && pdSheet.getLastRow() > 1) {
    data.pd = pdSheet.getRange(2, 1, pdSheet.getLastRow() - 1, 26).getValues();
    Logger.log('PD: ' + data.pd.length + ' rows');
  } else {
    data.pd = [];
  }

  // B2C
  var b2cSheet = ss.getSheetByName('B2C');
  if (b2cSheet && b2cSheet.getLastRow() > 1) {
    var b2cHeaders = b2cSheet.getRange(1, 1, 1, b2cSheet.getLastColumn()).getValues()[0];
    var b2cRows = b2cSheet.getRange(2, 1, b2cSheet.getLastRow() - 1, b2cSheet.getLastColumn()).getValues();
    data.b2c = { headers: b2cHeaders, rows: b2cRows };
    Logger.log('B2C: ' + b2cRows.length + ' rows');
  } else {
    data.b2c = { headers: [], rows: [] };
  }

  // شيت الطيران
  var flightSheet = ss.getSheetByName(CONFIG.FLIGHT_SHEET);
  if (flightSheet && flightSheet.getLastRow() > 1) {
    var fHeaders = flightSheet.getRange(1, 1, 1, flightSheet.getLastColumn()).getValues()[0];
    var fRows = flightSheet.getRange(2, 1, flightSheet.getLastRow() - 1, flightSheet.getLastColumn()).getValues();
    data.flights = { headers: fHeaders, rows: fRows };
    Logger.log('Flights: ' + fRows.length + ' rows');
  } else {
    data.flights = { headers: [], rows: [] };
  }

  return data;
}


/**
 * البحث عن الحاج بالاسم + PNR
 * يبحث في Presonal Details أولاً ثم B2C
 */
function findPilgrim_(firstName, lastName, pnr, allData, bookingId) {
  firstName = (firstName || '').toUpperCase().trim();
  lastName = (lastName || '').toUpperCase().trim();
  pnr = (pnr || '').toUpperCase().trim();

  var result = null;

  // === 1. البحث في Presonal Details بـ PNR ===
  if (pnr && allData.pd.length > 0) {
    var pnrCandidates = [];
    for (var i = 0; i < allData.pd.length; i++) {
      var contract = String(allData.pd[i][CONFIG.PD.CONTRACT_NAME] || '').toUpperCase();
      if (contract.indexOf(pnr) !== -1) {
        pnrCandidates.push(i);
      }
    }

    if (pnrCandidates.length > 0) {
      var match = matchName_(firstName, lastName, allData.pd, pnrCandidates, CONFIG.PD.FIRST_NAME_EN, CONFIG.PD.LAST_NAME_EN);
      if (match !== -1) {
        result = buildPDResult_(allData.pd[match]);
        result.source = 'PD';
      } else if (pnrCandidates.length === 1) {
        result = buildPDResult_(allData.pd[pnrCandidates[0]]);
        result.source = 'PD (PNR only)';
      }
    }
  }

  // === 2. البحث في B2C بـ PNR ===
  if (!result && pnr && allData.b2c.rows.length > 0) {
    var pnrCol = findCol_(allData.b2c.headers, 'Airline PNR');
    var fnCol = findCol_(allData.b2c.headers, 'الاسم الأول (الإنجليزية)');
    var lnCol = findCol_(allData.b2c.headers, 'اسم العائلة (الإنجليزية)');

    if (pnrCol !== -1) {
      var b2cCandidates = [];
      for (var i = 0; i < allData.b2c.rows.length; i++) {
        var rowPnr = String(allData.b2c.rows[i][pnrCol] || '').toUpperCase().trim();
        if (rowPnr === pnr) {
          b2cCandidates.push(i);
        }
      }

      if (b2cCandidates.length > 0 && fnCol !== -1 && lnCol !== -1) {
        var match = matchName_(firstName, lastName, allData.b2c.rows, b2cCandidates, fnCol, lnCol);
        if (match !== -1) {
          result = buildB2CResult_(allData.b2c.rows[match], allData.b2c.headers);
          result.source = 'B2C';
        } else if (b2cCandidates.length === 1) {
          result = buildB2CResult_(allData.b2c.rows[b2cCandidates[0]], allData.b2c.headers);
          result.source = 'B2C (PNR only)';
        }
      }
    }
  }

  // === 3. بحث بالاسم فقط في PD ===
  if (!result && firstName.length >= 2 && lastName.length >= 2 && allData.pd.length > 0) {
    var allIdx = [];
    for (var i = 0; i < allData.pd.length; i++) allIdx.push(i);
    var match = matchName_(firstName, lastName, allData.pd, allIdx, CONFIG.PD.FIRST_NAME_EN, CONFIG.PD.LAST_NAME_EN);
    if (match !== -1) {
      result = buildPDResult_(allData.pd[match]);
      result.source = 'PD (name only)';
    }
  }

  // === 4. جلب الرحلات الحالية من شيت الطيران ===
  if (result && pnr) {
    var cur = findCurrentFlights_(pnr, allData.flights);
    if (cur) {
      result.curOutFlight1 = cur.outFlight1;
      result.curOutDate1 = cur.outDate1;
      result.curOutFlight2 = cur.outFlight2;
      result.curOutDate2 = cur.outDate2;
      result.curRetFlight1 = cur.retFlight1;
      result.curRetDate1 = cur.retDate1;
      result.curRetFlight2 = cur.retFlight2;
      result.curRetDate2 = cur.retDate2;
    }
  }

  return result;
}


/**
 * الرحلات الحالية من شيت الطيران
 */
function findCurrentFlights_(pnr, flightsData) {
  if (!flightsData.rows.length) return null;

  var pnrCol = findCol_(flightsData.headers, 'PNR');
  if (pnrCol === -1) return null;

  for (var i = 0; i < flightsData.rows.length; i++) {
    var rowPnr = String(flightsData.rows[i][pnrCol] || '').toUpperCase().trim();
    if (rowPnr.indexOf(pnr) !== -1) {
      var r = flightsData.rows[i];
      return {
        outFlight1: String(r[21] || ''), outDate1: String(r[22] || ''),
        outFlight2: String(r[28] || ''), outDate2: String(r[29] || ''),
        retFlight1: String(r[35] || ''), retDate1: String(r[36] || ''),
        retFlight2: String(r[42] || ''), retDate2: String(r[43] || '')
      };
    }
  }
  return null;
}


/**
 * مطابقة اسم عامة
 */
function matchName_(firstName, lastName, data, candidates, fnCol, lnCol) {
  var noSpace = function(s) { return s.replace(/[\s\-'.]+/g, ''); };
  var fn = noSpace(firstName + lastName);
  var fnR = noSpace(lastName + firstName);

  for (var c = 0; c < candidates.length; c++) {
    var i = candidates[c];
    var sf = String(data[i][fnCol] || '').toUpperCase().trim();
    var sl = String(data[i][lnCol] || '').toUpperCase().trim();
    if (!sf && !sl) continue;

    if (sf === firstName && sl === lastName) return i;
    if (sf === lastName && sl === firstName) return i;

    var full = noSpace(sf + sl);
    if (full === fn || full === fnR) return i;

    if (sl === lastName && (sf.indexOf(firstName) !== -1 || firstName.indexOf(sf) !== -1)) return i;
    if (sl === firstName && (sf.indexOf(lastName) !== -1 || lastName.indexOf(sf) !== -1)) return i;
  }
  return -1;
}


/**
 * بناء نتيجة من PD
 */
function buildPDResult_(row) {
  return {
    sysFirstName: String(row[CONFIG.PD.FIRST_NAME_EN] || ''),
    sysLastName: String(row[CONFIG.PD.LAST_NAME_EN] || ''),
    serial: String(row[CONFIG.PD.SERIAL] || ''),
    passport: String(row[CONFIG.PD.PASSPORT] || ''),
    pkgNum: String(row[CONFIG.PD.PKG_NUM] || ''),
    pkgName: String(row[CONFIG.PD.PKG_NAME] || ''),
    flightType: String(row[CONFIG.PD.FLIGHT_TYPE] || ''),
    curOutFlight1: '', curOutDate1: '', curOutFlight2: '', curOutDate2: '',
    curRetFlight1: '', curRetDate1: '', curRetFlight2: '', curRetDate2: '',
    source: ''
  };
}


/**
 * بناء نتيجة من B2C
 */
function buildB2CResult_(row, headers) {
  var get = function(name) {
    var idx = findCol_(headers, name);
    return idx !== -1 ? String(row[idx] || '') : '';
  };
  // B2C: outbound flights at cols 32-45, return flights at cols 46-59
  var outFlightIdx = findCol_(headers, 'FlightNo1');
  var retFlightIdx = outFlightIdx !== -1 ? outFlightIdx + 14 : -1;
  var safeGet = function(i) { return i >= 0 && i < row.length ? String(row[i] || '') : ''; };
  return {
    sysFirstName: get('الاسم الأول (الإنجليزية)'),
    sysLastName: get('اسم العائلة (الإنجليزية)'),
    serial: get('الرقم التسلسلي'),
    passport: get('رقم جواز السفر'),
    pkgNum: get('رقم الباقة'),
    pkgName: get('اسم الباقة'),
    flightType: get('نوع عقد الطيران'),
    curOutFlight1: safeGet(outFlightIdx), curOutDate1: safeGet(outFlightIdx + 1),
    curOutFlight2: safeGet(outFlightIdx + 7), curOutDate2: safeGet(outFlightIdx + 8),
    curRetFlight1: safeGet(retFlightIdx), curRetDate1: safeGet(retFlightIdx + 1),
    curRetFlight2: safeGet(retFlightIdx + 7), curRetDate2: safeGet(retFlightIdx + 8),
    source: ''
  };
}


function findCol_(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === name) return i;
  }
  return -1;
}


