/**
 * فحص عميق للبيانات — كشف كل العيوب المنطقية
 */
function deepDataAudit() {
  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  // خرائط الأعمدة
  var cols = {};
  for (var i = 0; i < headers.length; i++) {
    cols[String(headers[i] || '').trim()] = i;
  }

  var SA_AIRPORTS = ['JED', 'MED', 'RUH', 'DMM', 'YNB', 'AHB', 'TIF'];

  var issues = {
    outbound_leg1_empty_leg2_filled: 0,
    return_leg_ends_in_sa: 0,
    return_leg1_empty_leg2_filled: 0,
    landing_time_empty: 0,
    takeoff_time_empty: 0,
    no_legs_at_all: 0,
    details: []
  };

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var rowNum = r + 1;

    // Outbound leg 1
    var out1From = String(row[cols['From 1']] || '').trim();
    var out1To = String(row[cols['To 1']] || '').trim();
    var out1Flight = String(row[cols['FlightNo1 (ذهاب)']] || '').trim();
    var out1TakeoffTime = row[cols['TAKEOFF TIME 1']];
    var out1LandingTime = row[cols['LANDING TIME 1']];

    // Outbound leg 2
    var out2From = String(row[cols['From 2']] || '').trim();
    var out2To = String(row[cols['To 2']] || '').trim();
    var out2Flight = String(row[cols['FlightNo 2 (ذهاب)']] || '').trim();
    var out2LandingTime = row[cols['LANDING TIME 2']];

    // Return leg 1 — يوجد عمودان اسمهما "From 1" و "To 1" — الثاني هو العودة
    // نتحقق بالموقع: الأعمدة AE-AK
    var retFlightCol = headers.indexOf('FlightNo1 (عودة)');
    var ret1Flight = retFlightCol !== -1 ? String(row[retFlightCol] || '').trim() : '';
    var ret1From = retFlightCol !== -1 ? String(row[retFlightCol + 3] || '').trim() : '';
    var ret1To = retFlightCol !== -1 ? String(row[retFlightCol + 4] || '').trim() : '';
    var ret1LandingTime = retFlightCol !== -1 ? row[retFlightCol + 6] : '';

    var ret2FlightCol = headers.indexOf('FlightNo 2 (عودة)');
    var ret2Flight = ret2FlightCol !== -1 ? String(row[ret2FlightCol] || '').trim() : '';
    var ret2From = ret2FlightCol !== -1 ? String(row[ret2FlightCol + 3] || '').trim() : '';
    var ret2To = ret2FlightCol !== -1 ? String(row[ret2FlightCol + 4] || '').trim() : '';

    var problems = [];

    // فحص 1: Outbound leg 1 فارغ مع leg 2 مملوء
    if (!out1Flight && !out1From && out2Flight) {
      issues.outbound_leg1_empty_leg2_filled++;
      problems.push('out1_empty_out2_filled: ذهاب leg 1 فارغ بينما leg 2 = ' + out2From + '→' + out2To);
    }

    // فحص 2: Return leg 1 فارغ مع leg 2 مملوء
    if (!ret1Flight && ret2Flight) {
      issues.return_leg1_empty_leg2_filled++;
      problems.push('ret1_empty_ret2_filled: عودة leg 1 فارغ مع leg 2 = ' + ret2From + '→' + ret2To);
    }

    // فحص 3: رحلة العودة تنتهي داخل السعودية
    var lastReturnTo = ret2To || ret1To;
    if (lastReturnTo && SA_AIRPORTS.indexOf(lastReturnTo) !== -1) {
      issues.return_leg_ends_in_sa++;
      problems.push('return_ends_in_sa: العودة تنتهي في ' + lastReturnTo + ' (داخل السعودية!)');
    }

    // فحص 4: رحلة leg 1 مُعلنة لكن LANDING TIME فارغ
    if (out1Flight && (!out1LandingTime || out1LandingTime === '')) {
      issues.landing_time_empty++;
      problems.push('out1_landing_empty: ذهاب leg 1 = ' + out1Flight + ' لكن LANDING TIME فارغ');
    }

    // فحص 5: لا رحلات أصلاً
    if (!out1Flight && !out2Flight && !ret1Flight && !ret2Flight) {
      issues.no_legs_at_all++;
      problems.push('no_legs: صف بلا أي رحلة');
    }

    if (problems.length > 0 && issues.details.length < 20) {
      issues.details.push({
        row: rowNum,
        pnr: String(row[cols['PNR']] || ''),
        name: String(row[cols['الاسم الأول (النظام)']] || '') + ' ' + String(row[cols['اسم العائلة (النظام)']] || ''),
        problems: problems
      });
    }
  }

  issues.totalRows = data.length - 1;
  return issues;
}
