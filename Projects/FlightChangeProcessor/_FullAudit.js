/**
 * فحص شامل لشيت V4 — كشف كل التناقضات
 */
function fullV4Audit() {
  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  var audit = {
    rowCount: data.length - 1,
    columnCount: headers.length,
    headers: headers.map(function(h, i) { return { index: i, name: String(h || '') }; }),
    anomalies: [],
    columnProfiles: []
  };

  // فحص كل عمود: ما نوع القيم المتوقعة vs الفعلية
  for (var c = 0; c < headers.length; c++) {
    var header = String(headers[c] || '').trim();
    if (!header) continue;

    var profile = {
      col: colLetter_(c),
      header: header,
      nonEmpty: 0,
      sampleValues: [],
      anomalyCount: 0,
      detectedTypes: { date: 0, time: 0, url: 0, text: 0, number: 0, empty: 0 }
    };

    var expectedType = expectedTypeOf_(header);
    profile.expected = expectedType;

    var anomalyRows = [];
    for (var r = 1; r < data.length; r++) {
      var val = data[r][c];
      if (val === null || val === undefined || val === '') {
        profile.detectedTypes.empty++;
        continue;
      }
      profile.nonEmpty++;

      var actual = detectType_(val);
      profile.detectedTypes[actual] = (profile.detectedTypes[actual] || 0) + 1;

      // كشف التناقض
      if (expectedType && !typeMatches_(expectedType, actual, val)) {
        profile.anomalyCount++;
        if (anomalyRows.length < 3) {
          anomalyRows.push({
            row: r + 1,
            value: String(val).substring(0, 80),
            detectedType: actual
          });
        }
      }

      if (profile.sampleValues.length < 3) {
        profile.sampleValues.push(String(val).substring(0, 60));
      }
    }

    if (profile.anomalyCount > 0) {
      audit.anomalies.push({
        col: profile.col,
        header: profile.header,
        expectedType: expectedType,
        anomalyCount: profile.anomalyCount,
        totalNonEmpty: profile.nonEmpty,
        examples: anomalyRows
      });
    }

    audit.columnProfiles.push(profile);
  }

  return audit;
}


function expectedTypeOf_(header) {
  var h = header.toLowerCase();
  if (/اسم|name/.test(h)) return 'text';
  if (/^date /.test(h) || /^تاريخ/.test(h)) return 'date';
  if (/time/.test(h) || /وقت/.test(h)) return 'time';
  if (/رابط|link|^url$|pdflink/.test(h)) return 'url';
  if (/pnr|booking id|incident|رقم التغيير|رقم الجواز|رقم الباقة/.test(h)) return 'text';
  if (/^to |^from |مطار|airport|city/.test(h)) return 'text';
  if (/flight|رحلة|flightno/.test(h)) return 'text';
  if (/حالة|state|status|نوع|type/.test(h)) return 'text';
  if (/ملاحظ|notes/.test(h)) return 'text';
  if (/تسلسلي|serial/.test(h)) return 'mixed';
  return null;
}


function detectType_(val) {
  var s = String(val);
  if (val instanceof Date) {
    // تمييز date vs time
    if (s.indexOf('1899') !== -1 || /^\d{1,2}:\d{2}/.test(s)) return 'time';
    return 'date';
  }
  if (/^https?:\/\//.test(s)) return 'url';
  if (/^\d{1,2}:\d{2}/.test(s)) return 'time';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return 'date';
  if (/^\d+(\.\d+)?$/.test(s)) return 'number';
  return 'text';
}


function typeMatches_(expected, actual, val) {
  if (expected === actual) return true;
  if (expected === 'date' && actual === 'time') return false;
  if (expected === 'time' && actual === 'url') return false;
  if (expected === 'time' && actual === 'text') {
    // قد تكون time بصيغة نصية مثل "12:15 PM"
    return /^\d{1,2}:\d{2}/.test(String(val)) || /^(0|1|2|3|4|5|6|7|8|9)/.test(String(val));
  }
  if (expected === 'date' && actual === 'text') return /\d{4}/.test(String(val));
  if (expected === 'mixed') return true;
  if (expected === 'text') return actual === 'text' || actual === 'number';
  if (expected === 'url' && actual === 'url') return true;
  return actual === expected;
}


function colLetter_(n) {
  var s = '';
  n++;
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
