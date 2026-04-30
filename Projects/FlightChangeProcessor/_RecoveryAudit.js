/**
 * _RecoveryAudit.js — جرد الضرر لخطة إنقاذ FCP (المرحلة 2 + 1.3)
 * أُضيف 2026-04-28.
 *
 * يفعل ثلاثة أشياء:
 *   1) يضع banner DEPRECATED على شيت V4 الموحّد.
 *   2) يفحص الـ 411 صف ويحدّد الصفوف المعطوبة (Gmail URLs/أنماط خاطئة في أعمدة وقت/تاريخ/IATA).
 *   3) يطابق إيميلات Gmail (TKT) مع صفوف V4 ويحدّد الإيميلات الضائعة (Processed بدون صف، أو غير معالج أصلاً).
 *
 * المخرجات: 3 تابات في spreadsheet V4 نفسه:
 *   - FCP-Recovery-Banner (يُحدَّث في الصف الأول للشيت الموحّد مباشرة، ليس tab)
 *   - FCP-Recovery-Corrupted (الصفوف المعطوبة)
 *   - FCP-Recovery-MissingEmails (الإيميلات بدون صف)
 *
 * استدعِ runRecoveryAudit للتنفيذ.
 */

var RECOVERY_AUDIT = {
  BANNER_TEXT: '⛔ DEPRECATED 2026-04-28 — هذا الشيت يحوي بيانات FCP v1 المهجور. لا تعتمد عليه. الاستبدال جاري عبر FCP v2.',
  CORRUPT_TAB: 'FCP-Recovery-Corrupted',
  MISSING_TAB: 'FCP-Recovery-MissingEmails',

  // أعمدة 1-based في الشيت الموحد V4
  COLS: {
    PNR: 2,
    BOOKING: 3,
    INCIDENT: 4,
    EMAIL_DATE: 5,
    FN_PDF: 6, LN_PDF: 7,
    // أوقات (TAKEOFF/LANDING TIME)
    TIMES: [19, 23, 26, 30, 33, 37, 40, 44],
    // تواريخ (Date TAKEOFF/LANDING)
    DATES: [18, 22, 25, 29, 32, 36, 39, 43],
    // مطارات (From/To — IATA 3 حروف)
    IATA: [20, 21, 27, 28, 34, 35, 41, 42],
    // أرقام رحلات
    FLIGHTS: [17, 24, 31, 38],
    // ملاحظات/مصدر/تحذيرات
    SOURCE: 50, NOTES: 51, WARN: 52,
    TOTAL: 54
  }
};

function runRecoveryAudit() {
  var t0 = new Date();
  Logger.log('=== FCP Recovery Audit — ' + t0.toISOString() + ' ===');

  var unified = getOrCreateUnifiedSheet_();
  var ss = unified.getParent();

  // 1) Banner على الشيت الموحّد
  applyDeprecatedBanner_(unified);

  // 2) فحص الصفوف المعطوبة
  var corrupt = findCorruptedRows_(unified);
  Logger.log('  📛 صفوف معطوبة: ' + corrupt.rows.length + ' من ' + corrupt.totalScanned);
  writeCorruptTab_(ss, corrupt);

  // 3) إيميلات ضائعة
  var missing = findMissingEmails_(unified);
  Logger.log('  📧 إيميلات ضائعة: ' + missing.length);
  writeMissingTab_(ss, missing);

  var duration = ((new Date() - t0) / 1000).toFixed(1) + 's';
  Logger.log('=== Done in ' + duration + ' ===');

  return {
    timestamp: t0.toISOString(),
    duration: duration,
    bannerApplied: true,
    totalRowsScanned: corrupt.totalScanned,
    corruptedRows: corrupt.rows.length,
    corruptionTypes: corrupt.summary,
    missingEmails: missing.length,
    missingBreakdown: countMissingByType_(missing),
    unifiedSpreadsheetUrl: ss.getUrl()
  };
}

/* ----------------------- 1) Banner ----------------------- */

function applyDeprecatedBanner_(sheet) {
  // أدرج صفاً علوياً جديداً (إن لم يكن موجوداً) ثم اكتب الـ banner.
  var firstCell = sheet.getRange(1, 1).getValue();
  if (String(firstCell).indexOf('DEPRECATED') !== -1) {
    Logger.log('  ✓ Banner موجود مسبقاً.');
    return;
  }
  sheet.insertRowBefore(1);
  var width = RECOVERY_AUDIT.COLS.TOTAL;
  var range = sheet.getRange(1, 1, 1, width);
  range.merge();
  range.setValue(RECOVERY_AUDIT.BANNER_TEXT);
  range.setBackground('#c0392b');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(2);  // banner + هيدر
  Logger.log('  ✓ Banner أُضيف.');
}

/* ----------------------- 2) Corrupted rows ----------------------- */

function findCorruptedRows_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.min(sheet.getLastColumn(), RECOVERY_AUDIT.COLS.TOTAL);
  // البيانات تبدأ من الصف 3 الآن لو banner أُدرج (banner=1, header=2, data=3+).
  // لو لم يُدرج banner لأي سبب، نبدأ من 2.
  var headerRow = isBannerPresent_(sheet) ? 2 : 1;
  var startData = headerRow + 1;
  if (lastRow < startData) return { rows: [], totalScanned: 0, summary: {} };

  var data = sheet.getRange(startData, 1, lastRow - startData + 1, lastCol).getValues();
  var iataRe = /^[A-Z]{3}$/;
  var timeRe = /^\d{1,2}:\d{2}$/;
  var dateRe = /^\d{4}-\d{2}-\d{2}$/;
  var flightRe = /^[A-Z0-9]{2}-?\d{1,5}$/i;
  var gmailRe = /mail\.google\.com|gmail/i;

  var corrupted = [];
  var summary = { gmail_in_time: 0, gmail_in_date: 0, gmail_in_iata: 0, bad_iata: 0, bad_time: 0, bad_date: 0, bad_flight: 0 };

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var issues = [];
    var sheetRowNum = startData + i;

    RECOVERY_AUDIT.COLS.TIMES.forEach(function(c) {
      var v = String(row[c - 1] || '').trim();
      if (!v) return;
      if (gmailRe.test(v)) { issues.push('gmail-in-time:col' + c); summary.gmail_in_time++; }
      else if (!timeRe.test(v) && v.length > 0) { issues.push('bad-time:col' + c + '=' + v.substring(0, 30)); summary.bad_time++; }
    });
    RECOVERY_AUDIT.COLS.DATES.forEach(function(c) {
      var v = String(row[c - 1] || '').trim();
      if (!v) return;
      if (gmailRe.test(v)) { issues.push('gmail-in-date:col' + c); summary.gmail_in_date++; }
      else if (!dateRe.test(v) && !(row[c - 1] instanceof Date)) {
        issues.push('bad-date:col' + c + '=' + v.substring(0, 30)); summary.bad_date++;
      }
    });
    RECOVERY_AUDIT.COLS.IATA.forEach(function(c) {
      var v = String(row[c - 1] || '').trim().toUpperCase();
      if (!v) return;
      if (gmailRe.test(v)) { issues.push('gmail-in-iata:col' + c); summary.gmail_in_iata++; }
      else if (!iataRe.test(v)) { issues.push('bad-iata:col' + c + '=' + v.substring(0, 30)); summary.bad_iata++; }
    });
    RECOVERY_AUDIT.COLS.FLIGHTS.forEach(function(c) {
      var v = String(row[c - 1] || '').trim();
      if (!v) return;
      if (!flightRe.test(v)) { issues.push('bad-flight:col' + c + '=' + v.substring(0, 30)); summary.bad_flight++; }
    });

    if (issues.length > 0) {
      corrupted.push({
        sheetRow: sheetRowNum,
        pnr: row[RECOVERY_AUDIT.COLS.PNR - 1],
        booking: row[RECOVERY_AUDIT.COLS.BOOKING - 1],
        incident: row[RECOVERY_AUDIT.COLS.INCIDENT - 1],
        emailDate: row[RECOVERY_AUDIT.COLS.EMAIL_DATE - 1],
        firstName: row[RECOVERY_AUDIT.COLS.FN_PDF - 1],
        lastName: row[RECOVERY_AUDIT.COLS.LN_PDF - 1],
        issueCount: issues.length,
        issues: issues.join(' | ')
      });
    }
  }

  return { rows: corrupted, totalScanned: data.length, summary: summary };
}

function isBannerPresent_(sheet) {
  var v = sheet.getRange(1, 1).getValue();
  return String(v).indexOf('DEPRECATED') !== -1;
}

function writeCorruptTab_(ss, corrupt) {
  var tab = ss.getSheetByName(RECOVERY_AUDIT.CORRUPT_TAB);
  if (!tab) tab = ss.insertSheet(RECOVERY_AUDIT.CORRUPT_TAB);
  tab.clear();
  var headers = ['Sheet Row', 'PNR', 'Booking ID', 'Incident #', 'Email Date',
                 'First Name', 'Last Name', '# Issues', 'Issues'];
  tab.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#c0392b').setFontColor('#fff');
  tab.setFrozenRows(1);
  if (corrupt.rows.length === 0) {
    tab.getRange(2, 1).setValue('✅ لا صفوف معطوبة وُجدت في هذا الفحص.');
    return;
  }
  var values = corrupt.rows.map(function(r) {
    return [r.sheetRow, r.pnr, r.booking, r.incident, r.emailDate, r.firstName, r.lastName, r.issueCount, r.issues];
  });
  tab.getRange(2, 1, values.length, headers.length).setValues(values);

  // ملخص في صف أخير
  var sumRow = values.length + 3;
  tab.getRange(sumRow, 1).setValue('ملخص أنواع الفساد:');
  tab.getRange(sumRow, 1).setFontWeight('bold');
  var idx = 1;
  Object.keys(corrupt.summary).forEach(function(k) {
    tab.getRange(sumRow + idx, 1).setValue(k);
    tab.getRange(sumRow + idx, 2).setValue(corrupt.summary[k]);
    idx++;
  });
}

/* ----------------------- 3) Missing emails ----------------------- */

function findMissingEmails_(unified) {
  // اجمع PNRs و Incidents الموجودة في الشيت الموحّد
  var lastRow = unified.getLastRow();
  var headerRow = isBannerPresent_(unified) ? 2 : 1;
  var startData = headerRow + 1;
  var rowPNRs = {}, rowIncidents = {};
  if (lastRow >= startData) {
    var data = unified.getRange(startData, 1, lastRow - startData + 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      var pnr = String(data[i][RECOVERY_AUDIT.COLS.PNR - 1] || '').trim();
      var inc = String(data[i][RECOVERY_AUDIT.COLS.INCIDENT - 1] || '').trim();
      if (pnr) rowPNRs[pnr] = true;
      if (inc) rowIncidents[inc] = true;
    }
  }

  var missing = [];

  // (أ) إيميلات Processed بدون صف في V4
  scanLabelForMissing_('label:' + CONFIG.PROCESSED_LABEL, 'processed_no_row', rowPNRs, rowIncidents, missing);
  // (ب) إيميلات بـ TKT لم يُعالج (لا processed ولا skipped)
  var unprocessedQuery = 'label:' + CONFIG.GMAIL_LABEL +
                         ' -label:' + CONFIG.PROCESSED_LABEL +
                         ' -label:' + CONFIG.SKIPPED_LABEL;
  scanLabelForMissing_(unprocessedQuery, 'unprocessed', rowPNRs, rowIncidents, missing);

  return missing;
}

function scanLabelForMissing_(query, type, rowPNRs, rowIncidents, missing) {
  var start = 0, batchSize = 100, scanned = 0;
  while (scanned < 1000) {
    var batch = GmailApp.search(query, start, batchSize);
    if (batch.length === 0) break;
    for (var i = 0; i < batch.length; i++) {
      var thr = batch[i];
      var first = thr.getMessages()[0];
      var subject = first.getSubject() || '';
      var body = first.getPlainBody().substring(0, 8000);
      var pnr = extractPNRFromBody_(body);
      var incident = extractIncidentNumber_(subject);
      var hasPNR = pnr && rowPNRs[pnr];
      var hasIncident = incident && rowIncidents[incident];
      if (type === 'processed_no_row' && !hasPNR && !hasIncident) {
        missing.push({
          type: type,
          threadId: thr.getId(),
          permalink: thr.getPermalink(),
          date: first.getDate(),
          subject: subject,
          pnr: pnr || '',
          incident: incident || ''
        });
      }
      if (type === 'unprocessed') {
        missing.push({
          type: type,
          threadId: thr.getId(),
          permalink: thr.getPermalink(),
          date: first.getDate(),
          subject: subject,
          pnr: pnr || '',
          incident: incident || ''
        });
      }
    }
    if (batch.length < batchSize) break;
    start += batchSize;
    scanned += batch.length;
  }
}

function writeMissingTab_(ss, missing) {
  var tab = ss.getSheetByName(RECOVERY_AUDIT.MISSING_TAB);
  if (!tab) tab = ss.insertSheet(RECOVERY_AUDIT.MISSING_TAB);
  tab.clear();
  var headers = ['Type', 'Thread ID', 'Permalink', 'Date', 'Subject', 'PNR', 'Incident'];
  tab.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#d35400').setFontColor('#fff');
  tab.setFrozenRows(1);
  if (missing.length === 0) {
    tab.getRange(2, 1).setValue('✅ لا إيميلات ضائعة.');
    return;
  }
  var values = missing.map(function(m) {
    return [m.type, m.threadId, m.permalink, m.date, m.subject, m.pnr, m.incident];
  });
  tab.getRange(2, 1, values.length, headers.length).setValues(values);
}

function countMissingByType_(missing) {
  var c = {};
  for (var i = 0; i < missing.length; i++) c[missing[i].type] = (c[missing[i].type] || 0) + 1;
  return c;
}

/**
 * inspectSampleRows — يقرأ 5 صفوف عينة من الشيت الموحد ويُرجع نوع كل خلية
 * + قيمتها — لتشخيص لماذا regex لا يطابق.
 */
function inspectSampleRows() {
  var sheet = getOrCreateUnifiedSheet_();
  var lastRow = sheet.getLastRow();
  var headerRow = isBannerPresent_(sheet) ? 2 : 1;
  var startData = headerRow + 1;
  if (lastRow < startData) return { error: 'no data' };

  // 5 عيّنات: أول، ربع، نصف، ثلاثة أرباع، آخر
  var n = lastRow - startData + 1;
  var picks = [0, Math.floor(n / 4), Math.floor(n / 2), Math.floor(n * 3 / 4), n - 1].map(function(i) { return startData + i; });
  var samples = [];
  picks.forEach(function(rowNum) {
    var range = sheet.getRange(rowNum, 1, 1, RECOVERY_AUDIT.COLS.TOTAL);
    var vals = range.getValues()[0];
    var disps = range.getDisplayValues()[0];
    var sample = { sheetRow: rowNum, cells: {} };
    var cols = ['TIMES', 'DATES', 'IATA', 'FLIGHTS'];
    cols.forEach(function(catName) {
      RECOVERY_AUDIT.COLS[catName].forEach(function(c) {
        var raw = vals[c - 1], disp = disps[c - 1];
        sample.cells[catName + '_col' + c] = {
          raw: raw,
          rawType: (raw instanceof Date) ? 'Date' : typeof raw,
          rawStr: String(raw).substring(0, 40),
          disp: String(disp).substring(0, 40)
        };
      });
    });
    samples.push(sample);
  });
  return { rowsInspected: samples.length, samples: samples };
}
