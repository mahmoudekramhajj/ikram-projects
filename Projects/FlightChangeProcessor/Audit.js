/**
 * Audit.js — دالة تدقيق داخلية دورية
 *
 * الهدف: مقارنة عدد الإيميلات بعدد الصفوف في V4
 *        لضمان أن كل إيميل تمت معالجته وكل حاج استُخرج منه صفوف
 *
 * تُستخدم بعد كل تشغيل لـ scanEmailsV3 كمراجعة داخلية
 */

function auditEmailsVsRowsV4() {
  var startTime = new Date();
  var report = {
    timestamp: Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss'),
    emails: {},
    rows: {},
    gap: {},
    duration: null
  };

  // ========================================
  // 1. عدّ الإيميلات لكل label
  // ========================================
  report.emails.totalTKT = countThreads_('label:' + CONFIG.GMAIL_LABEL);
  report.emails.processed = countThreads_('label:' + CONFIG.PROCESSED_LABEL);
  report.emails.skipped = countThreads_('label:' + CONFIG.SKIPPED_LABEL);
  report.emails.unprocessed = report.emails.totalTKT - report.emails.processed - report.emails.skipped;

  // ========================================
  // 2. إحصائيات الرسائل (كل thread قد يحوي عدة رسائل)
  // ========================================
  report.emails.messagesProcessed = countMessages_('label:' + CONFIG.PROCESSED_LABEL);
  report.emails.messagesSkipped = countMessages_('label:' + CONFIG.SKIPPED_LABEL);

  // ========================================
  // 3. إحصائيات شيت V4 الموحّد
  // ========================================
  var sheet = getOrCreateUnifiedSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    report.rows = { total: 0 };
  } else {
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

    var uniquePNRs = {};
    var uniqueIncidents = {};
    var uniquePilgrims = {};  // PNR + firstName + lastName
    var rowsPerIncident = {};

    for (var i = 0; i < data.length; i++) {
      var pnr = String(data[i][1] || '').trim();
      var incident = String(data[i][3] || '').trim();
      var fn = String(data[i][5] || '').trim().toUpperCase();
      var ln = String(data[i][6] || '').trim().toUpperCase();

      if (pnr) uniquePNRs[pnr] = true;
      if (incident) {
        uniqueIncidents[incident] = true;
        rowsPerIncident[incident] = (rowsPerIncident[incident] || 0) + 1;
      }
      if (pnr || fn || ln) uniquePilgrims[pnr + '|' + fn + '|' + ln] = true;
    }

    // متوسط الصفوف لكل incident
    var incidentRowCounts = Object.keys(rowsPerIncident).map(function(k) { return rowsPerIncident[k]; });
    var avgRowsPerIncident = incidentRowCounts.length > 0
      ? (incidentRowCounts.reduce(function(a, b) { return a + b; }, 0) / incidentRowCounts.length).toFixed(2)
      : 0;
    var maxRowsPerIncident = incidentRowCounts.length > 0 ? Math.max.apply(null, incidentRowCounts) : 0;

    report.rows = {
      total: data.length,
      uniquePNRs: Object.keys(uniquePNRs).length,
      uniqueIncidents: Object.keys(uniqueIncidents).length,
      uniquePilgrims: Object.keys(uniquePilgrims).length,
      avgRowsPerIncident: avgRowsPerIncident,
      maxRowsPerIncident: maxRowsPerIncident
    };
  }

  // ========================================
  // 4. تحليل الفجوة
  // ========================================
  report.gap.processedEmailsVsUniqueIncidents =
    report.emails.processed - (report.rows.uniqueIncidents || 0);

  report.gap.expectedMinRowsIfSinglePilgrim =
    (report.rows.uniqueIncidents || 0);

  report.gap.skippedPercentage = report.emails.totalTKT > 0
    ? ((report.emails.skipped / report.emails.totalTKT) * 100).toFixed(1) + '%'
    : '0%';

  // علامات تحذيرية
  report.warnings = [];

  if (report.emails.unprocessed > 0) {
    report.warnings.push('⚠️ ' + report.emails.unprocessed + ' إيميل بـ TKT لم يُعالج بعد (لا Processed ولا Skipped)');
  }

  if (report.emails.skipped > report.emails.processed) {
    report.warnings.push('🔴 عدد Skipped أكبر من Processed — مشكلة كبيرة في الـ parser');
  }

  if (report.rows.total < report.emails.processed) {
    report.warnings.push('⚠️ صفوف V4 أقل من عدد الإيميلات المعالجة — احتمال فقدان بيانات');
  }

  if (report.rows.uniqueIncidents < report.emails.processed) {
    var missing = report.emails.processed - report.rows.uniqueIncidents;
    report.warnings.push('⚠️ ' + missing + ' إيميل Processed لا يوجد له incident في V4 (فقدان صامت محتمل)');
  }

  report.duration = ((new Date() - startTime) / 1000).toFixed(2) + 's';
  return report;
}


/**
 * عدّ threads لـ query معيّن بالـ pagination
 */
function countThreads_(query) {
  var count = 0;
  var start = 0;
  var batchSize = 500;

  while (true) {
    var batch = GmailApp.search(query, start, batchSize);
    count += batch.length;
    if (batch.length < batchSize) break;
    start += batchSize;
    if (start > 10000) break; // حماية
  }

  return count;
}


/**
 * عدّ messages (ليس threads) لـ query معيّن
 */
function countMessages_(query) {
  var count = 0;
  var start = 0;
  var batchSize = 100;

  while (true) {
    var batch = GmailApp.search(query, start, batchSize);
    if (batch.length === 0) break;

    for (var i = 0; i < batch.length; i++) {
      count += batch[i].getMessageCount();
    }

    if (batch.length < batchSize) break;
    start += batchSize;
    if (start > 10000) break;
  }

  return count;
}
