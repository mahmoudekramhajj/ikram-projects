/**
 * يحصي صفوف PD التي امتلأ فيها عمود Z (TICKET_URL)
 */
function pdTicketUrlStats() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName(TL.Config.SHEET_PD);
  if (!sheet) return { error: 'PD not found' };

  var ticketUrlCol = findColumnByHeader_(sheet, TL.Config.PD_HEADERS.TICKET_URL, TL.Config.PD_COL.TICKET_URL);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { totalPilgrims: 0, withUrl: 0 };

  var data = sheet.getRange(2, ticketUrlCol, lastRow - 1, 1).getValues();
  var withUrl = 0;
  var driveLinks = 0;
  var otherLinks = 0;
  var samples = [];

  for (var i = 0; i < data.length; i++) {
    var v = String(data[i][0] || '').trim();
    if (v) {
      withUrl++;
      if (v.indexOf('drive.google.com') !== -1) driveLinks++;
      else otherLinks++;
      if (samples.length < 3) samples.push(v);
    }
  }

  return {
    totalPilgrims: data.length,
    withUrl: withUrl,
    withoutUrl: data.length - withUrl,
    driveLinks: driveLinks,
    otherLinks: otherLinks,
    samples: samples,
    coveragePct: ((withUrl / data.length) * 100).toFixed(1) + '%'
  };
}

/**
 * إحصائيات RunLog مفلترة من timestamp معين
 * @param {string} sinceIso - مثال '2026-04-26T17:30:00'
 */
function runLogStatsSince(sinceIso) {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);
  var sheet = ss.getSheetByName('TL_RunLog');
  if (!sheet) return { error: 'No log' };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { count: 0 };
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var since = new Date(sinceIso).getTime();
  var byStatus = {};
  var byStage = {};
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    var ts = data[i][0] instanceof Date ? data[i][0].getTime() : new Date(data[i][0]).getTime();
    if (ts < since) continue;
    count++;
    var status = String(data[i][5] || 'EMPTY');
    var stage = String(data[i][4] || 'EMPTY');
    byStatus[status] = (byStatus[status] || 0) + 1;
    byStage[stage] = (byStage[stage] || 0) + 1;
  }
  return { sinceIso: sinceIso, count: count, byStatus: byStatus, byStage: byStage };
}
