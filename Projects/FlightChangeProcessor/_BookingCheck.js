/**
 * فحص: هل Booking ID موجود في V4؟
 */
function checkBookingIdInV4(args) {
  args = args || {};
  var bid = (args.bid || '').toUpperCase().trim();
  if (!bid) return { error: 'ضع bid' };

  var ss = SpreadsheetApp.openById(CONFIG.UNIFIED_SPREADSHEET_ID);
  var sh = ss.getSheetByName(CONFIG.UNIFIED_SHEET_NAME) || ss.getSheets()[0];
  var data = sh.getDataRange().getValues();
  var headers = data[0];

  var bidCols = [];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').toLowerCase();
    if (h.indexOf('booking') !== -1) bidCols.push(i);
  }

  var matches = [];
  for (var r = 1; r < data.length; r++) {
    for (var c = 0; c < bidCols.length; c++) {
      var val = String(data[r][bidCols[c]] || '').toUpperCase().trim();
      if (val === bid) {
        matches.push({ row: r + 1, values: summarizeRow_(headers, data[r]) });
        break;
      }
    }
  }

  return { searched: bid, found: matches.length, matches: matches };
}

/**
 * يبحث عن إيميل بـ Booking ID في المحتوى أو اسم PDF
 */
function findEmailByBookingId(args) {
  args = args || {};
  var bid = args.bid;
  if (!bid) return { error: 'ضع bid' };

  var query = bid;
  var threads = GmailApp.search(query, 0, 10);

  var result = [];
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    var labels = thread.getLabels().map(function(l) { return l.getName(); });
    var msg = messages[messages.length - 1];
    var atts = msg.getAttachments();
    var pdfNames = [];
    for (var a = 0; a < atts.length; a++) {
      if (atts[a].getContentType() === 'application/pdf') pdfNames.push(atts[a].getName());
    }

    result.push({
      threadId: thread.getId(),
      subject: msg.getSubject(),
      date: Utilities.formatDate(msg.getDate(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'),
      labels: labels,
      pdfs: pdfNames,
      incidentInSubject: (msg.getSubject().match(/Incident#?\s*(\d+)/i) || [])[1] || null
    });
  }

  return { query: query, found: result.length, results: result };
}
