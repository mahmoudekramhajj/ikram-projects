/**
 * _SkippedReview.js — مراجعة إيميلات Skipped واحداً تلو الآخر
 */


/**
 * قائمة كل إيميلات Skipped الحالية مع تفاصيلها
 */
function listSkippedEmails() {
  var query = 'label:' + CONFIG.SKIPPED_LABEL;
  var threads = GmailApp.search(query, 0, 100);

  var list = [];
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();

    // ثريد له Skipped label → نأخذ كل رسائله (بدل فلترة per-message)
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];

      var attachments = msg.getAttachments();
      var pdfNames = [];
      var hasPdf = false;
      for (var a = 0; a < attachments.length; a++) {
        if (attachments[a].getContentType() === 'application/pdf') {
          pdfNames.push(attachments[a].getName());
          hasPdf = true;
        }
      }

      list.push({
        index: list.length + 1,
        threadId: thread.getId(),
        messageId: msg.getId(),
        date: Utilities.formatDate(msg.getDate(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'),
        subject: msg.getSubject().substring(0, 100),
        from: msg.getFrom().substring(0, 60),
        hasPdf: hasPdf,
        pdfCount: pdfNames.length,
        pdfNames: pdfNames.slice(0, 3),
        snippet: msg.getPlainBody().substring(0, 200).replace(/\s+/g, ' '),
        threadLink: 'https://mail.google.com/mail/u/0/#inbox/' + thread.getId()
      });
    }
  }

  return { total: list.length, items: list };
}


/**
 * فحص إيميل واحد بالتفصيل
 */
function inspectSkippedEmail(args) {
  args = args || {};
  var index = args.index;
  var messageId = args.messageId;

  var msg;
  if (messageId) {
    msg = GmailApp.getMessageById(messageId);
  } else if (index) {
    var list = listSkippedEmails();
    if (index > list.items.length) return { error: 'index out of range' };
    msg = GmailApp.getMessageById(list.items[index - 1].messageId);
  } else {
    return { error: 'provide index or messageId' };
  }

  var attachments = msg.getAttachments();
  var pdfs = [];
  for (var a = 0; a < attachments.length; a++) {
    if (attachments[a].getContentType() === 'application/pdf') {
      pdfs.push({
        name: attachments[a].getName(),
        size: attachments[a].getSize()
      });
    }
  }

  return {
    messageId: msg.getId(),
    threadId: msg.getThread().getId(),
    date: Utilities.formatDate(msg.getDate(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'),
    subject: msg.getSubject(),
    from: msg.getFrom(),
    to: msg.getTo(),
    pdfCount: pdfs.length,
    pdfs: pdfs,
    hasHtml: !!msg.getBody(),
    plainBody: msg.getPlainBody().substring(0, 2000),
    threadLink: 'https://mail.google.com/mail/u/0/#inbox/' + msg.getThread().getId()
  };
}
