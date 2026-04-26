/**
 * فحص كل رسائل ثريد معين
 */
function inspectThread(args) {
  args = args || {};
  var tid = args.threadId;
  if (!tid) return { error: 'threadId missing' };

  var thread = GmailApp.getThreadById(tid);
  if (!thread) return { error: 'thread not found' };

  var messages = thread.getMessages();
  var result = {
    threadId: tid,
    labels: thread.getLabels().map(function(l) { return l.getName(); }),
    messageCount: messages.length,
    messages: []
  };

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var atts = msg.getAttachments();
    var pdfs = [];
    for (var a = 0; a < atts.length; a++) {
      if (atts[a].getContentType() === 'application/pdf') {
        pdfs.push({ name: atts[a].getName(), size: atts[a].getSize() });
      }
    }
    result.messages.push({
      index: i,
      messageId: msg.getId(),
      from: msg.getFrom(),
      date: Utilities.formatDate(msg.getDate(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'),
      subject: msg.getSubject(),
      bodySnippet: msg.getPlainBody().substring(0, 300).replace(/\s+/g,' '),
      pdfs: pdfs
    });
  }

  return result;
}
