/**
 * TelegramReport.js — تقرير دوري كل 3 ساعات على تيليجرام
 *
 * الإعداد (مرة واحدة من محرر GAS):
 *   setReportConfig('BOT_TOKEN_HERE', '-1001234567890')
 *
 * Script Properties المستخدمة:
 *   TL_BOT_TOKEN      — توكن البوت (أو نفس BOT_TOKEN المشترك)
 *   TL_REPORT_CHAT_ID — Chat ID للمجموعة المستهدفة
 *   TL_LAST_LINKED    — عدد المرتبطين آخر تقرير (لحساب الجديد)
 */

var TL_TELEGRAM_API = 'https://api.telegram.org/bot';

// ==================== الإرسال ====================

function sendTelegram_(chatId, token, text) {
  var url = TL_TELEGRAM_API + token + '/sendMessage';
  var payload = {
    chat_id: chatId,
    text: text,
    parse_mode: 'HTML'
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Telegram error ' + code + ': ' + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(resp.getContentText());
}

// ==================== جمع الإحصائيات ====================

function collectReportStats_() {
  var ss = SpreadsheetApp.openById(TL.Config.SS_ID);

  // إحصائيات PD (رابط التذكرة)
  var pdStats = pdTicketUrlStats();

  // إحصائيات Unresolved
  var unresolvedSheet = ss.getSheetByName('TL_UnresolvedTickets');
  var pendingCount = 0;
  var linkedCount = 0;
  if (unresolvedSheet && unresolvedSheet.getLastRow() > 1) {
    var statusData = unresolvedSheet.getRange(2, 14, unresolvedSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < statusData.length; i++) {
      var s = String(statusData[i][0] || '').trim();
      if (s === 'PENDING_MANUAL') pendingCount++;
      else if (s === 'LINKED') linkedCount++;
    }
  }

  // الجديد منذ آخر تقرير
  var props = PropertiesService.getScriptProperties();
  var lastLinked = parseInt(props.getProperty('TL_LAST_LINKED') || '0', 10);
  var currentLinked = pdStats.withUrl || 0;
  var newSinceLastReport = Math.max(0, currentLinked - lastLinked);

  // حفظ العدد الحالي للمقارنة القادمة
  props.setProperty('TL_LAST_LINKED', String(currentLinked));

  return {
    totalPilgrims: pdStats.totalPilgrims || 0,
    withUrl: currentLinked,
    withoutUrl: pdStats.withoutUrl || 0,
    coveragePct: pdStats.coveragePct || '0%',
    pendingManual: pendingCount,
    linkedUnresolved: linkedCount,
    newSinceLastReport: newSinceLastReport
  };
}

// ==================== بناء نص التقرير ====================

function buildReportText_(stats) {
  var now = new Date();
  var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var newBadge = stats.newSinceLastReport > 0
    ? '\n🆕 جديد منذ آخر تقرير: <b>+' + stats.newSinceLastReport + '</b>'
    : '';

  var pendingLine = stats.pendingManual > 0
    ? '\n⏳ معلّق (يحتاج مراجعة): <b>' + stats.pendingManual + '</b>'
    : '\n✅ لا توجد تذاكر معلّقة';

  return '📊 <b>تقرير TicketLinker</b>\n' +
    '🗓 ' + dateStr + ' — ' + timeStr + '\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    '🎫 مرتبط: <b>' + stats.withUrl + ' / ' + stats.totalPilgrims + '</b>' +
    ' (' + stats.coveragePct + ')' +
    newBadge +
    '\n❌ بدون رابط: <b>' + stats.withoutUrl + '</b>' +
    pendingLine + '\n' +
    '━━━━━━━━━━━━━━━━━\n' +
    '🤖 TicketLinker Auto';
}

// ==================== الدالة الرئيسية ====================

/**
 * يُرسل تقرير الحالة الحالية لتيليجرام
 * تُستدعى من التريجر كل 3 ساعات
 */
function scheduledSendReport() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TL_BOT_TOKEN') || props.getProperty('BOT_TOKEN');
  var chatId = props.getProperty('TL_REPORT_CHAT_ID');

  if (!token || !chatId) {
    Logger.log('[Report] Missing config — run setReportConfig() first');
    return { error: 'Missing TL_BOT_TOKEN or TL_REPORT_CHAT_ID in Script Properties' };
  }

  try {
    var stats = collectReportStats_();
    var text = buildReportText_(stats);
    sendTelegram_(chatId, token, text);
    Logger.log('[Report] Sent: ' + stats.withUrl + '/' + stats.totalPilgrims);
    return { ok: true, stats: stats };
  } catch (e) {
    Logger.log('[Report] Error: ' + e.message);
    return { error: e.message };
  }
}

// ==================== الإعداد والتثبيت ====================

/**
 * إعداد بيانات التقرير (شغّلها مرة واحدة)
 * @param {string} botToken — توكن البوت
 * @param {string} chatId  — Chat ID للمجموعة (مثال: -1001234567890)
 */
function setReportConfig(botToken, chatId) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('TL_BOT_TOKEN', botToken);
  props.setProperty('TL_REPORT_CHAT_ID', chatId);
  return { ok: true, chatId: chatId, tokenSet: !!botToken };
}

/**
 * يثبّت تريجر كل 3 ساعات لإرسال التقرير
 */
function installReportTrigger() {
  // احذف أي trigger قديم
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scheduledSendReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('scheduledSendReport')
    .timeBased()
    .everyHours(3)
    .create();

  return { ok: true, message: 'Report trigger installed: every 3 hours' };
}

/**
 * اختبار فوري — يرسل تقريراً الآن
 */
function testSendReport() {
  return scheduledSendReport();
}
