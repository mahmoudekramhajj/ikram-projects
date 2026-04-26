// ============================================
// 🔔 TicketsMonitor — مراقبة البلاغات الجديدة وإشعار المشغّلين كل 15 دقيقة
// ============================================

var TICKETS_LAST_SEEN_KEY = 'tickets_last_seen_id';
var ESCALATION_LOG_SHEET = 'EscalationLog';

// ============================================
// الدالة الرئيسية — تعمل عبر Trigger كل 15 دقيقة
// ============================================
function checkNewTickets_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var lastSeen = props.getProperty(TICKETS_LAST_SEEN_KEY) || '';

    var sheet = getReportsSheet_();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var newTickets = [];
    var maxId = lastSeen;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ticketId = String(row[0] || '').trim();
      if (!ticketId || !ticketId.indexOf('TKT-') === 0) continue;
      if (ticketId.indexOf('TKT-') !== 0) continue;
      if (lastSeen && ticketId <= lastSeen) continue;
      newTickets.push(row);
      if (ticketId > maxId) maxId = ticketId;
    }

    // تحقّق من ردود الإدارة (مدمج مع نفس التريغر)
    try { checkManagementReplies_(); } catch (e) { Logger.log('checkManagementReplies_ failed: ' + e.message); }

    if (newTickets.length === 0) {
      Logger.log('checkNewTickets_: no new tickets (last seen: ' + lastSeen + ')');
      return;
    }

    Logger.log('checkNewTickets_: found ' + newTickets.length + ' new ticket(s)');

    // أرسل ملخّص + بلاغات منفردة
    if (newTickets.length > 1) {
      var header = '🆕 <b>' + newTickets.length + ' بلاغات جديدة وصلت</b>\n━━━━━━━━━━━━━━\nسيتم إرسال تفاصيل كل بلاغ منفصلاً 👇';
      for (var k = 0; k < ADMIN_IDS.length; k++) {
        try { sendMessageSafe_(ADMIN_IDS[k], header); } catch (e) {}
      }
    }

    for (var j = 0; j < newTickets.length; j++) {
      try {
        notifyOperatorTicket_(newTickets[j]);
      } catch (e) {
        Logger.log('Failed to notify ticket ' + newTickets[j][0] + ': ' + e.message);
      }
      Utilities.sleep(500);
    }

    props.setProperty(TICKETS_LAST_SEEN_KEY, maxId);
    Logger.log('checkNewTickets_: updated last_seen to ' + maxId);
  } catch (e) {
    Logger.log('checkNewTickets_ ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ============================================
// بناء ملخّص ذكي للبلاغ
// ============================================
function buildTicketSummary_(row) {
  var ticketId = String(row[0] || '');
  var chatId = String(row[1] || '');
  var passport = String(row[2] || '');
  var name = String(row[3] || '-');
  var section = String(row[4] || '-');
  var issue = String(row[5] || '');
  var complaint = String(row[6] || '');
  var createdAt = String(row[9] || '');

  var sectionIcons = {
    'flight': '✈️',
    'hotel': '🏨',
    'transport': '🚌',
    'visa': '🛂',
    'personal': '👤'
  };
  var sectionNames = {
    'flight': 'طيران',
    'hotel': 'فندق',
    'transport': 'نقل',
    'visa': 'تأشيرة/تذكرة',
    'personal': 'شخصي'
  };
  var icon = sectionIcons[section] || '📌';
  var sectionName = sectionNames[section] || section;

  var msg = '🆕 <b>بلاغ جديد — ' + ticketId + '</b>\n';
  msg += '⏰ <code>' + createdAt + '</code>\n';
  msg += '━━━━━━━━━━━━━━\n';
  msg += '👤 الحاج: <b>' + name + '</b>\n';
  msg += '🆔 الجواز: <code>' + passport + '</code>\n';
  msg += '💬 ChatID: <code>' + chatId + '</code>\n';
  msg += '📌 القسم: ' + icon + ' ' + sectionName + '\n';
  if (issue) msg += '🏷️ نوع المشكلة: ' + issue + '\n';

  // محاولة جلب بيانات الحاج من Personal Details
  var pilgrim = null;
  try { pilgrim = findPilgrimByPassport_(passport); } catch (e) {}

  if (pilgrim && pilgrim.rowData) {
    var r = pilgrim.rowData;
    var contractName = String(r[3] || '').trim();
    var packageId = String(r[1] || '').trim();
    var ctype = String(r[25] || '').trim() || String(r[34] || '').trim();
    msg += '🔖 العقد: ';
    if (ctype) msg += '<b>' + ctype + '</b> — ';
    msg += '<code>' + (contractName || '-') + '</code>\n';
    if (packageId) msg += '📦 الباقة: <code>' + packageId + '</code>\n';
  } else {
    msg += '❓ <i>الحاج غير مسجَّل في Personal Details</i>\n';
  }

  msg += '━━━━━━━━━━━━━━\n';
  msg += '💬 <b>نص الشكوى (الأصلي):</b>\n';
  var maxLen = 600;
  if (complaint && complaint.trim()) {
    if (complaint.length > maxLen) {
      msg += complaint.substring(0, maxLen) + '...\n<i>(النص مقتطع)</i>';
    } else {
      msg += complaint;
    }
  } else {
    msg += '<i>(فارغ)</i>';
  }

  // ترجمة عربية للشكاوى غير العربية
  if (complaint && complaint.trim() && !isArabicText_(complaint)) {
    var translated = null;
    try { translated = translateComplaintToArabic_(complaint); } catch (e) {}
    if (translated) {
      msg += '\n\n🌐 <b>ترجمة عربية:</b>\n';
      if (translated.length > maxLen) {
        msg += translated.substring(0, maxLen) + '...';
      } else {
        msg += translated;
      }
    }
  }

  msg += '\n━━━━━━━━━━━━━━\n';

  // تحليل أولي تلقائي
  var analysis = [];
  if (!complaint || !complaint.trim()) {
    analysis.push('⚠️ نص الشكوى فارغ — قد يكون إرسال PDF أو خطأ');
  }
  if (!pilgrim) {
    analysis.push('❌ بيانات الحاج غير موجودة');
  }
  if (pilgrim && pilgrim.rowData) {
    var r2 = pilgrim.rowData;
    var ct = String(r2[25] || '').trim();
    if (ct === 'B2C') {
      analysis.push('ℹ️ B2C — تذكرة الحاج محجوزة بنفسه عبر نسك');
    } else if (ct === 'B2B') {
      analysis.push('🏢 B2B — عقد جماعي عبر إكرام');
    }
  }

  // كشف بلاغات مكرَّرة (نفس chatId + قريب)
  try {
    var sheet = getReportsSheet_();
    var allData = sheet.getDataRange().getValues();
    var sameChatRecent = 0;
    for (var x = 1; x < allData.length; x++) {
      if (String(allData[x][1]) === chatId && allData[x][0] !== ticketId) {
        sameChatRecent++;
      }
    }
    if (sameChatRecent >= 3) {
      analysis.push('🔁 الحاج لديه ' + sameChatRecent + ' بلاغات سابقة');
    }
  } catch (e) {}

  if (analysis.length) {
    msg += '🤖 <b>تحليل أولي:</b>\n';
    msg += analysis.map(function(a) { return '• ' + a; }).join('\n');
  }

  return msg;
}

// ============================================
// أزرار التفاعل (المرحلة 1 — أزرار بسيطة لا تعمل بعد، تشير للمرحلة 2)
// ============================================
function buildTicketButtons_(ticketId, chatId, passport) {
  return [
    [
      { text: '📨 رد سريع', callback_data: 'tm_reply:' + ticketId },
      { text: '✏️ كتابة رد', callback_data: 'tm_custom:' + ticketId }
    ],
    [
      { text: '🚀 للإدارة', callback_data: 'tm_escalate:' + ticketId },
      { text: '✅ إغلاق', callback_data: 'tm_close:' + ticketId }
    ],
    [
      { text: '🔄 جلب البيانات الكاملة', callback_data: 'tm_full:' + ticketId }
    ]
  ];
}

// ============================================
// إشعار بلاغ منفرد
// ============================================
function notifyOperatorTicket_(row) {
  var msg = buildTicketSummary_(row);
  var buttons = buildTicketButtons_(row[0], row[1], row[2]);
  for (var i = 0; i < ADMIN_IDS.length; i++) {
    try {
      sendMessageSafe_(ADMIN_IDS[i], msg, { inline_keyboard: buttons });
    } catch (e) {
      Logger.log('Failed to send to admin ' + ADMIN_IDS[i] + ': ' + e.message);
    }
  }
}

// ============================================
// تركيب Trigger كل 15 دقيقة
// ============================================
function setupTicketsCheckTrigger() {
  // إزالة triggers قديمة أولاً
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkNewTickets_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('checkNewTickets_')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('✅ Tickets monitor trigger installed (every 15 minutes)');
  return 'تم تركيب trigger المراقبة (كل 15 دقيقة)';
}

function removeTicketsCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkNewTickets_') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' trigger(s)');
  return 'تمت إزالة ' + removed + ' trigger(s)';
}

// ============================================
// إعادة تعيين آخر TKT مرئي (للاختبار)
// ============================================
function resetTicketsLastSeen() {
  PropertiesService.getScriptProperties().deleteProperty(TICKETS_LAST_SEEN_KEY);
  Logger.log('Tickets last_seen reset');
  return 'تم مسح المؤشر — التشغيل التالي سيرسل كل البلاغات';
}

function getTicketsLastSeen() {
  var v = PropertiesService.getScriptProperties().getProperty(TICKETS_LAST_SEEN_KEY) || '(غير مُعيَّن)';
  Logger.log('Last seen: ' + v);
  return v;
}

// دالة مساعدة لتفعيل أذونات Gmail (تشغّلها مرة من Editor)
function authorizeGmail() {
  GmailApp.createDraft('test@example.com', 'auth test', 'auth test');
  return 'Gmail authorization granted';
}

// فحص مفتاح Claude API
function checkClaudeApiKey() {
  var k = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (k) return 'CLAUDE_API_KEY موجود — أول 10 أحرف: ' + k.substring(0, 10) + '...';
  return 'CLAUDE_API_KEY غير موجود';
}


function setTicketsLastSeen(ticketId) {
  PropertiesService.getScriptProperties().setProperty(TICKETS_LAST_SEEN_KEY, String(ticketId));
  return 'تم تعيين آخر TKT مرئي إلى ' + ticketId;
}

// إرسال إشعار اختبار لبلاغ محدّد بدون لمس last_seen
function notifyTicketById(ticketId) {
  var sheet = getReportsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === ticketId) {
      notifyOperatorTicket_(data[i]);
      return 'تم إرسال إشعار اختبار للبلاغ ' + ticketId;
    }
  }
  return 'البلاغ غير موجود: ' + ticketId;
}

// ============================================
// تشغيل اختبار يدوي (يحاكي Trigger)
// ============================================
function testCheckNewTickets() {
  checkNewTickets_();
}

// ============================================
// 🎯 المرحلة 2: معالجة أزرار المشغّلين (callback_query)
// ============================================
function handleTicketMonitorCallback_(opChatId, data, callback) {
  // data format: tm_<action>:<ticketId>
  var parts = data.replace('tm_', '').split(':');
  var action = parts[0];
  var ticketId = parts.slice(1).join(':');

  if (!ticketId) {
    sendMessageSafe_(opChatId, '⚠️ بلاغ غير محدّد');
    return;
  }

  var ticket = getTicketById_(ticketId);
  if (!ticket) {
    sendMessageSafe_(opChatId, '❌ البلاغ ' + ticketId + ' غير موجود');
    return;
  }

  switch (action) {
    case 'full':
      tmActionFull_(opChatId, ticket);
      break;
    case 'close':
      tmActionClose_(opChatId, ticket);
      break;
    case 'custom':
      tmActionCustom_(opChatId, ticket);
      break;
    case 'reply':
      tmActionQuickReplyMenu_(opChatId, ticket);
      break;
    case 'qr':
      tmActionQuickReplySend_(opChatId, ticket, parts[2] || '');
      break;
    case 'escalate':
      tmActionEscalate_(opChatId, ticket);
      break;
    case 'send':
      tmActionSend_(opChatId, ticket);
      break;
    case 'edit':
      tmActionEdit_(opChatId, ticket);
      break;
    case 'cancel':
      tmActionCancel_(opChatId, ticket);
      break;
    default:
      sendMessageSafe_(opChatId, '⚠️ إجراء غير معروف: ' + action);
  }
}

// ============================================
// جلب بلاغ من الشيت بالـ ID
// ============================================
function getTicketById_(ticketId) {
  var sheet = getReportsSheet_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === ticketId) {
      return { row: data[i], rowIndex: i + 1 };
    }
  }
  return null;
}

// ============================================
// 🔄 جلب البيانات الكاملة (رحلات + فنادق + باقة)
// ============================================
function tmActionFull_(opChatId, ticket) {
  var row = ticket.row;
  var passport = String(row[2] || '');
  var name = String(row[3] || '-');

  var msg = '🔍 <b>بيانات كاملة — ' + row[0] + '</b>\n';
  msg += '👤 ' + name + ' (' + passport + ')\n';
  msg += '━━━━━━━━━━━━━━\n';

  // الطيران
  try {
    var arrival = getPilgrimArrival_(passport, 'ar');
    var departure = getPilgrimDeparture_(passport, 'ar');
    if (arrival) {
      msg += '✈️ <b>الوصول:</b>\n';
      msg += '   ' + arrival.flightNo + ' → ' + arrival.airportAr + '\n';
      msg += '   📅 ' + Utilities.formatDate(new Date(arrival.date), 'Asia/Riyadh', 'yyyy-MM-dd') + '\n';
    }
    if (departure) {
      msg += '✈️ <b>المغادرة:</b>\n';
      msg += '   ' + departure.flightNo + ' من ' + departure.airportAr + '\n';
      msg += '   📅 ' + Utilities.formatDate(new Date(departure.date), 'Asia/Riyadh', 'yyyy-MM-dd') + '\n';
    }
  } catch (e) {
    msg += '⚠️ تعذّر جلب بيانات الطيران\n';
  }

  // الفنادق
  try {
    var pilgrim = findPilgrimByPassport_(passport);
    if (pilgrim && pilgrim.rowData) {
      var packageId = String(pilgrim.rowData[1] || '').trim();
      var hotels = getPackageHotels_(packageId);
      if (hotels && hotels.length) {
        msg += '🏨 <b>الفنادق:</b>\n';
        for (var i = 0; i < hotels.length; i++) {
          var h = hotels[i];
          var city = getCityDisplayName_(h.city, 'ar');
          msg += '   ' + (i + 1) + '. ' + h.nameAr + ' (' + city + ')\n';
          msg += '      ' + Utilities.formatDate(new Date(h.checkIn), 'Asia/Riyadh', 'yyyy-MM-dd');
          msg += ' → ' + Utilities.formatDate(new Date(h.checkOut), 'Asia/Riyadh', 'yyyy-MM-dd') + '\n';
        }
      }
    }
  } catch (e) {
    msg += '⚠️ تعذّر جلب بيانات الفنادق\n';
  }

  msg += '━━━━━━━━━━━━━━\n';
  msg += '💬 <b>الشكوى الكاملة:</b>\n' + (row[6] || '-');

  sendMessageSafe_(opChatId, msg);
}

// ============================================
// ✅ إغلاق البلاغ
// ============================================
function tmActionClose_(opChatId, ticket) {
  setCache_('tm_pending_' + opChatId, {
    action: 'closing',
    ticketId: ticket.row[0],
    targetChatId: ticket.row[1],
    rowIndex: ticket.rowIndex
  }, 600);

  sendMessageSafe_(opChatId,
    '✅ <b>إغلاق البلاغ ' + ticket.row[0] + '</b>\n\n' +
    'اكتب ملاحظة الإغلاق (مختصرة):'
  );
}

// ============================================
// ✏️ كتابة رد مخصّص للحاج
// ============================================
function tmActionCustom_(opChatId, ticket) {
  setCache_('tm_pending_' + opChatId, {
    action: 'custom_reply',
    ticketId: ticket.row[0],
    targetChatId: ticket.row[1]
  }, 600);

  sendMessageSafe_(opChatId,
    '✏️ <b>رد مخصّص — ' + ticket.row[0] + '</b>\n' +
    '👤 الحاج: ' + ticket.row[3] + '\n\n' +
    'اكتب الرسالة التي تريد إرسالها للحاج:'
  );
}

// ============================================
// 📨 قائمة الردود السريعة
// ============================================
function tmActionQuickReplyMenu_(opChatId, ticket) {
  var ticketId = ticket.row[0];
  var buttons = [
    [{ text: '🔄 طلب تحديث البيانات', callback_data: 'tm_qr:' + ticketId + ':refresh' }],
    [{ text: '📸 طلب صورة من نسك', callback_data: 'tm_qr:' + ticketId + ':screenshot' }],
    [{ text: '🚀 رُفع للإدارة', callback_data: 'tm_qr:' + ticketId + ':escalated' }],
    [{ text: '✅ تم التحقق — البيانات صحيحة', callback_data: 'tm_qr:' + ticketId + ':verified' }],
    [{ text: '⬅️ رجوع', callback_data: 'tm_full:' + ticketId }]
  ];
  sendMessageSafe_(opChatId, '📨 <b>اختر قالباً جاهزاً:</b>', { inline_keyboard: buttons });
}

// ============================================
// إرسال رد سريع جاهز
// ============================================
function tmActionQuickReplySend_(opChatId, ticket, template) {
  var name = ticket.row[3];
  var ticketId = ticket.row[0];
  var targetChatId = ticket.row[1];
  var msg = '';

  switch (template) {
    case 'refresh':
      msg = 'السلام عليكم ورحمة الله وبركاته،\n\nنرجو منكم التكرّم بالضغط على زر "🔄 تحديث البيانات" في القائمة الرئيسية، ثم فتح القسم الذي تستفسرون عنه لرؤية البيانات بأحدث صورة.\n\nجزاكم الله خيراً\nفريق إكرام الضيف';
      break;
    case 'screenshot':
      msg = 'السلام عليكم ورحمة الله وبركاته،\n\nللتحقق الدقيق من بياناتكم، نرجو منكم التكرّم بإرسال صورة من فاتورة منصة نسك (Nusuk) إلى البريد الإلكتروني:\n📧 ikramdhtravel@arbhaj.com\n\nبمجرد وصولها سيتم التحقق والتعديل بإذن الله.\n\nجزاكم الله خيراً\nفريق إكرام الضيف';
      break;
    case 'escalated':
      msg = 'السلام عليكم ورحمة الله وبركاته،\n\nتم رفع موضوعكم إلى الإدارة، ونحن بانتظار توجيهاتهم. سنُحدِّثكم فور وصول الرد بإذن الله.\n\nجزاكم الله خيراً\nفريق إكرام الضيف';
      break;
    case 'verified':
      msg = 'السلام عليكم ورحمة الله وبركاته،\n\nتم التحقق من بياناتكم وهي مطابقة لما هو مسجَّل لكم على منصة نسك. كل شيء سليم بإذن الله.\n\nنرجو الضغط على "🔄 تحديث البيانات" لرؤية البيانات بأحدث صورة.\n\nجزاكم الله خيراً\nفريق إكرام الضيف';
      break;
    default:
      sendMessageSafe_(opChatId, '⚠️ قالب غير معروف');
      return;
  }

  try {
    sendMessageSafe_(targetChatId, msg);
    sendMessageSafe_(opChatId, '✅ تم إرسال الرد للحاج <b>' + name + '</b> (' + ticketId + ')');
    // تحديث ملاحظة الحل في الشيت
    updateTicketResolved_(ticket, 'Auto-template: ' + template);
  } catch (e) {
    sendMessageSafe_(opChatId, '❌ فشل الإرسال: ' + e.message);
  }
}

// ============================================
// 🗃️ شيت EscalationLog — تتبّع رسائل الإدارة
// ============================================
function getEscalationLogSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(ESCALATION_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ESCALATION_LOG_SHEET);
    sheet.appendRow([
      'TicketId', 'ThreadId', 'DraftId', 'Recipients', 'Subject',
      'EscalatedAt', 'Status', 'MessageCount', 'LastReplyFrom', 'LastReplyAt'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function logEscalation_(ticketId, threadId, draftId, recipients, subject) {
  try {
    var sheet = getEscalationLogSheet_();
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([
      ticketId, threadId, draftId, recipients, subject,
      now, 'draft', 0, '', ''
    ]);
  } catch (e) {
    Logger.log('logEscalation_ error: ' + e.message);
  }
}

// ============================================
// 📨 فحص ردود الإدارة على threads المُسجَّلة
// ============================================
function checkManagementReplies_() {
  try {
    var sheet = getEscalationLogSheet_();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var ourEmail = '';
    try { ourEmail = Session.getActiveUser().getEmail() || ''; } catch (e) {}

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var ticketId = row[0];
      var threadId = row[1];
      var status = row[6];
      var lastSeenCount = parseInt(row[7] || 0);

      if (!threadId || status === 'closed') continue;

      try {
        var thread = GmailApp.getThreadById(threadId);
        if (!thread) continue;

        var messages = thread.getMessages();
        var count = messages.length;

        // انتقال draft → sent
        if (count >= 1 && status === 'draft') {
          sheet.getRange(i + 1, 7).setValue('sent');
          status = 'sent';
        }

        // التحقق من ردود جديدة (count زاد)
        if (count > lastSeenCount) {
          // ابحث عن أحدث رسالة ليست من إيميلنا
          var newReplies = [];
          for (var j = lastSeenCount; j < count; j++) {
            var m = messages[j];
            var from = m.getFrom();
            if (ourEmail && from.indexOf(ourEmail) !== -1) continue; // رسالتنا
            newReplies.push(m);
          }

          if (newReplies.length > 0) {
            // إشعار بأحدث رد
            var latest = newReplies[newReplies.length - 1];
            notifyManagementReply_(ticketId, threadId, latest);

            // تحديث الشيت
            sheet.getRange(i + 1, 7).setValue('replied');
            sheet.getRange(i + 1, 9).setValue(latest.getFrom());
            sheet.getRange(i + 1, 10).setValue(
              Utilities.formatDate(latest.getDate(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss')
            );
          }

          sheet.getRange(i + 1, 8).setValue(count);
        }
      } catch (e) {
        Logger.log('Failed to check thread ' + threadId + ' for ' + ticketId + ': ' + e.message);
      }
    }
  } catch (e) {
    Logger.log('checkManagementReplies_ ERROR: ' + e.message);
  }
}

// ============================================
// إشعار رد إدارة جديد
// ============================================
function notifyManagementReply_(ticketId, threadId, message) {
  var from = message.getFrom();
  var subject = message.getSubject();
  var body = (message.getPlainBody() || '').substring(0, 1500);
  var date = message.getDate();

  var msg = '📨 <b>رد جديد من الإدارة على بلاغ ' + ticketId + '</b>\n';
  msg += '👤 من: <code>' + from + '</code>\n';
  msg += '📋 الموضوع: ' + subject + '\n';
  msg += '⏰ ' + Utilities.formatDate(date, 'Asia/Riyadh', 'yyyy-MM-dd HH:mm') + '\n';
  msg += '━━━━━━━━━━━━━━\n';
  msg += body;
  msg += '\n━━━━━━━━━━━━━━';

  var gmailUrl = 'https://mail.google.com/mail/u/0/#inbox/' + threadId;
  var buttons = [
    [{ text: '💬 رد للحاج', callback_data: 'tm_custom:' + ticketId }],
    [{ text: '🔗 فتح في Gmail', url: gmailUrl }],
    [{ text: '✅ إغلاق البلاغ', callback_data: 'tm_close:' + ticketId }]
  ];

  for (var i = 0; i < ADMIN_IDS.length; i++) {
    try {
      sendMessageSafe_(ADMIN_IDS[i], msg, { inline_keyboard: buttons });
    } catch (e) {
      Logger.log('Failed to notify reply for ' + ticketId + ': ' + e.message);
    }
  }
}

// ============================================
// 🚀 رفع للإدارة (إنشاء مسوّدة إيميل بالعربية دائماً)
// ============================================
function tmActionEscalate_(opChatId, ticket) {
  var row = ticket.row;
  var ticketId = row[0];
  var name = row[3];
  var passport = row[2];
  var section = row[4];
  var complaint = row[6];

  // ترجمة قسم البلاغ للعربية
  var sectionAr = {
    'flight': 'الطيران',
    'hotel': 'الفندق',
    'transport': 'النقل',
    'visa': 'التأشيرة/التذكرة',
    'personal': 'البيانات الشخصية',
    'package': 'الباقة'
  }[section] || section;

  var emailBody = 'السلام عليكم ورحمة الله وبركاته،\n\n';
  emailBody += 'تواصل معنا الحاج/ة ' + name + ' (' + ticketId + ') عبر بوت إكرام الضيف، بشكوى تخصّ ' + sectionAr + '.\n\n';

  emailBody += '📌 بيانات الحاج:\n';
  emailBody += '- الاسم: ' + name + '\n';
  emailBody += '- جواز السفر: ' + passport + '\n';

  try {
    var pilgrim = findPilgrimByPassport_(passport);
    if (pilgrim && pilgrim.rowData) {
      var r = pilgrim.rowData;
      emailBody += '- الجنسية: ' + (r[12] || '-') + '\n';
      emailBody += '- العقد: ' + (r[3] || '-') + '\n';
      emailBody += '- الباقة: ' + (r[1] || '-') + '\n';
    }
  } catch (e) {}

  emailBody += '\n💬 نص الشكوى:\n';

  // ترجمة الشكوى للعربية إن كانت بلغة أخرى
  if (complaint && complaint.trim()) {
    if (isArabicText_(complaint)) {
      emailBody += complaint + '\n';
    } else {
      var translated = null;
      try { translated = translateComplaintToArabic_(complaint); } catch (e) {}
      if (translated) {
        emailBody += translated + '\n\n';
        emailBody += '— النص الأصلي:\n' + complaint + '\n';
      } else {
        emailBody += complaint + '\n';
      }
    }
  } else {
    emailBody += '(لم يُرفق نص شكوى)\n';
  }

  emailBody += '\n🙏 المطلوب من الإدارة:\n';
  emailBody += 'مراجعة الموضوع والتوجيه بالإجراء المناسب.\n\n';
  emailBody += 'جزاكم الله خيراً\n';
  emailBody += 'فريق إكرام الضيف — بوت الحج\n';
  emailBody += 'البلاغ المرتبط: ' + ticketId;

  var subject = 'بلاغ ' + ticketId + ' — ' + name + ' — ' + sectionAr;
  var recipients = 'ikramdhtravel@arbhaj.com';
  var cc = 'MALEK_ALDAIRI1998@outlook.com';

  try {
    var draft = GmailApp.createDraft(recipients, subject, emailBody, { cc: cc });

    // استخراج threadId و draftId
    var threadId = '';
    var draftId = draft.getId();
    try {
      var draftMsg = draft.getMessage();
      if (draftMsg) {
        var thread = draftMsg.getThread();
        if (thread) threadId = thread.getId();
      }
    } catch (e) {}

    // تسجيل التصعيد للمتابعة التلقائية
    logEscalation_(ticketId, threadId, draftId, recipients + ',' + cc, subject);

    sendMessageSafe_(opChatId,
      '✅ <b>تم إنشاء مسوّدة إيميل في Gmail Drafts</b>\n\n' +
      '📧 إلى: ikramdhtravel@arbhaj.com\n' +
      '📧 نسخة: MALEK_ALDAIRI1998@outlook.com\n' +
      '📋 الموضوع: ' + subject + '\n\n' +
      '🔍 <b>المتابعة التلقائية:</b> سيتم رصد رد الإدارة وإشعارك فور وصوله.\n\n' +
      'افتح Gmail لمراجعة المسوّدة وإرسالها يدوياً.'
    );
  } catch (e) {
    sendMessageSafe_(opChatId, '❌ فشل إنشاء المسوّدة: ' + e.message);
  }
}

// ============================================
// تحديث حالة البلاغ في الشيت
// ============================================
function updateTicketResolved_(ticket, note) {
  try {
    var sheet = getReportsSheet_();
    var rowIdx = ticket.rowIndex;
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(rowIdx, 8).setValue('resolved'); // Status
    sheet.getRange(rowIdx, 9).setValue('Operator'); // ResolvedBy
    sheet.getRange(rowIdx, 11).setValue(now); // ResolvedAt
    sheet.getRange(rowIdx, 12).setValue(note || ''); // AdminResponse
  } catch (e) {
    Logger.log('updateTicketResolved_ error: ' + e.message);
  }
}

// ============================================
// Wrapper بسيط (fallback عند فشل Claude)
// ============================================
function wrapReplyText_(rawText, pilgrimName) {
  var greeting = 'السلام عليكم ورحمة الله وبركاته،';
  var addressing = pilgrimName ? ('\n\nأهلاً بك ' + pilgrimName + ' 🌷') : '';
  var body = rawText.trim();
  var footer = '\n\nجزاكم الله خيراً\nفريق إكرام الضيف';
  return greeting + addressing + '\n\n' + body + footer;
}

// ============================================
// 🤖 Claude polish: إعادة صياغة + ترجمة احترافية
// ============================================
function polishReplyWithClaude_(rawText, pilgrimName, pilgrimLang, complaintText) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null; // fallback إلى Wrapper البسيط

  var langNames = {
    'ar': 'Arabic',
    'en': 'English',
    'fr': 'French',
    'de': 'German',
    'it': 'Italian',
    'es': 'Spanish'
  };
  var pilgrimLangName = langNames[pilgrimLang] || 'Arabic';
  var needsBilingual = pilgrimLang && pilgrimLang !== 'ar';

  var systemPrompt = [
    'You are the customer-care editor for "Ikram Aldyf Hajj Travel" (إكرام الضيف).',
    'Your job: take a SHORT operator note and turn it into a polished, professional reply to a pilgrim via Telegram bot.',
    '',
    'STRICT rules:',
    '1. Always start with the Arabic Islamic greeting: السلام عليكم ورحمة الله وبركاته،',
    '2. Always end with: جزاكم الله خيراً ⏎ فريق إكرام الضيف',
    '3. Tone: warm, respectful, brief, professional. Use Hajj-appropriate language.',
    '4. Preserve the operator note INTENT exactly — never invent facts, dates, names, flight numbers, or hotel info.',
    '5. If pilgrim language is non-Arabic, produce BILINGUAL: Arabic first, then the second language. Mirror greeting/closing in both languages.',
    '6. Use line breaks for readability. Use light emoji ONLY where natural (✈️ 🏨 🌷 ✅).',
    '7. Output the FINAL message text only, no explanations, no markdown.',
    '',
    'Pilgrim language: ' + pilgrimLangName + (needsBilingual ? ' (BILINGUAL: Arabic + ' + pilgrimLangName + ')' : ' (Arabic only)'),
    'Pilgrim name: ' + (pilgrimName || '-')
  ].join('\n');

  var userPrompt = [
    'OPERATOR NOTE (the intent/content to convey):',
    '"""',
    rawText,
    '"""'
  ];
  if (complaintText) {
    userPrompt.push('');
    userPrompt.push('CONTEXT — pilgrim original complaint:');
    userPrompt.push('"""');
    userPrompt.push(complaintText.substring(0, 600));
    userPrompt.push('"""');
  }
  userPrompt.push('');
  userPrompt.push('Output ONLY the polished reply text — no preface, no explanations.');

  var payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1500,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt.join('\n') }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    if (response.getResponseCode() !== 200) {
      Logger.log('Claude polish error code: ' + response.getResponseCode());
      return null;
    }
    var data = JSON.parse(response.getContentText());
    if (!data.content || !data.content[0] || !data.content[0].text) return null;
    return data.content[0].text.trim();
  } catch (e) {
    Logger.log('Claude polish error: ' + e.message);
    return null;
  }
}

// ============================================
// قراءة لغة الحاج من جلسته
// ============================================
function getPilgrimLang_(targetChatId) {
  try {
    var session = getSession_(targetChatId);
    if (session && session.language) return session.language;
  } catch (e) {}
  return 'ar';
}

// ============================================
// كشف ما إذا كان النص عربياً غالباً
// ============================================
function isArabicText_(text) {
  if (!text) return true;
  var s = String(text);
  var arabic = (s.match(/[\u0600-\u06FF]/g) || []).length;
  var latin = (s.match(/[a-zA-Z]/g) || []).length;
  // إن كانت الحروف العربية > 30% من النص أو أكثر من اللاتينية → عربي
  return arabic > latin || (arabic > 0 && arabic / (arabic + latin + 1) > 0.3);
}

// ============================================
// 🌐 ترجمة الشكوى للعربية عبر Claude (مع كاش)
// ============================================
function translateComplaintToArabic_(text) {
  if (!text || !text.trim()) return null;
  if (isArabicText_(text)) return null; // أصلاً عربي

  var cacheKey = 'tm_xlate_' + Utilities.base64Encode(text.substring(0, 200)).substring(0, 50);
  var cached = getCache_(cacheKey);
  if (cached) return cached;

  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;

  var prompt = [
    'Translate the following pilgrim complaint into clear Modern Standard Arabic.',
    'Rules:',
    '- Keep flight numbers, codes, dates, and proper nouns AS-IS (no transliteration).',
    '- Preserve the meaning and tone exactly.',
    '- Output ONLY the Arabic translation, no preface, no quotes.',
    '',
    'Text to translate:',
    '"""',
    text.substring(0, 2000),
    '"""'
  ].join('\n');

  var payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1000,
    messages: [{ role: 'user', content: prompt }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
    if (response.getResponseCode() !== 200) return null;
    var data = JSON.parse(response.getContentText());
    if (!data.content || !data.content[0] || !data.content[0].text) return null;
    var translated = data.content[0].text.trim();
    setCache_(cacheKey, translated, 21600); // 6 ساعات
    return translated;
  } catch (e) {
    Logger.log('translateComplaintToArabic_ error: ' + e.message);
    return null;
  }
}

// ============================================
// معالجة نص المشغّل (custom reply / close note)
// ============================================
function handleTicketMonitorInput_(opChatId, text, pendingTm) {
  var action = pendingTm.action;
  var ticketId = pendingTm.ticketId;
  var targetChatId = pendingTm.targetChatId;

  if (action === 'custom_reply') {
    var ticket = getTicketById_(ticketId);
    var pilgrimName = ticket ? ticket.row[3] : '';
    var complaintText = ticket ? ticket.row[6] : '';
    var pilgrimLang = getPilgrimLang_(targetChatId);

    // 1) جرّب Claude أولاً
    var polished = null;
    var usedClaude = false;
    try {
      polished = polishReplyWithClaude_(text, pilgrimName, pilgrimLang, complaintText);
      if (polished) usedClaude = true;
    } catch (e) {
      Logger.log('Claude polish failed: ' + e.message);
    }

    // 2) Fallback إلى Wrapper بسيط
    if (!polished) polished = wrapReplyText_(text, pilgrimName);

    // 3) احفظ النص المنسَّق في الكاش لإرساله بعد الموافقة
    setCache_('tm_preview_' + opChatId, {
      ticketId: ticketId,
      targetChatId: targetChatId,
      wrappedText: polished,
      rawText: text,
      usedClaude: usedClaude,
      pilgrimLang: pilgrimLang
    }, 600);

    // 4) أعرض المعاينة + 3 أزرار
    var preview = '👁️ <b>معاينة الرسالة قبل الإرسال</b>\n';
    preview += '🎫 ' + ticketId + ' — ' + (pilgrimName || '-') + '\n';
    preview += '🌐 لغة الحاج: <b>' + pilgrimLang + '</b>';
    preview += usedClaude ? ' | 🤖 Claude polish' : ' | 📝 Simple wrap';
    preview += '\n━━━━━━━━━━━━━━\n';
    preview += polished;
    preview += '\n━━━━━━━━━━━━━━';

    var buttons = [
      [
        { text: '✅ أرسل', callback_data: 'tm_send:' + ticketId },
        { text: '✏️ عدّل', callback_data: 'tm_edit:' + ticketId }
      ],
      [
        { text: '❌ إلغاء', callback_data: 'tm_cancel:' + ticketId }
      ]
    ];

    sendMessageSafe_(opChatId, preview, { inline_keyboard: buttons });
    CacheService.getScriptCache().remove('tm_pending_' + opChatId);
    return;

  } else if (action === 'closing') {
    try {
      var ticket2 = getTicketById_(ticketId);
      if (ticket2) {
        updateTicketResolved_(ticket2, text);
        sendMessageSafe_(opChatId, '✅ تم إغلاق البلاغ ' + ticketId + ' بالملاحظة:\n«' + text + '»');
      }
    } catch (e) {
      sendMessageSafe_(opChatId, '❌ فشل الإغلاق: ' + e.message);
    }
  }

  // مسح الكاش بعد المعالجة
  CacheService.getScriptCache().remove('tm_pending_' + opChatId);
}

// ============================================
// أزرار المعاينة (send / edit / cancel)
// ============================================
function tmActionSend_(opChatId, ticket) {
  var preview = getCache_('tm_preview_' + opChatId);
  if (!preview) {
    sendMessageSafe_(opChatId, '⚠️ المعاينة منتهية الصلاحية. أعد كتابة الرد.');
    return;
  }

  try {
    sendMessageSafe_(preview.targetChatId, preview.wrappedText);
    sendMessageSafe_(opChatId, '✅ تم إرسال الرد للحاج <b>' + ticket.row[3] + '</b> (' + preview.ticketId + ')');
    updateTicketResolved_(ticket, 'Wrapped reply: ' + preview.rawText.substring(0, 80));
  } catch (e) {
    sendMessageSafe_(opChatId, '❌ فشل الإرسال: ' + e.message);
  }
  CacheService.getScriptCache().remove('tm_preview_' + opChatId);
}

function tmActionEdit_(opChatId, ticket) {
  var preview = getCache_('tm_preview_' + opChatId);
  CacheService.getScriptCache().remove('tm_preview_' + opChatId);

  setCache_('tm_pending_' + opChatId, {
    action: 'custom_reply',
    ticketId: ticket.row[0],
    targetChatId: ticket.row[1]
  }, 600);

  var oldText = preview ? preview.rawText : '';
  var msg = '✏️ <b>أعد كتابة الرسالة — ' + ticket.row[0] + '</b>\n';
  msg += '👤 الحاج: ' + ticket.row[3] + '\n';
  if (oldText) msg += '\n📝 النص السابق:\n<code>' + oldText + '</code>\n';
  msg += '\nاكتب النص الجديد:';
  sendMessageSafe_(opChatId, msg);
}

function tmActionCancel_(opChatId, ticket) {
  CacheService.getScriptCache().remove('tm_preview_' + opChatId);
  CacheService.getScriptCache().remove('tm_pending_' + opChatId);
  sendMessageSafe_(opChatId, '❌ تم إلغاء الرد للبلاغ ' + ticket.row[0]);
}
