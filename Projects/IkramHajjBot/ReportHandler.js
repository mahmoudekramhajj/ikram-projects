// ============================================
// نظام البلاغات — إبلاغ الحاج عن أخطاء في معلوماته
// ============================================

// خريطة المشاكل الشائعة لكل قسم
var REPORT_ISSUES = {
  flight:    ['report_flight_date', 'report_flight_number', 'report_flight_airline', 'report_other'],
  hotel:     ['report_hotel_name', 'report_hotel_dates', 'report_hotel_room', 'report_other'],
  visa:      ['report_visa_status', 'report_ticket_number', 'report_ticket_link', 'report_other'],
  package:   ['report_package_type', 'report_package_details', 'report_other'],
  transport: ['report_transport_time', 'report_transport_point', 'report_other'],
  personal:  ['report_name_wrong', 'report_passport_wrong', 'report_nationality_wrong', 'report_other']
};

// أيقونات الأقسام
var SECTION_ICONS = {
  flight: '✈️', hotel: '🏨', visa: '🎫',
  package: '📦', transport: '🚌', personal: '📋'
};

// ============================================
// عرض قائمة الأقسام
// ============================================
function handleReportError_(chatId, session) {
  var lang = session.language || 'ar';

  var buttons = [
    [{ text: T_('report_sec_flight', lang),   callback_data: 'report_sec_flight' },
     { text: T_('report_sec_hotel', lang),    callback_data: 'report_sec_hotel' }],
    [{ text: T_('report_sec_visa', lang),     callback_data: 'report_sec_visa' },
     { text: T_('report_sec_package', lang),  callback_data: 'report_sec_package' }],
    [{ text: T_('report_sec_transport', lang), callback_data: 'report_sec_transport' },
     { text: T_('report_sec_personal', lang), callback_data: 'report_sec_personal' }],
    [{ text: T_('btn_cancel_report', lang),   callback_data: 'report_cancel' }]
  ];

  sendMessage_(chatId, T_('report_select_section', lang), { inline_keyboard: buttons });
}

// ============================================
// عرض المشاكل الشائعة للقسم المختار
// ============================================
function handleReportSection_(chatId, session, section) {
  var lang = session.language || 'ar';
  var issues = REPORT_ISSUES[section];
  if (!issues) return;

  var buttons = [];
  for (var i = 0; i < issues.length; i++) {
    buttons.push([{ text: T_(issues[i], lang), callback_data: 'report_issue_' + section + '_' + issues[i] }]);
  }
  buttons.push([{ text: T_('btn_cancel_report', lang), callback_data: 'report_cancel' }]);

  sendMessage_(chatId, T_('report_select_issue', lang), { inline_keyboard: buttons });
}

// ============================================
// الحاج اختار مشكلة — نطلب منه المعلومة الصحيحة
// ============================================
function handleReportIssueSelected_(chatId, session, section, issueKey) {
  var lang = session.language || 'ar';

  // حفظ المشكلة الحالية في الجلسة
  var reportData = getReportDraft_(chatId) || { issues: [] };
  reportData.currentSection = section;
  reportData.currentIssueKey = issueKey;
  saveReportDraft_(chatId, reportData);

  // طلب المعلومة الصحيحة
  var prompt = (issueKey === 'report_other')
    ? T_('report_write_other', lang)
    : T_('report_write_correct', lang);

  updateSession_(chatId, { inputState: 'awaiting_report_correction' });
  sendMessage_(chatId, prompt);
}

// ============================================
// الحاج أرسل المعلومة الصحيحة — تسجيل المشكلة
// ============================================
function handleReportCorrectionInput_(chatId, text, session) {
  var lang = session.language || 'ar';
  var reportData = getReportDraft_(chatId);

  if (!reportData || !reportData.currentSection) {
    updateSession_(chatId, { inputState: '' });
    sendMainMenu_(chatId, lang);
    return;
  }

  // إضافة المشكلة لقائمة المشاكل
  reportData.issues.push({
    section: reportData.currentSection,
    issueKey: reportData.currentIssueKey,
    issueText: T_(reportData.currentIssueKey, 'ar'), // نحفظ النص العربي دائماً
    correction: text
  });

  // مسح الحالة المؤقتة
  reportData.currentSection = null;
  reportData.currentIssueKey = null;
  saveReportDraft_(chatId, reportData);

  updateSession_(chatId, { inputState: '' });

  // عرض خيارات: إضافة مشكلة أخرى أو إرسال
  var buttons = [
    [{ text: T_('btn_add_more', lang), callback_data: 'report_error' }],
    [{ text: T_('btn_submit_report', lang), callback_data: 'report_submit' }],
    [{ text: T_('btn_cancel_report', lang), callback_data: 'report_cancel' }]
  ];

  sendMessage_(chatId, T_('report_issue_saved', lang), { inline_keyboard: buttons });
}

// ============================================
// إرسال البلاغ النهائي
// ============================================
function handleReportSubmit_(chatId, session) {
  var lang = session.language || 'ar';
  var reportData = getReportDraft_(chatId);

  if (!reportData || !reportData.issues || reportData.issues.length === 0) {
    sendMessage_(chatId, T_('report_cancelled', lang));
    sendMainMenu_(chatId, lang);
    clearReportDraft_(chatId);
    return;
  }

  var pilgrim = findPilgrimByPassport_(session.passport);
  var pilgrimName = pilgrim ? pilgrim.name : '-';

  // إنشاء التذاكر في الشيت
  var ticketIds = [];
  for (var i = 0; i < reportData.issues.length; i++) {
    var issue = reportData.issues[i];
    var ticketId = createTicket_(chatId, session.passport, pilgrimName, issue.section, issue.issueText, issue.correction);
    if (ticketId) ticketIds.push(ticketId);
  }

  var mainId = ticketIds.length > 0 ? ticketIds[0] : 'TKT-????';

  // إرسال تأكيد للحاج
  sendMessage_(chatId, T_('report_submitted', lang, { id: mainId, count: reportData.issues.length }), {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });

  // إرسال إشعار لمجموعة الإدارة
  sendReportNotification_(chatId, session.passport, pilgrimName, reportData.issues, mainId);

  // مسح المسودة
  clearReportDraft_(chatId);
}

// ============================================
// إلغاء البلاغ
// ============================================
function handleReportCancel_(chatId, session) {
  var lang = session.language || 'ar';
  clearReportDraft_(chatId);
  updateSession_(chatId, { inputState: '' });
  sendMessage_(chatId, T_('report_cancelled', lang));
  sendMainMenu_(chatId, lang);
}

// ============================================
// عرض بلاغات الحاج
// ============================================
function handleMyReports_(chatId, session) {
  var lang = session.language || 'ar';
  var tickets = getMyTickets_(session.passport);

  if (!tickets || tickets.length === 0) {
    sendMessage_(chatId, T_('my_reports_title', lang) + '\n\n' + T_('my_reports_empty', lang), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
    return;
  }

  var text = T_('my_reports_title', lang) + '\n\n';

  for (var i = 0; i < tickets.length; i++) {
    var t = tickets[i];
    var statusKey = 'report_status_' + t.status;
    var icon = SECTION_ICONS[t.section] || '📋';

    text += '<b>#' + t.id + '</b> ' + T_(statusKey, lang) + '\n';
    text += icon + ' ' + t.issue + '\n';
    if (t.correction) text += '✏️ ' + t.correction + '\n';
    text += '🕐 ' + t.createdAt + '\n\n';
  }

  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
}

// ============================================
// إشعار مجموعة الإدارة بالبلاغ الجديد
// ============================================
function sendReportNotification_(chatId, passport, pilgrimName, issues, mainId) {
  if (!REPORTS_GROUP) return;

  var text = '🔔 <b>بلاغ جديد #' + mainId + '</b>\n' +
    '━━━━━━━━━━━━━━\n' +
    '👤 ' + pilgrimName + ' | <code>' + passport + '</code>\n\n';

  for (var i = 0; i < issues.length; i++) {
    var issue = issues[i];
    var icon = SECTION_ICONS[issue.section] || '📋';
    text += (i + 1) + '️⃣ ' + icon + ' ' + issue.issueText + '\n';
    text += '   ✏️ الصحيح: ' + issue.correction + '\n\n';
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm');
  text += '🕐 ' + now;

  var buttons = {
    inline_keyboard: [
      [{ text: '✅ تم الحل', callback_data: 'rpt_resolve_' + mainId },
       { text: '❌ مرفوض', callback_data: 'rpt_reject_' + mainId }]
    ]
  };

  sendMessage_(REPORTS_GROUP, text, buttons);
}

// ============================================
// معالجة رد الإدارة (حل/رفض) من المجموعة
// ============================================
function handleReportAction_(callbackData, callbackFrom) {
  var parts = callbackData.split('_');
  // rpt_resolve_TKT-0001 or rpt_reject_TKT-0001
  var action = parts[1]; // resolve or reject
  var ticketId = parts.slice(2).join('_'); // TKT-0001

  var resolvedBy = callbackFrom.first_name || 'Admin';
  var newStatus = (action === 'resolve') ? 'resolved' : 'rejected';

  // تحديث كل التذاكر المرتبطة بنفس mainId
  var tickets = updateTicketStatus_(ticketId, newStatus, resolvedBy);
  if (!tickets || tickets.length === 0) return;

  // جلب معلومات الحاج لإبلاغه
  var chatId = tickets[0].chatId;
  var passport = tickets[0].passport;

  // جلب لغة الحاج
  var session = getSession_(chatId);
  var lang = (session && session.language) ? session.language : 'ar';

  if (newStatus === 'resolved') {
    var details = '';
    for (var i = 0; i < tickets.length; i++) {
      var icon = SECTION_ICONS[tickets[i].section] || '📋';
      details += icon + ' ' + tickets[i].issue + ' — ✅\n';
    }
    sendMessage_(chatId, T_('report_resolved_notify', lang, { id: ticketId, details: details }), {
      inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
    });
  } else {
    sendMessage_(chatId, T_('report_rejected_notify', lang, { id: ticketId }), {
      inline_keyboard: [
        [{ text: T_('btn_whatsapp', lang), url: 'https://wa.me/966125111940' }],
        [{ text: T_('btn_back', lang), callback_data: 'show_menu' }]
      ]
    });
  }
}

// ============================================
// إدارة المسودة (draft) — حفظ في الكاش
// ============================================
function getReportDraft_(chatId) {
  return getCache_('report_draft_' + chatId);
}

function saveReportDraft_(chatId, data) {
  setCache_('report_draft_' + chatId, data, 1800); // 30 دقيقة
}

function clearReportDraft_(chatId) {
  var cache = CacheService.getScriptCache();
  cache.remove('report_draft_' + chatId);
}
