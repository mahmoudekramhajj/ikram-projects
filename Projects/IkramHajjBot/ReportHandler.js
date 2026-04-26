// ============================================
// نظام البلاغات — إبلاغ الحاج عن أخطاء في معلوماته
// ============================================

var REPORTS_SHEET = 'PilgrimReports';
var REPORTS_GROUP = '-1003610412573'; // مجموعة مركز البلاغات

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

// أعمدة شيت التذاكر
var RT = {
  TICKET_ID: 0, CHAT_ID: 1, PASSPORT: 2, NAME: 3,
  SECTION: 4, ISSUE: 5, CORRECTION: 6, STATUS: 7,
  RESOLVED_BY: 8, CREATED_AT: 9, RESOLVED_AT: 10, ADMIN_RESPONSE: 11
};

// ============================================
// شيت التذاكر — إنشاء تلقائي
// ============================================
function getReportsSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(REPORTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REPORTS_SHEET);
    sheet.appendRow(['TicketID', 'ChatID', 'Passport', 'Name', 'Section', 'Issue', 'Correction', 'Status', 'ResolvedBy', 'CreatedAt', 'ResolvedAt', 'AdminResponse']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================
// إنشاء تذكرة بلاغ
// ============================================
function createTicket_(chatId, passport, pilgrimName, section, issueText, correction) {
  try {
    var sheet = getReportsSheet_();
    var data = sheet.getDataRange().getValues();
    var ticketId = 'TKT-' + String(data.length).padStart(4, '0');
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([ticketId, String(chatId), passport, pilgrimName, section, issueText, correction, 'open', '', now, '', '']);
    return ticketId;
  } catch (e) {
    Logger.log('createTicket_ error: ' + e.message);
    return null;
  }
}

// ============================================
// جلب تذاكر حاج معين
// ============================================
function getMyTickets_(passport) {
  try {
    var sheet = getReportsSheet_();
    var data = sheet.getDataRange().getValues();
    var key = String(passport).toUpperCase().trim();
    var tickets = [];

    for (var i = data.length - 1; i >= 1 && tickets.length < 10; i--) {
      if (String(data[i][RT.PASSPORT]).toUpperCase().trim() === key) {
        tickets.push({
          id: String(data[i][RT.TICKET_ID]),
          section: String(data[i][RT.SECTION]),
          issue: String(data[i][RT.ISSUE]),
          correction: String(data[i][RT.CORRECTION]),
          status: String(data[i][RT.STATUS]),
          createdAt: String(data[i][RT.CREATED_AT] || ''),
          resolvedAt: String(data[i][RT.RESOLVED_AT] || ''),
          adminResponse: String(data[i][RT.ADMIN_RESPONSE] || '')
        });
      }
    }
    return tickets;
  } catch (e) {
    Logger.log('getMyTickets_ error: ' + e.message);
    return [];
  }
}

// ============================================
// تحديث حالة تذكرة (حل / رفض) + رد الإدارة
// ============================================
function updateTicketStatus_(ticketId, newStatus, resolvedBy, adminResponse) {
  try {
    var sheet = getReportsSheet_();
    var data = sheet.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');
    var tickets = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][RT.TICKET_ID]) === ticketId) {
        sheet.getRange(i + 1, RT.STATUS + 1).setValue(newStatus);
        sheet.getRange(i + 1, RT.RESOLVED_BY + 1).setValue(resolvedBy);
        sheet.getRange(i + 1, RT.RESOLVED_AT + 1).setValue(now);
        if (adminResponse) {
          sheet.getRange(i + 1, RT.ADMIN_RESPONSE + 1).setValue(adminResponse);
        }
        tickets.push({
          chatId: String(data[i][RT.CHAT_ID]),
          passport: String(data[i][RT.PASSPORT]),
          section: String(data[i][RT.SECTION]),
          issue: String(data[i][RT.ISSUE]),
          correction: String(data[i][RT.CORRECTION])
        });
      }
    }
    return tickets;
  } catch (e) {
    Logger.log('updateTicketStatus_ error: ' + e.message);
    return [];
  }
}

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

  var reportData = getReportDraft_(chatId) || { issues: [] };
  reportData.currentSection = section;
  reportData.currentIssueKey = issueKey;
  saveReportDraft_(chatId, reportData);

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

  reportData.issues.push({
    section: reportData.currentSection,
    issueKey: reportData.currentIssueKey,
    issueText: T_(reportData.currentIssueKey, 'ar'),
    correction: text
  });

  reportData.currentSection = null;
  reportData.currentIssueKey = null;
  saveReportDraft_(chatId, reportData);

  updateSession_(chatId, { inputState: '' });

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

  var ticketIds = [];
  for (var i = 0; i < reportData.issues.length; i++) {
    var issue = reportData.issues[i];
    var ticketId = createTicket_(chatId, session.passport, pilgrimName, issue.section, issue.issueText, issue.correction);
    if (ticketId) ticketIds.push(ticketId);
  }

  var mainId = ticketIds.length > 0 ? ticketIds[0] : 'TKT-????';

  sendMessage_(chatId, T_('report_submitted', lang, { id: mainId, count: reportData.issues.length }), {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });

  sendReportNotification_(chatId, session.passport, pilgrimName, reportData.issues, mainId);

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
// إشعار المدير بالبلاغ الجديد
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
    text += '   ✏️ ' + issue.correction + '\n\n';
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
// المدير ضغط حل/رفض — نطلب منه كتابة رد
// ============================================
function handleReportAction_(adminChatId, callbackData, callbackFrom) {
  var parts = callbackData.split('_');
  var action = parts[1]; // resolve or reject
  var ticketId = parts.slice(2).join('_'); // TKT-0001

  // حفظ الإجراء المعلّق في الكاش
  setCache_('rpt_pending_' + adminChatId, {
    action: action,
    ticketId: ticketId,
    resolvedBy: callbackFrom.first_name || 'Admin'
  }, 600); // 10 دقائق

  var prompt = (action === 'resolve')
    ? '✅ <b>بلاغ ' + ticketId + '</b>\n\nاكتب <b>رد الإدارة</b> للحاج:'
    : '❌ <b>بلاغ ' + ticketId + '</b>\n\nاكتب <b>سبب الرفض</b> للحاج:';

  sendMessage_(adminChatId, prompt, {
    inline_keyboard: [[{ text: '⏩ بدون رد', callback_data: 'rpt_skip_' + ticketId }]]
  });
}

// ============================================
// المدير أرسل نص الرد — تنفيذ الحل/الرفض
// ============================================
function handleReportAdminInput_(adminChatId, responseText, pending) {
  var newStatus = (pending.action === 'resolve') ? 'resolved' : 'rejected';

  // تحديث الشيت
  var tickets = updateTicketStatus_(pending.ticketId, newStatus, pending.resolvedBy, responseText);

  // مسح الإجراء المعلّق
  CacheService.getScriptCache().remove('rpt_pending_' + adminChatId);

  // تأكيد للمدير
  var statusIcon = (newStatus === 'resolved') ? '✅' : '❌';
  sendMessage_(adminChatId, statusIcon + ' تم تحديث البلاغ <b>' + pending.ticketId + '</b>\n💬 الرد: ' + responseText);

  // إشعار الحاج
  if (tickets && tickets.length > 0) {
    notifyPilgrimResolution_(tickets, pending.ticketId, newStatus, responseText);
  }
}

// ============================================
// المدير ضغط "بدون رد" — تنفيذ بدون نص
// ============================================
function handleReportSkip_(adminChatId, callbackData) {
  var ticketId = callbackData.replace('rpt_skip_', '');
  var pending = getCache_('rpt_pending_' + adminChatId);

  if (!pending) {
    sendMessage_(adminChatId, '⚠️ انتهت صلاحية الإجراء. اضغط الزر مرة أخرى.');
    return;
  }

  var newStatus = (pending.action === 'resolve') ? 'resolved' : 'rejected';
  var tickets = updateTicketStatus_(pending.ticketId, newStatus, pending.resolvedBy, '');

  CacheService.getScriptCache().remove('rpt_pending_' + adminChatId);

  var statusIcon = (newStatus === 'resolved') ? '✅' : '❌';
  sendMessage_(adminChatId, statusIcon + ' تم تحديث البلاغ <b>' + pending.ticketId + '</b>');

  if (tickets && tickets.length > 0) {
    notifyPilgrimResolution_(tickets, pending.ticketId, newStatus, '');
  }
}

// ============================================
// إرسال إشعار الحل/الرفض للحاج
// ============================================
function notifyPilgrimResolution_(tickets, ticketId, newStatus, adminResponse) {
  var chatId = tickets[0].chatId;
  var session = getSession_(chatId);
  var lang = (session && session.language) ? session.language : 'ar';

  if (newStatus === 'resolved') {
    var details = '';
    for (var i = 0; i < tickets.length; i++) {
      var icon = SECTION_ICONS[tickets[i].section] || '📋';
      details += icon + ' ' + tickets[i].issue + ' — ✅\n';
    }

    var text = T_('report_resolved_notify', lang, { id: ticketId, details: details });
    if (adminResponse) {
      text += '\n\n💬 <b>' + T_('lbl_admin_response', lang) + ':</b>\n' + adminResponse;
    }

    sendMessage_(chatId, text, {
      inline_keyboard: [
        [{ text: T_('btn_my_reports', lang), callback_data: 'my_reports' }],
        [{ text: T_('btn_back', lang), callback_data: 'show_menu' }]
      ]
    });
  } else {
    var text = T_('report_rejected_notify', lang, { id: ticketId });
    if (adminResponse) {
      text += '\n\n💬 <b>' + T_('lbl_reject_reason', lang) + ':</b>\n' + adminResponse;
    }

    sendMessage_(chatId, text, {
      inline_keyboard: [
        [{ text: T_('btn_my_reports', lang), callback_data: 'my_reports' }],
        [{ text: T_('btn_whatsapp', lang), url: 'https://wa.me/966125111940' }],
        [{ text: T_('btn_back', lang), callback_data: 'show_menu' }]
      ]
    });
  }
}

// ============================================
// 📋 بلاغاتي — عرض البلاغات مع التفاصيل والرد
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
    var icon = SECTION_ICONS[t.section] || '📋';
    var statusKey = 'report_status_' + t.status;
    var statusText = T_(statusKey, lang);
    if (statusText === statusKey) statusText = t.status;

    text += '<b>' + t.id + '</b> ' + icon + ' ' + t.issue + '\n';
    text += '📊 ' + statusText + '\n';
    text += '✏️ ' + t.correction + '\n';

    // رد الإدارة إن وُجد
    var resp = t.adminResponse;
    if (resp && resp !== '' && resp !== 'undefined' && resp !== 'null') {
      if (t.status === 'resolved') {
        text += '💬 <b>' + T_('lbl_admin_response', lang) + ':</b> ' + resp + '\n';
      } else if (t.status === 'rejected') {
        text += '💬 <b>' + T_('lbl_reject_reason', lang) + ':</b> ' + resp + '\n';
      }
    }

    // التاريخ
    var dateStr = (t.resolvedAt && t.resolvedAt !== '' && t.resolvedAt !== 'undefined')
      ? t.resolvedAt.substring(0, 16)
      : t.createdAt.substring(0, 16);
    text += '📅 ' + dateStr + '\n\n';
  }

  sendMessage_(chatId, text, {
    inline_keyboard: [[{ text: T_('btn_back', lang), callback_data: 'show_menu' }]]
  });
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
