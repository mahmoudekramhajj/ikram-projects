/**
 * ClaudeAPI.js — نظام التشغيل عن بعد
 * يسمح لـ Claude بتشغيل أي دالة واستقبال النتيجة عبر HTTP
 *
 * الاستخدام (عبر doGet في Router.js):
 * GET ?action=run&fn=functionName&key=SECRET_KEY
 * GET ?action=list&key=SECRET_KEY
 * GET ?action=ping&key=SECRET_KEY
 * GET ?action=tickets&key=SECRET_KEY
 * GET ?action=updateTicket&key=SECRET_KEY&id=TKT-0001&status=resolved&reply=TEXT&by=Claude
 */

var CLAUDE_API_KEY = 'ekram2026claude';

/**
 * نقطة الدخول — يستدعيها doGet في Router.js عند وجود key
 */
function handleClaudeAPI_(e) {
  var params = e ? e.parameter : {};

  if (params.key !== CLAUDE_API_KEY) {
    return jsonResponse_({ error: 'Unauthorized', code: 401 });
  }

  var action = params.action || 'ping';

  switch (action) {
    case 'ping':
      return jsonResponse_({
        status: 'ok',
        project: getProjectName_(),
        timestamp: new Date().toISOString(),
        timezone: Session.getScriptTimeZone()
      });

    case 'list':
      return jsonResponse_({
        project: getProjectName_(),
        functions: listAvailableFunctions_()
      });

    case 'run':
      return handleRun_(params);

    case 'log':
      return handleLog_(params);

    case 'tickets':
      return handleGetTickets_(params);

    case 'updateTicket':
      return handleUpdateTicket_(params);

    case 'findChat':
      return handleFindChat_(params);

    case 'searchPresonal':
      return handleSearchPresonal_(params);

    default:
      return jsonResponse_({ error: 'Unknown action: ' + action, code: 400 });
  }
}

// ============================================
// جلب جميع التذاكر (أو تصفية بالحالة)
// GET ?action=tickets&key=...&status=open
// ============================================
function handleGetTickets_(params) {
  var sheet = getReportsSheet_();
  var data = sheet.getDataRange().getValues();
  var filterStatus = params.status || '';
  var tickets = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var ticket = {
      ticketId:      String(row[RT.TICKET_ID]),
      chatId:        String(row[RT.CHAT_ID]),
      passport:      String(row[RT.PASSPORT]),
      name:          String(row[RT.NAME]),
      section:       String(row[RT.SECTION]),
      issue:         String(row[RT.ISSUE]),
      correction:    String(row[RT.CORRECTION]),
      status:        String(row[RT.STATUS]),
      resolvedBy:    String(row[RT.RESOLVED_BY]),
      createdAt:     String(row[RT.CREATED_AT] || ''),
      resolvedAt:    String(row[RT.RESOLVED_AT] || ''),
      adminResponse: String(row[RT.ADMIN_RESPONSE] || '')
    };

    if (!filterStatus || ticket.status === filterStatus) {
      tickets.push(ticket);
    }
  }

  return jsonResponse_({ success: true, count: tickets.length, tickets: tickets });
}

// ============================================
// تحديث تذكرة واحدة أو عدة تذاكر
// GET ?action=updateTicket&key=...&id=TKT-0001&status=resolved&reply=TEXT&by=Claude
// لعدة تذاكر: id=TKT-0001,TKT-0002,TKT-0003
// ============================================
function handleUpdateTicket_(params) {
  var ids = (params.id || '').split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  var newStatus = params.status || 'resolved';
  var reply = params.reply || '';
  var by = params.by || 'Claude';

  if (ids.length === 0) {
    return jsonResponse_({ error: 'Missing id parameter', code: 400 });
  }

  var sheet = getReportsSheet_();
  var data = sheet.getDataRange().getValues();
  var now = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss');
  var updated = [];

  for (var i = 1; i < data.length; i++) {
    var ticketId = String(data[i][RT.TICKET_ID]);
    if (ids.indexOf(ticketId) !== -1) {
      sheet.getRange(i + 1, RT.STATUS + 1).setValue(newStatus);
      sheet.getRange(i + 1, RT.RESOLVED_BY + 1).setValue(by);
      sheet.getRange(i + 1, RT.RESOLVED_AT + 1).setValue(now);
      if (reply) {
        sheet.getRange(i + 1, RT.ADMIN_RESPONSE + 1).setValue(reply);
      }
      updated.push({
        ticketId: ticketId,
        chatId: String(data[i][RT.CHAT_ID]),
        name: String(data[i][RT.NAME]),
        status: newStatus
      });
    }
  }

  return jsonResponse_({ success: true, updated: updated, count: updated.length });
}

// ============================================
// البحث عن ChatId بالجواز في BotSessions
// GET ?action=findChat&key=...&passport=CGN4W8CP4
// ============================================
function handleFindChat_(params) {
  var passport = (params.passport || '').toUpperCase().trim();
  if (!passport) {
    return jsonResponse_({ error: 'Missing passport parameter', code: 400 });
  }

  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BotSessions');
  if (!sheet) {
    return jsonResponse_({ error: 'BotSessions sheet not found', code: 404 });
  }

  var data = sheet.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).toUpperCase().trim() === passport) {
      results.push({
        chatId: String(data[i][0]),
        passport: String(data[i][1]),
        language: String(data[i][3]),
        authStatus: String(data[i][4])
      });
    }
  }

  return jsonResponse_({ success: true, count: results.length, sessions: results });
}

// ============================================
// البحث في Presonal Details بالجواز
// GET ?action=searchPresonal&key=...&passport=535375711
// ============================================
function handleSearchPresonal_(params) {
  var passport = (params.passport || '').toUpperCase().trim();
  if (!passport) {
    return jsonResponse_({ error: 'Missing passport parameter', code: 400 });
  }

  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PERSONAL_SHEET);
  if (!sheet) {
    return jsonResponse_({ error: PERSONAL_SHEET + ' sheet not found', code: 404 });
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][PD.PASSPORT]).toUpperCase().trim() === passport) {
      var result = {
        row: i + 1,
        seq: String(data[i][PD.SEQ]),
        group: String(data[i][PD.GROUP]),
        type: String(data[i][PD.TYPE]),
        category: String(data[i][PD.CATEGORY]),
        passport: String(data[i][PD.PASSPORT]),
        nameAr: String(data[i][PD.FIRST_NAME_AR]) + ' ' + String(data[i][PD.LAST_NAME_AR]),
        nameEn: String(data[i][PD.FIRST_NAME_EN]) + ' ' + String(data[i][PD.LAST_NAME_EN]),
        email: String(data[i][PD.EMAIL]),
        phone: String(data[i][PD.PHONE]),
        guideName: String(data[i][PD.GUIDE_NAME]),
        packageNo: String(data[i][PD.PACKAGE_NO]),
        packageName: String(data[i][PD.PACKAGE_NAME]),
        flightType: String(data[i][PD.FLIGHT_TYPE]),
        contractName: String(data[i][PD.CONTRACT_NAME]),
        visaStatus: String(data[i][PD.VISA_STATUS]),
        ticketNo: String(data[i][PD.TICKET_NO]),
        ticketUrl: String(data[i][PD.TICKET_URL]),
        invoiceNo: String(data[i][PD.INVOICE_NO]),
        camp: String(data[i][PD.CAMP])
      };
      return jsonResponse_({ success: true, result: result });
    }
  }

  return jsonResponse_({ success: true, result: null, message: 'Passport not found in ' + PERSONAL_SHEET });
}

// ============================================
// تشغيل دالة عامة
// ============================================
function handleRun_(params) {
  var fnName = params.fn;
  if (!fnName) {
    return jsonResponse_({ error: 'Missing fn parameter', code: 400 });
  }

  var fn = globalThis[fnName];
  if (typeof fn !== 'function') {
    return jsonResponse_({ error: 'Function not found: ' + fnName, code: 404 });
  }

  var args = [];
  if (params.args) {
    try {
      args = JSON.parse(params.args);
      if (!Array.isArray(args)) args = [args];
    } catch (e) {
      args = [params.args];
    }
  }

  var startTime = new Date();
  var result, error, logs;

  try {
    Logger.clear();
    result = fn.apply(null, args);
    logs = Logger.getLog();
  } catch (e) {
    error = { message: e.message, stack: e.stack, name: e.name };
    logs = Logger.getLog();
  }

  var duration = (new Date() - startTime) / 1000;
  var response = {
    project: getProjectName_(),
    function: fnName,
    duration: duration + 's',
    timestamp: new Date().toISOString()
  };

  if (error) { response.success = false; response.error = error; }
  else { response.success = true; response.result = result; }

  if (logs) {
    response.logs = logs.split('\n').filter(function(l) { return l.length > 0; });
  }

  return jsonResponse_(response);
}

function handleLog_(params) {
  return jsonResponse_({ project: getProjectName_(), log: Logger.getLog() });
}

function listAvailableFunctions_() {
  var functions = [];
  var skip = ['doGet', 'doPost', 'onOpen', 'onEdit', 'onInstall',
              'include', 'jsonResponse_', 'getProjectName_',
              'listAvailableFunctions_', 'handleRun_', 'handleLog_',
              'handleClaudeAPI_', 'handleGetTickets_', 'handleUpdateTicket_',
              'registerProject_'];

  for (var name in globalThis) {
    try {
      if (typeof globalThis[name] === 'function' &&
          skip.indexOf(name) === -1 &&
          !name.startsWith('_') &&
          name !== 'CLAUDE_API_KEY') {
        functions.push(name);
      }
    } catch (e) {}
  }
  return functions.sort();
}

function getProjectName_() {
  try { return DriveApp.getFileById(ScriptApp.getScriptId()).getName(); }
  catch (e) { return 'Unknown Project'; }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
