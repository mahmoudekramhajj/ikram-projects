/**
 * ClaudeAPI.js — نظام التشغيل عن بعد
 * يسمح لـ Claude بتشغيل أي دالة واستقبال النتيجة عبر HTTP
 *
 * الاستخدام:
 * GET ?action=run&fn=functionName&key=SECRET_KEY
 * GET ?action=list&key=SECRET_KEY  (قائمة الدوال المتاحة)
 * GET ?action=ping&key=SECRET_KEY  (فحص الاتصال)
 */

// المفتاح السري — يجب تغييره لكل مشروع أو استخدام واحد موحّد
var CLAUDE_API_KEY = 'ekram2026claude';

/**
 * معالج طلبات API — يُستدعى من doGet الرئيسي
 * في المشاريع التي لها doGet خاص (Web Apps)، يُضاف في doGet:
 *   if (params.action && params.key) return handleClaudeAPI_(e);
 * في المشاريع بدون doGet، يُستخدم doGet مباشرة
 */
function handleClaudeAPI_(e) {
  var params = e ? e.parameter : {};

  // التحقق من المفتاح
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

    default:
      return jsonResponse_({ error: 'Unknown action: ' + action, code: 400 });
  }
}

/**
 * تشغيل دالة وإرجاع نتيجتها
 */
function handleRun_(params) {
  var fnName = params.fn;
  if (!fnName) {
    return jsonResponse_({ error: 'Missing fn parameter', code: 400 });
  }

  // التحقق من وجود الدالة
  var fn = globalThis[fnName];
  if (typeof fn !== 'function') {
    return jsonResponse_({ error: 'Function not found: ' + fnName, code: 404 });
  }

  // تحضير المعاملات
  var args = [];
  if (params.args) {
    try {
      args = JSON.parse(params.args);
      if (!Array.isArray(args)) args = [args];
    } catch (e) {
      args = [params.args];
    }
  }

  // تشغيل الدالة مع التقاط الأخطاء والـ Logger
  var startTime = new Date();
  var result, error, logs;

  try {
    // التقاط Logger output
    Logger.clear();
    result = fn.apply(null, args);
    logs = Logger.getLog();
  } catch (e) {
    error = {
      message: e.message,
      stack: e.stack,
      name: e.name
    };
    logs = Logger.getLog();
  }

  var duration = (new Date() - startTime) / 1000;

  var response = {
    project: getProjectName_(),
    function: fnName,
    duration: duration + 's',
    timestamp: new Date().toISOString()
  };

  if (error) {
    response.success = false;
    response.error = error;
  } else {
    response.success = true;
    response.result = result;
  }

  if (logs) {
    response.logs = logs.split('\n').filter(function(l) { return l.length > 0; });
  }

  return jsonResponse_(response);
}

/**
 * قراءة آخر سجلات التشغيل
 */
function handleLog_(params) {
  return jsonResponse_({
    project: getProjectName_(),
    log: Logger.getLog()
  });
}

/**
 * قائمة الدوال المتاحة للتشغيل
 */
function listAvailableFunctions_() {
  var functions = [];
  var skip = ['doGet', 'doPost', 'onOpen', 'onEdit', 'onInstall',
              'include', 'jsonResponse_', 'getProjectName_',
              'listAvailableFunctions_', 'handleRun_', 'handleLog_',
              'registerProject_'];

  for (var name in globalThis) {
    try {
      if (typeof globalThis[name] === 'function' &&
          skip.indexOf(name) === -1 &&
          !name.startsWith('_') &&
          name !== 'CLAUDE_API_KEY') {
        functions.push(name);
      }
    } catch (e) {
      // تخطي الدوال المحمية
    }
  }

  return functions.sort();
}

/**
 * تسجيل المشروع في شيت مركزي
 */
function registerProject_() {
  var ss = SpreadsheetApp.openById('1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s');
  var sheet = ss.getSheetByName('Claude Registry');

  if (!sheet) {
    sheet = ss.insertSheet('Claude Registry');
    sheet.appendRow(['المشروع', 'Script ID', 'Web App URL', 'آخر تسجيل', 'الحالة']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  var projectName = getProjectName_();
  var scriptId = ScriptApp.getScriptId();
  var url = ScriptApp.getService().getUrl();
  var now = new Date();

  // البحث عن صف موجود للتحديث
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === projectName || data[i][1] === scriptId) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[projectName, scriptId, url, now, 'Active']]);
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([projectName, scriptId, url, now, 'Active']);
  }

  Logger.log('Registered: ' + projectName + ' → ' + url);
  return { project: projectName, url: url, status: 'registered' };
}

/**
 * اسم المشروع الحالي
 */
function getProjectName_() {
  try {
    return DriveApp.getFileById(ScriptApp.getScriptId()).getName();
  } catch (e) {
    return 'Unknown Project';
  }
}

/**
 * إرجاع JSON response
 */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
