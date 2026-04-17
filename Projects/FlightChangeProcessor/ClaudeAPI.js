/**
 * ClaudeAPI.js — نظام التشغيل عن بعد
 * يسمح لـ Claude بتشغيل أي دالة واستقبال النتيجة عبر HTTP
 *
 * الاستخدام:
 * GET ?action=run&fn=functionName&key=SECRET_KEY
 * GET ?action=list&key=SECRET_KEY  (قائمة الدوال المتاحة)
 * GET ?action=ping&key=SECRET_KEY  (فحص الاتصال)
 */

var CLAUDE_API_KEY = 'ekram2026claude';

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

    default:
      return jsonResponse_({ error: 'Unknown action: ' + action, code: 400 });
  }
}

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

function handleLog_(params) {
  return jsonResponse_({
    project: getProjectName_(),
    log: Logger.getLog()
  });
}

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
    } catch (e) {}
  }

  return functions.sort();
}

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

function getProjectName_() {
  try {
    return DriveApp.getFileById(ScriptApp.getScriptId()).getName();
  } catch (e) {
    return 'FlightChangeProcessor';
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return handleClaudeAPI_(e);
}

function doPost(e) {
  return handleClaudeAPI_(e);
}
