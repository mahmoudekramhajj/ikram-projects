/**
 * ClaudeAPI.js — نقطة HTTP للتشغيل عن بُعد
 *
 * نقاط الوصول:
 *   GET ?action=ping&key=ekram2026claude
 *   GET ?action=list&key=ekram2026claude
 *   GET ?action=run&fn=FUNCTION_NAME&key=ekram2026claude
 *   GET ?action=run&fn=FUNCTION_NAME&args=[1,2]&key=ekram2026claude
 *   GET ?action=log&key=ekram2026claude
 *
 * ملاحظة: المفتاح ثابت هنا للتوافق مع باقي المشاريع.
 * لا توضع أي أسرار أخرى (Anthropic key) هنا — تلك في Script Properties.
 */

var CLAUDE_API_KEY = 'ekram2026claude';

function doGet(e) {
  return handleClaudeAPI_(e);
}

function doPost(e) {
  return handleClaudeAPI_(e);
}

function handleClaudeAPI_(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  if (params.key !== CLAUDE_API_KEY) {
    return jsonResponse_({ error: 'Unauthorized', code: 401 });
  }

  var action = params.action || 'ping';

  try {
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
        return jsonResponse_({
          project: getProjectName_(),
          log: Logger.getLog()
        });

      default:
        return jsonResponse_({ error: 'Unknown action: ' + action, code: 400 });
    }
  } catch (err) {
    return jsonResponse_({
      error: err.message,
      stack: err.stack,
      code: 500
    });
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

function listAvailableFunctions_() {
  var functions = [];
  var skip = [
    'doGet', 'doPost', 'onOpen', 'onEdit', 'onInstall',
    'jsonResponse_', 'getProjectName_',
    'listAvailableFunctions_', 'handleRun_',
    'handleClaudeAPI_', 'CLAUDE_API_KEY'
  ];

  for (var name in globalThis) {
    try {
      if (typeof globalThis[name] === 'function' &&
          skip.indexOf(name) === -1 &&
          !name.startsWith('_')) {
        functions.push(name);
      }
    } catch (e) {}
  }
  return functions.sort();
}

function getProjectName_() {
  try {
    return DriveApp.getFileById(ScriptApp.getScriptId()).getName();
  } catch (e) {
    return 'TicketLinker';
  }
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}
