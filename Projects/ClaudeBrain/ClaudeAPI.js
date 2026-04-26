/**
 * ClaudeAPI.js — نظام التشغيل عن بعد
 * GET ?action=ping&key=SECRET — فحص الاتصال
 * GET ?action=list&key=SECRET — قائمة الدوال
 * GET ?action=run&fn=NAME&key=SECRET — تشغيل دالة
 */

function handleClaudeAPI_(e) {
  var params = e ? e.parameter : {};
  var config = getConfig_();

  if (params.key !== config.CLAUDE_API_REMOTE_KEY) {
    return jsonResponse_({ error: 'Unauthorized', code: 401 });
  }

  var action = params.action || 'ping';

  switch (action) {
    case 'ping':
      return jsonResponse_({
        status: 'ok',
        project: 'ClaudeBrain',
        version: BRAIN_VERSION,
        timestamp: new Date().toISOString(),
        timezone: Session.getScriptTimeZone()
      });

    case 'list':
      return jsonResponse_({
        project: 'ClaudeBrain',
        functions: listAvailableFunctions_()
      });

    case 'run':
      return handleRun_(params);

    default:
      return jsonResponse_({ error: 'Unknown action: ' + action, code: 400 });
  }
}

function handleRun_(params) {
  var fnName = params.fn;
  if (!fnName) return jsonResponse_({ error: 'Missing fn', code: 400 });

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
    project: 'ClaudeBrain',
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
  var skip = ['doGet', 'doPost', 'handleClaudeAPI_', 'handleRun_',
              'listAvailableFunctions_', 'jsonResponse_', 'errorResponse_',
              'getConfig_', 'getProp_', 'logEvent_', 'getHeader_',
              'computeHMAC_', 'verifyHMAC_', 'constantTimeEquals_',
              'callClaude_', 'processSuccessResponse_', 'calculateCost_',
              'checkBudget_', 'recordSpend_', 'getSpendForKey_',
              'getDailyKey_', 'getHourlyKey_', 'pad2_',
              'handleEvent_', 'handlePing_', 'handleEcho_',
              'getIdentityPrompt_'];

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
