/**
 * ClaudeBrain — المخ المركزي لمشروع إكرام الضيف
 * Version: 1.0.0 (M1 — Skeleton)
 *
 * نقطة الدخول الرئيسية. يستقبل الأحداث، يتحقق من المصادقة،
 * يفحص الميزانية، ويمرر للمعالجة.
 */

var BRAIN_VERSION = '1.0.0-M1';

/**
 * معالج GET — للفحص السريع والـ ClaudeAPI.js
 */
function doGet(e) {
  // إن كان طلب ClaudeAPI.js (action + key)
  if (e && e.parameter && e.parameter.action) {
    return handleClaudeAPI_(e);
  }

  // افتراضياً: صفحة ترحيب بسيطة
  return jsonResponse_({
    service: 'ClaudeBrain',
    version: BRAIN_VERSION,
    status: 'alive',
    timestamp: new Date().toISOString(),
    message: 'استخدم POST مع X-Signature للتفاعل مع المخ'
  });
}

/**
 * معالج POST — نقطة الدخول الرئيسية للأحداث
 */
function doPost(e) {
  var startTime = new Date();
  var requestId = Utilities.getUuid();

  try {
    // 1. قراءة الجسم
    var rawBody = (e && e.postData && e.postData.contents) || '{}';
    var payload = {};
    try {
      payload = JSON.parse(rawBody);
    } catch (err) {
      return errorResponse_('INVALID_JSON', 'الجسم ليس JSON صالحاً', 400, requestId);
    }

    // 2. التحقق من التوقيع HMAC
    var signature = getHeader_(e, 'X-Signature') || (payload._signature);
    var authResult = verifyHMAC_(rawBody, signature);
    if (!authResult.valid) {
      logEvent_('AUTH_FAILED', { reason: authResult.reason, requestId: requestId });
      return errorResponse_('UNAUTHORIZED', authResult.reason, 401, requestId);
    }

    // 3. فحص الميزانية اليومية
    var budgetOk = checkBudget_();
    if (!budgetOk.ok) {
      logEvent_('BUDGET_EXCEEDED', { spent: budgetOk.spent, limit: budgetOk.limit, requestId: requestId });
      return errorResponse_('BUDGET_EXCEEDED',
        'تم تجاوز الحد اليومي ($' + budgetOk.limit + '). الإنفاق الحالي: $' + budgetOk.spent,
        429, requestId);
    }

    // 4. معالجة الحدث
    var eventType = payload.event_type || 'ping';
    var response = handleEvent_(eventType, payload, requestId);

    // 5. تسجيل النجاح
    var duration = (new Date() - startTime) / 1000;
    logEvent_('REQUEST_OK', {
      requestId: requestId,
      eventType: eventType,
      duration: duration + 's'
    });

    return jsonResponse_({
      success: true,
      requestId: requestId,
      version: BRAIN_VERSION,
      duration: duration + 's',
      data: response
    });

  } catch (err) {
    logEvent_('REQUEST_ERROR', {
      requestId: requestId,
      error: err.message,
      stack: err.stack
    });
    return errorResponse_('INTERNAL_ERROR', err.message, 500, requestId);
  }
}

/**
 * مُوجّه الأحداث — M1 يدعم ping فقط
 * M2 وما بعدها: سيُضاف التوجيه للـ intents
 */
function handleEvent_(eventType, payload, requestId) {
  switch (eventType) {
    case 'ping':
      return handlePing_(payload);

    case 'echo':
      return handleEcho_(payload, requestId);

    default:
      return {
        warning: 'نوع الحدث "' + eventType + '" لم يُنفَّذ بعد',
        available_events: ['ping', 'echo'],
        note: 'هذا M1 — الأحداث الفعلية تُضاف في M2+'
      };
  }
}

/**
 * ping — اختبار أن المخ حي ومتصل بـ Claude
 */
function handlePing_(payload) {
  var message = (payload && payload.message) || 'مرحباً';

  // استدعاء Claude للتأكد من الاتصال (Haiku، أرخص نموذج)
  var claudeResponse = callClaude_({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 100,
    messages: [{
      role: 'user',
      content: message + '\n\nرد برسالة ترحيب قصيرة بالعربية (سطر واحد).'
    }]
  });

  return {
    event: 'ping',
    brain_says: 'نعم، المخ يعمل',
    claude_says: claudeResponse.text,
    model_used: claudeResponse.model,
    tokens: claudeResponse.usage
  };
}

/**
 * echo — يرجّع ما أرسلته (اختبار المصادقة فقط، لا يستدعي Claude)
 */
function handleEcho_(payload, requestId) {
  return {
    event: 'echo',
    received: payload,
    requestId: requestId,
    note: 'هذا اختبار بدون استدعاء Claude'
  };
}

/**
 * قراءة header من الطلب (GAS لا يوفّرها مباشرة، نستخدم parameter كبديل)
 */
function getHeader_(e, name) {
  if (e && e.parameter && e.parameter[name]) {
    return e.parameter[name];
  }
  return null;
}

/**
 * رد JSON قياسي
 */
function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * رد خطأ قياسي
 */
function errorResponse_(code, message, httpCode, requestId) {
  return jsonResponse_({
    success: false,
    error: { code: code, message: message, http: httpCode },
    requestId: requestId,
    timestamp: new Date().toISOString()
  });
}

/**
 * تسجيل حدث في Logger + Console
 */
function logEvent_(type, data) {
  var entry = {
    time: new Date().toISOString(),
    type: type,
    data: data
  };
  console.log(JSON.stringify(entry));
}
