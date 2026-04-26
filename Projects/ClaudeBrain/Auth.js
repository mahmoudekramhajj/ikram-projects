/**
 * Auth.js — التحقق من توقيع HMAC-SHA256
 * كل طلب POST يجب أن يحمل X-Signature = HMAC(body, HMAC_SECRET)
 */

/**
 * التحقق من صحة التوقيع
 * @param {string} rawBody — جسم الطلب الخام
 * @param {string} signature — التوقيع القادم مع الطلب
 * @returns {Object} — { valid: boolean, reason?: string }
 */
function verifyHMAC_(rawBody, signature) {
  if (!signature) {
    return { valid: false, reason: 'Missing X-Signature' };
  }

  var config = getConfig_();
  var expected = computeHMAC_(rawBody, config.HMAC_SECRET);

  // مقارنة ثابتة الزمن (constant-time) لمنع timing attacks
  if (!constantTimeEquals_(signature.toLowerCase(), expected.toLowerCase())) {
    return { valid: false, reason: 'Invalid signature' };
  }

  return { valid: true };
}

/**
 * حساب HMAC-SHA256 بصيغة hex
 */
function computeHMAC_(message, secret) {
  var raw = Utilities.computeHmacSha256Signature(message, secret);
  var hex = '';
  for (var i = 0; i < raw.length; i++) {
    var byte = raw[i] & 0xff;
    if (byte < 16) hex += '0';
    hex += byte.toString(16);
  }
  return hex;
}

/**
 * مقارنة نصّين في وقت ثابت
 */
function constantTimeEquals_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  var result = 0;
  for (var i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return result === 0;
}

/**
 * دالة مساعدة — توليد توقيع لاختبار يدوي
 * (تُستدعى من محرر GAS للحصول على توقيع اختباري)
 */
function generateTestSignature() {
  var config = getConfig_();
  var testBody = JSON.stringify({ event_type: 'ping', message: 'اختبار' });
  var sig = computeHMAC_(testBody, config.HMAC_SECRET);

  Logger.log('Body: ' + testBody);
  Logger.log('Signature: ' + sig);
  Logger.log('');
  Logger.log('curl command:');
  Logger.log('curl -X POST "[WEB_APP_URL]?X-Signature=' + sig + '" \\');
  Logger.log('  -H "Content-Type: application/json" \\');
  Logger.log("  -d '" + testBody + "'");

  return { body: testBody, signature: sig };
}
