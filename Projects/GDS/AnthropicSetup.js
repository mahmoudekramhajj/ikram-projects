/**
 * AnthropicSetup.js — إعداد واختبار مفتاح Anthropic API
 *
 * دوال عامة:
 *   setAnthropicKey(key)   — حفظ المفتاح في Script Properties
 *   testAnthropicKey()     — اختبار المفتاح بمكالمة ping صغيرة
 *   checkAnthropicKey()    — التحقق من وجود المفتاح (بدون إرسال لـ Anthropic)
 *
 * ⚠️ المفتاح لا يُسجَّل أبداً في Logger أو أي مكان مرئي — فقط في Script Properties.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.AnthropicSetup = {
  /**
   * حفظ المفتاح في Script Properties.
   * يُرجع prefix فقط للتحقق (ليس المفتاح كاملاً).
   */
  setKey: function(key) {
    if (!key || typeof key !== 'string') {
      return { status: 'error', reason: 'invalid key' };
    }
    var trimmed = key.trim();
    if (!trimmed.startsWith('sk-ant-')) {
      return { status: 'error', reason: 'key does not start with sk-ant-' };
    }
    if (trimmed.length < 40) {
      return { status: 'error', reason: 'key too short' };
    }

    GDS2.State.set(GDS2.Config.PROP.ANTHROPIC_API_KEY, trimmed);

    return {
      status: 'ok',
      saved: true,
      key_length: trimmed.length,
      key_prefix: trimmed.substring(0, 14) + '...',
      key_suffix: '...' + trimmed.substring(trimmed.length - 4)
    };
  },

  /**
   * التحقق من وجود المفتاح بدون إرسال أي طلب.
   */
  checkKey: function() {
    var key = GDS2.State.get(GDS2.Config.PROP.ANTHROPIC_API_KEY);
    if (!key) {
      return { exists: false };
    }
    return {
      exists: true,
      key_length: key.length,
      key_prefix: key.substring(0, 14) + '...',
      key_suffix: '...' + key.substring(key.length - 4)
    };
  },

  /**
   * اختبار المفتاح بمكالمة صغيرة لـ Anthropic.
   */
  test: function() {
    var key = GDS2.State.get(GDS2.Config.PROP.ANTHROPIC_API_KEY);
    if (!key) {
      return { status: 'error', reason: 'key not set — run setAnthropicKey first' };
    }

    var startTime = new Date();

    try {
      var response = UrlFetchApp.fetch(GDS2.Config.CLAUDE_API_URL, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'x-api-key': key,
          'anthropic-version': GDS2.Config.CLAUDE_ANTHROPIC_VERSION
        },
        payload: JSON.stringify({
          model: GDS2.Config.CLAUDE_MODEL,
          max_tokens: 30,
          messages: [
            { role: 'user', content: 'Reply with exactly this word and nothing else: pong' }
          ]
        }),
        muteHttpExceptions: true
      });

      var code = response.getResponseCode();
      var body = response.getContentText();
      var elapsed = (new Date() - startTime) / 1000;

      if (code === 200) {
        var json = JSON.parse(body);
        var reply = (json.content && json.content[0] && json.content[0].text) || '';
        return {
          status: 'ok',
          http_code: code,
          model: json.model,
          claude_reply: reply,
          usage: json.usage,
          elapsed_sec: elapsed
        };
      }

      // خطأ HTTP
      var parsed = null;
      try { parsed = JSON.parse(body); } catch (e) {}
      return {
        status: 'error',
        http_code: code,
        error_body: parsed || body.substring(0, 500),
        elapsed_sec: elapsed
      };
    } catch (e) {
      return {
        status: 'error',
        exception: e.message,
        elapsed_sec: (new Date() - startTime) / 1000
      };
    }
  }
};

// ==================== Global Entry Points ====================

function setAnthropicKey(key) {
  return GDS2.AnthropicSetup.setKey(key);
}

function checkAnthropicKey() {
  return GDS2.AnthropicSetup.checkKey();
}

function testAnthropicKey() {
  return GDS2.AnthropicSetup.test();
}
