/**
 * AnthropicClient.js — عميل Claude API كامل
 *
 * - Rate limit داخلي (sleep 1.5s بين الاستدعاءات)
 * - Retry مع exponential backoff للأخطاء 5xx
 * - معالجة خاصة لـ 429 (rate limit من Anthropic)
 * - Parsing آمن للـ JSON المُرجَع (يتعامل مع markdown code blocks)
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.AnthropicClient = {
  _lastCallTime: 0,

  /**
   * استدعاء Claude API مع نص فقط.
   * @param {Object} prompts - { system, user }
   * @param {Object} options - { maxTokens, temperature }
   */
  call: function(prompts, options) {
    return GDS2.AnthropicClient._sendRequest(prompts, options, null);
  },

  /**
   * استدعاء Claude API مع ملف PDF (document input).
   * @param {Object} prompts - { system, user }
   * @param {Array} pdfBytes - مصفوفة بايتس PDF
   * @param {Object} options - { maxTokens, temperature }
   */
  callWithPdf: function(prompts, pdfBytes, options) {
    if (!pdfBytes || !pdfBytes.length) {
      return { status: 'error', reason: 'empty_pdf' };
    }
    var base64 = Utilities.base64Encode(pdfBytes);
    return GDS2.AnthropicClient._sendRequest(prompts, options, base64);
  },

  _sendRequest: function(prompts, options, pdfBase64) {
    options = options || {};
    var maxTokens = options.maxTokens || GDS2.Config.CLAUDE_MAX_TOKENS;

    var key = GDS2.State.get(GDS2.Config.PROP.ANTHROPIC_API_KEY);
    if (!key) {
      return { status: 'error', reason: 'key_not_set' };
    }

    GDS2.AnthropicClient._enforceRateLimit();

    var userContent;
    if (pdfBase64) {
      userContent = [
        {
          type: 'document',
          source: {
            type: 'base64',
            media_type: 'application/pdf',
            data: pdfBase64
          }
        },
        { type: 'text', text: prompts.user }
      ];
    } else {
      userContent = prompts.user;
    }

    var payload = {
      model: GDS2.Config.CLAUDE_MODEL,
      max_tokens: maxTokens,
      system: prompts.system,
      messages: [
        { role: 'user', content: userContent }
      ]
    };

    if (options.temperature !== undefined) {
      payload.temperature = options.temperature;
    }

    var maxAttempts = GDS2.Config.CLAUDE_MAX_RETRIES;
    var lastError = null;

    for (var attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        var response = UrlFetchApp.fetch(GDS2.Config.CLAUDE_API_URL, {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'x-api-key': key,
            'anthropic-version': GDS2.Config.CLAUDE_ANTHROPIC_VERSION
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        });

        GDS2.AnthropicClient._lastCallTime = new Date().getTime();

        var code = response.getResponseCode();
        var body = response.getContentText();

        if (code === 200) {
          var json;
          try {
            json = JSON.parse(body);
          } catch (parseErr) {
            return {
              status: 'error',
              reason: 'claude_response_not_json',
              body_sample: body.substring(0, 300),
              attempt: attempt
            };
          }

          var text = '';
          if (json.content && json.content.length > 0 && json.content[0].text) {
            text = json.content[0].text;
          }

          return {
            status: 'ok',
            text: text,
            usage: json.usage,
            model: json.model,
            stop_reason: json.stop_reason,
            attempt: attempt
          };
        }

        // 429: rate limit — ننتظر طويلاً
        if (code === 429) {
          lastError = { code: code, body: body.substring(0, 300) };
          GDS2.Log.warn('Anthropic 429 rate limit', { attempt: attempt });
          if (attempt < maxAttempts) {
            Utilities.sleep(30000);
            continue;
          }
        }

        // 5xx: خطأ خادم — exponential backoff
        if (code >= 500) {
          lastError = { code: code, body: body.substring(0, 300) };
          GDS2.Log.warn('Anthropic 5xx', { code: code, attempt: attempt });
          if (attempt < maxAttempts) {
            Utilities.sleep(Math.pow(2, attempt) * 1000);
            continue;
          }
        }

        // أخطاء أخرى (400, 401, 403, 404): لا retry
        return {
          status: 'error',
          reason: 'http_' + code,
          http_code: code,
          body_sample: body.substring(0, 500),
          attempt: attempt
        };
      } catch (e) {
        lastError = { exception: e.message };
        GDS2.Log.warn('Anthropic exception', { attempt: attempt, error: e.message });
        if (attempt < maxAttempts) {
          Utilities.sleep(Math.pow(2, attempt) * 1000);
        }
      }
    }

    return {
      status: 'error',
      reason: 'max_attempts_reached',
      last_error: lastError,
      attempts: maxAttempts
    };
  },

  /**
   * Parse JSON response من Claude، يتعامل مع markdown code blocks.
   */
  parseJSON: function(text) {
    if (!text) return { ok: false, error: 'empty_text' };

    var cleaned = String(text).trim();

    // إزالة markdown code blocks إن وُجدت
    if (cleaned.indexOf('```') !== -1) {
      cleaned = cleaned.replace(/^```(?:json)?\s*/m, '').replace(/\s*```\s*$/m, '');
    }

    try {
      return { ok: true, data: JSON.parse(cleaned) };
    } catch (e) {
      // محاولة البحث عن كائن JSON في النص
      var match = cleaned.match(/\{[\s\S]*\}/);
      if (match) {
        try {
          return { ok: true, data: JSON.parse(match[0]) };
        } catch (e2) {}
      }
      return {
        ok: false,
        error: 'invalid_json',
        parse_error: e.message,
        sample: cleaned.substring(0, 300)
      };
    }
  },

  // ==================== Private ====================

  _enforceRateLimit: function() {
    var now = new Date().getTime();
    var elapsed = now - GDS2.AnthropicClient._lastCallTime;
    var required = GDS2.Config.CLAUDE_RATE_LIMIT_MS;
    if (GDS2.AnthropicClient._lastCallTime > 0 && elapsed < required) {
      Utilities.sleep(required - elapsed);
    }
  }
};
