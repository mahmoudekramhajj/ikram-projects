/**
 * Claude.js — استدعاء Anthropic Messages API مع retry على 529
 */

/**
 * استدعاء Claude API
 * @param {Object} request — { model, max_tokens, messages, system?, tools? }
 * @returns {Object} — { text, model, usage, cost, stop_reason }
 */
function callClaude_(request) {
  var config = getConfig_();

  var model = request.model || config.DEFAULT_MODEL;
  var payload = {
    model: model,
    max_tokens: request.max_tokens || 1024,
    messages: request.messages
  };

  if (request.system) payload.system = request.system;
  if (request.tools) payload.tools = request.tools;
  if (request.tool_choice) payload.tool_choice = request.tool_choice;

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': config.CLAUDE_API_KEY,
      'anthropic-version': config.ANTHROPIC_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // Retry loop مع exponential backoff للـ 529 / 503 / 429
  var lastErr = null;
  var maxAttempts = config.RETRY_MAX;

  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var response = UrlFetchApp.fetch(config.ANTHROPIC_URL, options);
      var code = response.getResponseCode();
      var body = response.getContentText();

      if (code >= 200 && code < 300) {
        var parsed = JSON.parse(body);
        return processSuccessResponse_(parsed, model);
      }

      // 529 Overloaded / 503 Service Unavailable / 429 Rate Limit → retry
      if (code === 529 || code === 503 || code === 429) {
        lastErr = new Error('API ' + code + ': ' + body);
        if (attempt < maxAttempts) {
          var waitMs = config.RETRY_BASE_MS * Math.pow(3, attempt - 1);
          console.log('Claude API ' + code + ' — retry ' + attempt + '/' + maxAttempts + ' after ' + waitMs + 'ms');
          Utilities.sleep(waitMs);
          continue;
        }
      }

      // أخطاء أخرى (4xx/5xx) → لا retry
      throw new Error('Claude API error ' + code + ': ' + body);

    } catch (err) {
      lastErr = err;
      if (attempt >= maxAttempts) throw err;
      Utilities.sleep(config.RETRY_BASE_MS * attempt);
    }
  }

  throw lastErr || new Error('Claude call failed after ' + maxAttempts + ' attempts');
}

/**
 * معالجة رد Claude الناجح
 */
function processSuccessResponse_(parsed, model) {
  // استخراج النص من content blocks
  var text = '';
  var toolUses = [];

  if (parsed.content && Array.isArray(parsed.content)) {
    for (var i = 0; i < parsed.content.length; i++) {
      var block = parsed.content[i];
      if (block.type === 'text') {
        text += block.text;
      } else if (block.type === 'tool_use') {
        toolUses.push({
          id: block.id,
          name: block.name,
          input: block.input
        });
      }
    }
  }

  // حساب التكلفة
  var usage = parsed.usage || {};
  var inputTokens = usage.input_tokens || 0;
  var outputTokens = usage.output_tokens || 0;
  var cost = calculateCost_(model, inputTokens, outputTokens);

  // تسجيل التكلفة في Budget
  recordSpend_(cost);

  return {
    text: text,
    tool_uses: toolUses,
    model: model,
    usage: {
      input_tokens: inputTokens,
      output_tokens: outputTokens
    },
    cost: cost,
    stop_reason: parsed.stop_reason,
    raw: parsed
  };
}
