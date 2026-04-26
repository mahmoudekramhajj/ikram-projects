/**
 * Config.js — قراءة الإعدادات من Script Properties
 * لا تكتب أي قيمة سرية في هذا الملف.
 */

/**
 * قراءة قيمة من Script Properties
 */
function getProp_(key, defaultValue) {
  var props = PropertiesService.getScriptProperties();
  var val = props.getProperty(key);
  if (val === null || val === undefined) {
    if (defaultValue !== undefined) return defaultValue;
    throw new Error('Missing Script Property: ' + key);
  }
  return val;
}

/**
 * الإعدادات المطلوبة للمخ
 */
function getConfig_() {
  return {
    // مفاتيح حساسة
    CLAUDE_API_KEY: getProp_('CLAUDE_API_KEY'),
    HMAC_SECRET: getProp_('HMAC_SECRET'),

    // إعدادات النماذج
    DEFAULT_MODEL: getProp_('DEFAULT_MODEL', 'claude-haiku-4-5-20251001'),
    ANTHROPIC_VERSION: getProp_('ANTHROPIC_VERSION', '2023-06-01'),
    ANTHROPIC_URL: getProp_('ANTHROPIC_URL', 'https://api.anthropic.com/v1/messages'),

    // الميزانية
    DAILY_BUDGET_USD: parseFloat(getProp_('DAILY_BUDGET_USD', '2.0')),
    HOURLY_BUDGET_USD: parseFloat(getProp_('HOURLY_BUDGET_USD', '0.5')),

    // Retry
    RETRY_MAX: parseInt(getProp_('RETRY_MAX', '3'), 10),
    RETRY_BASE_MS: parseInt(getProp_('RETRY_BASE_MS', '1000'), 10),

    // المفتاح الموحّد لـ ClaudeAPI.js (التحكم عن بعد)
    CLAUDE_API_REMOTE_KEY: getProp_('CLAUDE_API_REMOTE_KEY', 'ekram2026claude')
  };
}

/**
 * أسعار النماذج (USD لكل مليون tokens) — للحساب
 * مصدر: Anthropic pricing
 */
var MODEL_PRICING = {
  'claude-haiku-4-5-20251001':   { input: 1.00,  output: 5.00  },
  'claude-sonnet-4-6':           { input: 3.00,  output: 15.00 },
  'claude-opus-4-6':             { input: 15.00, output: 75.00 }
};

/**
 * حساب تكلفة استدعاء (بالدولار)
 */
function calculateCost_(model, inputTokens, outputTokens) {
  var pricing = MODEL_PRICING[model];
  if (!pricing) return 0;
  var cost = (inputTokens / 1000000) * pricing.input
           + (outputTokens / 1000000) * pricing.output;
  return Math.round(cost * 10000) / 10000;  // 4 منازل عشرية
}
