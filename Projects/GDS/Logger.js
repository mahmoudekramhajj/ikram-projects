/**
 * Logger.js — تسجيل منظّم عبر Stackdriver
 *
 * استخدام:
 *   GDS2.Log.info('pipeline started');
 *   GDS2.Log.error('parse failed', { passport: 'A123', reason: 'timeout' });
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Log = {
  info: function(msg, ctx) {
    Logger.log('[INFO] ' + msg + GDS2.Log._fmt(ctx));
  },

  warn: function(msg, ctx) {
    Logger.log('[WARN] ' + msg + GDS2.Log._fmt(ctx));
    console.warn(msg, ctx || '');
  },

  error: function(msg, ctx) {
    Logger.log('[ERROR] ' + msg + GDS2.Log._fmt(ctx));
    console.error(msg, ctx || '');
  },

  debug: function(msg, ctx) {
    Logger.log('[DEBUG] ' + msg + GDS2.Log._fmt(ctx));
  },

  _fmt: function(ctx) {
    if (!ctx) return '';
    try {
      if (typeof ctx === 'string') return ' | ' + ctx;
      return ' | ' + JSON.stringify(ctx);
    } catch (e) {
      return ' | [context stringify failed: ' + e.message + ']';
    }
  }
};
