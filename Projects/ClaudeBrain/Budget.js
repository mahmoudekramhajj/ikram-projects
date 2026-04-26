/**
 * Budget.js — تتبّع الإنفاق اليومي والساعي، وإيقاف المخ عند التجاوز
 * الإنفاق يُحفظ في Script Properties (مفتاح لكل يوم/ساعة).
 */

/**
 * فحص ما إذا كان الإنفاق ضمن الحد المسموح
 * @returns {Object} — { ok: boolean, spent: number, limit: number, scope: string }
 */
function checkBudget_() {
  var config = getConfig_();

  var daily = getSpendForKey_(getDailyKey_());
  if (daily >= config.DAILY_BUDGET_USD) {
    return { ok: false, spent: daily.toFixed(4), limit: config.DAILY_BUDGET_USD, scope: 'daily' };
  }

  var hourly = getSpendForKey_(getHourlyKey_());
  if (hourly >= config.HOURLY_BUDGET_USD) {
    return { ok: false, spent: hourly.toFixed(4), limit: config.HOURLY_BUDGET_USD, scope: 'hourly' };
  }

  return { ok: true, daily_spent: daily, hourly_spent: hourly, daily_limit: config.DAILY_BUDGET_USD };
}

/**
 * تسجيل إنفاق جديد (يُستدعى بعد كل استدعاء Claude ناجح)
 */
function recordSpend_(costUsd) {
  if (!costUsd || costUsd <= 0) return;

  var props = PropertiesService.getScriptProperties();
  var dailyKey = getDailyKey_();
  var hourlyKey = getHourlyKey_();

  var daily = getSpendForKey_(dailyKey) + costUsd;
  var hourly = getSpendForKey_(hourlyKey) + costUsd;

  props.setProperty(dailyKey, daily.toFixed(6));
  props.setProperty(hourlyKey, hourly.toFixed(6));
}

/**
 * قراءة إنفاق مسجّل لمفتاح معيّن
 */
function getSpendForKey_(key) {
  var val = PropertiesService.getScriptProperties().getProperty(key);
  if (!val) return 0;
  var num = parseFloat(val);
  return isNaN(num) ? 0 : num;
}

/**
 * مفتاح الإنفاق اليومي — SPEND_DAY_2026-04-23
 */
function getDailyKey_() {
  var d = new Date();
  var yyyy = d.getFullYear();
  var mm = pad2_(d.getMonth() + 1);
  var dd = pad2_(d.getDate());
  return 'SPEND_DAY_' + yyyy + '-' + mm + '-' + dd;
}

/**
 * مفتاح الإنفاق الساعي — SPEND_HOUR_2026-04-23-14
 */
function getHourlyKey_() {
  var d = new Date();
  var yyyy = d.getFullYear();
  var mm = pad2_(d.getMonth() + 1);
  var dd = pad2_(d.getDate());
  var hh = pad2_(d.getHours());
  return 'SPEND_HOUR_' + yyyy + '-' + mm + '-' + dd + '-' + hh;
}

function pad2_(n) {
  return (n < 10) ? '0' + n : '' + n;
}

/**
 * تقرير سريع عن الإنفاق — يمكن استدعاؤه من محرر GAS
 */
function budgetReport() {
  var config = getConfig_();
  var report = {
    today: {
      key: getDailyKey_(),
      spent: getSpendForKey_(getDailyKey_()).toFixed(4),
      limit: config.DAILY_BUDGET_USD
    },
    current_hour: {
      key: getHourlyKey_(),
      spent: getSpendForKey_(getHourlyKey_()).toFixed(4),
      limit: config.HOURLY_BUDGET_USD
    }
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

/**
 * إعادة ضبط عدّاد اليوم (للطوارئ/الاختبار) — استخدمه بحذر
 */
function resetTodaySpend() {
  PropertiesService.getScriptProperties().deleteProperty(getDailyKey_());
  PropertiesService.getScriptProperties().deleteProperty(getHourlyKey_());
  Logger.log('Today spend counters reset.');
}
