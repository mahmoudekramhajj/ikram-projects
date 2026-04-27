/**
 * Diagnostics.js — دوال تشخيصية + عرض التقارير في واجهة المستخدم
 *
 * تقارير تُعرض في نوافذ منبثقة (Toast/Alert) من القائمة.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Diagnostics = {
  /**
   * تقرير الحالة العامة.
   */
  status: function() {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { total: 0 };

    // نقرأ: جواز (6), رابط (26), BJ (62), BK (63), BL (64), BM (65)
    var lastCol = GDS2.Config.COL.NUSUK_AUTH; // 65
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    var stats = {
      total: 0,
      with_url: 0,
      no_url: 0,
      processed: 0,
      pending: 0,
      manual: 0,
      stopped: 0,
      nusuk_auth: 0,
      fail_once: 0
    };

    for (var i = 0; i < data.length; i++) {
      var passport = String(data[i][GDS2.Config.COL.PASSPORT - 1] || '').trim();
      if (!passport) continue;
      stats.total++;

      var url = String(data[i][GDS2.Config.COL.TICKET_URL - 1] || '').trim();
      var bj = String(data[i][GDS2.Config.COL.LAST_URL - 1] || '').trim();
      var bk = data[i][GDS2.Config.COL.MANUAL - 1];
      var bl = Number(data[i][GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      var bm = data[i][GDS2.Config.COL.NUSUK_AUTH - 1];

      if (url) stats.with_url++;
      else stats.no_url++;

      if (!url) continue; // صفوف بلا URL: لا تُصنَّف في pending/processed

      if (bk === true || String(bk).toUpperCase() === 'TRUE') {
        stats.manual++;
      } else if (bm === true || String(bm).toUpperCase() === 'TRUE') {
        stats.nusuk_auth++;
      } else if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) {
        stats.stopped++;
      } else if (bj && bj === url) {
        stats.processed++;
      } else if (bl === 1) {
        stats.fail_once++;
        stats.pending++;
      } else {
        stats.pending++;
      }
    }

    return stats;
  },

  /**
   * قائمة بالحجاج الذين توقفت معالجتهم (BL >= max).
   */
  failures: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var bl = Number(row[GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      return bl >= GDS2.Config.MAX_FAIL_ATTEMPTS;
    });
  },

  manual: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var bk = row[GDS2.Config.COL.MANUAL - 1];
      return bk === true || String(bk).toUpperCase() === 'TRUE';
    });
  },

  nusuk: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var bm = row[GDS2.Config.COL.NUSUK_AUTH - 1];
      return bm === true || String(bm).toUpperCase() === 'TRUE';
    });
  },

  pending: function() {
    return GDS2.Diagnostics._listByFilter(function(row) {
      var url = String(row[GDS2.Config.COL.TICKET_URL - 1] || '').trim();
      var bj = String(row[GDS2.Config.COL.LAST_URL - 1] || '').trim();
      var bk = row[GDS2.Config.COL.MANUAL - 1];
      var bl = Number(row[GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      var bm = row[GDS2.Config.COL.NUSUK_AUTH - 1];

      if (!url) return false;
      if (bk === true || String(bk).toUpperCase() === 'TRUE') return false;
      if (bm === true || String(bm).toUpperCase() === 'TRUE') return false;
      if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) return false;
      if (bj === url) return false;
      return true;
    });
  },

  // ==================== UI helpers ====================

  alertStatus: function() {
    var s = GDS2.Diagnostics.status();
    var msg = [
      '📊 إحصائيات B2C',
      '',
      'إجمالي الصفوف: ' + s.total,
      'بـ ticketUrl: ' + s.with_url,
      '',
      '✅ معالجة ناجحة: ' + s.processed,
      '⏳ معلّقة: ' + s.pending,
      '🟡 فشل مرة واحدة: ' + s.fail_once,
      '🔴 متوقفة نهائياً: ' + s.stopped,
      '✍️ مُعدَّلة يدوياً: ' + s.manual,
      '🔒 تحتاج دخول نسك: ' + s.nusuk_auth
    ].join('\n');
    GDS2.Diagnostics._alert(msg, 'تقرير الحالة');
  },

  alertLastRun: function() {
    var status = GDS2.Pipeline.getStatus();
    if (!status.last_run) {
      GDS2.Diagnostics._alert('لم يُشغَّل Pipeline بعد.', 'آخر تشغيل');
      return;
    }
    var r = status.last_run;
    var msg = [
      '📋 آخر تشغيل',
      '',
      'بدأ: ' + r.started_at,
      'انتهى: ' + r.finished_at,
      'المدة: ' + (r.duration_sec || 0).toFixed(1) + ' ثانية',
      '',
      'صفوف مُفحوصة: ' + r.rows_scanned,
      '✅ نجحت: ' + r.success_count,
      '👨‍👩‍👧 عائلة: ' + r.family_applied,
      '❌ فشلت: ' + r.failure_count,
      '🚨 كوارث: ' + r.disaster_count,
      '🔒 nusuk auth: ' + r.nusuk_auth_count,
      '⏭️ متخطاة: ' + r.skipped_already,
      (r.timed_out ? '⏰ توقف عند الحد الزمني' : '')
    ].filter(Boolean).join('\n');
    GDS2.Diagnostics._alert(msg, 'آخر تشغيل');
  },

  alertList: function(list, title) {
    if (!list || list.length === 0) {
      GDS2.Diagnostics._alert('لا توجد صفوف مطابقة.', title);
      return;
    }
    var lines = [title + ' (' + list.length + ' حاج)', ''];
    var max = Math.min(list.length, 30);
    for (var i = 0; i < max; i++) {
      var item = list[i];
      lines.push('صف ' + item.row + ': ' + item.name + ' (' + item.passport + ')');
    }
    if (list.length > max) {
      lines.push('');
      lines.push('... و ' + (list.length - max) + ' آخرين');
    }
    GDS2.Diagnostics._alert(lines.join('\n'), title);
  },

  alertIATA: function() {
    var list = GDS2.IATARegistry.getList();
    var msg = [
      'قائمة IATA (' + list.length + ' كود)',
      '',
      list.join(', ')
    ].join('\n');
    GDS2.Diagnostics._alert(msg, 'IATA');
  },

  alertAirlines: function() {
    var list = GDS2.AirlineRegistry.getList();
    var lines = ['شركات الطيران (' + list.length + ')', ''];
    for (var i = 0; i < list.length; i++) {
      lines.push(list[i].code + ' — ' + list[i].name);
    }
    GDS2.Diagnostics._alert(lines.join('\n'), 'Airlines');
  },

  alertTelegramConfig: function() {
    var cfg = GDS2.TelegramNotifier.checkConfig();
    var msg = [
      'حالة Telegram',
      '',
      'Token: ' + (cfg.token_set ? '✅ مُعدّ' : '❌ غير مُعدّ'),
      'Chat ID: ' + (cfg.chat_id_set ? '✅ ' + cfg.chat_id : '❌ غير مُعدّ'),
      'جاهز: ' + (cfg.ready ? '✅' : '❌')
    ].join('\n');
    GDS2.Diagnostics._alert(msg, 'Telegram');
  },

  // ==================== Private ====================

  _listByFilter: function(filterFn) {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, GDS2.Config.COL.NUSUK_AUTH).getValues();
    var results = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!filterFn(row)) continue;

      var passport = String(row[GDS2.Config.COL.PASSPORT - 1] || '').trim();
      if (!passport) continue;

      var firstEn = String(row[GDS2.Config.COL.FIRST_NAME_EN - 1] || '').trim();
      var lastEn = String(row[GDS2.Config.COL.LAST_NAME_EN - 1] || '').trim();
      var firstAr = String(row[GDS2.Config.COL.FIRST_NAME_AR - 1] || '').trim();
      var lastAr = String(row[GDS2.Config.COL.LAST_NAME_AR - 1] || '').trim();
      var name = (firstEn + ' ' + lastEn).trim() || (firstAr + ' ' + lastAr).trim() || '(بلا اسم)';

      results.push({
        row: i + 2,
        passport: passport,
        name: name
      });
    }

    return results;
  },

  _alert: function(msg, title) {
    try {
      SpreadsheetApp.getUi().alert(title || 'GDS', msg, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // لا UI متاح (مثلاً عبر ClaudeAPI)
      Logger.log('[' + (title || 'GDS') + '] ' + msg);
    }
  }
};

// ==================== Global Entry Points ====================

function gdsShowStatus() { GDS2.Diagnostics.alertStatus(); }

function gdsShowLastRun() { GDS2.Diagnostics.alertLastRun(); }

function gdsShowFailures() {
  var list = GDS2.Diagnostics.failures();
  GDS2.Diagnostics.alertList(list, '🔴 حجاج متوقفون نهائياً (BL=2)');
}

function gdsShowManual() {
  var list = GDS2.Diagnostics.manual();
  GDS2.Diagnostics.alertList(list, '✍️ حجاج مُعدَّلون يدوياً');
}

function gdsShowNusuk() {
  var list = GDS2.Diagnostics.nusuk();
  GDS2.Diagnostics.alertList(list, '🔒 حجاج يحتاجون دخول نسك');
}

function gdsShowPending() {
  var list = GDS2.Diagnostics.pending();
  GDS2.Diagnostics.alertList(list, '⏳ حجاج معلّقون (بدون معالجة)');
}

/**
 * فحص صف B2C كاملاً بجواز معيّن — للتشخيص.
 * يُرجع كل الأعمدة المهمة بمسمياتها العربية.
 */
function inspectRow(passport) {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRow = ss.getLastRow();
  if (lastRow < 2) return { error: 'sheet_empty' };
  var target = String(passport).trim();
  var data = ss.getRange(2, 1, lastRow - 1, GDS2.Config.COL.RET3_ARR_TIME).getValues();
  var C = GDS2.Config.COL;
  for (var i = 0; i < data.length; i++) {
    var p = String(data[i][C.PASSPORT - 1] || '').trim();
    if (p !== target) continue;
    var r = data[i];
    return {
      row: i + 2,
      passport: p,
      name_en: (String(r[C.FIRST_NAME_EN-1]||'') + ' ' + String(r[C.LAST_NAME_EN-1]||'')).trim(),
      name_ar: (String(r[C.FIRST_NAME_AR-1]||'') + ' ' + String(r[C.LAST_NAME_AR-1]||'')).trim(),
      package_no: r[C.PACKAGE_NO-1],
      package_name: r[C.PACKAGE_NAME-1],
      contract_type: r[C.CONTRACT_TYPE-1],
      ticket_no: r[C.TICKET_NO-1],
      ticket_url: r[C.TICKET_URL-1],
      pnr: r[C.PNR-1],
      bj_last_url: r[C.LAST_URL-1],
      bk_manual: r[C.MANUAL-1],
      bl_fail_count: r[C.FAIL_COUNT-1],
      bm_nusuk: r[C.NUSUK_AUTH-1],
      dep1: { fn: r[33-1], depDate: r[34-1], depTime: r[35-1], from: r[36-1], to: r[37-1], arrDate: r[38-1], arrTime: r[39-1] },
      dep2: { fn: r[40-1], depDate: r[41-1], depTime: r[42-1], from: r[43-1], to: r[44-1], arrDate: r[45-1], arrTime: r[46-1] },
      ret1: { fn: r[47-1], depDate: r[48-1], depTime: r[49-1], from: r[50-1], to: r[51-1], arrDate: r[52-1], arrTime: r[53-1] },
      ret2: { fn: r[54-1], depDate: r[55-1], depTime: r[56-1], from: r[57-1], to: r[58-1], arrDate: r[59-1], arrTime: r[60-1] },
      dep0: { fn: r[71-1], depDate: r[72-1], depTime: r[73-1], from: r[74-1], to: r[75-1], arrDate: r[76-1], arrTime: r[77-1] },
      ret3: { fn: r[78-1], depDate: r[79-1], depTime: r[80-1], from: r[81-1], to: r[82-1], arrDate: r[83-1], arrTime: r[84-1] }
    };
  }
  return { error: 'passport_not_found', searched: target };
}

/**
 * فحص Presonal Details — البحث بالجواز أو بجزء من الاسم.
 * يُرجع كل التطابقات مع الأعمدة المهمة.
 */
function inspectPD(query) {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_PD);
  if (!ss) return { error: 'PD_sheet_not_found' };
  var lastRow = ss.getLastRow();
  var lastCol = ss.getLastColumn();
  if (lastRow < 2) return { error: 'PD_empty' };
  var headers = ss.getRange(1, 1, 1, lastCol).getValues()[0];
  var data = ss.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var q = String(query).trim().toLowerCase();
  var matches = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var hit = false;
    for (var c = 0; c < row.length; c++) {
      var v = String(row[c] || '').toLowerCase();
      if (v === q || (v.indexOf(q) !== -1 && q.length >= 4)) { hit = true; break; }
    }
    if (!hit) continue;
    var obj = { row: i + 2 };
    for (var h = 0; h < headers.length; h++) {
      var key = String(headers[h] || ('col_' + (h+1))).trim();
      if (key) obj[key] = row[h];
    }
    matches.push(obj);
  }
  return { query: query, total_matches: matches.length, matches: matches.slice(0, 10), pd_headers: headers };
}

/**
 * إحصاء الصفوف المتوقفة (BL=MAX) مع تصنيف نمط المشكلة لكل صف.
 * الأنماط:
 *  - missing_dep2: dep2.from أو dep2.to فارغ (الإصلاح الأخير قد يحلها)
 *  - bad_url: ticket_url ليس http
 *  - has_data: عنده dep2 و url صحيح (سبب آخر)
 *  - empty: bj فارغ
 */
function auditStopped() {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRow = ss.getLastRow();
  if (lastRow < 2) return { error: 'sheet_empty' };
  var C = GDS2.Config.COL;
  var data = ss.getRange(2, 1, lastRow - 1, C.RET3_ARR_TIME).getValues();
  var stats = { total_stopped: 0, missing_dep2: 0, bad_url: 0, has_data: 0, empty: 0, by_pattern: {} };
  var samples = { missing_dep2: [], bad_url: [], has_data: [] };

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var bl = Number(r[C.FAIL_COUNT-1]) || 0;
    if (bl < GDS2.Config.MAX_FAIL_ATTEMPTS) continue;
    stats.total_stopped++;
    var passport = String(r[C.PASSPORT-1]||'').trim();
    var name = (String(r[C.FIRST_NAME_EN-1]||'') + ' ' + String(r[C.LAST_NAME_EN-1]||'')).trim();
    var url = String(r[C.TICKET_URL-1]||'').trim();
    var dep2_from = String(r[43-1]||'').trim();
    var dep2_to = String(r[44-1]||'').trim();
    var pattern = '';
    if (!url || url.indexOf('http') !== 0) pattern = 'bad_url';
    else if (!dep2_from || !dep2_to) pattern = 'missing_dep2';
    else pattern = 'has_data';
    stats[pattern]++;
    if (samples[pattern].length < 5) {
      samples[pattern].push({ row: i+2, passport: passport, name: name, url_prefix: url.substring(0,60), dep2: dep2_from + '-' + dep2_to });
    }
  }
  stats.samples = samples;
  return stats;
}

/**
 * مزامنة ticket_url من PD إلى B2C لجواز معيّن.
 * يستخدم لإصلاح صفوف B2C التي ticket_url فيها خاطئ (اسم ملف بدلاً من رابط).
 * يمسح BJ تلقائياً ليُعاد المعالجة في الدورة القادمة.
 */
function syncTicketUrlFromPD(passport) {
  var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
  var pd = ss.getSheetByName(GDS2.Config.SHEET_PD);
  if (!pd) return { error: 'pd_not_found' };
  var pdLastRow = pd.getLastRow();
  var pdLastCol = pd.getLastColumn();
  var headers = pd.getRange(1, 1, 1, pdLastCol).getValues()[0];
  var passportColIdx = -1, urlColIdx = -1, ticketNoColIdx = -1;
  for (var h = 0; h < headers.length; h++) {
    var hh = String(headers[h] || '').trim();
    if (hh === 'رقم جواز السفر') passportColIdx = h;
    else if (hh === 'رابط التذكرة') urlColIdx = h;
    else if (hh === 'رقم التذكرة') ticketNoColIdx = h;
  }
  if (passportColIdx < 0 || urlColIdx < 0) return { error: 'pd_columns_not_found', headers: headers };

  var pdData = pd.getRange(2, 1, pdLastRow - 1, pdLastCol).getValues();
  var pdUrl = '', pdTicketNo = '', pdRow = -1;
  for (var i = 0; i < pdData.length; i++) {
    if (String(pdData[i][passportColIdx] || '').trim() === String(passport).trim()) {
      pdUrl = String(pdData[i][urlColIdx] || '').trim();
      pdTicketNo = ticketNoColIdx >= 0 ? String(pdData[i][ticketNoColIdx] || '').trim() : '';
      pdRow = i + 2;
      break;
    }
  }
  if (pdRow < 0) return { error: 'passport_not_found_in_pd', passport: passport };
  if (!pdUrl) return { error: 'pd_url_empty', pd_row: pdRow };

  var b2c = ss.getSheetByName(GDS2.Config.SHEET_B2C);
  var b2cLastRow = b2c.getLastRow();
  var passports = b2c.getRange(2, GDS2.Config.COL.PASSPORT, b2cLastRow - 1, 1).getValues();
  for (var j = 0; j < passports.length; j++) {
    if (String(passports[j][0] || '').trim() === String(passport).trim()) {
      var row = j + 2;
      var oldUrl = String(b2c.getRange(row, GDS2.Config.COL.TICKET_URL).getValue() || '').trim();
      var oldTicketNo = String(b2c.getRange(row, GDS2.Config.COL.TICKET_NO).getValue() || '').trim();
      b2c.getRange(row, GDS2.Config.COL.TICKET_URL).setValue(pdUrl);
      var changes = { url: { old: oldUrl, new: pdUrl } };
      if (pdTicketNo && oldTicketNo !== pdTicketNo) {
        b2c.getRange(row, GDS2.Config.COL.TICKET_NO).setValue(pdTicketNo);
        changes.ticket_no = { old: oldTicketNo, new: pdTicketNo };
      }
      b2c.getRange(row, GDS2.Config.COL.LAST_URL).setValue('');
      b2c.getRange(row, GDS2.Config.COL.FAIL_COUNT).setValue(0);
      return { status: 'ok', b2c_row: row, pd_row: pdRow, passport: passport, changes: changes, bj_cleared: true, bl_reset: true };
    }
  }
  return { error: 'passport_not_found_in_b2c', passport: passport };
}

function gdsListTriggers() {
  var t = GDS2.Triggers.list();
  var lines = ['Triggers (' + t.count + ')', ''];
  for (var i = 0; i < t.triggers.length; i++) {
    lines.push(t.triggers[i].handler + ' — ' + t.triggers[i].event);
  }
  GDS2.Diagnostics._alert(lines.join('\n'), 'Triggers');
}

function gdsShowIATA() { GDS2.Diagnostics.alertIATA(); }

function gdsShowAirlines() { GDS2.Diagnostics.alertAirlines(); }

function gdsCheckTelegram() { GDS2.Diagnostics.alertTelegramConfig(); }

// للاستدعاء عبر ClaudeAPI (يُرجع JSON)
function getStatusJSON() { return GDS2.Diagnostics.status(); }
function getPendingJSON() { return GDS2.Diagnostics.pending(); }
function getFailuresJSON() { return GDS2.Diagnostics.failures(); }
