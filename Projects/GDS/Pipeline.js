/**
 * Pipeline.js — المنسّق الرئيسي لمعالجة تذاكر الحجاج
 *
 * يستخدم:
 *   - LockService لمنع التشغيل المتزامن
 *   - حد زمني 5.5 دقيقة (آمن من حد GAS 6 دقائق)
 *   - skip logic لتجنب إعادة معالجة ما تم
 *   - Family optimization تلقائي (عبر FlightWriter → FamilyProcessor)
 *   - إشعارات تيليغرام للفشل والكوارث
 *
 * دوال عامة:
 *   runPipeline()       — التشغيل الافتراضي (كل الصفوف)
 *   runCanary()         — تشغيل محدود بعدد CANARY_SIZE (10 افتراضياً)
 *   runBatch(limit)     — تشغيل محدود بعدد محدد
 *   getPipelineStatus() — آخر تشغيل + إحصائيات
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Pipeline = {
  LOCK_TIMEOUT_MS: 30000,

  /**
   * التشغيل الرئيسي.
   * @param {Object} options - { limit, notifyTelegram }
   * @return {Object} stats
   */
  run: function(options) {
    options = options || {};
    var limit = options.limit || 99999;
    var notifyTelegram = options.notifyTelegram !== false; // default true

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(GDS2.Pipeline.LOCK_TIMEOUT_MS)) {
      GDS2.Log.warn('Pipeline: could not acquire lock — another run in progress');
      return { status: 'locked', reason: 'another_run_in_progress' };
    }

    var stats = {
      started_at: new Date().toISOString(),
      started_ms: new Date().getTime(),
      rows_scanned: 0,
      success_count: 0,
      family_applied: 0,
      failure_count: 0,
      disaster_count: 0,
      nusuk_auth_count: 0,
      skipped_already: 0,
      skipped_manual: 0,
      skipped_max_fails: 0,
      skipped_nusuk: 0,
      skipped_no_url: 0,
      total_tokens: { input: 0, output: 0 },
      timed_out: false,
      first_row: null,
      last_row: null
    };

    try {
      GDS2.Pipeline._runInternal(stats, limit, notifyTelegram);
    } finally {
      lock.releaseLock();
    }

    var elapsed = (new Date().getTime() - stats.started_ms) / 1000;
    stats.duration_sec = elapsed;
    stats.finished_at = new Date().toISOString();

    // حفظ آخر تشغيل
    GDS2.State.set(GDS2.Config.PROP.LAST_RUN, JSON.stringify(stats));

    GDS2.Log.info('Pipeline: finished', {
      duration: elapsed,
      success: stats.success_count,
      family: stats.family_applied,
      failure: stats.failure_count,
      disaster: stats.disaster_count,
      timed_out: stats.timed_out
    });

    return stats;
  },

  /**
   * تشغيل Canary: عدد محدود مع تقرير مفصّل.
   */
  runCanary: function() {
    return GDS2.Pipeline.run({ limit: GDS2.Config.CANARY_SIZE });
  },

  /**
   * الحصول على حالة آخر تشغيل.
   */
  getStatus: function() {
    var last = GDS2.State.get(GDS2.Config.PROP.LAST_RUN);
    if (!last) return { status: 'no_previous_run' };
    try {
      return { status: 'ok', last_run: JSON.parse(last) };
    } catch (e) {
      return { status: 'error', reason: 'corrupt_state' };
    }
  },

  // ==================== Private ====================

  _runInternal: function(stats, limit, notifyTelegram) {
    var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // قراءة البيانات اللازمة للـ skip logic لكل الصفوف
    // أعمدة: passport (6), URL (26), BJ (62), BK (63), BL (64), BM (65)
    var fullWidth = GDS2.Config.COL.NUSUK_AUTH; // 65
    var allData = sheet.getRange(2, 1, lastRow - 1, fullWidth).getValues();

    var processed = 0;

    for (var i = 0; i < allData.length; i++) {
      // فحص الميزانية الزمنية
      if (new Date().getTime() - stats.started_ms > GDS2.Config.MAX_RUNTIME_MS) {
        stats.timed_out = true;
        GDS2.Log.info('Pipeline: time budget exceeded');
        break;
      }

      if (processed >= limit) break;

      stats.rows_scanned++;
      var rowIdx = i + 2;

      var row = allData[i];
      var passport = String(row[GDS2.Config.COL.PASSPORT - 1] || '').trim();
      var url = String(row[GDS2.Config.COL.TICKET_URL - 1] || '').trim();
      var bj = String(row[GDS2.Config.COL.LAST_URL - 1] || '').trim();
      var bk = row[GDS2.Config.COL.MANUAL - 1];
      var bl = Number(row[GDS2.Config.COL.FAIL_COUNT - 1]) || 0;
      var bm = row[GDS2.Config.COL.NUSUK_AUTH - 1];

      if (!passport) continue; // صف فارغ

      // Skip logic (قبل أي معالجة)
      if (!url) {
        stats.skipped_no_url++;
        continue;
      }
      if (bk === true || String(bk).toUpperCase() === 'TRUE') {
        stats.skipped_manual++;
        continue;
      }
      if (bl >= GDS2.Config.MAX_FAIL_ATTEMPTS) {
        stats.skipped_max_fails++;
        continue;
      }
      if (bm === true || String(bm).toUpperCase() === 'TRUE') {
        stats.skipped_nusuk++;
        continue;
      }
      if (bj === url) {
        // تم معالجته بنفس الرابط → تخطَّ
        stats.skipped_already++;
        continue;
      }

      // سجّل أول/آخر صف معالَج
      if (stats.first_row === null) stats.first_row = rowIdx;
      stats.last_row = rowIdx;

      // المعالجة الفعلية
      var result = GDS2.Pipeline._processRow(passport, notifyTelegram);

      // تحديث الإحصائيات
      GDS2.Pipeline._updateStats(stats, result);

      processed++;

      // تحديث in-memory allData لعكس تأثير FamilyProcessor
      // (لو نجحت المعالجة، family قد تُغيّر BJ لصفوف أخرى)
      if (result.status === 'written' && result.family_result && result.family_result.applied_count > 0) {
        var applied = result.family_result.applied || [];
        for (var j = 0; j < applied.length; j++) {
          var appliedRow = applied[j].row;
          var idx = appliedRow - 2;
          if (idx >= 0 && idx < allData.length) {
            allData[idx][GDS2.Config.COL.LAST_URL - 1] = url; // BJ
            allData[idx][GDS2.Config.COL.FAIL_COUNT - 1] = 0; // BL
          }
        }
      }
    }
  },

  _processRow: function(passport, notifyTelegram) {
    try {
      var result = processAndWrite(passport);

      // إشعارات تيليغرام (اختيارية)
      if (notifyTelegram && result.status !== 'written') {
        GDS2.Pipeline._notifyIfNeeded(result);
      }

      return result;
    } catch (e) {
      GDS2.Log.error('Pipeline._processRow exception', { passport: passport, error: e.message });
      return { status: 'error', passport: passport, exception: e.message };
    }
  },

  _notifyIfNeeded: function(result) {
    var pilgrim = result.pilgrim || {
      name: result.name || '—',
      passport: result.passport || '—',
      url: result.url || ''
    };

    // قاعدة عامة: إشعار واحد لكل صف (بعد توقفه نهائياً)
    // السبب: بعد التوقف، BL=MAX أو BM=TRUE → Pipeline يتخطاه مستقبلاً

    // كارثة — تُرسل مرة واحدة (BL سيُضبط على MAX من processAndWrite)
    if (result.status === 'disaster') {
      GDS2.TelegramNotifier.sendDisaster(pilgrim, result.disaster || {});
      return;
    }

    // Nusuk auth — تُرسل مرة واحدة (BM=TRUE من processAndWrite)
    if (result.nusuk_marked) {
      GDS2.TelegramNotifier.sendNusukAuth(pilgrim);
      return;
    }

    // باقي الفشل: فقط عند التوقف النهائي (BL وصل للحد)
    if (!result.stopped) {
      return;
    }

    var reasonAr = GDS2.Pipeline._describeReasonArabic(result);
    var details = GDS2.Pipeline._describeDetailsArabic(result);
    GDS2.TelegramNotifier.sendFailure(pilgrim, reasonAr, details);
  },

  _describeReasonArabic: function(result) {
    if (result.status === 'validation_failed') {
      var errs = result.validation && result.validation.errors;
      var firstReason = (errs && errs.length > 0) ? errs[0].reason : '';
      var map = {
        'pilgrim_not_in_ticket': 'الحاج غير مذكور في التذكرة',
        'too_many_segments': 'عدد قطع الرحلة أكثر من المسموح',
        'missing_pnr': 'رقم PNR مفقود',
        'missing_dep2': 'رحلة الذهاب الأساسية مفقودة',
        'missing_ret1': 'رحلة العودة الأساسية مفقودة',
        'missing_segments': 'بيانات الرحلات مفقودة',
        'dep2_not_to_saudi': 'الوصول النهائي ليس لجدة أو المدينة',
        'ret1_not_from_saudi': 'العودة لا تبدأ من جدة أو المدينة',
        'return_before_departure': 'تاريخ العودة قبل الذهاب',
        'arrival_before_takeoff': 'الهبوط قبل الإقلاع',
        'out_of_season': 'تاريخ خارج موسم الحج',
        'bad_format': 'صيغة بيانات غير صحيحة'
      };
      return 'توقفت المعالجة نهائياً — فشل التحقق: ' + (map[firstReason] || firstReason || 'سبب غير محدد');
    }

    if (result.stage === 'download') {
      return 'توقفت المعالجة نهائياً — فشل تحميل ملف التذكرة من الرابط';
    }
    if (result.stage === 'ocr') {
      return 'توقفت المعالجة نهائياً — فشل قراءة نص PDF (OCR)';
    }
    if (result.stage === 'claude') {
      return 'توقفت المعالجة نهائياً — فشل تحليل التذكرة عبر الذكاء الاصطناعي';
    }
    if (result.stage === 'find_row') {
      return 'توقفت المعالجة نهائياً — الجواز غير موجود في B2C';
    }
    return 'توقفت المعالجة نهائياً — خطأ: ' + (result.stage || result.status || 'غير محدد');
  },

  _describeDetailsArabic: function(result) {
    var d = result.details || {};

    if (result.stage === 'download') {
      if (d.status === 'error' && d.http_code) return 'الرابط يرد بخطأ HTTP ' + d.http_code;
      if (d.reason === 'file_too_small') return 'الملف صغير جداً (' + (d.size || 0) + ' بايت)';
      if (d.reason === 'exception') return 'خطأ شبكي أثناء التحميل: ' + (d.exception || 'غير محدد');
      if (d.reason === 'empty_url') return 'الرابط فارغ';
      return d.reason || '';
    }

    if (result.stage === 'ocr') {
      if (d.exception) return 'خطأ OCR: ' + d.exception;
      return d.reason || '';
    }

    if (result.stage === 'claude') {
      if (d.reason === 'max_attempts_reached') return 'فشل بعد عدة محاولات مع Claude API';
      if (d.reason === 'claude_response_not_json') return 'Claude أرجع رداً غير JSON صحيح';
      if (d.http_code) return 'Claude رد بخطأ HTTP ' + d.http_code;
      return d.reason || '';
    }

    if (result.status === 'validation_failed') {
      var errs = result.validation && result.validation.errors;
      if (!errs || errs.length === 0) return '';
      var lines = [];
      for (var i = 0; i < Math.min(errs.length, 3); i++) {
        var e = errs[i];
        lines.push('• ' + (e.layer || '') + ': ' + (e.reason || '') + (e.field ? ' (' + e.field + ')' : '') + (e.value !== undefined ? ' = ' + e.value : ''));
      }
      return lines.join('\n');
    }

    return '';
  },

  _updateStats: function(stats, result) {
    if (result.status === 'written') {
      stats.success_count++;
      if (result.family_result && result.family_result.applied_count > 0) {
        stats.family_applied += result.family_result.applied_count;
      }
    } else if (result.status === 'disaster') {
      stats.disaster_count++;
    } else if (result.status === 'validation_failed') {
      stats.failure_count++;
    } else if (result.stage === 'download' && result.details && result.details.status === 'not_pdf') {
      // يُعامَل كـ nusuk auth needed
      stats.nusuk_auth_count++;
      // تعليم BM=TRUE للصف
      if (result.pilgrim) {
        try {
          // نحتاج معرفة رقم الصف
          // سنتركه للمرحلة التالية — الآن نكتفي بالإحصائية
        } catch (e) {}
      }
    } else {
      stats.failure_count++;
    }

    // تجميع tokens
    if (result.parse_meta && result.parse_meta.tokens) {
      stats.total_tokens.input += (result.parse_meta.tokens.input_tokens || 0);
      stats.total_tokens.output += (result.parse_meta.tokens.output_tokens || 0);
    }
  }
};

// ==================== Global Entry Points ====================

function runPipeline() {
  return GDS2.Pipeline.run();
}

function runCanary() {
  return GDS2.Pipeline.runCanary();
}

function runBatch(limit) {
  return GDS2.Pipeline.run({ limit: limit });
}

function getPipelineStatus() {
  return GDS2.Pipeline.getStatus();
}
