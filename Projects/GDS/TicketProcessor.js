/**
 * TicketProcessor.js — منسّق معالجة تذكرة واحدة (orchestrator للمرحلة 3)
 *
 * الدور: يربط PdfDownloader + PdfTextExtractor + AnthropicClient + Prompts
 * بدون أي كتابة على شيت B2C (ذلك في المرحلة 5).
 *
 * دوال عامة:
 *   parseOneTicket(passport)    — معالجة حاج واحد
 *   testParseFirst3()           — اختبار على أول 3 حجاج في B2C
 *   parseByRow(rowNumber)       — معالجة حاج بموقع الصف
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.TicketProcessor = {
  /**
   * محاولة صبورة (Patient Retry) — محاولة تحميل أخيرة بمهلة أطول.
   * تُستدعى قبل تثبيت إيقاف نهائي بسبب فشل شبكي.
   * @param {string} url
   * @return {Object} { status: 'ok', bytes } أو { status: 'failed' }
   */
  patientDownload: function(url) {
    GDS2.Log.info('patientDownload: waiting 30s before final attempt', { url: url.substring(0, 80) });
    Utilities.sleep(30000);
    var result = GDS2.PdfDownloader.downloadBlob(url);
    return result;
  },

  /**
   * رأي Claude ثانٍ (Verifier) — يستدعي Claude بـ prompt مختلف للمراجعة.
   * @param {Array} pdfBytes
   * @param {Object} ctx - { name, passport, previous_result }
   * @return {Object} نتيجة Claude الجديدة (قد تختلف عن الأولى)
   */
  verifyWithClaude: function(pdfBytes, ctx) {
    GDS2.Log.info('verifyWithClaude: second opinion', { passport: ctx.passport });
    var prompts = GDS2.PromptBuilder.build('verifier', ctx);
    var response = GDS2.AnthropicClient.callWithPdf(prompts, pdfBytes, { maxTokens: 2000 });
    if (response.status !== 'ok') {
      return { status: 'error', error: response };
    }
    var parsed = GDS2.AnthropicClient.parseJSON(response.text);
    return {
      status: 'ok',
      parsed: parsed.ok ? parsed.data : null,
      raw: response.text
    };
  },

  /**
   * معالجة تذكرة حاج واحد.
   * @param {string} passport
   * @return {Object} تقرير كامل — بدون كتابة لأي شيت
   */
  processOne: function(passport, options) {
    var overallStart = new Date();

    // 1. البحث عن الصف في B2C
    var rowInfo = GDS2.TicketProcessor._findRow(passport);
    if (rowInfo.error) return { status: 'error', stage: 'lookup', reason: rowInfo.error };

    var name = rowInfo.name;
    var url = rowInfo.url;

    if (!url) return { status: 'error', stage: 'lookup', reason: 'no_ticket_url' };

    // 2. تحميل PDF كـ blob (بدون حفظ في Drive)
    var dl = GDS2.PdfDownloader.downloadBlob(url);

    // رقابة: إذا هذه آخر محاولة ومحاولة التحميل فشلت بخطأ عابر → patient retry
    if (dl.status !== 'ok' && options && options.audit === true &&
        dl.reason !== 'empty_url' && dl.reason !== 'file_too_small' && dl.status !== 'not_pdf') {
      GDS2.Log.info('Audit mode: running patient download retry', { passport: passport });
      var dl2 = GDS2.TicketProcessor.patientDownload(url);
      if (dl2.status === 'ok') {
        dl = dl2;
        dl.audit_recovered = true;
      }
    }

    // Rate limit (429) — تخطٍ مؤقت بدون تسجيل فشل
    if (dl.status === 'rate_limited') {
      return {
        status: 'rate_limited',
        stage: 'download',
        passport: passport,
        name: name,
        url: url,
        details: dl,
        total_elapsed_sec: (new Date() - overallStart) / 1000
      };
    }

    if (dl.status !== 'ok') {
      return {
        status: 'error',
        stage: 'download',
        passport: passport,
        name: name,
        url: url,
        details: dl,
        total_elapsed_sec: (new Date() - overallStart) / 1000
      };
    }

    // 3. Claude parsing مباشرة من PDF (بدون OCR!)
    var prompts = GDS2.PromptBuilder.build('parser', {
      name: name,
      passport: passport
    });

    var claudeResponse = GDS2.AnthropicClient.callWithPdf(prompts, dl.bytes, { maxTokens: 2000 });

    if (claudeResponse.status !== 'ok') {
      return {
        status: 'error',
        stage: 'claude',
        passport: passport,
        name: name,
        url: url,
        details: claudeResponse,
        total_elapsed_sec: (new Date() - overallStart) / 1000
      };
    }

    // 5. Parse JSON المُرجَع
    var parsed = GDS2.AnthropicClient.parseJSON(claudeResponse.text);

    // 6. رقابة Claude ثانية (Verifier): للنتائج "السلبية" نراجع بـ Claude ثانٍ
    var verifierApplied = false;
    var verifierResult = null;
    if (parsed.ok && parsed.data) {
      var data = parsed.data;
      var needsVerify = (data.found === false) || (data.too_many_segments === true);
      if (needsVerify) {
        var v = GDS2.TicketProcessor.verifyWithClaude(dl.bytes, {
          name: name,
          passport: passport,
          previous_result: data
        });
        verifierApplied = true;
        verifierResult = v;

        if (v.status === 'ok' && v.parsed) {
          // قبول النتيجة الأفضل (لصالح الحاج)
          var changed = false;
          if (data.found === false && v.parsed.found === true) {
            parsed.data = v.parsed;
            changed = true;
          } else if (data.too_many_segments === true && v.parsed.too_many_segments === false) {
            parsed.data = v.parsed;
            changed = true;
          }
          parsed.data.verifier_applied = true;
          parsed.data.verifier_changed_result = changed;
        }
      }
    }

    return {
      status: 'ok',
      stage: 'complete',
      passport: passport,
      name: name,
      url: url,
      row: rowInfo.row,
      pdf: {
        size: dl.size,
        elapsed_sec: dl.elapsed_sec
      },
      claude: {
        usage: claudeResponse.usage,
        model: claudeResponse.model,
        stop_reason: claudeResponse.stop_reason,
        attempt: claudeResponse.attempt
      },
      json_parse_ok: parsed.ok,
      parsed_data: parsed.ok ? parsed.data : null,
      raw_text_sample: claudeResponse.text.substring(0, 500),
      total_elapsed_sec: (new Date() - overallStart) / 1000
    };
  },

  /**
   * معالجة حاج بموقع الصف (للاختبار السريع).
   */
  processByRow: function(rowNumber) {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var sheet = ss.getSheetByName(GDS2.Config.SHEET_B2C);
    if (rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      return { status: 'error', reason: 'row_out_of_range', row: rowNumber };
    }
    var passport = String(sheet.getRange(rowNumber, GDS2.Config.COL.PASSPORT).getValue() || '').trim();
    if (!passport) return { status: 'error', reason: 'empty_passport_at_row', row: rowNumber };
    return GDS2.TicketProcessor.processOne(passport);
  },

  /**
   * اختبار على أول 3 صفوف في B2C.
   */
  testFirst3: function() {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var sheet = ss.getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = Math.min(sheet.getLastRow(), 4);
    if (lastRow < 2) return { status: 'error', reason: 'b2c_empty' };

    var results = [];
    for (var row = 2; row <= lastRow; row++) {
      var result = GDS2.TicketProcessor.processByRow(row);
      results.push({ row: row, result: result });
    }

    return {
      status: 'ok',
      tested: results.length,
      results: results
    };
  },

  // ==================== Private ====================

  _findRow: function(passport) {
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID);
    var sheet = ss.getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { error: 'b2c_empty' };

    var target = String(passport).trim();
    // قراءة A-Z (1-26) دفعة واحدة
    var data = sheet.getRange(2, 1, lastRow - 1, 26).getValues();

    for (var i = 0; i < data.length; i++) {
      var p = String(data[i][GDS2.Config.COL.PASSPORT - 1] || '').trim();
      if (p === target) {
        var firstNameEn = String(data[i][GDS2.Config.COL.FIRST_NAME_EN - 1] || '').trim();
        var lastNameEn = String(data[i][GDS2.Config.COL.LAST_NAME_EN - 1] || '').trim();
        var firstNameAr = String(data[i][GDS2.Config.COL.FIRST_NAME_AR - 1] || '').trim();
        var lastNameAr = String(data[i][GDS2.Config.COL.LAST_NAME_AR - 1] || '').trim();

        var name = (firstNameEn + ' ' + lastNameEn).trim();
        if (!name) name = (firstNameAr + ' ' + lastNameAr).trim();

        var url = String(data[i][GDS2.Config.COL.TICKET_URL - 1] || '').trim();

        return {
          row: i + 2,
          passport: p,
          name: name,
          name_en: (firstNameEn + ' ' + lastNameEn).trim(),
          name_ar: (firstNameAr + ' ' + lastNameAr).trim(),
          url: url
        };
      }
    }

    return { error: 'passport_not_found' };
  }
};

// ==================== Global Entry Points ====================

function parseOneTicket(passport) {
  return GDS2.TicketProcessor.processOne(passport);
}

function parseByRow(rowNumber) {
  return GDS2.TicketProcessor.processByRow(rowNumber);
}

function testParseFirst3() {
  return GDS2.TicketProcessor.testFirst3();
}

/**
 * Pipeline كامل: Parse + Validate + Disaster + Risk — مع دعم audit mode.
 */
function processWithValidation(passport, options) {
  var parseResult = GDS2.TicketProcessor.processOne(passport, options);
  if (parseResult.status !== 'ok') return parseResult;

  var pilgrim = {
    name: parseResult.name,
    passport: parseResult.passport,
    url: parseResult.url
  };

  var data = parseResult.parsed_data;

  var validation = GDS2.TicketValidator.validate(data);
  var disaster = GDS2.DisasterDetector.check(data, pilgrim);
  var risks = data && data.risks ? GDS2.RiskNotifier.format(data.risks, validation.warnings, pilgrim) : null;

  return {
    status: 'ok',
    pilgrim: pilgrim,
    parse_meta: {
      elapsed_sec: parseResult.total_elapsed_sec,
      tokens: parseResult.claude ? parseResult.claude.usage : null,
      model: parseResult.claude ? parseResult.claude.model : null,
      audit_mode: options && options.audit === true,
      verifier_applied: data && data.verifier_applied,
      verifier_changed: data && data.verifier_changed_result
    },
    parsed_data: data,
    validation: validation,
    disaster: disaster,
    risk_report: risks,
    risk_text: risks ? GDS2.RiskNotifier.toText(risks) : null,
    ready_to_write: validation.valid && !disaster.is_disaster
  };
}

function processAndWrite(passport) {
  // تحقق من BL الحالي — إذا = MAX-1، هذه آخر محاولة → audit mode
  var ssCheck = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRowCheck = ssCheck.getLastRow();
  var auditMode = false;
  if (lastRowCheck >= 2) {
    var ppsCheck = ssCheck.getRange(2, GDS2.Config.COL.PASSPORT, lastRowCheck - 1, 1).getValues();
    for (var kk = 0; kk < ppsCheck.length; kk++) {
      if (String(ppsCheck[kk][0] || '').trim() === String(passport).trim()) {
        var rowCheck = kk + 2;
        var currentBL = Number(ssCheck.getRange(rowCheck, GDS2.Config.COL.FAIL_COUNT).getValue()) || 0;
        auditMode = (currentBL === GDS2.Config.MAX_FAIL_ATTEMPTS - 1);
        break;
      }
    }
  }

  var parseResult = processWithValidation(passport, { audit: auditMode });

  // Rate limit مؤقت (429) — تخطٍ بدون تسجيل فشل ولا إشعار
  if (parseResult.status === 'rate_limited') {
    return parseResult;
  }

  // لو الفشل قبل Claude (download/ocr/claude): سجّل فشل على الصف
  if (parseResult.status !== 'ok') {
    var r = parseResult;
    var ss = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    var lastRowSearch = ss.getLastRow();
    if (lastRowSearch >= 2) {
      var passports2 = ss.getRange(2, GDS2.Config.COL.PASSPORT, lastRowSearch - 1, 1).getValues();
      for (var k = 0; k < passports2.length; k++) {
        if (String(passports2[k][0] || '').trim() === String(passport).trim()) {
          var row2 = k + 2;

          // فشل عادي (تحميل/OCR/Claude) — ألغينا حالة "نسك محمي" خاصة
          var fr = GDS2.FlightWriter.recordFailure(row2, r.stage || 'error');
          r.row = row2;
          r.bl_after = fr.new_fail_count;
          r.stopped = fr.stopped;

          // إصلاح: pilgrim info
          if (!r.pilgrim) {
            r.pilgrim = {
              name: r.name || '—',
              passport: r.passport || passport,
              url: r.url || ''
            };
          }
          break;
        }
      }
    }
    return r;
  }

  var pilgrim = parseResult.pilgrim;
  var validation = parseResult.validation;
  var disaster = parseResult.disaster;
  var result = parseResult; // alias

  // البحث عن صف B2C
  var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRow = sheet.getLastRow();
  var passports = sheet.getRange(2, GDS2.Config.COL.PASSPORT, lastRow - 1, 1).getValues();
  var targetRow = -1;
  for (var i = 0; i < passports.length; i++) {
    if (String(passports[i][0] || '').trim() === String(pilgrim.passport).trim()) {
      targetRow = i + 2;
      break;
    }
  }
  if (targetRow < 0) {
    return {
      status: 'error',
      stage: 'find_row',
      reason: 'passport_not_found_in_b2c',
      pilgrim: pilgrim
    };
  }

  // Disaster → لا كتابة، لكن نُثبّت BL = MAX لإيقاف المحاولات (منع تكرار الإشعار)
  if (disaster.is_disaster) {
    var sheetD = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
    sheetD.getRange(targetRow, GDS2.Config.COL.FAIL_COUNT).setValue(GDS2.Config.MAX_FAIL_ATTEMPTS);
    sheetD.getRange(targetRow, 1, 1, GDS2.Config.COL.NUSUK_AUTH).setBackground(GDS2.Config.ROW_COLORS.ERROR);
    return {
      status: 'disaster',
      row: targetRow,
      pilgrim: pilgrim,
      disaster: disaster,
      stopped: true
    };
  }

  // Validation failed → تسجيل فشل
  if (!validation.valid) {
    var failResult = GDS2.FlightWriter.recordFailure(targetRow, 'validation_failed');
    return {
      status: 'validation_failed',
      row: targetRow,
      pilgrim: pilgrim,
      validation: validation,
      write_result: failResult,
      bl_after: failResult.new_fail_count,
      stopped: failResult.stopped
    };
  }

  // نجاح — كتابة شاملة
  var writeResult = GDS2.FlightWriter.writeSuccess(targetRow, validation.normalized, pilgrim.url);

  // تطبيق على العائلة (لو وُجد)
  var familyResult = null;
  if (writeResult.status === 'ok') {
    familyResult = GDS2.FamilyProcessor.applyToFamily(validation.normalized, pilgrim.url, targetRow);
  }

  return {
    status: writeResult.status === 'ok' ? 'written' : writeResult.status,
    row: targetRow,
    pilgrim: pilgrim,
    risk_text: result.risk_text,
    write_result: writeResult,
    family_result: familyResult,
    elapsed_sec: result.parse_meta.elapsed_sec
  };
}

/**
 * اختبار كتابة على أول حاج فقط (قبل batch أكبر).
 */
function testWriteFirstOne() {
  var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var passport = String(sheet.getRange(2, GDS2.Config.COL.PASSPORT).getValue() || '').trim();
  if (!passport) return { status: 'error', reason: 'empty_passport' };
  return processAndWrite(passport);
}

/**
 * اختبار كتابة على أول 3 حجاج.
 */
function testWriteFirst3() {
  var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRow = Math.min(sheet.getLastRow(), 4);
  var results = [];
  for (var row = 2; row <= lastRow; row++) {
    var passport = String(sheet.getRange(row, GDS2.Config.COL.PASSPORT).getValue() || '').trim();
    if (!passport) continue;
    var r = processAndWrite(passport);
    results.push({
      row: row,
      status: r.status,
      passport: passport,
      elapsed: r.elapsed_sec,
      family_applied: r.family_result ? r.family_result.applied_count : 0
    });
  }
  return { tested: results.length, results: results };
}

function testFullPipelineFirst3() {
  var sheet = SpreadsheetApp.openById(GDS2.Config.SS_ID).getSheetByName(GDS2.Config.SHEET_B2C);
  var lastRow = Math.min(sheet.getLastRow(), 4);
  var results = [];
  for (var row = 2; row <= lastRow; row++) {
    var passport = String(sheet.getRange(row, GDS2.Config.COL.PASSPORT).getValue() || '').trim();
    if (!passport) continue;
    var result = processWithValidation(passport);
    results.push({
      row: row,
      ready_to_write: result.ready_to_write,
      validation_valid: result.validation ? result.validation.valid : null,
      error_count: result.validation ? (result.validation.errors || []).length : 0,
      warning_count: result.validation ? (result.validation.warnings || []).length : 0,
      is_disaster: result.disaster ? result.disaster.is_disaster : null,
      detail: result
    });
  }
  return { tested: results.length, results: results };
}
