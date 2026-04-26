/**
 * _SkippedClassifier.js — تصنيف Skipped في labels فرعية للمراجعة اليدوية
 *
 * المبدأ: لا إيميل يُهمَل. كل إيميل Skipped يُصنَّف تحت فئة واضحة.
 */


// Labels الفرعية
var SKIP_LABELS = {
  CORRESPONDENCE:    'TKT-Skipped/Correspondence',      // ردود، شكاوى، أسئلة
  OTHER_COMPANIES:   'TKT-Skipped/OtherCompanies',      // حجاج شركات أخرى (PNR ليس في قاعدتنا)
  NO_FLIGHT_DATA:    'TKT-Skipped/NoFlightData',        // محتوى عام بلا بيانات رحلة
  EMPTY_CONTENT:     'TKT-Skipped/EmptyContent',        // فارغ أو "راجع المرفق" بلا مرفق
  LOW_CONFIDENCE:    'TKT-Skipped/LowConfidence',       // Claude ثقة < 70%
  NEEDS_REVIEW:      'TKT-Skipped/NeedsReview',         // حالة غامضة — يراجعها موظف
  PARSE_FAILED:      'TKT-Skipped/ParseFailed',         // فشل تقني
  NOT_SEASON:        'TKT-Skipped/NotSeason',           // تواريخ خارج موسم 2026
};


/**
 * يصنّف كل إيميلات Skipped الحالية إلى labels فرعية
 * يحافظ على label TKT-Skipped الأصلي (لا يحذفه)
 */
function classifyAllSkipped() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return { error: 'CLAUDE_API_KEY missing' };

  // تأكد من وجود كل Labels
  var labels = {};
  for (var k in SKIP_LABELS) {
    labels[k] = getOrCreateLabel_(SKIP_LABELS[k]);
  }

  var query = 'label:' + CONFIG.SKIPPED_LABEL;
  var threads = GmailApp.search(query, 0, 100);

  var stats = {
    total: threads.length,
    classified: {},
    errors: 0
  };
  for (var kk in SKIP_LABELS) stats.classified[kk] = 0;

  Logger.log('🏷️ تصنيف ' + threads.length + ' thread في labels فرعية');

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    var primaryMsg = messages[messages.length - 1]; // آخر رسالة

    try {
      var classification = classifyEmail_(primaryMsg);

      if (!classification || !classification.category) {
        Logger.log('  [' + (i+1) + '/' + threads.length + '] ⚠️ تصنيف فشل → NeedsReview');
        thread.addLabel(labels.NEEDS_REVIEW);
        stats.classified.NEEDS_REVIEW++;
        continue;
      }

      var cat = classification.category;
      if (!labels[cat]) cat = 'NEEDS_REVIEW';

      thread.addLabel(labels[cat]);
      stats.classified[cat]++;

      var subj = primaryMsg.getSubject().substring(0, 60);
      Logger.log('  [' + (i+1) + '/' + threads.length + '] ' + SKIP_LABELS[cat] + ' ← ' + subj);
    } catch (e) {
      Logger.log('  [' + (i+1) + '/' + threads.length + '] ❌ ' + e.message);
      thread.addLabel(labels.NEEDS_REVIEW);
      stats.classified.NEEDS_REVIEW++;
      stats.errors++;
    }
  }

  Logger.log('');
  Logger.log('=== التوزيع ===');
  for (var kkk in stats.classified) {
    if (stats.classified[kkk] > 0) {
      Logger.log('  ' + SKIP_LABELS[kkk] + ': ' + stats.classified[kkk]);
    }
  }

  return stats;
}


/**
 * يحلّل إيميل واحد ويصنّفه باستخدام Claude
 */
function classifyEmail_(message) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;

  var subject = message.getSubject() || '';
  var body = message.getPlainBody() || '';
  var attachments = message.getAttachments();
  var hasPdf = false;
  for (var a = 0; a < attachments.length; a++) {
    if (attachments[a].getContentType() === 'application/pdf') { hasPdf = true; break; }
  }

  // إن كان الإيميل فارغاً أو بلا محتوى ذي معنى
  if (!body || body.trim().length < 50) {
    return { category: 'EMPTY_CONTENT', reason: 'body_too_short' };
  }

  var prompt = [
    'أنت مصنّف إيميلات لشركة حج. صنّف الإيميل التالي في فئة واحدة:',
    '',
    'الفئات:',
    '• CORRESPONDENCE — ردّ من حاج، شكوى، سؤال، تأكيد، مراسلة إدارية بلا بيانات طيران',
    '• OTHER_COMPANIES — فيه بيانات حجز لكن حاج شركة أخرى (ليس عميلنا)',
    '• NO_FLIGHT_DATA — محتوى عام لا يحوي بيانات رحلة قابلة للاستخراج',
    '• EMPTY_CONTENT — يقول "راجع المرفق" أو "تم إرسال التذكرة" لكن لا مرفق في هذا الإيميل',
    '• LOW_CONFIDENCE — فيه بيانات لكن غير واضحة/ناقصة يصعب الوثوق بها',
    '• NEEDS_REVIEW — حالة غامضة تحتاج موظف يقرر',
    '',
    'أرجع JSON فقط:',
    '{"category": "CORRESPONDENCE", "reason": "شرح قصير بالعربية"}',
    '',
    '=== الموضوع ===',
    subject,
    '',
    '=== PDF مرفق: ' + (hasPdf ? 'نعم' : 'لا') + ' ===',
    '',
    '=== النص ===',
    body.substring(0, 4000)
  ].join('\n');

  var payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 300,
    messages: [{ role: 'user', content: prompt }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
    if (response.getResponseCode() !== 200) return null;

    var data = JSON.parse(response.getContentText());
    if (!data.content || !data.content[0]) return null;

    var text = data.content[0].text.trim()
      .replace(/^```(?:json)?\s*\n?/, '')
      .replace(/\n?```\s*$/, '').trim();

    var m = text.match(/\{[\s\S]*\}/);
    if (!m) return null;

    return JSON.parse(m[0]);
  } catch (e) {
    return null;
  }
}


/**
 * عرض التوزيع الحالي للـ labels الفرعية
 */
function showSkippedDistribution() {
  var result = { total: 0, breakdown: {} };
  for (var k in SKIP_LABELS) {
    var query = 'label:' + SKIP_LABELS[k];
    var threads = GmailApp.search(query, 0, 100);
    result.breakdown[SKIP_LABELS[k]] = threads.length;
    result.total += threads.length;
  }
  return result;
}
