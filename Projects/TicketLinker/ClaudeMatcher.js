/**
 * ClaudeMatcher.js — المرحلة 4: قراءة PDF بـ Claude API
 *
 * متى تُستدعى: لو فشلت المرحلة 1 و3 (لم يُطابق الاسم في pool ولا في PD كامل).
 *
 * المنطق:
 *   1. نستخرج نص PDF (OCR via Drive.Files.copy)
 *   2. نرسل النص لـ Claude API مع prompt واضح
 *   3. Claude يستخرج: fullName, firstName, lastName, pnr, ticketNumber
 *   4. نعيد البحث في PD بالاسم الجديد
 *   5. لو وُجد → ✓، لو لا → نمرّر للمرحلة 5 مع البيانات المستخرَجة
 */

/**
 * يستدعي Claude API
 * @param {string} pdfText
 * @return {Object} {fullName, firstName, lastName, pnr, ticketNumber, dob, allPassengers}
 */
function callClaudeForPdfExtraction_(pdfText) {
  var apiKey = PropertiesService.getScriptProperties().getProperty(TL.Config.PROP.ANTHROPIC_API_KEY);
  if (!apiKey) {
    throw new Error('ANTHROPIC_API_KEY not set in Script Properties');
  }

  // قاعدة ثابتة: لا يوجد رقم جواز في أي تذكرة طيران في العالم.
  // المطابقة تعتمد على الاسم فقط — الجواز يُؤخذ من قاعدة البيانات (PD).
  var prompt = 'هذا نص تذكرة طيران. ' +
    'استخرج البيانات التالية بصيغة JSON دقيقة (بدون أي نص آخر قبل أو بعد JSON):\n' +
    '{\n' +
    '  "fullName": "اسم المسافر الكامل بالإنجليزية كما يظهر في التذكرة (Passenger Name / Traveler)",\n' +
    '  "firstName": "الاسم الأول فقط (بدون ألقاب Mr/Mrs/Miss)",\n' +
    '  "lastName": "اسم العائلة فقط",\n' +
    '  "pnr": "رمز الحجز PNR أو Booking Ref (6 أحرف/أرقام)",\n' +
    '  "ticketNumber": "رقم التذكرة إن وُجد",\n' +
    '  "allPassengers": ["مصفوفة بكل أسماء الركاب — مفيد جداً لو التذكرة جماعية"],\n' +
    '  "flights": [{"flightNo":"", "depDate":"", "depTime":"", "from":"", "to":"", "arrDate":"", "arrTime":""}]\n' +
    '}\n\n' +
    'تعليمات مهمة:\n' +
    '- لا تبحث عن رقم جواز سفر — لا تذكرة في العالم تحتوي رقم جواز\n' +
    '- ركّز على استخراج الاسم بدقة (fullName هو الأهم)\n' +
    '- لو حقل غير موجود استخدم null أو [] (لا تخمّن)\n\n' +
    'النص:\n' + pdfText.substring(0, 12000);

  var payload = {
    model: TL.Config.CLAUDE_MODEL,
    max_tokens: TL.Config.CLAUDE_MAX_TOKENS,
    messages: [{ role: 'user', content: prompt }]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': TL.Config.CLAUDE_ANTHROPIC_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(TL.Config.CLAUDE_API_URL, options);
  var code = response.getResponseCode();
  var body = response.getContentText();

  if (code !== 200) {
    throw new Error('Claude API error ' + code + ': ' + body.substring(0, 500));
  }

  var data = JSON.parse(body);
  var text = data.content && data.content[0] && data.content[0].text || '';

  // استخرج JSON من النص (قد يكون محاطاً بـ markdown أو شرح)
  var jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('No JSON found in Claude response: ' + text.substring(0, 200));
  }

  return JSON.parse(jsonMatch[0]);
}

/**
 * المرحلة 4 الكاملة:
 *   1. استخراج نص PDF
 *   2. استدعاء Claude
 *   3. إعادة البحث في PD بالاسم/PNR الجديد
 */
function matchViaPdfClaudeExtraction_(pdfInfo) {
  var result = {
    fileId: pdfInfo.fileId,
    fileName: pdfInfo.fileName,
    folderName: pdfInfo.folderName,
    stage: 4,
    status: null,
    passport: null,
    claudeData: null,
    notes: []
  };

  // 1. استخراج نص PDF
  var text = extractPdfText_(pdfInfo.fileId);
  if (!text || text.length < 50) {
    result.status = 'PDF_TEXT_EMPTY';
    result.notes.push('PDF text extraction failed or too short');
    return result;
  }

  // 2. استدعاء Claude
  var claudeData;
  try {
    claudeData = callClaudeForPdfExtraction_(text);
    result.claudeData = claudeData;
  } catch (e) {
    result.status = 'CLAUDE_ERROR';
    result.notes.push('Claude API error: ' + e.message);
    return result;
  }

  if (!claudeData || !claudeData.fullName) {
    result.status = 'CLAUDE_NO_NAME';
    result.notes.push('Claude returned no fullName');
    return result;
  }

  // 3. إعادة البحث في PD بالاسم الجديد
  var pdMatch = searchNameInFullPd_(claudeData.fullName);
  if (pdMatch && pdMatch.collision) {
    result.status = 'COLLISION';
    result.collision = pdMatch.matches;
    return result;
  }
  if (pdMatch) {
    result.status = 'OK';
    result.passport = pdMatch.passport;
    result.matchedFirstName = pdMatch.firstName;
    result.matchedLastName = pdMatch.lastName;
    result.notes.push('Found via Claude-extracted name: ' + claudeData.fullName);
    return result;
  }

  // 4. جرّب firstName + lastName من Claude مباشرة
  if (claudeData.firstName && claudeData.lastName) {
    var combined = claudeData.firstName + ' ' + claudeData.lastName;
    pdMatch = searchNameInFullPd_(combined);
    if (pdMatch && !pdMatch.collision) {
      result.status = 'OK';
      result.passport = pdMatch.passport;
      result.matchedFirstName = pdMatch.firstName;
      result.matchedLastName = pdMatch.lastName;
      result.notes.push('Found via firstName+lastName from Claude');
      return result;
    }
  }

  // فشلت كل المحاولات → سيُحال للمرحلة 5
  result.status = 'GO_TO_STAGE_5';
  result.notes.push('Claude extracted name but not found in PD');
  return result;
}

// ==================== اختبار ====================

function testClaudeOnPdf(fileId) {
  var text = extractPdfText_(fileId);
  if (!text) return { error: 'PDF text empty' };
  return callClaudeForPdfExtraction_(text);
}

/**
 * تعيين Anthropic API key (للاستدعاء عن بُعد)
 */
function setAnthropicApiKey(key) {
  PropertiesService.getScriptProperties().setProperty(TL.Config.PROP.ANTHROPIC_API_KEY, key);
  return { ok: true, set: TL.Config.PROP.ANTHROPIC_API_KEY };
}
