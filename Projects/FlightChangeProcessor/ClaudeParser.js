/**
 * ClaudeParser.js — استخراج بيانات PDF نسك عبر Claude API
 *
 * يحل محل: extractTextFromPDF_ + parseNusukTicket_ + extractAllPassengers_
 *
 * المزايا:
 * • يقرأ PDF عربي/إنجليزي مباشرة (بدون OCR)
 * • يستخرج كل أفراد العائلة
 * • يُخرج JSON نظيف مع confidence score
 * • حماية ذاتية أدنى داخلية
 */


var CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
var CLAUDE_MODEL = 'claude-sonnet-4-5';
var CLAUDE_MAX_TOKENS = 2000;


/**
 * استخراج البيانات من PDF نسك عبر Claude
 *
 * @param {Blob} pdfBlob - ملف PDF
 * @return {Object|null} - البيانات المستخرجة أو null عند الفشل
 */
function parseWithClaude_(pdfBlob) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) {
    Logger.log('❌ CLAUDE_API_KEY غير موجود في Script Properties');
    return null;
  }

  var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());

  var systemPrompt = 'أنت نظام استخراج بيانات من تذاكر نسك الإلكترونية. تستخرج البيانات بدقة وترجعها JSON فقط، بدون أي نص آخر.';

  var userPrompt = [
    'استخرج من PDF نسك المرفق:',
    '',
    '1. كل المسافرين بأسمائهم الكاملة (عربي أو إنجليزي كما في PDF).',
    '   - كل مسافر له: firstName و lastName منفصلين',
    '   - age = "A" لـ Adult, "C" لـ Child, "I" لـ Infant',
    '',
    '2. PNR (Airline Reference)',
    '',
    '3. Booking ID (Nusuk Hajj ID، يبدأ بـ A متبوعاً بأرقام)',
    '',
    '4. رحلات الذهاب (outbound): كل leg منفصلاً',
    '   - flightNumber, fromCity (IATA 3 حرف), toCity (IATA 3 حرف)',
    '   - depDate (YYYY-MM-DD), depTime (HH:MM)',
    '   - arrDate (YYYY-MM-DD), arrTime (HH:MM)',
    '',
    '5. رحلات العودة (return): بنفس الهيكل',
    '',
    '6. overall_confidence: ثقتك من 0-100 في دقة البيانات المستخرجة',
    '',
    'أرجع JSON فقط بهذا الشكل:',
    '{',
    '  "pnr": "...",',
    '  "bookingId": "...",',
    '  "passengers": [{"firstName": "...", "lastName": "...", "age": "A"}],',
    '  "outboundLegs": [{"flightNumber": "...", "fromCity": "...", "toCity": "...", "depDate": "...", "depTime": "...", "arrDate": "...", "arrTime": "..."}],',
    '  "returnLegs": [...],',
    '  "overall_confidence": 95',
    '}',
    '',
    'إن لم يكن PDF تذكرة نسك صالحة، أرجع: {"error": "not_nusuk_ticket", "overall_confidence": 0}'
  ].join('\n');

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: CLAUDE_MAX_TOKENS,
    system: systemPrompt,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'document',
          source: {
            type: 'base64',
            media_type: 'application/pdf',
            data: pdfBase64
          }
        },
        {
          type: 'text',
          text: userPrompt
        }
      ]
    }]
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

  // retry عند 529/503/429
  var maxRetries = 3;
  var response = null;
  var lastError = null;

  for (var attempt = 0; attempt < maxRetries; attempt++) {
    try {
      response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
      var code = response.getResponseCode();

      if (code === 200) {
        break;
      }

      if (code === 529 || code === 503 || code === 429) {
        lastError = 'HTTP ' + code;
        if (attempt < maxRetries - 1) {
          Utilities.sleep(1000 * Math.pow(3, attempt)); // 1s, 3s, 9s
          continue;
        }
      }

      // خطأ آخر — لا retry
      Logger.log('  ❌ Claude API error ' + code + ': ' + response.getContentText().substring(0, 200));
      return null;
    } catch (e) {
      lastError = e.message;
      if (attempt < maxRetries - 1) {
        Utilities.sleep(1000 * Math.pow(3, attempt));
        continue;
      }
      Logger.log('  ❌ Claude fetch failed: ' + e.message);
      return null;
    }
  }

  if (!response || response.getResponseCode() !== 200) {
    Logger.log('  ❌ Claude failed after retries: ' + lastError);
    return null;
  }

  var data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log('  ❌ Invalid JSON response from Claude');
    return null;
  }

  if (!data.content || !data.content[0] || !data.content[0].text) {
    Logger.log('  ❌ Claude response has no content');
    return null;
  }

  // استخراج JSON من النص
  var text = data.content[0].text.trim();

  // إزالة code fences إن وُجدت
  text = text.replace(/^```(?:json)?\s*\n?/, '').replace(/\n?```\s*$/, '').trim();

  var parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    // محاولة استخراج JSON من داخل النص
    var jsonMatch = text.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        parsed = JSON.parse(jsonMatch[0]);
      } catch (e2) {
        Logger.log('  ❌ Cannot parse JSON from Claude response');
        return null;
      }
    } else {
      Logger.log('  ❌ No JSON found in Claude response');
      return null;
    }
  }

  // فحص الخطأ
  if (parsed.error) {
    Logger.log('  ⏭️ Claude: ' + parsed.error + ' (confidence ' + (parsed.overall_confidence || 0) + ')');
    return { skipped: true, reason: parsed.error, confidence: parsed.overall_confidence || 0 };
  }

  // فحص الثقة
  var confidence = parsed.overall_confidence || 0;
  if (confidence < 70) {
    Logger.log('  ⏭️ Claude confidence ' + confidence + '% — تخطي');
    return { skipped: true, reason: 'low_confidence', confidence: confidence };
  }

  // تطبيع البيانات لتتوافق مع هيكل parseNusukTicket_ القديم
  var result = {
    pnr: parsed.pnr || '',
    bookingId: parsed.bookingId || '',
    passengers: (parsed.passengers || []).map(function(p) {
      var fn = String(p.firstName || '').trim().toUpperCase();
      var ln = String(p.lastName || '').trim().toUpperCase();
      return {
        firstName: fn,
        lastName: ln,
        fullName: (fn + ' ' + ln).trim(),
        age: p.age || 'A'
      };
    }),
    outboundLegs: parsed.outboundLegs || [],
    returnLegs: parsed.returnLegs || [],
    passengerName: '',
    firstName: '',
    lastName: '',
    arrivalCity: '',
    arrivalDate: '',
    arrivalTime: '',
    departureCity: '',
    departureDate: '',
    departureTime: '',
    confidence: confidence,
    source: 'claude-v5'
  };

  // التوافق — أول مسافر
  if (result.passengers.length > 0) {
    result.passengerName = result.passengers[0].fullName;
    result.firstName = result.passengers[0].firstName;
    result.lastName = result.passengers[0].lastName;
  }

  // حساب تواريخ المملكة
  try { calculateKSADates_(result); } catch (e) {}

  Logger.log('  ✅ Claude: ' + result.passengers.length + ' مسافر | ثقة ' + confidence + '%');
  return result;
}


/**
 * استخراج بيانات تغيير رحلة من نص إيميل (بدون PDF)
 *
 * @param {string} plainBody - نص الإيميل
 * @return {Object|null} - البيانات المستخرجة أو skip
 */
function parseTextWithClaude_(plainBody) {
  if (!plainBody || plainBody.length < 100) return null;

  var apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;

  var systemPrompt = 'أنت نظام استخراج بيانات تغييرات الطيران من إيميلات نسك. تُرجع JSON فقط.';

  var userPrompt = [
    'استخرج من نص الإيميل التالي (من نسك) بيانات تغيير الرحلات:',
    '',
    '1. كل المسافرين (إن ذُكروا) — firstName + lastName',
    '2. PNR / Reservation / Airline PNR / Booking',
    '3. Booking ID / Reference (مثلاً A26020191393250)',
    '4. رحلات الذهاب (outbound) — كل leg: flightNumber, fromCity (IATA), toCity (IATA), depDate (YYYY-MM-DD), depTime (HH:MM), arrDate, arrTime',
    '5. رحلات العودة (return) — نفس الهيكل',
    '6. overall_confidence (0-100)',
    '',
    'ملاحظات مهمة:',
    '• إن لم يكن النص يحوي بيانات رحلة فعلية (مثلاً مجرد "راجع المرفق" أو مراسلة إدارية)، أرجع: {"error": "no_flight_data", "overall_confidence": 0}',
    '• إن كان النص مجرد رد/شكوى/سؤال، أرجع: {"error": "correspondence", "overall_confidence": 0}',
    '• التواريخ بالميلادي فقط (YYYY-MM-DD)، حوّل "الخميس 21 مايو 2026" → "2026-05-21"',
    '• المدن: استخرج IATA فقط (AMS, JED, FCO, KWI إلخ)',
    '• إن لم يُذكر رقم الرحلة، ضع "" فارغة',
    '',
    'أرجع JSON فقط:',
    '{',
    '  "pnr": "...",',
    '  "bookingId": "...",',
    '  "passengers": [{"firstName": "...", "lastName": "...", "age": "A"}],',
    '  "outboundLegs": [{...}],',
    '  "returnLegs": [...],',
    '  "overall_confidence": 90',
    '}',
    '',
    '=== نص الإيميل ===',
    plainBody.substring(0, 8000)
  ].join('\n');

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: CLAUDE_MAX_TOKENS,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }]
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

  // retry على 529/503/429
  var maxRetries = 3;
  var response = null;

  for (var attempt = 0; attempt < maxRetries; attempt++) {
    try {
      response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
      var code = response.getResponseCode();
      if (code === 200) break;
      if (code === 529 || code === 503 || code === 429) {
        if (attempt < maxRetries - 1) {
          Utilities.sleep(1000 * Math.pow(3, attempt));
          continue;
        }
      }
      Logger.log('  ❌ Claude text error ' + code);
      return null;
    } catch (e) {
      if (attempt < maxRetries - 1) {
        Utilities.sleep(1000 * Math.pow(3, attempt));
        continue;
      }
      return null;
    }
  }

  if (!response || response.getResponseCode() !== 200) return null;

  var data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {
    return null;
  }

  if (!data.content || !data.content[0]) return null;

  var text = data.content[0].text.trim()
    .replace(/^```(?:json)?\s*\n?/, '')
    .replace(/\n?```\s*$/, '').trim();

  var parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    var m = text.match(/\{[\s\S]*\}/);
    if (m) {
      try { parsed = JSON.parse(m[0]); } catch (e2) { return null; }
    } else return null;
  }

  if (parsed.error) {
    Logger.log('  ⏭️ Claude-text: ' + parsed.error);
    return { skipped: true, reason: parsed.error, confidence: 0 };
  }

  var confidence = parsed.overall_confidence || 0;
  if (confidence < 70) {
    Logger.log('  ⏭️ Claude-text confidence ' + confidence + '%');
    return { skipped: true, reason: 'low_confidence_text', confidence: confidence };
  }

  // تحقق من وجود بيانات رحلة فعلية
  var hasOutbound = parsed.outboundLegs && parsed.outboundLegs.length > 0;
  var hasReturn = parsed.returnLegs && parsed.returnLegs.length > 0;
  if (!hasOutbound && !hasReturn) {
    return { skipped: true, reason: 'no_legs_extracted', confidence: confidence };
  }

  var result = {
    pnr: parsed.pnr || '',
    bookingId: parsed.bookingId || '',
    passengers: (parsed.passengers || []).map(function(p) {
      var fn = String(p.firstName || '').trim().toUpperCase();
      var ln = String(p.lastName || '').trim().toUpperCase();
      return {
        firstName: fn,
        lastName: ln,
        fullName: (fn + ' ' + ln).trim(),
        age: p.age || 'A'
      };
    }),
    outboundLegs: parsed.outboundLegs || [],
    returnLegs: parsed.returnLegs || [],
    passengerName: '',
    firstName: '',
    lastName: '',
    arrivalCity: '', arrivalDate: '', arrivalTime: '',
    departureCity: '', departureDate: '', departureTime: '',
    confidence: confidence,
    source: 'claude-v5-text'
  };

  if (result.passengers.length > 0) {
    result.passengerName = result.passengers[0].fullName;
    result.firstName = result.passengers[0].firstName;
    result.lastName = result.passengers[0].lastName;
  }

  try { calculateKSADates_(result); } catch (e) {}

  Logger.log('  ✅ Claude-text: ' + result.outboundLegs.length + '+' + result.returnLegs.length + ' رحلات | ثقة ' + confidence + '%');
  return result;
}


/**
 * حماية أدنى — فحوصات سريعة على نتائج Claude قبل الكتابة
 * @return {Object} - {valid: bool, reason: string}
 */
function validateClaudeResult_(result, allData) {
  if (!result || result.skipped) return { valid: false, reason: result ? result.reason : 'no_result' };

  // 1. مسافرون أو PNR — الإيميلات النصية من نسك قد تكتفي بـ PNR
  var hasPassengers = result.passengers && result.passengers.length > 0;
  var hasPnr = result.pnr && String(result.pnr).trim().length >= 4;
  if (!hasPassengers && !hasPnr) {
    return { valid: false, reason: 'no_passengers_and_no_pnr' };
  }

  // 2. PNR موجود في قاعدتنا (إن وُجد PNR)
  if (result.pnr) {
    var pnrUpper = result.pnr.toUpperCase().trim();
    var foundInPD = false;
    for (var i = 0; i < allData.pd.length && !foundInPD; i++) {
      var contract = String(allData.pd[i][CONFIG.PD.CONTRACT_NAME] || '').toUpperCase();
      if (contract.indexOf(pnrUpper) !== -1) foundInPD = true;
    }
    var foundInB2C = false;
    if (!foundInPD && allData.b2c.rows.length > 0) {
      var pnrCol = findCol_(allData.b2c.headers, 'Airline PNR');
      if (pnrCol !== -1) {
        for (var j = 0; j < allData.b2c.rows.length && !foundInB2C; j++) {
          var rowPnr = String(allData.b2c.rows[j][pnrCol] || '').toUpperCase().trim();
          if (rowPnr === pnrUpper) foundInB2C = true;
        }
      }
    }
    if (!foundInPD && !foundInB2C) {
      return { valid: false, reason: 'pnr_not_in_database:' + result.pnr };
    }
  }

  // 3. تواريخ ضمن موسم الحج 2026 (مايو-أغسطس)
  var validFrom = new Date('2026-04-01');
  var validTo = new Date('2026-08-31');
  var allLegs = (result.outboundLegs || []).concat(result.returnLegs || []);
  for (var k = 0; k < allLegs.length; k++) {
    var leg = allLegs[k];
    if (leg.depDate) {
      var d = new Date(leg.depDate);
      if (!isNaN(d.getTime())) {
        if (d < validFrom || d > validTo) {
          return { valid: false, reason: 'date_out_of_season:' + leg.depDate };
        }
      }
    }
  }

  return { valid: true, reason: 'ok' };
}
