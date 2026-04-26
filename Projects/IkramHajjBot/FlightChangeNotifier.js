// ============================================
// FlightChangeNotifier.js
// إخطار الحاج عند تغيير رحلته في V3
// يقرأ شيت "تغييرات الطيران V3" ويُرسل رسائل تيليغرام
// ============================================

/**
 * الدالة الرئيسية
 * تقرأ V3، تطابق بـ passport، تُرسل (أو Dry-run حسب الإعداد)
 * تُرجع { total, sent, dryRun, notRegistered, error }
 */
function notifyFlightChanges() {
  // ١. فحص kill switch
  if (!FLIGHT_NOTIFY_ENABLED) {
    Logger.log('🛑 FLIGHT_NOTIFY_ENABLED = false — لا إرسال');
    return { total: 0, reason: 'disabled' };
  }

  // ٢. فتح شيت V3 من الجدول الخارجي
  var ss = SpreadsheetApp.openById(FLIGHT_CHANGES_SHEET_ID);
  var sheet = ss.getSheetByName(FLIGHT_CHANGES_SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ شيت V3 غير موجود: ' + FLIGHT_CHANGES_SHEET_NAME);
    return { total: 0, reason: 'sheet_not_found' };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, reason: 'empty_sheet' };

  // ٣. بناء passport → chatId map (قراءة BotSessions مرّة واحدة)
  var passportMap = buildPassportToChatMap_();
  Logger.log('📇 عدد الحجاج المسجلين في البوت: ' + Object.keys(passportMap).length);

  // ٤. قراءة كل صفوف V3 (الأعمدة A → AP = 42 عمود)
  var data = sheet.getRange(2, 1, lastRow - 1, 42).getValues();
  var stats = { total: 0, sent: 0, dryRun: 0, notRegistered: 0, error: 0, skippedNight: 0 };
  var updates = []; // [{ rowIndex, status, timestamp }]
  var processed = 0;

  // ٥. منع الإزعاج الليلي (11م - 7ص توقيت الرياض)
  var hour = parseInt(Utilities.formatDate(new Date(), 'Asia/Riyadh', 'H'), 10);
  var isNightTime = (hour >= 23 || hour < 7);

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowIndex = i + 2; // index في الشيت
    var notifyStatus = String(row[40] || '').trim(); // AO

    // تخطي: غير فارغ = مُعالَج سابقاً
    if (notifyStatus && notifyStatus !== '-') continue;

    // حد أقصى للتشغيل
    if (processed >= FLIGHT_NOTIFY_MAX_PER_RUN) break;

    // تخطي ليلاً
    if (isNightTime && !FLIGHT_NOTIFY_DRY_RUN) {
      stats.skippedNight++;
      continue;
    }

    stats.total++;
    processed++;

    // استخراج passport من العمود L (index 11)
    var passport = String(row[11] || '').trim();
    if (!passport) {
      updates.push({ rowIndex: rowIndex, status: '⚠️ لا جواز', timestamp: '' });
      stats.error++;
      continue;
    }

    // مطابقة مع chatId
    var entry = passportMap[passport.toUpperCase()];
    if (!entry) {
      updates.push({ rowIndex: rowIndex, status: '⚠️ غير مسجل في البوت', timestamp: '' });
      stats.notRegistered++;
      continue;
    }

    var chatId = entry.chatId;
    var lang = entry.language || 'ar';

    // بناء الرسالة
    var message = buildFlightChangeMessage_(row, lang);
    var pdfLink = String(row[36] || '').trim(); // AK
    var inlineButtons = buildInlineButtons_(pdfLink, lang);

    // Dry-run: كتابة بدون إرسال
    if (FLIGHT_NOTIFY_DRY_RUN) {
      Logger.log('[DRY] → chatId ' + chatId + ' | passport ' + passport + '\n' + message);
      updates.push({
        rowIndex: rowIndex,
        status: '✅ [DRY] chatId:' + chatId,
        timestamp: new Date().toISOString()
      });
      stats.dryRun++;
      continue;
    }

    // إرسال فعلي
    try {
      var res = sendMessageSafe_(chatId, message, inlineButtons);
      if (res && res.success) {
        updates.push({
          rowIndex: rowIndex,
          status: '✅ أُرسل',
          timestamp: new Date().toISOString()
        });
        stats.sent++;
      } else {
        updates.push({
          rowIndex: rowIndex,
          status: '❌ خطأ: ' + (res && res.error ? res.error : 'unknown'),
          timestamp: ''
        });
        stats.error++;
      }
    } catch (e) {
      updates.push({
        rowIndex: rowIndex,
        status: '❌ استثناء: ' + e.message,
        timestamp: ''
      });
      stats.error++;
    }

    // Rate limiting
    Utilities.sleep(100);
  }

  // ٦. كتابة الحالات للشيت دفعة واحدة
  if (updates.length > 0) {
    writeNotifyStatuses_(sheet, updates);
  }

  Logger.log('📊 notifyFlightChanges | إجمالي: ' + stats.total +
             ' | أُرسل: ' + stats.sent +
             ' | Dry: ' + stats.dryRun +
             ' | غير مسجل: ' + stats.notRegistered +
             ' | ليل: ' + stats.skippedNight +
             ' | خطأ: ' + stats.error);
  return stats;
}


/**
 * بناء فهرس passport → { chatId, language } من BotSessions
 */
function buildPassportToChatMap_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('BotSessions');
  if (!sheet) {
    Logger.log('❌ شيت BotSessions غير موجود');
    return {};
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  // A: chatId, B: passport, D: language, E: authStatus
  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var chatId = String(data[i][0] || '').trim();
    var passport = String(data[i][1] || '').trim().toUpperCase();
    var lang = String(data[i][3] || 'ar').trim() || 'ar';
    var authStatus = String(data[i][4] || '').trim();

    if (!chatId || !passport) continue;
    if (authStatus !== 'verified') continue; // نخطر فقط المتحقَّق منهم

    map[passport] = { chatId: chatId, language: lang };
  }
  return map;
}


/**
 * بناء نص الرسالة (صيغة د — هجين بأرقام عربية-هندية)
 * @param row  مصفوفة الصف من V3 (42 عمود)
 * @param lang  اللغة (ar فقط حالياً، الباقي مؤجل)
 */
function buildFlightChangeMessage_(row, lang) {
  // استخراج من الأعمدة
  var ob1Flight = String(row[16] || '').trim(); // Q
  var ob1Route  = String(row[17] || '').trim(); // R  ("من XXX إلى YYY")
  var ob1Date   = String(row[18] || '').trim(); // S
  var ob1Time   = String(row[19] || '').trim(); // T

  var ob2Flight = String(row[20] || '').trim(); // U
  var ob2Route  = String(row[21] || '').trim(); // V
  var ob2Date   = String(row[22] || '').trim(); // W
  var ob2Time   = String(row[23] || '').trim(); // X

  var rt1Flight = String(row[24] || '').trim(); // Y
  var rt1Route  = String(row[25] || '').trim(); // Z
  var rt1Date   = String(row[26] || '').trim(); // AA
  var rt1Time   = String(row[27] || '').trim(); // AB

  var rt2Flight = String(row[28] || '').trim(); // AC
  var rt2Route  = String(row[29] || '').trim(); // AD
  var rt2Date   = String(row[30] || '').trim(); // AE
  var rt2Time   = String(row[31] || '').trim(); // AF

  // بناء النص
  var lines = [];
  lines.push('⚠️ <b>تحديث مهم على رحلتك</b>');
  lines.push('');
  lines.push('السلام عليكم ورحمة الله');
  lines.push('');
  lines.push('نودّ إعلامك بتغيير في تفاصيل رحلتك:');

  // رحلات الذهاب — اعرض إذا أي حقل موجود
  var hasOb1 = ob1Flight || ob1Route || ob1Date;
  var hasOb2 = ob2Flight || ob2Route || ob2Date;
  if (hasOb1) {
    lines.push('');
    lines.push('✈️ <b>الذهاب</b>');
    lines.push(formatLegLines_(ob1Flight, ob1Route, ob1Date, ob1Time));
    if (hasOb2) {
      lines.push('');
      lines.push('🔁 <i>ترانزيت</i>');
      lines.push(formatLegLines_(ob2Flight, ob2Route, ob2Date, ob2Time));
    }
  }

  // رحلات العودة
  var hasRt1 = rt1Flight || rt1Route || rt1Date;
  var hasRt2 = rt2Flight || rt2Route || rt2Date;
  if (hasRt1) {
    lines.push('');
    lines.push('🛬 <b>العودة</b>');
    lines.push(formatLegLines_(rt1Flight, rt1Route, rt1Date, rt1Time));
    if (hasRt2) {
      lines.push('');
      lines.push('🔁 <i>ترانزيت</i>');
      lines.push(formatLegLines_(rt2Flight, rt2Route, rt2Date, rt2Time));
    }
  }

  lines.push('');
  lines.push('👇 <b>للاطلاع الكامل</b>');
  lines.push('١. اضغط «🔄 تحديث البيانات»');
  lines.push('٢. ثم اضغط «✈️ رحلتي»');

  return lines.join('\n');
}


/**
 * تنسيق سطور الرحلة الواحدة — تسميات عربية + قيم تقنية إنجليزية
 */
function formatLegLines_(flight, route, date, time) {
  var lines = [];
  if (flight) lines.push('الرحلة: <code>' + flight + '</code>');
  if (route)  lines.push('المسار: ' + cleanRoute_(route));
  if (date)   lines.push('التاريخ: ' + toArabicDigits_(formatArabicDate_(date)));
  if (time)   lines.push('الإقلاع: ' + toArabicDigits_(formatArabicTime_(time)));
  return lines.join('\n');
}


/**
 * تنظيف صيغة المسار (من XXX إلى YYY) → XXX → YYY
 */
function cleanRoute_(routeText) {
  var m = String(routeText).match(/من\s+([A-Z]{3})\s+إلى\s+([A-Z]{3})/);
  if (m) return m[1] + ' → ' + m[2];
  return routeText;
}


// ملاحظة: دوال toArabicDigits_ و formatArabicDate_ و formatArabicTime_ مُوحَّدة في Helpers.js

/**
 * بناء أزرار inline
 */
function buildInlineButtons_(pdfLink, lang) {
  var buttons = [];
  if (pdfLink && /^https?:\/\//.test(pdfLink)) {
    buttons.push([{ text: '📄 التذكرة الجديدة', url: pdfLink }]);
  }
  buttons.push([{ text: '❓ تواصل معنا', callback_data: 'contact_us' }]);
  return { inline_keyboard: buttons };
}


/**
 * كتابة حالات الإخطار للشيت دفعة واحدة
 */
function writeNotifyStatuses_(sheet, updates) {
  // تجميع المتتاليات لتقليل عمليات الكتابة
  updates.forEach(function(u) {
    sheet.getRange(u.rowIndex, 41).setValue(u.status);     // AO
    if (u.timestamp) {
      sheet.getRange(u.rowIndex, 42).setValue(u.timestamp); // AP
    }
  });
}


/**
 * قراءة رؤوس شيت الطيران الأصلي — لمعرفة الهيكل
 */
function debugFlightSheetHeaders() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(FLIGHTS_SHEET);
  if (!sheet) return { error: 'not_found', name: FLIGHTS_SHEET };
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var firstRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    var letter = i < 26 ? String.fromCharCode(65 + i) : 'A' + String.fromCharCode(65 + i - 26);
    out.push({ index: i, column: letter, header: headers[i], sample: firstRow[i] });
  }
  return { count: headers.length, columns: out };
}


/**
 * قراءة رؤوس V3 + أول صف — لاكتشاف أعمدة الشيت الفعلية
 */
function debugV3Headers() {
  var ss = SpreadsheetApp.openById(FLIGHT_CHANGES_SHEET_ID);
  var sheet = ss.getSheetByName(FLIGHT_CHANGES_SHEET_NAME);
  if (!sheet) return { error: 'sheet_not_found' };
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var firstRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var out = [];
  for (var i = 0; i < headers.length; i++) {
    out.push({ index: i, column: String.fromCharCode(65 + (i < 26 ? i : 0)) + (i >= 26 ? String.fromCharCode(65 + i - 26) : ''), header: headers[i], sample: firstRow[i] });
  }
  return { count: headers.length, columns: out };
}


/**
 * معاينة رسائل الإخطار — يُرجع نماذج فعلية بدون أي إرسال
 * @param count  عدد النماذج (افتراضي 3)
 */
function previewFlightMessages(count) {
  count = count || 3;
  var ss = SpreadsheetApp.openById(FLIGHT_CHANGES_SHEET_ID);
  var sheet = ss.getSheetByName(FLIGHT_CHANGES_SHEET_NAME);
  if (!sheet) return { error: 'sheet_not_found' };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { error: 'empty' };

  var data = sheet.getRange(2, 1, lastRow - 1, 42).getValues();
  var passportMap = buildPassportToChatMap_();
  var samples = [];

  for (var i = 0; i < data.length && samples.length < count; i++) {
    var row = data[i];
    var passport = String(row[11] || '').trim();
    var pnr = String(row[1] || '').trim();
    var name = String(row[5] || '').trim() + ' ' + String(row[6] || '').trim();
    var chgNum = String(row[0] || '').trim();

    if (!passport) continue;
    var entry = passportMap[passport.toUpperCase()];

    var msg = buildFlightChangeMessage_(row, entry ? entry.language : 'ar');
    var pdfLink = String(row[36] || '').trim();

    samples.push({
      chg: chgNum,
      pnr: pnr,
      name: name,
      passport: passport,
      registered: !!entry,
      chatId: entry ? entry.chatId : null,
      message: msg,
      pdfLink: pdfLink
    });
  }

  return { count: samples.length, samples: samples };
}


/**
 * دالة اختبار يدوية سريعة — يرى الـ logs فقط
 */
function testNotifyDryRun() {
  // تأكيد وضع التجربة
  Logger.log('FLIGHT_NOTIFY_ENABLED = ' + FLIGHT_NOTIFY_ENABLED);
  Logger.log('FLIGHT_NOTIFY_DRY_RUN = ' + FLIGHT_NOTIFY_DRY_RUN);
  if (!FLIGHT_NOTIFY_ENABLED) {
    Logger.log('⚠️ FLIGHT_NOTIFY_ENABLED = false — غيّره لـ true قبل الاختبار');
    return;
  }
  if (!FLIGHT_NOTIFY_DRY_RUN) {
    Logger.log('⚠️ FLIGHT_NOTIFY_DRY_RUN = false — سيُرسل فعلياً!');
    return;
  }
  return notifyFlightChanges();
}
