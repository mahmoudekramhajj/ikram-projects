/**
 * FlightValidator.js — التحقق المنطقي من بيانات الرحلات
 *
 * يُستدعى قبل كتابة أي صف في شيت V3.
 * يُعيد قائمة تحذيرات بصفّ الحاج (لا يرفض الصف — الحاج يُكتب دائماً).
 */


/**
 * الدالة الرئيسية — تتحقق من كل رحلات الحاج وتُعيد تحذيرات
 * @param outLegs  array of { flightNumber, fromCity, toCity, depDate, depTime, arrDate, arrTime }
 * @param retLegs  نفس الهيكل للعودة
 * @return string — تحذيرات مفصولة بـ " | " (فارغ إذا كل شيء منطقي)
 */
function validateFlightLegs_(outLegs, retLegs) {
  var warnings = [];

  // فحص رحلات الذهاب
  var outW = validateLegGroup_(outLegs || [], 'الذهاب');
  for (var i = 0; i < outW.length; i++) warnings.push(outW[i]);

  // فحص رحلات العودة
  var retW = validateLegGroup_(retLegs || [], 'العودة');
  for (var j = 0; j < retW.length; j++) warnings.push(retW[j]);

  // فحص العودة بعد الذهاب
  if (outLegs && outLegs.length > 0 && retLegs && retLegs.length > 0) {
    var lastOut = outLegs[outLegs.length - 1];
    var firstRet = retLegs[0];
    var cmp = compareDateTime_(lastOut, firstRet, 'arr', 'dep');
    if (cmp === 'invalid') {
      // لا يمكن مقارنة — لا تحذير
    } else if (cmp > 0) {
      warnings.push('⚠️ رحلة العودة تنطلق قبل وصول الذهاب');
    }
  }

  return warnings.join(' | ');
}


/**
 * فحص مجموعة رحلات (ذهاب أو عودة)
 */
function validateLegGroup_(legs, groupName) {
  var warnings = [];

  for (var i = 0; i < legs.length; i++) {
    var leg = legs[i];
    if (!leg) continue;

    var legLabel = groupName + ' ' + (i + 1);

    // قاعدة ١: نفس المطار (AMM→AMM)
    if (leg.fromCity && leg.toCity && leg.fromCity === leg.toCity) {
      warnings.push('⚠️ ' + legLabel + ': نفس المطار (' + leg.fromCity + '→' + leg.toCity + ')');
    }

    // قاعدة ٢: رقم رحلة مشبوه
    if (leg.flightNumber) {
      var fn = String(leg.flightNumber).trim();
      // يجب أن يبدأ بحرفين كود طيران ثم رقم ≥1
      if (!/^[A-Z]{2}-?\d{1,5}$/i.test(fn)) {
        warnings.push('⚠️ ' + legLabel + ': رقم رحلة غير صالح (' + fn + ')');
      }
    }

    // قاعدة ٣: مطار غير معروف
    if (leg.fromCity && !isValidAirport_(leg.fromCity)) {
      warnings.push('⚠️ ' + legLabel + ': مطار مغادرة غير معروف (' + leg.fromCity + ')');
    }
    if (leg.toCity && !isValidAirport_(leg.toCity)) {
      warnings.push('⚠️ ' + legLabel + ': مطار وصول غير معروف (' + leg.toCity + ')');
    }
  }

  // قاعدة ٤: استمرارية الترانزيت (مطار وصول = مطار مغادرة التالي)
  // استثناء: MED↔JED طبيعي — الحاج ينتقل براً بين المدينة ومكة (جدة)
  var SAUDI_TRANSFER_AIRPORTS = { 'JED': 1, 'MED': 1 };
  for (var k = 0; k < legs.length - 1; k++) {
    var cur = legs[k];
    var next = legs[k + 1];
    if (cur && next && cur.toCity && next.fromCity && cur.toCity !== next.fromCity) {
      // استثناء: كلاهما في السعودية (JED أو MED) = انتقال بري طبيعي
      if (SAUDI_TRANSFER_AIRPORTS[cur.toCity] && SAUDI_TRANSFER_AIRPORTS[next.fromCity]) {
        continue; // لا تحذير
      }
      warnings.push('⚠️ ' + groupName + ': فجوة في الترانزيت (' + cur.toCity + ' → ' + next.fromCity + ')');
    }
  }

  // قاعدة ٥: التسلسل الزمني (رحلة 2 تنطلق بعد وصول رحلة 1)
  for (var m = 0; m < legs.length - 1; m++) {
    var leg1 = legs[m];
    var leg2 = legs[m + 1];
    var cmp = compareDateTime_(leg1, leg2, 'arr', 'dep');
    if (cmp === 'invalid') continue; // لا يمكن تحليل — نتخطى
    if (cmp > 0) {
      warnings.push('⚠️ ' + groupName + ': الرحلة ' + (m + 2) + ' تنطلق قبل وصول الرحلة ' + (m + 1));
    }
  }

  return warnings;
}


/**
 * قائمة مبسّطة بأهم المطارات في رحلات الحج
 * مرجع IATA — نضيف للقائمة بحسب الحاجة
 */
var VALID_AIRPORTS_ = {
  // السعودية
  'JED': 1, 'MED': 1, 'RUH': 1, 'DMM': 1, 'TIF': 1, 'YNB': 1,
  // الخليج والشرق الأوسط
  'DXB': 1, 'AUH': 1, 'SHJ': 1, 'DOH': 1, 'KWI': 1, 'BAH': 1, 'MCT': 1,
  'AMM': 1, 'BEY': 1, 'DAM': 1, 'BGW': 1, 'CAI': 1, 'ALY': 1,
  // شمال أفريقيا
  'TUN': 1, 'ALG': 1, 'CMN': 1, 'RAK': 1, 'AAE': 1, 'ORN': 1, 'MIR': 1,
  // تركيا
  'IST': 1, 'SAW': 1, 'ESB': 1, 'ADB': 1, 'AYT': 1,
  // إيران
  'IKA': 1, 'THR': 1, 'MHD': 1,
  // أوروبا (الشائعة)
  'LHR': 1, 'LGW': 1, 'STN': 1, 'LTN': 1, 'MAN': 1, 'BHX': 1,
  'CDG': 1, 'ORY': 1, 'FRA': 1, 'MUC': 1, 'AMS': 1, 'BRU': 1, 'VIE': 1,
  'MAD': 1, 'BCN': 1, 'MXP': 1, 'FCO': 1, 'ZRH': 1, 'GVA': 1, 'CGN': 1,
  'DUS': 1, 'HAM': 1, 'BER': 1, 'CPH': 1, 'ARN': 1, 'OSL': 1, 'HEL': 1,
  // آسيا
  'KUL': 1, 'SIN': 1, 'BKK': 1, 'JKT': 1, 'CGK': 1, 'SUB': 1, 'MES': 1,
  'DAC': 1, 'KHI': 1, 'ISB': 1, 'LHE': 1, 'CCU': 1, 'DEL': 1, 'BOM': 1,
  'MAA': 1, 'HYD': 1, 'BLR': 1, 'COK': 1, 'TRV': 1, 'CMB': 1,
  'KBL': 1, 'KHI': 1, 'TAS': 1, 'ALA': 1,
  // أفريقيا
  'CPT': 1, 'JNB': 1, 'DUR': 1, 'LOS': 1, 'ABV': 1, 'KAN': 1, 'ACC': 1,
  'ADD': 1, 'NBO': 1, 'DAR': 1, 'EBB': 1, 'KGL': 1, 'LAD': 1,
  // أمريكا
  'JFK': 1, 'EWR': 1, 'IAD': 1, 'ORD': 1, 'LAX': 1, 'SFO': 1, 'YYZ': 1, 'YUL': 1,
  'DFW': 1, 'SEA': 1, 'MIA': 1, 'DTW': 1, 'IAH': 1, 'ATL': 1, 'BOS': 1, 'PHL': 1,
  'DEN': 1, 'PHX': 1, 'LAS': 1, 'MSP': 1, 'CLT': 1, 'MCO': 1, 'FLL': 1,
  // أمريكا الوسطى/الجنوبية
  'PTY': 1, 'MEX': 1, 'GRU': 1, 'EZE': 1, 'BOG': 1, 'LIM': 1, 'SCL': 1,
  // أوروبا إضافية
  'ATH': 1, 'OTP': 1, 'BSL': 1, 'LYS': 1, 'NCE': 1, 'MRS': 1, 'BUD': 1, 'PRG': 1,
  'WAW': 1, 'SOF': 1, 'BEG': 1, 'SKP': 1, 'SJJ': 1, 'ZAG': 1, 'LUX': 1, 'KEF': 1,
  // شرق أوسط/عراق إضافية
  'ISU': 1, 'EBL': 1, 'NJF': 1, 'BSR': 1, 'KRT': 1, 'SAH': 1, 'ADE': 1,
  // شرق آسيا
  'HKG': 1, 'PEK': 1, 'PVG': 1, 'ICN': 1, 'NRT': 1, 'HND': 1,
  // قوقاز
  'GYD': 1, 'TBS': 1, 'EVN': 1,
  // أخرى
  'SAN': 1, 'MSQ': 1, 'DME': 1, 'SVO': 1, 'VKO': 1, 'LED': 1
};

function isValidAirport_(code) {
  if (!code) return true; // خال = لا تحذير
  var c = String(code).trim().toUpperCase();
  if (c.length !== 3) return false; // IATA = 3 حروف
  if (!/^[A-Z]{3}$/.test(c)) return false;
  return !!VALID_AIRPORTS_[c];
}


/**
 * مقارنة تاريخ/وقت رحلتين
 * @param leg1, leg2  كائنات leg
 * @param field1  'dep' أو 'arr'
 * @param field2  'dep' أو 'arr'
 * @return عدد (<0 إذا leg1 قبل leg2)، 0 مساوٍ، >0 leg1 بعد leg2، 'invalid' إذا لم يمكن التحليل
 */
function compareDateTime_(leg1, leg2, field1, field2) {
  var dt1 = parseLegDateTime_(leg1, field1);
  var dt2 = parseLegDateTime_(leg2, field2);
  if (!dt1 || !dt2) return 'invalid';
  return dt1.getTime() - dt2.getTime();
}


/**
 * يُحلّل date + time من كائن leg إلى Date
 * أمثلة على التنسيقات:
 *   depDate: "May 20, 2026" أو "2026-05-20"
 *   depTime: "09:10 PM" أو "21:10"
 */
function parseLegDateTime_(leg, field) {
  if (!leg) return null;
  var dateStr = String(leg[field + 'Date'] || '').trim();
  var timeStr = String(leg[field + 'Time'] || '').trim();
  if (!dateStr) return null;

  // تحليل التاريخ
  var date;
  try {
    date = new Date(dateStr);
    if (isNaN(date.getTime())) return null;
  } catch (e) {
    return null;
  }

  // تحليل الوقت (اختياري)
  if (timeStr) {
    var tm = timeStr.match(/(\d{1,2}):(\d{2})(?:\s*(AM|PM))?/i);
    if (tm) {
      var hour = parseInt(tm[1], 10);
      var minute = parseInt(tm[2], 10);
      var ampm = tm[3] ? tm[3].toUpperCase() : '';
      if (ampm === 'PM' && hour < 12) hour += 12;
      if (ampm === 'AM' && hour === 12) hour = 0;
      date.setHours(hour, minute, 0, 0);
    }
  }

  return date;
}
