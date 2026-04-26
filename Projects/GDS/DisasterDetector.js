/**
 * DisasterDetector.js — كشف "الكارثة": رابط يشير لتذكرة شخص آخر
 *
 * إذا Claude أرجع found=false، فالحاج ليس في التذكرة المُشار إليها بـ ticketUrl
 * → هذا يعني أن الرابط في PD خاطئ (أُدخل خطأً، أو تبدّل مع حاج آخر)
 * → إشعار تيليغرام فوري للعمل (Phase 6)
 * → لا كتابة، لا تحديث BJ، لا زيادة BL
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.DisasterDetector = {
  /**
   * فحص نتيجة Claude للكارثة.
   * @param {Object} parsedData - الـ JSON المُرجَع
   * @param {Object} pilgrim - { name, passport, url }
   * @return {Object} { is_disaster, type?, details? }
   */
  check: function(parsedData, pilgrim) {
    if (!parsedData) {
      return { is_disaster: false };
    }

    if (parsedData.found === false) {
      return {
        is_disaster: true,
        type: 'pilgrim_not_in_ticket',
        severity: 'critical',
        pilgrim: pilgrim,
        passengers_in_ticket: parsedData.all_passengers || [],
        message: 'الحاج ' + (pilgrim.name || pilgrim.passport) + ' غير موجود في التذكرة المُشار إليها. الرابط قد يكون لشخص آخر.'
      };
    }

    // تحقق أيضاً: إذا matched_name موجود، هل يختلف جذرياً عن اسم الحاج؟
    // (فحص إضافي للثقة — Claude قد يُرجع found=true بخطأ)
    // نتركها بسيطة الآن — نثق بـ Claude

    return { is_disaster: false };
  }
};

// ==================== Global Entry Points ====================

function checkDisaster(parsedData, pilgrim) {
  return GDS2.DisasterDetector.check(parsedData, pilgrim);
}
