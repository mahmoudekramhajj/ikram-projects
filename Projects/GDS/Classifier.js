/**
 * Classifier.js — تحويل parsed data إلى مصفوفات جاهزة للكتابة
 *
 * المخرجات:
 *   1. AG..BI (29 عمود): DEP1 (7) + DEP2 (7) + RET1 (7) + RET2 (7) + PNR (1)
 *   2. BS..BY (7 أعمدة): DEP0 — قطعة الذهاب الإضافية (3+ قطع)
 *   3. BZ..CF (7 أعمدة): RET3 — قطعة العودة الإضافية (3+ قطع)
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.Classifier = {
  /**
   * بناء مصفوفة AG-BI (29 عمود) من البيانات المُطبَّعة.
   * تشمل DEP1+DEP2+RET1+RET2+PNR (بدون DEP0/RET3).
   * @param {Object} normalizedData - من TicketValidator.validate().normalized
   * @return {Array} 29 قيمة
   */
  buildFlightArray: function(normalizedData) {
    if (!normalizedData || !normalizedData.segments) {
      return GDS2.Classifier._emptyArray();
    }
    var segs = normalizedData.segments;
    var flat = [];
    flat = flat.concat(GDS2.Classifier._segmentToArray(segs.dep1));
    flat = flat.concat(GDS2.Classifier._segmentToArray(segs.dep2));
    flat = flat.concat(GDS2.Classifier._segmentToArray(segs.ret1));
    flat = flat.concat(GDS2.Classifier._segmentToArray(segs.ret2));
    flat.push(normalizedData.pnr || '');
    return flat;
  },

  /**
   * بناء مصفوفة DEP0 (7 أعمدة) — قطعة ذهاب إضافية للحالات 3+ قطع.
   * @return {Array} 7 قيم (فارغة إن لم توجد قطعة 3)
   */
  buildDep0Array: function(normalizedData) {
    if (!normalizedData || !normalizedData.segments) {
      return GDS2.Classifier._emptySegmentArray();
    }
    return GDS2.Classifier._segmentToArray(normalizedData.segments.dep0);
  },

  /**
   * بناء مصفوفة RET3 (7 أعمدة) — قطعة عودة إضافية للحالات 3+ قطع.
   * @return {Array} 7 قيم (فارغة إن لم توجد قطعة 3)
   */
  buildRet3Array: function(normalizedData) {
    if (!normalizedData || !normalizedData.segments) {
      return GDS2.Classifier._emptySegmentArray();
    }
    return GDS2.Classifier._segmentToArray(normalizedData.segments.ret3);
  },

  _segmentToArray: function(seg) {
    if (!seg) return ['', '', '', '', '', '', ''];
    return [
      seg.flightNo || '',
      seg.depDate || '',
      seg.depTime || '',
      seg.from || '',
      seg.to || '',
      seg.arrDate || '',
      seg.arrTime || ''
    ];
  },

  _emptySegmentArray: function() {
    return ['', '', '', '', '', '', ''];
  },

  _emptyArray: function() {
    var arr = [];
    for (var i = 0; i < 29; i++) arr.push('');
    return arr;
  }
};
