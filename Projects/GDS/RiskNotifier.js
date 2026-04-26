/**
 * RiskNotifier.js — تنسيق ملاحظات المخاطر من المخّ للعرض
 *
 * يتلقى risks[] من Claude + warnings من TicketValidator
 * ويُرجع تقريراً منظماً جاهزاً للإرسال عبر تيليغرام في Phase 6.
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.RiskNotifier = {
  /**
   * تلخيص المخاطر + التحذيرات في شكل منظم.
   * @param {Array} risks - من parsed_data.risks
   * @param {Array} validationWarnings - من TicketValidator.warnings
   * @param {Object} pilgrim - { name, passport }
   * @return {Object} تقرير منظم أو null إذا لا شيء للإبلاغ
   */
  format: function(risks, validationWarnings, pilgrim) {
    var hasRisks = risks && risks.length > 0;
    var hasWarnings = validationWarnings && validationWarnings.length > 0;
    if (!hasRisks && !hasWarnings) return null;

    // تصنيف risks حسب المستوى
    var critical = [];
    var warnings = [];
    var info = [];

    if (hasRisks) {
      for (var i = 0; i < risks.length; i++) {
        var r = risks[i];
        var level = (r.level || 'info').toLowerCase();
        if (level === 'critical') critical.push(r);
        else if (level === 'warn' || level === 'warning') warnings.push(r);
        else info.push(r);
      }
    }

    return {
      pilgrim: pilgrim,
      summary: {
        critical_count: critical.length,
        warn_count: warnings.length,
        info_count: info.length,
        validation_warnings_count: hasWarnings ? validationWarnings.length : 0
      },
      critical: critical,
      warnings: warnings,
      info: info,
      validation_warnings: validationWarnings || []
    };
  },

  /**
   * تحويل التقرير لرسالة نصية جاهزة لتيليغرام.
   * @param {Object} report - ناتج format()
   * @return {string}
   */
  toText: function(report) {
    if (!report) return '';

    var lines = [];
    var p = report.pilgrim || {};
    lines.push('🧳 الحاج: ' + (p.name || 'غير محدد') + ' (' + (p.passport || '—') + ')');
    lines.push('');

    if (report.critical.length > 0) {
      lines.push('🚨 حرج:');
      for (var i = 0; i < report.critical.length; i++) {
        lines.push('  • ' + report.critical[i].message);
      }
      lines.push('');
    }

    if (report.warnings.length > 0) {
      lines.push('⚠️ تحذيرات:');
      for (var j = 0; j < report.warnings.length; j++) {
        lines.push('  • ' + report.warnings[j].message);
      }
      lines.push('');
    }

    if (report.info.length > 0) {
      lines.push('ℹ️ ملاحظات:');
      for (var k = 0; k < report.info.length; k++) {
        lines.push('  • ' + report.info[k].message);
      }
      lines.push('');
    }

    if (report.validation_warnings.length > 0) {
      lines.push('🔧 تطبيعات تلقائية:');
      for (var v = 0; v < report.validation_warnings.length; v++) {
        var w = report.validation_warnings[v];
        if (w.layer === 'flightNo') {
          lines.push('  • رقم رحلة: "' + w.from + '" → "' + w.to + '"');
        } else if (w.layer === 'time') {
          lines.push('  • وقت ' + w.segment + '.' + w.field + ': "' + w.from + '" → "' + w.to + '"');
        } else if (w.layer === 'iata' && w.action === 'auto_added_to_registry') {
          lines.push('  • مطار جديد أُضيف للسجل: ' + w.code);
        }
      }
    }

    return lines.join('\n').trim();
  }
};

// ==================== Global Entry Points ====================

function formatRisks(risks, warnings, pilgrim) {
  return GDS2.RiskNotifier.format(risks, warnings, pilgrim);
}
