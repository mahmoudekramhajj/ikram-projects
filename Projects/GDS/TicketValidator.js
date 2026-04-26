/**
 * TicketValidator.js — 5 طبقات تحقق + تطبيع
 *
 * الطبقات:
 *   1. Structure — وجود الحقول الإلزامية (pnr, dep2, ret1)
 *   2. IATA — صيغة ثلاثة أحرف + إضافة تلقائية للجديد
 *   3. Date — ضمن موسم الحج (2026-04-01 إلى 2026-07-31) + arrival >= takeoff
 *   4. Time — تطبيع لـ HH:MM (تسامح مع PM/AM، نقطة، أرقام عربية)
 *   5. Flight number — تطبيع (حذف شرطات/مسافات + إضافة رمز الشركة إن ناقص)
 *   6. Logic — DEP2.to ∈ {JED, MED}، RET1.from ∈ {JED, MED}، RET1 > DEP2
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.TicketValidator = {
  /**
   * @param {Object} parsed - JSON المُرجَع من Claude
   * @return {Object} { valid, errors, warnings, normalized, is_disaster, too_many_segments }
   */
  validate: function(parsed) {
    if (!parsed || typeof parsed !== 'object') {
      return { valid: false, errors: [{ layer: 'structure', reason: 'null_or_invalid_data' }] };
    }

    // فحص "الكارثة": الحاج غير موجود في التذكرة
    if (parsed.found === false) {
      return {
        valid: false,
        is_disaster: true,
        passengers_in_ticket: parsed.all_passengers || [],
        errors: [{ layer: 'lookup', reason: 'pilgrim_not_in_ticket' }]
      };
    }

    // فحص too_many_segments — صارم: نرفض إذا Claude يقول too_many
    // (تجنّب قبول بيانات ناقصة/معطوبة عند عدم دمج Technical Stop)
    if (parsed.too_many_segments === true) {
      return {
        valid: false,
        too_many_segments: true,
        errors: [{ layer: 'structure', reason: 'too_many_segments' }]
      };
    }

    // Deep clone للتطبيع بدون تعديل الأصل
    var normalized = JSON.parse(JSON.stringify(parsed));
    var errors = [];
    var warnings = [];

    // Layer 1: Structure
    errors = errors.concat(GDS2.TicketValidator._validateStructure(normalized));
    if (errors.length > 0) {
      return { valid: false, errors: errors };
    }

    // تطبيع + تحقق لكل segment (حتى 6: dep0, dep1, dep2, ret1, ret2, ret3)
    var segKeys = ['dep0', 'dep1', 'dep2', 'ret1', 'ret2', 'ret3'];
    for (var i = 0; i < segKeys.length; i++) {
      var key = segKeys[i];
      var seg = normalized.segments[key];
      if (!seg) continue;

      // تطبيع الوقت
      var depTimeResult = GDS2.TicketValidator._normalizeTime(seg.depTime);
      var arrTimeResult = GDS2.TicketValidator._normalizeTime(seg.arrTime);
      if (depTimeResult.changed) warnings.push({ layer: 'time', segment: key, field: 'depTime', from: seg.depTime, to: depTimeResult.value });
      if (arrTimeResult.changed) warnings.push({ layer: 'time', segment: key, field: 'arrTime', from: seg.arrTime, to: arrTimeResult.value });
      seg.depTime = depTimeResult.value;
      seg.arrTime = arrTimeResult.value;

      // تطبيع رقم الرحلة
      var flightResult = GDS2.TicketValidator._normalizeFlightNumber(seg.flightNo, normalized.airline_code);
      if (flightResult.changed) warnings.push({ layer: 'flightNo', segment: key, from: seg.flightNo, to: flightResult.value });
      seg.flightNo = flightResult.value;

      // تطبيع IATA
      seg.from = String(seg.from || '').toUpperCase().trim();
      seg.to = String(seg.to || '').toUpperCase().trim();

      // تحقق Time format
      errors = errors.concat(GDS2.TicketValidator._validateTimes(key, seg));

      // تحقق Date
      errors = errors.concat(GDS2.TicketValidator._validateDates(key, seg));

      // تحقق IATA (مع إضافة تلقائية)
      var iataResult = GDS2.TicketValidator._validateIATA(key, seg);
      errors = errors.concat(iataResult.errors);
      warnings = warnings.concat(iataResult.warnings);
    }

    // Layer 6: Logic
    errors = errors.concat(GDS2.TicketValidator._validateLogic(normalized));

    return {
      valid: errors.length === 0,
      errors: errors,
      warnings: warnings,
      normalized: normalized
    };
  },

  // ==================== Private ====================

  _validateStructure: function(data) {
    var errors = [];
    if (!data.pnr) errors.push({ layer: 'structure', reason: 'missing_pnr' });
    if (!data.segments) {
      errors.push({ layer: 'structure', reason: 'missing_segments' });
      return errors;
    }
    if (!data.segments.dep2) errors.push({ layer: 'structure', reason: 'missing_dep2' });
    if (!data.segments.ret1) errors.push({ layer: 'structure', reason: 'missing_ret1' });
    return errors;
  },

  _normalizeTime: function(timeStr) {
    if (!timeStr) return { value: '', changed: false };
    var original = String(timeStr);
    var s = original.trim();

    // تحويل الأرقام العربية
    var arabicMap = { '٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9' };
    s = s.replace(/[٠-٩]/g, function(d) { return arabicMap[d]; });

    // PM/AM handling
    var ampmMatch = s.match(/^(\d{1,2})[\:\.]?(\d{2})\s*(AM|PM|am|pm)/);
    if (ampmMatch) {
      var h = parseInt(ampmMatch[1], 10);
      var m = parseInt(ampmMatch[2], 10);
      var isPM = /^[Pp]/.test(ampmMatch[3]);
      if (isPM && h < 12) h += 12;
      if (!isPM && h === 12) h = 0;
      var result = (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
      return { value: result, changed: result !== original };
    }

    // HH.MM → HH:MM
    s = s.replace('.', ':');

    // HHMM → HH:MM
    var hhmmMatch = s.match(/^(\d{2})(\d{2})$/);
    if (hhmmMatch) {
      var result2 = hhmmMatch[1] + ':' + hhmmMatch[2];
      return { value: result2, changed: result2 !== original };
    }

    // Standard HH:MM أو H:MM
    var stdMatch = s.match(/^(\d{1,2}):(\d{2})/);
    if (stdMatch) {
      var hr = parseInt(stdMatch[1], 10);
      var mn = parseInt(stdMatch[2], 10);
      var result3 = (hr < 10 ? '0' : '') + hr + ':' + (mn < 10 ? '0' : '') + mn;
      return { value: result3, changed: result3 !== original };
    }

    return { value: s, changed: s !== original };
  },

  _normalizeFlightNumber: function(flightNo, airlineCode) {
    if (!flightNo) return { value: '', changed: false };
    var original = String(flightNo);
    var cleaned = original.toUpperCase().trim()
      .replace(/\s+/g, '')
      .replace(/-/g, '');

    // إذا كل الرقم أرقام فقط، أضف رمز الشركة
    if (/^\d+$/.test(cleaned) && airlineCode) {
      cleaned = String(airlineCode).toUpperCase() + cleaned;
    }

    return { value: cleaned, changed: cleaned !== original };
  },

  _validateTimes: function(segKey, seg) {
    var errors = [];
    var pattern = /^([01]\d|2[0-3]):[0-5]\d$/;
    if (seg.depTime && !pattern.test(seg.depTime)) {
      errors.push({ layer: 'time', segment: segKey, field: 'depTime', value: seg.depTime, reason: 'bad_format' });
    }
    if (seg.arrTime && !pattern.test(seg.arrTime)) {
      errors.push({ layer: 'time', segment: segKey, field: 'arrTime', value: seg.arrTime, reason: 'bad_format' });
    }
    return errors;
  },

  _validateDates: function(segKey, seg) {
    var errors = [];
    var start = GDS2.Config.SEASON_START;
    var end = GDS2.Config.SEASON_END;

    if (seg.depDate) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(seg.depDate)) {
        errors.push({ layer: 'date', segment: segKey, field: 'depDate', value: seg.depDate, reason: 'bad_format' });
      } else if (seg.depDate < start || seg.depDate > end) {
        errors.push({ layer: 'date', segment: segKey, field: 'depDate', value: seg.depDate, reason: 'out_of_season' });
      }
    }

    if (seg.arrDate) {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(seg.arrDate)) {
        errors.push({ layer: 'date', segment: segKey, field: 'arrDate', value: seg.arrDate, reason: 'bad_format' });
      } else if (seg.arrDate < start || seg.arrDate > end) {
        errors.push({ layer: 'date', segment: segKey, field: 'arrDate', value: seg.arrDate, reason: 'out_of_season' });
      }
    }

    // Arrival >= Departure
    if (seg.depDate && seg.arrDate && seg.arrDate < seg.depDate) {
      errors.push({
        layer: 'date', segment: segKey, reason: 'arrival_before_takeoff',
        takeoff: seg.depDate, arrival: seg.arrDate
      });
    }

    return errors;
  },

  _validateIATA: function(segKey, seg) {
    var errors = [];
    var warnings = [];
    var fields = ['from', 'to'];

    for (var i = 0; i < fields.length; i++) {
      var code = seg[fields[i]];
      if (!code) continue;

      if (!/^[A-Z]{3}$/.test(code)) {
        errors.push({ layer: 'iata', segment: segKey, field: fields[i], value: code, reason: 'bad_format' });
        continue;
      }

      // إضافة تلقائية لأي كود جديد
      var added = GDS2.IATARegistry.add(code);
      if (added) {
        warnings.push({ layer: 'iata', segment: segKey, field: fields[i], code: code, action: 'auto_added_to_registry' });
      }
    }

    return { errors: errors, warnings: warnings };
  },

  _validateLogic: function(data) {
    var errors = [];
    if (!data.segments) return errors;
    var segs = data.segments;
    var saudi = GDS2.Config.SAUDI_ARRIVAL_AIRPORTS;

    // DEP2.to ∈ {JED, MED}
    if (segs.dep2 && segs.dep2.to && saudi.indexOf(segs.dep2.to) === -1) {
      errors.push({ layer: 'logic', reason: 'dep2_not_to_saudi', value: segs.dep2.to });
    }

    // RET1.from ∈ {JED, MED}
    if (segs.ret1 && segs.ret1.from && saudi.indexOf(segs.ret1.from) === -1) {
      errors.push({ layer: 'logic', reason: 'ret1_not_from_saudi', value: segs.ret1.from });
    }

    // RET1 بعد DEP2
    if (segs.dep2 && segs.ret1 && segs.dep2.arrDate && segs.ret1.depDate) {
      if (segs.ret1.depDate < segs.dep2.arrDate) {
        errors.push({
          layer: 'logic', reason: 'return_before_departure',
          dep2_arrival: segs.dep2.arrDate, ret1_takeoff: segs.ret1.depDate
        });
      }
    }

    return errors;
  }
};

// ==================== Global Entry Points ====================

function validateTicket(parsedJson) {
  return GDS2.TicketValidator.validate(parsedJson);
}
