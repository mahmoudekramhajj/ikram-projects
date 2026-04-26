/**
 * AirlineRegistry.js — سجل شركات الطيران ورموزها IATA
 *
 * القائمة الأساسية لأكثر الشركات استخداماً للحجاج، قابلة للتوسعة.
 * المطابقة تدعم عدة أشكال: العربي، الإنجليزي، الاختصارات.
 *
 * دوال عامة:
 *   initAirlineRegistry()          — تهيئة أولى بالقائمة الافتراضية
 *   getAirlineList()               — القائمة الحالية
 *   getAirlineCode(name)           — البحث عن الكود بالاسم
 *   addAirline(name, code, aliases) — إضافة شركة جديدة
 */

var GDS2 = (typeof GDS2 !== 'undefined') ? GDS2 : {};

GDS2.AirlineRegistry = {
  /**
   * القائمة الافتراضية — أكثر الشركات استخداماً للحجاج.
   */
  DEFAULT: [
    // الشركات الخليجية
    { code: 'SV', name: 'الخطوط السعودية', aliases: ['Saudia', 'Saudi Arabian Airlines', 'Saudi Arabian'] },
    { code: 'XY', name: 'طيران ناس', aliases: ['Flynas', 'Nas Air'] },
    { code: 'QR', name: 'القطرية', aliases: ['Qatar Airways', 'Qatar'] },
    { code: 'EK', name: 'طيران الإمارات', aliases: ['Emirates', 'Emirates Airlines'] },
    { code: 'EY', name: 'الاتحاد', aliases: ['Etihad', 'Etihad Airways'] },
    { code: 'GF', name: 'الخليج', aliases: ['Gulf Air'] },
    { code: 'KU', name: 'الكويتية', aliases: ['Kuwait Airways'] },
    { code: 'FZ', name: 'فلاي دبي', aliases: ['flydubai'] },
    { code: 'WY', name: 'الطيران العُماني', aliases: ['Oman Air'] },
    { code: 'J9', name: 'الجزيرة', aliases: ['Jazeera Airways'] },
    { code: 'BH', name: 'الطيران البحريني', aliases: ['Bahrain Air'] },

    // الشركات العربية والشرق الأوسطية
    { code: 'MS', name: 'مصر للطيران', aliases: ['EgyptAir', 'Egypt Air'] },
    { code: 'RJ', name: 'الملكية الأردنية', aliases: ['Royal Jordanian'] },
    { code: 'ME', name: 'طيران الشرق الأوسط', aliases: ['Middle East Airlines', 'MEA'] },
    { code: 'TU', name: 'الخطوط التونسية', aliases: ['Tunisair'] },
    { code: 'AT', name: 'الملكية المغربية', aliases: ['Royal Air Maroc', 'RAM'] },
    { code: 'AH', name: 'الخطوط الجزائرية', aliases: ['Air Algerie'] },
    { code: 'SD', name: 'الخطوط السودانية', aliases: ['Sudan Airways'] },
    { code: 'IY', name: 'اليمنية', aliases: ['Yemenia'] },
    { code: 'IA', name: 'الخطوط العراقية', aliases: ['Iraqi Airways'] },
    { code: 'RB', name: 'السورية', aliases: ['Syrian Air'] },

    // التركية
    { code: 'TK', name: 'الخطوط التركية', aliases: ['Turkish Airlines', 'Turkish'] },
    { code: 'XQ', name: 'صن إكسبريس', aliases: ['SunExpress'] },
    { code: 'PC', name: 'بيغاسوس', aliases: ['Pegasus'] },

    // الأوروبية
    { code: 'BA', name: 'البريطانية', aliases: ['British Airways', 'British'] },
    { code: 'AF', name: 'إير فرانس', aliases: ['Air France'] },
    { code: 'LH', name: 'لوفتهانزا', aliases: ['Lufthansa'] },
    { code: 'KL', name: 'KLM', aliases: ['KLM', 'Royal Dutch Airlines'] },
    { code: 'LX', name: 'السويسرية', aliases: ['Swiss International Air Lines', 'Swiss'] },
    { code: 'OS', name: 'النمساوية', aliases: ['Austrian Airlines'] },
    { code: 'SN', name: 'بروكسل', aliases: ['Brussels Airlines'] },
    { code: 'IB', name: 'إيبيريا', aliases: ['Iberia'] },
    { code: 'TP', name: 'التاب البرتغالية', aliases: ['TAP Portugal', 'TAP Air Portugal'] },
    { code: 'AZ', name: 'آي تي أيه', aliases: ['ITA Airways', 'Alitalia'] },

    // أمريكا الشمالية
    { code: 'AC', name: 'إير كندا', aliases: ['Air Canada'] },
    { code: 'AA', name: 'الأمريكية', aliases: ['American Airlines'] },
    { code: 'DL', name: 'دلتا', aliases: ['Delta Air Lines', 'Delta'] },
    { code: 'UA', name: 'يونايتد', aliases: ['United Airlines', 'United'] },
    { code: 'WS', name: 'ويست جيت', aliases: ['WestJet'] },

    // آسيا الشرقية/الجنوبية
    { code: 'PK', name: 'باكستان', aliases: ['Pakistan International Airlines', 'PIA'] },
    { code: 'BG', name: 'بيمان', aliases: ['Biman Bangladesh Airlines', 'Biman'] },
    { code: 'AI', name: 'إير إنديا', aliases: ['Air India'] },
    { code: 'UL', name: 'سريلانكان', aliases: ['SriLankan Airlines'] },
    { code: 'MH', name: 'الماليزية', aliases: ['Malaysia Airlines', 'MAS'] },
    { code: 'GA', name: 'غارودا', aliases: ['Garuda Indonesia', 'Garuda'] },
    { code: 'SQ', name: 'سنغافورة', aliases: ['Singapore Airlines'] },
    { code: 'TG', name: 'التايلندية', aliases: ['Thai Airways'] },
    { code: 'CZ', name: 'الصين الجنوبية', aliases: ['China Southern Airlines'] }
  ],

  /**
   * تهيئة أولى: يحفظ القائمة الافتراضية إن لم تكن موجودة.
   */
  init: function() {
    var existing = GDS2.State.getJSON(GDS2.Config.PROP.AIRLINE_REGISTRY);
    if (existing && existing.length > 0) {
      return {
        status: 'already_exists',
        count: existing.length
      };
    }

    GDS2.State.setJSON(GDS2.Config.PROP.AIRLINE_REGISTRY, GDS2.AirlineRegistry.DEFAULT);
    GDS2.Log.info('AirlineRegistry: initialized', { count: GDS2.AirlineRegistry.DEFAULT.length });
    return {
      status: 'initialized',
      count: GDS2.AirlineRegistry.DEFAULT.length
    };
  },

  /**
   * القائمة الحالية من Script Properties.
   */
  getList: function() {
    var list = GDS2.State.getJSON(GDS2.Config.PROP.AIRLINE_REGISTRY);
    if (!list || list.length === 0) return GDS2.AirlineRegistry.DEFAULT;
    return list;
  },

  /**
   * البحث عن كود IATA شركة طيران بالاسم (عربي/إنجليزي/اختصار).
   * @return {string|null} الكود أو null إذا لم توجد مطابقة
   */
  getCode: function(airlineName) {
    if (!airlineName) return null;
    var name = String(airlineName).trim();
    if (!name) return null;

    var nameLower = name.toLowerCase();
    var list = GDS2.AirlineRegistry.getList();

    for (var i = 0; i < list.length; i++) {
      var entry = list[i];

      // مطابقة مباشرة بالاسم العربي
      if (entry.name === name) return entry.code;

      // مطابقة بالـ aliases (case-insensitive)
      if (entry.aliases && entry.aliases.length > 0) {
        for (var j = 0; j < entry.aliases.length; j++) {
          if (entry.aliases[j].toLowerCase() === nameLower) return entry.code;
        }
      }

      // مطابقة جزئية بالاسم العربي (الاسم يحوي أو محوي في entry)
      if (entry.name.indexOf(name) !== -1 || name.indexOf(entry.name) !== -1) {
        return entry.code;
      }

      // مطابقة الكود مباشرة (إن مرر الكود بدل الاسم)
      if (entry.code === name.toUpperCase()) return entry.code;
    }

    return null;
  },

  /**
   * إضافة شركة جديدة.
   */
  add: function(name, code, aliases) {
    if (!name || !code) return { status: 'error', reason: 'name and code required' };

    var upperCode = String(code).toUpperCase().trim();
    if (!/^[A-Z0-9]{2,3}$/.test(upperCode)) {
      return { status: 'error', reason: 'invalid code format', code: code };
    }

    var list = GDS2.AirlineRegistry.getList();

    // لا تكرار
    for (var i = 0; i < list.length; i++) {
      if (list[i].code === upperCode) {
        return { status: 'exists', code: upperCode };
      }
    }

    list.push({
      code: upperCode,
      name: name,
      aliases: aliases || []
    });

    GDS2.State.setJSON(GDS2.Config.PROP.AIRLINE_REGISTRY, list);
    GDS2.Log.info('AirlineRegistry: added', { code: upperCode, name: name });
    return { status: 'added', code: upperCode, total: list.length };
  }
};

// ==================== Global Entry Points ====================

function initAirlineRegistry() {
  return GDS2.AirlineRegistry.init();
}

function getAirlineList() {
  var list = GDS2.AirlineRegistry.getList();
  return {
    count: list.length,
    airlines: list
  };
}

function getAirlineCode(name) {
  return { input: name, code: GDS2.AirlineRegistry.getCode(name) };
}

function addAirline(name, code, aliases) {
  return GDS2.AirlineRegistry.add(name, code, aliases);
}

/**
 * دالة مجمّعة: تشغيل كل المرحلة 1 دفعة واحدة.
 */
function initAllRegistries() {
  var iata = GDS2.IATARegistry.extractFromHistory();
  var airlines = GDS2.AirlineRegistry.init();
  return {
    iata: iata,
    airlines: airlines
  };
}
