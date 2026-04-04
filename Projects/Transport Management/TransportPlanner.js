/**
 * Transport Management App — Planning Engine
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v1.0 — محرك التخطيط التلقائي — يقرأ بيانات الحجاج ويحسب الـ 28 عملية نقل
 */

// ============================================================
// AIRLINE → TERMINAL MAPPING (مطار جدة)
// ============================================================

// TERMINAL_MAP يُبنى ديناميكياً من شيت Airline_Terminals
var _terminalMapCache = null;

function buildAirlineTerminalMap_(ss) {
  if (_terminalMapCache) return _terminalMapCache;

  var map = {};
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.AIRLINE_TERMINALS);
  if (!sheet) {
    // fallback ثابت
    map = {
      'Emirates': 'T1', 'Etihad': 'T1', 'Etihad Airways': 'T1',
      'Gulf Air': 'T1', 'EgyptAir': 'T1', 'Egyptair': 'T1',
      'Qatar Airways': 'T1', 'Royal Jordanian': 'T1', 'Saudia': 'T1',
      'Turkish Airlines': 'T1', 'Ethiopian Airlines': 'T1',
      'Flyadeal': 'T1', 'flynas': 'T1',
      'AJet': 'N', 'VF': 'N', 'Wizz Air': 'N', 'W4': 'N',
      'Aegean Airlines': 'N', 'Air Cairo': 'N', 'AnadoluJet': 'N',
      'Flydubai': 'N', 'Pegasus Airlines': 'N', 'Pegasus': 'N',
      'WEST ISLE AIR INC.': 'N'
    };
    _terminalMapCache = map;
    return map;
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var code = cleanValue_(data[i][0]);    // Airline Code
    var nameEn = cleanValue_(data[i][1]);  // Airline Name EN
    var nameAr = cleanValue_(data[i][2]);  // Airline Name AR
    var terminal = cleanValue_(data[i][3]); // Terminal

    var t = (terminal === 'North Terminal') ? 'N' : 'T1';

    if (code) map[code] = t;
    if (nameEn) map[nameEn] = t;
    if (nameAr) map[nameAr] = t;
  }

  _terminalMapCache = map;
  return map;
}

function getJeddahTerminal_(airlineName) {
  if (!airlineName) return 'T1';
  var name = String(airlineName).trim();
  var map = _terminalMapCache || {};

  // بحث مباشر
  if (map[name]) return map[name];
  // بحث جزئي
  for (var key in map) {
    if (name.indexOf(key) >= 0 || key.indexOf(name) >= 0) {
      return map[key];
    }
  }
  return 'T1';
}

// ============================================================
// MAIN: توليد خطة النقل الكاملة
// ============================================================

function generateTransportPlan(dateFilter, locationFilter) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);

  // 1. قراءة البيانات
  var journeyRows = readJourneySheet_(ss);
  var pdMap = buildPDMap_(ss);           // passport → { contractType, guide, phone }
  var pkgTransport = getTransportMap_(ss); // packageId → transport type (bus/train)
  var pkgHotels = buildPackageHotelMap_(ss); // packageId → { makkahHotel, madinahHotel, makkahShiftHotel }
  var airlineTerminals = buildAirlineTerminalMap_(ss); // يبني _terminalMapCache

  // 2. تجميع العمليات
  var plan = {}; // { "yyyy-MM-dd": { opId: { pilgrims: [], count: 0, byGroup: {} } } }

  for (var i = 0; i < journeyRows.length; i++) {
    var p = journeyRows[i];
    if (!p.name && !p.passport) continue;

    var pd = pdMap[p.passport] || {};
    // B2B/B2C من رحلة الحاج مباشرة (عمود ArrivalFlightType)
    var isB2C = (p.arrivalFlightType || '').toUpperCase() === 'B2C';
    var transport = pkgTransport[p.packageId] || '';
    var isTrain = isTrainTransport_(transport);

    // === تعبئة أسماء الفنادق الفارغة من الباقات ===
    var pkgH = pkgHotels[p.packageId] || {};
    if (!p.makkahEn && pkgH.makkahHotel) p.makkahEn = pkgH.makkahHotel;
    if (!p.madinahEn && pkgH.madinahHotel) p.madinahEn = pkgH.madinahHotel;
    if (!p.makkahShiftEn && pkgH.makkahShiftHotel) p.makkahShiftEn = pkgH.makkahShiftHotel;

    // === تحديد الوجهة الفعلية عند الوصول ===
    var actualDest = resolveActualDestination_(p);

    // === A. استقبال المطار (B2B فقط) ===
    if (!isB2C && p.arrivalDate && p.arrivalFlight) {
      var arrOp = resolveArrivalOp_(p, actualDest.city);
      if (arrOp) {
        addToPlan_(plan, p.arrivalDate, arrOp.opId, p, pd, {
          time: p.arrivalTime,
          flight: p.arrivalFlight,
          airline: p.arrivalAirlineEn,
          destHotel: actualDest.hotel
        });

        // تحذير: حاج وصل قبل بداية فندقه
        if (actualDest.earlyArrival) {
          addToPlan_(plan, p.arrivalDate, 'WARN_EARLY', p, pd, {
            destHotel: actualDest.hotel
          });
        }
      }
    }

    // === B. النقل بين المدن ===
    // فقط إذا الحاج ذهب فعلاً للفندق الأول (لم يتخطاه)
    if (!actualDest.skippedFirst) {
      var intercity = resolveIntercityOp_(p, isTrain);
      if (intercity) {
        addToPlan_(plan, intercity.date, intercity.opId, p, pd, {
          time: intercity.time,
          originHotel: intercity.originHotel,
          destHotel: intercity.destHotel
        });
        if (intercity.arrivalOpId) {
          addToPlan_(plan, intercity.arrivalDate || intercity.date, intercity.arrivalOpId, p, pd, {
            time: intercity.arrivalTime,
            destHotel: intercity.arrivalDestHotel
          });
        }
      }
    }

    // === C. التحويل داخل مكة (Shifting) ===
    if (p.makkahShiftEn && p.makkahEn && p.makkahShiftEn !== p.makkahEn) {
      var shiftDate = resolveShiftingDate_(p);
      if (shiftDate) {
        addToPlan_(plan, shiftDate, 12, p, pd, {
          originHotel: p.makkahEn,
          destHotel: p.makkahShiftEn
        });
      }
    }

    // === D. عمليات الحج ===
    if (p.campName) {
      var campOp = resolveCampOp_(p.campName);
      if (campOp) {
        var hajjOriginHotel = p.makkahShiftEn || p.makkahEn;
        addToPlan_(plan, 'hajj_to_mina', campOp.toMina, p, pd, { originHotel: hajjOriginHotel, camp: p.campName });

        if (campOp.ritual) {
          addToPlan_(plan, 'hajj_arafat', 16, p, pd, {});
          addToPlan_(plan, 'hajj_muzdalifah', 17, p, pd, {});
          addToPlan_(plan, 'hajj_back_mina', 18, p, pd, {});
        }

        addToPlan_(plan, 'hajj_from_mina', campOp.fromMina, p, pd, { destHotel: hajjOriginHotel, camp: p.campName });
      }
    }

    // === E. المغادرة — جميع الحجاج ===
    if (p.returnDeptDate && p.returnFlight) {
      var depOp = resolveDepartureOp_(p);
      if (depOp) {
        var depOriginHotel = isMadinaCity_(p.lastHouse) ? p.madinahEn : (p.makkahShiftEn || p.makkahEn);
        addToPlan_(plan, p.returnDeptDate, depOp.opId, p, pd, {
          time: p.returnDeptTime,
          flight: p.returnFlight,
          airline: p.returnAirlineEn,
          originHotel: depOriginHotel,
          busTime: depOp.busTime
        });
      }
    }
  }

  // 3. فلتر التاريخ
  if (dateFilter) {
    var filtered = {};
    if (plan[dateFilter]) filtered[dateFilter] = plan[dateFilter];
    plan = filtered;
  }

  // 4. فلتر الموقع — إزالة العمليات التي لا تخص الموقع المختار
  if (locationFilter) {
    plan = filterPlanByLocation_(plan, locationFilter);
  }

  // 5. تحويل للصيغة النهائية
  return formatPlanOutput_(plan);
}

// ============================================================
// RESOLVE: تحديد عملية الوصول (1-5)
// ============================================================

// ============================================================
// RESOLVE: تحديد الوجهة الفعلية عند الوصول
// مقارنة تاريخ الوصول مع فترة الفندق الأول
// ============================================================

function resolveActualDestination_(p) {
  var arrDate = p.arrivalDate;
  var firstEnd = p.firstHouseEnd;
  var firstStart = p.firstHouseStart;
  var firstCity = getCityCode_(p.firstHouse);
  var lastCity = getCityCode_(p.lastHouse);

  // إذا لا يوجد فندقين مختلفين — الوجهة هي الفندق الأول
  if (!firstCity || !lastCity || firstCity === lastCity) {
    var hotel = firstCity === 'Makkah' ? (p.makkahEn || p.makkahShiftEn) : p.madinahEn;
    return { city: firstCity || 'Makkah', hotel: hotel || p.firstHouse, skippedFirst: false, earlyArrival: false };
  }

  // تحذير: وصل قبل بداية فندقه
  var earlyArrival = false;
  if (arrDate && firstStart && arrDate < firstStart) {
    earlyArrival = true;
  }

  // وصل بعد انتهاء الفندق الأول → يتخطى للفندق التالي
  if (arrDate && firstEnd && arrDate > firstEnd) {
    var skipHotel = lastCity === 'Makkah' ? (p.makkahEn || p.makkahShiftEn) : p.madinahEn;
    return { city: lastCity, hotel: skipHotel || p.lastHouse, skippedFirst: true, earlyArrival: false };
  }

  // عادي — يذهب للفندق الأول
  var normalHotel = firstCity === 'Makkah' ? (p.makkahEn || p.makkahShiftEn) : p.madinahEn;
  return { city: firstCity, hotel: normalHotel || p.firstHouse, skippedFirst: false, earlyArrival: earlyArrival };
}

function isMadinaCity_(city) {
  if (!city) return false;
  return city === 'Madina' || city === 'Madinah' || city === 'MED' || city.indexOf('مدين') >= 0;
}

function isJeddahCity_(city) {
  if (!city) return false;
  return city === 'Jeddah' || city === 'JED' || city.indexOf('جدة') >= 0;
}

function resolveArrivalOp_(p, actualDestCity) {
  var airportCity = String(p.arrivalCity || '').trim();
  var destIsMadinah = (actualDestCity === 'Madina');

  // مطار المدينة
  if (isMadinaCity_(airportCity)) {
    if (destIsMadinah) {
      return { opId: 1 }; // مطار المدينة → فنادق المدينة
    } else {
      // وصل المدينة لكن وجهته مكة (تخطى فندق المدينة)
      return { opId: 1 }; // يظل نفس المطار، الوجهة تتحدد من الفندق
    }
  }

  // مطار جدة — تحديد الصالة
  if (isJeddahCity_(airportCity)) {
    var terminal = getJeddahTerminal_(p.arrivalAirlineEn);

    if (terminal === 'N') {
      return { opId: destIsMadinah ? 5 : 3 };
    } else {
      return { opId: destIsMadinah ? 4 : 2 };
    }
  }

  return null;
}

// ============================================================
// RESOLVE: تحديد عملية النقل بين المدن (6-11)
// ============================================================

function resolveIntercityOp_(p, isTrain) {
  var firstHouse = String(p.firstHouse || '').trim();
  var lastHouse = String(p.lastHouse || '').trim();

  // لازم يكون عنده فندقين في مدينتين مختلفتين
  if (!firstHouse || !lastHouse) return null;
  var firstCity = getCityCode_(firstHouse);
  var lastCity = getCityCode_(lastHouse);
  if (firstCity === lastCity) return null;

  // الاتجاه: من المدينة الأولى للمدينة الثانية
  var fromMakkah = (firstCity === 'Makkah');
  var transferDate = p.firstHouseEnd; // تاريخ مغادرة الفندق الأول

  if (!transferDate) return null;

  var makkahHotel = p.makkahShiftEn || p.makkahEn;
  var madinahHotel = p.madinahEn;

  if (isTrain) {
    if (fromMakkah) {
      return {
        opId: 8,  // فندق مكة → محطة قطار مكة
        date: transferDate,
        time: null,
        originHotel: makkahHotel,
        destHotel: '',
        arrivalOpId: 9, // محطة قطار المدينة → فندق المدينة
        arrivalDate: transferDate,
        arrivalTime: null,
        arrivalDestHotel: madinahHotel
      };
    } else {
      return {
        opId: 10, // فندق المدينة → محطة قطار المدينة
        date: transferDate,
        time: null,
        originHotel: madinahHotel,
        destHotel: '',
        arrivalOpId: 11, // محطة قطار مكة → فندق مكة
        arrivalDate: transferDate,
        arrivalTime: null,
        arrivalDestHotel: makkahHotel
      };
    }
  } else {
    if (fromMakkah) {
      return {
        opId: 6,
        date: transferDate,
        time: TRANSPORT_CONFIG.CHECKOUT_TIME,
        originHotel: makkahHotel,
        destHotel: madinahHotel
      };
    } else {
      return {
        opId: 7,
        date: transferDate,
        time: TRANSPORT_CONFIG.CHECKOUT_TIME,
        originHotel: madinahHotel,
        destHotel: makkahHotel
      };
    }
  }
}

// ============================================================
// RESOLVE: تاريخ التحويل (Shifting) داخل مكة
// ============================================================

function resolveShiftingDate_(p) {
  // Shifting يحدث بين الفندقين بمكة
  // عادةً بعد انتهاء إقامة الفندق الأول بمكة
  // نستخدم MakkahShiftAr/En كاسم الفندق الثاني
  // التاريخ: نبحث عن الفجوة بين الفندقين

  var firstHouse = String(p.firstHouse || '').trim();
  var lastHouse = String(p.lastHouse || '').trim();

  // إذا الفندق الأول مكة — Shifting يحدث قبل الانتقال بين المدن
  if (firstHouse === 'Makkah' || firstHouse === 'Makkah Shifting') {
    // Shifting date = firstHouseEnd - 1 day or custom
    return p.firstHouseEnd; // تقريبي
  }

  // إذا الفندق الأخير مكة — Shifting يحدث بعد العودة من المدينة
  if (lastHouse === 'Makkah' || lastHouse === 'Makkah Shifting') {
    return p.lastHouseStart; // تقريبي
  }

  return null;
}

// ============================================================
// RESOLVE: تحديد عملية المخيم (13-21)
// ============================================================

function resolveCampOp_(campName) {
  var camp = String(campName || '').trim().toLowerCase();

  if (camp.indexOf('معيصم') >= 0 || camp.indexOf('maisam') >= 0 || camp.indexOf('maisem') >= 0) {
    return { toMina: 13, fromMina: 19, ritual: true };  // المعيصم — بالحافلة لعرفة
  }
  if (camp.indexOf('مجر') >= 0 || camp.indexOf('mujar') >= 0) {
    return { toMina: 14, fromMina: 20, ritual: false }; // مجر الكبش — قطار داخلي لعرفة
  }
  if (camp.indexOf('72') >= 0) {
    return { toMina: 15, fromMina: 21, ritual: false }; // 72 — قطار داخلي لعرفة
  }

  return null;
}

// ============================================================
// RESOLVE: تحديد عملية المغادرة (22-28)
// ============================================================

function resolveDepartureOp_(p) {
  var deptCity = String(p.returnDeptCity || '').trim();
  var lastHouse = String(p.lastHouse || '').trim();
  var lastCity = getCityCode_(lastHouse);
  var deptTime = formatTime_(p.returnDeptTime);

  // مطار المدينة
  if (isMadinaCity_(deptCity)) {
    if (lastCity === 'Madina') {
      var busTime24 = subtractHours_(deptTime, 3);
      return { opId: 24, busTime: busTime24 };
    } else {
      var busTime24m = subtractHours_(deptTime, 12);
      return { opId: 24, busTime: busTime24m };
    }
  }

  // مطار جدة
  if (isJeddahCity_(deptCity)) {
    var terminal = getJeddahTerminal_(p.returnAirlineEn);

    if (lastCity === 'Madina') {
      var margin = 12; // المدينة → جدة
      var busTime = subtractHours_(deptTime, margin);
      return { opId: terminal === 'N' ? 26 : 25, busTime: busTime };
    } else {
      var margin = 8; // مكة → جدة
      var busTime = subtractHours_(deptTime, margin);
      return { opId: terminal === 'N' ? 23 : 22, busTime: busTime };
    }
  }

  return null;
}

// ============================================================
// HELPER: إضافة حاج للخطة
// ============================================================

function addToPlan_(plan, date, opId, pilgrim, pd, extra) {
  var dateKey = String(date || 'unscheduled');
  if (!plan[dateKey]) plan[dateKey] = {};

  // بناء مفتاح فريد حسب نوع العملية
  var originHotel = (extra && extra.originHotel) || '';
  var destHotel = (extra && extra.destHotel) || '';
  var routeKey = '';

  if (originHotel && destHotel) {
    routeKey = originHotel + '>' + destHotel; // نقل بين فندقين
  } else if (destHotel) {
    routeKey = destHotel; // وصول — المفتاح هو فندق الوجهة
  } else if (originHotel) {
    routeKey = originHotel; // مغادرة — المفتاح هو فندق المصدر
  }

  var planKey = String(opId) + (routeKey ? '|' + routeKey : '');

  if (!plan[dateKey][planKey]) {
    plan[dateKey][planKey] = {
      opId: opId,
      originHotel: originHotel,
      destHotel: destHotel,
      pilgrims: [],
      byFlight: {}
    };
  }

  var entry = {
    bookingId: pilgrim.bookingId,
    name: pilgrim.name,
    passport: pilgrim.passport,
    packageId: pilgrim.packageId,
    groupNumber: pilgrim.groupNumber,
    gender: pilgrim.gender,
    nationalityEn: pilgrim.nationalityEn,
    campName: pilgrim.campName,
    guide: pd.guide || '',
    phone: pd.phone || ''
  };

  if (extra) {
    if (extra.flight) entry.flight = extra.flight;
    if (extra.airline) entry.airline = extra.airline;
    if (extra.time) entry.time = extra.time;
    if (extra.busTime) entry.busTime = extra.busTime;
    if (extra.camp) entry.camp = extra.camp;
    if (extra.originHotel) entry.originHotel = extra.originHotel;
    if (extra.destHotel) entry.destHotel = extra.destHotel;
  }

  plan[dateKey][planKey].pilgrims.push(entry);

  // تجميع بالرحلة
  if (extra && extra.flight) {
    var fKey = extra.flight + (extra.time ? ' ' + extra.time : '');
    if (!plan[dateKey][planKey].byFlight[fKey]) plan[dateKey][planKey].byFlight[fKey] = [];
    plan[dateKey][planKey].byFlight[fKey].push(entry);
  }
}

// ============================================================
// FORMAT: تحويل الخطة للصيغة النهائية
// ============================================================

function formatPlanOutput_(plan) {
  var output = { dates: [], summary: { totalPilgrims: 0, totalOps: 0 } };
  var dateKeys = Object.keys(plan).sort();

  for (var d = 0; d < dateKeys.length; d++) {
    var dateKey = dateKeys[d];
    var ops = plan[dateKey];
    var dateEntry = { date: dateKey, operations: [], totalPilgrims: 0 };

    // ترتيب بالرقم ثم بالفندق
    var planKeys = Object.keys(ops).sort(function(a, b) {
      var aId = Number(a.split('|')[0]);
      var bId = Number(b.split('|')[0]);
      if (aId !== bId) return aId - bId;
      return a.localeCompare(b);
    });

    for (var o = 0; o < planKeys.length; o++) {
      var key = planKeys[o];
      var opData = ops[key];
      var opId = opData.opId;
      var opDef = OPERATIONS[opId] || {};
      var originHotel = opData.originHotel || '';
      var destHotel = opData.destHotel || '';
      var count = opData.pilgrims.length;

      // بناء اسم العملية مع أسماء الفنادق الفعلية
      var opName = buildOpName_(opDef, originHotel, destHotel);

      // تجميع بالرحلات
      var flightGroups = [];
      for (var fk in opData.byFlight) {
        flightGroups.push({ flight: fk, count: opData.byFlight[fk].length });
      }
      flightGroups.sort(function(a, b) { return b.count - a.count; });

      dateEntry.operations.push({
        opId: opId,
        planKey: key,
        opName: opName,
        originHotel: originHotel,
        destHotel: destHotel,
        category: opDef.category || '',
        origin: opDef.origin || '',
        destination: opDef.dest || '',
        count: count,
        flightGroups: flightGroups,
        pilgrims: opData.pilgrims
      });

      dateEntry.totalPilgrims += count;
      output.summary.totalOps++;
    }

    output.dates.push(dateEntry);
    output.summary.totalPilgrims += dateEntry.totalPilgrims;
  }

  return JSON.stringify(output);
}

// ============================================================
// READ: قراءة شيت رحلة الحاج
// ============================================================

function readJourneySheet_(ss) {
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.JOURNEY);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var J = TRANSPORT_CONFIG.JOURNEY_COLS;
  var pilgrims = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = cleanValue_(row[J.NAME]);
    var passport = cleanValue_(row[J.PASSPORT]);
    if (!name && !passport) continue;

    pilgrims.push({
      bookingId: cleanValue_(row[J.BOOKING_ID]),
      packageId: cleanValue_(row[J.PACKAGE_ID]),
      groupNumber: cleanValue_(row[J.GROUP_NUMBER]),
      name: name,
      passport: passport,
      gender: cleanValue_(row[J.GENDER]),
      nationalityEn: cleanValue_(row[J.NATIONALITY_EN]),
      campName: cleanValue_(row[J.CAMP_NAME]),
      arrivalAirlineEn: cleanValue_(row[J.ARRIVAL_AIRLINE_EN]),
      arrivalTime: formatTime_(row[J.ARRIVAL_TIME]),
      arrivalCity: cleanValue_(row[J.ARRIVAL_CITY]),
      arrivalDate: formatDate_(row[J.ARRIVAL_DATE]),
      arrivalFlight: cleanValue_(row[J.ARRIVAL_FLIGHT]),
      arrivalFlightType: cleanValue_(row[J.ARRIVAL_FLIGHT_TYPE]),
      returnAirlineEn: cleanValue_(row[J.RETURN_AIRLINE_EN]),
      returnDeptCity: cleanValue_(row[J.RETURN_DEPT_CITY]),
      returnDeptDate: formatDate_(row[J.RETURN_DEPT_DATE]),
      returnDeptTime: formatTime_(row[J.RETURN_DEPT_TIME]),
      returnFlight: cleanValue_(row[J.RETURN_FLIGHT]),
      firstHouse: cleanValue_(row[J.FIRST_HOUSE]),
      firstHouseStart: formatDate_(row[J.FIRST_HOUSE_START]),
      firstHouseEnd: formatDate_(row[J.FIRST_HOUSE_END]),
      lastHouse: cleanValue_(row[J.LAST_HOUSE]),
      lastHouseStart: formatDate_(row[J.LAST_HOUSE_START]),
      lastHouseEnd: formatDate_(row[J.LAST_HOUSE_END]),
      makkahEn: cleanValue_(row[J.MAKKAH_EN]),
      makkahShiftEn: cleanValue_(row[J.MAKKAH_SHIFT_EN]),
      madinahEn: cleanValue_(row[J.MADINAH_EN])
    });
  }

  return pilgrims;
}

// ============================================================
// READ: بناء خريطة Presonal Details (passport → بيانات)
// ============================================================

function buildPDMap_(ss) {
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.PERSONAL_DETAILS);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var data = sheet.getDataRange().getValues();
  var PD = TRANSPORT_CONFIG.PD_COLS;
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var passport = String(row[PD.PASSPORT] || '').trim();
    if (!passport) continue;

    map[passport] = {
      contractType: String(row[PD.FLIGHT_CONTRACT_TYPE] || '').trim(),
      guide: String(row[PD.GUIDE_NAME] || '').trim(),
      phone: String(row[PD.PHONE] || '').trim(),
      nameAr: (String(row[PD.FIRST_NAME_AR] || '') + ' ' + String(row[PD.LAST_NAME_AR] || '')).trim(),
      camp: String(row[PD.CAMP] || '').trim(),
      ticketNumber: String(row[PD.TICKET_NUMBER] || '').trim()
    };
  }

  return map;
}

// ============================================================
// FILTER: فلتر حسب الموقع
// ============================================================

function filterPlanByLocation_(plan, locationId) {
  // تحديد العمليات المرتبطة بكل موقع
  var locationOps = {
    // المدينة
    'madinah_airport': [1, 24],  // مطار المدينة: وصول + مغادرة
    'madinah_station': [9, 10],  // محطة قطار المدينة
    'madinah_hotels': [1, 4, 5, 6, 7, 9, 10, 24, 25, 26, 28],  // فنادق المدينة
    // مكة
    'jeddah_t1': [2, 4, 22, 25],  // مطار جدة صالة 1
    'jeddah_north': [3, 5, 23, 26],  // مطار جدة الشمالية
    'makkah_station': [8, 11],  // محطة قطار مكة
    'makkah_hotels': [2, 3, 6, 7, 8, 11, 12, 13, 14, 15, 19, 20, 21, 22, 23, 27],  // فنادق مكة
    // المشاعر
    'camp_maisam': [13, 16, 17, 18, 19],
    'camp_mujar': [14, 20],
    'camp_72': [15, 21],
    'arafat': [16, 17],
    'muzdalifah': [17, 18]
  };

  var allowedOps = locationOps[locationId];
  if (!allowedOps) return plan; // لا فلتر

  var filteredPlan = {};
  for (var dateKey in plan) {
    var dateOps = plan[dateKey];
    var filteredOps = {};

    for (var planKey in dateOps) {
      var opData = dateOps[planKey];
      if (allowedOps.indexOf(opData.opId) >= 0) {
        filteredOps[planKey] = opData;
      }
    }

    if (Object.keys(filteredOps).length > 0) {
      filteredPlan[dateKey] = filteredOps;
    }
  }

  return filteredPlan;
}

// ============================================================
// HELPER: بناء اسم العملية مع الفنادق
// ============================================================

function buildOpName_(opDef, originHotel, destHotel) {
  if (!opDef || !opDef.name) return 'عملية غير معروفة';

  var defOrigin = opDef.origin || '';
  var defDest = opDef.dest || '';

  // استبدال المصدر والوجهة بأسماء الفنادق الفعلية
  var origin = originHotel || defOrigin;
  var dest = destHotel || defDest;

  var name = origin + ' → ' + dest;

  // إضافة نوع النقل
  if (opDef.transport === 'bus') name += ' (حافلة)';

  return name;
}

// ============================================================
// READ: خريطة فنادق الباقات (مصدر احتياطي لأسماء الفنادق)
// ============================================================

function buildPackageHotelMap_(ss) {
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.PACKAGES);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return {};

  var data = sheet.getRange(1, 1, lastRow, 54).getValues(); // حتى عمود BB
  var map = {};

  // أعمدة الفنادق في شيت الباقات:
  // فندق 1: City=L(11), NameEn=N(13)
  // فندق 2: City=AA(26), NameEn=AC(28)
  // فندق 3: City=AP(41), NameEn=AR(43)

  for (var i = 2; i < data.length; i++) {
    var nuskNo = cleanValue_(data[i][1]); // Nusk No = PackageId
    if (!nuskNo) continue;

    var h1City = cleanValue_(data[i][11]);
    var h1Name = cleanValue_(data[i][13]);
    var h2City = cleanValue_(data[i][26]);
    var h2Name = cleanValue_(data[i][28]);
    var h3City = cleanValue_(data[i][41]);
    var h3Name = cleanValue_(data[i][43]);

    var entry = { makkahHotel: '', madinahHotel: '', makkahShiftHotel: '' };
    var makkahCount = 0;

    var hotels = [
      { city: h1City, name: h1Name },
      { city: h2City, name: h2Name },
      { city: h3City, name: h3Name }
    ];

    for (var h = 0; h < hotels.length; h++) {
      var city = hotels[h].city;
      var name = hotels[h].name;
      if (!city || !name) continue;

      if (city === 'Med' || city === 'Madina' || city === 'Madinah' || city.indexOf('مدين') >= 0) {
        entry.madinahHotel = name;
      } else if (city === 'Mak' || city === 'Makkah' || city === 'Makkah Shifting' || city.indexOf('مكة') >= 0 || city.indexOf('مك') >= 0) {
        makkahCount++;
        if (makkahCount === 1) {
          entry.makkahHotel = name;
        } else {
          entry.makkahShiftHotel = name;
        }
      }
    }

    map[nuskNo] = entry;
  }

  return map;
}

// ============================================================
// API: دالة مكشوفة للواجهة
// ============================================================

function getTransportPlan(dateFilter, locationFilter) {
  return generateTransportPlan(dateFilter || null, locationFilter || null);
}

// ============================================================
// API: ملخص سريع — عدد الحجاج لكل يوم
// ============================================================

function getTransportSummary(locationFilter) {
  var ss = SpreadsheetApp.openById(TRANSPORT_CONFIG.SPREADSHEET_ID);
  var journeyRows = readJourneySheet_(ss);

  var summary = {}; // { "yyyy-MM-dd": { arrivals: N, departures: N, transfers: N } }

  for (var i = 0; i < journeyRows.length; i++) {
    var p = journeyRows[i];
    if (!p.name && !p.passport) continue;
    var isB2C = (p.arrivalFlightType || '').toUpperCase() === 'B2C';

    // وصول (B2B فقط)
    if (!isB2C && p.arrivalDate) {
      if (!summary[p.arrivalDate]) summary[p.arrivalDate] = { arrivals: 0, departures: 0, transfers: 0 };
      summary[p.arrivalDate].arrivals++;
    }

    // مغادرة
    if (p.returnDeptDate) {
      if (!summary[p.returnDeptDate]) summary[p.returnDeptDate] = { arrivals: 0, departures: 0, transfers: 0 };
      summary[p.returnDeptDate].departures++;
    }

    // نقل بين المدن
    if (p.firstHouseEnd && p.firstHouse && p.lastHouse) {
      var fc = getCityCode_(p.firstHouse);
      var lc = getCityCode_(p.lastHouse);
      if (fc !== lc) {
        if (!summary[p.firstHouseEnd]) summary[p.firstHouseEnd] = { arrivals: 0, departures: 0, transfers: 0 };
        summary[p.firstHouseEnd].transfers++;
      }
    }
  }

  return JSON.stringify(summary);
}

// ============================================================
// API: كشوفات عملية واحدة — أسماء الحجاج للتصدير
// ============================================================

function getOperationManifest(dateFilter, opId) {
  var planJson = generateTransportPlan(dateFilter);
  var plan = JSON.parse(planJson);

  for (var d = 0; d < plan.dates.length; d++) {
    var dateEntry = plan.dates[d];
    if (dateFilter && dateEntry.date !== dateFilter) continue;

    for (var o = 0; o < dateEntry.operations.length; o++) {
      var op = dateEntry.operations[o];
      if (op.opId === Number(opId)) {
        return JSON.stringify({
          date: dateEntry.date,
          operation: op.opName,
          count: op.count,
          flightGroups: op.flightGroups,
          hotelGroups: op.hotelGroups,
          pilgrims: op.pilgrims
        });
      }
    }
  }

  return JSON.stringify({ error: 'لا توجد بيانات لهذه العملية' });
}
