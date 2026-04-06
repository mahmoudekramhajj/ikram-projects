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
  var contractMap = buildContractMap_(ss); // packageId+hotel → contractId (سكن)
  var transportContracts = buildTransportContractMap_(ss); // route → contractId (نقل)

  // 2. تجميع العمليات
  var plan = {}; // { "yyyy-MM-dd": { opId: { pilgrims: [], count: 0, byGroup: {} } } }

  for (var i = 0; i < journeyRows.length; i++) {
    var p = journeyRows[i];
    if (!p.name && !p.passport) continue;

    var pd = pdMap[p.passport] || {};
    // B2B/B2C من رحلة الحاج مباشرة (عمود ArrivalFlightType)
    var isB2C = (p.arrivalFlightType || '').toUpperCase() === 'B2C';
    var transport = pkgTransport[p.packageId] || '';
    // fallback: قراءة التنقل من خريطة الباقات إذا لم يوجد في getTransportMap_
    if (!transport && pkgHotels[p.packageId]) {
      transport = pkgHotels[p.packageId].transport || '';
    }
    var isTrain = isTrainTransport_(transport);

    // === تعبئة أسماء الفنادق والتواريخ من الباقات ===
    var pkgH = pkgHotels[p.packageId] || {};
    if (!p.makkahEn && pkgH.makkahHotel) p.makkahEn = pkgH.makkahHotel;
    if (!p.madinahEn && pkgH.madinahHotel) p.madinahEn = pkgH.madinahHotel;
    if (!p.makkahShiftEn && pkgH.makkahShiftHotel) p.makkahShiftEn = pkgH.makkahShiftHotel;
    // حفظ تواريخ الباقة للاستخدام في Shifting والمغادرة
    p._pkgH = pkgH;
    // حفظ خريطة العقود
    p._contractMap = contractMap;
    p._transportContracts = transportContracts;

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

    // === E. المغادرة — حسب نوع النقل ===
    // مغادرة بالطيران
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

    // مغادرة بالقطار (حالات 27, 28) — تاريخ checkout آخر فندق
    if (isTrain) {
      var lastCity = getCityCode_(p.lastHouse);
      if (lastCity === 'Makkah') {
        // حالة 27: فندق مكة → محطة قطار مكة (مغادرة)
        var trainDate27 = pkgH.makkahShiftCheckOut || pkgH.makkahCheckOut || p.lastHouseEnd;
        if (trainDate27) {
          addToPlan_(plan, trainDate27, 27, p, pd, {
            time: '14:00',
            originHotel: p.makkahShiftEn || p.makkahEn,
            destHotel: 'محطة قطار مكة'
          });
        }
      } else if (lastCity === 'Madina') {
        // حالة 28: فندق المدينة → محطة قطار المدينة (مغادرة)
        var trainDate28 = pkgH.madinahCheckOut || p.lastHouseEnd;
        if (trainDate28) {
          addToPlan_(plan, trainDate28, 28, p, pd, {
            time: '14:00',
            originHotel: p.madinahEn,
            destHotel: 'محطة قطار المدينة'
          });
        }
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

  // فندق مكة للنقل بين المدن:
  // إذا الحاج يغادر مكة → آخر فندق مكة (الانتقالي إذا موجود)
  // إذا الحاج يصل مكة → أول فندق مكة
  var makkahHotelOut = p.makkahShiftEn || p.makkahEn; // آخر فندق مكة (للمغادرة)
  var makkahHotelIn = p.makkahEn || p.makkahShiftEn;  // أول فندق مكة (للوصول)
  var madinahHotel = p.madinahEn;

  if (isTrain) {
    if (fromMakkah) {
      return {
        opId: 8,  // فندق مكة → محطة قطار مكة
        date: transferDate,
        time: '12:00',
        originHotel: makkahHotelOut,
        destHotel: 'محطة قطار مكة',
        arrivalOpId: 9, // محطة قطار المدينة → فندق المدينة
        arrivalDate: transferDate,
        arrivalTime: null,
        arrivalDestHotel: madinahHotel
      };
    } else {
      return {
        opId: 10, // فندق المدينة → محطة قطار المدينة
        date: transferDate,
        time: '12:00',
        originHotel: madinahHotel,
        destHotel: 'محطة قطار المدينة',
        arrivalOpId: 11, // محطة قطار مكة → فندق مكة
        arrivalDate: transferDate,
        arrivalTime: null,
        arrivalDestHotel: makkahHotelIn
      };
    }
  } else {
    if (fromMakkah) {
      return {
        opId: 6,
        date: transferDate,
        time: '12:00',
        originHotel: makkahHotelOut,
        destHotel: madinahHotel
      };
    } else {
      return {
        opId: 7,
        date: transferDate,
        time: '12:00',
        originHotel: madinahHotel,
        destHotel: makkahHotelIn
      };
    }
  }
}

// ============================================================
// RESOLVE: تاريخ التحويل (Shifting) داخل مكة
// ============================================================

function resolveShiftingDate_(p) {
  // Shifting يحدث بين الفندقين بمكة
  // التاريخ = checkout الفندق الأول بمكة = check-in الفندق الثاني بمكة
  // نأخذ التواريخ من شيت الباقات (أدق من رحلة الحاج)

  var pkgH = p._pkgH || {};

  // تاريخ التحويل = checkout فندق مكة الأول = check-in فندق مكة الثاني
  if (pkgH.makkahCheckOut) return pkgH.makkahCheckOut;
  if (pkgH.makkahShiftCheckIn) return pkgH.makkahShiftCheckIn;

  // fallback: من بيانات رحلة الحاج
  var firstCity = getCityCode_(p.firstHouse);
  var lastCity = getCityCode_(p.lastHouse);

  if (firstCity === 'Makkah') {
    return p.firstHouseEnd;
  }
  if (lastCity === 'Makkah') {
    return p.lastHouseStart;
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
    // موعد تواجد الحافلة = ساعة قبل الانطلاق
    var depTime = extra.busTime || extra.time;
    if (depTime) entry.busReadyTime = subtractHours_(depTime, 1);
    // رقم عقد السكن — من الفندق الوجهة أو المصدر
    var cMap = pilgrim._contractMap || {};
    var hotelForContract = extra.destHotel || extra.originHotel || '';
    var contractKey = pilgrim.packageId + '|' + hotelForContract;
    entry.contractId = cMap[contractKey] || cMap[pilgrim.packageId + '|'] || '';
    // رقم عقد النقل — من نوع العملية
    var tContracts = pilgrim._transportContracts || {};
    entry.transportContractId = resolveTransportContract_(tContracts, opId, extra.originHotel || '', extra.destHotel || '');
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
    // المدينة — الكل
    'madinah_all': [1, 4, 5, 6, 7, 9, 10, 24, 25, 26, 28],
    // المدينة — تفصيلي
    'madinah_airport': [1, 24],  // مطار المدينة: وصول + مغادرة
    'madinah_station': [9, 10, 28],  // محطة قطار المدينة
    'madinah_hotels': [6, 7, 25, 26, 28],  // فنادق المدينة: مغادرة من/إلى الفنادق فقط
    // مكة — الكل
    'makkah_all': [2, 3, 4, 5, 6, 7, 8, 11, 12, 13, 14, 15, 19, 20, 21, 22, 23, 27],
    // مكة — تفصيلي
    'jeddah_t1': [2, 4, 22, 25],  // مطار جدة صالة 1
    'jeddah_north': [3, 5, 23, 26],  // مطار جدة الشمالية
    'makkah_station': [8, 11, 27],  // محطة قطار مكة
    'makkah_hotels': [6, 7, 12, 22, 23, 27],  // فنادق مكة: مغادرة من الفنادق فقط
    // المشاعر — الكل
    'mashaaer_all': [13, 14, 15, 16, 17, 18, 19, 20, 21],
    // المشاعر — تفصيلي
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

  var name = origin + ' إلى ' + dest;

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

  // قراءة حتى عمود BQ (69) لتشمل التنقل + عقود نسك
  var data = sheet.getRange(1, 1, lastRow, 69).getValues();
  var map = {};

  // أعمدة الفنادق في شيت الباقات:
  // فندق 1: City=L(11), NameEn=N(13), CheckIn=O(14), CheckOut=P(15)
  // فندق 2: City=AA(26), NameEn=AC(28), CheckIn=AD(29), CheckOut=AE(30)
  // فندق 3: City=AP(41), NameEn=AR(43), CheckIn=AS(44), CheckOut=AT(45)
  // التنقل: BN(65) = حافلة/قطار

  for (var i = 2; i < data.length; i++) {
    var nuskNo = cleanValue_(data[i][1]); // Nusk No = PackageId
    if (!nuskNo) continue;

    var hotels = [
      { city: cleanValue_(data[i][11]), name: cleanValue_(data[i][13]),
        checkIn: formatDate_(data[i][14]), checkOut: formatDate_(data[i][15]) },
      { city: cleanValue_(data[i][26]), name: cleanValue_(data[i][28]),
        checkIn: formatDate_(data[i][29]), checkOut: formatDate_(data[i][30]) },
      { city: cleanValue_(data[i][41]), name: cleanValue_(data[i][43]),
        checkIn: formatDate_(data[i][44]), checkOut: formatDate_(data[i][45]) }
    ];

    var transport = cleanValue_(data[i][65]); // عمود التنقل BN

    var entry = {
      makkahHotel: '', madinahHotel: '', makkahShiftHotel: '',
      makkahCheckIn: '', makkahCheckOut: '',
      makkahShiftCheckIn: '', makkahShiftCheckOut: '',
      madinahCheckIn: '', madinahCheckOut: '',
      transport: transport
    };

    // جمع فنادق مكة مع تواريخهم
    var makkahHotels = [];

    for (var h = 0; h < hotels.length; h++) {
      var city = hotels[h].city;
      var name = hotels[h].name;
      if (!city || !name) continue;

      if (isMakkahCity_(city)) {
        makkahHotels.push(hotels[h]);
      } else if (isMadinaCity_(city)) {
        entry.madinahHotel = name;
        entry.madinahCheckIn = hotels[h].checkIn;
        entry.madinahCheckOut = hotels[h].checkOut;
      }
    }

    // ترتيب فنادق مكة بتاريخ الدخول (الأبكر أولاً)
    if (makkahHotels.length >= 2) {
      makkahHotels.sort(function(a, b) {
        return (a.checkIn || '').localeCompare(b.checkIn || '');
      });
      entry.makkahHotel = makkahHotels[0].name;
      entry.makkahCheckIn = makkahHotels[0].checkIn;
      entry.makkahCheckOut = makkahHotels[0].checkOut;
      entry.makkahShiftHotel = makkahHotels[1].name;
      entry.makkahShiftCheckIn = makkahHotels[1].checkIn;
      entry.makkahShiftCheckOut = makkahHotels[1].checkOut;
    } else if (makkahHotels.length === 1) {
      entry.makkahHotel = makkahHotels[0].name;
      entry.makkahCheckIn = makkahHotels[0].checkIn;
      entry.makkahCheckOut = makkahHotels[0].checkOut;
    }

    map[nuskNo] = entry;
  }

  return map;
}

// ============================================================
// READ: خريطة عقود السكن (packageId+hotel → contractId)
// ============================================================

function buildContractMap_(ss) {
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.HOUSING_CONTRACTS);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  var map = {};
  // Headers: ContractId, HousingProvider, ..., PackageId (index 5), ...
  // نبحث عن فهرس الأعمدة من الهيدر
  var headers = data[0];
  var colIdx = {};
  for (var h = 0; h < headers.length; h++) {
    var hdr = String(headers[h]).trim();
    if (hdr === 'ContractId') colIdx.contractId = h;
    else if (hdr === 'HousingProvider') colIdx.provider = h;
    else if (hdr === 'PackageId') colIdx.packageId = h;
  }

  if (colIdx.contractId === undefined || colIdx.packageId === undefined) return {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var contractId = cleanValue_(row[colIdx.contractId]);
    var provider = cleanValue_(row[colIdx.provider]);
    var packageId = cleanValue_(row[colIdx.packageId]);
    if (!contractId || !packageId) continue;

    // مفتاح: packageId|hotelName
    var key = packageId + '|' + provider;
    map[key] = contractId;
    // مفتاح احتياطي بدون اسم الفندق
    if (!map[packageId + '|']) map[packageId + '|'] = contractId;
  }

  return map;
}

// ============================================================
// READ: خريطة عقود النقل (مسار → رقم العقد)
// ============================================================

function buildTransportContractMap_(ss) {
  var sheet = findSheet_(ss, TRANSPORT_CONFIG.SHEETS.TRANSPORT_CONTRACTS);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  if (data.length < 5) return {};

  // الشيت ملخص — نبحث عن صفوف فيها "من" و "إلى" و "رقم العقد"
  var map = {}; // "من|إلى" → contractId

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    // البحث عن صفوف البيانات: عمود A = من، B = إلى (أو بالعكس RTL)
    // نبحث عن أي صف فيه رقم عقد طويل (يبدأ بـ 2026)
    var contractId = '';
    var from = '';
    var to = '';

    for (var c = 0; c < row.length; c++) {
      var val = cleanValue_(row[c]);
      if (!val) continue;
      // رقم العقد: رقم طويل يبدأ بـ 2026
      if (/^2026\d{10,}$/.test(val)) {
        contractId = val;
      }
    }

    if (!contractId) continue;

    // البحث عن "من" و "إلى" في نفس الصف
    // عمود B = من, عمود C = إلى (حسب الصورة: A=#, B=من, C=إلى, D=عدد الحجاج, E=عدد الباصات, F=رقم العقد, G=رقم المسار)
    from = cleanValue_(row[1]); // عمود B = من
    to = cleanValue_(row[2]);   // عمود C = إلى

    if (from && to) {
      var key = from + '|' + to;
      map[key] = contractId;
    }
  }

  return map;
}

// ربط عملية نقل برقم عقد النقل
function resolveTransportContract_(transportContracts, opId, originDesc, destDesc) {
  if (!transportContracts || Object.keys(transportContracts).length === 0) return '';

  // محاولة مطابقة مباشرة
  var keys = Object.keys(transportContracts);
  for (var i = 0; i < keys.length; i++) {
    var parts = keys[i].split('|');
    var from = parts[0] || '';
    var to = parts[1] || '';

    // مطابقة بالكلمات المفتاحية
    if (matchRoute_(from, to, opId, originDesc, destDesc)) {
      return transportContracts[keys[i]];
    }
  }
  return '';
}

function matchRoute_(from, to, opId, originDesc, destDesc) {
  // عمليات المطارات والفنادق
  if (opId >= 1 && opId <= 5) {
    // وصول: مطار → فنادق
    if (from.indexOf('مطار') >= 0 && to.indexOf('فناد') >= 0) return true;
    if (from.indexOf('مطار') >= 0 && to.indexOf('مساكن') >= 0) return true;
  }
  if (opId >= 22 && opId <= 26) {
    // مغادرة: فنادق → مطار
    var isMakkahDep = (opId === 22 || opId === 23);
    var isMadinahDep = (opId === 24 || opId === 25 || opId === 26);

    if (isMakkahDep && from.indexOf('مكة') >= 0 && to.indexOf('مطار') >= 0) return true;
    if (isMakkahDep && from.indexOf('فناد') >= 0 && to.indexOf('جدة') >= 0) return true;
    if (isMadinahDep && from.indexOf('المدينة') >= 0 && to.indexOf('مطار') >= 0) return true;
    if (isMadinahDep && from.indexOf('فناد') >= 0 && to.indexOf('المدينة') >= 0) return true;
  }
  // عمليات القطار
  if (opId === 8 || opId === 27) {
    // مساكن مكة → محطة قطار مكة
    if (from.indexOf('مكة') >= 0 && to.indexOf('قطار') >= 0 && to.indexOf('مكة') >= 0) return true;
    if (from.indexOf('مساكن') >= 0 && from.indexOf('مكة') >= 0 && to.indexOf('قطار') >= 0) return true;
  }
  if (opId === 9) {
    // محطة قطار المدينة → مساكن المدينة
    if (from.indexOf('قطار') >= 0 && from.indexOf('المدينة') >= 0 && to.indexOf('المدينة') >= 0) return true;
  }
  if (opId === 10 || opId === 28) {
    // مساكن المدينة → محطة قطار المدينة
    if (from.indexOf('المدينة') >= 0 && to.indexOf('قطار') >= 0 && to.indexOf('المدينة') >= 0) return true;
    if (from.indexOf('مساكن') >= 0 && from.indexOf('المدينة') >= 0 && to.indexOf('قطار') >= 0) return true;
  }
  if (opId === 11) {
    // محطة قطار مكة → مساكن مكة
    if (from.indexOf('قطار') >= 0 && from.indexOf('مكة') >= 0 && to.indexOf('مكة') >= 0) return true;
  }
  return false;
}

function isMakkahCity_(city) {
  if (!city) return false;
  return city === 'Mak' || city === 'Makkah' || city === 'Makkah Shifting'
    || city.indexOf('مكة') >= 0 || city.indexOf('مك') >= 0;
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
