// ============================================
// DataFetcher.js — قراءة ذكية من المصدر الأحدث لكل نوع بيانات
// ============================================

// ============================================
// قراءة الشيتات مع كاش
// ============================================

/**
 * قراءة شيت كامل (مع كاش)
 */
function getSheetData_(sheetName, cacheKey, ttl) {
  var cached = getCache_(cacheKey);
  if (cached) return cached;

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log('Sheet not found: ' + sheetName);
      return [];
    }
    var data = sheet.getDataRange().getValues();
    var rows = data.length > 1 ? data.slice(1) : []; // تخطي الهيدر

    // CacheService max 100KB — إذا كبيرة جداً لا نخزنها
    var json = JSON.stringify(rows);
    if (json.length < 95000) {
      setCache_(cacheKey, rows, ttl || CACHE_TTL);
    }
    return rows;
  } catch (e) {
    Logger.log('getSheetData_ error (' + sheetName + '): ' + e.message);
    return [];
  }
}

// ============================================
// Presonal Details — المصدر الأحدث للبيانات الشخصية
// ============================================

function getPersonalDetails_() {
  return getSheetData_(SHEETS.PERSONAL, 'adm_pd', CACHE_TTL);
}

/**
 * بحث حاج بالجواز أو الاسم أو رقم الحجز
 * يبحث في Presonal Details (المصدر الأحدث)
 */
function searchPilgrims_(query) {
  var data = getPersonalDetails_();
  if (!data || data.length === 0) return [];

  var q = String(query).toUpperCase().trim();
  var results = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row || !row[PD.PASSPORT]) continue;

    var passport = String(row[PD.PASSPORT] || '').toUpperCase().trim();
    var nameEn = String(row[PD.FIRST_NAME_EN] || '') + ' ' + String(row[PD.LAST_NAME_EN] || '');
    var nameAr = String(row[PD.FIRST_NAME_AR] || '') + ' ' + String(row[PD.LAST_NAME_AR] || '');
    var seq = String(row[PD.SEQ] || '');
    var group = String(row[PD.GROUP] || '');
    var phone = String(row[PD.PHONE] || '');

    if (passport === q ||
        seq === q ||
        group === q ||
        phone.indexOf(q) !== -1 ||
        nameEn.toUpperCase().indexOf(q) !== -1 ||
        nameAr.indexOf(q) !== -1) {
      results.push({
        index: i,
        seq: seq,
        group: group,
        passport: passport,
        nameEn: nameEn.trim(),
        nameAr: (nameAr || '').trim(),
        phone: row[PD.PHONE] || '',
        email: row[PD.EMAIL] || '',
        nationality: row[PD.NATIONALITY] || '',
        countryResidence: row[PD.COUNTRY_RESIDENCE] || '',
        guideName: row[PD.GUIDE_NAME] || '',
        packageNo: row[PD.PACKAGE_NO] || '',
        packageName: row[PD.PACKAGE_NAME] || '',
        visaStatus: row[PD.VISA_STATUS] || '',
        ticketNo: row[PD.TICKET_NO] || '',
        camp: row[PD.CAMP] || '',
        flightType: row[PD.FLIGHT_TYPE] || '',
        transportArrType: row[PD.TRANSPORT_ARR_TYPE] || '',
        transportDepType: row[PD.TRANSPORT_DEP_TYPE] || '',
        _row: row
      });

      if (results.length >= 20) break;
    }
  }

  return results;
}

/**
 * جلب بيانات حاج كاملة من جميع المصادر
 */
function getPilgrimFull_(seq) {
  var pd = getPersonalDetails_();
  var pilgrim = null;

  // 1. البحث في Presonal Details
  for (var i = 0; i < pd.length; i++) {
    if (String(pd[i][PD.SEQ]) === String(seq)) {
      pilgrim = pd[i];
      break;
    }
  }
  if (!pilgrim) return null;

  var result = {
    // بيانات شخصية (من PD — الأحدث)
    seq: pilgrim[PD.SEQ],
    group: pilgrim[PD.GROUP],
    nameEn: String(pilgrim[PD.FIRST_NAME_EN] || '') + ' ' + String(pilgrim[PD.LAST_NAME_EN] || ''),
    nameAr: String(pilgrim[PD.FIRST_NAME_AR] || '') + ' ' + String(pilgrim[PD.LAST_NAME_AR] || ''),
    passport: pilgrim[PD.PASSPORT],
    passportExp: pilgrim[PD.PASSPORT_EXP],
    dob: pilgrim[PD.DOB],
    gender: pilgrim[PD.GENDER],
    phone: pilgrim[PD.PHONE],
    email: pilgrim[PD.EMAIL],
    nationality: pilgrim[PD.NATIONALITY],
    countryResidence: pilgrim[PD.COUNTRY_RESIDENCE],
    guideName: pilgrim[PD.GUIDE_NAME],
    packageNo: pilgrim[PD.PACKAGE_NO],
    packageName: pilgrim[PD.PACKAGE_NAME],
    flightType: pilgrim[PD.FLIGHT_TYPE],
    contractName: pilgrim[PD.CONTRACT_NAME],
    visaStatus: pilgrim[PD.VISA_STATUS],
    ticketNo: pilgrim[PD.TICKET_NO],
    ticketUrl: pilgrim[PD.TICKET_URL],
    camp: pilgrim[PD.CAMP],
    transportArrType: pilgrim[PD.TRANSPORT_ARR_TYPE],
    transportArrTime: pilgrim[PD.TRANSPORT_ARR_TIME],
    transportDepType: pilgrim[PD.TRANSPORT_DEP_TYPE],
    transportDepTime: pilgrim[PD.TRANSPORT_DEP_TIME],
    // سيُملأ من المصادر الأخرى
    flight: null,
    hotels: null,
    reception: null,
    packageDetails: null,
    intercityTransport: null
  };

  // 2. بيانات الرحلة والفنادق (من رحلة الحاج 2 — الأحدث)
  var j2 = getJourney2Data_();
  var groupId = String(pilgrim[PD.GROUP]);
  var gender = String(pilgrim[PD.GENDER] || '').toLowerCase();
  var isMain = String(pilgrim[PD.TYPE] || '');

  for (var j = 0; j < j2.length; j++) {
    if (String(j2[j][J2.APPLICATION_ID]) === groupId) {
      // مطابقة الفرد داخل المجموعة
      var j2Gender = String(j2[j][J2.GENDER] || '').toLowerCase();
      var j2IsMain = String(j2[j][J2.IS_MAIN] || '');

      // تطابق الجنس + النوع، أو إذا المجموعة فرد واحد
      var genderMatch = (gender.indexOf('ذكر') !== -1 || gender === 'male') ?
        (j2Gender === 'male') : (j2Gender === 'female');
      var typeMatch = (isMain.indexOf('رئيسي') !== -1 || j2IsMain === '1');

      if (genderMatch || j2.filter(function(r) { return String(r[J2.APPLICATION_ID]) === groupId; }).length === 1) {
        result.flight = {
          arrAirline: j2[j][J2.ARR_AIRLINE_EN] || j2[j][J2.ARR_AIRLINE_AR],
          arrFlightNo: j2[j][J2.ARR_FLIGHT_NO],
          arrFrom: j2[j][J2.ARR_DEPART_CITY],
          arrTo: j2[j][J2.ARR_ARRIVE_CITY],
          arrDate: j2[j][J2.ARR_DEPART_DATE],
          arrTime: j2[j][J2.ARR_DEPART_TIME],
          arrLandDate: j2[j][J2.ARR_ARRIVE_DATE],
          arrLandTime: j2[j][J2.ARR_ARRIVAL_TIME],
          retAirline: j2[j][J2.RET_AIRLINE_EN] || j2[j][J2.RET_AIRLINE_AR],
          retFlightNo: j2[j][J2.RET_FLIGHT_NO],
          retFrom: j2[j][J2.RET_DEPART_CITY],
          retTo: j2[j][J2.RET_ARRIVE_CITY],
          retDate: j2[j][J2.RET_DEPART_DATE],
          retTime: j2[j][J2.RET_DEPART_TIME],
          retLandDate: j2[j][J2.RET_ARRIVE_DATE],
          retLandTime: j2[j][J2.RET_ARRIVAL_TIME]
        };
        result.hotels = {
          first: { name: j2[j][J2.FIRST_HOUSE], start: j2[j][J2.FIRST_HOUSE_START], end: j2[j][J2.FIRST_HOUSE_END] },
          last: { name: j2[j][J2.LAST_HOUSE], start: j2[j][J2.LAST_HOUSE_START], end: j2[j][J2.LAST_HOUSE_END] },
          makkahAr: j2[j][J2.MAKKAH_AR], makkahEn: j2[j][J2.MAKKAH_EN],
          madinahAr: j2[j][J2.MADINAH_AR], madinahEn: j2[j][J2.MADINAH_EN]
        };
        break;
      }
    }
  }

  // 3. حالة الاستقبال (من رحلة الحاج — المصدر الوحيد)
  var journey = getJourneyData_();
  var pp = String(pilgrim[PD.PASSPORT] || '').toUpperCase().trim();
  for (var k = 0; k < journey.length; k++) {
    if (String(journey[k][JRN.PASSPORT] || '').toUpperCase().trim() === pp) {
      var status = journey[k][JRN.RECEPTION_STATUS];
      result.reception = {
        status: status ? 'received' : 'not_received',
        time: journey[k][JRN.RECEPTION_TIME] || '',
        staff: journey[k][JRN.RECEPTION_STAFF] || ''
      };
      break;
    }
  }

  // 4. تفاصيل الباقة + النقل بين المدن
  var pkgs = getPackagesData_();
  var pkgNo = String(pilgrim[PD.PACKAGE_NO] || '');
  for (var p = 0; p < pkgs.length; p++) {
    if (String(pkgs[p][PKG.NUSK_NO]) === pkgNo || String(pkgs[p][PKG.IKRAM_NO]) === pkgNo) {
      result.packageDetails = {
        nameAr: pkgs[p][PKG.NAME_AR],
        nameEn: pkgs[p][PKG.NAME_EN],
        dateStart: pkgs[p][PKG.DATE_START],
        dateEnd: pkgs[p][PKG.DATE_END],
        noDays: pkgs[p][PKG.NO_DAYS],
        capacity: pkgs[p][PKG.NO_PILGRIM],
        sales: pkgs[p][PKG.SALES],
        remaining: pkgs[p][PKG.REMAINING]
      };
      result.intercityTransport = pkgs[p][PKG.TRANSPORT] || 'حافلة';
      break;
    }
  }

  return result;
}

// ============================================
// رحلة الحاج 2 — المصدر الأحدث للرحلة والفنادق
// ============================================

function getJourney2Data_() {
  return getSheetData_(SHEETS.JOURNEY2, 'adm_j2', CACHE_TTL);
}

// ============================================
// رحلة الحاج (المدمج) — فقط لحالة الاستقبال
// ============================================

function getJourneyData_() {
  return getSheetData_(SHEETS.JOURNEY, 'adm_journey', CACHE_TTL);
}

// ============================================
// الباقات
// ============================================

function getPackagesData_() {
  // الباقات تبدأ من الصف 3 (صفين هيدر)
  var cached = getCache_('adm_pkgs');
  if (cached) return cached;

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.PACKAGES);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var rows = data.length > 2 ? data.slice(2) : []; // تخطي صفين هيدر
    // فلترة الصفوف الفارغة
    rows = rows.filter(function(r) { return r[PKG.NUSK_NO] || r[PKG.NAME_AR]; });
    setCache_('adm_pkgs', rows, CACHE_TTL);
    return rows;
  } catch (e) {
    Logger.log('getPackagesData_ error: ' + e.message);
    return [];
  }
}

/**
 * إحصائيات الباقات
 */
function getPackagesStats_() {
  var pkgs = getPackagesData_();
  var stats = [];
  var totalCapacity = 0, totalSold = 0;

  for (var i = 0; i < pkgs.length; i++) {
    var row = pkgs[i];
    if (!row[PKG.NAME_AR] && !row[PKG.NAME_EN]) continue;

    var capacity = Number(row[PKG.NO_PILGRIM]) || 0;
    var sold = Number(row[PKG.SALES]) || 0;
    var remaining = Number(row[PKG.REMAINING]) || 0;
    var pct = capacity > 0 ? Math.round((sold / capacity) * 100) : 0;

    totalCapacity += capacity;
    totalSold += sold;

    var icon = pct >= 80 ? '🔴' : pct >= 50 ? '🟡' : '🟢';

    stats.push({
      nuskNo: row[PKG.NUSK_NO],
      nameAr: row[PKG.NAME_AR],
      nameEn: row[PKG.NAME_EN],
      capacity: capacity,
      sold: sold,
      remaining: remaining,
      percent: pct,
      icon: icon,
      dateStart: row[PKG.DATE_START],
      dateEnd: row[PKG.DATE_END],
      transport: row[PKG.TRANSPORT] || 'حافلة'
    });
  }

  return {
    packages: stats,
    totalCapacity: totalCapacity,
    totalSold: totalSold,
    totalRemaining: totalCapacity - totalSold,
    totalPercent: totalCapacity > 0 ? Math.round((totalSold / totalCapacity) * 100) : 0
  };
}

// ============================================
// الطيران
// ============================================

function getFlightsData_() {
  var cached = getCache_('adm_flights');
  if (cached) return cached;

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.FLIGHTS);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var rows = data.length > 2 ? data.slice(2) : [];
    rows = rows.filter(function(r) { return r[FLT.PNR]; });
    setCache_('adm_flights', rows, CACHE_TTL);
    return rows;
  } catch (e) {
    Logger.log('getFlightsData_ error: ' + e.message);
    return [];
  }
}

/**
 * رحلات حسب الفترة
 */
function getFlightsByPeriod_(period) {
  var flights = getFlightsData_();
  var today = getTodayString_();
  var weekEnd = getDateOffset_(today, 7);
  var results = [];

  for (var i = 0; i < flights.length; i++) {
    var row = flights[i];
    if (!row[FLT.PNR]) continue;

    var goDate = normDate_(row[FLT.GO1_TAKEOFF_DATE]);
    var retDate = normDate_(row[FLT.RET1_TAKEOFF_DATE]);

    var include = false;
    if (period === 'today') {
      include = (goDate === today || retDate === today);
    } else if (period === 'week') {
      include = (goDate >= today && goDate <= weekEnd) || (retDate >= today && retDate <= weekEnd);
    } else {
      include = true;
    }

    if (include) {
      results.push({
        pnr: row[FLT.PNR],
        supplier: row[FLT.SUPPLIER],
        status: row[FLT.STATUS],
        country: row[FLT.COUNTRY],
        city: row[FLT.CITY],
        airline: row[FLT.AIRLINE],
        pax: row[FLT.PAX],
        sales: row[FLT.SALES],
        remaining: row[FLT.REMAINING],
        go1: {
          flight: row[FLT.GO1_FLIGHT],
          date: goDate,
          time: row[FLT.GO1_TAKEOFF_TIME],
          from: row[FLT.GO1_FROM],
          to: row[FLT.GO1_TO],
          landDate: normDate_(row[FLT.GO1_LAND_DATE]),
          landTime: row[FLT.GO1_LAND_TIME]
        },
        ret1: {
          flight: row[FLT.RET1_FLIGHT],
          date: retDate,
          time: row[FLT.RET1_TAKEOFF_TIME],
          from: row[FLT.RET1_FROM],
          to: row[FLT.RET1_TO],
          landDate: normDate_(row[FLT.RET1_LAND_DATE]),
          landTime: row[FLT.RET1_LAND_TIME]
        }
      });
    }
  }

  return results;
}

// ============================================
// لوحة القيادة — إحصائيات مجمّعة
// ============================================

function getDashboardStats_() {
  var cached = getCache_('adm_dash');
  if (cached) return cached;

  // عدد الحجاج من PD (الأدق)
  var pd = getPersonalDetails_();
  var totalPilgrims = pd.filter(function(r) { return r[PD.PASSPORT]; }).length;

  // الباقات
  var pkgStats = getPackagesStats_();

  // الرحلات
  var flights = getFlightsData_();
  var totalFlights = flights.length;

  // وصول/مغادرة اليوم (من رحلة الحاج 2)
  var j2 = getJourney2Data_();
  var today = getTodayString_();
  var arrivalsToday = 0, departuresToday = 0;

  for (var i = 0; i < j2.length; i++) {
    if (normDate_(j2[i][J2.ARR_ARRIVE_DATE]) === today) arrivalsToday++;
    if (normDate_(j2[i][J2.RET_DEPART_DATE]) === today) departuresToday++;
  }

  // حالة الاستقبال
  var journey = getJourneyData_();
  var received = 0;
  for (var k = 0; k < journey.length; k++) {
    if (journey[k][JRN.RECEPTION_STATUS]) received++;
  }

  var stats = {
    totalPilgrims: totalPilgrims,
    totalPackages: pkgStats.packages.length,
    totalFlights: totalFlights,
    salesPercent: pkgStats.totalPercent,
    totalSold: pkgStats.totalSold,
    totalCapacity: pkgStats.totalCapacity,
    arrivalsToday: arrivalsToday,
    departuresToday: departuresToday,
    received: received,
    countries: countUniqueCountries_(flights)
  };

  setCache_('adm_dash', stats, 300); // كاش 5 دقائق للوحة القيادة
  return stats;
}

/**
 * عدد الدول الفريدة
 */
function countUniqueCountries_(flights) {
  var countries = {};
  for (var i = 0; i < flights.length; i++) {
    var c = String(flights[i][FLT.COUNTRY] || '').trim();
    if (c) countries[c] = true;
  }
  return Object.keys(countries).length;
}

// ============================================
// إشغال الفنادق (مجمّع من الباقات)
// ============================================

function getHotelOccupancy_() {
  var pkgs = getPackagesData_();
  var hotels = {};

  for (var i = 0; i < pkgs.length; i++) {
    var row = pkgs[i];
    if (!row[PKG.NAME_AR]) continue;

    var sold = Number(row[PKG.SALES]) || 0;
    var capacity = Number(row[PKG.NO_PILGRIM]) || 0;

    // 3 فنادق لكل باقة
    var hotelSets = [
      { name: row[PKG.H1_NAME_AR], nameEn: row[PKG.H1_NAME_EN], city: row[PKG.H1_CITY], rooms: row[PKG.H1_ROOMS], beds: row[PKG.H1_BEDS] },
      { name: row[PKG.H2_NAME_AR], nameEn: row[PKG.H2_NAME_EN], city: row[PKG.H2_CITY], rooms: row[PKG.H2_ROOMS], beds: row[PKG.H2_BEDS] },
      { name: row[PKG.H3_NAME_AR], nameEn: row[PKG.H3_NAME_EN], city: row[PKG.H3_CITY], rooms: row[PKG.H3_ROOMS], beds: row[PKG.H3_BEDS] }
    ];

    for (var h = 0; h < hotelSets.length; h++) {
      var ht = hotelSets[h];
      var hName = String(ht.name || '').trim();
      if (!hName || hName === '-' || hName === '0') continue;

      var key = hName;
      if (!hotels[key]) {
        hotels[key] = {
          name: hName,
          nameEn: ht.nameEn || '',
          city: ht.city || '',
          totalRooms: 0,
          totalBeds: 0,
          packages: []
        };
      }
      hotels[key].totalRooms += Number(ht.rooms) || 0;
      hotels[key].totalBeds += Number(ht.beds) || 0;
      hotels[key].packages.push({
        name: row[PKG.NAME_AR],
        sold: sold,
        capacity: capacity
      });
    }
  }

  return Object.keys(hotels).map(function(k) { return hotels[k]; });
}

// ============================================
// المخيمات (من Presonal Details عمود AB)
// ============================================

function getCampStats_() {
  var pd = getPersonalDetails_();
  var camps = {};

  for (var i = 0; i < pd.length; i++) {
    var camp = String(pd[i][PD.CAMP] || '').trim();
    if (!camp) continue;

    if (!camps[camp]) {
      camps[camp] = { name: camp, count: 0 };
    }
    camps[camp].count++;
  }

  var result = Object.keys(camps).map(function(k) { return camps[k]; });
  result.sort(function(a, b) { return b.count - a.count; });
  return result;
}

// ============================================
// المبيعات (مجمّع من الباقات + الطيران)
// ============================================

function getSalesStats_() {
  var pkgStats = getPackagesStats_();
  var flights = getFlightsData_();

  // مبيعات حسب الدولة
  var byCountry = {};
  for (var i = 0; i < flights.length; i++) {
    var country = String(flights[i][FLT.COUNTRY] || '').trim();
    var pax = Number(flights[i][FLT.PAX]) || 0;
    var sold = Number(flights[i][FLT.SALES]) || 0;
    if (!country) continue;

    if (!byCountry[country]) {
      byCountry[country] = { country: country, pax: 0, sold: 0, flights: 0 };
    }
    byCountry[country].pax += pax;
    byCountry[country].sold += sold;
    byCountry[country].flights++;
  }

  var countries = Object.keys(byCountry).map(function(k) { return byCountry[k]; });
  countries.sort(function(a, b) { return b.sold - a.sold; });

  return {
    packages: pkgStats,
    byCountry: countries,
    totalFlights: flights.length
  };
}
