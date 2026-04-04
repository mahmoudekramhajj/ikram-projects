/**
 * ══════════════════════════════════════════════════════════════
 * Ikram Hajj - Sales & Operations Report v3.0
 * ══════════════════════════════════════════════════════════════
 *
 * المصادر:
 *  - شيت الطيران: المبيعات + الرحلات التفصيلية + الباقات المرتبطة
 *  - شيت رحلة الحاج: بلد الإقامة (cached 5 دقائق)
 *  - شيت الباقات: City Of Start + التنقل (cached 5 دقائق)
 *  - شيت الباقات + رحلة الحاج: تفاصيل كاملة للتصدير (عند الطلب)
 *
 * ══════════════════════════════════════════════════════════════
 */

var REPORT_CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  FLIGHTS_SHEET: 'الطيران',
  PACKAGES_SHEET: 'الباقات',
  JOURNEY_SHEET: 'رحلة الحاج ',
  DATA_START: 3,
  CACHE_TTL: 300
};

// ══════════════════════════════════════
// Web App Entry Point
// ══════════════════════════════════════

function doGet(e) {
  return HtmlService.createTemplateFromFile('SalesIndex')
    .evaluate()
    .setTitle('Ikram Sales Report')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ══════════════════════════════════════
// Data Fetching
// ══════════════════════════════════════

function getSalesData() {
  try {
    var ss = SpreadsheetApp.openById(REPORT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(REPORT_CONFIG.FLIGHTS_SHEET);
    if (!sheet) return JSON.stringify({ success: false, error: 'Sheet not found' });

    var lastRow = sheet.getLastRow();
    if (lastRow < REPORT_CONFIG.DATA_START) return JSON.stringify({ success: false, error: 'No data' });

    var numRows = lastRow - REPORT_CONFIG.DATA_START + 1;
    var data = sheet.getRange(REPORT_CONFIG.DATA_START, 1, numRows, 92).getValues();

    // ─── بيانات مرجعية (cached) ───
    var residenceByPkg = getCached(ss, 'resByFlight', buildResidenceMap);
    var guideByContract = getCached(ss, 'guideMap', buildGuideMap);
    var pkgInfoByNusk = getCached(ss, 'pkgInfoMap', buildPackageInfoMap);

    var flights = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      var pnr = row[1];
      var country = row[4];
      var pax = toNum(row[7]);

      if (!(pnr && String(pnr).trim() !== '') && !(country && String(country).trim() !== '') && !(pax > 0)) continue;

      var sold = toNum(row[86]);
      var remaining = toNum(row[87]);
      var pctRaw = row[88];

      if (remaining === 0 && pax > 0 && sold >= 0) remaining = pax - sold;

      var salesPct = 0;
      if (pctRaw) {
        salesPct = parseFloat(pctRaw);
        if (!isNaN(salesPct) && salesPct > 0 && salesPct <= 1) salesPct = salesPct * 100;
        if (isNaN(salesPct)) salesPct = 0;
      }
      if (salesPct === 0 && pax > 0 && sold > 0) salesPct = (sold / pax) * 100;

      var pkgs = extractPackages(row);

      // ─── بناء الرحلات أولاً ───
      var dep1 = buildLeg(row, 21);
      var dep2 = buildLeg(row, 28);
      var ret1 = buildLeg(row, 35);
      var ret2 = buildLeg(row, 42);

      // ─── بلدان الإقامة (حسب رقم رحلة الذهاب الفعلي) مع الأعداد ───
      var resCounts = {};
      var depFlights = [];
      if (dep1) depFlights.push(dep1.flight);
      if (dep2) depFlights.push(dep2.flight);
      for (var fi = 0; fi < depFlights.length; fi++) {
        var fnCounts = residenceByPkg[depFlights[fi]] || {};
        for (var rc in fnCounts) {
          resCounts[rc] = (resCounts[rc] || 0) + fnCounts[rc];
        }
      }
      var resList = Object.keys(resCounts);

      // ─── أسماء المرشدين (ربط عبر أكواد PNR داخل اسم العقد) ───
      var pnrRaw = clean(row[1]);
      var pnrCodes = pnrRaw ? pnrRaw.split(/\s*-\s*/) : [];
      var guideCounts = {};
      for (var gi = 0; gi < pnrCodes.length; gi++) {
        var pCode = pnrCodes[gi].trim();
        if (!pCode) continue;
        var gcForCode = guideByContract[pCode] || {};
        for (var gn in gcForCode) {
          guideCounts[gn] = (guideCounts[gn] || 0) + gcForCode[gn];
        }
      }
      var guideList = Object.keys(guideCounts);

      // ─── معلومات الباقة (مكة/مدينة أولاً) ───
      var cityStart = '', transport = '';
      for (var pi = 0; pi < pkgs.length; pi++) {
        var pInfo = pkgInfoByNusk[pkgs[pi].id];
        if (pInfo && !cityStart) { cityStart = pInfo.city; transport = pInfo.transport; }
      }

      flights.push({
        no: row[0],
        pnr: clean(pnr) || ('ROW-' + (i + REPORT_CONFIG.DATA_START)),
        supplier: clean(row[2]),
        status: clean(row[3]),
        country: clean(row[4]),
        city: clean(row[5]),
        airline: clean(row[6]),
        pax: pax,
        sold: sold,
        remaining: remaining,
        salesPct: Math.round(salesPct * 10) / 10,
        contractNo: clean(row[90]),
        packages: pkgs,
        residenceCountries: resList,
        residenceCounts: resCounts,
        guides: guideList,
        guideCounts: guideCounts,
        cityStart: cityStart,
        transport: transport,
        dep1: dep1,
        dep2: dep2,
        ret1: ret1,
        ret2: ret2
      });
    }

    var report = generateReport(flights);

    return JSON.stringify({
      success: true,
      flights: flights,
      report: report,
      totalRowsRead: numRows,
      flightsFound: flights.length,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    Logger.log('getSalesData ERROR: ' + error.toString());
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

/**
 * بناء بيانات leg واحدة (ذهاب أو إياب)
 */
function buildLeg(row, startIdx) {
  var flight = clean(row[startIdx]);
  if (!flight) return null;
  return {
    flight: flight,
    date: formatDate(row[startIdx + 1]),
    time: clean(row[startIdx + 2]),
    from: clean(row[startIdx + 3]),
    to: clean(row[startIdx + 4]),
    arrDate: formatDate(row[startIdx + 5]),
    arrTime: clean(row[startIdx + 6])
  };
}

// ══════════════════════════════════════
// Generic Cache Layer
// ══════════════════════════════════════

function getCached(ss, key, builder) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(key);
    if (cached) return JSON.parse(cached);

    var map = builder(ss);
    var str = JSON.stringify(map);
    if (str.length < 100000) cache.put(key, str, REPORT_CONFIG.CACHE_TTL);
    return map;
  } catch (e) {
    Logger.log('Cache ' + key + ' error: ' + e.toString());
    try { return builder(ss); } catch (e2) { return {}; }
  }
}

// ══════════════════════════════════════
// Residence Country Map (per-FlightNumber from رحلة الحاج)
// ══════════════════════════════════════

function buildResidenceMap(ss) {
  var map = {}; // flightNumber → [residenceCountries]
  var sheet = ss.getSheetByName(REPORT_CONFIG.JOURNEY_SHEET);
  if (!sheet) return map;
  var n = sheet.getLastRow() - 1;
  if (n < 1) return map;

  // عمود Y (25 بترقيم 1) = ArrivalFlightNumber
  // عمود O (15 بترقيم 1) = CountryResidenceEn
  var colFN = sheet.getRange(2, 25, n, 1).getValues();
  var colRes = sheet.getRange(2, 15, n, 1).getValues();

  for (var i = 0; i < n; i++) {
    var fn = clean(colFN[i][0]);
    var res = clean(colRes[i][0]);
    if (!fn || !res) continue;
    if (!map[fn]) map[fn] = {};
    map[fn][res] = (map[fn][res] || 0) + 1;
  }
  return map;
}

// ══════════════════════════════════════
// Guide Map — ربط عبر أكواد PNR داخل اسم العقد
// Guide Rabih col V (22) = اسم العقد يحتوي أكواد PNR
// الطيران col B (2) = PNR (أكواد مفصولة بـ " - ")
// ══════════════════════════════════════

var NO_GUIDE = 'بدون مرشد';

function buildGuideMap(ss) {
  var map = {}; // pnrCode → {guideName: count}

  var grSheet = ss.getSheetByName('Guide Rabih');
  if (!grSheet) return map;
  var grN = grSheet.getLastRow() - 1;
  if (grN < 1) return map;

  var grContract = grSheet.getRange(2, 22, grN, 1).getValues(); // V = اسم العقد
  var grGuide    = grSheet.getRange(2, 16, grN, 1).getValues(); // P = اسم المرشد

  // مرحلة 1: تجميع حسب اسم العقد الكامل
  var byContract = {}; // contractName → {guide: count}
  for (var i = 0; i < grN; i++) {
    var contract = clean(grContract[i][0]);
    var gd       = clean(grGuide[i][0]) || NO_GUIDE;
    if (!contract) continue;
    if (!byContract[contract]) byContract[contract] = {};
    byContract[contract][gd] = (byContract[contract][gd] || 0) + 1;
  }

  // مرحلة 2: استخراج أكواد PNR من اسم العقد وبناء map بكل كود
  // مثال: "Wizz MXP 17May 03Jun 30 HSISPR  CNGJ6A" → أكواد: HSISPR, CNGJ6A
  // أكواد PNR: 4-8 حروف كبيرة/أرقام، عادة في نهاية النص
  var pnrPattern = /\b([A-Z0-9]{4,8})\b/g;
  for (var contractName in byContract) {
    // استخراج كل الأكواد الممكنة من آخر جزء بعد الأرقام
    var match;
    var codes = [];
    pnrPattern.lastIndex = 0;
    while ((match = pnrPattern.exec(contractName)) !== null) {
      var code = match[1];
      // تجاهل الأكواد التي تبدو كتواريخ أو أرقام PAX
      if (/^\d+$/.test(code)) continue;        // أرقام فقط
      if (/^\d{1,2}[A-Z][a-z]{2}$/.test(code)) continue; // تاريخ مثل 17May
      codes.push(code);
    }
    // أيضاً: عقود NUSUK — استخدم الاسم الكامل كمفتاح
    if (contractName.indexOf('NUSUK') === 0) {
      codes.push(contractName);
    }
    for (var ci = 0; ci < codes.length; ci++) {
      if (!map[codes[ci]]) map[codes[ci]] = {};
      var guides = byContract[contractName];
      for (var gn in guides) {
        map[codes[ci]][gn] = (map[codes[ci]][gn] || 0) + guides[gn];
      }
    }
  }
  return map;
}

// ══════════════════════════════════════
// Guide Report (guide-centric view — called on demand)
// ══════════════════════════════════════

function getGuideReport(guideName) {
  try {
    var ss = SpreadsheetApp.openById(REPORT_CONFIG.SPREADSHEET_ID);

    // 1) Guide Rabih → passport → guide + details
    var grSheet = ss.getSheetByName('Guide Rabih');
    if (!grSheet) return JSON.stringify({ success: false, error: 'Guide Rabih not found' });
    var grLastRow = grSheet.getLastRow();
    if (grLastRow < 2) return JSON.stringify({ success: false, error: 'No data' });
    var grData = grSheet.getRange(2, 1, grLastRow - 1, 20).getValues();
    // Columns: 0=serial, 4=gender, 5=passport, 6=passportExpiry, 10=firstNameEN, 11=familyNameEN,
    //          12=dob, 15=guideName, 16=residence, 17=nationality, 18=pkgNo, 19=pkgName

    // فلتر حسب المرشد — مطابقة كاملة أولاً، ثم جزئية للأسماء المركبة
    var matchPassports = {}; // passport → {guide, firstName, familyName, gender, dob, passportExpiry, nationality, residence, pkgNo, pkgName}
    for (var i = 0; i < grData.length; i++) {
      var rawGuide = clean(grData[i][15]);
      var guideMatch = false;
      if (guideName === NO_GUIDE) {
        guideMatch = !rawGuide;
      } else if (rawGuide === guideName) {
        // مطابقة كاملة: "شكري سعدي - زكريا كارنو فوياج" === "شكري سعدي - زكريا كارنو فوياج"
        guideMatch = true;
      } else {
        // مطابقة جزئية: "شكري سعدي" يطابق "شكري سعدي - زكريا كارنو فوياج"
        var parts = rawGuide.split(/\s*-\s*/);
        for (var p = 0; p < parts.length; p++) {
          if (parts[p].trim() === guideName) { guideMatch = true; break; }
        }
      }
      if (!guideMatch) continue;

      var pp = clean(grData[i][5]);
      if (!pp) continue;
      matchPassports[pp] = {
        guide: rawGuide || NO_GUIDE,
        firstName: clean(grData[i][10]),
        familyName: clean(grData[i][11]),
        gender: clean(grData[i][4]),
        dob: formatDate(grData[i][12]),
        passportExpiry: formatDate(grData[i][6]),
        nationality: clean(grData[i][17]),
        residence: clean(grData[i][16]),
        pkgNo: clean(grData[i][18]),
        pkgName: clean(grData[i][19])
      };
    }

    // 2) رحلة الحاج → ربط الجواز بالرحلة والفنادق
    var jSheet = ss.getSheetByName(REPORT_CONFIG.JOURNEY_SHEET);
    var pilgrims = [];
    if (jSheet) {
      var jN = jSheet.getLastRow() - 1;
      if (jN > 0) {
        // نقرأ أعمدة: B(2)=PkgId, H(8)=Name, I(9)=Passport, V(22)=DepCity, W(23)=DepDate, X(24)=DepTime, Y(25)=FlightNo
        //              AK(37)=FirstHouse, AL(38)=FirstHouseStart, AN(40)=LastHouse
        var jData = jSheet.getRange(2, 1, jN, 48).getValues();
        for (var j = 0; j < jN; j++) {
          var pp2 = clean(jData[j][8]); // Passport
          if (!pp2 || !matchPassports[pp2]) continue;
          var pd = matchPassports[pp2];
          pilgrims.push({
            name: pd.firstName + ' ' + pd.familyName,
            firstName: pd.firstName,
            familyName: pd.familyName,
            passport: pp2,
            gender: pd.gender,
            dob: pd.dob,
            passportExpiry: pd.passportExpiry,
            nationality: pd.nationality,
            residence: pd.residence,
            pkgNo: pd.pkgNo || clean(jData[j][1]),
            pkgName: pd.pkgName,
            flightNo: clean(jData[j][24]),
            depCity: clean(jData[j][21]),
            depDate: formatDate(jData[j][22]),
            depTime: clean(jData[j][23]),
            firstHotel: clean(jData[j][36]),
            firstHotelStart: formatDate(jData[j][37]),
            lastHotel: clean(jData[j][39]),
            lastHotelStart: formatDate(jData[j][40])
          });
        }
      }
    }

    // 3) الباقات → تفاصيل إضافية
    var pkgInfoByNusk = getCached(ss, 'pkgInfoMap', buildPackageInfoMap);

    // 4) إحصائيات
    var stats = { packages: {}, flights: {}, countries: {}, residences: {} };
    for (var k = 0; k < pilgrims.length; k++) {
      var pl = pilgrims[k];
      if (pl.pkgNo) stats.packages[pl.pkgNo] = pl.pkgName || pl.pkgNo;
      if (pl.flightNo) stats.flights[pl.flightNo] = { depCity: pl.depCity, depDate: pl.depDate };
      if (pl.nationality) stats.countries[pl.nationality] = (stats.countries[pl.nationality] || 0) + 1;
      if (pl.residence) stats.residences[pl.residence] = (stats.residences[pl.residence] || 0) + 1;
    }

    return JSON.stringify({
      success: true,
      guide: guideName,
      totalPilgrims: pilgrims.length,
      totalPackages: Object.keys(stats.packages).length,
      totalFlights: Object.keys(stats.flights).length,
      pilgrims: pilgrims,
      stats: stats,
      pkgInfo: pkgInfoByNusk
    });
  } catch (error) {
    Logger.log('getGuideReport ERROR: ' + error.toString());
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// ══════════════════════════════════════
// Package Info Map (City Of Start, Transport)
// ══════════════════════════════════════

function buildPackageInfoMap(ss) {
  var map = {};
  var sheet = ss.getSheetByName(REPORT_CONFIG.PACKAGES_SHEET);
  if (!sheet) return map;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return map;

  var n = lastRow - 2;
  // B(1)=NuskNo, J(9+1=10)=CityOfStart, BN(65+1=66)=Transport
  var colB = sheet.getRange(3, 2, n, 1).getValues();
  var colJ = sheet.getRange(3, 10, n, 1).getValues();
  var colBN = sheet.getRange(3, 66, n, 1).getValues();

  for (var i = 0; i < n; i++) {
    var nusk = clean(colB[i][0]);
    if (!nusk) continue;
    map[nusk] = {
      city: clean(colJ[i][0]),
      transport: clean(colBN[i][0])
    };
  }
  return map;
}

// ══════════════════════════════════════
// Package Extraction
// ══════════════════════════════════════

function extractPackages(row) {
  var packages = [];
  for (var idx = 49; idx <= 67; idx += 2) {
    var id = clean(row[idx]);
    var name = clean(row[idx + 1]);
    if (id || name) packages.push({ id: id, name: name || id });
  }
  return packages;
}

// ══════════════════════════════════════
// Export Data (Packages + Pilgrims)
// ══════════════════════════════════════

function getExportData(inputJson) {
  try {
    var input = JSON.parse(inputJson);
    // دعم الشكل القديم (مصفوفة) والجديد (كائن)
    var packageIds = Array.isArray(input) ? input : (input.pkgIds || []);
    var flightNos = input.flightNos || [];
    var ss = SpreadsheetApp.openById(REPORT_CONFIG.SPREADSHEET_ID);

    // ─── 1) بيانات الباقات ───
    var pkgSheet = ss.getSheetByName(REPORT_CONFIG.PACKAGES_SHEET);
    var packagesMap = {};
    if (pkgSheet) {
      var pkgLastRow = pkgSheet.getLastRow();
      if (pkgLastRow >= 3) {
        var pkgData = pkgSheet.getRange(3, 1, pkgLastRow - 2, 69).getValues();
        for (var i = 0; i < pkgData.length; i++) {
          var nuskNo = clean(pkgData[i][1]);
          if (!nuskNo) continue;
          packagesMap[nuskNo] = {
            nuskNo: nuskNo,
            nameAr: clean(pkgData[i][2]),
            nameEn: clean(pkgData[i][60]),
            startDate: formatDate(pkgData[i][6]),
            endDate: formatDate(pkgData[i][7]),
            capacity: toNum(pkgData[i][10]),
            sold: toNum(pkgData[i][57]),
            cityStart: clean(pkgData[i][9]),
            transport: clean(pkgData[i][65])
          };
        }
      }
    }

    // ─── 2) Guide Rabih — نفس هيكل PD + اسم المرشد ───
    var grSheet = ss.getSheetByName('Guide Rabih');
    var pdMap = {}; // passport → {firstName, familyName, dob, passportExpiry, gender, nationality, guide}
    if (grSheet) {
      var grLastRow = grSheet.getLastRow();
      if (grLastRow >= 2) {
        // أعمدة: 4=الجنس, 5=جواز, 6=انتهاء الجواز, 10=الاسم الأول EN, 11=اسم العائلة EN, 12=تاريخ الميلاد, 15=اسم المرشد, 17=الجنسية
        var grData = grSheet.getRange(2, 1, grLastRow - 1, 18).getValues();
        for (var p = 0; p < grData.length; p++) {
          var pp = clean(grData[p][5]); // رقم الجواز
          if (!pp) continue;
          pdMap[pp] = {
            firstName: clean(grData[p][10]),
            familyName: clean(grData[p][11]),
            gender: clean(grData[p][4]),
            nationality: clean(grData[p][17]),
            dob: formatDate(grData[p][12]),
            passportExpiry: formatDate(grData[p][6]),
            guide: clean(grData[p][15]) || 'بدون مرشد'
          };
        }
      }
    }

    // ─── 3) بيانات الحجاج — مرتبطة برقم الرحلة الفعلي + مثراة من PD ───
    var journeySheet = ss.getSheetByName(REPORT_CONFIG.JOURNEY_SHEET);
    var pilgrimsMap = {};
    if (journeySheet) {
      var jLastRow = journeySheet.getLastRow();
      if (jLastRow >= 2) {
        var jData = journeySheet.getRange(2, 1, jLastRow - 1, 25).getValues();

        var fnSet = {};
        for (var k = 0; k < flightNos.length; k++) fnSet[String(flightNos[k])] = true;

        for (var j = 0; j < jData.length; j++) {
          var pkgId = clean(jData[j][1]);
          var flightNo = clean(jData[j][24]);
          if (!pkgId) continue;
          if (flightNos.length > 0 && !fnSet[flightNo]) continue;

          if (!pilgrimsMap[pkgId]) pilgrimsMap[pkgId] = [];
          var name = clean(jData[j][7]);
          var passport = clean(jData[j][8]);
          var residence = clean(jData[j][14]);
          var gender = clean(jData[j][11]);
          var nationality = clean(jData[j][12]);

          if (name || passport) {
            var pd = pdMap[passport] || {};
            pilgrimsMap[pkgId].push({
              name: name,
              firstName: pd.firstName || name.split(' ')[0] || '',
              familyName: pd.familyName || name.split(' ').slice(1).join(' ') || '',
              passport: passport,
              residence: residence,
              flightNo: flightNo,
              gender: gender || pd.gender || '',
              nationality: nationality || pd.nationality || '',
              dob: pd.dob || '',
              passportExpiry: pd.passportExpiry || '',
              guide: pd.guide || ''
            });
          }
        }
      }
    }

    return JSON.stringify({ success: true, packages: packagesMap, pilgrims: pilgrimsMap });
  } catch (error) {
    Logger.log('getExportData ERROR: ' + error.toString());
    return JSON.stringify({ success: false, error: error.toString() });
  }
}

// ══════════════════════════════════════
// Report Generation
// ══════════════════════════════════════

function generateReport(flights) {
  var byCountry = {}, byAirline = {}, bySupplier = {}, byPackage = {};
  var totals = { pnrs: 0, pax: 0, sold: 0, remaining: 0 };

  for (var i = 0; i < flights.length; i++) {
    var f = flights[i];
    totals.pnrs++; totals.pax += f.pax; totals.sold += f.sold; totals.remaining += f.remaining;

    agg(byCountry, f.country || 'Unknown', f);
    agg(byAirline, f.airline || 'Unknown', f);

    var s = f.supplier || 'Unknown';
    if (!bySupplier[s]) bySupplier[s] = { name: s, pnrs: 0, pax: 0, sold: 0, remaining: 0, countries: [] };
    bySupplier[s].pnrs++; bySupplier[s].pax += f.pax; bySupplier[s].sold += f.sold; bySupplier[s].remaining += f.remaining;
    if (f.country && bySupplier[s].countries.indexOf(f.country) === -1) bySupplier[s].countries.push(f.country);

    for (var p = 0; p < f.packages.length; p++) {
      var pkg = f.packages[p], pKey = pkg.name || pkg.id || 'Unknown';
      if (!byPackage[pKey]) byPackage[pKey] = { name: pKey, id: pkg.id, pnrs: 0, pax: 0, sold: 0, remaining: 0, countries: [] };
      byPackage[pKey].pnrs++; byPackage[pKey].pax += f.pax; byPackage[pKey].sold += f.sold; byPackage[pKey].remaining += f.remaining;
      if (f.country && byPackage[pKey].countries.indexOf(f.country) === -1) byPackage[pKey].countries.push(f.country);
    }
  }

  totals.salesPct = totals.pax > 0 ? Math.round((totals.sold / totals.pax) * 1000) / 10 : 0;

  return {
    totals: totals,
    countries: toArr(byCountry),
    airlines: toArr(byAirline),
    suppliers: toArr(bySupplier),
    packages: toArr(byPackage)
  };
}

function agg(obj, key, f) {
  if (!obj[key]) obj[key] = { name: key, pnrs: 0, pax: 0, sold: 0, remaining: 0 };
  obj[key].pnrs++; obj[key].pax += f.pax; obj[key].sold += f.sold; obj[key].remaining += f.remaining;
}

function toArr(obj) {
  var arr = [];
  for (var k in obj) { var it = obj[k]; it.salesPct = it.pax > 0 ? Math.round((it.sold / it.pax) * 1000) / 10 : 0; arr.push(it); }
  return arr.sort(function(a, b) { return b.pax - a.pax; });
}

// ══════════════════════════════════════
// Helpers
// ══════════════════════════════════════

function clean(val) { if (!val) return ''; return String(val).trim(); }

function toNum(val) { if (!val && val !== 0) return 0; var n = Number(val); return isNaN(n) ? 0 : Math.round(n); }

function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var d = val.getDate(), m = val.getMonth() + 1, y = val.getFullYear();
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
  }
  return String(val).trim();
}
