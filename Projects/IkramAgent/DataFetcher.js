/**
 * DataFetcher.gs - جلب البيانات لوكيل إكرام
 * الإصدار 2.1 - مُصحح بالكامل
 * 
 * هيكل البيانات:
 * - شيت الطيران: العمود E = الدولة، العمود F = المدينة
 * - أرقام نسك في الطيران: الأعمدة 50, 52, 54, 56, 58, 60, 62, 64, 66, 68
 * - شيت الباقات: العمود B = رقم نسك، العمود C = الاسم
 */

// ═══════════════════════════════════════════════════════════
// جلب الدول المتاحة (التي لها باقات فعلاً)
// ═══════════════════════════════════════════════════════════

function getAvailableCountries() {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var flightsSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.FLIGHTS);
    var packagesSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    
    if (!flightsSheet || !packagesSheet) return [];
    
    // 1. جلب أرقام نسك من الباقات المتاحة
    var packagesLastRow = packagesSheet.getLastRow();
    if (packagesLastRow < AGENT_CONFIG.DATA_START_ROW) return [];
    
    var packagesData = packagesSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1, 
      packagesLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var availableNusks = {};
    for (var i = 0; i < packagesData.length; i++) {
      var nuskNo = packagesData[i][1]; // العمود B
      var capacity = Number(packagesData[i][10]) || 0; // العمود K
      var salesCount = Number(packagesData[i][57]) || 0; // العمود BF
      var remaining = capacity - salesCount;
      
      if (nuskNo && remaining > 0) {
        availableNusks[Math.floor(Number(nuskNo))] = true;
      }
    }
    
    // 2. جلب الدول من الطيران وربطها بالباقات
    var flightsLastRow = flightsSheet.getLastRow();
    if (flightsLastRow < AGENT_CONFIG.DATA_START_ROW) return [];
    
    var flightsData = flightsSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      flightsLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 70
    ).getValues();
    
    var countriesWithPackages = {};
    var nuskColumns = [49, 51, 53, 55, 57, 59, 61, 63, 65, 67]; // 0-indexed
    
    for (var j = 0; j < flightsData.length; j++) {
      var country = String(flightsData[j][4]).trim(); // العمود E (index 4)
      if (!country) continue;
      
      // التحقق من وجود باقات متاحة لهذه الرحلة
      for (var k = 0; k < nuskColumns.length; k++) {
        var nuskVal = flightsData[j][nuskColumns[k]];
        if (nuskVal) {
          var nuskNum = Math.floor(Number(nuskVal));
          if (availableNusks[nuskNum]) {
            countriesWithPackages[country] = true;
            break;
          }
        }
      }
    }
    
    // 3. إضافة الأعلام
    var countryFlags = {
      'Italy': '🇮🇹', 'Australia': '🇦🇺', 'Canada': '🇨🇦', 'USA': '🇺🇸',
      'Switzerland': '🇨🇭', 'Austria': '🇦🇹', 'Germany': '🇩🇪',
      'Netherland': '🇳🇱', 'Netherlands': '🇳🇱', 'France': '🇫🇷',
      'UK': '🇬🇧', 'Spain': '🇪🇸', 'Denmark': '🇩🇰', 'Belgium': '🇧🇪',
      'Sweden': '🇸🇪', 'South Africa': '🇿🇦', 'Greece': '🇬🇷',
      'Japan': '🇯🇵', 'Finland': '🇫🇮', 'Norway': '🇳🇴'
    };
    
    var result = [];
    for (var c in countriesWithPackages) {
      var flag = countryFlags[c] || '🌍';
      result.push(flag + ' ' + c);
    }
    
    return result.sort();
    
  } catch (e) {
    Logger.log('getAvailableCountries Error: ' + e.toString());
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
// جلب الفنادق المتاحة
// ═══════════════════════════════════════════════════════════

function getAvailableHotels() {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < AGENT_CONFIG.DATA_START_ROW) return [];
    
    var data = sheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      lastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var hotels = {};
    
    for (var i = 0; i < data.length; i++) {
      var capacity = Number(data[i][10]) || 0;
      var salesCount = Number(data[i][57]) || 0;
      var remaining = capacity - salesCount;
      
      if (remaining <= 0) continue;
      
      // فندق 1: العمود L (11) = المدينة، العمود M (12) = الاسم
      var hotel1City = String(data[i][11]).trim();
      var hotel1Name = String(data[i][12]).trim();
      if (hotel1Name && hotel1Name !== '') {
        var icon1 = hotel1City === 'Mak' ? '🕋' : '🕌';
        hotels[icon1 + ' ' + hotel1Name] = true;
      }
      
      // فندق 2: العمود AA (26) = المدينة، العمود AB (27) = الاسم
      var hotel2City = String(data[i][26]).trim();
      var hotel2Name = String(data[i][27]).trim();
      if (hotel2Name && hotel2Name !== '') {
        var icon2 = hotel2City === 'Mak' ? '🕋' : '🕌';
        hotels[icon2 + ' ' + hotel2Name] = true;
      }
    }
    
    var result = [];
    for (var h in hotels) {
      result.push(h);
    }
    
    return result.sort();
    
  } catch (e) {
    Logger.log('getAvailableHotels Error: ' + e.toString());
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
// جلب مدن الطيران المتاحة
// ═══════════════════════════════════════════════════════════

function getAvailableCities() {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var flightsSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.FLIGHTS);
    var packagesSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    
    if (!flightsSheet || !packagesSheet) return [];
    
    // 1. جلب أرقام نسك من الباقات المتاحة
    var packagesLastRow = packagesSheet.getLastRow();
    var packagesData = packagesSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      packagesLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var availableNusks = {};
    for (var i = 0; i < packagesData.length; i++) {
      var nuskNo = packagesData[i][1];
      var capacity = Number(packagesData[i][10]) || 0;
      var salesCount = Number(packagesData[i][57]) || 0;
      if (nuskNo && (capacity - salesCount) > 0) {
        availableNusks[Math.floor(Number(nuskNo))] = true;
      }
    }
    
    // 2. جلب المدن من الطيران
    var flightsLastRow = flightsSheet.getLastRow();
    var flightsData = flightsSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      flightsLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 70
    ).getValues();
    
    var citiesWithPackages = {};
    var nuskColumns = [49, 51, 53, 55, 57, 59, 61, 63, 65, 67];
    
    for (var j = 0; j < flightsData.length; j++) {
      var city = String(flightsData[j][5]).trim(); // العمود F (index 5)
      if (!city) continue;
      
      for (var k = 0; k < nuskColumns.length; k++) {
        var nuskVal = flightsData[j][nuskColumns[k]];
        if (nuskVal) {
          var nuskNum = Math.floor(Number(nuskVal));
          if (availableNusks[nuskNum]) {
            citiesWithPackages[city] = true;
            break;
          }
        }
      }
    }
    
    var result = [];
    for (var c in citiesWithPackages) {
      result.push('✈️ ' + c);
    }
    
    return result.sort();
    
  } catch (e) {
    Logger.log('getAvailableCities Error: ' + e.toString());
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
// جلب الباقات حسب الدولة
// ═══════════════════════════════════════════════════════════

function getPackagesByCountry(country) {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var flightsSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.FLIGHTS);
    var packagesSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    
    if (!flightsSheet || !packagesSheet) return [];
    
    // 1. جلب أرقام نسك المرتبطة بالدولة من شيت الطيران
    var flightsLastRow = flightsSheet.getLastRow();
    var flightsData = flightsSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      flightsLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 70
    ).getValues();
    
    var linkedNusks = {};
    var nuskColumns = [49, 51, 53, 55, 57, 59, 61, 63, 65, 67];
    var cleanCountry = country.replace(/[🇮🇹🇦🇺🇨🇦🇺🇸🇨🇭🇦🇹🇩🇪🇳🇱🇫🇷🇬🇧🇪🇸🇩🇰🇧🇪🇸🇪🇿🇦🇬🇷🇯🇵🇫🇮🇳🇴🌍]/gu, '').trim();
    
    for (var i = 0; i < flightsData.length; i++) {
      var flightCountry = String(flightsData[i][4]).trim();
      if (flightCountry !== cleanCountry) continue;
      
      for (var j = 0; j < nuskColumns.length; j++) {
        var nuskVal = flightsData[i][nuskColumns[j]];
        if (nuskVal) {
          linkedNusks[Math.floor(Number(nuskVal))] = true;
        }
      }
    }
    
    // 2. جلب الباقات المطابقة
    var packagesLastRow = packagesSheet.getLastRow();
    var packagesData = packagesSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      packagesLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var result = [];
    
    for (var k = 0; k < packagesData.length; k++) {
      var nuskNo = packagesData[k][1];
      if (!nuskNo) continue;
      
      var nuskNum = Math.floor(Number(nuskNo));
      if (!linkedNusks[nuskNum]) continue;
      
      var capacity = Number(packagesData[k][10]) || 0;
      var salesCount = Number(packagesData[k][57]) || 0;
      var remaining = capacity - salesCount;
      
      if (remaining <= 0) continue;
      
      result.push({
        nuskNo: nuskNum,
        name: packagesData[k][2] || '',
        nameEn: packagesData[k][60] || packagesData[k][2] || '',
        price: Number(packagesData[k][5]) || 0,
        noDays: packagesData[k][8] || '',
        capacity: capacity,
        salesCount: salesCount,
        remaining: remaining,
        hotel1Name: packagesData[k][12] || '',
        hotel1NameEn: packagesData[k][13] || '',
        hotel2Name: packagesData[k][27] || '',
        hotel2NameEn: packagesData[k][28] || ''
      });
    }
    
    // ترتيب حسب السعر
    result.sort(function(a, b) { return a.price - b.price; });
    
    return result;
    
  } catch (e) {
    Logger.log('getPackagesByCountry Error: ' + e.toString());
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
// جلب الباقات حسب الفندق
// ═══════════════════════════════════════════════════════════

function getPackagesByHotel(hotelName) {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      lastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var cleanName = hotelName.replace(/[🕋🕌]/g, '').trim();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var hotel1 = String(data[i][12]).trim();
      var hotel1En = String(data[i][13]).trim();
      var hotel2 = String(data[i][27]).trim();
      var hotel2En = String(data[i][28]).trim();
      
      var hasHotel = (hotel1 === cleanName || hotel1En === cleanName ||
                      hotel2 === cleanName || hotel2En === cleanName);
      
      if (!hasHotel) continue;
      
      var capacity = Number(data[i][10]) || 0;
      var salesCount = Number(data[i][57]) || 0;
      var remaining = capacity - salesCount;
      
      if (remaining <= 0) continue;
      
      result.push({
        nuskNo: data[i][1],
        name: data[i][2] || '',
        nameEn: data[i][60] || data[i][2] || '',
        price: Number(data[i][5]) || 0,
        noDays: data[i][8] || '',
        capacity: capacity,
        salesCount: salesCount,
        remaining: remaining,
        hotel1Name: hotel1,
        hotel1NameEn: hotel1En,
        hotel2Name: hotel2,
        hotel2NameEn: hotel2En
      });
    }
    
    result.sort(function(a, b) { return a.price - b.price; });
    return result;
    
  } catch (e) {
    Logger.log('getPackagesByHotel Error: ' + e.toString());
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
// جلب الباقات حسب مدينة الطيران
// ═══════════════════════════════════════════════════════════

function getPackagesByCity(cityName) {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var flightsSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.FLIGHTS);
    var packagesSheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    
    if (!flightsSheet || !packagesSheet) return [];
    
    // 1. جلب أرقام نسك المرتبطة بالمدينة
    var flightsLastRow = flightsSheet.getLastRow();
    var flightsData = flightsSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      flightsLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 70
    ).getValues();
    
    var linkedNusks = {};
    var nuskColumns = [49, 51, 53, 55, 57, 59, 61, 63, 65, 67];
    var cleanCity = cityName.replace(/✈️/g, '').trim();
    
    for (var i = 0; i < flightsData.length; i++) {
      var city = String(flightsData[i][5]).trim();
      if (city !== cleanCity) continue;
      
      for (var j = 0; j < nuskColumns.length; j++) {
        var nuskVal = flightsData[i][nuskColumns[j]];
        if (nuskVal) {
          linkedNusks[Math.floor(Number(nuskVal))] = true;
        }
      }
    }
    
    // 2. جلب الباقات المطابقة
    var packagesLastRow = packagesSheet.getLastRow();
    var packagesData = packagesSheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      packagesLastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var result = [];
    
    for (var k = 0; k < packagesData.length; k++) {
      var nuskNo = packagesData[k][1];
      if (!nuskNo) continue;
      
      var nuskNum = Math.floor(Number(nuskNo));
      if (!linkedNusks[nuskNum]) continue;
      
      var capacity = Number(packagesData[k][10]) || 0;
      var salesCount = Number(packagesData[k][57]) || 0;
      var remaining = capacity - salesCount;
      
      if (remaining <= 0) continue;
      
      result.push({
        nuskNo: nuskNum,
        name: packagesData[k][2] || '',
        nameEn: packagesData[k][60] || packagesData[k][2] || '',
        price: Number(packagesData[k][5]) || 0,
        noDays: packagesData[k][8] || '',
        capacity: capacity,
        salesCount: salesCount,
        remaining: remaining,
        hotel1Name: packagesData[k][12] || '',
        hotel1NameEn: packagesData[k][13] || '',
        hotel2Name: packagesData[k][27] || '',
        hotel2NameEn: packagesData[k][28] || ''
      });
    }
    
    result.sort(function(a, b) { return a.price - b.price; });
    return result;
    
  } catch (e) {
    Logger.log('getPackagesByCity Error: ' + e.toString());
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
// ملخص الباقات للذكاء الاصطناعي
// ═══════════════════════════════════════════════════════════

function getPackagesSummaryForAI() {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.PACKAGES);
    if (!sheet) return '';
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(
      AGENT_CONFIG.DATA_START_ROW, 1,
      lastRow - AGENT_CONFIG.DATA_START_ROW + 1, 62
    ).getValues();
    
    var summary = [];
    var count = 0;
    
    for (var i = 0; i < data.length && count < 15; i++) {
      var name = data[i][2];
      if (!name) continue;
      
      var capacity = Number(data[i][10]) || 0;
      var salesCount = Number(data[i][57]) || 0;
      var remaining = capacity - salesCount;
      
      if (remaining <= 0) continue;
      
      var price = Number(data[i][5]) || 0;
      var days = data[i][8] || '';
      
      summary.push(name + ': ' + price + ' SAR, ' + days + ' days, ' + remaining + ' available');
      count++;
    }
    
    return summary.join('\n');
    
  } catch (e) {
    Logger.log('getPackagesSummaryForAI Error: ' + e.toString());
    return '';
  }
}

// ═══════════════════════════════════════════════════════════
// دوال الاختبار
// ═══════════════════════════════════════════════════════════

function testDataFetcher() {
  Logger.log('=== اختبار DataFetcher ===');
  
  var countries = getAvailableCountries();
  Logger.log('الدول المتاحة (' + countries.length + '):');
  Logger.log(countries.join(', '));
  
  Logger.log('');
  
  if (countries.length > 0) {
    var testCountry = countries[0];
    Logger.log('اختبار جلب باقات: ' + testCountry);
    var packages = getPackagesByCountry(testCountry);
    Logger.log('عدد الباقات: ' + packages.length);
    if (packages.length > 0) {
      Logger.log('أول باقة: ' + packages[0].name + ' - ' + packages[0].price + ' SAR');
    }
  }
  
  Logger.log('');
  
  var hotels = getAvailableHotels();
  Logger.log('الفنادق المتاحة (' + hotels.length + '):');
  Logger.log(hotels.slice(0, 5).join(', ') + '...');
  
  Logger.log('');
  
  var cities = getAvailableCities();
  Logger.log('المدن المتاحة (' + cities.length + '):');
  Logger.log(cities.slice(0, 5).join(', ') + '...');
}