/**
 * نظام البحث - شركة إكرام الضيف للسياحة
 * الإصدار 27.0 - الدول المستهدفة + توفر الغرف + إخفاء المكتمل
 */

var CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  PACKAGES_SHEET: 'الباقات',
  FLIGHTS_SHEET: 'الطيران',
  HOTELS_SHEET: 'الفنادق',
  SETTINGS_SHEET: 'الإعدادات',
  ROOM_AVAIL_SHEET: 'توفر_الغرف',
  DATA_START: 3,
  HOTELS_START: 2,
  ROOM_AVAIL_START: 3
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Ikram AlDyf for Hajj')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPassword() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
    if (!sheet) return '1234';
    return sheet.getRange('B1').getValue().toString();
  } catch (e) {
    return '1234';
  }
}

function getAllData() {
  try {
    var packages = getPackages();
    var flights = getFlights();
    var hotels = getHotels();
    var roomAvail = getRoomAvailability();
    var stats = calculateStats(packages, flights);
    
    // دمج توفر الغرف مع الباقات
    for (var k = 0; k < packages.length; k++) {
      var nusk = Number(packages[k].nuskNo);
      if (roomAvail[nusk]) {
        packages[k].roomAvail = roomAvail[nusk];
      }
    }
    
    return JSON.stringify({
      packages: packages || [],
      flights: flights || [],
      hotels: hotels || [],
      stats: stats,
      success: true
    });
    
  } catch (error) {
    Logger.log('getAllData ERROR: ' + error.toString());
    return JSON.stringify({
      packages: [],
      flights: [],
      hotels: [],
      stats: { availablePackages: 0, totalPilgrims: 0, soldPilgrims: 0, remainingPilgrims: 0, bookingPercent: 0, remainingTickets: 0 },
      success: false,
      error: error.toString()
    });
  }
}

/**
 * حساب الإحصائيات المحدثة
 */
function calculateStats(packages, flights) {
  var availablePackages = 0;
  var totalPilgrims = 0;
  var soldPilgrims = 0;
  var remainingPilgrims = 0;
  var remainingTickets = 0;
  
  for (var i = 0; i < packages.length; i++) {
    var p = packages[i];
    totalPilgrims += p.capacity || 0;
    soldPilgrims += p.salesCount || 0;
    remainingPilgrims += p.remaining || 0;
    if (p.remaining > 0) {
      availablePackages++;
    }
  }
  
  for (var j = 0; j < flights.length; j++) {
    var f = flights[j];
    if (f.remaining > 0) {
      remainingTickets += f.remaining;
    }
  }
  
  var bookingPercent = totalPilgrims > 0 ? Math.round((soldPilgrims / totalPilgrims) * 1000) / 10 : 0;
  
  return {
    availablePackages: availablePackages,
    totalPilgrims: totalPilgrims,
    soldPilgrims: soldPilgrims,
    remainingPilgrims: remainingPilgrims,
    bookingPercent: bookingPercent,
    remainingTickets: remainingTickets
  };
}

function getPackages() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.PACKAGES_SHEET);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START) return [];
    
    var numRows = lastRow - CONFIG.DATA_START + 1;
    var data = sheet.getRange(CONFIG.DATA_START, 1, numRows, 63).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var name = row[2];
      if (!name || String(name).trim() === '') continue;
      
      var capacity = toNumber(row[10]);
      var salesCount = toNumber(row[57]);
      var remaining = capacity - salesCount;
      var salesPercent = parsePercent(row[59]);
      
      if (salesPercent === 0 && capacity > 0 && salesCount > 0) {
        salesPercent = (salesCount / capacity) * 100;
      }
      
      var hotel1Nights = calculateNights(row[14], row[15]);
      var hotel2Nights = calculateNights(row[29], row[30]);
      var hotel3Nights = calculateNights(row[44], row[45]);
      
      result.push({
        no: row[0],
        nuskNo: row[1],
        name: cleanString(name),
        nameEn: cleanString(row[60]) || cleanString(name),
        category: cleanString(row[3]) || 'Standard',
        ikramNo: row[4],
        price: toNumber(row[5]),
        dateStart: formatDate(row[6]),
        dateEnd: formatDate(row[7]),
        noDays: row[8],
        cityStart: cleanString(row[9]),
        capacity: capacity,
        bookingLink: row[61] || '',
        
        hotel1City: cleanString(row[11]),
        hotel1Name: cleanString(row[12]),
        hotel1NameEn: cleanString(row[13]),
        hotel1CheckIn: formatDateShort(row[14]),
        hotel1CheckOut: formatDateShort(row[15]),
        hotel1Nights: hotel1Nights,
        
        hotel2City: cleanString(row[26]),
        hotel2Name: cleanString(row[27]),
        hotel2NameEn: cleanString(row[28]),
        hotel2CheckIn: formatDateShort(row[29]),
        hotel2CheckOut: formatDateShort(row[30]),
        hotel2Nights: hotel2Nights,
        
        hotel3City: cleanString(row[41]),
        hotel3Name: cleanString(row[42]),
        hotel3NameEn: cleanString(row[43]),
        hotel3CheckIn: formatDateShort(row[44]),
        hotel3CheckOut: formatDateShort(row[45]),
        hotel3Nights: hotel3Nights,
        
        salesCount: salesCount,
        remaining: remaining,
        salesPercent: Math.round(salesPercent * 100) / 100,
        isSoldOut: remaining <= 0,
        roomAvail: null
      });
    }
    return result;
  } catch (error) {
    Logger.log('getPackages ERROR: ' + error.toString());
    return [];
  }
}

function getFlights() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.FLIGHTS_SHEET);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START) return [];
    
    var numRows = lastRow - CONFIG.DATA_START + 1;
    var data = sheet.getRange(CONFIG.DATA_START, 1, numRows, 103).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var pnr = row[1];
      if (!pnr || String(pnr).trim() === '') continue;
      
      var linkedNusk = [];
      var pkgIndices = [49, 51, 53, 55, 57, 59, 61, 63, 65, 67];
      for (var j = 0; j < pkgIndices.length; j++) {
        var val = row[pkgIndices[j]];
        if (val && !isNaN(Number(val))) linkedNusk.push(Number(val));
      }
      
      var pax = toNumber(row[7]);
      var salesCount = toNumber(row[86]);
      var remaining = pax - salesCount;
      var salesPercent = parsePercent(row[88]);
      
      if (salesPercent === 0 && pax > 0 && salesCount > 0) {
        salesPercent = (salesCount / pax) * 100;
      }
      
      var flight1No = cleanFlightData(row[21]);
      var flight2No = cleanFlightData(row[28]);
      var flight3No = cleanFlightData(row[35]);
      var flight4No = cleanFlightData(row[42]);
      
      var isDirect = !flight1No && flight2No;
      
      // استخراج الدول المستهدفة (أعمدة CO-CY = 93-103 = index 92-102)
      var targetCountries = [];
      for (var tc = 92; tc <= 102; tc++) {
        var tcVal = cleanString(row[tc]);
        if (tcVal) targetCountries.push(tcVal);
      }
      
      result.push({
        no: row[0],
        pnr: cleanString(pnr),
        supplier: cleanString(row[2]),
        status: cleanString(row[3]),
        country: cleanString(row[4]),
        countryEn: cleanString(row[4]),
        city: cleanString(row[5]),
        cityCode: extractCityCode(row[5]),
        airline: cleanString(row[6]),
        pax: pax,
        noDays: row[8],
        priceNusk: toNumber(row[15]),
        isDirect: isDirect,
        
        flight1No: flight1No,
        flight1Date: cleanDate(row[22]),
        flight1TimeDep: cleanTime(row[23]),
        flight1From: cleanFlightData(row[24]),
        flight1To: cleanFlightData(row[25]),
        flight1DateArr: cleanDate(row[26]),
        flight1TimeArr: cleanTime(row[27]),
        
        flight2No: flight2No,
        flight2Date: cleanDate(row[29]),
        flight2TimeDep: cleanTime(row[30]),
        flight2From: cleanFlightData(row[31]),
        flight2To: cleanFlightData(row[32]),
        flight2DateArr: cleanDate(row[33]),
        flight2TimeArr: cleanTime(row[34]),
        
        flight3No: flight3No,
        flight3Date: cleanDate(row[36]),
        flight3TimeDep: cleanTime(row[37]),
        flight3From: cleanFlightData(row[38]),
        flight3To: cleanFlightData(row[39]),
        flight3DateArr: cleanDate(row[40]),
        flight3TimeArr: cleanTime(row[41]),
        
        flight4No: flight4No,
        flight4Date: cleanDate(row[43]),
        flight4TimeDep: cleanTime(row[44]),
        flight4From: cleanFlightData(row[45]),
        flight4To: cleanFlightData(row[46]),
        flight4DateArr: cleanDate(row[47]),
        flight4TimeArr: cleanTime(row[48]),
        
        linkedNusk: linkedNusk,
        targetCountries: targetCountries,
        salesCount: salesCount,
        remaining: remaining,
        salesPercent: Math.round(salesPercent * 100) / 100,
        isSoldOut: remaining <= 0
      });
    }
    return result;
  } catch (error) {
    Logger.log('getFlights ERROR: ' + error.toString());
    return [];
  }
}

/**
 * قراءة بيانات توفر الغرف
 * يُرجع: map بـ Nusk No → { h1: {dbl, tri, quad}, h2: {...}, h3: {...}, lastUpdate }
 */
function getRoomAvailability() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.ROOM_AVAIL_SHEET);
    if (!sheet) return {};
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.ROOM_AVAIL_START) return {};
    
    var numRows = lastRow - CONFIG.ROOM_AVAIL_START + 1;
    var data = sheet.getRange(CONFIG.ROOM_AVAIL_START, 1, numRows, 14).getValues();
    var result = {};
    
    for (var i = 0; i < data.length; i++) {
      var nusk = data[i][0];
      if (!nusk) continue;
      
      result[Number(nusk)] = {
        h1: {
          dbl: parseAvailability(data[i][3]),
          tri: parseAvailability(data[i][4]),
          quad: parseAvailability(data[i][5])
        },
        h2: {
          dbl: parseAvailability(data[i][6]),
          tri: parseAvailability(data[i][7]),
          quad: parseAvailability(data[i][8])
        },
        h3: {
          dbl: parseAvailability(data[i][9]),
          tri: parseAvailability(data[i][10]),
          quad: parseAvailability(data[i][11])
        },
        lastUpdate: data[i][12] ? formatDateTimeRA(data[i][12]) : '',
        notes: cleanString(data[i][13])
      };
    }
    return result;
  } catch (error) {
    Logger.log('getRoomAvailability ERROR: ' + error.toString());
    return {};
  }
}

/**
 * تحويل نص التوفر إلى كود مختصر
 */
function parseAvailability(val) {
  if (!val) return '';
  var str = String(val).trim();
  if (str.indexOf('متوفر') !== -1 && str.indexOf('غير') === -1) return 'available';
  if (str.indexOf('غير متوفر') !== -1) return 'unavailable';
  if (str.indexOf('غير مؤكد') !== -1) return 'unconfirmed';
  return '';
}

/**
 * تنسيق التاريخ/الوقت لتوفر الغرف
 */
function formatDateTimeRA(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      return ('0' + val.getDate()).slice(-2) + '/' +
             ('0' + (val.getMonth() + 1)).slice(-2) + ' ' +
             ('0' + val.getHours()).slice(-2) + ':' +
             ('0' + val.getMinutes()).slice(-2);
    }
    return String(val).trim();
  } catch (e) {
    return '';
  }
}

function getHotels() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.HOTELS_SHEET);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.HOTELS_START) return [];
    
    var numRows = lastRow - CONFIG.HOTELS_START + 1;
    var data = sheet.getRange(CONFIG.HOTELS_START, 1, numRows, 4).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[2]) continue;
      result.push({
        nameAr: cleanString(row[0]),
        city: cleanString(row[1]),
        nameEn: cleanString(row[2]),
        mapLink: row[3] || ''
      });
    }
    return result;
  } catch (error) {
    Logger.log('getHotels ERROR: ' + error.toString());
    return [];
  }
}

/**
 * دالة تنظيف النصوص - إزالة المسافات الزائدة
 */
function cleanString(val) {
  if (!val) return '';
  return String(val).trim();
}

function cleanFlightData(val) {
  if (!val) return '';
  var str = String(val).trim();
  if (str === '' || str === '0' || str === 'null' || str === 'undefined') return '';
  return str;
}

function cleanDate(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      if (val.getFullYear() < 2000) return '';
      return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2) + '-' + ('0' + val.getDate()).slice(-2);
    }
    var str = String(val).trim();
    if (str.indexOf('1899') !== -1 || str.indexOf('1900') !== -1) return '';
    return str;
  } catch (e) {
    return '';
  }
}

function cleanTime(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) {
      if (val.getFullYear() < 2000) return '';
      return ('0' + val.getHours()).slice(-2) + ':' + ('0' + val.getMinutes()).slice(-2);
    }
    var str = String(val).trim();
    if (str.indexOf('1899') !== -1 || str.indexOf('1900') !== -1) return '';
    var match = str.match(/(\d{1,2}):(\d{2})/);
    if (match) return match[0];
    return '';
  } catch (e) {
    return '';
  }
}

function extractCityCode(city) {
  if (!city) return '';
  var str = String(city).trim().toUpperCase();
  if (str.length === 3) return str;
  var codes = {
    'PARIS': 'CDG', 'LONDON': 'LHR', 'SYDNEY': 'SYD', 'ROME': 'FCO',
    'MILAN': 'MXP', 'GENEVA': 'GVA', 'ZURICH': 'ZRH', 'TORONTO': 'YYZ',
    'NEW YORK': 'JFK', 'LOS ANGELES': 'LAX', 'JEDDAH': 'JED', 'MADINAH': 'MED'
  };
  return codes[str] || str.substring(0, 3);
}

function calculateNights(checkIn, checkOut) {
  if (!checkIn || !checkOut) return 0;
  try {
    var d1 = checkIn instanceof Date ? checkIn : new Date(checkIn);
    var d2 = checkOut instanceof Date ? checkOut : new Date(checkOut);
    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return 0;
    if (d1.getFullYear() < 2000 || d2.getFullYear() < 2000) return 0;
    var diff = Math.abs(d2 - d1);
    return Math.ceil(diff / (1000 * 60 * 60 * 24));
  } catch (e) {
    return 0;
  }
}

function toNumber(val) {
  if (!val && val !== 0) return 0;
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

function parsePercent(val) {
  if (!val && val !== 0) return 0;
  var num = typeof val === 'string' ? parseFloat(val.replace('%', '').replace(',', '.')) : Number(val);
  if (isNaN(num)) return 0;
  if (num > 0 && num <= 1) num = num * 100;
  return num;
}

function formatDate(val) {
  if (!val) return null;
  try {
    if (val instanceof Date) {
      if (val.getFullYear() < 2000) return null;
      return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2) + '-' + ('0' + val.getDate()).slice(-2);
    }
    return String(val).trim();
  } catch (e) {
    return String(val).trim();
  }
}

function formatDateShort(val) {
  if (!val) return null;
  try {
    var months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    if (val instanceof Date) {
      if (val.getFullYear() < 2000) return null;
      return val.getDate() + ' ' + months[val.getMonth()];
    }
    var d = new Date(val);
    if (!isNaN(d.getTime()) && d.getFullYear() >= 2000) {
      return d.getDate() + ' ' + months[d.getMonth()];
    }
    return String(val).trim();
  } catch (e) {
    return String(val).trim();
  }
}