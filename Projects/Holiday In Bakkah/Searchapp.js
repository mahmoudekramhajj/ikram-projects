/**
 * نظام البحث - شركة هوليداي إن بكة
 * Holiday In Bakkah - Search System
 * الإصدار 6.0 - Modular Structure
 * 
 * هيكل الملفات:
 * - SearchApp.gs      → دوال الخادم + include
 * - Index.html        → هيكل HTML فقط
 * - Styles.html       → CSS
 * - Scripts_Core.html → المتغيرات، الترجمة، الفلاتر، الإحصائيات
 * - Scripts_Packages.html → عرض الباقات + المشاركة
 * - Scripts_Flights.html  → عرض الطيران
 * - Scripts_Hotels.html   → عرض الفنادق
 */

var CONFIG = {
  SPREADSHEET_ID: '1ggPKRVnGAwIJqKHmxwRnPkcRWEpcDDmT--KQyg6o15I',
  PACKAGES_SHEET: 'الباقات',
  FLIGHTS_SHEET: 'الطيران',
  HOTELS_SHEET: 'الفنادق',
  DATA_START: 3,
  HOTELS_START: 2
};

/* ══════════════════════════════════════
   دالة include لتضمين ملفات HTML
   ══════════════════════════════════════ */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ══════════════════════════════════════
   نقطة الدخول - Web App
   ══════════════════════════════════════ */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holiday In Bakkah - هوليداي إن بكة')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/* ══════════════════════════════════════
   جلب كل البيانات
   ══════════════════════════════════════ */
function getAllData() {
  try {
    var hotels = getHotels();
    var flights = getFlights();
    var packages = getPackages(flights, hotels);
    
    return JSON.stringify({
      packages: packages || [],
      flights: flights || [],
      hotels: hotels || [],
      success: true
    });
    
  } catch (error) {
    Logger.log('getAllData ERROR: ' + error.toString());
    return JSON.stringify({
      packages: [],
      flights: [],
      hotels: [],
      success: false,
      error: error.toString()
    });
  }
}

/* ══════════════════════════════════════
   جلب الفنادق
   ══════════════════════════════════════ */
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
      if (!row[0]) continue;
      result.push({
        nameAr: row[0],
        city: row[1],
        nameEn: row[2] || row[0],
        mapLink: row[3] || ''
      });
    }
    return result;
  } catch (error) {
    Logger.log('getHotels ERROR: ' + error.toString());
    return [];
  }
}

/* ══════════════════════════════════════
   جلب الباقات مع ربط الفنادق والرحلات
   ══════════════════════════════════════ */
function getPackages(flights, hotels) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.PACKAGES_SHEET);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START) return [];
    
    var numRows = lastRow - CONFIG.DATA_START + 1;
    var data = sheet.getRange(CONFIG.DATA_START, 1, numRows, 62).getValues();
    var result = [];
    
    // Build hotel lookup by Arabic name
    var hotelLookup = {};
    if (hotels && hotels.length > 0) {
      hotels.forEach(function(h) {
        if (h.nameAr) hotelLookup[h.nameAr.trim()] = h;
      });
    }
    
    // Build flight lookup by Nusk No
    var flightsByNusk = {};
    if (flights && flights.length > 0) {
      flights.forEach(function(flight) {
        if (flight.linkedPackages && flight.linkedPackages.length > 0) {
          flight.linkedPackages.forEach(function(nuskNo) {
            if (!flightsByNusk[nuskNo]) {
              flightsByNusk[nuskNo] = [];
            }
            flightsByNusk[nuskNo].push(flight);
          });
        }
      });
    }
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var name = row[2];
      if (!name || String(name).trim() === '') continue;
      
      var nuskNo = toNumber(row[1]);
      var holidayNo = toNumber(row[4]);
      var salesCount = toNumber(row[57]);
      var remaining = toNumber(row[58]);
      var salesPercent = parsePercent(row[59]);
      var capacity = toNumber(row[10]);
      
      if (salesPercent === 0 && capacity > 0 && salesCount > 0) {
        salesPercent = (salesCount / capacity) * 100;
      }
      
      // Get hotel info with map links
      var hotel1Info = getHotelInfo(row[12], hotelLookup);
      var hotel2Info = getHotelInfo(row[27], hotelLookup);
      var hotel3Info = getHotelInfo(row[42], hotelLookup);
      
      // Get linked flights
      var packageFlights = flightsByNusk[nuskNo] || [];
      
      result.push({
        no: row[0],
        nuskNo: nuskNo,
        name: name,
        nameEn: row[60] || name,
        category: row[3] || 'Standard',
        holidayNo: holidayNo,
        price: toNumber(row[5]),
        dateStart: formatDate(row[6]),
        dateEnd: formatDate(row[7]),
        noDays: toNumber(row[8]),
        cityStart: row[9],
        capacity: capacity,
        bookingLink: row[61] || '',
        
        // Hotel 1
        hotel1City: row[11],
        hotel1Name: row[12],
        hotel1NameEn: row[13] || hotel1Info.nameEn,
        hotel1CheckIn: formatDateShort(row[14]),
        hotel1CheckOut: formatDateShort(row[15]),
        hotel1Nights: calculateNights(row[14], row[15]),
        hotel1MapLink: hotel1Info.mapLink,
        
        // Hotel 2
        hotel2City: row[26],
        hotel2Name: row[27],
        hotel2NameEn: row[28] || hotel2Info.nameEn,
        hotel2CheckIn: formatDateShort(row[29]),
        hotel2CheckOut: formatDateShort(row[30]),
        hotel2Nights: calculateNights(row[29], row[30]),
        hotel2MapLink: hotel2Info.mapLink,
        
        // Hotel 3
        hotel3City: row[41],
        hotel3Name: row[42],
        hotel3NameEn: row[43] || hotel3Info.nameEn,
        hotel3CheckIn: formatDateShort(row[44]),
        hotel3CheckOut: formatDateShort(row[45]),
        hotel3Nights: calculateNights(row[44], row[45]),
        hotel3MapLink: hotel3Info.mapLink,
        
        // Sales
        salesCount: salesCount,
        remaining: remaining,
        salesPercent: Math.round(salesPercent * 100) / 100,
        isSoldOut: salesPercent >= 100,
        
        // Linked flights
        flights: packageFlights
      });
    }
    return result;
  } catch (error) {
    Logger.log('getPackages ERROR: ' + error.toString());
    return [];
  }
}

function getHotelInfo(hotelName, hotelLookup) {
  if (!hotelName) return { nameEn: '', mapLink: '' };
  var key = String(hotelName).trim();
  var hotel = hotelLookup[key];
  if (hotel) {
    return { nameEn: hotel.nameEn || '', mapLink: hotel.mapLink || '' };
  }
  return { nameEn: '', mapLink: '' };
}

/* ══════════════════════════════════════
   جلب الرحلات
   ══════════════════════════════════════ */
function getFlights() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.FLIGHTS_SHEET);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.DATA_START) return [];
    
    var numRows = lastRow - CONFIG.DATA_START + 1;
    var data = sheet.getRange(CONFIG.DATA_START, 1, numRows, 91).getValues();
    var result = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var pnr = row[1];
      if (!pnr || String(pnr).trim() === '') continue;
      
      var linkedPackages = [];
      var pkgIndices = [49, 51, 53, 55, 57, 59, 61, 63, 65, 67];
      for (var j = 0; j < pkgIndices.length; j++) {
        var val = row[pkgIndices[j]];
        if (val && !isNaN(Number(val))) {
          linkedPackages.push(Number(val));
        }
      }
      
      var pax = toNumber(row[7]);
      var salesCount = toNumber(row[86]);
      var remaining = toNumber(row[87]);
      var salesPercent = parsePercent(row[88]);
      
      if (salesPercent === 0 && pax > 0 && salesCount > 0) {
        salesPercent = (salesCount / pax) * 100;
      }
      
// تحديد نوع رحلة الذهاب: مباشرة أم ترانزيت
      var depLeg1No = row[21], depLeg1Date = row[22], depLeg1From = row[24], depLeg1To = row[25];
      var depLeg2No = row[28], depLeg2Date = row[29], depLeg2From = row[31], depLeg2To = row[32];
      var depIsDirect = !depLeg1No || String(depLeg1No).trim() === '';
      
      // تحديد نوع رحلة العودة: مباشرة أم ترانزيت
      var retLeg1No = row[35], retLeg1Date = row[36], retLeg1From = row[38], retLeg1To = row[39];
      var retLeg2No = row[42], retLeg2Date = row[43], retLeg2From = row[45], retLeg2To = row[46];
      var retIsDirect = !retLeg2No || String(retLeg2No).trim() === '';

      result.push({
        no: row[0],
        pnr: pnr,
        supplier: row[2],
        status: row[3],
        country: row[4],
        city: row[5],
        airline: row[6],
        pax: pax,
        noDays: row[8],
        priceNusk: toNumber(row[15]),
        
        // رحلة الذهاب
        depIsDirect: depIsDirect,
        dep1No: depIsDirect ? depLeg2No : depLeg1No,
        dep1Date: formatDate(depIsDirect ? depLeg2Date : depLeg1Date),
        dep1From: depIsDirect ? depLeg2From : depLeg1From,
        dep1To: depIsDirect ? depLeg2To : depLeg1To,
        dep2No: depIsDirect ? null : depLeg2No,
        dep2Date: depIsDirect ? null : formatDate(depLeg2Date),
        dep2From: depIsDirect ? null : depLeg2From,
        dep2To: depIsDirect ? null : depLeg2To,
        
        // رحلة العودة
        retIsDirect: retIsDirect,
        ret1No: retLeg1No,
        ret1Date: formatDate(retLeg1Date),
        ret1From: retLeg1From,
        ret1To: retLeg1To,
        ret2No: retIsDirect ? null : retLeg2No,
        ret2Date: retIsDirect ? null : formatDate(retLeg2Date),
        ret2From: retIsDirect ? null : retLeg2From,
        ret2To: retIsDirect ? null : retLeg2To,
        
        linkedPackages: linkedPackages,
        salesCount: salesCount,
        remaining: remaining,
        salesPercent: Math.round(salesPercent * 100) / 100,
        isSoldOut: salesPercent >= 100
      });
    }
    return result;
  } catch (error) {
    Logger.log('getFlights ERROR: ' + error.toString());
    return [];
  }
}

/* ══════════════════════════════════════
   دوال مساعدة
   ══════════════════════════════════════ */
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
      return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2) + '-' + ('0' + val.getDate()).slice(-2);
    }
    return String(val);
  } catch (e) {
    return String(val);
  }
}

function formatDateShort(val) {
  if (!val) return null;
  try {
    var months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    if (val instanceof Date) {
      return val.getDate() + ' ' + months[val.getMonth()];
    }
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return d.getDate() + ' ' + months[d.getMonth()];
    }
    return String(val);
  } catch (e) {
    return String(val);
  }
}

function calculateNights(checkIn, checkOut) {
  if (!checkIn || !checkOut) return 0;
  try {
    var d1 = checkIn instanceof Date ? checkIn : new Date(checkIn);
    var d2 = checkOut instanceof Date ? checkOut : new Date(checkOut);
    if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return 0;
    var diff = Math.abs(d2 - d1);
    return Math.ceil(diff / (1000 * 60 * 60 * 24));
  } catch (e) {
    return 0;
  }
}