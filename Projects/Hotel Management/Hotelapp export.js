/**
 * Hotel Management App — Export & Forecast
 * إكرام الضيف للسياحة | حج 1447هـ — 2026م
 * v3.0 — Excel export, reports, forecast
 */

// ============================================================
// EXPORT TO EXCEL
// ============================================================

function exportToExcel(reportType, hotelName, hotelCity, filterData) {
  var ss = SpreadsheetApp.create('تقرير_' + reportType + '_' + Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyyMMdd_HHmm'));
  var sheet = ss.getActiveSheet();
  
  switch (reportType) {
    case 'pilgrims':
      exportPilgrimsReport_(sheet, hotelName, hotelCity, filterData);
      break;
    case 'departures':
      exportDeparturesReport_(sheet, hotelName, hotelCity, filterData);
      break;
    case 'inventory':
      exportInventoryReport_(sheet, hotelName, hotelCity);
      break;
    case 'dashboard':
      exportDashboardReport_(sheet);
      break;
    case 'forecast':
      exportForecastReport_(ss);
      break;
  }
  
  var fileId = ss.getId();
  var url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?format=xlsx';
  
  ScriptApp.newTrigger('deleteFile_')
    .timeBased()
    .after(60 * 60 * 1000)
    .create();
  PropertiesService.getScriptProperties().setProperty('tempFileId', fileId);
  
  return { success: true, url: url, fileName: ss.getName() + '.xlsx' };
}

function deleteFile_() {
  try {
    var fileId = PropertiesService.getScriptProperties().getProperty('tempFileId');
    if (fileId) DriveApp.getFileById(fileId).setTrashed(true);
  } catch(e) {}
}

function exportPilgrimsReport_(sheet, hotelName, hotelCity, filterData) {
  sheet.setName('حجاج ' + hotelName);
  var pilgrims = getPilgrimsForHotel(hotelName, hotelCity);
  
  // Apply filters if provided
  if (filterData) {
    pilgrims = applyPilgrimFilters_(pilgrims, filterData);
  }
  
  var headers = ['الاسم', 'الجواز', 'الجنس', 'الجنسية', 'البلد', 'المجموعة', 'الباقة',
    'رحلة الوصول', 'تاريخ الوصول', 'وقت الوصول', 'الوصول المتوقع',
    'بداية العقد', 'نهاية العقد', 'نوع الغرفة',
    'حالة الوصول', 'ساعات الانتظار', 'أيام التأخر',
    'حالة Check-in', 'رقم الغرفة', 'وقت Check-in',
    'المرشد', 'قائد المجموعة', 'هاتف القائد'];
  
  sheet.appendRow(headers);
  
  pilgrims.forEach(function(p) {
    sheet.appendRow([
      p.name, p.passport, p.gender, p.nationality, p.countryResidence, p.groupNumber, p.packageName,
      p.arrivalFlight, p.arrivalDate, p.arrivalTime, p.expectedArrival,
      p.hotelCheckIn, p.hotelCheckOut, p.roomType,
      p.arrivalStatus, p.earlyHours || '', p.lateDays || '',
      p.checkInStatus, p.roomNumber, p.checkInTime,
      p.tourGuide, p.groupLeader, p.leaderPhone
    ]);
  });
  
  formatExportSheet_(sheet, headers.length);
}

function applyPilgrimFilters_(pilgrims, filters) {
  return pilgrims.filter(function(p) {
    if (filters.date && p.arrivalDate !== filters.date && p.hotelCheckIn !== filters.date) return false;
    if (filters.status && p.checkInStatus !== filters.status) return false;
    if (filters.flight && p.arrivalFlight !== filters.flight) return false;
    if (filters.search) {
      var s = filters.search.toLowerCase();
      if (!(p.name || '').toLowerCase().includes(s) && !(p.passport || '').toLowerCase().includes(s)) return false;
    }
    return true;
  });
}

function exportDeparturesReport_(mainSheet, hotelName, hotelCity, filterData) {
  var ss = mainSheet.getParent();
  var departures = getDeparturesForHotel(hotelName, hotelCity);

  if (filterData && filterData.date) {
    departures = departures.filter(function(d) { return d.departureDate === filterData.date; });
  }

  // ── 1. تجميع حسب المحطة (رحلة / فندق / قطار) ──
  var groups = {};
  departures.forEach(function(d) {
    var isAirport = d.destination && d.destination.indexOf('مطار') >= 0;
    var isTrain   = d.transport === 'قطار';

    var groupKey;
    if (isAirport) {
      // تجميع حسب الرحلة الجوية أولاً
      groupKey = (d.linkedFlight || d.destination || 'مطار') + '|airport';
    } else if (isTrain) {
      groupKey = 'قطار|train';
    } else {
      // حافلة لفندق آخر أو تحويل
      groupKey = (d.destinationHotel || d.destination || 'أخرى') + '|bus';
    }

    if (!groups[groupKey]) {
      groups[groupKey] = {
        label: isAirport ? (d.linkedFlight || d.destination)
               : isTrain ? 'قطار'
               : (d.destinationHotel || d.destination || 'أخرى'),
        type: isAirport ? 'airport' : isTrain ? 'train' : 'bus',
        destination: d.destination || '',
        rows: []
      };
    }
    groups[groupKey].rows.push(d);
  });

  var DEP_HEADERS = [
    'الاسم', 'الجواز', 'المجموعة', 'تاريخ المغادرة', 'وقت المغادرة',
    'الوجهة', 'الفندق التالي', 'وسيلة النقل', 'رحلة العودة', 'وقت الرحلة', 'نوع المغادرة'
  ];

  // ── 2. ملخص كل المغادرات (الشيت الأول) ──
  mainSheet.setName('ملخص المغادرات');
  mainSheet.appendRow(['الفندق: ' + hotelName + (filterData && filterData.date ? ' | تاريخ: ' + filterData.date : ' | كل المغادرات')]);
  mainSheet.appendRow(['إجمالي الحجاج: ' + departures.length + ' | المجموعات: ' + Object.keys(groups).length]);
  mainSheet.appendRow(['']);
  mainSheet.appendRow(DEP_HEADERS);
  departures.sort(function(a, b) {
    return (a.departureDate + (a.departureTime || '')) > (b.departureDate + (b.departureTime || '')) ? 1 : -1;
  }).forEach(function(d) {
    mainSheet.appendRow([
      d.name, d.passport, d.groupNumber, d.departureDate, d.departureTime,
      d.destination, d.destinationHotel, d.transport, d.linkedFlight, d.linkedTime, d.type
    ]);
  });
  mainSheet.getRange(4, 1, 1, DEP_HEADERS.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
  mainSheet.setFrozenRows(4);
  for (var c = 1; c <= DEP_HEADERS.length; c++) mainSheet.autoResizeColumn(c);

  // ── 3. شيت منفصل لكل مجموعة ──
  var sortedKeys = Object.keys(groups).sort(function(a, b) {
    // ترتيب: مطار أولاً ← قطار ← حافلة
    var order = { airport: 1, train: 2, bus: 3 };
    return (order[groups[a].type] || 4) - (order[groups[b].type] || 4);
  });

  sortedKeys.forEach(function(key) {
    var g = groups[key];
    var typeIcon = g.type === 'airport' ? '✈ ' : g.type === 'train' ? '🚂 ' : '🚌 ';
    var sheetName = (typeIcon + g.label).substring(0, 100).replace(/[\/\\?*\[\]]/g, '-');

    var s = ss.insertSheet(sheetName);

    // عنوان
    var titleParts = [
      'الفندق: ' + hotelName,
      g.type === 'airport' ? ('رحلة: ' + g.label + ' | ' + g.destination) :
      g.type === 'train'   ? 'محطة القطار' :
      ('الفندق التالي: ' + g.label + (g.destination ? ' — ' + g.destination : ''))
    ];
    if (filterData && filterData.date) titleParts.push('تاريخ: ' + filterData.date);

    s.appendRow([titleParts.join(' | ')]);
    s.appendRow(['عدد الحجاج: ' + g.rows.length]);
    s.appendRow(['']);
    s.appendRow(DEP_HEADERS);

    // بيانات مرتبة حسب الوقت
    g.rows.sort(function(a, b) {
      return (a.departureDate + (a.departureTime || '')) > (b.departureDate + (b.departureTime || '')) ? 1 : -1;
    }).forEach(function(d) {
      s.appendRow([
        d.name, d.passport, d.groupNumber, d.departureDate, d.departureTime,
        d.destination, d.destinationHotel, d.transport, d.linkedFlight, d.linkedTime, d.type
      ]);
    });

    // تنسيق
    var bgColor = g.type === 'airport' ? '#3b1f07' : g.type === 'train' ? '#064e3b' : '#1e3a5f';
    s.getRange(1, 1).setFontWeight('bold').setFontSize(12).setFontColor('#ffffff');
    s.getRange(1, 1, 3, DEP_HEADERS.length).setBackground(bgColor);
    s.getRange(4, 1, 1, DEP_HEADERS.length).setFontWeight('bold').setBackground('#0f172a').setFontColor('#e2e8f0');
    s.setFrozenRows(4);
    for (var c = 1; c <= DEP_HEADERS.length; c++) s.autoResizeColumn(c);
  });
}

function exportInventoryReport_(sheet, hotelName, hotelCity) {
  sheet.setName('مخزون ' + hotelName);
  var inventory = getRoomInventory(hotelName, hotelCity);
  
  var headers = ['التاريخ', 'ثنائي', 'ثلاثي', 'رباعي', 'إجمالي الغرف', 'إجمالي الأسرّة'];
  sheet.appendRow(headers);
  
  var dates = Object.keys(inventory).sort();
  dates.forEach(function(d) {
    var inv = inventory[d];
    sheet.appendRow([d, inv.dbl, inv.tri, inv.quad, inv.rooms, inv.beds]);
  });
  
  formatExportSheet_(sheet, headers.length);
}

function exportDashboardReport_(sheet) {
  sheet.setName('لوحة التحكم');
  var stats = getDashboardStats();
  
  var headers = ['الفندق', 'المدينة', 'الإجمالي', 'وصلوا', 'لم يصلوا', 'متأخرون', 'No Show', 'نسبة Check-in'];
  sheet.appendRow(headers);
  
  for (var key in stats) {
    var s = stats[key];
    sheet.appendRow([s.hotel, s.city, s.total, s.arrived, s.pending, s.delayed, s.noShow || 0, s.checkinPct + '%']);
  }
  
  formatExportSheet_(sheet, headers.length);
}

function exportForecastReport_(ss) {
  var forecast = getForecastData();
  
  var first = true;
  for (var hotel in forecast) {
    var sheet;
    if (first) { sheet = ss.getActiveSheet(); first = false; }
    else { sheet = ss.insertSheet(); }
    
    var shortName = hotel.length > 30 ? hotel.substring(0, 30) : hotel;
    sheet.setName(shortName);
    
    var headers = ['التاريخ', 'الوصول المتوقع', 'المغادرة المتوقعة', 'الإشغال', 'وصول مبكر', 'وصول متأخر', 'No Show'];
    sheet.appendRow(['الفندق: ' + hotel + ' | المدينة: ' + (forecast[hotel].city || '')]);
    sheet.appendRow(headers);
    
    var dates = Object.keys(forecast[hotel].dates).sort();
    dates.forEach(function(d) {
      var f = forecast[hotel].dates[d];
      sheet.appendRow([d, f.arrivals, f.departures, f.occupancy, f.early, f.late, f.noshow || 0]);
    });
    
    formatExportSheet_(sheet, headers.length);
  }
  
  var summarySheet = ss.insertSheet('ملخص');
  summarySheet.appendRow(['الفندق', 'المدينة', 'إجمالي الحجاج', 'ذروة الإشغال', 'وصول مبكر', 'وصول متأخر', 'No Show']);
  
  for (var hotel in forecast) {
    var dates = forecast[hotel].dates;
    var totalArr = 0, peakOcc = 0, totalEarly = 0, totalLate = 0, totalNoShow = 0;
    for (var d in dates) {
      totalArr += dates[d].arrivals;
      if (dates[d].occupancy > peakOcc) peakOcc = dates[d].occupancy;
      totalEarly += dates[d].early;
      totalLate += dates[d].late;
      totalNoShow += (dates[d].noshow || 0);
    }
    summarySheet.appendRow([hotel, forecast[hotel].city, totalArr, peakOcc, totalEarly, totalLate, totalNoShow]);
  }
  
  formatExportSheet_(summarySheet, 7);
  ss.setActiveSheet(summarySheet);
  ss.moveActiveSheet(1);
}

function formatExportSheet_(sheet, colCount) {
  if (sheet.getLastRow() === 0) return;
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#064e3b');
  headerRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  for (var c = 1; c <= colCount; c++) {
    sheet.autoResizeColumn(c);
  }
}

// ============================================================
// DATA: FORECAST
// ============================================================

function getForecastData() {
  var mainSS = SpreadsheetApp.openById(HOTEL_CONFIG.SPREADSHEET_ID);
  var journeySheet = findSheet_(mainSS, HOTEL_CONFIG.SHEETS.JOURNEY);
  var journeyData = journeySheet.getDataRange().getValues();
  var J = HOTEL_CONFIG.JOURNEY_COLS;
  
  var forecast = {};
  
  for (var i = 1; i < journeyData.length; i++) {
    var row = journeyData[i];
    var firstHouse = row[J.FIRST_HOUSE];
    var hotel = getHotelName_(row, firstHouse);
    if (!hotel || hotel === 'NULL') continue;
    
    if (!forecast[hotel]) forecast[hotel] = { city: firstHouse, dates: {} };
    
    var arrDate = formatDate_(row[J.ARRIVAL_DATE]);
    var checkIn = formatDate_(row[J.FIRST_HOUSE_START]);
    var checkOut = formatDate_(row[J.FIRST_HOUSE_END]);
    
    // Check No Show (Madinah)
    var isNoShow = false;
    if (firstHouse === 'Madina' && arrDate && checkOut) {
      var ad = new Date(arrDate); ad.setHours(0,0,0,0);
      var ce = new Date(checkOut); ce.setHours(0,0,0,0);
      if (ad > ce) isNoShow = true;
    }
    
    if (arrDate) {
      if (!forecast[hotel].dates[arrDate]) forecast[hotel].dates[arrDate] = { arrivals: 0, departures: 0, occupancy: 0, early: 0, late: 0, noshow: 0 };
      
      if (isNoShow) {
        forecast[hotel].dates[arrDate].noshow++;
      } else {
        forecast[hotel].dates[arrDate].arrivals++;
        
        if (arrDate > checkIn) forecast[hotel].dates[arrDate].late++;
        if (arrDate === checkIn && firstHouse === 'Madina') {
          var exp = calculateExpectedArrival_(row, firstHouse);
          if (exp && parseInt(exp.split(':')[0]) < 13) {
            forecast[hotel].dates[arrDate].early++;
          }
        }
      }
    }
    
    if (!isNoShow && checkIn && checkOut) {
      var d1 = new Date(checkIn);
      var d2 = new Date(checkOut);
      for (var d = new Date(d1); d < d2; d.setDate(d.getDate() + 1)) {
        var dk = Utilities.formatDate(d, 'Asia/Riyadh', 'yyyy-MM-dd');
        if (!forecast[hotel].dates[dk]) forecast[hotel].dates[dk] = { arrivals: 0, departures: 0, occupancy: 0, early: 0, late: 0, noshow: 0 };
        forecast[hotel].dates[dk].occupancy++;
      }
    }
    
    if (!isNoShow && checkOut) {
      if (!forecast[hotel].dates[checkOut]) forecast[hotel].dates[checkOut] = { arrivals: 0, departures: 0, occupancy: 0, early: 0, late: 0, noshow: 0 };
      forecast[hotel].dates[checkOut].departures++;
    }
  }
  
  return forecast;
}
