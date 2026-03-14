/**
 * ══════════════════════════════════════════════════════════════════
 * سكريبت إدارة الحج والعمرة - الإصدار v5
 * ══════════════════════════════════════════════════════════════════
 * 
 * المكونات:
 *   1. فلترة الفنادق (بناءً على المدينة) + الاسم الإنجليزي التلقائي
 *   2. جلب بيانات الرحلات من API
 * 
 * التحديثات v5:
 *   - تصحيح أرقام أعمدة الرحلات لتتطابق مع هيكل الشيت الفعلي
 *   - ذهاب 1 (ترانزيت): V-AB (22-28)
 *   - ذهاب 2 (أو مباشر): AC-AI (29-35)
 *   - عودة 1: AJ-AP (36-42)
 *   - عودة 2 (ترانزيت): AQ-AW (43-49)
 * 
 * ══════════════════════════════════════════════════════════════════
 */


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    الإعدادات العامة                           ║
// ╚═══════════════════════════════════════════════════════════════╝

// ============ إعدادات الفنادق ============
const HOTEL_CONFIG = {
  MAIN_SHEET: "الباقات",
  HOTELS_SHEET: "الفنادق",
  DATA_START_ROW: 3,
  HOTELS_START_ROW: 2,
  
  // الفندق الأول
  COL_CITY_1: 12,           // العمود L
  COL_HOTEL_1: 13,          // العمود M (عربي)
  COL_HOTEL_EN_1: 14,       // العمود N (إنجليزي)
  
  // الفندق الثاني
  COL_CITY_2: 27,           // العمود AA
  COL_HOTEL_2: 28,          // العمود AB (عربي)
  COL_HOTEL_EN_2: 29,       // العمود AC (إنجليزي)
  
  // الفندق الثالث
  COL_CITY_3: 42,           // العمود AP
  COL_HOTEL_3: 43,          // العمود AQ (عربي)
  COL_HOTEL_EN_3: 44        // العمود AR (إنجليزي)
};

// مصفوفة أزواج (مدينة، فندق عربي، فندق إنجليزي)
const HOTEL_PAIRS = [
  { cityCol: HOTEL_CONFIG.COL_CITY_1, hotelCol: HOTEL_CONFIG.COL_HOTEL_1, hotelEnCol: HOTEL_CONFIG.COL_HOTEL_EN_1, name: "الأول" },
  { cityCol: HOTEL_CONFIG.COL_CITY_2, hotelCol: HOTEL_CONFIG.COL_HOTEL_2, hotelEnCol: HOTEL_CONFIG.COL_HOTEL_EN_2, name: "الثاني" },
  { cityCol: HOTEL_CONFIG.COL_CITY_3, hotelCol: HOTEL_CONFIG.COL_HOTEL_3, hotelEnCol: HOTEL_CONFIG.COL_HOTEL_EN_3, name: "الثالث" }
];

// ============ إعدادات الرحلات ============
const FLIGHT_CONFIG = {
  API_KEY: "1d82a75bd2msh6da5259e10fbb77p1229ebjsne04353e515dd",
  API_HOST: "aerodatabox.p.rapidapi.com",
  SHEET_NAME: "الطيران",
  DATA_START_ROW: 3
};

// مجموعات الرحلات (4 رحلات) - مصححة حسب هيكل الشيت الفعلي
const FLIGHT_GROUPS = [
  // ذهاب 1 (ترانزيت): V-AB
  { 
    name: "ذهاب 1", 
    flightNo: 22,      // V - رقم الرحلة
    dateTakeoff: 23,   // W - تاريخ الإقلاع
    timeTakeoff: 24,   // X - وقت الإقلاع
    from: 25,          // Y - من
    to: 26,            // Z - إلى
    dateLanding: 27,   // AA - تاريخ الهبوط
    timeLanding: 28    // AB - وقت الهبوط
  },
  // ذهاب 2 (أو مباشر): AC-AI
  { 
    name: "ذهاب 2", 
    flightNo: 29,      // AC - رقم الرحلة
    dateTakeoff: 30,   // AD - تاريخ الإقلاع
    timeTakeoff: 31,   // AE - وقت الإقلاع
    from: 32,          // AF - من
    to: 33,            // AG - إلى
    dateLanding: 34,   // AH - تاريخ الهبوط
    timeLanding: 35    // AI - وقت الهبوط
  },
  // عودة 1: AJ-AP
  { 
    name: "عودة 1", 
    flightNo: 36,      // AJ - رقم الرحلة
    dateTakeoff: 37,   // AK - تاريخ الإقلاع
    timeTakeoff: 38,   // AL - وقت الإقلاع
    from: 39,          // AM - من
    to: 40,            // AN - إلى
    dateLanding: 41,   // AO - تاريخ الهبوط
    timeLanding: 42    // AP - وقت الهبوط
  },
  // عودة 2 (ترانزيت): AQ-AW
  { 
    name: "عودة 2", 
    flightNo: 43,      // AQ - رقم الرحلة
    dateTakeoff: 44,   // AR - تاريخ الإقلاع
    timeTakeoff: 45,   // AS - وقت الإقلاع
    from: 46,          // AT - من
    to: 47,            // AU - إلى
    dateLanding: 48,   // AV - تاريخ الهبوط
    timeLanding: 49    // AW - وقت الهبوط
  }
];


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    القائمة الرئيسية                           ║
// ╚═══════════════════════════════════════════════════════════════╝

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // قائمة الفنادق
  ui.createMenu('🏨 إدارة الفنادق')
    .addItem('🔄 تحديث جميع الصفوف', 'updateAllHotelRows')
    .addItem('🧪 فحص الإعداد', 'testHotelSetup')
    .addToUi();
    addRoomTypeMenu();
  
  // قائمة الرحلات
  ui.createMenu('✈️ بيانات الرحلات')
    .addItem('🔄 تحديث جميع الرحلات', 'updateAllFlights')
    .addItem('📍 تحديث الصف الحالي', 'updateCurrentRow')
    .addSeparator()
    .addItem('✨ إكمال الفارغة فقط', 'fillEmptyFlightsOnly')
    .addItem('⏩ إكمال من الصف الحالي', 'fillFromCurrentRow')
    .addSeparator()
    .addItem('🧪 اختبار الاتصال', 'testAPIConnection')
    .addItem('🔍 تشخيص رحلة', 'debugFlightData')
    .addItem('📋 فحص هيكل الأعمدة', 'verifyColumnStructure')
    .addToUi();

    addB2CFlightMenu();
    addTourGuideMenu();
  
  // قائمة مخيم مِنى
  addMinaCampMenu();
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    الـ Trigger الرئيسي                        ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * Installable Trigger - للتعديلات (الفنادق + الرحلات)
 */
function onEditInstallable(e) {
  if (!e || !e.range) return;
  tgOnEdit(e);
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // معالجة تعديلات الفنادق
  if (handleHotelEdit(sheet, row, col)) return;
  
  // معالجة تعديلات الرحلات
  if (sheet.getName() === FLIGHT_CONFIG.SHEET_NAME && row >= FLIGHT_CONFIG.DATA_START_ROW) {
    for (const group of FLIGHT_GROUPS) {
      // التفعيل عند إدخال رقم الرحلة أو تاريخ الإقلاع
      if (col === group.flightNo || col === group.dateTakeoff) {
        updateFlightGroup(sheet, row, group);
        return;
      }
    }
  }
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    دوال الفنادق                               ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * الحصول على بيانات الفنادق (الاسم العربي والإنجليزي) بناءً على المدينة
 */
function getHotelsData(city) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hotelsSheet = ss.getSheetByName(HOTEL_CONFIG.HOTELS_SHEET);
  
  if (!hotelsSheet) return { names: [], map: {} };
  
  const lastRow = hotelsSheet.getLastRow();
  if (lastRow < HOTEL_CONFIG.HOTELS_START_ROW) return { names: [], map: {} };
  
  // قراءة الأعمدة الثلاثة: A (عربي)، B (المدينة)، C (إنجليزي)
  const hotelsData = hotelsSheet.getRange(
    HOTEL_CONFIG.HOTELS_START_ROW, 1,
    lastRow - HOTEL_CONFIG.HOTELS_START_ROW + 1, 3
  ).getValues();
  
  const hotelNames = [];
  const hotelMap = {}; // خريطة: اسم عربي -> اسم إنجليزي
  const cleanCity = String(city).trim().toLowerCase();
  
  for (let i = 0; i < hotelsData.length; i++) {
    const hotelNameAr = String(hotelsData[i][0]).trim();
    const hotelCity = String(hotelsData[i][1]).trim().toLowerCase();
    const hotelNameEn = String(hotelsData[i][2]).trim();
    
    if (!hotelNameAr) continue;
    
    if (hotelCity === cleanCity) {
      hotelNames.push(hotelNameAr);
      hotelMap[hotelNameAr] = hotelNameEn;
    }
  }
  
  return { names: hotelNames, map: hotelMap };
}

/**
 * الحصول على الاسم الإنجليزي للفندق
 */
function getHotelEnglishName(hotelNameAr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hotelsSheet = ss.getSheetByName(HOTEL_CONFIG.HOTELS_SHEET);
  
  if (!hotelsSheet) return "";
  
  const lastRow = hotelsSheet.getLastRow();
  if (lastRow < HOTEL_CONFIG.HOTELS_START_ROW) return "";
  
  const hotelsData = hotelsSheet.getRange(
    HOTEL_CONFIG.HOTELS_START_ROW, 1,
    lastRow - HOTEL_CONFIG.HOTELS_START_ROW + 1, 3
  ).getValues();
  
  const cleanName = String(hotelNameAr).trim();
  
  for (let i = 0; i < hotelsData.length; i++) {
    const nameAr = String(hotelsData[i][0]).trim();
    const nameEn = String(hotelsData[i][2]).trim();
    
    if (nameAr === cleanName) {
      return nameEn;
    }
  }
  
  return "";
}

/**
 * تحديث القائمة المنسدلة للفندق
 */
function updateHotelDropdown(sheet, row, cityCol, hotelCol, hotelEnCol) {
  try {
    const city = sheet.getRange(row, cityCol).getValue();
    const hotelCell = sheet.getRange(row, hotelCol);
    const hotelEnCell = sheet.getRange(row, hotelEnCol);
    
    if (!city || String(city).trim() === "") {
      hotelCell.clearDataValidations();
      hotelCell.setValue("");
      hotelEnCell.setValue("");
      return;
    }
    
    const hotelsData = getHotelsData(city);
    
    if (hotelsData.names.length === 0) {
      hotelCell.clearDataValidations();
      hotelCell.setValue("");
      hotelEnCell.setValue("");
      return;
    }
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(hotelsData.names, true)
      .setAllowInvalid(false)
      .build();
    
    hotelCell.setDataValidation(rule);
    
  } catch (error) {
    console.error("خطأ في updateHotelDropdown: " + error.message);
  }
}

/**
 * تحديث الاسم الإنجليزي عند اختيار الفندق
 */
function updateHotelEnglishName(sheet, row, hotelCol, hotelEnCol) {
  try {
    const hotelNameAr = sheet.getRange(row, hotelCol).getValue();
    const hotelEnCell = sheet.getRange(row, hotelEnCol);
    
    if (!hotelNameAr || String(hotelNameAr).trim() === "") {
      hotelEnCell.setValue("");
      return;
    }
    
    const hotelNameEn = getHotelEnglishName(hotelNameAr);
    hotelEnCell.setValue(hotelNameEn);
    
  } catch (error) {
    console.error("خطأ في updateHotelEnglishName: " + error.message);
  }
}

/**
 * معالجة تعديلات الفنادق
 */
function handleHotelEdit(sheet, row, col) {
  if (sheet.getName() !== HOTEL_CONFIG.MAIN_SHEET) return false;
  if (row < HOTEL_CONFIG.DATA_START_ROW) return false;
  
  for (const pair of HOTEL_PAIRS) {
    // إذا تغيرت المدينة → تحديث القائمة المنسدلة
    if (col === pair.cityCol) {
      updateHotelDropdown(sheet, row, pair.cityCol, pair.hotelCol, pair.hotelEnCol);
      return true;
    }
    
    // إذا تغير الفندق (العربي) → تحديث الاسم الإنجليزي
    if (col === pair.hotelCol) {
      updateHotelEnglishName(sheet, row, pair.hotelCol, pair.hotelEnCol);
      return true;
    }
  }
  
  return false;
}

/**
 * تحديث جميع صفوف الفنادق
 */
function updateAllHotelRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(HOTEL_CONFIG.MAIN_SHEET);
  
  if (!sheet) {
    ss.toast("شيت '" + HOTEL_CONFIG.MAIN_SHEET + "' غير موجود!", "❌ خطأ", 5);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < HOTEL_CONFIG.DATA_START_ROW) {
    ss.toast("لا توجد بيانات للتحديث", "⚠️ تنبيه", 3);
    return;
  }
  
  let updatedCount = 0;
  
  for (let row = HOTEL_CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    for (const pair of HOTEL_PAIRS) {
      // تحديث القائمة المنسدلة
      updateHotelDropdown(sheet, row, pair.cityCol, pair.hotelCol, pair.hotelEnCol);
      // تحديث الاسم الإنجليزي
      updateHotelEnglishName(sheet, row, pair.hotelCol, pair.hotelEnCol);
    }
    updatedCount++;
  }
  
  ss.toast("تم التحديث بنجاح!\nعدد الصفوف: " + updatedCount, "✅", 5);
}

/**
 * فحص إعداد الفنادق
 */
function testHotelSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  let report = "📋 تقرير فحص الإعداد\n";
  report += "═══════════════════════════════\n\n";
  
  const mainSheet = ss.getSheetByName(HOTEL_CONFIG.MAIN_SHEET);
  const hotelsSheet = ss.getSheetByName(HOTEL_CONFIG.HOTELS_SHEET);
  
  report += mainSheet 
    ? "✅ شيت الباقات: موجود (" + mainSheet.getLastRow() + " صف)\n"
    : "❌ شيت الباقات: غير موجود!\n";
    
  report += hotelsSheet
    ? "✅ شيت الفنادق: موجود (" + (hotelsSheet.getLastRow() - 1) + " فندق)\n"
    : "❌ شيت الفنادق: غير موجود!\n";
  
  report += "\n📍 هيكل شيت الفنادق:\n";
  report += "   العمود A: اسم الفندق (عربي)\n";
  report += "   العمود B: المدينة (Mak/Med)\n";
  report += "   العمود C: اسم الفندق (إنجليزي)\n";
  
  report += "\n📍 أعمدة الفنادق في الباقات:\n";
  report += "   الفندق 1: المدينة L → الفندق M → الإنجليزي N\n";
  report += "   الفندق 2: المدينة AA → الفندق AB → الإنجليزي AC\n";
  report += "   الفندق 3: المدينة AP → الفندق AQ → الإنجليزي AR\n";
  
  report += "\n🔍 اختبار الفلترة:\n";
  const makData = getHotelsData("Mak");
  const medData = getHotelsData("Med");
  report += "   Mak → " + makData.names.length + " فندق\n";
  report += "   Med → " + medData.names.length + " فندق\n";
  
  if (makData.names.length > 0) {
    report += "\n   مثال (مكة): " + makData.names[0] + " → " + makData.map[makData.names[0]] + "\n";
  }
  if (medData.names.length > 0) {
    report += "   مثال (المدينة): " + medData.names[0] + " → " + medData.map[medData.names[0]] + "\n";
  }
  
  ui.alert(report);
}


// ╔═══════════════════════════════════════════════════════════════╗
// ║                    دوال الرحلات                               ║
// ╚═══════════════════════════════════════════════════════════════╝

/**
 * جلب بيانات رحلة من API
 */
function fetchFlightData(flightNumber, flightDate) {
  try {
    const cleanFlightNo = String(flightNumber).trim().toUpperCase().replace(/\s+/g, '');
    const dateStr = formatDateForAPI(flightDate);
    if (!dateStr) return null;
    
    const url = `https://aerodatabox.p.rapidapi.com/flights/number/${cleanFlightNo}/${dateStr}?withAircraftImage=false&withLocation=false&dateLocalRole=Both`;
    
    const options = {
      method: "GET",
      headers: {
        "X-RapidAPI-Key": FLIGHT_CONFIG.API_KEY,
        "X-RapidAPI-Host": FLIGHT_CONFIG.API_HOST
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) return null;
    
    const data = JSON.parse(response.getContentText());
    if (!data || data.length === 0) return null;
    
    const flight = data[0];
    
    return {
      from: flight.departure?.airport?.iata || flight.departure?.airport?.icao || "",
      to: flight.arrival?.airport?.iata || flight.arrival?.airport?.icao || "",
      departureTime: extractTime(flight.departure?.scheduledTime?.local || flight.departure?.scheduledTime?.utc),
      arrivalTime: extractTime(flight.arrival?.scheduledTime?.local || flight.arrival?.scheduledTime?.utc),
      departureDate: extractDate(flight.departure?.scheduledTime?.local || flight.departure?.scheduledTime?.utc),
      arrivalDate: extractDate(flight.arrival?.scheduledTime?.local || flight.arrival?.scheduledTime?.utc)
    };
    
  } catch (error) {
    console.error("خطأ في جلب بيانات الرحلة: " + error.message);
    return null;
  }
}

/**
 * تنسيق التاريخ لـ API
 */
function formatDateForAPI(dateValue) {
  try {
    let date;
    
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      const dateStr = String(dateValue).trim();
      const monthNames = {
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
      };
      
      const parts = dateStr.toLowerCase().split(/\s+/);
      if (parts.length >= 2) {
        const day = parts[0].padStart(2, '0');
        const month = monthNames[parts[1].substring(0, 3)];
        const year = parts[2] || new Date().getFullYear();
        if (month) return `${year}-${month}-${day}`;
      }
      
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) return null;
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
  } catch (e) {
    return null;
  }
}

/**
 * استخراج الوقت من سلسلة التاريخ/الوقت
 */
function extractTime(dateTimeStr) {
  if (!dateTimeStr) return "";
  try {
    const match = String(dateTimeStr).match(/[\sT](\d{2}:\d{2})/);
    return match ? match[1] : "";
  } catch (e) {
    return "";
  }
}

/**
 * استخراج التاريخ من سلسلة التاريخ/الوقت
 */
function extractDate(dateTimeStr) {
  if (!dateTimeStr) return "";
  try {
    const match = String(dateTimeStr).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      return `${parseInt(match[3])} ${months[parseInt(match[2]) - 1]}`;
    }
    return "";
  } catch (e) {
    return "";
  }
}

/**
 * تحديث مجموعة رحلة واحدة
 */
function updateFlightGroup(sheet, row, group) {
  const flightNo = sheet.getRange(row, group.flightNo).getValue();
  const flightDate = sheet.getRange(row, group.dateTakeoff).getValue();
  
  if (!flightNo || !flightDate) return;
  
  const data = fetchFlightData(flightNo, flightDate);
  
  if (!data) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "لم يتم العثور على بيانات للرحلة: " + flightNo + " (" + group.name + ")",
      "⚠️ تنبيه", 5
    );
    return;
  }
  
  sheet.getRange(row, group.timeTakeoff).setValue(data.departureTime);
  sheet.getRange(row, group.from).setValue(data.from);
  sheet.getRange(row, group.to).setValue(data.to);
  sheet.getRange(row, group.dateLanding).setValue(data.arrivalDate);
  sheet.getRange(row, group.timeLanding).setValue(data.arrivalTime);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ " + group.name + ": " + flightNo, "تم جلب البيانات", 3);
}

/**
 * تحديث جميع الرحلات
 */
function updateAllFlights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FLIGHT_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    ss.toast("شيت '" + FLIGHT_CONFIG.SHEET_NAME + "' غير موجود!", "❌ خطأ", 5);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < FLIGHT_CONFIG.DATA_START_ROW) {
    ss.toast("لا توجد بيانات للتحديث", "⚠️ تنبيه", 3);
    return;
  }
  
  let successCount = 0, failCount = 0, totalProcessed = 0;
  
  for (let row = FLIGHT_CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    for (const group of FLIGHT_GROUPS) {
      const flightNo = sheet.getRange(row, group.flightNo).getValue();
      const flightDate = sheet.getRange(row, group.dateTakeoff).getValue();
      
      if (!flightNo || !flightDate) continue;
      
      totalProcessed++;
      const data = fetchFlightData(flightNo, flightDate);
      
      if (data) {
        sheet.getRange(row, group.timeTakeoff).setValue(data.departureTime);
        sheet.getRange(row, group.from).setValue(data.from);
        sheet.getRange(row, group.to).setValue(data.to);
        sheet.getRange(row, group.dateLanding).setValue(data.arrivalDate);
        sheet.getRange(row, group.timeLanding).setValue(data.arrivalTime);
        successCount++;
      } else {
        failCount++;
      }
      
      Utilities.sleep(500);
    }
  }
  
  ss.toast("نجح: " + successCount + " | فشل: " + failCount + " | الإجمالي: " + totalProcessed, "✅ اكتمل التحديث", 5);
}

/**
 * تحديث الصف الحالي
 */
function updateCurrentRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const row = ss.getActiveCell().getRow();
  
  if (sheet.getName() !== FLIGHT_CONFIG.SHEET_NAME) {
    ss.toast("يجب أن تكون في شيت '" + FLIGHT_CONFIG.SHEET_NAME + "'", "⚠️", 3);
    return;
  }
  
  if (row < FLIGHT_CONFIG.DATA_START_ROW) {
    ss.toast("اختر صف بيانات (من الصف " + FLIGHT_CONFIG.DATA_START_ROW + " فما فوق)", "⚠️", 3);
    return;
  }
  
  let updated = 0;
  
  for (const group of FLIGHT_GROUPS) {
    const flightNo = sheet.getRange(row, group.flightNo).getValue();
    const flightDate = sheet.getRange(row, group.dateTakeoff).getValue();
    
    if (!flightNo || !flightDate) continue;
    
    const data = fetchFlightData(flightNo, flightDate);
    
    if (data) {
      sheet.getRange(row, group.timeTakeoff).setValue(data.departureTime);
      sheet.getRange(row, group.from).setValue(data.from);
      sheet.getRange(row, group.to).setValue(data.to);
      sheet.getRange(row, group.dateLanding).setValue(data.arrivalDate);
      sheet.getRange(row, group.timeLanding).setValue(data.arrivalTime);
      updated++;
    }
    
    Utilities.sleep(500);
  }
  
  ss.toast("تم تحديث " + updated + " رحلات في الصف " + row, "✅", 3);
}

/**
 * إكمال الرحلات الفارغة فقط
 */
function fillEmptyFlightsOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(FLIGHT_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    ss.toast("شيت '" + FLIGHT_CONFIG.SHEET_NAME + "' غير موجود!", "❌ خطأ", 5);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < FLIGHT_CONFIG.DATA_START_ROW) {
    ss.toast("لا توجد بيانات", "⚠️", 3);
    return;
  }
  
  let filledCount = 0, skippedCount = 0;
  const startTime = new Date().getTime();
  const MAX_TIME = 5 * 60 * 1000;
  
  for (let row = FLIGHT_CONFIG.DATA_START_ROW; row <= lastRow; row++) {
    if (new Date().getTime() - startTime > MAX_TIME) {
      ss.toast("⏱️ توقف مؤقت - أكمل " + filledCount + " رحلة\nشغّل الدالة مرة أخرى", "تنبيه", 10);
      return;
    }
    
    for (const group of FLIGHT_GROUPS) {
      const flightNo = sheet.getRange(row, group.flightNo).getValue();
      const flightDate = sheet.getRange(row, group.dateTakeoff).getValue();
      const existingTime = sheet.getRange(row, group.timeTakeoff).getValue();
      
      if (!flightNo || !flightDate) continue;
      if (existingTime && String(existingTime).trim() !== "") {
        skippedCount++;
        continue;
      }
      
      const data = fetchFlightData(flightNo, flightDate);
      
      if (data) {
        sheet.getRange(row, group.timeTakeoff).setValue(data.departureTime);
        sheet.getRange(row, group.from).setValue(data.from);
        sheet.getRange(row, group.to).setValue(data.to);
        sheet.getRange(row, group.dateLanding).setValue(data.arrivalDate);
        sheet.getRange(row, group.timeLanding).setValue(data.arrivalTime);
        filledCount++;
      }
      
      Utilities.sleep(300);
    }
  }
  
  ss.toast("✅ اكتمل!\nمُلئت: " + filledCount + " | تُخطيت: " + skippedCount, "النتيجة", 5);
}

/**
 * إكمال من الصف الحالي
 */
function fillFromCurrentRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const startRow = ss.getActiveCell().getRow();
  
  if (sheet.getName() !== FLIGHT_CONFIG.SHEET_NAME) {
    ss.toast("يجب أن تكون في شيت '" + FLIGHT_CONFIG.SHEET_NAME + "'", "⚠️", 3);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  let filledCount = 0;
  const startTime = new Date().getTime();
  const MAX_TIME = 5 * 60 * 1000;
  
  for (let row = startRow; row <= lastRow; row++) {
    if (new Date().getTime() - startTime > MAX_TIME) {
      ss.toast("⏱️ توقف في الصف " + row + "\nأكمل " + filledCount + " رحلة", "تنبيه", 10);
      return;
    }
    
    for (const group of FLIGHT_GROUPS) {
      const flightNo = sheet.getRange(row, group.flightNo).getValue();
      const flightDate = sheet.getRange(row, group.dateTakeoff).getValue();
      const existingTime = sheet.getRange(row, group.timeTakeoff).getValue();
      
      if (!flightNo || !flightDate) continue;
      if (existingTime && String(existingTime).trim() !== "") continue;
      
      const data = fetchFlightData(flightNo, flightDate);
      
      if (data) {
        sheet.getRange(row, group.timeTakeoff).setValue(data.departureTime);
        sheet.getRange(row, group.from).setValue(data.from);
        sheet.getRange(row, group.to).setValue(data.to);
        sheet.getRange(row, group.dateLanding).setValue(data.arrivalDate);
        sheet.getRange(row, group.timeLanding).setValue(data.arrivalTime);
        filledCount++;
      }
      
      Utilities.sleep(300);
    }
  }
  
  ss.toast("✅ اكتمل! مُلئت: " + filledCount + " رحلة", "النتيجة", 5);
}

/**
 * اختبار اتصال API
 */
function testAPIConnection() {
  const ui = SpreadsheetApp.getUi();
  const testFlight = "SV22";
  const testDate = "2026-05-17";
  
  const url = `https://aerodatabox.p.rapidapi.com/flights/number/${testFlight}/${testDate}?withAircraftImage=false&withLocation=false&dateLocalRole=Both`;
  
  try {
    const options = {
      method: "GET",
      headers: {
        "X-RapidAPI-Key": FLIGHT_CONFIG.API_KEY,
        "X-RapidAPI-Host": FLIGHT_CONFIG.API_HOST
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      ui.alert("✅ الاتصال ناجح!\n\nرحلة الاختبار: " + testFlight + "\nعدد النتائج: " + (data ? data.length : 0));
    } else if (responseCode === 403) {
      ui.alert("❌ خطأ في API Key");
    } else if (responseCode === 429) {
      ui.alert("⚠️ تجاوزت حد الطلبات");
    } else {
      ui.alert("❌ خطأ: " + responseCode);
    }
  } catch (error) {
    ui.alert("❌ خطأ في الاتصال:\n\n" + error.message);
  }
}

/**
 * تشخيص بيانات رحلة - لعرض البيانات الخام من API
 */
function debugFlightData() {
  const ui = SpreadsheetApp.getUi();
  
  // طلب رقم الرحلة
  const flightResponse = ui.prompt('🔍 تشخيص رحلة', 'أدخل رقم الرحلة (مثال: VF205):', ui.ButtonSet.OK_CANCEL);
  if (flightResponse.getSelectedButton() !== ui.Button.OK) return;
  const flightNo = flightResponse.getResponseText().trim().toUpperCase().replace(/\s+/g, '');
  
  // طلب التاريخ
  const dateResponse = ui.prompt('🔍 تشخيص رحلة', 'أدخل التاريخ (مثال: 2026-05-31):', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const dateStr = dateResponse.getResponseText().trim();
  
  // جلب البيانات من API
  const url = `https://aerodatabox.p.rapidapi.com/flights/number/${flightNo}/${dateStr}?withAircraftImage=false&withLocation=false&dateLocalRole=Both`;
  
  try {
    const options = {
      method: "GET",
      headers: {
        "X-RapidAPI-Key": FLIGHT_CONFIG.API_KEY,
        "X-RapidAPI-Host": FLIGHT_CONFIG.API_HOST
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      ui.alert("❌ خطأ في الاستجابة: " + responseCode);
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data || data.length === 0) {
      ui.alert("⚠️ لم يتم العثور على بيانات للرحلة: " + flightNo);
      return;
    }
    
    const flight = data[0];
    
    // بناء تقرير مفصل
    let report = "📋 بيانات الرحلة: " + flightNo + "\n";
    report += "═══════════════════════════════\n\n";
    
    report += "【 المغادرة 】\n";
    report += "  المطار: " + (flight.departure?.airport?.iata || "غير متوفر") + "\n";
    report += "  الوقت المحلي: " + (flight.departure?.scheduledTime?.local || "غير متوفر") + "\n";
    report += "  الوقت UTC: " + (flight.departure?.scheduledTime?.utc || "غير متوفر") + "\n\n";
    
    report += "【 الوصول 】\n";
    report += "  المطار: " + (flight.arrival?.airport?.iata || "غير متوفر") + "\n";
    report += "  الوقت المحلي: " + (flight.arrival?.scheduledTime?.local || "غير متوفر") + "\n";
    report += "  الوقت UTC: " + (flight.arrival?.scheduledTime?.utc || "غير متوفر") + "\n\n";
    
    report += "【 البيانات المستخرجة 】\n";
    report += "  وقت الإقلاع: " + extractTime(flight.departure?.scheduledTime?.local) + "\n";
    report += "  تاريخ الإقلاع: " + extractDate(flight.departure?.scheduledTime?.local) + "\n";
    report += "  وقت الهبوط: " + extractTime(flight.arrival?.scheduledTime?.local) + "\n";
    report += "  تاريخ الهبوط: " + extractDate(flight.arrival?.scheduledTime?.local) + "\n";
    
    ui.alert(report);
    
    // طباعة البيانات الخام في السجل
    console.log("=== RAW API Response for " + flightNo + " ===");
    console.log(JSON.stringify(data, null, 2));
    
  } catch (error) {
    ui.alert("❌ خطأ: " + error.message);
  }
}

/**
 * تشخيص بيانات رحلة - لعرض البيانات الخام من API
 */
function debugFlightData() {
  const ui = SpreadsheetApp.getUi();
  
  // طلب رقم الرحلة
  const flightResponse = ui.prompt('🔍 تشخيص رحلة', 'أدخل رقم الرحلة (مثال: VF205):', ui.ButtonSet.OK_CANCEL);
  if (flightResponse.getSelectedButton() !== ui.Button.OK) return;
  const flightNo = flightResponse.getResponseText().trim();
  
  // طلب التاريخ
  const dateResponse = ui.prompt('🔍 تشخيص رحلة', 'أدخل التاريخ (مثال: 2026-05-31):', ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const flightDate = dateResponse.getResponseText().trim();
  
  const url = `https://aerodatabox.p.rapidapi.com/flights/number/${flightNo}/${flightDate}?withAircraftImage=false&withLocation=false&dateLocalRole=Both`;
  
  try {
    const options = {
      method: "GET",
      headers: {
        "X-RapidAPI-Key": FLIGHT_CONFIG.API_KEY,
        "X-RapidAPI-Host": FLIGHT_CONFIG.API_HOST
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      ui.alert("❌ خطأ: " + responseCode);
      return;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data || data.length === 0) {
      ui.alert("❌ لا توجد بيانات لهذه الرحلة");
      return;
    }
    
    const flight = data[0];
    
    let report = "📋 بيانات الرحلة الخام من API\n";
    report += "═══════════════════════════════\n\n";
    report += "رقم الرحلة: " + flightNo + "\n";
    report += "التاريخ: " + flightDate + "\n\n";
    
    report += "【 المغادرة 】\n";
    report += "  المطار: " + (flight.departure?.airport?.iata || "غير متوفر") + "\n";
    report += "  الوقت المحلي: " + (flight.departure?.scheduledTime?.local || "غير متوفر") + "\n";
    report += "  الوقت UTC: " + (flight.departure?.scheduledTime?.utc || "غير متوفر") + "\n\n";
    
    report += "【 الوصول 】\n";
    report += "  المطار: " + (flight.arrival?.airport?.iata || "غير متوفر") + "\n";
    report += "  الوقت المحلي: " + (flight.arrival?.scheduledTime?.local || "غير متوفر") + "\n";
    report += "  الوقت UTC: " + (flight.arrival?.scheduledTime?.utc || "غير متوفر") + "\n\n";
    
    report += "【 البيانات المستخرجة 】\n";
    report += "  تاريخ الهبوط: " + extractDate(flight.arrival?.scheduledTime?.local || flight.arrival?.scheduledTime?.utc) + "\n";
    report += "  وقت الهبوط: " + extractTime(flight.arrival?.scheduledTime?.local || flight.arrival?.scheduledTime?.utc) + "\n";
    
    ui.alert(report);
    
    // طباعة البيانات الكاملة في السجل
    console.log("=== RAW API RESPONSE ===");
    console.log(JSON.stringify(data, null, 2));
    
  } catch (error) {
    ui.alert("❌ خطأ: " + error.message);
  }
}

/**
 * فحص هيكل الأعمدة - للتأكد من صحة الإعدادات
 */
function verifyColumnStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(FLIGHT_CONFIG.SHEET_NAME);
  
  if (!sheet) {
    ui.alert("❌ شيت '" + FLIGHT_CONFIG.SHEET_NAME + "' غير موجود!");
    return;
  }
  
  let report = "📋 فحص هيكل أعمدة الرحلات\n";
  report += "═══════════════════════════════\n\n";
  report += "صف البداية: " + FLIGHT_CONFIG.DATA_START_ROW + "\n\n";
  
  for (const group of FLIGHT_GROUPS) {
    report += "【 " + group.name + " 】\n";
    report += "  رقم الرحلة (عمود " + group.flightNo + "): " + sheet.getRange(2, group.flightNo).getValue() + "\n";
    report += "  تاريخ الإقلاع (عمود " + group.dateTakeoff + "): " + sheet.getRange(2, group.dateTakeoff).getValue() + "\n";
    report += "  وقت الإقلاع (عمود " + group.timeTakeoff + "): " + sheet.getRange(2, group.timeTakeoff).getValue() + "\n";
    report += "  من (عمود " + group.from + "): " + sheet.getRange(2, group.from).getValue() + "\n";
    report += "  إلى (عمود " + group.to + "): " + sheet.getRange(2, group.to).getValue() + "\n";
    report += "  تاريخ الهبوط (عمود " + group.dateLanding + "): " + sheet.getRange(2, group.dateLanding).getValue() + "\n";
    report += "  وقت الهبوط (عمود " + group.timeLanding + "): " + sheet.getRange(2, group.timeLanding).getValue() + "\n\n";
  }
  
  ui.alert(report);
}