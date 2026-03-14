/**
 * ══════════════════════════════════════════════════════════════════
 * تطبيق تجميع غرف المرشدين — شركة إكرام الضيف للسياحة
 * حج 1447هـ — 2026م
 * الإصدار: 2.2
 * ──────────────────────────────────────────────────────────────────
 * التعديلات عن v2.1:
 *   - دخول بجواز السفر (رابط واحد للجميع)
 *   - إبقاء ?g= للمسؤولين
 *   - تحديث generateGuideLinks بعمود جواز المرشد
 * ══════════════════════════════════════════════════════════════════
 */

var CONFIG = {
  SPREADSHEET_ID: '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s',
  
  SHEETS: {
    TOUR_GUIDE: 'Tour Guide',
    PILGRIMS: 'رحلة الحاج',
    ROOM_TYPE: 'Room Type Preview',
    GUIDE_GROUPING: 'Guide Grouping',
    GUIDE_LINKS: 'روابط المرشدين'
  },
  
  GUIDE: {
    NAME: 0,
    PASSPORT: 4,
    REGISTERED: 9,
    UNIQUE: 10
  },
  
  PILGRIM: {
    BOOKING_ID: 0,
    APPLICANT_ID: 5,
    GROUP_NUMBER: 6,
    NAME: 7,
    PASSPORT: 8,
    GENDER: 11
  },
  
  RT: {
    GROUP_NUMBER: 0,
    MED_HOTEL: 15,
    MED_TYPE: 16,
    MED_SHARED: 19,
    MAK1_HOTEL: 20,
    MAK1_TYPE: 21,
    MAK1_SHARED: 24,
    MAK2_HOTEL: 25,
    MAK2_TYPE: 26,
    MAK2_SHARED: 29
  },
  
  START_ROW: 2
};

var HOTEL_POSITIONS = [
  { hotel: 'MED_HOTEL', type: 'MED_TYPE', shared: 'MED_SHARED' },
  { hotel: 'MAK1_HOTEL', type: 'MAK1_TYPE', shared: 'MAK1_SHARED' },
  { hotel: 'MAK2_HOTEL', type: 'MAK2_TYPE', shared: 'MAK2_SHARED' }
];


// ══════════════════════════════════════════════════════════════════
// نقطة الدخول — v2.2
// ──────────────────────────────────────────────────────────────────
// إذا ?g= موجود → دخول مباشر (للمسؤولين)
// إذا لا → صفحة إدخال جواز السفر
// ══════════════════════════════════════════════════════════════════

function doGet(e) {
  var guideName = '';
  var mode = 'login'; // login = صفحة إدخال الجواز
  
  if (e && e.parameter && e.parameter.g) {
    guideName = decodeURIComponent(e.parameter.g);
    mode = 'direct';
  }
  
  var template = HtmlService.createTemplateFromFile('GuideIndex');
  template.guideName = guideName;
  template.mode = mode;
  
  return template.evaluate()
    .setTitle('إكرام الضيف — تجميع الغرف')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// ══════════════════════════════════════════════════════════════════
// الدالة الجديدة: التحقق من جواز المرشد — v2.2
// ──────────────────────────────────────────────────────────────────
// المُدخل: رقم جواز سفر المرشد
// المُخرج: {success, guideName} أو {success:false, error}
// المصدر: شيت "روابط المرشدين" عمود E (جواز المرشد) → عمود B (اسمه)
// ══════════════════════════════════════════════════════════════════

function loginWithPassport(passport) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.GUIDE_LINKS);
    
    if (!sheet) {
      return { success: false, error: 'شيت روابط المرشدين غير موجود' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: false, error: 'لا توجد بيانات مرشدين' };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var inputPassport = String(passport).trim().toUpperCase();
    
    for (var i = 0; i < data.length; i++) {
      var guidePassport = String(data[i][4]).trim().toUpperCase(); // عمود E (index 4)
      if (guidePassport && guidePassport === inputPassport) {
        var guideName = String(data[i][1]).trim(); // عمود B (اسم المرشد)
        if (guideName) {
          return { success: true, guideName: guideName };
        }
      }
    }
    
    return { success: false, error: 'رقم الجواز غير مسجل. تواصل مع إدارة العمليات.' };
    
  } catch (err) {
    return { success: false, error: 'خطأ في التحقق: ' + err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
// الدالة 1: جلب الفنادق
// ══════════════════════════════════════════════════════════════════

function getGuideHotels(guideName) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var guidePassports = getGuidePassports_(ss, guideName);
    if (guidePassports.length === 0) {
      return { success: false, error: 'لا يوجد حجاج مسجلين لهذا المرشد' };
    }
    var pilgrimData = getPilgrimsByPassports_(ss, guidePassports);
    var groupNumbers = {};
    for (var p in pilgrimData) {
      var gn = String(pilgrimData[p].groupNumber);
      if (gn) groupNumbers[gn] = true;
    }
    var roomTypeData = getRoomTypeByGroups_(ss, Object.keys(groupNumbers));
    var hotels = {};
    for (var passport in pilgrimData) {
      var pil = pilgrimData[passport];
      var gn = String(pil.groupNumber);
      var rtRow = roomTypeData[gn];
      if (!rtRow) continue;
      for (var i = 0; i < HOTEL_POSITIONS.length; i++) {
        var pos = HOTEL_POSITIONS[i];
        var hotelName = String(rtRow[CONFIG.RT[pos.hotel]] || '').trim();
        var sharedBeds = parseInt(rtRow[CONFIG.RT[pos.shared]]) || 0;
        if (hotelName && sharedBeds > 0) {
          if (!hotels[hotelName]) {
            hotels[hotelName] = { hotelName: hotelName, pilgrimCount: 0, alreadyGrouped: false };
          }
          hotels[hotelName].pilgrimCount++;
        }
      }
    }
    var groupedHotels = getGroupedHotels_(ss, guideName);
    for (var h in hotels) {
      if (groupedHotels[h]) hotels[h].alreadyGrouped = true;
    }
    var result = [];
    for (var h in hotels) result.push(hotels[h]);
    return { success: true, hotels: result, totalPilgrims: guidePassports.length };
  } catch (err) {
    return { success: false, error: 'خطأ في جلب البيانات: ' + err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
// الدالة 2: جلب حجاج المرشد لفندق معين
// ══════════════════════════════════════════════════════════════════

function getGuidePilgrims(guideName, hotelName) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var guidePassports = getGuidePassports_(ss, guideName);
    var pilgrimData = getPilgrimsByPassports_(ss, guidePassports);
    var groupNumbers = {};
    for (var p in pilgrimData) {
      groupNumbers[String(pilgrimData[p].groupNumber)] = true;
    }
    var roomTypeData = getRoomTypeByGroups_(ss, Object.keys(groupNumbers));
    
    var sharedPilgrims = [];
    for (var passport in pilgrimData) {
      var pil = pilgrimData[passport];
      var rtRow = roomTypeData[String(pil.groupNumber)];
      if (!rtRow) continue;
      for (var i = 0; i < HOTEL_POSITIONS.length; i++) {
        var pos = HOTEL_POSITIONS[i];
        var hn = String(rtRow[CONFIG.RT[pos.hotel]] || '').trim();
        var shared = parseInt(rtRow[CONFIG.RT[pos.shared]]) || 0;
        if (shared > 0 && normalizeHotel_(hn) === normalizeHotel_(hotelName)) {
          var roomType = String(rtRow[CONFIG.RT[pos.type]] || '');
          var capacity = getRoomCapacity_(roomType);
          pil.roomType = roomType;
          pil.capacity = capacity;
          pil.roomTypeLabel = getRoomTypeLabel_(capacity);
          sharedPilgrims.push(pil);
          break;
        }
      }
    }
    
    if (sharedPilgrims.length === 0) {
      return { success: false, error: 'لا يوجد حجاج مشتركين لهذا الفندق' };
    }
    
    var byBooking = {};
    for (var i = 0; i < sharedPilgrims.length; i++) {
      var p = sharedPilgrims[i];
      var bid = String(p.bookingId);
      if (!byBooking[bid]) byBooking[bid] = [];
      byBooking[bid].push(p);
    }
    
    var sectionMap = {};
    var prefilled = [];
    
    for (var bid in byBooking) {
      var members = byBooking[bid];
      var byGenderCap = {};
      for (var m = 0; m < members.length; m++) {
        var key = members[m].gender + '|' + members[m].capacity;
        if (!byGenderCap[key]) byGenderCap[key] = [];
        byGenderCap[key].push(members[m]);
      }
      for (var key in byGenderCap) {
        var group = byGenderCap[key];
        var cap = group[0].capacity;
        var isMultiple = members.length > 1;
        if (group.length === cap) {
          for (var g = 0; g < group.length; g++) {
            group[g].isFamily = isMultiple;
            group[g].familyBookingId = isMultiple ? bid : '';
          }
          prefilled.push({
            gender: group[0].gender, capacity: cap,
            roomTypeLabel: group[0].roomTypeLabel, bookingId: bid, pilgrims: group
          });
          continue;
        }
        if (!sectionMap[key]) {
          sectionMap[key] = {
            gender: group[0].gender, capacity: cap,
            roomTypeLabel: group[0].roomTypeLabel, pilgrims: []
          };
        }
        for (var g = 0; g < group.length; g++) {
          group[g].isFamily = isMultiple;
          group[g].familyBookingId = isMultiple ? bid : '';
          sectionMap[key].pilgrims.push(group[g]);
        }
      }
    }
    
    var sections = [];
    for (var key in sectionMap) sections.push(sectionMap[key]);
    sections.sort(function(a, b) {
      var go = { 'Male': 1, 'Female': 2 };
      if ((go[a.gender]||3) !== (go[b.gender]||3)) return (go[a.gender]||3) - (go[b.gender]||3);
      return b.capacity - a.capacity;
    });
    
    var totalPrefilled = 0, menCount = 0, womenCount = 0;
    for (var i = 0; i < prefilled.length; i++) totalPrefilled += prefilled[i].pilgrims.length;
    for (var i = 0; i < sections.length; i++) {
      if (sections[i].gender === 'Male') menCount += sections[i].pilgrims.length;
      else womenCount += sections[i].pilgrims.length;
    }
    for (var i = 0; i < prefilled.length; i++) {
      for (var p = 0; p < prefilled[i].pilgrims.length; p++) {
        if (prefilled[i].pilgrims[p].gender === 'Male') menCount++; else womenCount++;
      }
    }
    
    return {
      success: true, sections: sections, prefilled: prefilled,
      summary: { total: sharedPilgrims.length, prefilledCount: totalPrefilled, menCount: menCount, womenCount: womenCount }
    };
  } catch (err) {
    return { success: false, error: 'خطأ: ' + err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
// الدالة 3: حفظ التجميع
// ══════════════════════════════════════════════════════════════════

function saveGrouping(guideName, hotelName, groups) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.GUIDE_GROUPING);
    if (!sheet) return { success: false, error: 'شيت Guide Grouping غير موجود' };
    
    var groupedHotels = getGroupedHotels_(ss, guideName);
    if (groupedHotels[hotelName]) {
      return { success: false, error: 'تم حفظ التوزيع مسبقاً لهذا الفندق' };
    }
    
    var confirmationNo = generateConfirmationNo_(sheet);
    var rows = [];
    var timestamp = new Date();
    
    for (var g = 0; g < groups.length; g++) {
      var group = groups[g];
      var pilgrims = group.pilgrims || [];
      for (var p = 0; p < pilgrims.length; p++) {
        var pil = pilgrims[p];
        rows.push([
          guideName, hotelName, group.groupIndex, group.category,
          group.capacity || 4, pil.applicantId || '', pil.name || '',
          pil.gender || '', pil.passport || '', pil.bookingId || '',
          pil.isFamily ? 'TRUE' : 'FALSE', 'Grouped', timestamp, confirmationNo
        ]);
      }
    }
    
    if (rows.length === 0) return { success: false, error: 'لا توجد بيانات للحفظ' };
    
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 14).setValues(rows);
    return { success: true, saved: rows.length, confirmationNo: confirmationNo };
  } catch (err) {
    return { success: false, error: 'خطأ في الحفظ: ' + err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
// الدالة 4: جلب التجميع المحفوظ
// ══════════════════════════════════════════════════════════════════

function getExistingGrouping(guideName, hotelName) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.GUIDE_GROUPING);
    if (!sheet) return null;
    var lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.START_ROW) return null;
    var data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - 1, 14).getValues();
    var groups = {};
    var confirmationNo = '';
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (String(row[0]).trim() !== guideName || String(row[1]).trim() !== hotelName) continue;
      var gi = String(row[2]);
      if (!confirmationNo && row[13]) confirmationNo = String(row[13]);
      if (!groups[gi]) {
        groups[gi] = { groupIndex: parseInt(gi), category: String(row[3]), capacity: parseInt(row[4]), pilgrims: [] };
      }
      groups[gi].pilgrims.push({
        applicantId: String(row[5]), name: String(row[6]), gender: String(row[7]),
        passport: String(row[8]), bookingId: String(row[9]), isFamily: String(row[10]) === 'TRUE'
      });
    }
    var result = [];
    for (var g in groups) result.push(groups[g]);
    result.sort(function(a, b) { return a.groupIndex - b.groupIndex; });
    if (result.length === 0) return null;
    return { groups: result, confirmationNo: confirmationNo };
  } catch (err) {
    return null;
  }
}


// ══════════════════════════════════════════════════════════════════
// الدالة 5: توليد PDF التأكيد
// ══════════════════════════════════════════════════════════════════

function generateConfirmationPDF(guideName, hotelName) {
  try {
    var existing = getExistingGrouping(guideName, hotelName);
    if (!existing) return { success: false, error: 'لا يوجد تجميع محفوظ' };
    var groups = existing.groups;
    var confNo = existing.confirmationNo || '—';
    
    var html = '<html dir="rtl"><head><meta charset="UTF-8"><style>';
    html += 'body{font-family:Arial,sans-serif;padding:30px;direction:rtl;color:#333}';
    html += '.header{text-align:center;border-bottom:3px solid #1a5276;padding-bottom:15px;margin-bottom:20px}';
    html += '.header h1{color:#1a5276;font-size:20px;margin:0}.header h2{color:#2980b9;font-size:16px;margin:5px 0}';
    html += '.conf-box{background:#eaf2f8;border:2px solid #2980b9;border-radius:8px;padding:12px;text-align:center;margin:15px 0}';
    html += '.conf-box .num{font-size:22px;font-weight:bold;color:#1a5276}.conf-box .label{font-size:12px;color:#666}';
    html += '.info-row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #eee;font-size:13px}';
    html += '.group-title{background:#1a5276;color:#fff;padding:8px 12px;border-radius:5px;margin:15px 0 8px;font-size:14px}';
    html += 'table{width:100%;border-collapse:collapse;margin-bottom:15px;font-size:12px}';
    html += 'th{background:#2980b9;color:#fff;padding:8px;text-align:right}';
    html += 'td{padding:6px 8px;border-bottom:1px solid #ddd}tr:nth-child(even){background:#f8f9fa}';
    html += '.family-tag{background:#fff3cd;color:#856404;padding:1px 6px;border-radius:8px;font-size:10px}';
    html += '.footer{margin-top:30px;padding-top:15px;border-top:2px solid #1a5276;text-align:center;font-size:11px;color:#888}';
    html += '.warning{background:#fdecea;border:1px solid #e74c3c;border-radius:5px;padding:10px;text-align:center;font-size:12px;color:#c0392b;margin:15px 0;font-weight:bold}';
    html += '.system-note-pdf{background:#fef9e7;border:1px solid #f9e79f;border-radius:5px;padding:6px 10px;font-size:11px;color:#7d6608;margin:5px 0}';
    html += '</style></head><body>';
    
    html += '<div class="header"><h1>شركة إكرام الضيف للسياحة</h1>';
    html += '<h2>تأكيد تجميع الغرف — حج 1447هـ / 2026م</h2></div>';
    html += '<div class="conf-box"><div class="label">رقم التأكيد</div><div class="num">' + confNo + '</div></div>';
    html += '<div class="info-row"><span>المرشد:</span><strong>' + guideName + '</strong></div>';
    html += '<div class="info-row"><span>الفندق:</span><strong>' + hotelName + '</strong></div>';
    var totalPilgrims = 0;
    for (var g = 0; g < groups.length; g++) totalPilgrims += groups[g].pilgrims.length;
    html += '<div class="info-row"><span>عدد المجموعات:</span><strong>' + groups.length + '</strong></div>';
    html += '<div class="info-row"><span>إجمالي الحجاج:</span><strong>' + totalPilgrims + '</strong></div>';
    
    for (var g = 0; g < groups.length; g++) {
      var grp = groups[g];
      html += '<div class="group-title">' + grp.category + ' — مجموعة ' + grp.groupIndex + ' (' + grp.pilgrims.length + '/' + grp.capacity + ')</div>';
      html += '<table><tr><th>#</th><th>الاسم</th><th>الجنس</th><th>الجواز</th><th>ملاحظة</th></tr>';
      for (var p = 0; p < grp.pilgrims.length; p++) {
        var pil = grp.pilgrims[p];
        var genderAr = pil.gender === 'Male' ? 'ذكر' : 'أنثى';
        var note = pil.isFamily ? '<span class="family-tag">👨‍👩‍👧 عائلة</span>' : '';
        html += '<tr><td>' + (p+1) + '</td><td>' + pil.name + '</td><td>' + genderAr + '</td><td>' + pil.passport + '</td><td>' + note + '</td></tr>';
      }
      html += '</table>';
      if (grp.pilgrims.length < grp.capacity) {
        html += '<div class="system-note-pdf">ℹ️ سيتم إكمال هذه المجموعة تلقائياً (' + (grp.capacity - grp.pilgrims.length) + ' سرير متبقي)</div>';
      }
    }
    
    html += '<div class="warning">⚠️ هذا التسكين نهائي ولا يمكن تعديله إلا بطلب رسمي لإدارة العمليات</div>';
    var now = new Date();
    html += '<div class="footer">تاريخ التوليد: ' + now.getFullYear() + '-' + pad_(now.getMonth()+1) + '-' + pad_(now.getDate()) + ' ' + pad_(now.getHours()) + ':' + pad_(now.getMinutes());
    html += '<br>رقم التأكيد: ' + confNo + '<br>شركة إكرام الضيف للسياحة — جميع الحقوق محفوظة</div>';
    html += '</body></html>';
    
    var blob = HtmlService.createHtmlOutput(html).getBlob()
      .setName('تأكيد_تجميع_' + guideName + '_' + hotelName + '.pdf').getAs('application/pdf');
    var base64 = Utilities.base64Encode(blob.getBytes());
    var fileName = 'تأكيد_تجميع_' + guideName.replace(/\s+/g, '_') + '_' + confNo + '.pdf';
    return { success: true, pdfBase64: base64, fileName: fileName };
  } catch (err) {
    return { success: false, error: 'خطأ في توليد PDF: ' + err.message };
  }
}


// ══════════════════════════════════════════════════════════════════
// توليد روابط المرشدين — v2.2 (محدّثة)
// ──────────────────────────────────────────────────────────────────
// التعديل: إضافة عمود E (جواز المرشد) — يُملأ يدوياً
// ══════════════════════════════════════════════════════════════════

function generateGuideLinks() {
  var DEPLOY_URL = 'https://script.google.com/macros/s/AKfycbwirA68K9peWAsnAmcYZV3sDU1Bpy9TWQKt9lvtlF0kDGcvtxgqvAmJzhBdvD_eUf9Bdg/exec';
  
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var guideSheet = ss.getSheetByName(CONFIG.SHEETS.TOUR_GUIDE);
  var lastRow = guideSheet.getLastRow();
  var data = guideSheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  var guidesMap = {};
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    var registered = String(data[i][9]);
    var unique = String(data[i][10]);
    if (!name) continue;
    if (registered.indexOf('✅') === -1 || unique.indexOf('✅') === -1) continue;
    if (!guidesMap[name]) guidesMap[name] = { name: name, count: 0 };
    guidesMap[name].count++;
  }
  
  // ── جلب جوازات المرشدين الحالية (إن وُجدت) لعدم مسحها ──
  var existingPassports = {};
  var linkSheet = ss.getSheetByName(CONFIG.SHEETS.GUIDE_LINKS);
  if (linkSheet) {
    var existingLastRow = linkSheet.getLastRow();
    if (existingLastRow >= 2) {
      var existingData = linkSheet.getRange(2, 1, existingLastRow - 1, 5).getValues();
      for (var i = 0; i < existingData.length; i++) {
        var eName = String(existingData[i][1]).trim();
        var ePassport = String(existingData[i][4]).trim();
        if (eName && ePassport) existingPassports[eName] = ePassport;
      }
    }
    linkSheet.clear();
  } else {
    linkSheet = ss.insertSheet(CONFIG.SHEETS.GUIDE_LINKS);
  }
  
  // ── رؤوس الأعمدة ──
  var headers = [['#', 'اسم المرشد', 'عدد الحجاج', 'الرابط', 'جواز المرشد']];
  linkSheet.getRange(1, 1, 1, 5).setValues(headers);
  linkSheet.getRange(1, 1, 1, 5)
    .setBackground('#1a5276').setFontColor('#ffffff')
    .setFontWeight('bold').setFontFamily('Tajawal');
  
  var guides = [];
  for (var name in guidesMap) guides.push(guidesMap[name]);
  guides.sort(function(a, b) { return a.name.localeCompare(b.name); });
  
  var rows = [];
  for (var i = 0; i < guides.length; i++) {
    var g = guides[i];
    var url = DEPLOY_URL + '?g=' + encodeURIComponent(g.name);
    var existingPassport = existingPassports[g.name] || '';
    rows.push([i + 1, g.name, g.count, url, existingPassport]);
  }
  
  if (rows.length > 0) {
    linkSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
  
  linkSheet.setColumnWidth(1, 40);
  linkSheet.setColumnWidth(2, 200);
  linkSheet.setColumnWidth(3, 100);
  linkSheet.setColumnWidth(4, 600);
  linkSheet.setColumnWidth(5, 150);
  linkSheet.setRightToLeft(true);
  linkSheet.setFrozenRows(1);
  
  for (var i = 0; i < rows.length; i++) {
    var cell = linkSheet.getRange(i + 2, 4);
    cell.setFormula('=HYPERLINK("' + rows[i][3] + '", "افتح الرابط")');
    cell.setFontColor('#2980b9');
  }
  
  // تلوين عمود جواز المرشد الفارغ
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][4]) {
      linkSheet.getRange(i + 2, 5).setBackground('#fdecea');
    }
  }
  
  Logger.log('تم توليد ' + rows.length + ' رابط. املأ عمود "جواز المرشد" (E) يدوياً.');
}


// ══════════════════════════════════════════════════════════════════
// دوال مساعدة
// ══════════════════════════════════════════════════════════════════

function generateConfirmationNo_(sheet) {
  var lastRow = sheet.getLastRow();
  var maxNum = 0;
  if (lastRow >= CONFIG.START_ROW) {
    var colN = sheet.getRange(CONFIG.START_ROW, 14, lastRow - 1, 1).getValues();
    for (var i = 0; i < colN.length; i++) {
      var match = String(colN[i][0]).match(/GRP-2026-(\d+)/);
      if (match) { var num = parseInt(match[1]); if (num > maxNum) maxNum = num; }
    }
  }
  return 'GRP-2026-' + padNum_(maxNum + 1, 4);
}

function padNum_(num, len) { var s = String(num); while (s.length < len) s = '0' + s; return s; }
function pad_(n) { return n < 10 ? '0' + n : '' + n; }

function getGuidePassports_(ss, guideName) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.TOUR_GUIDE);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return [];
  var data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - 1, 11).getValues();
  var passports = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (String(row[CONFIG.GUIDE.NAME]).trim() === guideName &&
        String(row[CONFIG.GUIDE.REGISTERED]).indexOf('✅') !== -1 &&
        String(row[CONFIG.GUIDE.UNIQUE]).indexOf('✅') !== -1) {
      var passport = String(row[CONFIG.GUIDE.PASSPORT]).trim();
      if (passport) passports.push(passport);
    }
  }
  return passports;
}

function getPilgrimsByPassports_(ss, passports) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PILGRIMS);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return {};
  var passportSet = {};
  for (var i = 0; i < passports.length; i++) passportSet[passports[i]] = true;
  var data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - 1, 12).getValues();
  var result = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var passport = String(row[CONFIG.PILGRIM.PASSPORT]).trim();
    if (passportSet[passport]) {
      result[passport] = {
        bookingId: String(row[CONFIG.PILGRIM.BOOKING_ID]).trim(),
        applicantId: String(row[CONFIG.PILGRIM.APPLICANT_ID]).trim(),
        groupNumber: String(row[CONFIG.PILGRIM.GROUP_NUMBER]).trim(),
        name: String(row[CONFIG.PILGRIM.NAME]).trim(),
        passport: passport,
        gender: String(row[CONFIG.PILGRIM.GENDER]).trim(),
        isFamily: false, familyBookingId: ''
      };
    }
  }
  return result;
}

function getRoomTypeByGroups_(ss, groupNumbers) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ROOM_TYPE);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return {};
  var gnSet = {};
  for (var i = 0; i < groupNumbers.length; i++) gnSet[String(groupNumbers[i])] = true;
  var data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - 1, 30).getValues();
  var result = {};
  for (var i = 0; i < data.length; i++) {
    var gn = String(data[i][CONFIG.RT.GROUP_NUMBER]).trim();
    if (gnSet[gn]) result[gn] = data[i];
  }
  return result;
}

function getGroupedHotels_(ss, guideName) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.GUIDE_GROUPING);
  var result = {};
  if (!sheet) return result;
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.START_ROW) return result;
  var data = sheet.getRange(CONFIG.START_ROW, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === guideName) result[String(data[i][1]).trim()] = true;
  }
  return result;
}

function getRoomCapacity_(roomType) {
  var t = String(roomType).trim().toLowerCase();
  if (t.indexOf('double') !== -1 || t.indexOf('dbl') !== -1) return 2;
  if (t.indexOf('triple') !== -1 || t.indexOf('trp') !== -1) return 3;
  return 4;
}

function getRoomTypeLabel_(capacity) {
  if (capacity === 2) return 'ثنائي';
  if (capacity === 3) return 'ثلاثي';
  return 'رباعي';
}

function normalizeHotel_(name) {
  return String(name || '').trim().toLowerCase()
    .replace(/[-–—]/g, ' ').replace(/\s+/g, ' ')
    .replace(/\s*(madinah|makkah|mecca|medina)\s*$/i, '').trim();
}

function debugGuideHotels() {
  var guideName = 'Abdul Rahman Azmi';
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // Step 1: Passports
  var passports = getGuidePassports_(ss, guideName);
  Logger.log('Step 1 — Passports found: ' + passports.length);
  Logger.log('Sample: ' + passports.slice(0, 3).join(', '));
  
  // Step 2: Pilgrims
  var pilgrimData = getPilgrimsByPassports_(ss, passports);
  var pilgrimCount = Object.keys(pilgrimData).length;
  Logger.log('Step 2 — Pilgrims matched: ' + pilgrimCount);
  
  if (pilgrimCount === 0) {
    // Check sheet structure
    var pSheet = ss.getSheetByName(CONFIG.SHEETS.PILGRIMS);
    Logger.log('Sheet name: "' + CONFIG.SHEETS.PILGRIMS + '" found: ' + !!pSheet);
    if (pSheet) {
      Logger.log('Rows: ' + pSheet.getLastRow());
      var headers = pSheet.getRange(1, 1, 1, 12).getValues()[0];
      Logger.log('Headers: ' + JSON.stringify(headers));
      var row2 = pSheet.getRange(2, 1, 1, 12).getValues()[0];
      Logger.log('Row 2 col I (passport): ' + row2[8]);
    }
    return;
  }
  
  // Step 3: Group numbers
  var gns = {};
  for (var p in pilgrimData) gns[String(pilgrimData[p].groupNumber)] = true;
  Logger.log('Step 3 — Group numbers: ' + Object.keys(gns).length);
  
  // Step 4: Room Type
  var rtData = getRoomTypeByGroups_(ss, Object.keys(gns));
  Logger.log('Step 4 — RT matches: ' + Object.keys(rtData).length);
  
  // Step 5: Hotels with shared > 0
  var hotels = {};
  for (var passport in pilgrimData) {
    var pil = pilgrimData[passport];
    var rtRow = rtData[String(pil.groupNumber)];
    if (!rtRow) continue;
    for (var i = 0; i < HOTEL_POSITIONS.length; i++) {
      var pos = HOTEL_POSITIONS[i];
      var hn = String(rtRow[CONFIG.RT[pos.hotel]] || '').trim();
      var shared = parseInt(rtRow[CONFIG.RT[pos.shared]]) || 0;
      Logger.log('  ' + passport + ' → hotel=' + hn + ' shared=' + shared);
      if (hn && shared > 0) hotels[hn] = (hotels[hn] || 0) + 1;
    }
  }
  Logger.log('Step 5 — Hotels: ' + JSON.stringify(hotels));
}
function checkSheetName() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n.indexOf('رحلة') !== -1 || n.indexOf('حاج') !== -1) {
      Logger.log('Found: "' + n + '" len=' + n.length);
    }
  }
}