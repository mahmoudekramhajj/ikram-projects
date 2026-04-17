/**
 * EmailReader.js — قراءة إيميلات تغيير الرحلات من Gmail
 * النسخة 2: مبسّط + مقارنة + فلترة تذاكر نسك فقط
 */


function scanEmails() {
  var startTime = new Date();
  var results = { processed: 0, skipped: 0, errors: 0 };

  var processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  var query = 'label:' + CONFIG.GMAIL_LABEL + ' -label:' + CONFIG.PROCESSED_LABEL;
  var threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN);

  Logger.log('وُجد ' + threads.length + ' إيميل جديد');
  if (threads.length === 0) return results;

  var changesFolder = getOrCreateChangesFolder_();
  var changesSheet = getOrCreateChangesSheet_();
  var comparisonSheet = getOrCreateComparisonSheet_();
  var allData = loadAllData_();

  // تحميل الصفوف الموجودة لمنع التكرار (PNR + اسم)
  var existingKeys = loadExistingKeys_(comparisonSheet);

  for (var t = 0; t < threads.length; t++) {
    if (new Date() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('⏰ توقف — انتهت المهلة بعد ' + t + ' إيميل');
      break;
    }

    var thread = threads[t];
    var messages = thread.getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var subject = message.getSubject();
      var body = message.getPlainBody();
      var date = message.getDate();
      var incidentNum = extractIncidentNumber_(subject);

      Logger.log('\n📧 #' + (t + 1) + ': ' + subject);

      try {
        var pnr = extractPNRFromBody_(body);

        // فلترة PDFs — فقط تذاكر نسك
        var attachments = message.getAttachments();
        var nusukPDFs = attachments.filter(function(att) {
          var name = att.getName().toLowerCase();
          return name.endsWith('.pdf') && (name.indexOf('flight_eticket') !== -1 || name.indexOf('eticket') !== -1);
        });

        if (nusukPDFs.length === 0) {
          Logger.log('  ⏭️ بدون تذكرة نسك — تخطي');
          results.skipped++;
          continue;
        }

        for (var p = 0; p < nusukPDFs.length; p++) {
          var pdf = nusukPDFs[p];
          Logger.log('  📄 ' + pdf.getName());

          // حفظ في Drive
          var driveFile = changesFolder.createFile(pdf);
          var shareLink = createShareLink_(driveFile);

          // تحليل PDF
          var text = extractTextFromPDF_(driveFile.getId());
          var ticketData = parseNusukTicket_(text);

          if (!ticketData) {
            Logger.log('  ❌ لم يتم التعرف على التذكرة');
            results.errors++;
            continue;
          }

          if (!pnr && ticketData.pnr) pnr = ticketData.pnr;
          var activePnr = pnr || ticketData.pnr;

          // استخراج كل المسافرين من PDF
          var passengers = ticketData.passengers || [];
          if (passengers.length === 0 && (ticketData.firstName || ticketData.lastName)) {
            passengers = [{ firstName: ticketData.firstName, lastName: ticketData.lastName, fullName: ticketData.passengerName }];
          }

          if (passengers.length === 0) {
            Logger.log('  ⚠️ لا يوجد أسماء في PDF');
          }

          Logger.log('  👥 ' + passengers.length + ' مسافر في PDF');

          // كتابة صف لكل مسافر
          for (var px = 0; px < passengers.length; px++) {
            var pax = passengers[px];

            // فحص التكرار — نفس PNR + اسم = تخطي
            var dupKey = (activePnr + '|' + (pax.firstName || '') + '|' + (pax.lastName || '')).toUpperCase();
            if (existingKeys[dupKey]) {
              Logger.log('  ⏭️ مكرر: ' + pax.firstName + ' ' + pax.lastName + ' | PNR: ' + activePnr);
              continue;
            }
            existingKeys[dupKey] = true;

            Logger.log('  ✈️ ' + pax.firstName + ' ' + pax.lastName + ' | PNR: ' + activePnr);

            // === المحاولة 1: مطابقة الاسم كما هو (بدون فلاتر) ===
            var match = findPilgrim_(pax.firstName, pax.lastName, activePnr, allData, ticketData.bookingId);

            // === المحاولة 2: لو فشلت — ننظف الاسم (نزيل أكواد المطارات) ونحاول مرة ثانية ===
            if (!match) {
              var cleanedFull = deepClean_(pax.fullName || (pax.firstName + ' ' + pax.lastName));
              if (cleanedFull && cleanedFull !== (pax.firstName + ' ' + pax.lastName).trim()) {
                var cleanParts = cleanedFull.split(/\s+/);
                var cleanFirst = cleanParts[0] || '';
                var cleanLast = cleanParts.length >= 2 ? cleanParts.slice(1).join(' ') : '';
                Logger.log('  🔄 محاولة ثانية بعد التنظيف: ' + cleanFirst + ' ' + cleanLast);
                match = findPilgrim_(cleanFirst, cleanLast, activePnr, allData, ticketData.bookingId);
                if (match) {
                  // تحديث الاسم بالنسخة النظيفة
                  pax.firstName = cleanFirst;
                  pax.lastName = cleanLast;
                  pax.fullName = cleanedFull;
                  Logger.log('  ✅ تطابق بعد التنظيف!');
                }
              }
            }

            var status = match ? 'تم المطابقة' : 'لم يُطابَق';
            var notes = '';

            if (match) {
              Logger.log('  ✅ ' + match.sysFirstName + ' ' + match.sysLastName + ' [' + match.source + ']');
            } else {
              notes = 'PNR: ' + activePnr;
              Logger.log('  ⚠️ لم يُطابَق — PNR: ' + activePnr);
            }

            var changeNum = 'CHG-' + String(changesSheet.getLastRow()).padStart(4, '0');

            var rowData = {
              changeNum: changeNum,
              pnr: activePnr,
              bookingId: ticketData.bookingId,
              incidentNum: incidentNum,
              pdfFirstName: pax.firstName || '',
              pdfLastName: pax.lastName || '',
              sysFirstName: match ? match.sysFirstName : '',
              sysLastName: match ? match.sysLastName : '',
              serialNum: match ? match.serial : '',
              passport: match ? match.passport : '',
              pkgNum: match ? match.pkgNum : '',
              pkgName: match ? match.pkgName : '',
              flightType: match ? match.flightType : '',
              bookingType: match ? (String(match.source || '').indexOf('B2C') !== -1 ? 'B2C' : 'B2B') : '',
              outboundLegs: ticketData.outboundLegs,
              returnLegs: ticketData.returnLegs,
              curOutFlight1: match ? match.curOutFlight1 : '',
              curOutDate1: match ? match.curOutDate1 : '',
              curOutFlight2: match ? match.curOutFlight2 : '',
              curOutDate2: match ? match.curOutDate2 : '',
              curRetFlight1: match ? match.curRetFlight1 : '',
              curRetDate1: match ? match.curRetDate1 : '',
              curRetFlight2: match ? match.curRetFlight2 : '',
              curRetDate2: match ? match.curRetDate2 : '',
              pdfLink: shareLink,
              emailDate: date,
              status: status,
              notes: notes
            };

            writeChangeRow_(changesSheet, rowData);

            var compStatus = 'لا يوجد اسم';
            if (pax.firstName || pax.lastName) {
              compStatus = match ? 'متطابق' : 'غير متطابق';
            }
            rowData.comparisonStatus = compStatus;
            writeComparisonRow_(comparisonSheet, rowData);
          }

          results.processed++;
        }

      } catch (e) {
        Logger.log('  ❌ خطأ: ' + e.message);
        results.errors++;
      }
    }

    thread.addLabel(processedLabel);
  }

  Logger.log('\n=== ✅ ' + results.processed + ' | ⏭️ ' + results.skipped + ' | ❌ ' + results.errors + ' ===');
  return results;
}


// === دوال مساعدة ===

function extractPNRFromBody_(body) {
  var patterns = [
    /PNR#([A-Z0-9]{5,8})/i,
    /reference\s+(?:PNR#)?([A-Z0-9]{5,8})/i,
    /booking\s*#([A-Z0-9]{5,8})/i,
    /your\s+reference\s+([A-Z0-9]{5,8})/i,
    /Regarding\s+your\s+reference\s+([A-Z0-9]{5,8})/i,
    /Reservation\s+([A-Z0-9]{5,8})/i
  ];
  for (var i = 0; i < patterns.length; i++) {
    var m = body.match(patterns[i]);
    if (m && m[1].length >= 5 && m[1].length <= 8) return m[1].trim();
  }
  return '';
}

function extractIncidentNumber_(subject) {
  var m = subject.match(/Incident#?\s*(\d+)/i);
  return m ? m[1] : '';
}

function createShareLink_(file) {
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateLabel_(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log('تم إنشاء label: ' + labelName);
  }
  return label;
}

/**
 * تحميل المفاتيح الموجودة (PNR + اسم) من شيت المقارنة لمنع التكرار
 */
function loadExistingKeys_(sheet) {
  var keys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;

  var data = sheet.getRange(2, 2, lastRow - 1, 4).getValues(); // B=PNR, D=firstName, E=lastName
  for (var i = 0; i < data.length; i++) {
    var key = (String(data[i][0]) + '|' + String(data[i][2]) + '|' + String(data[i][3])).toUpperCase();
    keys[key] = true;
  }
  return keys;
}


/**
 * حذف الصفوف المكررة من شيتي المقارنة والتغييرات
 */
function removeDuplicates() {
  var removed = { comparison: 0, changes: 0 };

  // === شيت المقارنة ===
  var compSheet = getOrCreateComparisonSheet_();
  removed.comparison = deduplicateSheet_(compSheet, 1, 3, 4); // B=PNR(col2→idx1), D=first(col4→idx3), E=last(col5→idx4)

  // === شيت التغييرات ===
  var chgSheet = getOrCreateChangesSheet_();
  removed.changes = deduplicateSheet_(chgSheet, 1, 4, 5); // B=PNR(col2→idx1), E=first(col5→idx4), F=last(col6→idx5)

  // إعادة ترقيم عمود # في المقارنة
  var compLastRow = compSheet.getLastRow();
  if (compLastRow >= 2) {
    var nums = [];
    for (var i = 1; i <= compLastRow - 1; i++) nums.push([i]);
    compSheet.getRange(2, 1, nums.length, 1).setValues(nums);
  }

  return removed;
}


/**
 * حذف المكررات من شيت واحد — يبقي أول ظهور ويحذف الباقي
 */
function deduplicateSheet_(sheet, pnrIdx, fnIdx, lnIdx) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0; // أقل من صفين → لا مكررات

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var seen = {};
  var rowsToDelete = [];

  for (var i = 0; i < data.length; i++) {
    var key = (String(data[i][pnrIdx]) + '|' + String(data[i][fnIdx]) + '|' + String(data[i][lnIdx])).toUpperCase();
    if (seen[key]) {
      rowsToDelete.push(i + 2); // +2 لأن البيانات تبدأ من صف 2
    } else {
      seen[key] = true;
    }
  }

  // حذف من الأسفل للأعلى حتى لا تتغير أرقام الصفوف
  for (var d = rowsToDelete.length - 1; d >= 0; d--) {
    sheet.deleteRow(rowsToDelete[d]);
  }

  return rowsToDelete.length;
}


function getOrCreateChangesFolder_() {
  var parentFolder = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  var folders = parentFolder.getFoldersByName(CONFIG.CHANGES_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  var newFolder = parentFolder.createFolder(CONFIG.CHANGES_FOLDER_NAME);
  Logger.log('تم إنشاء مجلد: ' + CONFIG.CHANGES_FOLDER_NAME);
  return newFolder;
}

function createAutoTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'scanEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scanEmails').timeBased().everyMinutes(5).create();
  Logger.log('✅ Trigger created');
}

function stopAutoTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'scanEmails') ScriptApp.deleteTrigger(t);
  });
  Logger.log('🛑 Triggers removed');
}


/**
 * إزالة label TKT-Processed من جميع الإيميلات لإعادة المعالجة
 */
function resetProcessedLabels() {
  var label = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  if (!label) {
    Logger.log('⚠️ Label not found: ' + CONFIG.PROCESSED_LABEL);
    return { removed: 0 };
  }

  var threads = label.getThreads();
  var count = threads.length;
  for (var i = 0; i < threads.length; i++) {
    threads[i].removeLabel(label);
  }
  Logger.log('✅ تم إزالة TKT-Processed من ' + count + ' إيميل');
  return { removed: count };
}
