/**
 * EmailReader.js — قراءة إيميلات تغيير الرحلات من Gmail
 * النسخة 2: مبسّط + مقارنة + فلترة تذاكر نسك فقط
 */


function scanEmails() {
  var startTime = new Date();
  var results = { processed: 0, skipped: 0, errors: 0, textProcessed: 0 };

  var processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  var skippedLabel = getOrCreateLabel_(CONFIG.SKIPPED_LABEL);
  var query = 'label:' + CONFIG.GMAIL_LABEL +
              ' -label:' + CONFIG.PROCESSED_LABEL +
              ' -label:' + CONFIG.SKIPPED_LABEL;
  var threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN);

  Logger.log('وُجد ' + threads.length + ' إيميل جديد');
  if (threads.length === 0) return results;

  var changesFolder = getOrCreateChangesFolder_();
  var changesSheet = getOrCreateChangesSheet_();
  var comparisonSheet = getOrCreateComparisonSheet_();
  var allData = loadAllData_();

  // تحميل الصفوف الموجودة من كل شيت بمفرده — يسمح باسترداد الـ 185 الضائعة
  var existingComparisonKeys = loadExistingKeys_(comparisonSheet, 'COMPARISON');
  var existingChangesKeys = loadExistingKeys_(changesSheet, 'CHANGES');

  for (var t = 0; t < threads.length; t++) {
    if (new Date() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('⏰ توقف — انتهت المهلة بعد ' + t + ' إيميل');
      break;
    }

    var thread = threads[t];
    var threadProcessed = false;  // ← جديد: هل كُتب صف واحد على الأقل لهذا الـ thread؟
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

        // === نوع B: إيميل نصّي بدون PDF — محاولة استخراج الرحلات من النص ===
        if (nusukPDFs.length === 0) {
          var textLegs = extractFlightLegs_(body);
          if (textLegs && textLegs.length > 0) {
            Logger.log('  📝 إيميل نصّي — ' + textLegs.length + ' رحلة من النص');
            var written = processTextOnlyEmail_(
              message, body, subject, date, incidentNum, pnr, textLegs,
              allData, changesSheet, comparisonSheet,
              existingChangesKeys, existingComparisonKeys
            );
            if (written > 0) {
              results.textProcessed += written;
              threadProcessed = true;
            }
          } else {
            Logger.log('  ⏭️ بدون تذكرة نسك ولا نص رحلات — تخطي');
            results.skipped++;
          }
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

            // فحص التكرار لكل شيت بمفرده — يسمح باسترداد الصف إذا كان في شيت واحد فقط
            var dupKey = (activePnr + '|' + (pax.firstName || '') + '|' + (pax.lastName || '')).toUpperCase();
            var inChanges = existingChangesKeys[dupKey];
            var inComparison = existingComparisonKeys[dupKey];

            if (inChanges && inComparison) {
              Logger.log('  ⏭️ مكرر في الشيتين: ' + pax.firstName + ' ' + pax.lastName + ' | PNR: ' + activePnr);
              continue;
            }

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

            var compStatus = 'لا يوجد اسم';
            if (pax.firstName || pax.lastName) {
              compStatus = match ? 'متطابق' : 'غير متطابق';
            }
            rowData.comparisonStatus = compStatus;

            // كتابة انتقائية — فقط في الشيت الناقص
            if (!inChanges) {
              writeChangeRow_(changesSheet, rowData);
              existingChangesKeys[dupKey] = true;
              threadProcessed = true;  // صف كُتب فعلاً
              if (inComparison) {
                Logger.log('  ♻️ استرداد: ' + pax.firstName + ' ' + pax.lastName + ' (كان ناقصاً من شيت التغييرات)');
              }
            }

            if (!inComparison) {
              writeComparisonRow_(comparisonSheet, rowData);
              existingComparisonKeys[dupKey] = true;
            }
          }

          results.processed++;
        }

      } catch (e) {
        Logger.log('  ❌ خطأ: ' + e.message);
        results.errors++;
      }
    }

    // ← Fix الأهم: الـ label يُضاف فقط إذا كُتب صف على الأقل
    if (threadProcessed) {
      thread.addLabel(processedLabel);
    } else {
      thread.addLabel(skippedLabel);
    }
  }

  Logger.log('\n=== ✅ PDF: ' + results.processed + ' | 📝 نص: ' + results.textProcessed + ' | ⏭️ متجاهَل: ' + results.skipped + ' | ❌ ' + results.errors + ' ===');
  return results;
}


/**
 * معالجة إيميل نصّي (بلا PDF) — نوع B
 * يستخرج الرحلات من extractFlightLegs_ + يطابق الحاج بالـ PNR/الإيميل المستلم
 * يرجع عدد الصفوف المكتوبة
 */
function processTextOnlyEmail_(message, body, subject, date, incidentNum, pnr, textLegs, allData, changesSheet, comparisonSheet, existingChangesKeys, existingComparisonKeys) {
  // استخراج إيميل المستلم الأصلي من النص (الحاج الحقيقي — ليس Mahmoud)
  // نمط: "إلى: X@Y.Z" أو "To: X@Y.Z"
  var recipientEmail = '';
  var emailMatch = body.match(/(?:إلى|To)\s*:?\s*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  if (emailMatch) recipientEmail = emailMatch[1].toLowerCase();

  // إذا لم يوجد PNR في body — جرّب استخراجه من نمط "Booking XXXXXX"
  if (!pnr) {
    var bookMatch = body.match(/Booking\s+([A-Z0-9]{5,8})/i);
    if (bookMatch) pnr = bookMatch[1].toUpperCase();
  }

  // مطابقة الحاج: بـ PNR أولاً (بدون اسم لأن النص لا يحوي اسماً عادة) → ثم بالإيميل
  var match = null;
  if (pnr) {
    // محاولة: PNR في contract name (PD) مع أي اسم
    for (var i = 0; i < allData.pd.length && !match; i++) {
      var contract = String(allData.pd[i][CONFIG.PD.CONTRACT_NAME] || '').toUpperCase();
      if (contract.indexOf(pnr) !== -1) {
        match = buildPDResult_(allData.pd[i]);
        match.source = 'PD (text PNR)';
        // رحلات حالية
        var cur = findCurrentFlights_(pnr, allData.flights);
        if (cur) {
          match.curOutFlight1 = cur.outFlight1; match.curOutDate1 = cur.outDate1;
          match.curOutFlight2 = cur.outFlight2; match.curOutDate2 = cur.outDate2;
          match.curRetFlight1 = cur.retFlight1; match.curRetDate1 = cur.retDate1;
          match.curRetFlight2 = cur.retFlight2; match.curRetDate2 = cur.retDate2;
        }
      }
    }
  }
  if (!match && recipientEmail) {
    match = findByEmail_(recipientEmail, pnr, allData);
  }

  // تصنيف الرحلات إلى ذهاب/عودة
  var dataObj = { outboundLegs: [], returnLegs: [] };
  classifyLegs_(dataObj, textLegs);

  // Dedup key: PNR + incidentNum (لإيميلات متكررة لنفس الحادث — in-memory ضمن نفس scan run)
  var dupKey = ('TEXT|' + (pnr || '') + '|' + (incidentNum || '') + '|' + (recipientEmail || '')).toUpperCase();
  if (existingChangesKeys[dupKey]) {
    Logger.log('  ⏭️ مكرر نصّي: ' + incidentNum);
    return 0;
  }
  existingChangesKeys[dupKey] = true;

  var status = match ? 'تم المطابقة' : 'لم يُطابَق';
  var notes = 'إيميل نصّي' + (recipientEmail ? ' | ' + recipientEmail : '');
  Logger.log('  ' + (match ? '✅' : '⚠️') + ' نصّي — PNR: ' + pnr + ' | email: ' + recipientEmail);

  // رابط الإيميل بدلاً من PDF
  var threadId = message.getThread().getId();
  var emailLink = 'https://mail.google.com/mail/u/0/#inbox/' + threadId;

  var changeNum = 'CHG-' + String(changesSheet.getLastRow()).padStart(4, '0');

  var rowData = {
    changeNum: changeNum,
    pnr: pnr || '',
    bookingId: '',
    incidentNum: incidentNum,
    pdfFirstName: '',  // النص لا يحوي اسم عادة
    pdfLastName: '',
    sysFirstName: match ? match.sysFirstName : '',
    sysLastName: match ? match.sysLastName : '',
    serialNum: match ? match.serial : '',
    passport: match ? match.passport : '',
    pkgNum: match ? match.pkgNum : '',
    pkgName: match ? match.pkgName : '',
    flightType: match ? match.flightType : '',
    bookingType: match ? (String(match.source || '').indexOf('B2C') !== -1 ? 'B2C' : 'B2B') : '',
    outboundLegs: dataObj.outboundLegs,
    returnLegs: dataObj.returnLegs,
    curOutFlight1: match ? match.curOutFlight1 : '',
    curOutDate1: match ? match.curOutDate1 : '',
    curOutFlight2: match ? match.curOutFlight2 : '',
    curOutDate2: match ? match.curOutDate2 : '',
    curRetFlight1: match ? match.curRetFlight1 : '',
    curRetDate1: match ? match.curRetDate1 : '',
    curRetFlight2: match ? match.curRetFlight2 : '',
    curRetDate2: match ? match.curRetDate2 : '',
    pdfLink: emailLink,  // رابط الإيميل
    emailDate: date,
    status: status,
    notes: notes
  };

  writeChangeRow_(changesSheet, rowData);

  rowData.comparisonStatus = match ? 'متطابق' : 'غير متطابق';
  writeComparisonRow_(comparisonSheet, rowData);

  return 1;
}


/**
 * استرداد الإيميلات الضائعة — إزالة TKT-Processed من الـ threads التي لم تُعالَج فعلاً
 * ثم شغّل scanEmails() بعد هذه الدالة
 */
function recoverLostEmails(dryRun) {
  var processedLabel = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  var skippedLabel = GmailApp.getUserLabelByName(CONFIG.SKIPPED_LABEL);
  if (!processedLabel && !skippedLabel) {
    Logger.log('❌ لا Labels موجودة');
    return { error: 'no labels' };
  }

  // الـ threads المُعلَّمة TKT-Processed أو TKT-Skipped بدون مرفقات = التي تُجوهلت خطأً
  var query = '(label:' + CONFIG.PROCESSED_LABEL + ' OR label:' + CONFIG.SKIPPED_LABEL + ') -has:attachment';
  var threads = GmailApp.search(query, 0, 100);

  Logger.log('🔍 وُجد ' + threads.length + ' thread بـ TKT-Processed وبلا مرفق');

  if (dryRun) {
    var samples = [];
    for (var i = 0; i < Math.min(threads.length, 10); i++) {
      samples.push(threads[i].getFirstMessageSubject());
    }
    return { dryRun: true, count: threads.length, samples: samples };
  }

  for (var i = 0; i < threads.length; i++) {
    if (processedLabel) threads[i].removeLabel(processedLabel);
    if (skippedLabel) threads[i].removeLabel(skippedLabel);
  }

  Logger.log('✅ أُزيلت الـ labels من ' + threads.length + ' thread — شغّل scanEmails() الآن');
  return { removed: threads.length };
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
  // قائمة سوداء — كلمات إنجليزية تُطابق نمط PNR عرضاً (false positives)
  var BLACKLIST = {
    'NUMBER':1, 'NUMBERS':1, 'BOOKING':1, 'FLIGHT':1, 'FLIGHTS':1,
    'TICKET':1, 'TICKETS':1, 'CHANGE':1, 'CHANGES':1, 'SCHEDULE':1,
    'PLEASE':1, 'REGARDS':1, 'THANKS':1, 'DETAILS':1, 'INVOICE':1,
    'DEPART':1, 'ARRIVE':1, 'AIRPORT':1, 'CANCEL':1, 'UPDATE':1,
    'UPDATED':1, 'HISTORY':1, 'ALERT':1, 'NOTICE':1, 'ABOUT':1,
    'WITHIN':1, 'BECAUSE':1, 'SUBJECT':1, 'MESSAGE':1, 'CONFIRM':1
  };
  for (var i = 0; i < patterns.length; i++) {
    var m = body.match(patterns[i]);
    if (m && m[1].length >= 5 && m[1].length <= 8) {
      var code = m[1].trim().toUpperCase();
      if (BLACKLIST[code]) continue;  // تخطي الكلمات الإنجليزية الشائعة
      return code;
    }
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
function loadExistingKeys_(sheet, sheetType) {
  var keys = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;

  // حسب نوع الشيت، الأعمدة في موقع مختلف:
  // COMPARISON: B=PNR(idx1), D=firstName(idx3), E=lastName(idx4)
  // CHANGES:    B=PNR(idx1), E=firstName(idx4), F=lastName(idx5)
  var pnrIdx, fnIdx, lnIdx;
  if (sheetType === 'CHANGES') {
    pnrIdx = 1; fnIdx = 4; lnIdx = 5;
  } else { // COMPARISON (default — للتوافق مع الاستدعاء القديم)
    pnrIdx = 1; fnIdx = 3; lnIdx = 4;
  }

  var maxCols = Math.max(pnrIdx, fnIdx, lnIdx) + 1;
  var data = sheet.getRange(2, 1, lastRow - 1, maxCols).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = (String(data[i][pnrIdx]) + '|' + String(data[i][fnIdx]) + '|' + String(data[i][lnIdx])).toUpperCase();
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


/**
 * إزالة كل labels المعالجة (Processed + Skipped) — بداية صفرية كاملة
 */
function resetAllLabels() {
  var removed = { processed: 0, skipped: 0 };
  var proc = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  var skip = GmailApp.getUserLabelByName(CONFIG.SKIPPED_LABEL);
  if (proc) {
    var pt = proc.getThreads();
    for (var i = 0; i < pt.length; i++) pt[i].removeLabel(proc);
    removed.processed = pt.length;
  }
  if (skip) {
    var st = skip.getThreads();
    for (var j = 0; j < st.length; j++) st[j].removeLabel(skip);
    removed.skipped = st.length;
  }
  Logger.log('✅ أُزيلت labels: Processed=' + removed.processed + ', Skipped=' + removed.skipped);
  return removed;
}


// =====================================================
// === V3: scanEmailsV3 — الشيت الموحّد + تحقق منطقي ===
// =====================================================

/**
 * النسخة V3 — يكتب في شيت واحد فقط (Unified) مع dedup PNR+Incident+Name + validation
 */
function scanEmailsV3() {
  var startTime = new Date();
  var results = { processed: 0, skipped: 0, errors: 0, textProcessed: 0, warnings: 0 };

  var processedLabel = getOrCreateLabel_(CONFIG.PROCESSED_LABEL);
  var skippedLabel = getOrCreateLabel_(CONFIG.SKIPPED_LABEL);
  var query = 'label:' + CONFIG.GMAIL_LABEL +
              ' -label:' + CONFIG.PROCESSED_LABEL +
              ' -label:' + CONFIG.SKIPPED_LABEL;
  var threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN);

  Logger.log('V3: وُجد ' + threads.length + ' إيميل جديد');
  if (threads.length === 0) return results;

  var changesFolder = getOrCreateChangesFolder_();
  var unifiedSheet = getOrCreateUnifiedSheet_();
  var allData = loadAllData_();

  // مفاتيح موجودة — PNR + Incident + firstName + lastName
  var existingKeys = loadUnifiedKeys_(unifiedSheet);

  for (var t = 0; t < threads.length; t++) {
    if (new Date() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('⏰ توقف — انتهت المهلة بعد ' + t + ' إيميل');
      break;
    }

    var thread = threads[t];
    var threadProcessed = false;
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

        // فلترة PDFs نسك
        var attachments = message.getAttachments();
        var nusukPDFs = attachments.filter(function(att) {
          var name = att.getName().toLowerCase();
          return name.endsWith('.pdf') && (name.indexOf('flight_eticket') !== -1 || name.indexOf('eticket') !== -1);
        });

        // === نوع B: إيميل نصّي ===
        if (nusukPDFs.length === 0) {
          // V5 — جرّب Claude على النص أولاً
          var textData = parseTextWithClaude_(body);

          if (textData && !textData.skipped && textData.outboundLegs) {
            // Claude نجح في استخراج رحلات من النص
            Logger.log('  📝 Claude-text: ' + textData.passengers.length + ' مسافر + ' + (textData.outboundLegs.length + textData.returnLegs.length) + ' رحلة');

            // حماية ذاتية
            var validationText = validateClaudeResult_(textData, allData);
            if (!validationText.valid) {
              Logger.log('  ⛔ رفض text — ' + validationText.reason);
              results.skipped++;
              continue;
            }

            // تحويل إلى legs قديم (للتوافق مع processTextUnified_)
            var convertedLegs = [];
            for (var ol = 0; ol < textData.outboundLegs.length; ol++) {
              var leg = textData.outboundLegs[ol];
              convertedLegs.push({
                direction: 'outbound',
                flightNumber: leg.flightNumber || '',
                fromCity: leg.fromCity,
                toCity: leg.toCity,
                depDate: leg.depDate,
                depTime: leg.depTime,
                arrDate: leg.arrDate,
                arrTime: leg.arrTime
              });
            }
            for (var rl = 0; rl < textData.returnLegs.length; rl++) {
              var leg2 = textData.returnLegs[rl];
              convertedLegs.push({
                direction: 'return',
                flightNumber: leg2.flightNumber || '',
                fromCity: leg2.fromCity,
                toCity: leg2.toCity,
                depDate: leg2.depDate,
                depTime: leg2.depTime,
                arrDate: leg2.arrDate,
                arrTime: leg2.arrTime
              });
            }

            var activePnr = textData.pnr || pnr;
            var wroteText = processTextUnified_(
              message, body, subject, date, incidentNum, activePnr, convertedLegs,
              allData, unifiedSheet, existingKeys
            );
            if (wroteText) {
              results.textProcessed++;
              threadProcessed = true;
            } else {
              results.skipped++;
            }
            continue;
          }

          // إذا Claude قال skip أو فشل، جرّب Regex القديم
          var textLegs = extractFlightLegs_(body);
          if (textLegs && textLegs.length > 0) {
            Logger.log('  📝 Regex-text — ' + textLegs.length + ' رحلة');
            var wroteText2 = processTextUnified_(
              message, body, subject, date, incidentNum, pnr, textLegs,
              allData, unifiedSheet, existingKeys
            );
            if (wroteText2) {
              results.textProcessed++;
              threadProcessed = true;
            }
          } else {
            var reason = textData && textData.reason ? textData.reason : 'no_content';
            Logger.log('  ⏭️ بلا بيانات رحلة — ' + reason);
            results.skipped++;
          }
          continue;
        }

        // === نوع A: PDF نسك ===
        for (var p = 0; p < nusukPDFs.length; p++) {
          var pdf = nusukPDFs[p];
          Logger.log('  📄 ' + pdf.getName());

          var driveFile = changesFolder.createFile(pdf);
          var shareLink = createShareLink_(driveFile);

          // V5 — استخدام Claude بدل OCR+Regex
          var ticketData = parseWithClaude_(pdf);

          // Fallback للطريقة القديمة إذا Claude فشل كلياً (null)
          if (!ticketData) {
            Logger.log('  ⚠️ Claude فشل — محاولة OCR القديم');
            var text = extractTextFromPDF_(driveFile.getId());
            ticketData = parseNusukTicket_(text);
          }

          if (!ticketData) {
            Logger.log('  ❌ لم يتم التعرف على التذكرة');
            results.errors++;
            continue;
          }

          // إذا Claude قال skip (ثقة منخفضة أو ليس تذكرة) → لا نكتب
          if (ticketData.skipped) {
            Logger.log('  ⏭️ تخطي — ' + ticketData.reason);
            results.skipped++;
            continue;
          }

          // حماية أدنى — validation قبل الكتابة
          var validation = validateClaudeResult_(ticketData, allData);
          if (!validation.valid) {
            Logger.log('  ⛔ رفض — ' + validation.reason);
            results.skipped++;
            continue;
          }

          if (!pnr && ticketData.pnr) pnr = ticketData.pnr;
          var activePnr = pnr || ticketData.pnr;

          var passengers = ticketData.passengers || [];
          if (passengers.length === 0 && (ticketData.firstName || ticketData.lastName)) {
            passengers = [{ firstName: ticketData.firstName, lastName: ticketData.lastName, fullName: ticketData.passengerName }];
          }

          Logger.log('  👥 ' + passengers.length + ' مسافر في PDF');

          for (var px = 0; px < passengers.length; px++) {
            var pax = passengers[px];

            // مفتاح V3 = PNR + Incident + firstName + lastName
            var dupKey = (
              (activePnr || '') + '|' +
              (incidentNum || '') + '|' +
              (pax.firstName || '') + '|' +
              (pax.lastName || '')
            ).toUpperCase();

            if (existingKeys[dupKey]) {
              Logger.log('  ⏭️ مكرر: ' + pax.firstName + ' ' + pax.lastName);
              continue;
            }

            Logger.log('  ✈️ ' + pax.firstName + ' ' + pax.lastName + ' | PNR: ' + activePnr);

            // مطابقة الحاج (نفس منطق V2)
            var match = findPilgrim_(pax.firstName, pax.lastName, activePnr, allData, ticketData.bookingId);
            if (!match) {
              var cleanedFull = deepClean_(pax.fullName || (pax.firstName + ' ' + pax.lastName));
              if (cleanedFull && cleanedFull !== (pax.firstName + ' ' + pax.lastName).trim()) {
                var cp = cleanedFull.split(/\s+/);
                var cf = cp[0] || '';
                var cl = cp.length >= 2 ? cp.slice(1).join(' ') : '';
                match = findPilgrim_(cf, cl, activePnr, allData, ticketData.bookingId);
                if (match) { pax.firstName = cf; pax.lastName = cl; pax.fullName = cleanedFull; }
              }
            }

            var status = match ? 'تم المطابقة' : 'لم يُطابَق';
            if (match) Logger.log('  ✅ ' + match.sysFirstName + ' ' + match.sysLastName);
            else Logger.log('  ⚠️ لم يُطابَق');

            var rowData = {
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
              curRetFlight1: match ? match.curRetFlight1 : '',
              curRetDate1: match ? match.curRetDate1 : '',
              pdfLink: shareLink,
              emailDate: date,
              status: status,
              source: 'PDF',
              notes: match ? '' : ('PNR: ' + activePnr)
            };

            writeUnifiedRow_(unifiedSheet, rowData);
            existingKeys[dupKey] = true;
            threadProcessed = true;
          }
          results.processed++;
        }
      } catch (e) {
        Logger.log('  ❌ خطأ: ' + e.message);
        results.errors++;
      }
    }

    if (threadProcessed) thread.addLabel(processedLabel);
    else thread.addLabel(skippedLabel);
  }

  Logger.log('\n=== V3 ✅ PDF: ' + results.processed + ' | 📝 نص: ' + results.textProcessed + ' | ⏭️ متجاهَل: ' + results.skipped + ' | ❌ ' + results.errors + ' ===');
  return results;
}


/**
 * معالجة إيميل نصّي للـ V3 — يكتب في الشيت الموحّد
 */
function processTextUnified_(message, body, subject, date, incidentNum, pnr, textLegs, allData, unifiedSheet, existingKeys) {
  // استخراج إيميل المستلم الأصلي
  var recipientEmail = '';
  var emailMatch = body.match(/(?:إلى|To)\s*:?\s*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})/);
  if (emailMatch) recipientEmail = emailMatch[1].toLowerCase();

  if (!pnr) {
    var bookMatch = body.match(/Booking\s+([A-Z0-9]{5,8})/i);
    if (bookMatch) {
      var candidate = bookMatch[1].toUpperCase();
      // تطبيق نفس القائمة السوداء لتجنب "NUMBER" / "RESERVAT" / إلخ
      var BLACKLIST2 = {
        'NUMBER':1,'NUMBERS':1,'BOOKING':1,'FLIGHT':1,'FLIGHTS':1,
        'TICKET':1,'TICKETS':1,'CHANGE':1,'CHANGES':1,'SCHEDULE':1,
        'PLEASE':1,'REGARDS':1,'THANKS':1,'DETAILS':1,'INVOICE':1,
        'DEPART':1,'ARRIVE':1,'AIRPORT':1,'CANCEL':1,'UPDATE':1,
        'UPDATED':1,'HISTORY':1,'ALERT':1,'NOTICE':1,'ABOUT':1,
        'WITHIN':1,'BECAUSE':1,'SUBJECT':1,'MESSAGE':1,'CONFIRM':1,
        'RESERVA':1,'RESERVAT':1,'REFEREN':1,'REFERNC':1
      };
      if (!BLACKLIST2[candidate]) pnr = candidate;
    }
  }

  // مفتاح V3 للنص — PNR + Incident + "TEXT" + recipientEmail
  var dupKey = (
    (pnr || '') + '|' +
    (incidentNum || '') + '|' +
    'TEXT|' +
    (recipientEmail || '')
  ).toUpperCase();

  if (existingKeys[dupKey]) {
    Logger.log('  ⏭️ مكرر نصّي: ' + incidentNum);
    return false;
  }

  // مطابقة
  var match = null;
  if (pnr) {
    for (var i = 0; i < allData.pd.length && !match; i++) {
      var contract = String(allData.pd[i][CONFIG.PD.CONTRACT_NAME] || '').toUpperCase();
      if (contract.indexOf(pnr) !== -1) {
        match = buildPDResult_(allData.pd[i]);
        match.source = 'PD (text PNR)';
        var cur = findCurrentFlights_(pnr, allData.flights);
        if (cur) {
          match.curOutFlight1 = cur.outFlight1; match.curOutDate1 = cur.outDate1;
          match.curRetFlight1 = cur.retFlight1; match.curRetDate1 = cur.retDate1;
        }
      }
    }
  }
  if (!match && recipientEmail) match = findByEmail_(recipientEmail, pnr, allData);

  var dataObj = { outboundLegs: [], returnLegs: [] };
  classifyLegs_(dataObj, textLegs);

  var status = match ? 'تم المطابقة' : 'لم يُطابَق';
  Logger.log('  ' + (match ? '✅' : '⚠️') + ' نصّي — PNR: ' + pnr + ' | ' + recipientEmail);

  var threadId = message.getThread().getId();
  var emailLink = 'https://mail.google.com/mail/u/0/#inbox/' + threadId;

  var rowData = {
    pnr: pnr || '',
    bookingId: '',
    incidentNum: incidentNum,
    pdfFirstName: '',
    pdfLastName: '',
    sysFirstName: match ? match.sysFirstName : '',
    sysLastName: match ? match.sysLastName : '',
    serialNum: match ? match.serial : '',
    passport: match ? match.passport : '',
    pkgNum: match ? match.pkgNum : '',
    pkgName: match ? match.pkgName : '',
    flightType: match ? match.flightType : '',
    bookingType: match ? (String(match.source || '').indexOf('B2C') !== -1 ? 'B2C' : 'B2B') : '',
    outboundLegs: dataObj.outboundLegs,
    returnLegs: dataObj.returnLegs,
    curOutFlight1: match ? match.curOutFlight1 : '',
    curOutDate1: match ? match.curOutDate1 : '',
    curRetFlight1: match ? match.curRetFlight1 : '',
    curRetDate1: match ? match.curRetDate1 : '',
    pdfLink: emailLink,
    emailDate: date,
    status: status,
    source: 'نص',
    notes: 'إيميل نصّي' + (recipientEmail ? ' | ' + recipientEmail : '')
  };

  writeUnifiedRow_(unifiedSheet, rowData);
  existingKeys[dupKey] = true;
  return true;
}
