/**
 * إدارة طلبات المتابعة لوكيل إكرام
 * الإصدار 1.0
 */

/**
 * حفظ طلب متابعة جديد
 */
function saveLead(leadData) {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.LEADS);
    if (!sheet) {
      Logger.log('Leads sheet not found');
      return false;
    }
    
    var lastRow = sheet.getLastRow();
    var newId = lastRow > 0 ? lastRow : 1;
    
    sheet.appendRow([
      newId,                              // ID
      new Date(),                         // التاريخ
      leadData.platform || '',            // المنصة
      leadData.userId || '',              // رقم المستخدم
      leadData.name || '',                // الاسم
      leadData.email || '',               // البريد
      leadData.package || '',             // الباقة
      leadData.nuskNo || '',              // رقم نسك
      leadData.persons || '',             // عدد الأشخاص
      leadData.country || '',             // الدولة
      leadData.contactMethod || '',       // طريقة التواصل
      leadData.notes || '',               // ملاحظات
      'جديد',                             // الحالة
      '',                                 // الموظف المسؤول
      ''                                  // تاريخ المتابعة
    ]);
    
    // إرسال الإشعارات
    sendLeadNotifications(leadData, newId);
    
    return true;
    
  } catch (e) {
    Logger.log('saveLead Error: ' + e.toString());
    return false;
  }
}

/**
 * إرسال إشعارات عند طلب متابعة جديد
 */
function sendLeadNotifications(leadData, leadId) {
  try {
    var settings = getAllAgentSettings();
    
    // تجهيز نص الإشعار
    var notificationText = '🔔 *طلب متابعة جديد #' + leadId + '*\n\n' +
      '👤 الاسم: ' + (leadData.name || 'غير محدد') + '\n' +
      '📱 الرقم: ' + (leadData.userId || 'غير محدد') + '\n' +
      '📧 البريد: ' + (leadData.email || 'غير محدد') + '\n' +
      '📦 الباقة: ' + (leadData.package || 'غير محدد') + '\n' +
      '👥 عدد الأشخاص: ' + (leadData.persons || 'غير محدد') + '\n' +
      '🌍 الدولة: ' + (leadData.country || 'غير محدد') + '\n' +
      '📞 طريقة التواصل المفضلة: ' + (leadData.contactMethod || 'غير محدد') + '\n' +
      '💬 ملاحظات: ' + (leadData.notes || 'لا يوجد') + '\n' +
      '📲 المنصة: ' + (leadData.platform || 'غير محدد') + '\n\n' +
      '⏰ ' + new Date().toLocaleString('ar-SA');
    
    // إرسال إشعار WhatsApp
    if (settings.NOTIFICATION_PHONE) {
      sendWhatsAppNotification(notificationText);
    }
    
    // إرسال إشعار بريد إلكتروني
    if (settings.NOTIFICATION_EMAIL) {
      sendEmailNotification(leadData, leadId, settings.NOTIFICATION_EMAIL);
    }
    
  } catch (e) {
    Logger.log('sendLeadNotifications Error: ' + e.toString());
  }
}

/**
 * إرسال إشعار بريد إلكتروني
 */
function sendEmailNotification(leadData, leadId, email) {
  try {
    var subject = '🔔 طلب متابعة جديد #' + leadId + ' - ' + (leadData.name || 'عميل جديد');
    
    var htmlBody = '<div dir="rtl" style="font-family: Arial, sans-serif; padding: 20px;">' +
      '<h2 style="color: #2e7d32;">🔔 طلب متابعة جديد #' + leadId + '</h2>' +
      '<table style="border-collapse: collapse; width: 100%; max-width: 500px;">' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>👤 الاسم</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.name || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>📱 الرقم</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.userId || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>📧 البريد</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.email || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>📦 الباقة</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.package || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>👥 عدد الأشخاص</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.persons || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>🌍 الدولة</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.country || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>📞 طريقة التواصل</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.contactMethod || 'غير محدد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>💬 ملاحظات</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.notes || 'لا يوجد') + '</td></tr>' +
      '<tr><td style="padding: 10px; border: 1px solid #ddd; background: #f5f5f5;"><strong>📲 المنصة</strong></td>' +
      '<td style="padding: 10px; border: 1px solid #ddd;">' + (leadData.platform || 'غير محدد') + '</td></tr>' +
      '</table>' +
      '<p style="margin-top: 20px; color: #666;">⏰ ' + new Date().toLocaleString('ar-SA') + '</p>' +
      '<hr style="margin: 20px 0;">' +
      '<p style="color: #999; font-size: 12px;">وكيل إكرام - شركة إكرام الضيف للسياحة</p>' +
      '</div>';
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    
    Logger.log('Email notification sent to: ' + email);
    return true;
    
  } catch (e) {
    Logger.log('sendEmailNotification Error: ' + e.toString());
    return false;
  }
}

/**
 * تحديث حالة طلب المتابعة
 */
function updateLeadStatus(leadId, status, employee) {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.LEADS);
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == leadId) {
        sheet.getRange(i + 1, 13).setValue(status);           // الحالة
        sheet.getRange(i + 1, 14).setValue(employee || '');   // الموظف
        sheet.getRange(i + 1, 15).setValue(new Date());       // تاريخ المتابعة
        return true;
      }
    }
    
    return false;
    
  } catch (e) {
    Logger.log('updateLeadStatus Error: ' + e.toString());
    return false;
  }
}

/**
 * جلب الطلبات الجديدة
 */
function getNewLeads() {
  try {
    var ss = SpreadsheetApp.openById(AGENT_CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(AGENT_CONFIG.SHEETS.LEADS);
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var newLeads = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][12] === 'جديد') {
        newLeads.push({
          id: data[i][0],
          date: data[i][1],
          platform: data[i][2],
          userId: data[i][3],
          name: data[i][4],
          email: data[i][5],
          package: data[i][6],
          persons: data[i][8],
          country: data[i][9],
          contactMethod: data[i][10]
        });
      }
    }
    
    return newLeads;
    
  } catch (e) {
    Logger.log('getNewLeads Error: ' + e.toString());
    return [];
  }
}

/**
 * اختبار حفظ طلب متابعة
 */
function testSaveLead() {
  var testLead = {
    platform: 'Test',
    userId: '+966501234567',
    name: 'عميل تجريبي',
    email: 'test@example.com',
    package: 'باقة النخبة',
    nuskNo: '1001',
    persons: '3',
    country: 'السعودية',
    contactMethod: 'WhatsApp',
    notes: 'هذا اختبار للنظام'
  };
  
  var result = saveLead(testLead);
  
  if (result) {
    Logger.log('✅ Test lead saved successfully!');
  } else {
    Logger.log('❌ Failed to save test lead');
  }
  
  return result;
}

/**
 * إرسال تقرير يومي بالطلبات الجديدة
 */
function sendDailyLeadsReport() {
  try {
    var newLeads = getNewLeads();
    
    if (newLeads.length === 0) {
      Logger.log('No new leads to report');
      return;
    }
    
    var settings = getAllAgentSettings();
    
    var reportText = '📊 *تقرير الطلبات اليومي*\n\n' +
      '📅 التاريخ: ' + new Date().toLocaleDateString('ar-SA') + '\n' +
      '🔢 عدد الطلبات الجديدة: ' + newLeads.length + '\n\n';
    
    for (var i = 0; i < newLeads.length; i++) {
      var lead = newLeads[i];
      reportText += (i + 1) + '. ' + lead.name + ' - ' + lead.package + '\n';
    }
    
    reportText += '\n📱 راجع شيت طلبات_المتابعة للتفاصيل';
    
    // إرسال للواتساب
    if (settings.NOTIFICATION_PHONE) {
      sendWhatsAppNotification(reportText);
    }
    
    // إرسال للبريد
    if (settings.NOTIFICATION_EMAIL) {
      MailApp.sendEmail({
        to: settings.NOTIFICATION_EMAIL,
        subject: '📊 تقرير الطلبات اليومي - ' + newLeads.length + ' طلب جديد',
        body: reportText.replace(/\*/g, '').replace(/\n/g, '\r\n')
      });
    }
    
    Logger.log('Daily report sent with ' + newLeads.length + ' leads');
    
  } catch (e) {
    Logger.log('sendDailyLeadsReport Error: ' + e.toString());
  }
}