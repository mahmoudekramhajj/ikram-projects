// ============================================
// أوامر المدير — /broadcast + /stats
// ============================================

// ============================================
// التحقق من صلاحية المدير
// ============================================
function isAdmin_(chatId) {
  return ADMIN_IDS.indexOf(String(chatId)) !== -1;
}

// ============================================
// /broadcast — بدء إرسال رسالة جديدة
// ============================================
function startBroadcast_(chatId) {
  // تخزين بيانات البث المؤقتة في الكاش
  setCache_('admin_draft_' + chatId, {
    ar: '', en: '', fr: '', de: '', it: '', es: '',
    imageUrl: '', fileUrl: '', fileName: '',
    target: 'all', targetValue: '', priority: 'normal'
  }, 3600); // ساعة واحدة

  updateSession_(chatId, { inputState: 'admin_msg_ar' });
  sendMessage_(chatId, T_('admin_broadcast_start', 'ar'));
}

// ============================================
// معالجة مدخلات المدير (نص/صورة/ملف)
// ============================================
function handleAdminInput_(chatId, text, message, session) {
  var state = session.inputState || '';
  var draft = getCache_('admin_draft_' + chatId);
  if (!draft) {
    updateSession_(chatId, { inputState: '' });
    sendMessage_(chatId, '⚠️ انتهت صلاحية المسودة. أعد /broadcast');
    return;
  }

  var langOrder = ['ar', 'en', 'fr', 'de', 'it', 'es'];
  var langNames = { ar: 'العربية', en: 'English', fr: 'Français', de: 'Deutsch', it: 'Italiano', es: 'Español' };

  // === إدخال النصوص ===
  for (var i = 0; i < langOrder.length; i++) {
    var lang = langOrder[i];
    if (state === 'admin_msg_' + lang) {
      var skip = text.toLowerCase() === 'تخطي' || text.toLowerCase() === 'skip';

      if (!skip) {
        draft[lang] = text;
      }
      setCache_('admin_draft_' + chatId, draft, 3600);

      // الانتقال للغة التالية أو المرفقات
      var nextIdx = i + 1;
      if (nextIdx < langOrder.length) {
        var nextLang = langOrder[nextIdx];
        updateSession_(chatId, { inputState: 'admin_msg_' + nextLang });
        sendMessage_(chatId, T_('admin_broadcast_next_lang', 'ar', { lang: langNames[nextLang] }));
      } else {
        // انتهت اللغات → المرفقات
        updateSession_(chatId, { inputState: 'admin_msg_media' });
        sendMessage_(chatId, T_('admin_broadcast_media', 'ar'));
      }
      return;
    }
  }

  // === إدخال المرفقات (صورة أو ملف) ===
  if (state === 'admin_msg_media') {
    var skip = text.toLowerCase() === 'تخطي' || text.toLowerCase() === 'skip';

    if (!skip) {
      // التحقق من صورة
      if (message.photo && message.photo.length > 0) {
        var fileId = message.photo[message.photo.length - 1].file_id;
        var fileInfo = getFile_(fileId);
        if (fileInfo && fileInfo.file_path) {
          draft.imageUrl = 'https://api.telegram.org/file/bot' + BOT_TOKEN + '/' + fileInfo.file_path;
        }
      }
      // التحقق من ملف (PDF)
      if (message.document) {
        var docInfo = getFile_(message.document.file_id);
        if (docInfo && docInfo.file_path) {
          draft.fileUrl = 'https://api.telegram.org/file/bot' + BOT_TOKEN + '/' + docInfo.file_path;
          draft.fileName = message.document.file_name || 'document';
        }
      }
    }

    setCache_('admin_draft_' + chatId, draft, 3600);
    updateSession_(chatId, { inputState: 'admin_msg_target' });

    // أزرار الاستهداف
    sendMessage_(chatId, T_('admin_broadcast_target', 'ar'), {
      inline_keyboard: [
        [{ text: '👥 الكل', callback_data: 'admin_target_all' }],
        [{ text: '📦 باقة محددة', callback_data: 'admin_target_package' }],
        [{ text: '🌍 جنسية محددة', callback_data: 'admin_target_nationality' }],
        [{ text: '✈️ رحلة محددة', callback_data: 'admin_target_flight' }]
      ]
    });
    return;
  }

  // === إدخال قيمة الاستهداف ===
  if (state === 'admin_msg_target_value') {
    draft.targetValue = text.trim();
    setCache_('admin_draft_' + chatId, draft, 3600);
    showBroadcastConfirmation_(chatId, draft);
    return;
  }
}

// ============================================
// معالجة أزرار المدير
// ============================================
function handleAdminCallback_(chatId, data, session) {
  if (!isAdmin_(chatId)) return;

  var draft = getCache_('admin_draft_' + chatId);
  if (!draft) return;

  // === اختيار الاستهداف ===
  if (data === 'admin_target_all') {
    draft.target = 'all';
    setCache_('admin_draft_' + chatId, draft, 3600);
    showBroadcastConfirmation_(chatId, draft);
    return;
  }

  if (data === 'admin_target_package' || data === 'admin_target_nationality' || data === 'admin_target_flight') {
    var target = data.replace('admin_target_', '');
    draft.target = target;
    setCache_('admin_draft_' + chatId, draft, 3600);
    updateSession_(chatId, { inputState: 'admin_msg_target_value' });

    var labels = { package: 'رقم الباقة', nationality: 'الجنسية (بالإنجليزية)', flight: 'رقم الرحلة' };
    sendMessage_(chatId, T_('admin_broadcast_target_value', 'ar', { target: labels[target] || target }));
    return;
  }

  // === تأكيد الإرسال ===
  if (data === 'admin_confirm_send') {
    updateSession_(chatId, { inputState: '' });

    // حفظ في AdminMessages وإرسال
    var msgId = addAdminMessage_({
      title: (draft.ar || draft.en || '').substring(0, 50),
      ar: draft.ar, en: draft.en, fr: draft.fr, de: draft.de, it: draft.it, es: draft.es,
      imageUrl: draft.imageUrl, fileUrl: draft.fileUrl, fileName: draft.fileName,
      target: draft.target, targetValue: draft.targetValue, priority: draft.priority
    });

    // إرسال فوري (بدل انتظار trigger)
    processAdminMessages_();

    var cache = CacheService.getScriptCache();
    cache.remove('admin_draft_' + chatId);
    return;
  }

  // === إلغاء ===
  if (data === 'admin_confirm_cancel') {
    updateSession_(chatId, { inputState: '' });
    var cache = CacheService.getScriptCache();
    cache.remove('admin_draft_' + chatId);
    sendMessage_(chatId, T_('admin_broadcast_cancelled', 'ar'));
    return;
  }

  // === تحديد أولوية عاجلة ===
  if (data === 'admin_priority_urgent') {
    draft.priority = 'urgent';
    setCache_('admin_draft_' + chatId, draft, 3600);
    showBroadcastConfirmation_(chatId, draft);
    return;
  }
}

// ============================================
// عرض ملخص الرسالة قبل الإرسال
// ============================================
function showBroadcastConfirmation_(chatId, draft) {
  updateSession_(chatId, { inputState: 'admin_msg_confirm' });

  var preview = (draft.ar || draft.en || '').substring(0, 100);
  if (preview.length === 100) preview += '...';

  var targetLabels = { all: '👥 الكل', package: '📦 باقة: ' + draft.targetValue, nationality: '🌍 جنسية: ' + draft.targetValue, flight: '✈️ رحلة: ' + draft.targetValue };
  var targetLabel = targetLabels[draft.target] || draft.target;

  var mediaLabel = 'لا';
  if (draft.imageUrl) mediaLabel = '📸 صورة';
  if (draft.fileUrl) mediaLabel += (mediaLabel !== 'لا' ? ' + ' : '') + '📄 ملف';

  var text = T_('admin_broadcast_confirm', 'ar', { preview: preview, target: targetLabel, media: mediaLabel });

  sendMessage_(chatId, text, {
    inline_keyboard: [
      [{ text: '🚨 عاجل', callback_data: 'admin_priority_urgent' }],
      [{ text: '✅ إرسال الآن', callback_data: 'admin_confirm_send' }],
      [{ text: '❌ إلغاء', callback_data: 'admin_confirm_cancel' }]
    ]
  });
}

// ============================================
// /stats — إحصائيات البوت
// ============================================
function handleStats_(chatId) {
  var stats = getBotStats_();

  // آخر رسالة مرسلة
  var lastMsg = '';
  try {
    var sheet = getAdminMessagesSheet_();
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][AM.STATUS]).toLowerCase() === 'sent') {
        lastMsg = String(data[i][AM.TITLE] || '') + ' — ' + (data[i][AM.SENT_COUNT] || 0) + ' مستلم';
        break;
      }
    }
  } catch (e) {}
  if (!lastMsg) lastMsg = 'لا توجد';

  var text = '📊 <b>إحصائيات البوت</b>\n━━━━━━━━━━━━━━\n\n' +
    '👥 مسجلين: <b>' + stats.total + '</b>\n' +
    '✅ نشطين: <b>' + stats.active + '</b>\n' +
    '🚫 حظروا البوت: <b>' + stats.blocked + '</b>\n' +
    '💤 غير نشطين: <b>' + stats.inactive + '</b>\n\n' +
    '📱 سجلوا جوال: <b>' + stats.phone + '</b>\n' +
    '🏨 سجلوا غرف: <b>' + stats.rooms + '</b>\n' +
    '📷 رفعوا صورة: <b>' + stats.photo + '</b>\n\n' +
    '📢 آخر رسالة: ' + lastMsg;

  sendMessage_(chatId, text);
}
