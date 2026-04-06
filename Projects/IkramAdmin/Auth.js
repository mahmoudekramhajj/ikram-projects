// ============================================
// Auth.js — المصادقة وإدارة الجلسات والأدوار
// ============================================

// ============================================
// التحقق من الهوية
// ============================================

/**
 * التحقق مما إذا كان المستخدم مدير (بالـ Chat ID)
 */
function isAdmin_(chatId) {
  return ADMIN_IDS.indexOf(String(chatId)) !== -1;
}

/**
 * معالجة اختيار اللغة
 */
function handleLanguageSelection_(chatId, lang) {
  var session = getSession_(chatId);
  session.language = lang;

  // المدراء يدخلون مباشرة بدون مصادقة
  if (isAdmin_(chatId)) {
    session.status = 'verified';
    session.role = ROLES.ADMIN;
    session.name = 'Admin';
    session.userId = String(chatId);
    saveSession_(session); // كاش فقط — فوري
    sendMessage_(chatId, T_('auth_admin_auto', lang));
    sendMainMenu_(chatId, session);
    return;
  }

  session.status = 'awaiting_auth';
  saveSession_(session);
  sendMessage_(chatId, T_('auth_prompt', lang));
}

/**
 * معالجة إدخال بيانات التسجيل (رقم جوال أو اسم)
 */
function handleAuthInput_(chatId, text, session) {
  var lang = session.language || 'ar';
  var query = String(text).trim();

  // البحث في شيت المستخدمين
  var user = findUser_(query);

  if (!user) {
    sendMessage_(chatId, T_('auth_failed', lang));
    return;
  }

  // تحديث الجلسة
  session.status = 'verified';
  session.userId = user.id;
  session.name = user.name;
  session.role = user.role || ROLES.EMPLOYEE;
  session.department = user.department || '';
  saveSession_(session);

  var roleName = T_('role_' + session.role, lang);
  sendMessage_(chatId, T_('auth_success', lang, { name: session.name, role: roleName }));
  sendMainMenu_(chatId, session);
}

/**
 * البحث عن مستخدم في شيت المستخدمين
 */
function findUser_(query) {
  var data = getUsersData_();
  if (!data || data.length === 0) return null;

  var q = String(query).toLowerCase().trim();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row || !row[0]) continue;

    // البحث بالاسم أو الجوال أو الإيميل
    var name = String(row[1] || '').toLowerCase();
    var phone = String(row[2] || '').replace(/\s+/g, '');
    var email = String(row[3] || '').toLowerCase();
    var role = String(row[4] || '').toLowerCase();
    var dept = String(row[5] || '');

    if (q === phone || q === email || name.indexOf(q) !== -1) {
      return {
        id: String(row[0]),
        name: row[1] || '',
        phone: phone,
        email: email,
        role: role || ROLES.EMPLOYEE,
        department: dept
      };
    }
  }
  return null;
}

/**
 * قراءة بيانات المستخدمين (مع كاش)
 */
function getUsersData_() {
  var cached = getCache_('admin_users');
  if (cached) return cached;

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.USERS);
    if (!sheet) return [];

    var data = sheet.getDataRange().getValues();
    // تخطي الهيدر
    var rows = data.slice(1);
    setCache_('admin_users', rows, CACHE_TTL);
    return rows;
  } catch (e) {
    Logger.log('getUsersData_ error: ' + e.message);
    return [];
  }
}

// ============================================
// إدارة الجلسات
// ============================================

/**
 * جلب الجلسة من الكاش أولاً (سريع)
 * إذا لم توجد في الكاش، جلسة جديدة فوراً (بدون قراءة الشيت)
 * الشيت يُقرأ فقط عند الحاجة عبر loadSessionFromSheet_
 */
function getSession_(chatId) {
  var key = 'adm_session_' + String(chatId);
  var cached = getCache_(key);
  if (cached) return cached;

  // جلسة جديدة فوراً — بدون فتح الشيت (سرعة)
  var session = {
    chatId: String(chatId),
    userId: '',
    name: '',
    role: '',
    department: '',
    language: 'ar',
    status: 'new',
    lastActivity: new Date().toISOString(),
    inputState: ''
  };

  // للأدمن: تعيين فوري بدون شيت
  if (isAdmin_(chatId)) {
    session.status = 'verified';
    session.role = ROLES.ADMIN;
    session.name = 'Admin';
    session.userId = String(chatId);
  }

  setCache_(key, session, CACHE_TTL);
  return session;
}

/**
 * حفظ الجلسة في الكاش فقط (سريع)
 * الكتابة في الشيت تتم عبر saveSessionToSheet_ لاحقاً
 */
function saveSession_(session) {
  var key = 'adm_session_' + session.chatId;
  session.lastActivity = new Date().toISOString();
  setCache_(key, session, CACHE_TTL);
}

/**
 * حفظ الجلسة في الشيت (يُستدعى بعد الرد على المستخدم)
 */
function saveSessionToSheet_(session) {
  try {
    var sheet = getSessionSheet_();
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var found = false;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === session.chatId) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          session.chatId, session.userId, session.name,
          session.role, session.department, session.language,
          session.status, session.lastActivity
        ]]);
        found = true;
        break;
      }
    }

    if (!found) {
      sheet.appendRow([
        session.chatId, session.userId, session.name,
        session.role, session.department, session.language,
        session.status, session.lastActivity
      ]);
    }
  } catch (e) {
    Logger.log('saveSessionToSheet_ error: ' + e.message);
  }
}

/**
 * الحصول على شيت الجلسات (إنشاء تلقائي إذا لم يوجد)
 */
function getSessionSheet_() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEETS.BOT_SESSIONS);

    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.BOT_SESSIONS);
      sheet.appendRow([
        'ChatID', 'UserID', 'Name', 'Role',
        'Department', 'Language', 'Status', 'LastActivity'
      ]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    }

    return sheet;
  } catch (e) {
    Logger.log('getSessionSheet_ error: ' + e.message);
    return null;
  }
}

/**
 * تسجيل نشاط — يُخزّن في الكاش ويُكتب بالدفعات لاحقاً
 * لا يفتح الشيت أبداً أثناء الرد (سرعة)
 */
function logActivity_(chatId, action, query, duration) {
  try {
    var cache = CacheService.getScriptCache();
    var logKey = 'adm_log_' + Date.now();
    var entry = JSON.stringify({
      ts: new Date().toISOString(),
      chat: String(chatId),
      action: action,
      query: truncate_(query, 100),
      dur: duration || ''
    });
    cache.put(logKey, entry, 3600); // ساعة واحدة
  } catch (e) {
    // لا شيء — السجل غير حرج
  }
}

/**
 * التحقق من صلاحية عملية معينة
 */
function hasPermission_(session, feature) {
  var role = session.role;
  if (role === ROLES.ADMIN) return true;

  var restricted = {
    broadcast: [ROLES.ADMIN],
    user_mgmt: [ROLES.ADMIN],
    sales: [ROLES.ADMIN, ROLES.MANAGER]
  };

  if (restricted[feature]) {
    return restricted[feature].indexOf(role) !== -1;
  }

  return true; // الميزات غير المقيّدة متاحة للجميع
}

/**
 * تسجيل الخروج
 */
function handleLogout_(chatId, session) {
  var lang = session.language || 'ar';
  var key = 'adm_session_' + String(chatId);
  CacheService.getScriptCache().remove(key);

  session.status = 'new';
  session.userId = '';
  session.name = '';
  session.role = '';
  session.inputState = '';
  saveSession_(session);

  sendMessage_(chatId, T_('welcome', lang), {
    inline_keyboard: [
      [{ text: '🇸🇦 العربية', callback_data: 'lang_ar' },
       { text: '🇬🇧 English', callback_data: 'lang_en' }]
    ]
  });
}
