// ============================================
// i18n — ملف الترجمة (عربي / إنجليزي)
// ============================================

var I18N = {

  // --- عام ---
  welcome: {
    ar: '👋 <b>مرحباً بك في IkramAdmin</b>\n\nنظام إدارة إكرام الضيف الذكي\n\nاختر اللغة:',
    en: '👋 <b>Welcome to IkramAdmin</b>\n\nSmart Ekram Aldyf Management System\n\nChoose language:'
  },
  error_general: {
    ar: '❌ حدث خطأ. حاول مرة أخرى.',
    en: '❌ An error occurred. Please try again.'
  },
  unauthorized: {
    ar: '🔒 غير مصرّح لك باستخدام هذا البوت.\n\nتواصل مع الإدارة للحصول على صلاحية.',
    en: '🔒 You are not authorized to use this bot.\n\nContact management for access.'
  },
  no_permission: {
    ar: '⛔ ليس لديك صلاحية لهذه العملية.',
    en: '⛔ You do not have permission for this action.'
  },

  // --- المصادقة ---
  auth_prompt: {
    ar: '🔐 <b>تسجيل الدخول</b>\n\nأرسل <b>رقم الجوال</b> أو <b>اسمك</b> المسجّل في النظام:',
    en: '🔐 <b>Login</b>\n\nSend your registered <b>phone number</b> or <b>name</b>:'
  },
  auth_success: {
    ar: '✅ <b>تم تسجيل الدخول بنجاح!</b>\n\n👤 {name}\n🏷️ الدور: {role}',
    en: '✅ <b>Login successful!</b>\n\n👤 {name}\n🏷️ Role: {role}'
  },
  auth_failed: {
    ar: '❌ لم يتم العثور على حسابك.\n\nتأكد من رقم الجوال أو الاسم وأعد المحاولة:',
    en: '❌ Account not found.\n\nCheck your phone number or name and try again:'
  },
  auth_admin_auto: {
    ar: '✅ <b>مرحباً أيها المدير!</b>\n\nتم التعرف عليك تلقائياً.',
    en: '✅ <b>Welcome, Admin!</b>\n\nYou were recognized automatically.'
  },

  // --- القائمة الرئيسية ---
  menu_title: {
    ar: '📋 <b>القائمة الرئيسية</b>\n\nاختر ما تريد أو اكتب سؤالك مباشرة:',
    en: '📋 <b>Main Menu</b>\n\nChoose an option or type your question:'
  },
  btn_dashboard:   { ar: '📊 لوحة القيادة',    en: '📊 Dashboard' },
  btn_search:      { ar: '🔍 بحث عن حاج',      en: '🔍 Search Pilgrim' },
  btn_packages:    { ar: '📦 حالة الباقات',     en: '📦 Package Status' },
  btn_flights:     { ar: '✈️ جدول الرحلات',     en: '✈️ Flight Schedule' },
  btn_hotels:      { ar: '🏨 إشغال الفنادق',    en: '🏨 Hotel Occupancy' },
  btn_camps:       { ar: '⛺ حالة المخيمات',    en: '⛺ Camp Status' },
  btn_sales:       { ar: '💰 تقرير المبيعات',   en: '💰 Sales Report' },
  btn_export:      { ar: '📤 تصدير ملف',        en: '📤 Export File' },
  btn_ai:          { ar: '💬 اسألني أي شيء',    en: '💬 Ask Me Anything' },
  btn_settings:    { ar: '⚙️ الإعدادات',        en: '⚙️ Settings' },
  btn_back:        { ar: '🔙 القائمة الرئيسية', en: '🔙 Main Menu' },

  // --- لوحة القيادة ---
  dashboard_title: {
    ar: '📊 <b>لوحة القيادة</b>\n━━━━━━━━━━━━━━━━━━',
    en: '📊 <b>Dashboard</b>\n━━━━━━━━━━━━━━━━━━'
  },
  dash_pilgrims:  { ar: '👥 إجمالي الحجاج',     en: '👥 Total Pilgrims' },
  dash_packages:  { ar: '📦 الباقات',            en: '📦 Packages' },
  dash_flights:   { ar: '✈️ الرحلات',            en: '✈️ Flights' },
  dash_sold:      { ar: '💰 نسبة البيع',         en: '💰 Sales Rate' },
  dash_arrivals:  { ar: '🛬 وصول اليوم',         en: '🛬 Arrivals Today' },
  dash_departures:{ ar: '🛫 مغادرة اليوم',       en: '🛫 Departures Today' },

  // --- بحث الحاج ---
  search_prompt: {
    ar: '🔍 أرسل <b>اسم الحاج</b> أو <b>رقم الجواز</b> أو <b>رقم الحجز</b>:',
    en: '🔍 Send <b>pilgrim name</b>, <b>passport number</b>, or <b>booking ID</b>:'
  },
  search_results: {
    ar: '🔍 نتائج البحث ({count}):',
    en: '🔍 Search results ({count}):'
  },
  search_no_results: {
    ar: '❌ لا توجد نتائج لـ "<b>{query}</b>"',
    en: '❌ No results for "<b>{query}</b>"'
  },
  search_too_many: {
    ar: '⚠️ وُجد {count} نتيجة. حدّد البحث أكثر:',
    en: '⚠️ Found {count} results. Please refine your search:'
  },

  // --- تفاصيل الحاج ---
  pilgrim_detail_title: {
    ar: '👤 <b>بيانات الحاج</b>\n━━━━━━━━━━━━━━━━━━',
    en: '👤 <b>Pilgrim Details</b>\n━━━━━━━━━━━━━━━━━━'
  },
  lbl_name:       { ar: 'الاسم',          en: 'Name' },
  lbl_passport:   { ar: 'الجواز',         en: 'Passport' },
  lbl_phone:      { ar: 'الجوال',         en: 'Phone' },
  lbl_email:      { ar: 'الإيميل',        en: 'Email' },
  lbl_nationality:{ ar: 'الجنسية',        en: 'Nationality' },
  lbl_package:    { ar: 'الباقة',          en: 'Package' },
  lbl_guide:      { ar: 'المرشد الديني',   en: 'Religious Guide' },
  lbl_camp:       { ar: 'المخيم',          en: 'Camp' },
  lbl_visa:       { ar: 'التأشيرة',       en: 'Visa' },
  lbl_ticket:     { ar: 'التذكرة',        en: 'Ticket' },
  lbl_transport:  { ar: 'التنقل',         en: 'Transport' },
  lbl_flight_arr: { ar: 'رحلة الوصول',    en: 'Arrival Flight' },
  lbl_flight_ret: { ar: 'رحلة العودة',    en: 'Return Flight' },
  lbl_hotel:      { ar: 'الفندق',         en: 'Hotel' },
  lbl_checkin:    { ar: 'الدخول',          en: 'Check-in' },
  lbl_checkout:   { ar: 'الخروج',         en: 'Check-out' },
  lbl_reception:  { ar: 'الاستقبال',      en: 'Reception' },
  lbl_received:   { ar: '✅ تم الاستقبال', en: '✅ Received' },
  lbl_not_received: { ar: '⏳ لم يصل بعد', en: '⏳ Not arrived yet' },

  // --- الباقات ---
  packages_title: {
    ar: '📦 <b>حالة الباقات</b>\n━━━━━━━━━━━━━━━━━━',
    en: '📦 <b>Package Status</b>\n━━━━━━━━━━━━━━━━━━'
  },
  lbl_capacity:  { ar: 'السعة', en: 'Capacity' },
  lbl_sold:      { ar: 'مباع',  en: 'Sold' },
  lbl_remaining: { ar: 'متبقي', en: 'Remaining' },

  // --- الرحلات ---
  flights_title: {
    ar: '✈️ <b>جدول الرحلات</b>\n━━━━━━━━━━━━━━━━━━',
    en: '✈️ <b>Flight Schedule</b>\n━━━━━━━━━━━━━━━━━━'
  },
  flights_today: { ar: 'رحلات اليوم', en: "Today's Flights" },
  flights_week:  { ar: 'رحلات الأسبوع', en: 'This Week' },
  btn_flights_today: { ar: '📅 اليوم',    en: '📅 Today' },
  btn_flights_week:  { ar: '📆 الأسبوع',  en: '📆 This Week' },
  btn_flights_all:   { ar: '📋 الكل',     en: '📋 All' },
  no_flights: {
    ar: '✈️ لا توجد رحلات {period}',
    en: '✈️ No flights {period}'
  },

  // --- الفنادق ---
  hotels_title: {
    ar: '🏨 <b>إشغال الفنادق</b>\n━━━━━━━━━━━━━━━━━━',
    en: '🏨 <b>Hotel Occupancy</b>\n━━━━━━━━━━━━━━━━━━'
  },

  // --- المخيمات ---
  camps_title: {
    ar: '⛺ <b>حالة المخيمات</b>\n━━━━━━━━━━━━━━━━━━',
    en: '⛺ <b>Camp Status</b>\n━━━━━━━━━━━━━━━━━━'
  },

  // --- المبيعات ---
  sales_title: {
    ar: '💰 <b>تقرير المبيعات</b>\n━━━━━━━━━━━━━━━━━━',
    en: '💰 <b>Sales Report</b>\n━━━━━━━━━━━━━━━━━━'
  },

  // --- التصدير ---
  export_prompt: {
    ar: '📤 اختر نوع التقرير المطلوب تصديره:',
    en: '📤 Choose the report type to export:'
  },
  export_generating: {
    ar: '⏳ جاري إنشاء الملف...',
    en: '⏳ Generating file...'
  },
  export_done: {
    ar: '✅ تم إنشاء الملف بنجاح!',
    en: '✅ File generated successfully!'
  },

  // --- AI ---
  ai_thinking: {
    ar: '🤔 جاري التحليل...',
    en: '🤔 Analyzing...'
  },
  ai_prompt: {
    ar: '💬 اكتب سؤالك بالعربي أو الإنجليزي:\n\nأمثلة:\n• كم حاج في باقة 5؟\n• أعطني رحلات هذا الأسبوع\n• ما نسبة البيع الإجمالية؟',
    en: '💬 Type your question in Arabic or English:\n\nExamples:\n• How many pilgrims in package 5?\n• Show me this week\'s flights\n• What is the overall sales rate?'
  },

  // --- الإعدادات ---
  settings_title: {
    ar: '⚙️ <b>الإعدادات</b>',
    en: '⚙️ <b>Settings</b>'
  },
  btn_change_lang: { ar: '🌐 تغيير اللغة', en: '🌐 Change Language' },
  btn_logout:      { ar: '🚪 تسجيل خروج',  en: '🚪 Logout' },

  // --- الإدارة ---
  btn_broadcast:  { ar: '📢 بث رسالة',       en: '📢 Broadcast' },
  btn_users:      { ar: '👥 إدارة المستخدمين', en: '👥 Manage Users' },
  btn_stats:      { ar: '📈 إحصائيات البوت',  en: '📈 Bot Stats' },

  // الأدوار
  role_admin:    { ar: 'مدير',  en: 'Admin' },
  role_manager:  { ar: 'مشرف',  en: 'Manager' },
  role_employee: { ar: 'موظف',  en: 'Employee' }
};

// ============================================
// دالة الترجمة
// ============================================
function T_(key, lang, params) {
  var entry = I18N[key];
  if (!entry) return '[' + key + ']';
  var text = entry[lang] || entry['ar'] || '';
  if (params) {
    for (var k in params) {
      text = text.replace(new RegExp('\\{' + k + '\\}', 'g'), params[k]);
    }
  }
  return text;
}
