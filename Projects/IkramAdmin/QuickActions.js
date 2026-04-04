// ============================================
// QuickActions.js — معالجة الأزرار والاستعلامات
// ============================================

// ============================================
// لوحة القيادة
// ============================================

function handleDashboard_(chatId, session) {
  var lang = session.language || 'ar';
  var stats = getDashboardStats_();

  var soldBar = buildProgressBar_(stats.salesPercent);

  var text = T_('dashboard_title', lang) + '\n\n';
  text += '👥 ' + T_('dash_pilgrims', lang) + ': <b>' + formatNumber_(stats.totalPilgrims) + '</b>\n';
  text += '📦 ' + T_('dash_packages', lang) + ': <b>' + stats.totalPackages + '</b>\n';
  text += '✈️ ' + T_('dash_flights', lang) + ': <b>' + stats.totalFlights + '</b> (' + stats.countries + ' ' + (lang === 'ar' ? 'دولة' : 'countries') + ')\n';
  text += '💰 ' + T_('dash_sold', lang) + ': <b>' + stats.salesPercent + '%</b> (' + formatNumber_(stats.totalSold) + '/' + formatNumber_(stats.totalCapacity) + ')\n';
  text += soldBar + '\n\n';
  text += '🛬 ' + T_('dash_arrivals', lang) + ': <b>' + stats.arrivalsToday + '</b>\n';
  text += '🛫 ' + T_('dash_departures', lang) + ': <b>' + stats.departuresToday + '</b>\n';
  text += '✅ ' + (lang === 'ar' ? 'تم استقبالهم' : 'Received') + ': <b>' + formatNumber_(stats.received) + '</b>';

  sendMessage_(chatId, text, backButton_(lang));
}

// ============================================
// بحث الحاج
// ============================================

function handlePilgrimSearch_(chatId, query, session) {
  var lang = session.language || 'ar';
  var results = searchPilgrims_(query);

  if (results.length === 0) {
    sendMessage_(chatId, T_('search_no_results', lang, { query: query }), backButton_(lang));
    return;
  }

  if (results.length === 1) {
    handlePilgrimDetail_(chatId, results[0].seq, session);
    return;
  }

  if (results.length > 10) {
    sendMessage_(chatId, T_('search_too_many', lang, { count: results.length }), backButton_(lang));
    return;
  }

  // عرض قائمة النتائج
  var text = T_('search_results', lang, { count: results.length }) + '\n\n';
  var keyboard = [];

  for (var i = 0; i < results.length; i++) {
    var r = results[i];
    var displayName = r.nameEn || r.nameAr || r.passport;
    text += (i + 1) + '. <b>' + displayName + '</b> — ' + r.passport + '\n';
    text += '   📦 ' + (r.packageName || '-') + '\n';

    keyboard.push([{
      text: (i + 1) + '. ' + truncate_(displayName, 25),
      callback_data: 'pilgrim_' + r.seq
    }]);
  }

  keyboard.push([{ text: T_('btn_back', lang), callback_data: 'back_main' }]);
  sendMessage_(chatId, text, { inline_keyboard: keyboard });
}

// ============================================
// تفاصيل حاج كاملة
// ============================================

function handlePilgrimDetail_(chatId, seq, session) {
  var lang = session.language || 'ar';
  var p = getPilgrimFull_(seq);

  if (!p) {
    sendMessage_(chatId, T_('search_no_results', lang, { query: seq }), backButton_(lang));
    return;
  }

  var name = lang === 'ar' ? (p.nameAr || p.nameEn) : (p.nameEn || p.nameAr);

  var text = T_('pilgrim_detail_title', lang) + '\n\n';
  text += '👤 ' + T_('lbl_name', lang) + ': <b>' + name + '</b>\n';
  text += '🛂 ' + T_('lbl_passport', lang) + ': <code>' + p.passport + '</code>\n';
  if (p.phone) text += '📱 ' + T_('lbl_phone', lang) + ': ' + p.phone + '\n';
  if (p.email) text += '📧 ' + T_('lbl_email', lang) + ': ' + p.email + '\n';
  text += '🌍 ' + T_('lbl_nationality', lang) + ': ' + p.nationality + '\n';
  text += '📦 ' + T_('lbl_package', lang) + ': <b>' + (p.packageName || '-') + '</b>\n';
  text += '🧕 ' + T_('lbl_guide', lang) + ': ' + (p.guideName || '-') + '\n';
  text += '⛺ ' + T_('lbl_camp', lang) + ': ' + (p.camp || '-') + '\n';
  text += '📋 ' + T_('lbl_visa', lang) + ': ' + (p.visaStatus || '-') + '\n';
  if (p.ticketNo) text += '🎫 ' + T_('lbl_ticket', lang) + ': ' + p.ticketNo + '\n';

  // النقل
  if (p.intercityTransport) {
    text += '🚌 ' + T_('lbl_transport', lang) + ' (' + (lang === 'ar' ? 'بين المدن' : 'inter-city') + '): <b>' + p.intercityTransport + '</b>\n';
  }

  // الطيران
  if (p.flight) {
    text += '\n' + T_('lbl_flight_arr', lang) + ':\n';
    text += '  ✈️ ' + (p.flight.arrAirline || '') + ' ' + (p.flight.arrFlightNo || '') + '\n';
    text += '  📅 ' + formatDate_(p.flight.arrDate) + ' ' + formatTime_(p.flight.arrTime) + '\n';
    text += '  🛫 ' + (p.flight.arrFrom || '') + ' → 🛬 ' + (p.flight.arrTo || '') + '\n';

    text += '\n' + T_('lbl_flight_ret', lang) + ':\n';
    text += '  ✈️ ' + (p.flight.retAirline || '') + ' ' + (p.flight.retFlightNo || '') + '\n';
    text += '  📅 ' + formatDate_(p.flight.retDate) + ' ' + formatTime_(p.flight.retTime) + '\n';
    text += '  🛫 ' + (p.flight.retFrom || '') + ' → 🛬 ' + (p.flight.retTo || '') + '\n';
  }

  // الفنادق
  if (p.hotels) {
    text += '\n🏨 ' + T_('lbl_hotel', lang) + ':\n';
    if (p.hotels.first && p.hotels.first.name) {
      text += '  1️⃣ <b>' + p.hotels.first.name + '</b>\n';
      text += '     ' + T_('lbl_checkin', lang) + ': ' + formatDate_(p.hotels.first.start);
      text += ' → ' + T_('lbl_checkout', lang) + ': ' + formatDate_(p.hotels.first.end) + '\n';
    }
    if (p.hotels.last && p.hotels.last.name) {
      text += '  2️⃣ <b>' + p.hotels.last.name + '</b>\n';
      text += '     ' + T_('lbl_checkin', lang) + ': ' + formatDate_(p.hotels.last.start);
      text += ' → ' + T_('lbl_checkout', lang) + ': ' + formatDate_(p.hotels.last.end) + '\n';
    }
  }

  // الاستقبال
  if (p.reception) {
    text += '\n' + T_('lbl_reception', lang) + ': ';
    if (p.reception.status === 'received') {
      text += T_('lbl_received', lang) + ' (' + formatDate_(p.reception.time) + ')';
    } else {
      text += T_('lbl_not_received', lang);
    }
  }

  sendMessage_(chatId, text, backButton_(lang));
}

// ============================================
// حالة الباقات
// ============================================

function handlePackageStatus_(chatId, session) {
  var lang = session.language || 'ar';
  var stats = getPackagesStats_();

  var text = T_('packages_title', lang) + '\n\n';
  text += '📊 ' + (lang === 'ar' ? 'إجمالي' : 'Total') + ': ';
  text += '<b>' + formatNumber_(stats.totalSold) + '/' + formatNumber_(stats.totalCapacity) + '</b>';
  text += ' (' + stats.totalPercent + '%)\n';
  text += buildProgressBar_(stats.totalPercent) + '\n\n';

  for (var i = 0; i < stats.packages.length; i++) {
    var p = stats.packages[i];
    var name = lang === 'ar' ? (p.nameAr || p.nameEn) : (p.nameEn || p.nameAr);
    text += p.icon + ' <b>' + truncate_(name, 30) + '</b>\n';
    text += '   ' + T_('lbl_capacity', lang) + ': ' + p.capacity;
    text += ' | ' + T_('lbl_sold', lang) + ': ' + p.sold;
    text += ' | ' + T_('lbl_remaining', lang) + ': ' + p.remaining;
    text += ' (' + p.percent + '%)\n';
  }

  sendMessage_(chatId, text, backButton_(lang));
}

// ============================================
// جدول الرحلات
// ============================================

function handleFlightSchedule_(chatId, session, period) {
  var lang = session.language || 'ar';
  var flights = getFlightsByPeriod_(period);

  var periodLabel = period === 'today' ? T_('flights_today', lang) :
                    period === 'week' ? T_('flights_week', lang) : (lang === 'ar' ? 'جميع الرحلات' : 'All Flights');

  if (flights.length === 0) {
    sendMessage_(chatId, T_('no_flights', lang, { period: periodLabel }), backButton_(lang));
    return;
  }

  var text = T_('flights_title', lang) + '\n';
  text += '📋 ' + periodLabel + ' (' + flights.length + ')\n\n';

  var max = Math.min(flights.length, 15);
  for (var i = 0; i < max; i++) {
    var f = flights[i];
    text += '✈️ <b>' + f.pnr + '</b> — ' + (f.airline || f.supplier) + '\n';
    text += '   🌍 ' + (f.country || '') + ' | 👥 ' + (f.pax || 0) + ' PAX\n';
    if (f.go1 && f.go1.date) {
      text += '   🛫 ' + (f.go1.from || '') + ' → ' + (f.go1.to || '');
      text += ' | ' + formatDate_(f.go1.date) + ' ' + formatTime_(f.go1.time) + '\n';
    }
    if (f.ret1 && f.ret1.date) {
      text += '   🛬 ' + (f.ret1.from || '') + ' → ' + (f.ret1.to || '');
      text += ' | ' + formatDate_(f.ret1.date) + ' ' + formatTime_(f.ret1.time) + '\n';
    }
    text += '\n';
  }

  if (flights.length > max) {
    text += '... ' + (lang === 'ar' ? 'و ' + (flights.length - max) + ' رحلة أخرى' : 'and ' + (flights.length - max) + ' more');
  }

  sendMessage_(chatId, text, {
    inline_keyboard: [
      [{ text: T_('btn_flights_today', lang), callback_data: 'flights_today' },
       { text: T_('btn_flights_week', lang), callback_data: 'flights_week' }],
      [{ text: T_('btn_back', lang), callback_data: 'back_main' }]
    ]
  });
}

// ============================================
// إشغال الفنادق
// ============================================

function handleHotelOccupancy_(chatId, session) {
  var lang = session.language || 'ar';
  var hotels = getHotelOccupancy_();

  if (hotels.length === 0) {
    sendMessage_(chatId, T_('hotels_title', lang) + '\n\n' + (lang === 'ar' ? 'لا توجد بيانات' : 'No data'), backButton_(lang));
    return;
  }

  var text = T_('hotels_title', lang) + '\n\n';

  // تجميع حسب المدينة
  var makkah = hotels.filter(function(h) { return String(h.city).match(/مك|Mak/i); });
  var madinah = hotels.filter(function(h) { return String(h.city).match(/مد|Med/i); });
  var other = hotels.filter(function(h) { return !String(h.city).match(/مك|Mak|مد|Med/i); });

  if (makkah.length > 0) {
    text += '🕋 <b>' + (lang === 'ar' ? 'مكة المكرمة' : 'Makkah') + '</b>\n';
    text += formatHotelList_(makkah, lang);
  }
  if (madinah.length > 0) {
    text += '\n🕌 <b>' + (lang === 'ar' ? 'المدينة المنورة' : 'Madinah') + '</b>\n';
    text += formatHotelList_(madinah, lang);
  }
  if (other.length > 0) {
    text += '\n📍 <b>' + (lang === 'ar' ? 'أخرى' : 'Other') + '</b>\n';
    text += formatHotelList_(other, lang);
  }

  sendMessage_(chatId, text, backButton_(lang));
}

function formatHotelList_(hotels, lang) {
  var text = '';
  for (var i = 0; i < hotels.length; i++) {
    var h = hotels[i];
    text += '  🏨 <b>' + h.name + '</b>';
    if (h.nameEn) text += ' (' + h.nameEn + ')';
    text += '\n';
    text += '     🚪 ' + (lang === 'ar' ? 'غرف' : 'Rooms') + ': ' + h.totalRooms;
    text += ' | 🛏️ ' + (lang === 'ar' ? 'أسرّة' : 'Beds') + ': ' + h.totalBeds + '\n';
  }
  return text;
}

// ============================================
// حالة المخيمات
// ============================================

function handleCampStatus_(chatId, session) {
  var lang = session.language || 'ar';
  var camps = getCampStats_();

  if (camps.length === 0) {
    sendMessage_(chatId, T_('camps_title', lang) + '\n\n' + (lang === 'ar' ? 'لا توجد بيانات' : 'No data'), backButton_(lang));
    return;
  }

  var total = 0;
  for (var i = 0; i < camps.length; i++) total += camps[i].count;

  var text = T_('camps_title', lang) + '\n\n';
  text += '👥 ' + (lang === 'ar' ? 'إجمالي المخصصين' : 'Total assigned') + ': <b>' + formatNumber_(total) + '</b>\n\n';

  for (var j = 0; j < camps.length; j++) {
    var c = camps[j];
    text += '⛺ <b>' + c.name + '</b>: ' + formatNumber_(c.count) + ' ' + (lang === 'ar' ? 'حاج' : 'pilgrims') + '\n';
  }

  sendMessage_(chatId, text, backButton_(lang));
}

// ============================================
// تقرير المبيعات
// ============================================

function handleSalesReport_(chatId, session) {
  var lang = session.language || 'ar';
  var sales = getSalesStats_();
  var pkg = sales.packages;

  var text = T_('sales_title', lang) + '\n\n';

  // ملخص عام
  text += '📊 <b>' + (lang === 'ar' ? 'ملخص عام' : 'Overview') + '</b>\n';
  text += '📦 ' + (lang === 'ar' ? 'الباقات' : 'Packages') + ': ' + pkg.packages.length + '\n';
  text += '✈️ ' + (lang === 'ar' ? 'الرحلات' : 'Flights') + ': ' + sales.totalFlights + '\n';
  text += '👥 ' + (lang === 'ar' ? 'مباع' : 'Sold') + ': <b>' + formatNumber_(pkg.totalSold) + '/' + formatNumber_(pkg.totalCapacity) + '</b> (' + pkg.totalPercent + '%)\n';
  text += buildProgressBar_(pkg.totalPercent) + '\n\n';

  // حسب الدولة
  text += '🌍 <b>' + (lang === 'ar' ? 'حسب الدولة' : 'By Country') + '</b>\n';
  var max = Math.min(sales.byCountry.length, 10);
  for (var i = 0; i < max; i++) {
    var c = sales.byCountry[i];
    text += '  🏳️ ' + c.country + ': ' + formatNumber_(c.sold) + ' ' + (lang === 'ar' ? 'مباع' : 'sold');
    text += ' | ' + c.flights + ' ' + (lang === 'ar' ? 'رحلة' : 'flights') + '\n';
  }

  sendMessage_(chatId, text, backButton_(lang));
}

// ============================================
// شريط التقدم
// ============================================

function buildProgressBar_(percent) {
  var filled = Math.round(percent / 10);
  var empty = 10 - filled;
  var bar = '';
  for (var i = 0; i < filled; i++) bar += '▓';
  for (var j = 0; j < empty; j++) bar += '░';
  return '[' + bar + '] ' + percent + '%';
}
