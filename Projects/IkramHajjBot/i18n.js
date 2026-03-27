// ============================================
// i18n — ملف الترجمة المركزي (ar/en/fr/nl)
// ============================================

var I18N = {

  // --- عام ---
  start_prompt: {
    ar: 'أرسل /start للبدء',
    en: 'Send /start to begin',
    fr: 'Envoyez /start pour commencer',
    nl: 'Stuur /start om te beginnen'
  },
  start_prompt_new: {
    ar: 'أرسل /start للبدء من جديد',
    en: 'Send /start to start over',
    fr: 'Envoyez /start pour recommencer',
    nl: 'Stuur /start om opnieuw te beginnen'
  },
  data_not_found: {
    ar: '❌ لم يتم العثور على بياناتك',
    en: '❌ Data not found',
    fr: '❌ Données non trouvées',
    nl: '❌ Gegevens niet gevonden'
  },
  faq_wip: {
    ar: '⚙️ قسم الأسئلة العامة قيد التطوير...',
    en: '⚙️ FAQ section is under development...',
    fr: '⚙️ Section FAQ en cours de développement...',
    nl: '⚙️ FAQ-sectie is in ontwikkeling...'
  },

  // --- المصادقة ---
  passport_prompt: {
    ar: '🔐 للوصول لبياناتك الشخصية:\n\nأرسل <b>رقم جواز السفر</b> الخاص بك',
    en: '🔐 To access your personal data:\n\nSend your <b>passport number</b>',
    fr: '🔐 Pour accéder à vos données:\n\nEnvoyez votre <b>numéro de passeport</b>',
    nl: '🔐 Om uw gegevens te bekijken:\n\nStuur uw <b>paspoortnummer</b>'
  },
  passport_invalid: {
    ar: '⚠️ رقم الجواز غير صحيح. تأكد وأعد الإرسال:',
    en: '⚠️ Invalid passport number. Please try again:',
    fr: '⚠️ Numéro de passeport invalide. Réessayez:',
    nl: '⚠️ Ongeldig paspoortnummer. Probeer opnieuw:'
  },
  passport_not_found: {
    ar: '❌ رقم الجواز غير موجود في النظام.\n\nتأكد من الرقم وأعد الإرسال:',
    en: '❌ Passport not found in our system.\n\nPlease check and try again:',
    fr: '❌ Passeport non trouvé.\n\nVérifiez et réessayez:',
    nl: '❌ Paspoort niet gevonden.\n\nControleer en probeer opnieuw:'
  },
  data_refreshed: {
    ar: '✅ تم تحديث بياناتك بنجاح!',
    en: '✅ Your data has been refreshed!',
    fr: '✅ Vos données ont été actualisées!',
    nl: '✅ Uw gegevens zijn bijgewerkt!'
  },
  switch_pilgrim: {
    ar: '🔀 تم تسجيل الخروج.\n\nأدخل <b>رقم جواز السفر</b> للحاج الآخر:',
    en: '🔀 Logged out.\n\nEnter the <b>passport number</b> of the other pilgrim:',
    fr: '🔀 Déconnecté.\n\nEntrez le <b>numéro de passeport</b> de l\'autre pèlerin:',
    nl: '🔀 Uitgelogd.\n\nVoer het <b>paspoortnummer</b> van de andere pelgrim in:'
  },

  // --- أزرار القائمة الرئيسية ---
  menu_title: {
    ar: '📋 القائمة الرئيسية — اختر ما تريد:',
    en: '📋 Main Menu — Choose an option:',
    fr: '📋 Menu Principal:',
    nl: '📋 Hoofdmenu:'
  },
  btn_flight:      { ar: '✈️ رحلتي',      en: '✈️ My Flight',     fr: '✈️ Mon Vol',       nl: '✈️ Mijn Vlucht' },
  btn_hotel:       { ar: '🏨 فندقي',       en: '🏨 My Hotel',      fr: '🏨 Mon Hôtel',     nl: '🏨 Mijn Hotel' },
  btn_package:     { ar: '📦 باقتي',       en: '📦 My Package',    fr: '📦 Mon Forfait',   nl: '📦 Mijn Pakket' },
  btn_transport:   { ar: '🚌 مواصلاتي',    en: '🚌 My Transport',  fr: '🚌 Mon Transport', nl: '🚌 Mijn Vervoer' },
  btn_emergency:   { ar: '🚨 طوارئ',       en: '🚨 Emergency',     fr: '🚨 Urgences',      nl: '🚨 Noodgevallen' },
  btn_contact:     { ar: '📞 تواصل معنا',  en: '📞 Contact Us',    fr: '📞 Nous Contacter', nl: '📞 Contact' },
  btn_faq:         { ar: '❓ أسئلة عن الحج', en: '❓ Hajj FAQ',     fr: '❓ FAQ Hajj',       nl: '❓ Hadj FAQ' },
  btn_change_lang: { ar: '🔄 تغيير اللغة', en: '🔄 Change Language', fr: '🔄 Changer la langue', nl: '🔄 Taal wijzigen' },
  btn_refresh:     { ar: '🔃 تحديث البيانات', en: '🔃 Refresh Data', fr: '🔃 Actualiser',   nl: '🔃 Vernieuwen' },
  btn_switch:      { ar: '🔀 تبديل الحاج', en: '🔀 Switch Pilgrim', fr: '🔀 Changer de pèlerin', nl: '🔀 Wissel pelgrim' },
  btn_back:        { ar: '🔙 القائمة الرئيسية', en: '🔙 Main Menu', fr: '🔙 Menu Principal', nl: '🔙 Hoofdmenu' },
  btn_whatsapp:    { ar: '💬 فتح واتساب',  en: '💬 Open WhatsApp', fr: '💬 Ouvrir WhatsApp', nl: '💬 WhatsApp openen' },
  btn_my_data:     { ar: '📋 بياناتي',    en: '📋 My Data',       fr: '📋 Mes Données',  nl: '📋 Mijn Gegevens' },
  btn_add_phone:   { ar: '📱 إضافة رقم الهاتف', en: '📱 Add Phone Number', fr: '📱 Ajouter téléphone', nl: '📱 Telefoonnummer toevoegen' },
  btn_add_room:    { ar: '🏨 إدخال رقم الغرفة', en: '🏨 Enter Room Number', fr: '🏨 Numéro de chambre', nl: '🏨 Kamernummer invoeren' },
  btn_add_photo:   { ar: '📸 إرسال صورة الجواز', en: '📸 Send Passport Photo', fr: '📸 Photo du passeport', nl: '📸 Pasfoto sturen' },

  // --- تسميات الفندق ---
  lbl_acc:     { ar: 'السكن',    en: 'Accommodation', fr: 'Hébergement', nl: 'Accommodatie' },
  lbl_hotel:   { ar: 'الفندق',   en: 'Hotel',         fr: 'Hôtel',       nl: 'Hotel' },
  lbl_checkin: { ar: 'الدخول',   en: 'Check-in',      fr: 'Arrivée',     nl: 'Inchecken' },
  lbl_checkout:{ ar: 'الخروج',   en: 'Check-out',     fr: 'Départ',      nl: 'Uitchecken' },
  lbl_map:     { ar: '📍 موقع ', en: '📍 Map: ',      fr: '📍 Carte: ',  nl: '📍 Kaart: ' },
  lbl_shifting:{ ar: 'مكة (تحويل)', en: 'Makkah (Shifting)', fr: 'La Mecque (Transfert)', nl: 'Mekka (Verplaatsing)' },

  // --- الطيران والباقة والمرشد ---
  lbl_transit: { ar: 'ترانزيت', en: 'Transit', fr: 'Transit', nl: 'Transit' },
  lbl_leg:     { ar: 'الرحلة',  en: 'Flight',  fr: 'Vol',     nl: 'Vlucht' },
  lbl_direct:  { ar: 'مباشرة',  en: 'Direct',  fr: 'Direct',  nl: 'Direct' },
  lbl_pkg_name:    { ar: 'اسم الباقة',    en: 'Package Name',   fr: 'Nom du forfait', nl: 'Pakketnaam' },
  lbl_guide:       { ar: 'المرشد',        en: 'Tour Guide',     fr: 'Guide',          nl: 'Gids' },
  lbl_guide_phone: { ar: 'هاتف المرشد',   en: 'Guide Phone',    fr: 'Tél. du guide',  nl: 'Gids telefoon' },

  // --- العد التنازلي ---
  countdown_before: {
    ar: '⏳ باقي على رحلتك <b>{days} يوم</b>',
    en: '⏳ <b>{days} days</b> until your trip',
    fr: '⏳ <b>{days} jours</b> avant votre voyage',
    nl: '⏳ Nog <b>{days} dagen</b> tot uw reis'
  },
  countdown_during: {
    ar: '🕋 أنت الآن في <b>اليوم {day}</b> من رحلتك',
    en: '🕋 You are on <b>day {day}</b> of your trip',
    fr: '🕋 Vous êtes au <b>jour {day}</b> de votre voyage',
    nl: '🕋 U bent op <b>dag {day}</b> van uw reis'
  },
  countdown_after: {
    ar: '🤲 تقبّل الله حجك وأعاده عليك بالخير',
    en: '🤲 May Allah accept your Hajj',
    fr: '🤲 Qu\'Allah accepte votre Hajj',
    nl: '🤲 Moge Allah uw Hadj accepteren'
  },

  // --- جمع البيانات ---
  phone_prompt: {
    ar: '📱 أدخل <b>رقم هاتفك السعودي</b> (الشريحة المحلية):\n\nمثال: 0512345678',
    en: '📱 Enter your <b>Saudi phone number</b> (local SIM):\n\nExample: 0512345678',
    fr: '📱 Entrez votre <b>numéro saoudien</b> (SIM locale):\n\nExemple: 0512345678',
    nl: '📱 Voer uw <b>Saudisch telefoonnummer</b> in (lokale SIM):\n\nVoorbeeld: 0512345678'
  },
  phone_invalid: {
    ar: '⚠️ رقم غير صحيح. أدخل رقم سعودي يبدأ بـ 05 (10 أرقام):',
    en: '⚠️ Invalid number. Enter a Saudi number starting with 05 (10 digits):',
    fr: '⚠️ Numéro invalide. Entrez un numéro saoudien commençant par 05 (10 chiffres):',
    nl: '⚠️ Ongeldig nummer. Voer een Saudisch nummer in beginnend met 05 (10 cijfers):'
  },
  phone_saved: {
    ar: '✅ تم حفظ رقمك: <b>{phone}</b>',
    en: '✅ Phone saved: <b>{phone}</b>',
    fr: '✅ Numéro enregistré: <b>{phone}</b>',
    nl: '✅ Nummer opgeslagen: <b>{phone}</b>'
  },
  room_select_hotel: {
    ar: '🏨 اختر الفندق لإدخال رقم غرفتك:',
    en: '🏨 Select hotel to enter room number:',
    fr: '🏨 Choisissez l\'hôtel pour entrer le numéro:',
    nl: '🏨 Kies hotel om kamernummer in te voeren:'
  },
  room_prompt: {
    ar: '🏨 أرسل <b>رقم غرفتك</b>:',
    en: '🏨 Send your <b>room number</b>:',
    fr: '🏨 Envoyez votre <b>numéro de chambre</b>:',
    nl: '🏨 Stuur uw <b>kamernummer</b>:'
  },
  room_invalid: {
    ar: '⚠️ رقم غرفة غير صحيح. أعد الإرسال:',
    en: '⚠️ Invalid room number. Try again:',
    fr: '⚠️ Numéro invalide. Réessayez:',
    nl: '⚠️ Ongeldig kamernummer. Probeer opnieuw:'
  },
  room_saved: {
    ar: '✅ تم حفظ رقم غرفتك: <b>{room}</b> (فندق {hotel})',
    en: '✅ Room saved: <b>{room}</b> (Hotel {hotel})',
    fr: '✅ Chambre enregistrée: <b>{room}</b> (Hôtel {hotel})',
    nl: '✅ Kamer opgeslagen: <b>{room}</b> (Hotel {hotel})'
  },
  photo_prompt: {
    ar: '📸 أرسل <b>صورة جواز سفرك</b> (صفحة البيانات):\n\n📷 التقط صورة واضحة وأرسلها هنا',
    en: '📸 Send a <b>photo of your passport</b> (data page):\n\n📷 Take a clear photo and send it here',
    fr: '📸 Envoyez une <b>photo de votre passeport</b> (page des données):\n\n📷 Prenez une photo claire',
    nl: '📸 Stuur een <b>foto van uw paspoort</b> (gegevenspagina):\n\n📷 Maak een duidelijke foto'
  },
  photo_saved: {
    ar: '✅ تم حفظ صورة الجواز بنجاح!',
    en: '✅ Passport photo saved successfully!',
    fr: '✅ Photo du passeport enregistrée!',
    nl: '✅ Pasfoto succesvol opgeslagen!'
  },
  photo_error: {
    ar: '❌ حدث خطأ أثناء حفظ الصورة. حاول مرة أخرى لاحقاً.',
    en: '❌ Error saving photo. Please try again later.',
    fr: '❌ Erreur lors de l\'enregistrement. Réessayez plus tard.',
    nl: '❌ Fout bij opslaan. Probeer het later opnieuw.'
  },
  photo_not_received: {
    ar: '📸 أرسل <b>صورة</b> وليس نصاً.\n\nالتقط صورة لجواز سفرك وأرسلها:',
    en: '📸 Send a <b>photo</b>, not text.\n\nTake a photo of your passport and send it:',
    fr: '📸 Envoyez une <b>photo</b>, pas du texte.\n\nPrenez une photo de votre passeport:',
    nl: '📸 Stuur een <b>foto</b>, geen tekst.\n\nMaak een foto van uw paspoort:'
  },
  notif_ask_phone: {
    ar: '📱 هل لديك شريحة سعودية؟ سجّل رقمك لنتمكن من التواصل معك:',
    en: '📱 Do you have a Saudi SIM? Register your number so we can reach you:',
    fr: '📱 Avez-vous une SIM saoudienne? Enregistrez votre numéro:',
    nl: '📱 Heeft u een Saudische SIM? Registreer uw nummer:'
  }
};

/**
 * جلب نص مترجم بالمفتاح واللغة
 * يدعم placeholders: T_('countdown_before', 'ar', {days: 5})
 */
function T_(key, lang, params) {
  var entry = I18N[key];
  if (!entry) return key;
  var text = entry[lang] || entry['en'] || key;
  if (params) {
    for (var k in params) {
      text = text.replace('{' + k + '}', params[k]);
    }
  }
  return text;
}
