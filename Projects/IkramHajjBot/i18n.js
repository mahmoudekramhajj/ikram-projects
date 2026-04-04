// ============================================
// i18n — ملف الترجمة المركزي (ar/en/fr/de/it/es)
// ============================================

var I18N = {

  // --- عام ---
  start_prompt: {
    ar: 'أرسل /start للبدء',
    en: 'Send /start to begin',
    fr: 'Envoyez /start pour commencer',
    de: 'Senden Sie /start um zu beginnen',
    it: 'Invia /start per iniziare',
    es: 'Envía /start para comenzar'
  },
  start_prompt_new: {
    ar: 'أرسل /start للبدء من جديد',
    en: 'Send /start to start over',
    fr: 'Envoyez /start pour recommencer',
    de: 'Senden Sie /start um neu zu beginnen',
    it: 'Invia /start per ricominciare',
    es: 'Envía /start para reiniciar'
  },
  data_not_found: {
    ar: '❌ لم يتم العثور على بياناتك',
    en: '❌ Data not found',
    fr: '❌ Données non trouvées',
    de: '❌ Daten nicht gefunden',
    it: '❌ Dati non trovati',
    es: '❌ Datos no encontrados'
  },
  faq_wip: {
    ar: '⚙️ قسم الأسئلة العامة قيد التطوير...',
    en: '⚙️ FAQ section is under development...',
    fr: '⚙️ Section FAQ en cours de développement...',
    de: '⚙️ FAQ-Bereich in Entwicklung...',
    it: '⚙️ Sezione FAQ in fase di sviluppo...',
    es: '⚙️ Sección de preguntas frecuentes en desarrollo...'
  },

  // --- المصادقة ---
  passport_prompt: {
    ar: '🔐 للوصول لبياناتك الشخصية:\n\nأرسل <b>رقم جواز السفر</b> الخاص بك',
    en: '🔐 To access your personal data:\n\nSend your <b>passport number</b>',
    fr: '🔐 Pour accéder à vos données:\n\nEnvoyez votre <b>numéro de passeport</b>',
    de: '🔐 Um auf Ihre Daten zuzugreifen:\n\nSenden Sie Ihre <b>Reisepassnummer</b>',
    it: '🔐 Per accedere ai tuoi dati:\n\nInvia il tuo <b>numero di passaporto</b>',
    es: '🔐 Para acceder a tus datos:\n\nEnvía tu <b>número de pasaporte</b>'
  },
  passport_invalid: {
    ar: '⚠️ رقم الجواز غير صحيح. تأكد وأعد الإرسال:',
    en: '⚠️ Invalid passport number. Please try again:',
    fr: '⚠️ Numéro de passeport invalide. Réessayez:',
    de: '⚠️ Ungültige Passnummer. Bitte versuchen Sie es erneut:',
    it: '⚠️ Numero di passaporto non valido. Riprova:',
    es: '⚠️ Número de pasaporte inválido. Inténtalo de nuevo:'
  },
  passport_not_found: {
    ar: '❌ رقم الجواز غير موجود في النظام.\n\nتأكد من الرقم وأعد الإرسال:',
    en: '❌ Passport not found in our system.\n\nPlease check and try again:',
    fr: '❌ Passeport non trouvé.\n\nVérifiez et réessayez:',
    de: '❌ Pass nicht in unserem System gefunden.\n\nBitte überprüfen und erneut versuchen:',
    it: '❌ Passaporto non trovato nel sistema.\n\nControlla e riprova:',
    es: '❌ Pasaporte no encontrado en el sistema.\n\nVerifica e intenta de nuevo:'
  },
  data_refreshed: {
    ar: '✅ تم تحديث بياناتك بنجاح!',
    en: '✅ Your data has been refreshed!',
    fr: '✅ Vos données ont été actualisées!',
    de: '✅ Ihre Daten wurden aktualisiert!',
    it: '✅ I tuoi dati sono stati aggiornati!',
    es: '✅ ¡Tus datos han sido actualizados!'
  },
  switch_pilgrim: {
    ar: '🔀 تم تسجيل الخروج.\n\nأدخل <b>رقم جواز السفر</b> للحاج الآخر:',
    en: '🔀 Logged out.\n\nEnter the <b>passport number</b> of the other pilgrim:',
    fr: '🔀 Déconnecté.\n\nEntrez le <b>numéro de passeport</b> de l\'autre pèlerin:',
    de: '🔀 Abgemeldet.\n\nGeben Sie die <b>Passnummer</b> des anderen Pilgers ein:',
    it: '🔀 Disconnesso.\n\nInserisci il <b>numero di passaporto</b> dell\'altro pellegrino:',
    es: '🔀 Sesión cerrada.\n\nIngresa el <b>número de pasaporte</b> del otro peregrino:'
  },

  // --- أزرار القائمة الرئيسية ---
  menu_title: {
    ar: '📋 القائمة الرئيسية — اختر ما تريد:',
    en: '📋 Main Menu — Choose an option:',
    fr: '📋 Menu Principal:',
    de: '📋 Hauptmenü — Wählen Sie eine Option:',
    it: '📋 Menu Principale — Scegli un\'opzione:',
    es: '📋 Menú Principal — Elige una opción:'
  },
  btn_flight:      { ar: '✈️ رحلتي',      en: '✈️ My Flight',     fr: '✈️ Mon Vol',       de: '✈️ Mein Flug',      it: '✈️ Il mio Volo',    es: '✈️ Mi Vuelo' },
  btn_hotel:       { ar: '🏨 فندقي',       en: '🏨 My Hotel',      fr: '🏨 Mon Hôtel',     de: '🏨 Mein Hotel',     it: '🏨 Il mio Hotel',   es: '🏨 Mi Hotel' },
  btn_package:     { ar: '📦 باقتي',       en: '📦 My Package',    fr: '📦 Mon Forfait',   de: '📦 Mein Paket',     it: '📦 Il mio Pacchetto', es: '📦 Mi Paquete' },
  btn_transport:   { ar: '🚌 مواصلاتي',    en: '🚌 My Transport',  fr: '🚌 Mon Transport', de: '🚌 Mein Transport', it: '🚌 Il mio Trasporto', es: '🚌 Mi Transporte' },
  btn_emergency:   { ar: '🚨 طوارئ',       en: '🚨 Emergency',     fr: '🚨 Urgences',      de: '🚨 Notfall',        it: '🚨 Emergenza',      es: '🚨 Emergencia' },
  btn_contact:     { ar: '📞 تواصل معنا',  en: '📞 Contact Us',    fr: '📞 Nous Contacter', de: '📞 Kontakt',       it: '📞 Contattaci',     es: '📞 Contáctanos' },
  btn_faq:         { ar: '❓ أسئلة عن الحج', en: '❓ Hajj FAQ',     fr: '❓ FAQ Hajj',       de: '❓ Hajj FAQ',       it: '❓ FAQ Hajj',        es: '❓ FAQ Hajj' },
  btn_change_lang: { ar: '🔄 تغيير اللغة', en: '🔄 Change Language', fr: '🔄 Changer la langue', de: '🔄 Sprache ändern', it: '🔄 Cambia lingua', es: '🔄 Cambiar idioma' },
  btn_refresh:     { ar: '🔃 تحديث البيانات', en: '🔃 Refresh Data', fr: '🔃 Actualiser',   de: '🔃 Daten aktualisieren', it: '🔃 Aggiorna dati', es: '🔃 Actualizar datos' },
  btn_switch:      { ar: '🔀 تبديل الحاج', en: '🔀 Switch Pilgrim', fr: '🔀 Changer de pèlerin', de: '🔀 Pilger wechseln', it: '🔀 Cambia pellegrino', es: '🔀 Cambiar peregrino' },
  btn_back:        { ar: '🔙 القائمة الرئيسية', en: '🔙 Main Menu', fr: '🔙 Menu Principal', de: '🔙 Hauptmenü', it: '🔙 Menu Principale', es: '🔙 Menú Principal' },
  btn_whatsapp:    { ar: '💬 فتح واتساب',  en: '💬 Open WhatsApp', fr: '💬 Ouvrir WhatsApp', de: '💬 WhatsApp öffnen', it: '💬 Apri WhatsApp', es: '💬 Abrir WhatsApp' },
  btn_my_data:     { ar: '📋 بياناتي',    en: '📋 My Data',       fr: '📋 Mes Données',  de: '📋 Meine Daten',    it: '📋 I miei Dati',    es: '📋 Mis Datos' },
  btn_add_phone:   { ar: '📱 إضافة رقم الهاتف', en: '📱 Add Phone Number', fr: '📱 Ajouter téléphone', de: '📱 Telefonnummer hinzufügen', it: '📱 Aggiungi telefono', es: '📱 Agregar teléfono' },
  btn_add_room:    { ar: '🏨 إدخال رقم الغرفة', en: '🏨 Enter Room Number', fr: '🏨 Numéro de chambre', de: '🏨 Zimmernummer eingeben', it: '🏨 Inserisci numero stanza', es: '🏨 Ingresar número de habitación' },
  btn_add_photo:   { ar: '📸 إرسال صورة الجواز', en: '📸 Send Passport Photo', fr: '📸 Photo du passeport', de: '📸 Passfoto senden', it: '📸 Invia foto passaporto', es: '📸 Enviar foto del pasaporte' },

  // --- تأكيد الوصول ---
  btn_confirm_arrival: { ar: '✅ تأكيد الوصول', en: '✅ Confirm Arrival', fr: '✅ Confirmer l\'arrivée', de: '✅ Ankunft bestätigen', it: '✅ Conferma arrivo', es: '✅ Confirmar llegada' },
  arrival_confirmed: {
    ar: '✅ <b>تم تسجيل وصولك بنجاح!</b>\n\n🕋 مرحباً بك في المملكة العربية السعودية\n🤲 حجاً مبروراً وذنباً مغفوراً',
    en: '✅ <b>Your arrival has been confirmed!</b>\n\n🕋 Welcome to Saudi Arabia\n🤲 May your Hajj be blessed',
    fr: '✅ <b>Votre arrivée a été confirmée!</b>\n\n🕋 Bienvenue en Arabie Saoudite\n🤲 Hajj béni',
    de: '✅ <b>Ihre Ankunft wurde bestätigt!</b>\n\n🕋 Willkommen in Saudi-Arabien\n🤲 Möge Ihre Hajj gesegnet sein',
    it: '✅ <b>Il tuo arrivo è stato confermato!</b>\n\n🕋 Benvenuto in Arabia Saudita\n🤲 Che il tuo Hajj sia benedetto',
    es: '✅ <b>¡Tu llegada ha sido confirmada!</b>\n\n🕋 Bienvenido a Arabia Saudita\n🤲 Que tu Hajj sea bendecido'
  },
  arrival_already: {
    ar: 'ℹ️ وصولك مسجّل مسبقاً بتاريخ <b>{time}</b>',
    en: 'ℹ️ Your arrival was already confirmed at <b>{time}</b>',
    fr: 'ℹ️ Votre arrivée a déjà été confirmée le <b>{time}</b>',
    de: 'ℹ️ Ihre Ankunft wurde bereits bestätigt am <b>{time}</b>',
    it: 'ℹ️ Il tuo arrivo è già stato confermato il <b>{time}</b>',
    es: 'ℹ️ Tu llegada ya fue confirmada el <b>{time}</b>'
  },
  arrival_not_today: {
    ar: '⚠️ تاريخ وصولك <b>{date}</b> — هذا الزر متاح قبل وبعد يوم الوصول فقط',
    en: '⚠️ Your arrival date is <b>{date}</b> — this button is only available around your arrival date',
    fr: '⚠️ Votre date d\'arrivée est le <b>{date}</b> — ce bouton n\'est disponible qu\'autour de cette date',
    de: '⚠️ Ihr Ankunftsdatum ist <b>{date}</b> — diese Schaltfläche ist nur um Ihr Ankunftsdatum verfügbar',
    it: '⚠️ La tua data di arrivo è <b>{date}</b> — questo pulsante è disponibile solo intorno alla data di arrivo',
    es: '⚠️ Tu fecha de llegada es <b>{date}</b> — este botón solo está disponible alrededor de la fecha de llegada'
  },

  // --- QR Code ---
  btn_my_qr: { ar: '📱 كود QR', en: '📱 QR Code', fr: '📱 Code QR', de: '📱 QR-Code', it: '📱 Codice QR', es: '📱 Código QR' },
  qr_caption: {
    ar: '📱 كود QR الخاص بك — أظهره للموظف عند الحاجة',
    en: '📱 Your QR Code — Show it to staff when needed',
    fr: '📱 Votre code QR — Montrez-le au personnel si nécessaire',
    de: '📱 Ihr QR-Code — Zeigen Sie ihn dem Personal bei Bedarf',
    it: '📱 Il tuo codice QR — Mostralo al personale quando necessario',
    es: '📱 Tu código QR — Muéstralo al personal cuando sea necesario'
  },

  // --- تسميات الفندق ---
  lbl_acc:     { ar: 'السكن',    en: 'Accommodation', fr: 'Hébergement', de: 'Unterkunft',   it: 'Alloggio',       es: 'Alojamiento' },
  lbl_hotel:   { ar: 'الفندق',   en: 'Hotel',         fr: 'Hôtel',       de: 'Hotel',        it: 'Hotel',          es: 'Hotel' },
  lbl_checkin: { ar: 'الدخول',   en: 'Check-in',      fr: 'Arrivée',     de: 'Check-in',     it: 'Check-in',       es: 'Check-in' },
  lbl_checkout:{ ar: 'الخروج',   en: 'Check-out',     fr: 'Départ',      de: 'Check-out',    it: 'Check-out',      es: 'Check-out' },
  lbl_map:     { ar: '📍 موقع ', en: '📍 Map: ',      fr: '📍 Carte: ',  de: '📍 Karte: ',   it: '📍 Mappa: ',     es: '📍 Mapa: ' },
  lbl_shifting:{ ar: 'مكة (تحويل)', en: 'Makkah (Shifting)', fr: 'La Mecque (Transfert)', de: 'Mekka (Transfer)', it: 'Mecca (Trasferimento)', es: 'La Meca (Traslado)' },

  // --- الطيران والباقة والمرشد ---
  lbl_transit: { ar: 'ترانزيت', en: 'Transit', fr: 'Transit', de: 'Transit', it: 'Transito', es: 'Tránsito' },
  lbl_leg:     { ar: 'الرحلة',  en: 'Flight',  fr: 'Vol',     de: 'Flug',    it: 'Volo',     es: 'Vuelo' },
  lbl_direct:  { ar: 'مباشرة',  en: 'Direct',  fr: 'Direct',  de: 'Direkt',  it: 'Diretto',  es: 'Directo' },
  lbl_pkg_name:    { ar: 'اسم الباقة',    en: 'Package Name',   fr: 'Nom du forfait', de: 'Paketname',       it: 'Nome pacchetto',    es: 'Nombre del paquete' },
  lbl_guide:       { ar: 'المرشد',        en: 'Tour Guide',     fr: 'Guide',          de: 'Reiseleiter',     it: 'Guida turistica',   es: 'Guía turístico' },
  lbl_guide_phone: { ar: 'هاتف المرشد',   en: 'Guide Phone',    fr: 'Tél. du guide',  de: 'Reiseleiter Tel.', it: 'Tel. guida',       es: 'Tel. del guía' },

  // --- العد التنازلي ---
  countdown_before: {
    ar: '⏳ باقي على رحلتك <b>{days} يوم</b>',
    en: '⏳ <b>{days} days</b> until your trip',
    fr: '⏳ <b>{days} jours</b> avant votre voyage',
    de: '⏳ Noch <b>{days} Tage</b> bis zu Ihrer Reise',
    it: '⏳ <b>{days} giorni</b> al tuo viaggio',
    es: '⏳ <b>{days} días</b> para tu viaje'
  },
  countdown_during: {
    ar: '🕋 أنت الآن في <b>اليوم {day}</b> من رحلتك',
    en: '🕋 You are on <b>day {day}</b> of your trip',
    fr: '🕋 Vous êtes au <b>jour {day}</b> de votre voyage',
    de: '🕋 Sie sind am <b>Tag {day}</b> Ihrer Reise',
    it: '🕋 Sei al <b>giorno {day}</b> del tuo viaggio',
    es: '🕋 Estás en el <b>día {day}</b> de tu viaje'
  },
  countdown_after: {
    ar: '🤲 تقبّل الله حجك وأعاده عليك بالخير',
    en: '🤲 May Allah accept your Hajj',
    fr: '🤲 Qu\'Allah accepte votre Hajj',
    de: '🤲 Möge Allah Ihre Hajj annehmen',
    it: '🤲 Che Allah accetti il tuo Hajj',
    es: '🤲 Que Allah acepte tu Hajj'
  },

  // --- جمع البيانات ---
  phone_prompt: {
    ar: '📱 أدخل <b>رقم هاتفك السعودي</b> (الشريحة المحلية):\n\nمثال: 0512345678',
    en: '📱 Enter your <b>Saudi phone number</b> (local SIM):\n\nExample: 0512345678',
    fr: '📱 Entrez votre <b>numéro saoudien</b> (SIM locale):\n\nExemple: 0512345678',
    de: '📱 Geben Sie Ihre <b>saudische Telefonnummer</b> ein (lokale SIM):\n\nBeispiel: 0512345678',
    it: '📱 Inserisci il tuo <b>numero saudita</b> (SIM locale):\n\nEsempio: 0512345678',
    es: '📱 Ingresa tu <b>número saudí</b> (SIM local):\n\nEjemplo: 0512345678'
  },
  phone_invalid: {
    ar: '⚠️ رقم غير صحيح. أدخل رقم سعودي يبدأ بـ 05 (10 أرقام):',
    en: '⚠️ Invalid number. Enter a Saudi number starting with 05 (10 digits):',
    fr: '⚠️ Numéro invalide. Entrez un numéro saoudien commençant par 05 (10 chiffres):',
    de: '⚠️ Ungültige Nummer. Geben Sie eine saudische Nummer ein, die mit 05 beginnt (10 Ziffern):',
    it: '⚠️ Numero non valido. Inserisci un numero saudita che inizia con 05 (10 cifre):',
    es: '⚠️ Número inválido. Ingresa un número saudí que comience con 05 (10 dígitos):'
  },
  phone_saved: {
    ar: '✅ تم حفظ رقمك: <b>{phone}</b>',
    en: '✅ Phone saved: <b>{phone}</b>',
    fr: '✅ Numéro enregistré: <b>{phone}</b>',
    de: '✅ Nummer gespeichert: <b>{phone}</b>',
    it: '✅ Numero salvato: <b>{phone}</b>',
    es: '✅ Número guardado: <b>{phone}</b>'
  },
  room_select_hotel: {
    ar: '🏨 اختر الفندق لإدخال رقم غرفتك:',
    en: '🏨 Select hotel to enter room number:',
    fr: '🏨 Choisissez l\'hôtel pour entrer le numéro:',
    de: '🏨 Wählen Sie das Hotel um die Zimmernummer einzugeben:',
    it: '🏨 Seleziona l\'hotel per inserire il numero della stanza:',
    es: '🏨 Selecciona el hotel para ingresar el número de habitación:'
  },
  room_prompt: {
    ar: '🏨 أرسل <b>رقم غرفتك</b>:',
    en: '🏨 Send your <b>room number</b>:',
    fr: '🏨 Envoyez votre <b>numéro de chambre</b>:',
    de: '🏨 Senden Sie Ihre <b>Zimmernummer</b>:',
    it: '🏨 Invia il tuo <b>numero di stanza</b>:',
    es: '🏨 Envía tu <b>número de habitación</b>:'
  },
  room_invalid: {
    ar: '⚠️ رقم غرفة غير صحيح. أعد الإرسال:',
    en: '⚠️ Invalid room number. Try again:',
    fr: '⚠️ Numéro invalide. Réessayez:',
    de: '⚠️ Ungültige Zimmernummer. Versuchen Sie es erneut:',
    it: '⚠️ Numero stanza non valido. Riprova:',
    es: '⚠️ Número de habitación inválido. Inténtalo de nuevo:'
  },
  room_saved: {
    ar: '✅ تم حفظ رقم غرفتك: <b>{room}</b> (فندق {hotel})',
    en: '✅ Room saved: <b>{room}</b> (Hotel {hotel})',
    fr: '✅ Chambre enregistrée: <b>{room}</b> (Hôtel {hotel})',
    de: '✅ Zimmer gespeichert: <b>{room}</b> (Hotel {hotel})',
    it: '✅ Stanza salvata: <b>{room}</b> (Hotel {hotel})',
    es: '✅ Habitación guardada: <b>{room}</b> (Hotel {hotel})'
  },
  photo_prompt: {
    ar: '📸 أرسل <b>صورة جواز سفرك</b> (صفحة البيانات):\n\n📷 التقط صورة واضحة وأرسلها هنا',
    en: '📸 Send a <b>photo of your passport</b> (data page):\n\n📷 Take a clear photo and send it here',
    fr: '📸 Envoyez une <b>photo de votre passeport</b> (page des données):\n\n📷 Prenez une photo claire',
    de: '📸 Senden Sie ein <b>Foto Ihres Reisepasses</b> (Datenseite):\n\n📷 Machen Sie ein klares Foto',
    it: '📸 Invia una <b>foto del tuo passaporto</b> (pagina dati):\n\n📷 Scatta una foto chiara',
    es: '📸 Envía una <b>foto de tu pasaporte</b> (página de datos):\n\n📷 Toma una foto clara'
  },
  photo_saved: {
    ar: '✅ تم حفظ صورة الجواز بنجاح!',
    en: '✅ Passport photo saved successfully!',
    fr: '✅ Photo du passeport enregistrée!',
    de: '✅ Passfoto erfolgreich gespeichert!',
    it: '✅ Foto del passaporto salvata con successo!',
    es: '✅ ¡Foto del pasaporte guardada exitosamente!'
  },
  photo_error: {
    ar: '❌ حدث خطأ أثناء حفظ الصورة. حاول مرة أخرى لاحقاً.',
    en: '❌ Error saving photo. Please try again later.',
    fr: '❌ Erreur lors de l\'enregistrement. Réessayez plus tard.',
    de: '❌ Fehler beim Speichern. Bitte versuchen Sie es später erneut.',
    it: '❌ Errore nel salvataggio. Riprova più tardi.',
    es: '❌ Error al guardar la foto. Inténtalo más tarde.'
  },
  photo_not_received: {
    ar: '📸 أرسل <b>صورة</b> وليس نصاً.\n\nالتقط صورة لجواز سفرك وأرسلها:',
    en: '📸 Send a <b>photo</b>, not text.\n\nTake a photo of your passport and send it:',
    fr: '📸 Envoyez une <b>photo</b>, pas du texte.\n\nPrenez une photo de votre passeport:',
    de: '📸 Senden Sie ein <b>Foto</b>, keinen Text.\n\nFotografieren Sie Ihren Reisepass:',
    it: '📸 Invia una <b>foto</b>, non testo.\n\nScatta una foto del passaporto:',
    es: '📸 Envía una <b>foto</b>, no texto.\n\nToma una foto de tu pasaporte:'
  },
  notif_ask_phone: {
    ar: '📱 هل لديك شريحة سعودية؟ سجّل رقمك لنتمكن من التواصل معك:',
    en: '📱 Do you have a Saudi SIM? Register your number so we can reach you:',
    fr: '📱 Avez-vous une SIM saoudienne? Enregistrez votre numéro:',
    de: '📱 Haben Sie eine saudische SIM? Registrieren Sie Ihre Nummer:',
    it: '📱 Hai una SIM saudita? Registra il tuo numero:',
    es: '📱 ¿Tienes una SIM saudí? Registra tu número:'
  },

  // --- الإعلانات ---
  btn_announcements: { ar: '📢 الإعلانات', en: '📢 Announcements', fr: '📢 Annonces', de: '📢 Ankündigungen', it: '📢 Annunci', es: '📢 Anuncios' },
  ann_title: {
    ar: '📢 <b>إعلانات الإدارة</b>\n━━━━━━━━━━━━━━',
    en: '📢 <b>Announcements</b>\n━━━━━━━━━━━━━━',
    fr: '📢 <b>Annonces</b>\n━━━━━━━━━━━━━━',
    de: '📢 <b>Ankündigungen</b>\n━━━━━━━━━━━━━━',
    it: '📢 <b>Annunci</b>\n━━━━━━━━━━━━━━',
    es: '📢 <b>Anuncios</b>\n━━━━━━━━━━━━━━'
  },
  ann_empty: {
    ar: 'لا توجد إعلانات حالياً',
    en: 'No announcements yet',
    fr: 'Pas d\'annonces pour le moment',
    de: 'Noch keine Ankündigungen',
    it: 'Nessun annuncio al momento',
    es: 'No hay anuncios por el momento'
  },
  ann_urgent: { ar: '🚨 عاجل', en: '🚨 Urgent', fr: '🚨 Urgent', de: '🚨 Dringend', it: '🚨 Urgente', es: '🚨 Urgente' },
  ann_from_admin: {
    ar: '📢 <b>رسالة من إدارة إكرام الضيف</b>\n━━━━━━━━━━━━━━\n\n',
    en: '📢 <b>Message from Ikram Aldyf Management</b>\n━━━━━━━━━━━━━━\n\n',
    fr: '📢 <b>Message de la direction Ikram Aldyf</b>\n━━━━━━━━━━━━━━\n\n',
    de: '📢 <b>Nachricht von der Ikram Aldyf Leitung</b>\n━━━━━━━━━━━━━━\n\n',
    it: '📢 <b>Messaggio dalla direzione Ikram Aldyf</b>\n━━━━━━━━━━━━━━\n\n',
    es: '📢 <b>Mensaje de la dirección Ikram Aldyf</b>\n━━━━━━━━━━━━━━\n\n'
  },
  ann_file_attached: { ar: '📎 ملف مرفق', en: '📎 File attached', fr: '📎 Fichier joint', de: '📎 Datei angehängt', it: '📎 File allegato', es: '📎 Archivo adjunto' },

  // --- أوامر المدير ---
  admin_not_authorized: {
    ar: '⛔ غير مصرّح لك باستخدام هذا الأمر',
    en: '⛔ You are not authorized to use this command',
    fr: '⛔ Vous n\'êtes pas autorisé',
    de: '⛔ Sie sind nicht autorisiert',
    it: '⛔ Non sei autorizzato',
    es: '⛔ No estás autorizado'
  },
  admin_broadcast_start: {
    ar: '📢 <b>إرسال رسالة جديدة</b>\n\nاكتب <b>الرسالة بالعربية</b>:',
    en: '📢 <b>New Broadcast</b>\n\nType the <b>Arabic message</b>:'
  },
  admin_broadcast_next_lang: {
    ar: '✅ تم حفظ النص.\n\nالآن اكتب <b>الترجمة {lang}</b> أو اكتب <b>تخطي</b>:',
    en: '✅ Saved.\n\nNow type the <b>{lang} translation</b> or type <b>skip</b>:'
  },
  admin_broadcast_media: {
    ar: '📎 أرسل <b>صورة</b> أو <b>ملف PDF</b>، أو اكتب <b>تخطي</b>:',
    en: '📎 Send a <b>photo</b> or <b>PDF file</b>, or type <b>skip</b>:'
  },
  admin_broadcast_target: {
    ar: '🎯 اختر <b>الاستهداف</b>:',
    en: '🎯 Choose <b>target</b>:'
  },
  admin_broadcast_target_value: {
    ar: '🎯 اكتب <b>القيمة</b> ({target}):',
    en: '🎯 Enter the <b>value</b> ({target}):'
  },
  admin_broadcast_confirm: {
    ar: '📋 <b>ملخص الرسالة</b>\n━━━━━━━━━━━━━━\n\n📝 النص: {preview}\n🎯 الاستهداف: {target}\n📎 مرفق: {media}\n\nهل تريد الإرسال؟',
    en: '📋 <b>Message Summary</b>\n━━━━━━━━━━━━━━\n\n📝 Text: {preview}\n🎯 Target: {target}\n📎 Attachment: {media}\n\nSend now?'
  },
  admin_broadcast_sent: {
    ar: '✅ تم إرسال الرسالة!\n\n📊 وصل: <b>{sent}</b> | فشل: <b>{failed}</b>',
    en: '✅ Message sent!\n\n📊 Delivered: <b>{sent}</b> | Failed: <b>{failed}</b>'
  },
  admin_broadcast_cancelled: {
    ar: '❌ تم إلغاء الرسالة',
    en: '❌ Broadcast cancelled'
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
