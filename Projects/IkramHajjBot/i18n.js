// ============================================
// i18n — ملف الترجمة المركزي (ar/en/fr/de/it/es)
// ============================================

var I18N = {

  // --- عام ---
  start_prompt: { ar: 'أرسل /start للبدء', en: 'Send /start to begin', fr: 'Envoyez /start pour commencer', de: 'Senden Sie /start um zu beginnen', it: 'Invia /start per iniziare', es: 'Envía /start para comenzar' },
  start_prompt_new: { ar: 'أرسل /start للبدء من جديد', en: 'Send /start to start over', fr: 'Envoyez /start pour recommencer', de: 'Senden Sie /start um neu zu beginnen', it: 'Invia /start per ricominciare', es: 'Envía /start para reiniciar' },
  data_not_found: { ar: '❌ لم يتم العثور على بياناتك', en: '❌ Data not found', fr: '❌ Données non trouvées', de: '❌ Daten nicht gefunden', it: '❌ Dati non trovati', es: '❌ Datos no encontrados' },
  faq_wip: { ar: '⚙️ قسم الأسئلة العامة قيد التطوير...', en: '⚙️ FAQ section is under development...', fr: '⚙️ Section FAQ en cours de développement...', de: '⚙️ FAQ-Bereich in Entwicklung...', it: '⚙️ Sezione FAQ in fase di sviluppo...', es: '⚙️ Sección de preguntas frecuentes en desarrollo...' },

  // --- المصادقة ---
  passport_prompt: { ar: '🔐 للوصول لبياناتك الشخصية:\n\nأرسل <b>رقم جواز السفر</b> الخاص بك', en: '🔐 To access your personal data:\n\nSend your <b>passport number</b>', fr: '🔐 Pour accéder à vos données:\n\nEnvoyez votre <b>numéro de passeport</b>', de: '🔐 Um auf Ihre Daten zuzugreifen:\n\nSenden Sie Ihre <b>Reisepassnummer</b>', it: '🔐 Per accedere ai tuoi dati:\n\nInvia il tuo <b>numero di passaporto</b>', es: '🔐 Para acceder a tus datos:\n\nEnvía tu <b>número de pasaporte</b>' },
  passport_invalid: { ar: '⚠️ رقم الجواز غير صحيح. تأكد وأعد الإرسال:', en: '⚠️ Invalid passport number. Please try again:', fr: '⚠️ Numéro de passeport invalide. Réessayez:', de: '⚠️ Ungültige Passnummer. Bitte versuchen Sie es erneut:', it: '⚠️ Numero di passaporto non valido. Riprova:', es: '⚠️ Número de pasaporte inválido. Inténtalo de nuevo:' },
  passport_not_found: { ar: '❌ رقم الجواز غير موجود في النظام.\n\nتأكد من الرقم وأعد الإرسال:', en: '❌ Passport not found in our system.\n\nPlease check and try again:', fr: '❌ Passeport non trouvé.\n\nVérifiez et réessayez:', de: '❌ Pass nicht in unserem System gefunden.\n\nBitte überprüfen und erneut versuchen:', it: '❌ Passaporto non trovato nel sistema.\n\nControlla e riprova:', es: '❌ Pasaporte no encontrado en el sistema.\n\nVerifica e intenta de nuevo:' },
  data_refreshed: { ar: '✅ تم تحديث بياناتك بنجاح!', en: '✅ Your data has been refreshed!', fr: '✅ Vos données ont été actualisées!', de: '✅ Ihre Daten wurden aktualisiert!', it: '✅ I tuoi dati sono stati aggiornati!', es: '✅ ¡Tus datos han sido actualizados!' },
  switch_pilgrim: { ar: '🔀 تم تسجيل الخروج.\n\nأدخل <b>رقم جواز السفر</b> للحاج الآخر:', en: '🔀 Logged out.\n\nEnter the <b>passport number</b> of the other pilgrim:', fr: '🔀 Déconnecté.\n\nEntrez le <b>numéro de passeport</b> de l\'autre pèlerin:', de: '🔀 Abgemeldet.\n\nGeben Sie die <b>Passnummer</b> des anderen Pilgers ein:', it: '🔀 Disconnesso.\n\nInserisci il <b>numero di passaporto</b> dell\'altro pellegrino:', es: '🔀 Sesión cerrada.\n\nIngresa el <b>número de pasaporte</b> del otro peregrino:' },

  // --- أزرار القائمة الرئيسية ---
  menu_title: { ar: '📋 القائمة الرئيسية — اختر ما تريد:', en: '📋 Main Menu — Choose an option:', fr: '📋 Menu Principal:', de: '📋 Hauptmenü — Wählen Sie eine Option:', it: '📋 Menu Principale — Scegli un\'opzione:', es: '📋 Menú Principal — Elige una opción:' },
  btn_flight:      { ar: '✈️ رحلتي',      en: '✈️ My Flight',     fr: '✈️ Mon Vol',       de: '✈️ Mein Flug',      it: '✈️ Il mio Volo',    es: '✈️ Mi Vuelo' },
  btn_hotel:       { ar: '🏨 فندقي',       en: '🏨 My Hotel',      fr: '🏨 Mon Hôtel',     de: '🏨 Mein Hotel',     it: '🏨 Il mio Hotel',   es: '🏨 Mi Hotel' },
  btn_package:     { ar: '📦 باقتي',       en: '📦 My Package',    fr: '📦 Mon Forfait',   de: '📦 Mein Paket',     it: '📦 Il mio Pacchetto', es: '📦 Mi Paquete' },
  btn_transport:   { ar: '🚌 مواصلاتي',    en: '🚌 My Transport',  fr: '🚌 Mon Transport', de: '🚌 Mein Transport', it: '🚌 Il mio Trasporto', es: '🚌 Mi Transporte' },
  btn_emergency:   { ar: '🚨 طوارئ',       en: '🚨 Emergency',     fr: '🚨 Urgences',      de: '🚨 Notfall',        it: '🚨 Emergenza',      es: '🚨 Emergencia' },
  btn_contact:     { ar: '📞 تواصل معنا',  en: '📞 Contact Us',    fr: '📞 Nous Contacter', de: '📞 Kontakt',       it: '📞 Contattaci',     es: '📞 Contáctanos' },
  btn_change_lang: { ar: '🔄 تغيير اللغة', en: '🔄 Change Language', fr: '🔄 Changer la langue', de: '🔄 Sprache ändern', it: '🔄 Cambia lingua', es: '🔄 Cambiar idioma' },
  btn_refresh:     { ar: '🔃 تحديث البيانات', en: '🔃 Refresh Data', fr: '🔃 Actualiser', de: '🔃 Daten aktualisieren', it: '🔃 Aggiorna dati', es: '🔃 Actualizar datos' },
  btn_switch:      { ar: '🔀 تبديل الحاج', en: '🔀 Switch Pilgrim', fr: '🔀 Changer de pèlerin', de: '🔀 Pilger wechseln', it: '🔀 Cambia pellegrino', es: '🔀 Cambiar peregrino' },
  btn_back:        { ar: '🔙 القائمة الرئيسية', en: '🔙 Main Menu', fr: '🔙 Menu Principal', de: '🔙 Hauptmenü', it: '🔙 Menu Principale', es: '🔙 Menú Principal' },
  btn_whatsapp:    { ar: '💬 فتح واتساب',  en: '💬 Open WhatsApp', fr: '💬 Ouvrir WhatsApp', de: '💬 WhatsApp öffnen', it: '💬 Apri WhatsApp', es: '💬 Abrir WhatsApp' },
  btn_my_data:     { ar: '📋 بياناتي',    en: '📋 My Data',       fr: '📋 Mes Données',  de: '📋 Meine Daten',    it: '📋 I miei Dati',    es: '📋 Mis Datos' },
  btn_add_phone:   { ar: '📱 إضافة رقم الهاتف', en: '📱 Add Phone Number', fr: '📱 Ajouter téléphone', de: '📱 Telefonnummer hinzufügen', it: '📱 Aggiungi telefono', es: '📱 Agregar teléfono' },
  btn_add_room:    { ar: '🏨 إدخال رقم الغرفة', en: '🏨 Enter Room Number', fr: '🏨 Numéro de chambre', de: '🏨 Zimmernummer eingeben', it: '🏨 Inserisci numero stanza', es: '🏨 Ingresar número de habitación' },
  btn_add_photo:   { ar: '📸 إرسال صورة الجواز', en: '📸 Send Passport Photo', fr: '📸 Photo du passeport', de: '📸 Passfoto senden', it: '📸 Invia foto passaporto', es: '📸 Enviar foto del pasaporte' },

  // --- تأكيد الوصول ---
  btn_confirm_arrival: { ar: '✅ تأكيد الوصول', en: '✅ Confirm Arrival', fr: '✅ Confirmer l\'arrivée', de: '✅ Ankunft bestätigen', it: '✅ Conferma arrivo', es: '✅ Confirmar llegada' },
  arrival_confirmed: { ar: '✅ <b>تم تسجيل وصولك بنجاح!</b>\n\n🕋 مرحباً بك في المملكة العربية السعودية\n🤲 حجاً مبروراً وذنباً مغفوراً', en: '✅ <b>Your arrival has been confirmed!</b>\n\n🕋 Welcome to Saudi Arabia\n🤲 May your Hajj be blessed', fr: '✅ <b>Votre arrivée a été confirmée!</b>\n\n🕋 Bienvenue en Arabie Saoudite\n🤲 Hajj béni', de: '✅ <b>Ihre Ankunft wurde bestätigt!</b>\n\n🕋 Willkommen in Saudi-Arabien\n🤲 Möge Ihre Hajj gesegnet sein', it: '✅ <b>Il tuo arrivo è stato confermato!</b>\n\n🕋 Benvenuto in Arabia Saudita\n🤲 Che il tuo Hajj sia benedetto', es: '✅ <b>¡Tu llegada ha sido confirmada!</b>\n\n🕋 Bienvenido a Arabia Saudita\n🤲 Que tu Hajj sea bendecido' },
  arrival_already: { ar: 'ℹ️ وصولك مسجّل مسبقاً بتاريخ <b>{time}</b>', en: 'ℹ️ Your arrival was already confirmed at <b>{time}</b>', fr: 'ℹ️ Votre arrivée a déjà été confirmée le <b>{time}</b>', de: 'ℹ️ Ihre Ankunft wurde bereits bestätigt am <b>{time}</b>', it: 'ℹ️ Il tuo arrivo è già stato confermato il <b>{time}</b>', es: 'ℹ️ Tu llegada ya fue confirmada el <b>{time}</b>' },
  arrival_not_today: { ar: '⚠️ تاريخ وصولك <b>{date}</b> — هذا الزر متاح قبل وبعد يوم الوصول فقط', en: '⚠️ Your arrival date is <b>{date}</b> — this button is only available around your arrival date', fr: '⚠️ Votre date d\'arrivée est le <b>{date}</b> — ce bouton n\'est disponible qu\'autour de cette date', de: '⚠️ Ihr Ankunftsdatum ist <b>{date}</b> — diese Schaltfläche ist nur um Ihr Ankunftsdatum verfügbar', it: '⚠️ La tua data di arrivo è <b>{date}</b> — questo pulsante è disponibile solo intorno alla data di arrivo', es: '⚠️ Tu fecha de llegada es <b>{date}</b> — este botón solo está disponible alrededor de la fecha de llegada' },

  // --- QR Code ---
  btn_my_qr: { ar: '📱 كود QR', en: '📱 QR Code', fr: '📱 Code QR', de: '📱 QR-Code', it: '📱 Codice QR', es: '📱 Código QR' },
  qr_caption: { ar: '📱 كود QR الخاص بك — أظهره للموظف عند الحاجة', en: '📱 Your QR Code — Show it to staff when needed', fr: '📱 Votre code QR — Montrez-le au personnel si nécessaire', de: '📱 Ihr QR-Code — Zeigen Sie ihn dem Personal bei Bedarf', it: '📱 Il tuo codice QR — Mostralo al personale quando necessario', es: '📱 Tu código QR — Muéstralo al personal cuando sea necesario' },

  // --- تسميات الفندق ---
  lbl_acc:      { ar: 'السكن',         en: 'Accommodation', fr: 'Hébergement', de: 'Unterkunft',   it: 'Alloggio',       es: 'Alojamiento' },
  lbl_hotel:    { ar: 'الفندق',        en: 'Hotel',         fr: 'Hôtel',       de: 'Hotel',        it: 'Hotel',          es: 'Hotel' },
  lbl_checkin:  { ar: 'الدخول',        en: 'Check-in',      fr: 'Arrivée',     de: 'Check-in',     it: 'Check-in',       es: 'Check-in' },
  lbl_checkout: { ar: 'الخروج',        en: 'Check-out',     fr: 'Départ',      de: 'Check-out',    it: 'Check-out',      es: 'Check-out' },
  lbl_map:      { ar: '📍 موقع ',      en: '📍 Map: ',      fr: '📍 Carte: ',  de: '📍 Karte: ',   it: '📍 Mappa: ',     es: '📍 Mapa: ' },
  lbl_shifting: { ar: 'مكة (تحويل)',   en: 'Makkah (Shifting)', fr: 'La Mecque (Transfert)', de: 'Mekka (Transfer)', it: 'Mecca (Trasferimento)', es: 'La Meca (Traslado)' },

  // --- الطيران والباقة والمرشد ---
  lbl_transit:      { ar: 'ترانزيت',     en: 'Transit',       fr: 'Transit',     de: 'Transit',      it: 'Transito',       es: 'Tránsito' },
  lbl_leg:          { ar: 'الرحلة',      en: 'Flight',        fr: 'Vol',         de: 'Flug',         it: 'Volo',           es: 'Vuelo' },
  lbl_direct:       { ar: 'مباشرة',      en: 'Direct',        fr: 'Direct',      de: 'Direkt',       it: 'Diretto',        es: 'Directo' },
  lbl_pkg_name:     { ar: 'اسم الباقة',  en: 'Package Name',  fr: 'Nom du forfait', de: 'Paketname', it: 'Nome pacchetto', es: 'Nombre del paquete' },
  lbl_guide:        { ar: 'المرشد',      en: 'Tour Guide',    fr: 'Guide',       de: 'Reiseleiter',  it: 'Guida turistica', es: 'Guía turístico' },
  lbl_guide_phone:  { ar: 'هاتف المرشد', en: 'Guide Phone',   fr: 'Tél. du guide', de: 'Reiseleiter Tel.', it: 'Tel. guida', es: 'Tel. del guía' },
  lbl_under_review: { ar: 'قيد التدقيق', en: 'Under Review',  fr: 'En cours de vérification', de: 'Wird überprüft', it: 'In revisione', es: 'En revisión' },

  // --- التذكرة والتأشيرة ---
  btn_visa_ticket:      { ar: '🎫 التذكرة والتأشيرة', en: '🎫 Ticket & Visa', fr: '🎫 Billet & Visa', de: '🎫 Ticket & Visum', it: '🎫 Biglietto & Visto', es: '🎫 Billete & Visa' },
  lbl_visa_status:      { ar: 'حالة التأشيرة', en: 'Visa Status', fr: 'Statut du visa', de: 'Visumstatus', it: 'Stato del visto', es: 'Estado del visado' },
  lbl_ticket_no:        { ar: 'رقم التذكرة', en: 'Ticket Number', fr: 'N° du billet', de: 'Ticketnummer', it: 'N. biglietto', es: 'N° de billete' },
  btn_download_ticket:  { ar: '📄 تحميل التذكرة', en: '📄 Download Ticket', fr: '📄 Télécharger le billet', de: '📄 Ticket herunterladen', it: '📄 Scarica biglietto', es: '📄 Descargar billete' },
  btn_view_visa:        { ar: '📋 عرض التأشيرة', en: '📋 View Visa', fr: '📋 Voir le visa', de: '📋 Visum anzeigen', it: '📋 Vedi visto', es: '📋 Ver visado' },
  btn_train_ticket:     { ar: '🚆 تذكرة القطار', en: '🚆 Train Ticket', fr: '🚆 Billet de train', de: '🚆 Zugticket', it: '🚆 Biglietto treno', es: '🚆 Billete de tren' },
  train_ticket_unavailable: { ar: '🚆 تذكرة القطار غير متوفرة حالياً — سيتم إضافتها قريباً', en: '🚆 Train ticket not available yet — coming soon', fr: '🚆 Billet de train non disponible — bientôt', de: '🚆 Zugticket noch nicht verfügbar', it: '🚆 Biglietto treno non ancora disponibile', es: '🚆 Billete de tren no disponible aún' },
  lbl_ticket_not_ready: { ar: 'لم تصدر التذكرة بعد', en: 'Ticket not issued yet', fr: 'Billet pas encore émis', de: 'Ticket noch nicht ausgestellt', it: 'Biglietto non ancora emesso', es: 'Billete aún no emitido' },
  lbl_visa_not_ready:   { ar: 'صورة التأشيرة غير متوفرة حالياً', en: 'Visa image not available', fr: 'Image du visa indisponible', de: 'Visumbild nicht verfügbar', it: 'Immagine visto non disponibile', es: 'Imagen del visado no disponible' },

  // --- العد التنازلي ---
  countdown_before: { ar: '⏳ باقي على رحلتك <b>{days} يوم</b>', en: '⏳ <b>{days} days</b> until your trip', fr: '⏳ <b>{days} jours</b> avant votre voyage', de: '⏳ Noch <b>{days} Tage</b> bis zu Ihrer Reise', it: '⏳ <b>{days} giorni</b> al tuo viaggio', es: '⏳ <b>{days} días</b> para tu viaje' },
  countdown_during: { ar: '🕋 أنت الآن في <b>اليوم {day}</b> من رحلتك', en: '🕋 You are on <b>day {day}</b> of your trip', fr: '🕋 Vous êtes au <b>jour {day}</b> de votre voyage', de: '🕋 Sie sind am <b>Tag {day}</b> Ihrer Reise', it: '🕋 Sei al <b>giorno {day}</b> del tuo viaggio', es: '🕋 Estás en el <b>día {day}</b> de tu viaje' },
  countdown_after: { ar: '🤲 تقبّل الله حجك وأعاده عليك بالخير', en: '🤲 May Allah accept your Hajj', fr: '🤲 Qu\'Allah accepte votre Hajj', de: '🤲 Möge Allah Ihre Hajj annehmen', it: '🤲 Che Allah accetti il tuo Hajj', es: '🤲 Que Allah acepte tu Hajj' },

  // --- جمع البيانات ---
  phone_prompt: { ar: '📱 أدخل <b>رقم هاتفك السعودي</b> (الشريحة المحلية):\n\nمثال: 0512345678', en: '📱 Enter your <b>Saudi phone number</b> (local SIM):\n\nExample: 0512345678', fr: '📱 Entrez votre <b>numéro saoudien</b>:\n\nExemple: 0512345678', de: '📱 Geben Sie Ihre <b>saudische Telefonnummer</b> ein:\n\nBeispiel: 0512345678', it: '📱 Inserisci il tuo <b>numero saudita</b>:\n\nEsempio: 0512345678', es: '📱 Ingresa tu <b>número saudí</b>:\n\nEjemplo: 0512345678' },
  phone_invalid: { ar: '⚠️ رقم غير صحيح. أدخل رقم سعودي يبدأ بـ 05 (10 أرقام):', en: '⚠️ Invalid number. Enter a Saudi number starting with 05 (10 digits):', fr: '⚠️ Numéro invalide.', de: '⚠️ Ungültige Nummer.', it: '⚠️ Numero non valido.', es: '⚠️ Número inválido.' },
  phone_saved: { ar: '✅ تم حفظ رقمك: <b>{phone}</b>', en: '✅ Phone saved: <b>{phone}</b>', fr: '✅ Numéro enregistré: <b>{phone}</b>', de: '✅ Nummer gespeichert: <b>{phone}</b>', it: '✅ Numero salvato: <b>{phone}</b>', es: '✅ Número guardado: <b>{phone}</b>' },
  room_select_hotel: { ar: '🏨 اختر الفندق لإدخال رقم غرفتك:', en: '🏨 Select hotel to enter room number:', fr: '🏨 Choisissez l\'hôtel:', de: '🏨 Wählen Sie das Hotel:', it: '🏨 Seleziona l\'hotel:', es: '🏨 Selecciona el hotel:' },
  room_prompt: { ar: '🏨 أرسل <b>رقم غرفتك</b>:', en: '🏨 Send your <b>room number</b>:', fr: '🏨 Envoyez votre <b>numéro de chambre</b>:', de: '🏨 Senden Sie Ihre <b>Zimmernummer</b>:', it: '🏨 Invia il tuo <b>numero di stanza</b>:', es: '🏨 Envía tu <b>número de habitación</b>:' },
  room_invalid: { ar: '⚠️ رقم غرفة غير صحيح. أعد الإرسال:', en: '⚠️ Invalid room number. Try again:', fr: '⚠️ Numéro invalide.', de: '⚠️ Ungültige Zimmernummer.', it: '⚠️ Numero stanza non valido.', es: '⚠️ Número de habitación inválido.' },
  room_saved: { ar: '✅ تم حفظ رقم غرفتك: <b>{room}</b> (فندق {hotel})', en: '✅ Room saved: <b>{room}</b> (Hotel {hotel})', fr: '✅ Chambre enregistrée: <b>{room}</b> (Hôtel {hotel})', de: '✅ Zimmer gespeichert: <b>{room}</b> (Hotel {hotel})', it: '✅ Stanza salvata: <b>{room}</b> (Hotel {hotel})', es: '✅ Habitación guardada: <b>{room}</b> (Hotel {hotel})' },
  photo_prompt: { ar: '📸 أرسل <b>صورة جواز سفرك</b> (صفحة البيانات):\n\n📷 التقط صورة واضحة وأرسلها هنا', en: '📸 Send a <b>photo of your passport</b> (data page):\n\n📷 Take a clear photo and send it here', fr: '📸 Envoyez une <b>photo de votre passeport</b>', de: '📸 Senden Sie ein <b>Foto Ihres Reisepasses</b>', it: '📸 Invia una <b>foto del tuo passaporto</b>', es: '📸 Envía una <b>foto de tu pasaporte</b>' },
  photo_saved: { ar: '✅ تم حفظ صورة الجواز بنجاح!', en: '✅ Passport photo saved successfully!', fr: '✅ Photo du passeport enregistrée!', de: '✅ Passfoto erfolgreich gespeichert!', it: '✅ Foto del passaporto salvata!', es: '✅ ¡Foto del pasaporte guardada!' },
  photo_error: { ar: '❌ حدث خطأ أثناء حفظ الصورة. حاول مرة أخرى لاحقاً.', en: '❌ Error saving photo. Please try again later.', fr: '❌ Erreur. Réessayez plus tard.', de: '❌ Fehler beim Speichern.', it: '❌ Errore nel salvataggio.', es: '❌ Error al guardar la foto.' },
  photo_not_received: { ar: '📸 أرسل <b>صورة</b> وليس نصاً.\n\nالتقط صورة لجواز سفرك وأرسلها:', en: '📸 Send a <b>photo</b>, not text.\n\nTake a photo of your passport and send it:', fr: '📸 Envoyez une <b>photo</b>, pas du texte.', de: '📸 Senden Sie ein <b>Foto</b>, keinen Text.', it: '📸 Invia una <b>foto</b>, non testo.', es: '📸 Envía una <b>foto</b>, no texto.' },
  notif_ask_phone: { ar: '📱 هل لديك شريحة سعودية؟ سجّل رقمك لنتمكن من التواصل معك:', en: '📱 Do you have a Saudi SIM? Register your number so we can reach you:', fr: '📱 Avez-vous une SIM saoudienne?', de: '📱 Haben Sie eine saudische SIM?', it: '📱 Hai una SIM saudita?', es: '📱 ¿Tienes una SIM saudí?' },

  // --- الإعلانات ---
  btn_announcements: { ar: '📢 الإعلانات', en: '📢 Announcements', fr: '📢 Annonces', de: '📢 Ankündigungen', it: '📢 Annunci', es: '📢 Anuncios' },
  ann_title: { ar: '📢 <b>إعلانات الإدارة</b>\n━━━━━━━━━━━━━━', en: '📢 <b>Announcements</b>\n━━━━━━━━━━━━━━', fr: '📢 <b>Annonces</b>\n━━━━━━━━━━━━━━', de: '📢 <b>Ankündigungen</b>\n━━━━━━━━━━━━━━', it: '📢 <b>Annunci</b>\n━━━━━━━━━━━━━━', es: '📢 <b>Anuncios</b>\n━━━━━━━━━━━━━━' },
  ann_empty: { ar: 'لا توجد إعلانات حالياً', en: 'No announcements yet', fr: 'Pas d\'annonces', de: 'Noch keine Ankündigungen', it: 'Nessun annuncio', es: 'No hay anuncios' },
  ann_urgent: { ar: '🚨 عاجل', en: '🚨 Urgent', fr: '🚨 Urgent', de: '🚨 Dringend', it: '🚨 Urgente', es: '🚨 Urgente' },
  ann_from_admin: { ar: '📢 <b>رسالة من إدارة إكرام الضيف</b>\n━━━━━━━━━━━━━━\n\n', en: '📢 <b>Message from Ikram Aldyf Management</b>\n━━━━━━━━━━━━━━\n\n', fr: '📢 <b>Message de la direction</b>\n━━━━━━━━━━━━━━\n\n', de: '📢 <b>Nachricht von der Leitung</b>\n━━━━━━━━━━━━━━\n\n', it: '📢 <b>Messaggio dalla direzione</b>\n━━━━━━━━━━━━━━\n\n', es: '📢 <b>Mensaje de la dirección</b>\n━━━━━━━━━━━━━━\n\n' },
  ann_file_attached: { ar: '📎 ملف مرفق', en: '📎 File attached', fr: '📎 Fichier joint', de: '📎 Datei angehängt', it: '📎 File allegato', es: '📎 Archivo adjunto' },

  // --- نظام البلاغات ---
  btn_report_error: { ar: '⚠️ إبلاغ عن خطأ', en: '⚠️ Report Issue', fr: '⚠️ Signaler', de: '⚠️ Melden', it: '⚠️ Segnala', es: '⚠️ Reportar' },
  btn_my_reports:   { ar: '📋 بلاغاتي', en: '📋 My Reports', fr: '📋 Mes rapports', de: '📋 Meine Berichte', it: '📋 Segnalazioni', es: '📋 Mis reportes' },
  report_select_section: { ar: '📋 <b>الإبلاغ عن خطأ في بياناتك</b>\n\nاختر القسم:', en: '📋 <b>Report an error</b>\n\nSelect the section:', fr: '📋 <b>Signaler une erreur</b>\n\nSection:', de: '📋 <b>Fehler melden</b>\n\nBereich:', it: '📋 <b>Segnala un errore</b>\n\nSezione:', es: '📋 <b>Reportar un error</b>\n\nSección:' },
  report_sec_flight:    { ar: '✈️ الطيران', en: '✈️ Flight', fr: '✈️ Vol', de: '✈️ Flug', it: '✈️ Volo', es: '✈️ Vuelo' },
  report_sec_hotel:     { ar: '🏨 الفندق', en: '🏨 Hotel', fr: '🏨 Hôtel', de: '🏨 Hotel', it: '🏨 Hotel', es: '🏨 Hotel' },
  report_sec_visa:      { ar: '🎫 التأشيرة/التذكرة', en: '🎫 Visa/Ticket', fr: '🎫 Visa', de: '🎫 Visum', it: '🎫 Visto', es: '🎫 Visa' },
  report_sec_package:   { ar: '📦 الباقة', en: '📦 Package', fr: '📦 Forfait', de: '📦 Paket', it: '📦 Pacchetto', es: '📦 Paquete' },
  report_sec_transport: { ar: '🚌 النقل', en: '🚌 Transport', fr: '🚌 Transport', de: '🚌 Transport', it: '🚌 Trasporto', es: '🚌 Transporte' },
  report_sec_personal:  { ar: '📋 بيانات شخصية', en: '📋 Personal Data', fr: '📋 Données personnelles', de: '📋 Persönliche Daten', it: '📋 Dati personali', es: '📋 Datos personales' },
  report_select_issue: { ar: '❗ اختر المشكلة:', en: '❗ Select the issue:', fr: '❗ Sélectionnez:', de: '❗ Wählen Sie:', it: '❗ Seleziona:', es: '❗ Selecciona:' },
  report_flight_date: { ar: 'تاريخ الرحلة خاطئ', en: 'Wrong flight date' },
  report_flight_number: { ar: 'رقم الرحلة خاطئ', en: 'Wrong flight number' },
  report_flight_airline: { ar: 'شركة الطيران خاطئة', en: 'Wrong airline' },
  report_hotel_name: { ar: 'اسم الفندق خاطئ', en: 'Wrong hotel name' },
  report_hotel_dates: { ar: 'تاريخ الدخول/الخروج خاطئ', en: 'Wrong hotel dates' },
  report_hotel_room: { ar: 'رقم الغرفة خاطئ', en: 'Wrong room number' },
  report_visa_status: { ar: 'حالة التأشيرة خاطئة', en: 'Wrong visa status' },
  report_ticket_number: { ar: 'رقم التذكرة خاطئ', en: 'Wrong ticket number' },
  report_ticket_link: { ar: 'رابط التذكرة لا يعمل', en: 'Ticket link not working' },
  report_package_type: { ar: 'نوع الباقة خاطئ', en: 'Wrong package type' },
  report_package_details: { ar: 'تفاصيل الباقة خاطئة', en: 'Wrong package details' },
  report_transport_time: { ar: 'وقت النقل خاطئ', en: 'Wrong transport time' },
  report_transport_point: { ar: 'نقطة التجمع خاطئة', en: 'Wrong meeting point' },
  report_name_wrong: { ar: 'الاسم خاطئ', en: 'Wrong name' },
  report_passport_wrong: { ar: 'رقم الجواز خاطئ', en: 'Wrong passport number' },
  report_nationality_wrong: { ar: 'الجنسية خاطئة', en: 'Wrong nationality' },
  report_other: { ar: '🔹 مشكلة أخرى', en: '🔹 Other issue' },
  report_write_correct: { ar: '✏️ اكتب <b>المعلومة الصحيحة</b>:', en: '✏️ Write the <b>correct information</b>:' },
  report_write_other: { ar: '✏️ اكتب <b>وصف المشكلة</b>:', en: '✏️ Describe the <b>issue</b>:' },
  report_issue_saved: { ar: '✅ تم تسجيل المشكلة.\n\nهل تريد إضافة مشكلة أخرى أو إرسال البلاغ؟', en: '✅ Issue recorded.\n\nAdd another or submit?' },
  btn_add_more: { ar: '➕ إضافة مشكلة أخرى', en: '➕ Add Another Issue' },
  btn_submit_report: { ar: '📤 إرسال البلاغ', en: '📤 Submit Report' },
  btn_cancel_report: { ar: '❌ إلغاء', en: '❌ Cancel' },
  report_submitted: { ar: '✅ <b>تم إرسال بلاغك بنجاح!</b>\n\n📋 رقم البلاغ: <b>{id}</b>\n📊 عدد المشاكل: {count}\n\nسيتم مراجعته وإبلاغك بالنتيجة.', en: '✅ <b>Report submitted!</b>\n\n📋 Ticket: <b>{id}</b>\n📊 Issues: {count}' },
  report_cancelled: { ar: '❌ تم إلغاء البلاغ', en: '❌ Report cancelled' },
  report_resolved_notify: { ar: '✅ <b>تم حل بلاغك #{id}</b>\n━━━━━━━━━━━━━━\n\n{details}\n🤲 شكراً لإبلاغنا', en: '✅ <b>Report #{id} resolved</b>\n\n{details}\n🤲 Thank you' },
  report_rejected_notify: { ar: '❌ <b>بلاغك #{id} تم رفضه</b>\n\nتواصل مع الدعم عبر واتساب:', en: '❌ <b>Report #{id} rejected</b>\n\nContact support via WhatsApp:' },
  report_status_open: { ar: '🔵 مفتوح', en: '🔵 Open' },
  report_status_resolved: { ar: '✅ محلول', en: '✅ Resolved' },
  report_status_rejected: { ar: '❌ مرفوض', en: '❌ Rejected' },
  my_reports_title: { ar: '📋 <b>بلاغاتي</b>\n━━━━━━━━━━━━━━', en: '📋 <b>My Reports</b>\n━━━━━━━━━━━━━━' },
  my_reports_empty: { ar: 'لا توجد بلاغات سابقة', en: 'No previous reports' },
  lbl_admin_response: { ar: 'رد الإدارة', en: 'Admin Response', fr: 'Réponse admin', de: 'Admin-Antwort', it: 'Risposta admin', es: 'Respuesta admin' },
  lbl_reject_reason: { ar: 'سبب الرفض', en: 'Rejection Reason', fr: 'Motif du refus', de: 'Ablehnungsgrund', it: 'Motivo rifiuto', es: 'Motivo de rechazo' },

  // --- أوامر المدير ---
  admin_not_authorized: { ar: '⛔ غير مصرّح لك باستخدام هذا الأمر', en: '⛔ You are not authorized' },
  admin_broadcast_start: { ar: '📢 <b>إرسال رسالة جديدة</b>\n\nاكتب <b>الرسالة بالعربية</b>:', en: '📢 <b>New Broadcast</b>\n\nType the <b>Arabic message</b>:' },
  admin_broadcast_next_lang: { ar: '✅ تم حفظ النص.\n\nالآن اكتب <b>الترجمة {lang}</b> أو اكتب <b>تخطي</b>:', en: '✅ Saved.\n\nNow type the <b>{lang} translation</b> or type <b>skip</b>:' },
  admin_broadcast_media: { ar: '📎 أرسل <b>صورة</b> أو <b>ملف PDF</b>، أو اكتب <b>تخطي</b>:', en: '📎 Send a <b>photo</b> or <b>PDF</b>, or type <b>skip</b>:' },
  admin_broadcast_target: { ar: '🎯 اختر <b>الاستهداف</b>:', en: '🎯 Choose <b>target</b>:' },
  admin_broadcast_target_value: { ar: '🎯 اكتب <b>القيمة</b> ({target}):', en: '🎯 Enter the <b>value</b> ({target}):' },
  admin_broadcast_confirm: { ar: '📋 <b>ملخص الرسالة</b>\n━━━━━━━━━━━━━━\n\n📝 النص: {preview}\n🎯 الاستهداف: {target}\n📎 مرفق: {media}\n\nهل تريد الإرسال؟', en: '📋 <b>Summary</b>\n\n📝 Text: {preview}\n🎯 Target: {target}\n📎 Attachment: {media}\n\nSend now?' },
  admin_broadcast_sent: { ar: '✅ تم إرسال الرسالة!\n\n📊 وصل: <b>{sent}</b> | فشل: <b>{failed}</b>', en: '✅ Message sent!\n\n📊 Delivered: <b>{sent}</b> | Failed: <b>{failed}</b>' },
  admin_broadcast_cancelled: { ar: '❌ تم إلغاء الرسالة', en: '❌ Broadcast cancelled' }
};

function T_(key, lang, params) {
  var entry = I18N[key];
  if (!entry) return key;
  var text = entry[lang] || entry['en'] || key;
  if (params) {
    var isArabic = (lang === 'ar');
    for (var k in params) {
      var val = params[k];
      // في العربية: تحويل الأرقام تلقائياً لعربية-هندية لحل RTL mixing
      if (isArabic && val !== null && val !== undefined && /[0-9]/.test(String(val))) {
        val = toArabicDigits_(val);
      }
      text = text.replace('{' + k + '}', val);
    }
  }
  return text;
}
