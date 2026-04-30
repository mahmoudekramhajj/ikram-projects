// ============================================
// إعدادات البوت والثوابت
// ============================================
var BOT_TOKEN = '8694589281:AAHvT-anZgLDk6s5YO8WStP7Y2zhq6-BDIE';
var TELEGRAM_API = 'https://api.telegram.org/bot' + BOT_TOKEN;
var SHEET_ID = '1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s';
var PILGRIM_DATA_SHEET = 'Pilgrim Data'; // شيت البيانات المجمّعة من الحجاج
var DRIVE_FOLDER_ID = ''; // TODO: أضف ID مجلد Google Drive لحفظ صور الجوازات
var VISA_FOLDER_ID = '1KaiLQhyeFhD1xBXj3Ya_HP08aVgyUkFT'; // مجلد صور التأشيرات (PDF باسم رقم الجواز)
var CACHE_TTL = 300; // مدة الكاش بالثواني (5 دقائق)

// ============================================
// أسماء الشيتات المصدرية (بدل رحلة الحاج المدمج)
// ============================================
var PERSONAL_SHEET = 'Presonal Details';
var PACKAGES_SHEET = 'الباقات';
var FLIGHTS_SHEET = 'الطيران';
var B2C_SHEET = 'B2C';

// ============================================
// إعدادات إخطار تغييرات الطيران (V3)
// ============================================
var FLIGHT_NOTIFY_ENABLED = false;  // kill switch — متوقف (آمن)
var FLIGHT_NOTIFY_DRY_RUN = true;   // وضع التجربة — يكتب في AO بدون إرسال فعلي
var FLIGHT_NOTIFY_MAX_PER_RUN = 50; // حد أقصى 50 إخطار في التشغيل الواحد
var FLIGHT_CHANGES_SHEET_ID = '1gTqnBSovdc0BXFZ0XSf8TxvYSYwdvC8qLNxTzV1lhS8';
var FLIGHT_CHANGES_SHEET_NAME = 'تغييرات الطيران V3';

// ============================================
// فهارس أعمدة Presonal Details
// ============================================
var PD = {
  SEQ: 0, GROUP: 1, TYPE: 2, CATEGORY: 3, GENDER: 4,
  PASSPORT: 5, PASSPORT_EXP: 6, PASSPORT_ISSUE: 7,
  FIRST_NAME_AR: 8, LAST_NAME_AR: 9,
  FIRST_NAME_EN: 10, LAST_NAME_EN: 11,
  DOB: 12, EMAIL: 13, PHONE: 14, GUIDE_NAME: 15,
  COUNTRY_RESIDENCE: 16, NATIONALITY: 17,
  PACKAGE_NO: 18, PACKAGE_NAME: 19,
  FLIGHT_TYPE: 20, CONTRACT_NAME: 21,
  VISA_STATUS: 22, IN_KSA: 23,
  TICKET_NO: 24, TICKET_URL: 25,
  INVOICE_NO: 26, CAMP: 27
};

// ============================================
// فهارس أعمدة الباقات
// ============================================
var PKG = {
  NO: 0, NUSK_NO: 1, NAME_AR: 2, CATEGORY: 3, IKRAM_NO: 4,
  PRICE: 5, DATE_START: 6, DATE_END: 7, NO_DAYS: 8,
  CITY_START: 9, NO_PILGRIM: 10,
  H1_CITY: 11, H1_NAME_AR: 12, H1_NAME_EN: 13, H1_CHECKIN: 14, H1_CHECKOUT: 15,
  H2_CITY: 26, H2_NAME_AR: 27, H2_NAME_EN: 28, H2_CHECKIN: 29, H2_CHECKOUT: 30,
  H3_CITY: 41, H3_NAME_AR: 42, H3_NAME_EN: 43, H3_CHECKIN: 44, H3_CHECKOUT: 45,
  NAME_EN: 60, TRANSPORT: 65
};

// ============================================
// فهارس أعمدة الطيران (B2B)
// ============================================
var FLT = {
  NO: 0, PNR: 1, SUPPLIER: 2, STATUS: 3, COUNTRY: 4,
  CITY: 5, AIRLINE: 6, PAX: 7,
  GO1_FLIGHT: 21, GO1_TAKEOFF_DATE: 22, GO1_TAKEOFF_TIME: 23,
  GO1_FROM: 24, GO1_TO: 25, GO1_LAND_DATE: 26, GO1_LAND_TIME: 27,
  GO2_FLIGHT: 28, GO2_TAKEOFF_DATE: 29, GO2_TAKEOFF_TIME: 30,
  GO2_FROM: 31, GO2_TO: 32, GO2_LAND_DATE: 33, GO2_LAND_TIME: 34,
  RET1_FLIGHT: 35, RET1_TAKEOFF_DATE: 36, RET1_TAKEOFF_TIME: 37,
  RET1_FROM: 38, RET1_TO: 39, RET1_LAND_DATE: 40, RET1_LAND_TIME: 41,
  RET2_FLIGHT: 42, RET2_TAKEOFF_DATE: 43, RET2_TAKEOFF_TIME: 44,
  RET2_FROM: 45, RET2_TO: 46, RET2_LAND_DATE: 47, RET2_LAND_TIME: 48,
  PKG1_NO: 50, PKG2_NO: 52, PKG3_NO: 54, PKG4_NO: 56, PKG5_NO: 58,
  PKG6_NO: 60, PKG7_NO: 62, PKG8_NO: 64, PKG9_NO: 66, PKG10_NO: 68,
  CONTRACT_NAME: 90  // CM — اسم العقد (نفس PD.CONTRACT_NAME)
};

// ============================================
// فهارس أعمدة B2C (نفس Presonal Details + طيران)
// ============================================
var B2CI = {
  PASSPORT: 5,
  ARR1_FLIGHT: 32, ARR1_DATE: 33, ARR1_TIME: 34, ARR1_FROM: 35, ARR1_TO: 36, ARR1_LAND_DATE: 37, ARR1_LAND_TIME: 38,
  ARR2_FLIGHT: 39, ARR2_DATE: 40, ARR2_TIME: 41, ARR2_FROM: 42, ARR2_TO: 43, ARR2_LAND_DATE: 44, ARR2_LAND_TIME: 45,
  RET1_FLIGHT: 46, RET1_DATE: 47, RET1_TIME: 48, RET1_FROM: 49, RET1_TO: 50, RET1_LAND_DATE: 51, RET1_LAND_TIME: 52,
  RET2_FLIGHT: 53, RET2_DATE: 54, RET2_TIME: 55, RET2_FROM: 56, RET2_TO: 57, RET2_LAND_DATE: 58, RET2_LAND_TIME: 59,
  PNR: 60,
  // DEP0 = قطعة الذهاب الإضافية (الأبعد عن السعودية) — للحالات 3 قطع
  // أعمدة BS-BY في الشيت (71-77 → 0-indexed: 70-76)
  DEP0_FLIGHT: 70, DEP0_DATE: 71, DEP0_TIME: 72, DEP0_FROM: 73, DEP0_TO: 74, DEP0_LAND_DATE: 75, DEP0_LAND_TIME: 76,
  // RET3 = قطعة العودة الإضافية (الأبعد عن السعودية) — للحالات 3 قطع
  // أعمدة BZ-CF في الشيت (78-84 → 0-indexed: 77-83)
  RET3_FLIGHT: 77, RET3_DATE: 78, RET3_TIME: 79, RET3_FROM: 80, RET3_TO: 81, RET3_LAND_DATE: 82, RET3_LAND_TIME: 83
};

// ============================================
// أعمدة الاستقبال في Pilgrim Data (بعد الأعمدة الحالية الـ 8)
// ============================================
var PDATA_RECEPTION_STATUS = 8;  // العمود I
var PDATA_RECEPTION_TIME = 9;    // العمود J
var PDATA_RECEPTION_STAFF = 10;  // العمود K

// ============================================
// المدراء — Chat IDs المصرّح لهم باستخدام /broadcast و /stats
// ============================================
var ADMIN_IDS = ['8566760392']; // إكرام الضيف

// ============================================
// أسماء الشيتات الجديدة
// ============================================
var ADMIN_MESSAGES_SHEET = 'AdminMessages';
var DELIVERY_LOG_SHEET = 'DeliveryLog';
var BOT_ACTIVITY_SHEET = 'BotActivity';

// خريطة أعمدة AdminMessages
var AM = {
  ID: 0, TITLE: 1, MSG_AR: 2, MSG_EN: 3, MSG_FR: 4, MSG_DE: 5, MSG_IT: 6, MSG_ES: 7,
  IMAGE_URL: 8, FILE_URL: 9, FILE_NAME: 10, TARGET: 11, TARGET_VALUE: 12,
  PRIORITY: 13, STATUS: 14, SENT_AT: 15, SENT_COUNT: 16, FAIL_COUNT: 17, CREATED_BY: 18
};

// خريطة اللغة → عمود الرسالة في AdminMessages
var LANG_TO_COL = { ar: AM.MSG_AR, en: AM.MSG_EN, fr: AM.MSG_FR, de: AM.MSG_DE, it: AM.MSG_IT, es: AM.MSG_ES };

// ============================================
// خريطة رموز المطارات (IATA) → اسم المطار + المدينة (عربي + إنجليزي)
// ============================================
var AIRPORTS = {
  // السعودية
  'JED': { ar: 'مطار الملك عبدالعزيز الدولي',           en: 'King Abdulaziz Intl',           cityAr: 'جدة',            cityEn: 'Jeddah' },
  'MED': { ar: 'مطار الأمير محمد بن عبدالعزيز الدولي',  en: 'Prince Mohammad Bin Abdulaziz', cityAr: 'المدينة المنورة', cityEn: 'Madinah' },
  'RUH': { ar: 'مطار الملك خالد الدولي',                en: 'King Khalid Intl',              cityAr: 'الرياض',          cityEn: 'Riyadh' },
  'DMM': { ar: 'مطار الملك فهد الدولي',                 en: 'King Fahd Intl',                cityAr: 'الدمام',          cityEn: 'Dammam' },
  // الخليج والشرق الأوسط
  'DXB': { ar: 'مطار دبي الدولي',                       en: 'Dubai Intl',                    cityAr: 'دبي',             cityEn: 'Dubai' },
  'AUH': { ar: 'مطار زايد الدولي',                      en: 'Zayed Intl',                    cityAr: 'أبو ظبي',         cityEn: 'Abu Dhabi' },
  'DOH': { ar: 'مطار حمد الدولي',                       en: 'Hamad Intl',                    cityAr: 'الدوحة',          cityEn: 'Doha' },
  'BAH': { ar: 'مطار البحرين الدولي',                   en: 'Bahrain Intl',                  cityAr: 'المنامة',         cityEn: 'Manama' },
  'KWI': { ar: 'مطار الكويت الدولي',                    en: 'Kuwait Intl',                   cityAr: 'الكويت',          cityEn: 'Kuwait' },
  'AMM': { ar: 'مطار الملكة علياء الدولي',              en: 'Queen Alia Intl',               cityAr: 'عمّان',           cityEn: 'Amman' },
  'BEY': { ar: 'مطار رفيق الحريري الدولي',              en: 'Rafic Hariri Intl',             cityAr: 'بيروت',           cityEn: 'Beirut' },
  'BGW': { ar: 'مطار بغداد الدولي',                     en: 'Baghdad Intl',                  cityAr: 'بغداد',           cityEn: 'Baghdad' },
  // أفريقيا
  'CAI': { ar: 'مطار القاهرة الدولي',                   en: 'Cairo Intl',                    cityAr: 'القاهرة',         cityEn: 'Cairo' },
  'HBE': { ar: 'مطار برج العرب الدولي',                 en: 'Borg El Arab Intl',             cityAr: 'الإسكندرية',      cityEn: 'Alexandria' },
  'TUN': { ar: 'مطار تونس قرطاج',                       en: 'Tunis–Carthage',                cityAr: 'تونس',            cityEn: 'Tunis' },
  'CMN': { ar: 'مطار محمد الخامس الدولي',               en: 'Mohammed V Intl',               cityAr: 'الدار البيضاء',   cityEn: 'Casablanca' },
  'RAK': { ar: 'مطار مراكش المنارة',                    en: 'Marrakesh Menara',              cityAr: 'مراكش',           cityEn: 'Marrakesh' },
  'ALG': { ar: 'مطار هواري بومدين الدولي',              en: 'Houari Boumediene',             cityAr: 'الجزائر',         cityEn: 'Algiers' },
  'ADD': { ar: 'مطار أديس أبابا بولي الدولي',           en: 'Addis Ababa Bole Intl',         cityAr: 'أديس أبابا',       cityEn: 'Addis Ababa' },
  'JNB': { ar: 'مطار أو. آر. تامبو الدولي',             en: 'O. R. Tambo Intl',              cityAr: 'جوهانسبرغ',       cityEn: 'Johannesburg' },
  'DKR': { ar: 'مطار بليز دياني الدولي',                en: 'Blaise Diagne Intl',            cityAr: 'داكار',           cityEn: 'Dakar' },
  'LOS': { ar: 'مطار مرتلا محمد الدولي',                en: 'Murtala Muhammed Intl',         cityAr: 'لاغوس',           cityEn: 'Lagos' },
  'BJL': { ar: 'مطار بانجول الدولي',                    en: 'Banjul Intl',                   cityAr: 'بانجول',          cityEn: 'Banjul' },
  // أوروبا — فرنسا
  'CDG': { ar: 'مطار شارل ديغول',                       en: 'Charles de Gaulle',             cityAr: 'باريس',           cityEn: 'Paris' },
  'ORY': { ar: 'مطار أورلي',                            en: 'Orly',                          cityAr: 'باريس',           cityEn: 'Paris' },
  'MRS': { ar: 'مطار مرسيليا بروفانس',                  en: 'Marseille Provence',            cityAr: 'مرسيليا',         cityEn: 'Marseille' },
  'LYS': { ar: 'مطار ليون سانت إكزوبيري',               en: 'Lyon-Saint Exupéry',            cityAr: 'ليون',            cityEn: 'Lyon' },
  // بريطانيا
  'LHR': { ar: 'مطار لندن هيثرو',                       en: 'London Heathrow',               cityAr: 'لندن',            cityEn: 'London' },
  'LGW': { ar: 'مطار لندن جاتويك',                      en: 'London Gatwick',                cityAr: 'لندن',            cityEn: 'London' },
  'STN': { ar: 'مطار لندن ستانستيد',                    en: 'London Stansted',               cityAr: 'لندن',            cityEn: 'London' },
  'LTN': { ar: 'مطار لندن لوتون',                       en: 'London Luton',                  cityAr: 'لندن',            cityEn: 'London' },
  'MAN': { ar: 'مطار مانشستر',                          en: 'Manchester',                    cityAr: 'مانشستر',         cityEn: 'Manchester' },
  'BHX': { ar: 'مطار برمنغهام',                         en: 'Birmingham',                    cityAr: 'برمنغهام',        cityEn: 'Birmingham' },
  'EDI': { ar: 'مطار إدنبرة',                           en: 'Edinburgh',                     cityAr: 'إدنبرة',          cityEn: 'Edinburgh' },
  // بقية أوروبا
  'BRU': { ar: 'مطار بروكسل الوطني',                    en: 'Brussels Airport',              cityAr: 'بروكسل',          cityEn: 'Brussels' },
  'AMS': { ar: 'مطار سخيبول',                           en: 'Schiphol',                      cityAr: 'أمستردام',        cityEn: 'Amsterdam' },
  'FRA': { ar: 'مطار فرانكفورت',                        en: 'Frankfurt',                     cityAr: 'فرانكفورت',       cityEn: 'Frankfurt' },
  'MUC': { ar: 'مطار ميونخ',                            en: 'Munich',                        cityAr: 'ميونخ',           cityEn: 'Munich' },
  'BER': { ar: 'مطار برلين براندنبورغ',                 en: 'Berlin Brandenburg',            cityAr: 'برلين',           cityEn: 'Berlin' },
  'DUS': { ar: 'مطار دوسلدورف',                         en: 'Düsseldorf',                    cityAr: 'دوسلدورف',        cityEn: 'Düsseldorf' },
  'CGN': { ar: 'مطار كولونيا/بون',                      en: 'Cologne Bonn',                  cityAr: 'كولونيا',         cityEn: 'Cologne' },
  'FCO': { ar: 'مطار روما فيوميتشينو',                  en: 'Rome Fiumicino',                cityAr: 'روما',            cityEn: 'Rome' },
  'MXP': { ar: 'مطار ميلانو مالبينسا',                  en: 'Milan Malpensa',                cityAr: 'ميلانو',          cityEn: 'Milan' },
  'BCN': { ar: 'مطار برشلونة–البرات',                   en: 'Barcelona–El Prat',             cityAr: 'برشلونة',         cityEn: 'Barcelona' },
  'MAD': { ar: 'مطار مدريد باراخاس',                    en: 'Madrid Barajas',                cityAr: 'مدريد',           cityEn: 'Madrid' },
  'VIE': { ar: 'مطار فيينا الدولي',                     en: 'Vienna Intl',                   cityAr: 'فيينا',           cityEn: 'Vienna' },
  'ZRH': { ar: 'مطار زيورخ',                            en: 'Zurich',                        cityAr: 'زيورخ',           cityEn: 'Zurich' },
  'GVA': { ar: 'مطار جنيف',                             en: 'Geneva',                        cityAr: 'جنيف',            cityEn: 'Geneva' },
  'ATH': { ar: 'مطار أثينا الدولي',                     en: 'Athens Intl',                   cityAr: 'أثينا',           cityEn: 'Athens' },
  'ARN': { ar: 'مطار ستوكهولم أرلاندا',                 en: 'Stockholm Arlanda',             cityAr: 'ستوكهولم',        cityEn: 'Stockholm' },
  'CPH': { ar: 'مطار كوبنهاغن كاستراب',                 en: 'Copenhagen Kastrup',            cityAr: 'كوبنهاغن',        cityEn: 'Copenhagen' },
  'OSL': { ar: 'مطار أوسلو',                            en: 'Oslo',                          cityAr: 'أوسلو',           cityEn: 'Oslo' },
  'HEL': { ar: 'مطار هلسنكي',                           en: 'Helsinki',                      cityAr: 'هلسنكي',          cityEn: 'Helsinki' },
  // تركيا
  'IST': { ar: 'مطار إسطنبول',                          en: 'Istanbul Airport',              cityAr: 'إسطنبول',         cityEn: 'Istanbul' },
  'SAW': { ar: 'مطار صبيحة كوكتشن',                     en: 'Sabiha Gökçen',                 cityAr: 'إسطنبول',         cityEn: 'Istanbul' },
  'ESB': { ar: 'مطار إسن بوغا',                         en: 'Esenboğa',                      cityAr: 'أنقرة',           cityEn: 'Ankara' },
  'ADB': { ar: 'مطار عدنان مندريس',                     en: 'Adnan Menderes',                cityAr: 'إزمير',           cityEn: 'İzmir' },
  // أمريكا الشمالية
  'JFK': { ar: 'مطار جون إف كينيدي الدولي',             en: 'John F. Kennedy Intl',          cityAr: 'نيويورك',         cityEn: 'New York' },
  'EWR': { ar: 'مطار نيوآرك ليبرتي الدولي',             en: 'Newark Liberty Intl',           cityAr: 'نيويورك',         cityEn: 'New York' },
  'LGA': { ar: 'مطار لاغوارديا',                        en: 'LaGuardia',                     cityAr: 'نيويورك',         cityEn: 'New York' },
  'ORD': { ar: 'مطار أوهير الدولي',                     en: "O'Hare Intl",                   cityAr: 'شيكاغو',          cityEn: 'Chicago' },
  'DTW': { ar: 'مطار ديترويت الحضري',                   en: 'Detroit Metropolitan',          cityAr: 'ديترويت',         cityEn: 'Detroit' },
  'IAD': { ar: 'مطار دالاس الدولي',                     en: 'Dulles Intl',                   cityAr: 'واشنطن',          cityEn: 'Washington' },
  'LAX': { ar: 'مطار لوس أنجلوس الدولي',                en: 'Los Angeles Intl',              cityAr: 'لوس أنجلوس',      cityEn: 'Los Angeles' },
  'YYZ': { ar: 'مطار تورنتو بيرسون الدولي',             en: 'Toronto Pearson Intl',          cityAr: 'تورنتو',          cityEn: 'Toronto' },
  'YUL': { ar: 'مطار مونتريال–ترودو الدولي',            en: 'Montréal–Trudeau Intl',         cityAr: 'مونتريال',        cityEn: 'Montreal' }
};

// ============================================
// مجموعات العمليات — المطارات (Telegram Chat IDs)
// ============================================
var OPS_GROUPS = {
  madinah: '-4849598886',        // مطار المدينة
  jeddah_t1: '-5220583519',      // مطار جدة — صالة 1
  jeddah_north: '-5267173490'    // مطار جدة — الصالة الشمالية
};

// مجموعات العمليات — المواقع
var OPS_LOCATION_GROUPS = {
  makkah: '-4916619724',         // عمليات مكة
  madinah: '-5284394785',        // عمليات المدينة
  mashaaer: '-5268778683'        // عمليات المشاعر
};

// ============================================
// جدول صالات مطار جدة حسب شركة الطيران
// ============================================
var AIRLINE_TERMINAL = {
  'Emirates': 'T1', 'Etihad': 'T1', 'Gulf Air': 'T1',
  'EgyptAir': 'T1', 'Qatar Airways': 'T1', 'Royal Jordanian': 'T1',
  'Saudia': 'T1', 'Turkish Airlines': 'T1', 'Ethiopian Airlines': 'T1',
  'Flyadeal': 'T1',
  'AJet': 'N', 'Wizz Air': 'N', 'Aegean Airlines': 'N',
  'Air Cairo': 'N', 'AnadoluJet': 'N', 'Flydubai': 'N',
  'Pegasus Airlines': 'N', 'WEST ISLE AIR INC.': 'N'
};

// ============================================
// جداول الترجمة — الجنسيات (EN → AR)
// ============================================
var NATIONALITY_AR_MAP = {
  'Turkish': 'تركي', 'Egyptian': 'مصري', 'Jordanian': 'أردني',
  'Palestinian': 'فلسطيني', 'Syrian': 'سوري', 'Lebanese': 'لبناني',
  'Iraqi': 'عراقي', 'Yemeni': 'يمني', 'Sudanese': 'سوداني',
  'Moroccan': 'مغربي', 'Algerian': 'جزائري', 'Tunisian': 'تونسي',
  'Libyan': 'ليبي', 'Emirati': 'إماراتي', 'Kuwaiti': 'كويتي',
  'Bahraini': 'بحريني', 'Qatari': 'قطري', 'Omani': 'عماني',
  'Pakistani': 'باكستاني', 'Indian': 'هندي', 'Indonesian': 'إندونيسي',
  'Malaysian': 'ماليزي', 'Bangladeshi': 'بنغلاديشي',
  'Afghan': 'أفغاني', 'Iranian': 'إيراني', 'Nigerian': 'نيجيري',
  'Somali': 'صومالي', 'Ethiopian': 'إثيوبي',
  'British': 'بريطاني', 'American': 'أمريكي', 'French': 'فرنسي',
  'German': 'ألماني', 'Italian': 'إيطالي', 'Spanish': 'إسباني',
  'Dutch': 'هولندي', 'Belgian': 'بلجيكي', 'Swedish': 'سويدي',
  'Australian': 'أسترالي', 'Canadian': 'كندي', 'Russian': 'روسي',
  'Chinese': 'صيني', 'Japanese': 'ياباني', 'Korean': 'كوري',
  'Saudi': 'سعودي', 'Uzbek': 'أوزبكي', 'Kazakh': 'كازاخي',
  'Tajik': 'طاجيكي', 'Kyrgyz': 'قيرغيزي', 'Turkmen': 'تركماني',
  'Bosnian': 'بوسني', 'Albanian': 'ألباني', 'Kosovar': 'كوسوفي',
  'Senegalese': 'سنغالي', 'Malian': 'مالي', 'Guinean': 'غيني',
  'Cameroonian': 'كاميروني', 'Ghanaian': 'غاني',
  'Sri Lankan': 'سريلانكي', 'Maldivian': 'مالديفي',
  'Filipino': 'فلبيني', 'Thai': 'تايلاندي', 'Bruneian': 'بروناوي'
};

// ============================================
// جداول الترجمة — الدول (EN → AR)
// ============================================
var COUNTRY_AR_MAP = {
  'Turkey': 'تركيا', 'Egypt': 'مصر', 'Jordan': 'الأردن',
  'Palestine': 'فلسطين', 'Syria': 'سوريا', 'Lebanon': 'لبنان',
  'Iraq': 'العراق', 'Yemen': 'اليمن', 'Sudan': 'السودان',
  'Morocco': 'المغرب', 'Algeria': 'الجزائر', 'Tunisia': 'تونس',
  'Libya': 'ليبيا', 'UAE': 'الإمارات', 'Kuwait': 'الكويت',
  'Bahrain': 'البحرين', 'Qatar': 'قطر', 'Oman': 'عمان',
  'Pakistan': 'باكستان', 'India': 'الهند', 'Indonesia': 'إندونيسيا',
  'Malaysia': 'ماليزيا', 'Bangladesh': 'بنغلاديش',
  'Afghanistan': 'أفغانستان', 'Iran': 'إيران', 'Nigeria': 'نيجيريا',
  'Somalia': 'الصومال', 'Ethiopia': 'إثيوبيا',
  'United Kingdom': 'بريطانيا', 'UK': 'بريطانيا',
  'United States': 'أمريكا', 'USA': 'أمريكا',
  'France': 'فرنسا', 'Germany': 'ألمانيا', 'Italy': 'إيطاليا',
  'Spain': 'إسبانيا', 'Netherlands': 'هولندا', 'Belgium': 'بلجيكا',
  'Sweden': 'السويد', 'Australia': 'أستراليا', 'Canada': 'كندا',
  'Russia': 'روسيا', 'China': 'الصين', 'Japan': 'اليابان',
  'Korea': 'كوريا', 'Saudi Arabia': 'السعودية',
  'Uzbekistan': 'أوزبكستان', 'Kazakhstan': 'كازاخستان',
  'Tajikistan': 'طاجيكستان', 'Kyrgyzstan': 'قيرغيزستان',
  'Turkmenistan': 'تركمانستان', 'Bosnia': 'البوسنة',
  'Albania': 'ألبانيا', 'Kosovo': 'كوسوفو',
  'Senegal': 'السنغال', 'Mali': 'مالي', 'Guinea': 'غينيا',
  'Cameroon': 'الكاميرون', 'Ghana': 'غانا',
  'Sri Lanka': 'سريلانكا', 'Maldives': 'المالديف',
  'Philippines': 'الفلبين', 'Thailand': 'تايلاند', 'Brunei': 'بروناي'
};

// ============================================
// جداول الترجمة — شركات الطيران (EN → AR)
// ============================================
var AIRLINE_AR_MAP = {
  'Emirates': 'طيران الإمارات', 'Etihad': 'الاتحاد للطيران',
  'Gulf Air': 'طيران الخليج', 'EgyptAir': 'مصر للطيران',
  'Qatar Airways': 'الخطوط القطرية', 'Royal Jordanian': 'الملكية الأردنية',
  'Saudia': 'الخطوط السعودية', 'Turkish Airlines': 'الخطوط التركية',
  'Ethiopian Airlines': 'الخطوط الإثيوبية', 'Flyadeal': 'طيران أديل',
  'AJet': 'إيه جت', 'Wizz Air': 'ويز إير',
  'Aegean Airlines': 'إيجين للطيران', 'Air Cairo': 'إير كايرو',
  'AnadoluJet': 'أناضولو جت', 'Flydubai': 'فلاي دبي',
  'Pegasus Airlines': 'بيغاسوس للطيران', 'WEST ISLE AIR INC.': 'ويست آيل إير'
};

// ============================================
// CacheService — دوال مساعدة
// ============================================
function getCache_(key) {
  var cache = CacheService.getScriptCache();
  var data = cache.get(key);
  if (data) {
    try { return JSON.parse(data); } catch(e) { return null; }
  }
  return null;
}

function setCache_(key, value, ttl) {
  var cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(value), ttl || CACHE_TTL);
}

function clearPilgrimCache_(passport) {
  var cache = CacheService.getScriptCache();
  var p = String(passport).toUpperCase().trim();
  // مسح كل cache keys المتعلقة بالحاج
  cache.removeAll([
    'pilgrim_' + p,        // findPilgrimByPassport_
    'flightdata_' + p,     // getPilgrimFlightData_ (legs structure)
    'b2cflight_' + p,      // getB2CFlightDetails_ (leg1/leg2)
    'visa_ticket_' + p     // visa info
  ]);
}

function clearSessionCache_(chatId) {
  var cache = CacheService.getScriptCache();
  cache.remove('session_' + String(chatId));
}

function clearTransportCache_(packageId) {
  var cache = CacheService.getScriptCache();
  cache.remove('transport_' + String(packageId).trim());
}

function clearHotelMapCache_(rowData) {
  var cache = CacheService.getScriptCache();
  var names = [
    String(rowData[42] || ''), String(rowData[43] || ''),
    String(rowData[44] || ''), String(rowData[45] || ''),
    String(rowData[46] || ''), String(rowData[47] || '')
  ];
  for (var i = 0; i < names.length; i++) {
    if (names[i] && names[i] !== '-') {
      cache.remove('hmap2_' + names[i]);
    }
  }
}
// deploy-fix-1776306155
