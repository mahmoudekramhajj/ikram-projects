# إكرام الضيف — Ekram Aldyf

## قواعد حرجة

- **تحقّق من البيانات الفعلية أولاً:** تحقّق دائماً من البيانات الفعلية (صفوف الشيتات، سجلات الحجاج، تفاصيل البلاغات) قبل اقتراح أي إجراء أو صياغة رد أو إجراء تعديلات. لا تفترض قيم البيانات أبداً.
- **لا عمليات تدميرية بدون إذن:** لا تشغّل عمليات تدميرية (حذف صفوف، إزالة مكررات، الكتابة فوق بيانات) دون إذن صريح من المستخدم. اعرض ما سيتغيّر دائماً قبل التنفيذ.
- **إفصاح إلزامي قبل أي تعديل:** قبل أي تعديل على كود أو بيانات، أجب على ثلاثة أسئلة: 1) أي بيانات ستُقرأ ومن أي شيت؟ 2) أي أعمدة ستُستخدم؟ 3) ما الذي سيتغيّر؟ — ثم انتظر الموافقة الصريحة قبل المتابعة.

---

## قواعد العمل معي
- إذا لم تفهم الفكرة 100%، اطرح سؤالاً واحداً فقط وانتظر إجابتي
- أبلّغني بالتعديلات المقترحة والنتائج المتوقعة قبل التنفيذ
- لا تنفذ أي شيء حتى أعطيك الموافقة الصريحة
- استخدم المهارات (Skills) المتاحة تلقائياً حسب المهمة بدون انتظار إرشاد — اختر المهارة الأنسب وتقمّصها مباشرة
- **عند كتابة "تابع":** اقرأ ملف `SESSION_LOG.md` في جذر المشروع، ثم أكمل العمل من آخر جلسة مسجّلة فيه دون طرح أسئلة
- **قبل تعديل أي تطبيق يقرأ من Google Sheets:** استخدم Agent لتحليل هيكل الشيت الفعلي — اعرض كل عنوان عمود وموقعه وعيّنة من البيانات. لا تبدأ بكتابة الكود حتى يرجع العميل بالتقرير
- **بناء تدريجي وليس شامل:** عند تنفيذ ميزة جديدة: 1) اعرض الأعمدة/البيانات التي ستُستخدم 2) اشرح منطق الفلترة بكلمات واضحة 3) ابنِ الجزء الأول فقط واتركني أختبره 4) بعد تأكيدي فقط، أضف الجزء التالي. لا تبني كل شيء دفعة واحدة
- **تحليل التأثير إلزامي:** قبل أي تعديل على تطبيق، اقرأ `DEPENDENCIES.md`، حدد الشيتات المتأثرة، واعرض قائمة التطبيقات الأخرى التي قد تتأثر بالتغيير. إذا كان عدد التطبيقات المتأثرة > 0، اعرض تحليل التأثير وانتظر موافقتي قبل الكتابة
- **منع فقدان السياق:** في بداية أي مهمة معقدة (أكثر من 3 خطوات)، أنشئ ملف `.claude/SESSION_STATE.md` وحدّثه بعد كل خطوة مهمة. الملف يحتوي: الهدف الحالي، الخطوات المنجزة، الخطوات القادمة، القرارات المهمة. إذا شعرت بفقدان السياق — اقرأ هذا الملف فوراً قبل المتابعة
- **المزامنة الثلاثية إلزامية عند إنهاء كل محادثة:** يجب أن يكون محتوى المشروع متطابقاً في 3 أماكن:
  1. **Apps Script (السحابة):** `clasp push --force && clasp deploy` — كل تعديل منشور ومفعّل
  2. **GitHub:** `git add <project> && git commit && git push` — كل تعديل مدفوع للـ remote
  3. **الملف المحلي:** آخر نسخة محفوظة في `C:\Users\mubar\Ekram-Aldyf\Projects\`
  قبل إنهاء أي محادثة، نفّذ الخطوات الثلاث وتحقق من التطابق. إذا فشلت خطوة، أبلّغني ولا تعلن إنهاء المحادثة

---

## نظرة عامة على المشروع

هذا مشروع عمليات حج (إكرام الضيف). الأنظمة الرئيسية: Google Sheets (بيانات الحجاج، الطيران، الفنادق)، Google Apps Script (الخلفية)، بوت تيليغرام، لوحة GitHub Pages. لا تخلط بين سياقَي B2B/B2C ولا تخلط بين نطاقات المشاريع.

**مشروع واحد متكامل** لإدارة رحلات الحج والعمرة، مقسّم إلى تطبيقات فرعية لسهولة المتابعة.
جميع التطبيقات تستمد بياناتها من **شيت Ikram Abuown** الموحّد.

**المسار:** `C:\Users\mubar\Ekram-Aldyf\Projects\` (Git repo)
**المسار القديم (Google Drive):** `D:\اكرام الضيف\Ekram Aldyf\Projects\` (للقراءة فقط)
**التوقيت:** Asia/Riyadh
**البيئة:** Google Apps Script V8 runtime
**Spreadsheet ID:** `1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s`

---

## البيانات والشيتات

- **تأكيد أعمدة Sheets قبل الكود:** عند العمل مع خرائط أعمدة Google Sheets، أكّد دائماً أحرف/أسماء الأعمدة بدقة مع المستخدم قبل كتابة الكود. لا تفترض أن مواضع الأعمدة متطابقة بين الشيتات (مثل B2B vs B2C، GO1 vs GO2، Flights vs All).

---

## الوصول المباشر للبيانات

سكريبت `SheetData/refresh.js` يستخدم Google Sheets API مع Service Account لتحميل جميع الشيتات كملفات JSON محلية.

```bash
cd "D:\اكرام الضيف\Ekram Aldyf\Projects\SheetData"
node refresh.js
```
الملفات تُحفظ في `SheetData/sheets/` — يمكن قراءتها مباشرة دون الحاجة لنسخ يدوية.

---

## التطبيقات الفرعية

| التطبيق | الوصف |
|---|---|
| **Ikram** | النظام الرئيسي — باقات + فنادق + رحلات v5 (45 deployment، آخرها V 20.2) |
| **IkramAgent** | وكيل ذكي متعدد القنوات (Telegram + WhatsApp) — AR/EN/FR — ⏸️ مؤجّل للعام القادم |
| **IkramHajjBot** | بوت تيليغرام لمتابعة رحلة الحاج (مفعّل) |
| **Pilgrim App** | تطبيق الحاج |
| **Guide App** | تطبيق المرشد السياحي |
| **TourGuide Manager** | إدارة المرشدين (فارغ تقريباً — 3KB) |
| **Reception Airport** | استقبال المطار |
| **Airport Search** | بحث في بيانات المطار |
| **Hotel Management** | إدارة الفنادق |
| **GDS** | إكمال بيانات رحلات B2C من AeroDataBox API |
| **B2C Sync** | مزامنة بيانات B2C |
| **Mina Camp Search** | بحث في مخيمات منى |
| **Sales Operations Report** | تقارير المبيعات |
| **Holiday In Bakkah** | باقات العطل في مكة (Spreadsheet منفصل) |
| **SheetData** | أدوات تصدير البيانات (Node.js + Google Sheets API) |
| **PNR Target Countries Sync** | مزامنة PNR للدول المستهدفة |
| **PilgrimStats** | شاشات إحصائيات الحجاج للشاشات الكبيرة — أفراد/مجاملة/B2B/B2C/مخيمات/دول (GAS Web App) |
| **IkramAdmin** | لوحة تحكم إدارية — إعدادات + إجراءات سريعة + تيليغرام |
| **Report Builder** | منشئ التقارير الذكي (Ikram Report Builder) |
| **Transport Management** | إدارة النقل — 28 عملية، 350 موظف، QR scan |

---

## خريطة تفاعل المشاريع

### الشيتات المركزية ومن يستخدمها

| الشيت | يقرأ منه | يكتب فيه |
|---|---|---|
| **رحلة الحاج** (6500+ صف) | Airport Search, IkramHajjBot, Pilgrim App, Guide App, Hotel Management, Mina Camp Search | Reception Airport, **JourneyMerger** (يكتب الشيت كاملاً) |
| **رحلة الحاج 2** (6500+ صف) | **JourneyMerger** (مصدر بيانات الرحلة) | منصة نسك (مصدر خارجي) |
| **Presonal Details** (6500+ صف) | **JourneyMerger** (مصدر البيانات الشخصية)، Pilgrim App, B2C Sync, TourGuide Manager, **PilgrimStats** | منصة نسك (مصدر خارجي) |
| **الباقات** (69 عمود) | Ikram, IkramAgent, IkramHajjBot, Hotel Management, Pilgrim App, Holiday In Bakkah | Ikram |
| **الطيران** (103 عمود) | Ikram, IkramAgent, GDS, B2C Sync, Sales Report, Holiday In Bakkah | Ikram, GDS, B2C Sync, PNR Sync |
| **الفنادق** | Ikram, Hotel Management, Holiday In Bakkah | — |
| **مخيم مني** | Mina Camp Search | Ikram (Minacamp.js) |
| **Guide Rabih** (المصدر المعتمد للمرشدين) | IkramHajjBot, Hotel Management, Pilgrim App, Guide App, Ikram (Minacamp.js, Roomtype builder.js), SheetData | — (يُحدَّث من نسك) |
| **Tour Guide** (قديم — لم يعد مستخدماً) | — | TourGuide Manager |
| **BotSessions** | IkramHajjBot | IkramHajjBot |

### APIs الخارجية

| API | المشاريع | الاستخدام |
|---|---|---|
| **AeroDataBox** (RapidAPI) | Ikram, Hotel Management, GDS | جلب بيانات الرحلات الجوية |
| **Telegram Bot API** | IkramHajjBot, IkramAgent | بوتات الدردشة |
| **Twilio WhatsApp** | IkramAgent | واتساب |
| **Claude AI** | IkramAgent | الرد على لغات غير مدعومة |
| **Google Drive API** | Airport Search, GDS | تصدير Excel/PDF |

### تصنيف التطبيقات

**تطبيقات ويب (8):**
| التطبيق | المستخدم |
|---|---|
| Ikram | الإدارة — لوحة تحكم الباقات والرحلات |
| Airport Search | موظفي المطار — بحث وتصدير |
| Hotel Management | موظفي الفنادق — توزيع غرف |
| Pilgrim App | الحاج — عرض بياناته |
| Guide App | المرشد — عرض مجموعته |
| Reception Airport | موظفي الاستقبال — تأكيد وصول |
| Mina Camp Search | الموظفين الميدانيين — بحث خيام |
| Sales Report | الإدارة — تقارير مبيعات |
| IkramAdmin | الإدارة — لوحة تحكم إدارية |
| Report Builder | الإدارة — منشئ التقارير |
| Transport Management | الإدارة — إدارة النقل |

**بوتات (2):**
| البوت | الوظيفة |
|---|---|
| IkramHajjBot | بعد الحجز — الحاج يتابع رحلته عبر Telegram |
| IkramAgent | قبل الحجز — العميل يبحث عن باقات (⏸️ مؤجّل) |

**تعمل بالخلفية (4):**
| المشروع | الوظيفة |
|---|---|
| GDS | إكمال بيانات رحلات B2C من API |
| B2C Sync | مزامنة أسماء الحجاج |
| PNR Sync | مزامنة الدول المستهدفة لكل PNR |
| SheetData | تصدير JSON محلي (Node.js) |

---

## هيكل شيت Ikram Abuown

### قائمة الشيتات (52 ورقة)

| الشيت | الأعمدة | الصفوف | الوصف |
|---|---|---|---|
| **الباقات** | 69 (A–BQ) | 60 | بيانات الباقات — تبدأ من الصف 3 |
| **الطيران** | 103 (A–CY) | 191 | بيانات الرحلات — تبدأ من الصف 3 |
| **الفنادق** | 52 (A–AZ) | 993 | قائمة الفنادق — تبدأ من الصف 2 |
| **رحلة الحاج** | 48 (A–AV) | 5965 | بيانات الحجاج الكاملة |
| **رحلة الحاج1** | 41 | 6057 | نسخة بديلة |
| **المستخدمين** | 26 | 1000 | مستخدمي النظام |
| **مخيم مني** | 28 | 5966 | بيانات مخيمات منى |
| **رحلة الحاج 2** | 42 | 6527 | الشكل الجديد من نسك — بيانات الرحلة بدون بيانات شخصية |
| **Presonal Details** | 33 | 6532 | البيانات الشخصية للحجاج (اسم، جواز، إيميل، جوال، مرشد، تأشيرة، تذكرة) |
| **Guide Rabih** | 33 | 6535 | **المصدر المعتمد لأسماء المرشدين** — نفس هيكل Presonal Details + عمود P مملوء بأسماء المرشدين (88 مرشد) |

### شيت الباقات — هيكل الأعمدة الكامل

#### معلومات أساسية (A–K)
| العمود | الفهرس | المحتوى |
|---|---|---|
| A | 0 | No |
| B | 1 | Nusk No. |
| C | 2 | Name Of Pakg (عربي) |
| D | 3 | Category |
| E | 4 | Ikram No |
| F | 5 | سعر الباقة الابتدائي |
| G | 6 | Date Of Start |
| H | 7 | Date Of End |
| I | 8 | No of Date |
| J | 9 | City Of Start |
| K | 10 | No. of Pilgrim |

#### الفندق الأول (L–Y) — هيدر: "معلومات الفندق الاول"
| العمود | الفهرس | المحتوى |
|---|---|---|
| L | 11 | City |
| M | 12 | Name of Hotel (عربي) |
| N | 13 | Name of Hotel English |
| O | 14 | Check-in |
| P | 15 | Check-out |
| Q | 16 | No. Dbl |
| R | 17 | Price of Dbl |
| S | 18 | No. Tri |
| T | 19 | Price of Tri |
| U | 20 | No. Quad |
| V | 21 | Price of Quad |
| W | 22 | No. Rooms |
| X | 23 | No. Beds |
| Y | 24 | (فارغ) |

#### الفندق الثاني (Z–AO) — هيدر: "معلومات الفندق الثاني"
| العمود | الفهرس | المحتوى |
|---|---|---|
| Z | 25 | الفرق |
| AA | 26 | City |
| AB | 27 | Name of Hotel (عربي) |
| AC | 28 | Name of Hotel English |
| AD–AK | 29–36 | Check-in → Price of Quad |
| AL | 37 | No. Rooms |
| AM | 38 | No. Beds |
| AN–AO | 39–40 | (فارغ / No. Beds) |

#### الفندق الثالث (AP–BB) — هيدر: "معلومات الفندق الثالث"
| العمود | الفهرس | المحتوى |
|---|---|---|
| AP | 41 | City |
| AQ | 42 | Name of Hotel (عربي) |
| AR | 43 | Name of Hotel English |
| AS–AZ | 44–51 | Check-in → Price of Quad |
| BA | 52 | No. Rooms |
| BB | 53 | No. Beds |

#### معلومات إضافية (BC–BN)
| العمود | الفهرس | المحتوى |
|---|---|---|
| BC | 54 | رابط الملف (Photo link) |
| BD | 55 | بداية الاعاشة |
| BE | 56 | نهاية الاعاشة |
| BF | 57 | المبيعات (البيع) |
| BG | 58 | المتبقي |
| BH | 59 | النسبة المئوية |
| BI | 60 | Name Of Pakg English |
| BJ | 61 | رابط الحجز |
| BK–BM | 62–64 | (فارغ) |
| BN | 65 | التنقل — وسيلة نقل الحاج بين فندق مكة وفندق المدينة (والعكس). باقي التنقلات بالحافلة دائماً |

#### عقود السكن — أرقام العقود في منصة نسك (BO–BQ)
| العمود | الفهرس | المحتوى |
|---|---|---|
| **BO** | 66 | رقم عقد نسك — الفندق 1 |
| **BP** | 67 | رقم عقد نسك — الفندق 2 |
| **BQ** | 68 | رقم عقد نسك — الفندق 3 |

### شيت الطيران — هيكل الأعمدة

#### معلومات أساسية (A–U)
| العمود | الفهرس | المحتوى |
|---|---|---|
| A | 0 | No |
| B | 1 | PNR |
| C | 2 | Supplier |
| D | 3 | Status |
| E | 4 | Country |
| F | 5 | City |
| G | 6 | Airline |
| H | 7 | PAX |
| I | 8 | No. Days |
| J–U | 9–20 | AMOUNT (Fare source, Fare, Currency, Fare SAR, Add, TOTAL, PRICE NUSK, الفرق, PRICE NUSK, difrent, Profit, Total Nusuk) |

#### رحلات الذهاب (V–AI)
| النطاق | المحتوى |
|---|---|
| V–AB (21–27) | ذهاب 1: FlightNo1, Date TAKEOFF, TIME, From, To, DATE LANDING, LANDING TIME |
| AC–AI (28–34) | ذهاب 2: FlightNo2, Date TAKEOFF, TIME, From, To, DATE LANDING, LANDING TIME |

#### رحلات العودة (AJ–AW)
| النطاق | المحتوى |
|---|---|
| AJ–AP (35–41) | عودة 1: FlightNo1, Date TAKEOFF, TIME, From, To, DATE LANDING, LANDING TIME |
| AQ–AW (42–48) | عودة 2: FlightNo2, Date TAKEOFF, TIME, From, To, DATE LANDING, LANDING TIME |

#### البرامج المرتبطة (AX–BQ)
| النطاق | المحتوى |
|---|---|
| AX–BQ (49–68) | PKGname 1–10 مع أسماء (10 باقات لكل رحلة) |

#### المبيعات والعقود (BR+)
| العمود | المحتوى |
|---|---|
| BR–BS (69–70) | Sales, Remaining |
| BT–BZ (71–77) | CONTRACT (Total, Currency, SAR), DEPOSIT 1 (Date, Amount, Currency, Total) |

### شيت رحلة الحاج — هيكل الأعمدة (48 عمود)

| العمود | الفهرس | المحتوى |
|---|---|---|
| A | 0 | BookingId |
| B | 1 | PackageId |
| C | 2 | ServiceProviderId |
| D | 3 | PackageYear |
| E | 4 | CampName |
| F | 5 | ApplicantId |
| G | 6 | Group Number |
| H | 7 | Name |
| I | 8 | Passport |
| J | 9 | Email |
| K | 10 | IsMain |
| L | 11 | Gender |
| M | 12 | NationalityEn |
| N | 13 | NationalityAr |
| O | 14 | CountryResidenceEn |
| P | 15 | CountryResidenceAr |
| Q | 16 | ArrivalAirlineNameAr |
| R | 17 | ArrivalAirlineNameEn |
| S | 18 | ArrivalArrivalTime |
| T | 19 | ArrivalArriveCity |
| U | 20 | ArrivalArriveDate |
| V | 21 | ArrivalDepartureCity |
| W | 22 | ArrivalDepartureDate |
| X | 23 | ArrivalDepartureTime |
| Y | 24 | ArrivalFlightNumber |
| Z | 25 | ArrivalFlightType |
| AA | 26 | ReturnAirlineNameAr |
| AB | 27 | ReturnAirlineNameEn |
| AC | 28 | ReturnArrivalTime |
| AD | 29 | ReturnArriveCity |
| AE | 30 | ReturnArriveDate |
| AF | 31 | ReturnDepartureCity |
| AG | 32 | ReturnDepartureDate |
| AH | 33 | ReturnDepartureTime |
| AI | 34 | ReturnFlightNumber |
| AJ | 35 | ReturnFlightType |
| AK | 36 | FirstHouse |
| AL | 37 | FirstHouseStartDate |
| AM | 38 | FirstHouseEndDate |
| AN | 39 | LastHouse |
| AO | 40 | LastHouseStartDate |
| AP | 41 | LastHouseEndDate |
| AQ | 42 | MakkahAr |
| AR | 43 | MakkahEn |
| AS | 44 | MakkahShiftingAr |
| AT | 45 | MakkahShiftingEn |
| AU | 46 | MadinahAr |
| AV | 47 | MadinahEn |

### شيت الفنادق
| العمود | الفهرس | المحتوى |
|---|---|---|
| A | 0 | الاسم (عربي) |
| B | 1 | الموقع (Med/Mak) |
| C | 2 | Name (English) |
| D | 3 | رابط الموقع (Google Maps) |

### شيت Presonal Details — هيكل الأعمدة (33 عمود)

| الفهرس | اسم العمود | الوصف |
|---|---|---|
| 0 | الرقم التسلسلي | معرّف فريد لكل حاج (= ApplicantId القديم في رحلة الحاج) |
| 1 | رقم المجموعة | معرّف المجموعة/الطلب (= ApplicationId في رحلة الحاج 2) |
| 2 | نوع الحاج | رئيسي / عضو عائلة |
| 3 | فئة الحجاج | أفراد / ... |
| 4 | الجنس | ذكر / انثى |
| 5 | رقم جواز السفر | رقم الجواز |
| 6 | تاريخ انتهاء صلاحية جواز السفر | |
| 7 | تاريخ الإصدار | |
| 8 | الاسم الأول (العربية) | |
| 9 | اسم العائلة (العربية) | |
| 10 | الاسم الأول (الإنجليزية) | |
| 11 | اسم العائلة (الإنجليزية) | |
| 12 | تاريخ الميلاد | |
| 13 | البريد الإلكتروني | |
| 14 | رقم الجوال | |
| 15 | اسم المرشد (انجليزي) | |
| 16 | بلد الإقامة | |
| 17 | الجنسية | |
| 18 | رقم الباقة | |
| 19 | اسم الباقة | |
| 20 | نوع عقد الطيران | B2B / B2C |
| 21 | اسم العقد | |
| 22 | حالة التأشيرة | |
| 23 | داخل المملكة | |
| 24 | رقم التذكرة | |
| 25 | رابط التذكرة | |
| 26 | رقم الفاتورة | |
| 27 | المخيم | |
| 28 | نوع النقل (الوصول) | |
| 29 | وقت الوصول - النقل | |
| 30 | نوع النقل (المغادرة) | |
| 31 | وقت المغادرة - النقل | |
| 32 | تفاصيل الحجز | رابط نسك |

### شيت رحلة الحاج 2 — هيكل الأعمدة (42 عمود)

| الفهرس | اسم العمود | الوصف |
|---|---|---|
| 0 | BookingId | رقم الحجز |
| 1 | PackageId | رقم الباقة |
| 2 | Service Provider | اسم مقدم الخدمة (كان ServiceProviderId) |
| 3 | PackageYear | سنة الباقة |
| 4 | CampName | اسم المخيم |
| 5 | ApplicationId | رقم الطلب/المجموعة (= رقم المجموعة في PD) |
| 6 | IsMain | 1=رئيسي, 0=عضو عائلة |
| 7 | Gender | Male/Female |
| 8 | NationalityEn | الجنسية بالإنجليزي |
| 9 | CountryResidenceEn | بلد الإقامة بالإنجليزي |
| 10-41 | بيانات الطيران والفنادق | نفس الأعمدة 16-47 في رحلة الحاج القديم |

---

## آلية دمج البيانات (JourneyMerger)

> **السبب:** منصة نسك غيّرت شكل البيانات — فصلت البيانات الشخصية عن بيانات الرحلة.
> **الحل:** سكريبت `JourneyMerger.js` يدمجهما في شيت "رحلة الحاج " بالشكل القديم (48 عمود).

### كيف يعمل:
```
رحلة الحاج 2 (رحلة + فنادق) ─┐
                                 ├──→ JourneyMerger ──→ رحلة الحاج  (48 عمود بالشكل القديم)
Presonal Details (بيانات شخصية) ─┘
```

### مفاتيح الربط:
- **بين الشيتين:** `ApplicationId` (رحلة الحاج 2) = `رقم المجموعة` (Presonal Details)
- **بين الأفراد داخل المجموعة:** `Gender + IsMain` ↔ `الجنس + نوع الحاج`

### خريطة الأعمدة المدموجة:

| فهرس الناتج | العمود | المصدر |
|---|---|---|
| 0-4 | BookingId → CampName | رحلة الحاج 2 |
| 5 | ApplicantId | PD → الرقم التسلسلي |
| 6 | Group Number | رحلة الحاج 2 → ApplicationId |
| 7 | Name | PD → الاسم الأول + اسم العائلة (إنجليزي) |
| 8 | Passport | PD → رقم جواز السفر |
| 9 | Email | PD → البريد الإلكتروني |
| 10-11 | IsMain, Gender | رحلة الحاج 2 |
| 12 | NationalityEn | رحلة الحاج 2 |
| 13 | NationalityAr | PD → الجنسية |
| 14 | CountryResidenceEn | رحلة الحاج 2 |
| 15 | CountryResidenceAr | PD → بلد الإقامة |
| 16-47 | طيران + فنادق | رحلة الحاج 2 (نفس الترتيب) |

### التشغيل:
- **يدوي:** من قائمة "🔄 دمج البيانات" في الشيت، أو تشغيل `mergeJourneyData()` من محرر السكريبت
- **تلقائي:** يمكن إعداد Trigger كل 15 دقيقة
- **يحافظ على:** بيانات Reception Airport (أعمدة 48+) عبر حفظها واستعادتها بمفتاح BookingId|Passport

### ملف السكريبت:
`Projects/Ikram/JourneyMerger.js` — تم رفعه عبر `clasp push` بتاريخ 2026-03-26

---

## APIs المستخدمة
- **Google Sheets API** — الوصول المباشر للبيانات (Service Account في SheetData)
- **AeroDataBox** — بيانات الرحلات الجوية (RapidAPI)
- **Telegram Bot API** — IkramHajjBot + IkramAgent
- **WhatsApp** — IkramAgent

---

## الملفات الرئيسية في كل تطبيق
- `appsscript.json` — إعدادات المشروع
- `CompleteScript.js` — السكريبت الرئيسي (Ikram)
- `AgentCore.js` — المحرك الرئيسي (IkramAgent)
- `AgentConfig.js` — الإعدادات والرسائل (IkramAgent)
- `DataFetcher.js` — جلب البيانات (IkramAgent)
- `LeadManager.js` — إدارة العملاء المحتملين (IkramAgent)
- `TelegramHandler.js / WhatsAppHandler.js` — معالجات القنوات
- `JourneyMerger.js` — دمج رحلة الحاج 2 + Presonal Details → رحلة الحاج (Ikram)
- `SheetData/refresh.js` — تحميل البيانات محلياً (Node.js)
- `SheetData/ExportToCSV.js` — تصدير تلقائي كل ساعة (GAS trigger)

---

## Git & المزامنة بين الأجهزة

**Repo:** `C:\Users\mubar\Ekram-Aldyf\` (محلي + GitHub)
**Remote:** `https://github.com/mahmoudekramhajj/ikram-projects.git`
**الفرع الافتراضي:** `master` ← **ليس main — تحقق دائماً بـ `git branch` قبل الدفع**
**سجل الجلسات:** `SESSION_LOG.md` — يُحدّث تلقائياً نهاية كل محادثة

### طريقة العمل:
1. **بدء العمل على أي جهاز:** `git pull` ثم اكتب "تابع"
2. **نهاية العمل:** Claude يحدّث SESSION_LOG.md → المستخدم يكتب "ارفع" → `git push`
3. **clasp push** لنشر التعديلات على Google Apps Script
4. **عند طلب PR:** أنشئ feature branch أولاً، لا تدفع مباشرة لـ master

> **تعليمة لـ Claude:**
> - عند "تابع": اقرأ SESSION_LOG.md وأكمل من آخر جلسة
> - نهاية كل محادثة: أضف جلسة جديدة في SESSION_LOG.md + حدّث "آخر نقطة عمل" أدناه

### قواعد النشر (clasp)
- عند إعادة النشر: **حدّث النشرة الموجودة** — لا تنشئ واحدة جديدة (يغيّر الرابط)
- بعد النشر: تحقق أن رابط Web App يعمل (ليس 404)
- التفويض النهائي يجب أن يتم يدوياً من المستخدم في المتصفح
- استخدم: `clasp push && clasp deploy --deploymentId <existing-id>`

### معرّفات السكريبتات (Script IDs)

| التطبيق | Script ID |
|---|---|
| Ikram | `1txmPRRjW5lx1k7RGhUXk0CxO3OOvvgiUeSmszZMsF3kcBtcgw7v76dbd` |
| IkramHajjBot | `1hJ5pojmbbYTy3Xv2fEt2yDwISpBof4GF8hywZHuL7LQvuJ-tgUzwrLYi` |
| IkramAdmin | `1f-zr0LuP9_NjDz0kLYHj5C60qQsVvWQCvf5pQzoFkR5sB843D-Z_ZZW1` |
| Airport Search | `1Izf0A04wscTnuIaex1-oipNR6y8LHWNr_w6qa2YbCIhS_EiKtyVkZcuz` |
| Hotel Management | `1FhtOVMzFALclBnYtBOP1fL89cgx6fLtEI_YxLijARg6QNv7nGGaWbJ3U` |
| Reception Airport | `14g-5fXLwm1-kAXuKSCFP-at1rJvemajuZMpBFUdOrohJbwpdUan-2jif` |
| Pilgrim App | `1a7NzodkK9oiwwl4K-cMP4Qhr6mWCFPpm01RfzvB6xCzsVxo44xGGIpYX` |
| Guide App | `1Ssf8TLv4pM0z5KARbzvu5m0hSHaiDHoZeg5qrxmvZS8v7JGVhPxKD30O` |
| Sales Operations Report | `1q86OcUWZOI29jfQMLxwohrwgaliq5yOi96wJVEfdX0zYrxa7k8jkObk-` |
| Report Builder | `1pZjXdk0So0FeOokQGTOy-BHFp7DDyxO7NFCMNOeBBB-Cp7HZ4inztU0F` |
| Transport Management | `1q7S0gljNiKtm3UCfHBmLBO_Dq3PuWmG22Ixcv-ATJUFySGiTjpyw8uQt` |
| PilgrimStats | `1GaNM_2vPTp4KNMAY9tmm5Z5pv38ucviz52mYZkVYiZeL7hA6fO5LiV3F` |
| Mina Camp Search | `1grSxyHbhrNbz5o8dsc3gy4vYQc3JRe9f-hrN8CGd36RF-ttweXtPEXu3` |
| GDS | `181VvHqb5RpNebfMkRMoqi9oNo-26rDbAPF_ZLjkLCka81xS0cJXsUax4` |
| B2C Sync | `1IJ9LsGSSlfdzHkbzn16zc4i--TWdrYvV4zUh4oSNv6O-xtMMyVYcrGLz` |
| Holiday In Bakkah | `1uxCF_apaYDy8xA29NZOW-W5eWefYHNjmvsY-6qr7bD82iwXMf5BrVFN_` |
| PNR Target Countries Sync | `1ZNY7dC2k96mDyKR3cZoNEf59vNAGXlC_BUx2Nkgs5KK4_9ZAxUK7Z51N` |
| IkramAgent | `1hwFQRZjt4ZfwgHcoM64IWdbxHPun4GXi8rrBV2ytHA5ttY0hBQuXdAZv` |
| TourGuide Manager | `1dSWbcN-CVc3IyFvJIk0ravd2UJNl7lhMFd_IhT37s5CkILegS_YMMOGE` |
| Ticket Processor | `1TUIJyz8kYhxRINh73A3rfwBGCrueP-XoF7e1Jk6PTOskcAflDLLKiXg-` |

---

## خرائط الأعمدة في الكود (Code Constants)

هذه الثوابت مُعرَّفة في `IkramHajjBot/Config.js` وتُستخدم في كل مشاريع GAS — **صفر-مبني (0-based index)**:

### PD — Presonal Details
```js
PD = {
  SEQ:0, GROUP:1, TYPE:2, CATEGORY:3, GENDER:4,
  PASSPORT:5, PASSPORT_EXP:6, PASSPORT_ISSUE:7,
  FIRST_NAME_AR:8, LAST_NAME_AR:9, FIRST_NAME_EN:10, LAST_NAME_EN:11,
  DOB:12, EMAIL:13, PHONE:14, GUIDE_NAME:15,
  COUNTRY_RESIDENCE:16, NATIONALITY:17,
  PACKAGE_NO:18, PACKAGE_NAME:19,
  FLIGHT_TYPE:20,   // B2B | B2C
  CONTRACT_NAME:21, VISA_STATUS:22, IN_KSA:23,
  TICKET_NO:24, TICKET_URL:25, INVOICE_NO:26, CAMP:27
}
```

### FLT — شيت الطيران (B2B)
```js
FLT = {
  NO:0, PNR:1, SUPPLIER:2, STATUS:3, COUNTRY:4, CITY:5, AIRLINE:6, PAX:7,
  GO1_FLIGHT:21, GO1_TAKEOFF_DATE:22, GO1_TAKEOFF_TIME:23, GO1_FROM:24, GO1_TO:25, GO1_LAND_DATE:26, GO1_LAND_TIME:27,
  GO2_FLIGHT:28, GO2_TAKEOFF_DATE:29, GO2_TAKEOFF_TIME:30, GO2_FROM:31, GO2_TO:32, GO2_LAND_DATE:33, GO2_LAND_TIME:34,
  RET1_FLIGHT:35, RET1_TAKEOFF_DATE:36, RET1_TAKEOFF_TIME:37, RET1_FROM:38, RET1_TO:39, RET1_LAND_DATE:40, RET1_LAND_TIME:41,
  RET2_FLIGHT:42, RET2_TAKEOFF_DATE:43, RET2_TAKEOFF_TIME:44, RET2_FROM:45, RET2_TO:46, RET2_LAND_DATE:47, RET2_LAND_TIME:48,
  PKG1_NO:50, ..., PKG10_NO:68,
  CONTRACT_NAME:90   // CM — مفتاح الربط مع PD.CONTRACT_NAME
}
```
> مفتاح الربط B2B: `pd[PD.CONTRACT_NAME]` === `flt[FLT.CONTRACT_NAME]`

### PKG — شيت الباقات
```js
PKG = {
  NO:0, NUSK_NO:1, NAME_AR:2, CATEGORY:3, IKRAM_NO:4,
  PRICE:5, DATE_START:6, DATE_END:7, NO_DAYS:8, CITY_START:9, NO_PILGRIM:10,
  H1_CITY:11, H1_NAME_AR:12, H1_NAME_EN:13, H1_CHECKIN:14, H1_CHECKOUT:15,
  H2_CITY:26, H2_NAME_AR:27, H2_NAME_EN:28, H2_CHECKIN:29, H2_CHECKOUT:30,
  H3_CITY:41, H3_NAME_AR:42, H3_NAME_EN:43, H3_CHECKIN:44, H3_CHECKOUT:45,
  NAME_EN:60, TRANSPORT:65
}
```
> مفتاح الربط: `pd[PD.PACKAGE_NO]` === `pkg[PKG.NUSK_NO]` (من صف 2، الصف 1 = header ثانٍ)

### B2CI — شيت B2C (يشترك مع PD في الأعمدة 0-31)
```js
B2CI = {
  PASSPORT:5,   // نفس PD.PASSPORT
  // ذهاب
  ARR1_FLIGHT:32, ARR1_DATE:33, ARR1_TIME:34, ARR1_FROM:35, ARR1_TO:36, ARR1_LAND_DATE:37, ARR1_LAND_TIME:38,
  ARR2_FLIGHT:39, ARR2_DATE:40, ARR2_TIME:41, ARR2_FROM:42, ARR2_TO:43, ARR2_LAND_DATE:44, ARR2_LAND_TIME:45,
  // عودة
  RET1_FLIGHT:46, RET1_DATE:47, RET1_TIME:48, RET1_FROM:49, RET1_TO:50, RET1_LAND_DATE:51, RET1_LAND_TIME:52,
  RET2_FLIGHT:53, RET2_DATE:54, RET2_TIME:55, RET2_FROM:56, RET2_TO:57, RET2_LAND_DATE:58, RET2_LAND_TIME:59,
  PNR:60,
  // قطع إضافية (3-leg)
  DEP0_FLIGHT:70, ..., DEP0_LAND_TIME:76,   // الذهاب الأبعد
  RET3_FLIGHT:77, ..., RET3_LAND_TIME:83    // العودة الأبعد
}
```
> مفتاح الربط B2C: `pd[PD.PASSPORT]` === `b2c[B2CI.PASSPORT]` (وليس CONTRACT_NAME)

### GDS2.Config.COL — نسخة 1-based للـ GDS
```
A=1 (SERIAL) | B=2 (GROUP) | C=3 (TYPE) | ... | F=6 (PASSPORT) | U=21 (CONTRACT_TYPE) | Z=26 (TICKET_URL)
```

---

## علاقة B2B و B2C

### تعريف
| | B2B | B2C |
|---|---|---|
| **الفلسفة** | إكرام الضيف اشترت التذاكر بالجملة من شركة طيران | الحاج اشترى تذكرته بشكل مستقل |
| **المسؤولية** | إكرام الضيف مسؤولة عن أي خطأ | الحاج يتواصل مع الناقل مباشرة |
| **مصدر بيانات الطيران** | شيت "الطيران" (B2B contracts) | شيت "B2C" (مُعبَّأ من GDS + PDF نسك) |
| **مفتاح الربط** | `pd[CONTRACT_NAME]` ↔ `flt[CONTRACT_NAME]` | `pd[PASSPORT]` ↔ `b2c[PASSPORT]` |
| **رابط التذكرة** | من B2B شيت أو PD[TICKET_URL] | من B2C شيت أو PD[TICKET_URL] |
| **الحقل المُحدِّد** | `pd[PD.FLIGHT_TYPE] === 'B2B'` | `pd[PD.FLIGHT_TYPE] === 'B2C'` |

### منطق اختيار المصدر في الكود
```js
if (flightType === 'B2C') {
  // البحث في شيت B2C بالجواز
} else {
  // B2B: البحث في شيت الطيران باسم العقد
  var contractName = pd[PD.CONTRACT_NAME];
  // ابحث عن flt[FLT.CONTRACT_NAME] === contractName
}
```

### قاعدة B2C الحرجة
- **حجاج B2C بدون رابط تذكرة** لا يدخلون شيت B2C — غيابهم ليس خطأ في الكود
- بلاغات طيران B2C ليست من مسؤوليتنا — نُبلّغ الحاج بالتواصل مع الناقل

### GDS — مصدر بيانات B2C
مشروع GDS يعبّئ شيت B2C من:
1. تحميل PDF التذكرة من Drive (عبر رابط `pd[TICKET_URL]`)
2. استخراج بيانات الرحلة بـ Claude AI
3. كتابة النتيجة في أعمدة ARR/RET في B2C

---

## معمارية IkramHajjBot

### هيكل الملفات (18 ملف)
| الملف | المسؤولية |
|---|---|
| `Router.js` | `doPost()` ← نقطة دخول Webhook، يوجّه لـ `handleMessage_` / `handleCallback_` |
| `Config.js` | BOT_TOKEN، SHEET_ID، ثوابت الأعمدة (PD/FLT/PKG/B2CI)، خرائط المطارات/الجنسيات |
| `DataFetcher.js` | `findPilgrimByPassport_()` — يبني Virtual Row من PD+الباقات+الطيران/B2C |
| `Handlers.js` | معالجة كل callback (`handleMyFlight_`, `handleMyHotel_`, ...) |
| `TelegramAPI.js` | `sendMessage_()`, `sendPhoto_()`, `answerCallback_()` |
| `Auth.js` | `handlePassportInput_()` — التحقق من الحاج بالجواز، `isAdmin_()` |
| `Menu.js` | `sendMainMenu_()`, `sendLanguageButtons_()` |
| `i18n.js` | `T_(key, lang)` — ترجمة الرسائل (AR/EN/FR/DE/IT/ES) |
| `TicketsMonitor.js` | مراقبة بلاغات الحجاج كل 15 دقيقة + إشعارات للمدير |
| `ReportHandler.js` | `handleReportError_()` — تسجيل بلاغ جديد في شيت Tickets |
| `Notifications.js` | إشعارات نظام (وصول، تغيير، إعلانات) |
| `FlightChangeNotifier.js` | إخطار الحجاج بتغييرات الطيران (kill switch: `FLIGHT_NOTIFY_ENABLED=false`) |
| `Broadcast.js` | `/broadcast` — إرسال جماعي للمدراء |
| `Helpers.js` | `formatTime_()`, `formatDate_()`, دوال مساعدة |
| `DataCollector.js` | جمع بيانات إضافية (غرفة، رقم جوال، صورة جواز) |
| `Setup.js` | `setWebhook_()` — ربط Bot بـ Webhook |
| `ClaudeAPI.js` | استدعاء Claude API لمعالجة النصوص |
| `Auth.js` | جلسات المستخدم عبر `BotSessions` sheet + CacheService |

### آلية العمل (Webhook)
```
الحاج يرسل رسالة تيليغرام
    ↓
Telegram API → POST إلى Web App URL (doPost)
    ↓
Router.js: message أو callback_query؟
    ↓ message                    ↓ callback_query
handleMessage_()           handleCallback_()
    ↓
تحقق الحالة (awaiting_passport / verified / admin)
    ↓
جلب بيانات الحاج (findPilgrimByPassport_)
    ↓ B2B                        ↓ B2C
شيت الطيران                  شيت B2C
(CONTRACT_NAME)              (PASSPORT)
    ↓
إرسال الرد (TelegramAPI.js)
```

### الأوامر والأزرار
| command/callback | الوظيفة |
|---|---|
| `/start` | ترحيب + اختيار اللغة |
| `/menu` | القائمة الرئيسية (للمستخدمين المُتحقَّق منهم) |
| `/broadcast` | إرسال جماعي (مدراء فقط) |
| `/stats` | إحصائيات (مدراء فقط) |
| `my_flight` | بيانات الطيران |
| `my_hotel` | بيانات الفندق |
| `my_package` | تفاصيل الباقة |
| `my_transport` | موعد النقل |
| `my_data` | البيانات الشخصية |
| `report_error` | فتح بلاغ جديد → شيت Tickets |
| `tm_*` | أزرار TicketsMonitor (مدراء) |
| `rpt_*` | أزرار إجراءات البلاغات (مدراء) |

### مجموعات العمليات (Ops Groups)
```js
OPS_GROUPS = {
  madinah: '-4849598886',      // مطار المدينة
  jeddah_t1: '-5220583519',    // مطار جدة - صالة 1
  jeddah_north: '-5267173490'  // مطار جدة - الصالة الشمالية
}
OPS_LOCATION_GROUPS = {
  makkah: '-4916619724',
  madinah: '-5284394785',
  mashaaer: '-5268778683'
}
```

### CacheService
- `session_{chatId}` → بيانات الجلسة (TTL 5 دقائق)
- `pilgrim_{passport}` → بيانات الحاج المُجمَّعة
- `flightdata_{passport}` → بيانات الطيران (legs)
- `b2cflight_{passport}` → بيانات B2C
- `transport_{packageId}` → بيانات النقل
- إلغاء الـ cache يدوياً: `clearPilgrimCache_(passport)`

---

## سير نشر Apps Script (clasp)

### الأدوات
- `clasp` — CLI رسمي لرفع الكود ونشره
- `deploy.sh` — سكريبت أتمتة يُنفذ الخطوات الأربع لكل مشروع
- `deployments.json` — يحفظ `deploymentId` الثابت لكل مشروع (URL لا يتغيّر)

### خطوات النشر (لمشروع واحد)
```bash
cd Projects/<ProjectName>
clasp push -f                          # رفع الكود
clasp deploy -i <deploymentId> -d "..."  # تحديث النشرة الموجودة (لا تنشئ جديدة)
curl <webAppUrl>                        # تحقق أن URL يعمل (200/302)
```

### النشر الكامل (كل المشاريع)
```bash
cd Projects
./deploy.sh              # ينشر كل 9 تطبيقات ويب
./deploy.sh "Ikram"      # ينشر مشروعاً واحداً فقط
```

### deploy.sh خطوات التفصيل
1. `clasp push -f` → رفع الكود
2. `clasp deploy -i <id> -d "auto YYYY-MM-DD"` → تحديث نفس النشرة
3. `curl -L <url>` → تحقق أن الـ URL يعمل (200 أو 302)
4. تحديث `deployments.json` بـ version + timestamp

### قواعد النشر الحرجة
- **لا تنشئ نشرة جديدة** — استخدم `clasp deploy -i <existingId>` دائماً وإلا تتغيّر الروابط
- **clasp push أولاً ثم deploy** — بدون deploy لا يصل الكود للمستخدمين
- **الـ Trigger قد ينفّذ كوداً قديماً** بعد deploy — احذف Trigger وأعد إنشاءه إن لزم
- كل مشروع له `deploymentId` ثابت في `deployments.json`

### تبعيات OAuth لكل مشروع
| المشروع | صلاحيات إضافية |
|---|---|
| FlightChangeProcessor | `gmail.modify`, `drive`, `documents`, `script.scriptapp` |
| GDS | `drive` (Advanced), `script.send_mail`, `userinfo.email` |
| IkramHajjBot | `gmail.compose`, `gmail.modify`, `drive`, `script.scriptapp` |
| TicketLinker | `drive` (Advanced), `documents`, `userinfo.email` |
| Ticket Processor | `drive` (Advanced) |
| Airport Search | `cloud-platform`, `script.send_mail` |

> **تحذير:** `Report Builder` يستخدم `America/New_York` — خطأ محتمل، الصحيح `Asia/Riyadh`

---

## الشيتات الإضافية (غير موثَّقة سابقاً)

### Tickets — بلاغات الحجاج
| العمود | الحقل | الوصف |
|---|---|---|
| ticketId | TKT-XXXX | معرّف البلاغ |
| chatId | Telegram chatId | معرّف المحادثة |
| passport | رقم الجواز | مفتاح الربط مع PD |
| name | اسم الحاج | من النص |
| section | hotel/flight/transport/personal/visa | نوع المشكلة |
| issue | نص المشكلة | |
| correction | الشيء الصحيح | |
| status | new/in_progress/resolved/closed | |
| createdAt/resolvedAt/resolvedBy | تواريخ | |

### Hotel Check-in — تسجيل دخول الفنادق
| العمود | الوصف |
|---|---|
| BookingId | مفتاح الربط |
| HotelName/HotelCity | تفاصيل الفندق |
| RoomGroup_ID | رقم مجموعة الغرفة |
| Room_Number | رقم الغرفة |
| CheckIn_Status | pending/checked_in |
| CheckIn_Time/CheckIn_By | وقت وموظف التسجيل |

### AdminMessages — رسائل إدارية للبوت
```js
AM = {
  ID:0, TITLE:1, MSG_AR:2, MSG_EN:3, MSG_FR:4, MSG_DE:5, MSG_IT:6, MSG_ES:7,
  IMAGE_URL:8, FILE_URL:9, FILE_NAME:10,
  TARGET:11, TARGET_VALUE:12,   // الجمهور المستهدف
  PRIORITY:13, STATUS:14, SENT_AT:15, SENT_COUNT:16, FAIL_COUNT:17, CREATED_BY:18
}
```

### تغييرات الطيران — FlightChangeProcessor
| الأعمدة | المحتوى |
|---|---|
| 0-3 | رقم التغيير CHG-XXXX، PNR، BookingID، رقم الحادثة |
| 4-9 | اسم الحاج، الرقم التسلسلي، الجواز، الباقة، اسم الباقة، نوع العقد |
| 10-17 | ذهاب1 (رقم رحلة، من، إلى، تواريخ، أوقات) |
| 18-25 | ذهاب2 |
| 26-33 | عودة1 |
| 34-41 | عودة2 |
| 42-49 | معلومات الوصول/المغادرة النهائية |

---

## بنية كل مشروع (ملفات .js)

### IkramHajjBot (18 ملف)
`Router` → `Handlers` → `DataFetcher` + `TelegramAPI` + `Auth` + `Menu` + `i18n`
`TicketsMonitor` + `ReportHandler` + `Notifications` + `FlightChangeNotifier` + `Broadcast`

### GDS (18 ملف)
`Pipeline` → `PdfDownloader` → `PdfTextExtractor` → `PromptBuilder` → `AnthropicClient` → `FlightWriter`
`Classifier` + `FamilyProcessor` + `State` + `Logger` + `SchemaManager` + `IATARegistry` + `AirlineRegistry`

### FlightChangeProcessor (11 ملف أساسي + 9 أدوات تشخيص)
`EmailReader` → `NusukParser` → `PilgrimMatcher` → `SheetWriter`
`FlightValidator` + `ClaudeParser` + `Audit` + `Config`
أدوات: `_Setup`, `_TriggerCheck`, `_DeepAudit`, `_FullAudit`, `_RecoveryAudit`, ...

### TicketLinker (13 ملف)
`Pipeline` → `Matcher` + `PnrResolver` + `ClaudeMatcher` → `Writer`
`SourceSync` + `HourlyTrigger` + `TelegramReport` + `UnresolvedHandler`

### Transport Management (6 ملف)
`TransportApp` (Web App) + `TransportData` + `TransportActions` + `TransportHelpers` + `TransportPlanner`

---

## سير العمل (Workflow)

- **إصلاحات شاملة لا تراكمية:** فضّل الإصلاحات الشاملة على دورات deploy-test التراكمية. اجمع التغييرات المرتبطة وانشرها مرة واحدة بدلاً من تعديلات صغيرة وإعادة نشر متكررة.

---

## التصحيح (Debugging)

- **افحص تعديلاتك أولاً:** عند التصحيح، افحص تعديلاتك الأخيرة على الكود أولاً قبل التحقيق في الأسباب الخارجية (بيانات الشيتات، أسماء الأعمدة، الـ APIs). إذا كان قد تم clasp push للتو، تحقق من عدم تضمّن ملفات غير ذات صلة.

---

## آخر نقطة عمل

- **التاريخ:** 2026-04-17
- **المشروع:** FlightChangeProcessor — معالجة تذاكر الطيران من إيميلات نسك
- **الحالة:** deployment @42 منشور. التحديثات: (1) إصلاح regex استخراج الأسماء لدعم صيغة نسك الجديدة `NAME NAME (Adult)` بدون ألقاب Mr/Mrs — حلّ 71 PDF فاشل (2) دعم المسافرين المتعددين في PDF واحد — صف منفصل لكل مسافر (3) منطق مطابقة تدريجي: جرب الاسم الخام أولاً، ثم نظّف أكواد المطارات من البداية وأعد المحاولة — حل مشكلة التقاط `DXB DUR FATIMA` كاسم مع الحفاظ على أسماء عائلة شرعية مثل `DAR` (4) استبدال الأسهم `→` بـ `من X إلى Y` لدعم RTL العربي (5) إضافة عمود "نوع الحجز" (B2B من PD / B2C) في الشيتين (6) دالة `deleteOldSimilarSheets` لتنظيف Drive — حُذف 22 ملف قديم. **النتيجة:** 253 صف، 251 متطابق (99.2%). الشيتات النشطة: تغييرات الطيران `1NhFiQ_0K7F90pX-mKfuM3UjOsFwhpYsEizyIL65LCEA`، مقارنة الأسماء `13fOtOFHpfsbxL1BwjmIqdPgBS4YOosrxl9tp3h3ci4w`. الملفات المعدّلة: NusukParser.js، EmailReader.js، SheetWriter.js، PilgrimMatcher.js.
- **المؤجل:** إصلاح شامل لـ `extractFlightLegs_` — تم تشخيص خلل `DOH→DOH` و `AUH→AUH` (السبب: blind pairing لـ dateTime entries + false positive في regex أرقام الرحلات يلتقط `XB 0` من `DXB 07h 40m`). المنطق الجديد المقترح: anchor-based (كل رقم رحلة = dep قبله + arr بعده) مع fallback لحالة آخر رحلة حيث الوصول يظهر قبل المغادرة + قطع النص عند أول حرف عربي لتجنب تكرار.
- **الخطوة التالية:** تنفيذ الإصلاح الشامل لـ `extractFlightLegs_` بالمنطق الجديد + مراجعة البيانات الناتجة فعلياً للتحقق من صحة المسارات + فحص لماذا الإيميلات الجديدة لم تُلتقط
