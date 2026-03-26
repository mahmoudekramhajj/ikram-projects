# إكرام الضيف — Ekram Aldyf

## قواعد العمل معي
- إذا لم تفهم الفكرة 100%، اطرح سؤالاً واحداً فقط وانتظر إجابتي
- أبلّغني بالتعديلات المقترحة والنتائج المتوقعة قبل التنفيذ
- لا تنفذ أي شيء حتى أعطيك الموافقة الصريحة
- **عند كتابة "تابع":** اقرأ ملف `SESSION_LOG.md` في جذر المشروع، ثم أكمل العمل من آخر جلسة مسجّلة فيه دون طرح أسئلة

---

## نظرة عامة

**مشروع واحد متكامل** لإدارة رحلات الحج والعمرة، مقسّم إلى تطبيقات فرعية لسهولة المتابعة.
جميع التطبيقات تستمد بياناتها من **شيت Ikram Abuown** الموحّد.

**المسار:** `C:\Users\mubar\Ekram-Aldyf\Projects\` (Git repo)
**المسار القديم (Google Drive):** `D:\اكرام الضيف\Ekram Aldyf\Projects\` (للقراءة فقط)
**التوقيت:** Asia/Riyadh
**البيئة:** Google Apps Script V8 runtime
**Spreadsheet ID:** `1z4b3BmTLDLvYUs8H8cPU8MJrOuvuN5GztZ9pLlYhF6s`

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

---

## خريطة تفاعل المشاريع

### الشيتات المركزية ومن يستخدمها

| الشيت | يقرأ منه | يكتب فيه |
|---|---|---|
| **رحلة الحاج** (6500+ صف) | Airport Search, IkramHajjBot, Pilgrim App, Guide App, Hotel Management, Mina Camp Search | Reception Airport, **JourneyMerger** (يكتب الشيت كاملاً) |
| **رحلة الحاج 2** (6500+ صف) | **JourneyMerger** (مصدر بيانات الرحلة) | منصة نسك (مصدر خارجي) |
| **Presonal Details** (6500+ صف) | **JourneyMerger** (مصدر البيانات الشخصية)، Pilgrim App, B2C Sync, Roomtype Builder, TourGuide Manager | منصة نسك (مصدر خارجي) |
| **الباقات** (69 عمود) | Ikram, IkramAgent, IkramHajjBot, Hotel Management, Pilgrim App, Holiday In Bakkah | Ikram |
| **الطيران** (103 عمود) | Ikram, IkramAgent, GDS, B2C Sync, Sales Report, Holiday In Bakkah | Ikram, GDS, B2C Sync, PNR Sync |
| **الفنادق** | Ikram, Hotel Management, Holiday In Bakkah | — |
| **مخيم مني** | Mina Camp Search | Ikram (Minacamp.js) |
| **Tour Guide** | Guide App, Ikram (Minacamp.js, TourGuide_Manager.js) | TourGuide Manager |
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
**سجل الجلسات:** `SESSION_LOG.md` — يُحدّث تلقائياً نهاية كل محادثة

### طريقة العمل:
1. **بدء العمل على أي جهاز:** `git pull` ثم اكتب "تابع"
2. **نهاية العمل:** Claude يحدّث SESSION_LOG.md → المستخدم يكتب "ارفع" → `git push`
3. **clasp push** لنشر التعديلات على Google Apps Script

> **تعليمة لـ Claude:**
> - عند "تابع": اقرأ SESSION_LOG.md وأكمل من آخر جلسة
> - نهاية كل محادثة: أضف جلسة جديدة في SESSION_LOG.md + حدّث "آخر نقطة عمل" أدناه

---

## آخر نقطة عمل

- **التاريخ:** 2026-03-26
- **المشروع:** JourneyMerger — دمج بيانات نسك الجديدة
- **الحالة:** سكريبت JourneyMerger.js جاهز ومرفوع عبر clasp push. يحتاج: أول تشغيل يدوي من محرر السكريبت + إعداد Trigger اختياري
- **الخطوة التالية:** تشغيل `mergeJourneyData()` من Google Apps Script + التحقق من النتائج + إعداد Trigger
