# إكرام الضيف — Ekram Aldyf

## قواعد العمل معي
- إذا لم تفهم الفكرة 100%، اطرح سؤالاً واحداً فقط وانتظر إجابتي
- أبلّغني بالتعديلات المقترحة والنتائج المتوقعة قبل التنفيذ
- لا تنفذ أي شيء حتى أعطيك الموافقة الصريحة

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
| **Ikram1** | النظام الرئيسي (إنتاج) — باقات + فنادق + رحلات v5 |
| **Ikram2** | النظام الرئيسي (نسخة ثانية/تطوير) |
| **IkramAgent** | وكيل ذكي متعدد القنوات (Telegram + WhatsApp) — يدعم AR/EN/FR |
| **IkramHajjBot** | بوت تيليغرام لمتابعة رحلة الحاج |
| **Pilgrim App** | تطبيق الحاج |
| **Guide App** | تطبيق المرشد السياحي |
| **TourGuide Manager** | إدارة المرشدين |
| **Reception Airport** | استقبال المطار |
| **Airport Search** | بحث في بيانات المطار |
| **Hotel Management** | إدارة الفنادق |
| **GDS** | نظام التوزيع العالمي للرحلات |
| **B2C Sync** | مزامنة بيانات B2C |
| **Mina Camp Search** | بحث في مخيمات منى |
| **Sales Operations Report** | تقارير المبيعات |
| **Holiday In Bakkah** | باقات العطل في مكة |
| **SheetData** | أدوات تصدير البيانات (Node.js + Google Sheets API) |
| **PNR Target Countries Sync** | مزامنة PNR للدول المستهدفة |

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

---

## APIs المستخدمة
- **Google Sheets API** — الوصول المباشر للبيانات (Service Account في SheetData)
- **AeroDataBox** — بيانات الرحلات الجوية (RapidAPI)
- **Telegram Bot API** — IkramHajjBot + IkramAgent
- **WhatsApp** — IkramAgent

---

## الملفات الرئيسية في كل تطبيق
- `appsscript.json` — إعدادات المشروع
- `CompleteScript.js` — السكريبت الرئيسي (Ikram1/2)
- `AgentCore.js` — المحرك الرئيسي (IkramAgent)
- `AgentConfig.js` — الإعدادات والرسائل (IkramAgent)
- `DataFetcher.js` — جلب البيانات (IkramAgent)
- `LeadManager.js` — إدارة العملاء المحتملين (IkramAgent)
- `TelegramHandler.js / WhatsAppHandler.js` — معالجات القنوات
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

- **التاريخ:** 2026-03-14
- **المشروع:** إعداد Git + SESSION_LOG
- **الحالة:** نسخ المشروع إلى مسار محلي، git init، أول commit (120 ملف)، إنشاء SESSION_LOG.md
- **الخطوة التالية:** إنشاء repo على GitHub + تنفيذ نفس الإعداد على جهاز المكتب + تحديد المطلوب من أعمدة عقود السكن/التنقل
