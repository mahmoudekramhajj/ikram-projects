# سجل الجلسات — Ekram Aldyf

> **تعليمة لـ Claude:** في نهاية كل محادثة، أضف جلسة جديدة في الأعلى (بعد هذا السطر مباشرة). عند بدء محادثة جديدة بكلمة "تابع"، اقرأ آخر جلسة وأكمل من حيث توقفنا.

---

## 2026-04-08 | تحديث شامل للمشاريع + رفع على GitHub

**ما تم إنجازه:**
- رفع كوميت شامل على `master` يتضمن 30 ملف (20 تعديل + 10 ملفات جديدة):
  - **Airport Search** — تحديثات الكود والواجهة
  - **Hotel Management** — تعديلات Checkin, Operations, Rooming, Export, Styles (11 ملف)
  - **Ikram** — تحديث الواجهة والبحث والستايل
  - **IkramAdmin** — تحديث Auth و Router
  - **Reception Airport** — تحسينات الاستقبال والواجهة
  - **PilgrimStats** — مشروع جديد كامل (إحصائيات الحجاج)

**التعديلات المطلوبة على صفحة المشاريع (index.html على فرع main):** — لم تُنفَّذ بعد
1. حذف: تطبيق الحاج + تطبيق المرشد
2. تقسيم التطبيقات إلى قسمين: "تطبيقات الحاج" و"تطبيقات الإدارة"
3. تطبيقات الحاج: بوت تيليجرام + نظام الباقات
4. تحديث روابط: الباقات، إدارة الفنادق، إدارة النقل، نظام المقابلات
5. رفع ملف الهوية البصرية + التأكد من توافق الموقع معها

**الخطوة التالية:**
- تنفيذ تعديلات صفحة المشاريع على فرع `main`
- رفع ملف الهوية البصرية على GitHub

---

## 2026-04-08 | PilgrimStats — مراجعة داشبورد ملخص العمليات + تصحيح البيانات

**ما تم إنجازه:**
- مراجعة شاملة لكل رقم في داشبورد `?view=summary` مقابل بيانات الشيت الحقيقية
- تصغير CSS (خطوط + padding) لحل مشكلة قطع العناصر في بطاقات الباقات والتأشيرات
- نشر v44 مع إصلاح الـ overflow

**الفروقات المكتشفة والمتفق على تصحيحها:**

| البيان | الداشبورد (قديم) | الشيت (صحيح) | الإجراء |
|---|---|---|---|
| إجمالي الحجاج | 6,800 | 6,800 | ثابت (سقف التعاقد) ✅ |
| إجمالي المرشدين | 200 | 200 | ثابت (المتاح) ✅ |
| الباقات | 52 مباعة / 7 متبقية | يُحسب من عمود BF (المبيعات vs السعة) | تفعيل حي |
| فئات الباقات | 4 فئات | 5 فئات (+ فاخرة انتقالية Luxury Shifting) | إضافة |
| مجاملة | 233 | 224 | تصحيح من الشيت |
| التأشيرات | أرقام قديمة | تحديث من عمود 22 | تفعيل حي |
| المخيمات | 2 مخيمات | 3 مخيمات (+ مخيم 72ب = 224 حاج) | إضافة |
| التذاكر في النظام | 63 | — | حذف |
| الدول | 9 دول | أعلى 10 دول | توسيع |

**الملفات المتأثرة:**
- `Projects/PilgrimStats/SummaryView.html` — CSS fixes (v44 deployed)
- `Projects/PilgrimStats/Code.js` — سيُعدّل لتفعيل البيانات الحية

**آخر deployment:** @44 — `AKfycbzmYrDYkiJPVhQBKrAKp1b6MS_33AFaP__CW-vrKiW6Hvbvfw2j3X-6p1thflkfH9OJPQ`

**الخطوة التالية:**
1. تفعيل `getSummaryStats()` ليقرأ من الشيتات بدل البيانات الثابتة
2. إضافة فاخرة انتقالية + مخيم 72ب + أعلى 10 دول
3. حذف "التذاكر في النظام"
4. نشر واختبار

---

## 2026-04-08 | إضافة تفضيلات الحاج للداشبورد

**ما تم إنجازه:**
- إضافة كارت "تفضيلات الحاج" (PilgrimStats) لصفحة المشاريع التقنية في قسم "تطبيقات الحاج"
- الرابط: `https://script.google.com/macros/s/AKfycbz83be2cu4xHN6Smi-wtRxA9btSX-Q92dfktg6zsGe0-z7RnwoEBsmODyorU5_J8DPi/exec`
- إصلاح مشكلة: GitHub Pages مُعد على فرع `main` لكن التعديل الأول رُفع خطأً على `master` — تم التصحيح بالتعديل مباشرة على `main`
- تغيير grid قسم "تطبيقات الحاج" من `grid-2` إلى `grid-3` (3 كروت)

**الملفات المتأثرة:**
- `index.html` (فرع main) — إضافة كارت + تعديل grid

**ملاحظة مهمة:**
- الداشبورد يُنشر من فرع `main` عبر GitHub Pages — أي تعديل على `index.html` يجب أن يكون على `main`
- فرع `master` يحتوي على المشاريع والسكريبتات

---

## 2026-03-30 | Sales Report v3.0+ — فلاتر + وضع المرشد

**ما تم إنجازه:**
- 8 فلاتر متتالية: Status, Country, Airline, Supplier, Package, Residence, DepCity, Guide
- فلتر الإقامة مرتبط بـ ArrivalFlightNumber (وليس PackageId)
- فلتر المرشد من شيت Guide Rabih عمود P
- بطاقات KPI إضافية (عدد حجاج الإقامة + المرشد)
- تصدير Excel: 4 شيتات + شيت لكل PNR بتنسيق شركات الطيران (Family/First Name, Passport, DOB, Gender, Expiry)
- بيانات الحجاج من Guide Rabih (firstName, familyName, dob, passportExpiry, guide)
- شيت Packages يعكس عدد الحجاج الفعلي حسب الفلاتر

**وضع المرشد (Guide Mode) — قيد التطوير:**
- Backend جاهز: `getGuideReport(guideName)` — يقرأ Guide Rabih + رحلة الحاج (48 عمود) + الباقات
- Frontend: guideModeArea + KPIs + جدول حجاج — يعمل لكن **يظهر 0 حجاج** (يحتاج debug)
- الأسماء المركبة ("أشرف حبيب - دار السلام") تُعرض كما هي في الفلتر، والبحث جزئي في getGuideReport
- **لم يُنفذ**: تصدير Excel في وضع المرشد

**آخر deployment:** @30 — `AKfycbwICN-v2Gvlr9oyFae8ahlPmKSKvpsNEeFdkAap9RW1vfLr9maiWgMlEeaZL26E5coQ`

**الخطوة التالية:**
1. تصحيح وضع المرشد (0 حجاج) — debug getGuideReport + renderGuideReport
2. تصدير Excel في وضع المرشد
3. اختبار شامل

---

## 2026-03-28 | ترحيل مصدر المرشدين إلى Guide Rabih
**ما تم:**
- تحليل شيت "Guide Rabih" الجديد: 6535 صف، 33 عمود (نفس هيكل Presonal Details)، عمود P = اسم المرشد (5565 حاج لديهم مرشد، 88 مرشد فريد)
- مقارنة البيانات: تطابق 99.9% مع Tour Guide (فرق واحد فقط = مسافة زائدة)
- تحديث 7 مشاريع لقراءة المرشدين من "Guide Rabih" بدلاً من "Tour Guide":
  1. **Hotel Management** — Hotelapp.js (SHEETS + GUIDE_COLS) + Hotelapp helpers.js (buildGuideMap_)
  2. **IkramHajjBot** — DataFetcher.js (getTourGuide_: أرقام أعمدة جديدة)
  3. **Ikram/Minacamp.js** — SHEETS.GUIDE + G columns + buildMinaGuideMap_
  4. **Ikram/Roomtype builder.js** — PD_SHEET من Presonal Details إلى Guide Rabih
  5. **Pilgrim App** — TOUR_GUIDE_SHEET + getGuidePackageStats + buildGuideMap
  6. **Guide App** — Code.js (CONFIG.SHEETS + GUIDE columns + generateGuideLinks + getGuidePassports_) + generateGuideLinks.js
  7. **SheetData** — analyze_guides.js + analyze_pkg_guides.js (مصدر JSON + أسماء الحقول)
- **لم يُعدَّل JourneyMerger** — إضافة عمود مرشد ستزيح بيانات Reception Airport (العمود 49+)
- **Mina Camp Search لا يحتاج تعديل** — يقرأ من "مخيم مني" الذي يُكتب بواسطة Minacamp.js (مُحدَّث)
- تحديث CLAUDE.md: إضافة Guide Rabih للشيتات + تحديث خريطة التفاعل + آخر نقطة عمل

**الملفات المتأثرة:**
- `Projects/Hotel Management/Hotelapp.js` + `Hotelapp helpers.js`
- `Projects/IkramHajjBot/DataFetcher.js`
- `Projects/Ikram/Minacamp.js` + `Roomtype builder.js`
- `Projects/Pilgrim App/PilgrimApp.js`
- `Projects/Guide App/Code.js` + `generateGuideLinks.js`
- `Projects/SheetData/analyze_guides.js` + `analyze_pkg_guides.js`
- `CLAUDE.md` + `SESSION_LOG.md`

**الخطوة التالية:**
- `clasp push` للمشاريع: Hotel Management, IkramHajjBot, Ikram, Pilgrim App, Guide App
- اختبار يدوي لكل تطبيق: التحقق أن أسماء المرشدين تظهر بشكل صحيح

---

## 2026-03-26 | جهاز المكتب
**ما تم:**
- ربط Git على جهاز المكتب بـ GitHub remote (mahmoudekramhajj/ekram-aldyf)
- سحب جميع التحديثات من جهاز المنزل (4 commits)
- إضافة تعليمة استخدام المهارات (Skills) تلقائياً في CLAUDE.md
- تأكيد: المشاريع الصغيرة كلها مطلوبة — لا دمج حالياً
- تأكيد: CLAUDE.md واحد يخدم كل التطبيقات (مشروع متكامل)

**الملفات المتأثرة:**
- CLAUDE.md (محدّث — تعليمة المهارات + آخر نقطة عمل)
- SESSION_LOG.md (محدّث)

**الخطوة التالية:**
- تشغيل setupDailyNotificationTrigger لـ IkramHajjBot
- اختبار شامل للبوت
- إضافة DRIVE_FOLDER_ID في Config.js

---

## 2026-03-26 | جهاز البيت
**ما تم:**
- تحليل الشكل الجديد لبيانات نسك: "رحلة الحاج 2" (42 عمود) + "Presonal Details" (33 عمود)
- اكتشاف الأعمدة المحذوفة (Name, Passport, Email, Group Number, NationalityAr, CountryResidenceAr)
- اكتشاف مفتاح الربط: ApplicationId ↔ رقم المجموعة + مطابقة الأفراد بـ Gender+IsMain (نجاح 100%)
- تأكيد أن ApplicantId القديم = الرقم التسلسلي في Presonal Details (تطابق 100%)
- إنشاء JourneyMerger.js — سكريبت دمج يُنتج نفس الشكل القديم (48 عمود) بدون تعديل أي تطبيق آخر
- اختبار محلي: 46-47/48 عمود متطابقة (الفرق فقط في الاسم الأوسط المفقود من PD)
- رفع السكريبت عبر clasp push (16 ملف)
- توثيق شامل: هيكل Presonal Details + رحلة الحاج 2 + آلية الدمج + خريطة الأعمدة في CLAUDE.md

**الملفات المتأثرة:**
- `Projects/Ikram/JourneyMerger.js` (جديد — سكريبت الدمج)
- `CLAUDE.md` (توثيق الشيتات الجديدة + آلية الدمج + تحديث خريطة التفاعل + آخر نقطة عمل)
- `SESSION_LOG.md` (محدّث)

**الخطوة التالية:**
- تشغيل `mergeJourneyData()` يدوياً من Google Apps Script (أول تشغيل + موافقة صلاحيات)
- التحقق من شيت "رحلة الحاج " بعد الدمج
- إعداد Trigger كل 15 دقيقة (اختياري)
- إضافة استدعاء `onOpen_JourneyMerger()` في دالة onOpen الموجودة (اختياري)

---

## 2026-03-17 | جهاز البيت
**ما تم:**
- مراجعة حالة المشروع واستئناف من جلسة 2026-03-15
- حفظ رؤية 2027 في ذاكرة Claude (نظام متكامل من التعاقد إلى إقفال الحسابات)
- تحديث فهرس الذاكرة MEMORY.md

**الملفات المتأثرة:**
- ذاكرة Claude: project_ikram_vision.md (جديد) + MEMORY.md (محدّث)

**الخطوة التالية:**
- إعداد Git على جهاز المكتب (git clone)
- إنشاء CLAUDE.md لـ Ikram و Hotel Management
- نقاش دمج المشاريع الصغيرة (TourGuide Manager، GDS، B2C Sync، PNR Sync)

---

## 2026-03-15 | جهاز البيت
**ما تم:**
- إنشاء repo على GitHub (private): `mahmoudekramhajj/ekram-aldyf`
- تثبيت GitHub CLI + تسجيل الدخول
- حذف Ikram1 (محذوف من Apps Script) + إعادة تسمية Ikram2 → Ikram
- مراجعة شاملة لـ IkramAgent: كود كامل لكن غير مفعّل — مؤجّل للعام القادم (مرحلة بيع الباقات انتهت)
- مقارنة IkramAgent (قبل الحجز) vs IkramHajjBot (بعد الحجز) — لا تكرار بينهم
- إنشاء خريطة تفاعل المشاريع الكاملة وإضافتها في CLAUDE.md

**الملفات المتأثرة:**
- CLAUDE.md (خريطة التفاعل + تحديث جدول التطبيقات + حذف Ikram1/Ikram2)
- SESSION_LOG.md (محدّث)
- Projects/Ikram1/ (محذوف)
- Projects/Ikram2/ → Projects/Ikram/ (أعيد تسميته)

**الخطوة التالية:**
- تنفيذ إعداد Git على جهاز المكتب (git clone)
- إنشاء CLAUDE.md لـ Ikram و Hotel Management
- نقاش دمج المشاريع الصغيرة

---

## 2026-03-14 | جهاز البيت
**ما تم:**
- مراجعة شاملة لتحديثات المكتب (clasp pull لـ 17 مشروع + CLAUDE.md لـ Airport Search)
- تحديث ملفات ذاكرة Claude
- إعداد Git: نسخ المشروع إلى `C:\Users\mubar\Ekram-Aldyf\` + أول commit (120 ملف)
- إنشاء SESSION_LOG.md لتسهيل الاستمرارية بين الأجهزة

**الملفات المتأثرة:**
- .gitignore (جديد)
- SESSION_LOG.md (جديد)
- CLAUDE.md (محدّث)

**الخطوة التالية:**
- إنشاء repo على GitHub وربطه
- تنفيذ نفس الإعداد على جهاز المكتب (git clone)
- تحديد المطلوب من أعمدة عقود السكن/التنقل في التطبيقات
