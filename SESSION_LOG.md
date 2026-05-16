# سجل الجلسات — Ekram Aldyf

> **تعليمة لـ Claude:** في نهاية كل محادثة، أضف جلسة جديدة في الأعلى (بعد هذا السطر مباشرة). عند بدء محادثة جديدة بكلمة "تابع"، اقرأ آخر جلسة وأكمل من حيث توقفنا.

---

## 2026-05-16 | فرز السكن/التنقل كرونولوجياً — إصلاح Makkah-first

**البلاغ:** الحاج `19AI08726` رأى تضارباً بين خطة التنقل وخطة السكن.

**ما اكتُشف:**
- البيانات الفعلية للحاج (Railway + B2C_v2): يصل MED بـ VF205 يوم 15 مايو، يقيم في Dar Almaqam/مكة (15 مايو–3 يونيو) ثم Concorde/المدينة (3–6 يونيو)، يغادر من JED بـ VF190 يوم 6 يونيو. خطته الفعلية **مكة-أولاً ثم المدينة**.
- شيت "تعديل المرشدين" يضعه H1=Concorde (المدينة) و H2=Dar Almaqam (مكة) — الترتيب غير كرونولوجي.
- `HajjBotServer/src/feature-hotel.js:buildPilgrimHotelList` كان يثق بترتيب الأعمدة H1→H2→H3 ولا يفرز. `feature-transport.js` يستهلك نفس القائمة → الخلل مزدوج (سكن + تنقّل).
- جلسة 2026-05-13 افترضت "ترتيب الشيت كرونولوجي بطبيعته" — هذه الفرضية **خاطئة**.

**النطاق:** فُحص الشيت كاملاً → **1,144 من 6,793 حاجاً (≈17%)** لديهم `H2.checkIn < H1.checkIn`. حالة عامة، ليست خاصة بهذا الحاج. باقات متكررة متأثرة: Concorde+Park Plaza، Concorde+Biak، Al Shakreen+EMAAR AL RAWDA 2، Concorde+Dar Almaqam.

**الإصلاح:**
- `commit aa112cb` (مدفوع لـ main، Railway نشره تلقائياً).
- في `feature-hotel.js` قبل Stage A: فرز `rawSlots` بـ `toIsoDate_(checkIn)`. الفنادق بلا تاريخ سليم → نهاية الترتيب (Infinity).
- إصلاح واحد يغطي السكن والتنقل (مصدر موحَّد).

**التحقّق بعد deploy:**
- `trace-pilgrim?passport=19AI08726`:
  - السكن 1 = مكة (المقام) 15 مايو→3 يونيو ✓
  - السكن 2 = المدينة (كونكورد) 3→6 يونيو ✓
  - التنقل: وصول→المقام، انتقال مكة→المدينة في 3 يونيو، مغادرة كونكورد→JED ✓
- Case 1 تطبَّق صحيحاً: 14-May→15-May (تأخير الدخول للوصول الفعلي).

**ذاكرة:**
- `project_hotel_chronological_sort.md` — جديد.
- `MEMORY.md` — index مُحدَّث.

**ما لم يُلمَس (خارج النطاق):**
- HajjBotServer scripts غير المُتعقَّبة (10 ملفات تحليل) + `daily-b2c-arrivals-2026-05-15.js` معدَّل.
- الـ working tree في الجذر يحوي تعديلات في IkramHajjBot/GDS/Guide App/TicketLinker — معلَّقة من جلسات سابقة، تنتظر قراراً منفصلاً.

**القاعدة المستفادة:** لا تثق بترتيب أعمدة H1/H2/H3 — افرز دائماً بـ `checkIn`.

---

## 2026-04-23 | ClaudeBrain M1 — إنشاء المخ المركزي (Orchestrator)

**ما تم إنجازه:**
- تصميم كامل لـ ClaudeBrain: Orchestrator مركزي لكل عمليات الأتمتة في إكرام الضيف
- قرارات معمارية مُعتمدة: خليط ذكي للنماذج (Haiku/Sonnet/Opus) + 6 شخصيات + الفقيه مستشار داخلي
- الأسئلة الشرعية تحويل 100% لـ (مرشد + 1966) — المخ لا يُفتي
- أول تطبيق مستهدف للربط: FlightChangeProcessor
- مفتاح Claude API مخصّص للمخ (منفصل عن باقي المشاريع)

**M1 (الهيكل الفارغ) — مكتمل:**
- `Projects/ClaudeBrain/` مشروع GAS جديد
- Script ID: `1ht-aJiQZE3JwgqUx563pbUykNLe--Xwfms2a-E1KhwEa8gtFnWReH98_`
- Deployment (M1 Clean): `AKfycbzi2Vft_DeQMSkVQGKBvbzmzO5yhjdnat498L5Dnsg2S4RPFKIyQCcqhhf5Rbobcjko`
- Web App URL: https://script.google.com/macros/s/AKfycbzi2Vft_DeQMSkVQGKBvbzmzO5yhjdnat498L5Dnsg2S4RPFKIyQCcqhhf5Rbobcjko/exec

**الملفات (8):**
- `Main.js` — doPost + doGet + event router
- `Config.js` — قراءة Script Properties + أسعار النماذج + حساب التكلفة
- `Claude.js` — استدعاء Anthropic API + retry exponential على 529/503/429
- `Auth.js` — HMAC-SHA256 verification (constant-time compare)
- `Budget.js` — حد يومي $2 + ساعي $0.5 مع تتبّع في Script Properties
- `Identity.js` — الهوية الأساسية (system prompt)
- `ClaudeAPI.js` — التحكم عن بُعد (ping/list/run)
- `_Setup.js` — فارغ (كان يحتوي المفاتيح مؤقتاً ثم نُظّف)

**الاختبارات الناجحة:**
- ✅ GET / → رد ترحيبي
- ✅ GET ?action=ping → ClaudeAPI يعمل
- ✅ POST echo مع HMAC صحيح → رد JSON
- ✅ POST ping → Claude Haiku ردّ بالعربية (1.87s، 39+31 tokens)
- ✅ POST مع HMAC خاطئ → 401
- ✅ POST بلا X-Signature → 401

**ما ليس موجوداً بعد (لاحقاً في M2+):**
- ❌ Tools (قراءة شيت، تيليغرام، إلخ)
- ❌ Intents (FlightChange، TicketHandler، إلخ)
- ❌ نظام الموافقة البشرية
- ❌ ربط FlightChangeProcessor

**الخطوة التالية:**
- M2: أول intent + أول tool (readPilgrim)
- بعدها: Tools قراءة للشيت (flight، package، transport)
- ثم: intent flight_change + ربط FlightChangeProcessor

**⚠️ عمل أمني لم يكتمل:**
- المفتاح `sk-ant-api03-_j74n_bO8...` ظهر في سجل المحادثة → يجب تدويره: احذفه من console.anthropic.com وأنشئ جديداً ثم استبدل `CLAUDE_API_KEY` في Script Properties

---

## 2026-04-08 | تصميم Infographic إحصائيات الحجاج حسب الجنسية

**ما تم إنجازه:**
- تصميم infographic احترافي لعرض تنفيذي: إحصائيات الحجاج حسب الجنسية (6,612 حاج، 78 دولة، 60 باقة)
- التصميم بالهوية البصرية لإكرام الضيف (ألوان + أنماط هندسية إسلامية)
- المنهجية: HTML/CSS + خطوط Tajawal + أعلام الدول (flagcdn.com) + Chrome screenshot → PNG 1920×1080
- مصدر البيانات: `AA.csv` (شيت Presonal Details) + `pilgrim_stats_v8.html`
- المحاولة الأولى كانت بـ Pillow/Python — نتائج ضعيفة (نص عربي مكسّر + 60% مساحة فارغة)
- المحاولة الثانية بـ HTML/CSS أعطت نتائج أفضل بكثير

**الملفات المنتجة:**
- `Downloads/pilgrim_stats_infographic.png` — الناتج النهائي (لا يُرفع لـ Git)

**الخطوة التالية:**
- تحسين التصميم حسب ملاحظات المسؤولين

---

## 2026-04-09 | Transport Management — تطبيق إدارة النقل v2.3

**ما تم إنجازه:**
- تصميم شامل لتطبيق إدارة النقل عبر جلسة أسئلة وأجوبة مفصلة
- تحديد 28 عملية نقل كاملة مع مصادر البيانات والهوامش الزمنية لكل حالة
- إنشاء مشروع GAS جديد (Transport Management) مع 10 ملفات
- إنشاء 6 شيتات جديدة في Ikram Abuown + شيت Airline_Terminals (21 شركة طيران)
- إضافة عمودي عقد السكن وعقد النقل
- نشر عدة إصدارات حتى v2.3

**مراجعة الـ 28 حالة (مكتملة):**
- الحالات 1-5: وصول المطارات — منطق تحديد الوجهة بمقارنة تاريخ الوصول مع تواريخ الفنادق
- الحالات 6-11: نقل بين المدن (حافلة 12:00 / قطار 14:00)
- الحالة 12: تحويل فنادق مكة (Shifting) — ترتيب بالتواريخ
- الحالات 13-21: عمليات الحج (منى/عرفة/مزدلفة) والعودة
- الحالات 22-28: مغادرة للمطارات (8س/3س/12س) ومحطات القطار

**قواعد مهمة:**
- الحافلة تتواجد ساعة قبل الانطلاق
- B2C: مسؤولية تحالف النقل
- الصالة من شيت Airline_Terminals

**ملاحظات UI للتنفيذ:**
- "قبل"/"بعد" بدل الأسهم + "إلى" بدل → + أرقام لاتينية
- فلتر الموقع: مدينة كاملة (افتراضي) + فلتر فرعي
- الهوية البصرية: (#476831, #232E64, #9B6FB1, #B3C79D) + شعار إكرام

**آخر deployment:** @15 — v2.3

**الخطوة التالية:**
1. إعادة كتابة محرك التخطيط بناءً على المراجعة التفصيلية
2. تطبيق الهوية البصرية + إصلاح UI
3. اختبار شامل ونشر

---

## 2026-04-08 | جهاز البيت
**المشروع:** PilgrimStats — شاشات عرض إحصائيات الحجاج على شاشات كبيرة

**ما تم:**
- إنشاء مشروع GAS جديد `PilgrimStats` مع `clasp create` + نشر كـ Web App (ANYONE_ANONYMOUS)
- بناء Backend (`Code.js`) يقرأ من شيت Presonal Details ويحسب:
  - إجمالي + أفراد vs مجاملة (عمود D) + B2B/B2C (عمود U) + 3 مخيمات (عمود AB) + توزيع بلد الإقامة (عمود Q) لكل فئة — الكل مع ذكور/إناث
  - كاش 5 دقائق
- بناء شاشات:
  - **KpiView** — إجمالي كبير + حلقة جنس + 7 بطاقات بأشرطة نسبية
  - **B2cCourtesyView** → "الأفراد والمجاملة" — شريط أفقي لا نهائي من بطاقات الدول + أعلام (flagcdn.com) + ترقيم ذهبي
  - **CountriesView / NationalitiesView / CampsView / IndexView**
- الهوية البصرية: خلفية كريمية + أخضر (#476831) + بنفسجي (#9B6FB1) + نيلي (#232E64) + ذهبي (#B08D3A) + شعار حقيقي
- قاموس أعلام 57 دولة + Fallback بفهرس مباشر للأعمدة

**ملفات:** `Projects/PilgrimStats/` (Code.js, KpiView.html, B2cCourtesyView.html, IndexView.html, CountriesView.html, NationalitiesView.html, CampsView.html, Styles.html)
**Deployment:** `AKfycbx6hAhMvjPBTa40qQgN4kUPwc1eN1FXJgRFK9fl0HeaOvfViMxa_aTmtzCPGTbJnKyFqw`
**Script ID:** `1GaNM_2vPTp4KNMAY9tmm5Z5pv38ucviz52mYZkVYiZeL7hA6fO5LiV3F`

**الخطوة التالية:**
- تحسين تصميم الشاشات حسب الملاحظات
- اختبار على شاشات كبيرة فعلية

---

## 2026-04-08 | Ikram — مراجعة الهوية البصرية + تحسينات الداشبورد

**ما تم إنجازه:**
- تطبيق الهوية البصرية على تطبيق Ikram (الباقات)
- تغيير الهيدر: حذف الاسم العربي/الإنجليزي والأيقونة، واستبدالها بشعار إكرام الضيف الكامل (شفاف بدون مربع أبيض)
- رابط الشعار: `https://image2url.com/r2/default/images/1775443354621-9813e5f5-b4a7-4d23-8a0c-131fbfe58ec7.png`
- إضافة إحصائية "Total Pilgrims" (إجمالي الحجاج) من شيت Presonal Details — تظهر قبل Tickets
- تغيير حساب المتبقي ليكون: `7000 - إجمالي الحجاج`
- تصحيح عنوان الصفحة إلى "الباقات / اكرام الضيف"
- إصلاح خطأ meta tag (description غير مسموح في GAS)
- إصلاح خطأ updateUILanguage (headerTitle/headerSubtitle محذوفة)

**الملفات المتأثرة:**
- `Projects/Ikram/Index.html` — هيدر جديد بالشعار + كارت Total Pilgrims
- `Projects/Ikram/Styles.html` — تحديث ألوان الهوية البصرية
- `Projects/Ikram/Scripts_Core.html` — إزالة مراجع headerTitle/headerSubtitle + إضافة getTotalPilgrims
- `Projects/Ikram/SearchApp.js` — عنوان الصفحة + دالة getTotalPilgrims من Presonal Details

**آخر deployment:** @69 — `AKfycbxVLLuk8yxPHOSmXSqxfnYN7wg0YmF4NIb6ItY69VJQdWaOeMblq0nj7jzk9ayxCXtc7w`

**الخطوة التالية:**
1. مراجعة الهوية البصرية في باقي التطبيقات
2. اختبار شامل

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

---

## 2026-05-16 — Closure جماعي على 239 فترة (Guide Admin Server)

**الحالة:** متوقّف يدوياً (أمر "stop"). **66/239 مكتمل، 173 متبقي.** لا فشل حقيقي.

### السياق
بعد دمج المستخدم لتسميات Al Shakreen/Concorde، توليد قائمة حية من 239 فترة، تشغيل `/api/admin/execute-closure` لكل واحدة لتطبيق pool-based per-stay repack.

### إنجازات الجلسة
1. تشخيص أن `curl -u` يفشل أحياناً بـ 401 على Windows (SSL renegotiation schannel) → تحوّل لـ `Authorization: Basic <b64>` صراحةً.
2. تشخيص أن "لا توجد بيانات للإرسال" + `forceDeleted ≥ 1` = نجاح متوقّع لتفريغ مرشدين يتامى (ليس فشلاً).
3. سكريبت v2 نظيف: `C:/tmp/run_closure_v2.sh` مع retry + سجل واحد.
4. أدوات تحليل: `analyze_results.js` و `build_done_set.js`.
5. ASSIGNMENT_CLOSED تم تعطيله ويبقى OFF.

### الاستئناف
1. تحقّق فلاج: `curl '.../exec?closureStatus=1&key=ekram2026claude'`.
2. `node C:/tmp/build_done_set.js` لتجديد todo (يلتقط الـ66 المنجزة).
3. `bash C:/tmp/run_closure_v2.sh` في background.
4. عند الاكتمال: `analyze_results.js` → 0 فشل → فعّل `?closure=on`.

### ملفات مرجع
- `C:/tmp/run_closure_v2.sh`
- `C:/tmp/closure_results/v2_*.json` (نتائج)
- `C:/tmp/closure_log_v2.txt` (سجل)
- ذاكرة جديدة: `project_no_empty_bed_closure.md`

