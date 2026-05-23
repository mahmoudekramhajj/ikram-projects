# سجل الجلسات — Ekram Aldyf

> **تعليمة لـ Claude:** في نهاية كل محادثة، أضف جلسة جديدة في الأعلى (بعد هذا السطر مباشرة). عند بدء محادثة جديدة بكلمة "تابع"، اقرأ آخر جلسة وأكمل من حيث توقفنا.

---

## 2026-05-23 | HajjBotServer — إضافة سطر "رقم الخيمة" لقائمة المخيمات

**الطلب:** إضافة "رقم الخيمة" لعرض المخيمات في بوت الحاج (مخيم منى ومخيم عرفة) ليكون مرئياً بين رقم المخيم ورقم الصوفابيد.

**التحقق قبل التعديل:** قراءة `src/feature-camps.js` — وجدت 3 أسطر فقط (الاسم/الموقع، رقم المخيم، رقم الصوفابيد) في كلا الدالتين `handleArafahCamp` و `handleMinaCamp`. لا حقل بيانات لرقم الخيمة في PD.

**التعديل:** سطر واحد مُضاف في كل دالة:
```js
text += '🏕️ <b>رقم الخيمة:</b> <i>قيد الإعداد</i>\n';
```

**النتيجة (4 أسطر بالترتيب):**
1. 📍 الاسم/الموقع (مخيم منى: `PD.CAMP` col 27، عرفة: ثابت)
2. 🔢 رقم المخيم — `قيد الإعداد`
3. 🏕️ رقم الخيمة — `قيد الإعداد` (جديد)
4. 🛏️ رقم الصوفابيد — `قيد الإعداد`

**النشر:**
- HajjBotServer submodule: commit `5fa1a48` على `main` → push GitHub ✓
- Fly deploy: `hajjbot-standby` → machine `08072e6c597748` started ✓
- تحقُّق: `curl https://hajjbot-standby.fly.dev/` → HTTP 200، uptime 22ث، tenants.ikram.exists=true ✓

**ملفات:** `Projects/HajjBotServer/src/feature-camps.js` (سطران مُضافان في دالتين).

**معلَّق:** عند توفّر مصدر بيانات حقيقي لأرقام المخيم/الخيمة/الصوفابيد (شيت أو حقول PD جديدة)، استبدل الـ placeholders `قيد الإعداد` بقيم ديناميكية.

**ذاكرة:** [project_camps_tent_number_line.md](.claude/projects/C--Users-mubar-Ekram-Aldyf/memory/project_camps_tent_number_line.md)

---

## 2026-05-22 (09:30 PM) | تذاكر قطار 23 مايو — name-trust + رفع 92 تذكرة

**الطلب:** معالجة مجلد `C:\Users\mubar\Downloads\قطار 23 مايو` (تذاكر قطار نسك ليوم 2026-05-23) بسكريبت `name-trust-train-tickets.js`.

**المُدخَلات:** 2 PDF multi-pax: `3_05232026_MAD_3DFAC24F4.pdf` (53 صفحة) + `53_05232026_MAD_13CA24E18.pdf` (41 صفحة) = 94 صفحة.

**التشغيل:** `node scripts/name-trust-train-tickets.js --src "..." --split --dry-run --concurrency=5` → Gemini 29.4ث.

**النتائج (94 صفحة):**
- 91 مطابقة (name=passport ok)
- **1 مُصحَّحة بالاسم:** `AMNA OSMAN` — نسك طبعت `A1I911069` (حرف I بدل رقم 1) → السكريبت صحّحها لـ `A11911069` عبر مطابقة الاسم في PD (سبب وجود السكريبت أصلاً)
- 0 غامضة، 0 مفقودة، 2 صفحة غلاف، 0 أخطاء Gemini
- → **92 جاهزة في `renamed_by_name/`**

**الرفع:** المستخدم رفع الـ92 ملف يدوياً إلى `TRAIN_FOLDER` Drive (`1A_gihLxAr2nV9RI3DXJhQib13cF-1ucN`) — معتاد لأن Service Account لا quota في My Drive (راجع `project_train_tickets_audit_fix.md`).

**ما بعد الرفع:**
- `POST /api/admin/reload` على `hajjbot-standby.fly.dev` → ok=true، B2C_v2=6792، PD=6794
- `train-watcher/status` قبل الرفع: snapshotSize=2920، intervalMs=180000 — المراقب سيلتقط الـ92 الجديدة تلقائياً في الدورة التالية وينفّذ `clearTrainCache()` (راجع `project_train_ticket_watcher.md`).

**التقرير:** `C:\Users\mubar\Downloads\قطار 23 مايو\name_trust_report.csv` محفوظ.

**لا تعديلات كود.** صفر مشاكل، 0 معلَّقات.

---

## 2026-05-21 (08:45 AM) | SheetData — تحديث دوري تلقائي كل 15 دقيقة

**الطلب:** المستخدم أراد تحديث دوري دائم للنسخ المحلية من شيتات Google لسرعة اتخاذ القرار.

**الوضع السابق:** `D:\اكرام الضيف\Ekram Aldyf\Projects\SheetData\refresh.js` (Node.js + Google Service Account readonly) يُصدّر 97 شيت كـ JSON. آخر تحديث يدوي قبل الجلسة: 08:28 صباحاً.

**المكوّنات المُضافة:**
1. **Junction**: `C:\SheetData` → `D:\اكرام الضيف\Ekram Aldyf\Projects\SheetData` (تجاوز مشكلة PowerShell 5.1 + UTF-8 بدون BOM في المسار العربي). أُنشئ عبر `C:\Users\mubar\mklink-sheet.bat` (chcp 65001).
2. **Wrapper**: `C:\Users\mubar\refresh-sheets.ps1` يستدعي `node refresh.js` ويكتب log مقصوص لـ 200 سطر.
3. **Scheduled Task**: `EkramAldyf-RefreshSheets` كل 15 دقيقة، RunLevel=LIMITED (لا باسورد). التشغيل التالي: 9:02 AM.

**تحقق:** تشغيل تجريبي = exit=0، `_index.json` تحدّث، 97 شيت كاملة.

**قيد:** لا يعمل وجهازك نائم — يستأنف بعد الاستيقاظ. للحظات حرجة شغّل يدوياً: `schtasks //Run //TN "EkramAldyf-RefreshSheets"`.

**ذاكرة:** [project_sheetdata_auto_refresh.md](.claude/projects/.../memory/project_sheetdata_auto_refresh.md) — تفاصيل البنية + الأوامر الإدارية + الدروس.

**ملفات جديدة (خارج repo):** `C:\Users\mubar\refresh-sheets.ps1` + `C:\Users\mubar\mklink-sheet.bat` + `C:\SheetData` junction. صفر تعديل كود في المشروع.

---

## 2026-05-20 (09:30 PM) | HajjBotServer — إصلاح رابط "قائمة الحجاج القادمين" في تنبيهات الوصول

**البلاغ:** الضغط على زر "قائمة الحجاج القادمين 📋" في تنبيه SM493 (T-1h) فتح `https://hajjbot-standby.fly.dev/arrivals?flight=SM493&date=2026-05-20` → "غير مصرّح بالوصول، أضف ?key=YOUR_TOKEN".

**السبب الجذري:** في `src/arrival-alerts.js:685` بناء الرابط لا يضيف `key`، لكن endpoint `/arrivals` في `server.js:69` محمي بـ `DASH_TOKEN`. كان الـbug صامتاً على Railway لأن `BASE_URL` كان فارغاً → الزر لا يظهر أصلاً. بعد نقل البوت لـ Fly في جلسة 05:00 PM وضبط `ARRIVAL_ALERTS_BASE_URL=https://hajjbot-standby.fly.dev`، الزر بدأ يظهر بدون مفتاح.

**الإصلاح (commit التالي):**
1. `src/arrival-alerts.js:96` — إضافة `const DASH_TOKEN = process.env.DASH_TOKEN || process.env.SYNC_TOKEN || '';`
2. `src/arrival-alerts.js:685` — إلحاق `&key=${encodeURIComponent(DASH_TOKEN)}` عند توفّر التوكن.

**التحقق:**
- `flyctl deploy -a hajjbot-standby` ← machine started, DNS verified ✓
- `curl /arrivals?flight=SM493&date=2026-05-20&key=<DASH_TOKEN>` → HTTP 200 ✓
- التنبيه القديم لـ SM493 المُرسَل قبل deploy لن يعمل (يحتاج المستخدم لصق المفتاح يدوياً). كل التنبيهات القادمة ستحوي المفتاح تلقائياً.

**ملفات معدّلة:** `Projects/HajjBotServer/src/arrival-alerts.js` (سطران).

---

## 2026-05-20 (05:00 PM) | HajjBotServer — مراجعة مهنية مستقلة لنقل Fly + إكمال النواقص

**السياق:** المستخدم طلب مراجعة كاملة مستقلة لعملية نقل البوت من Railway إلى Fly. الجلسات السابقة وثّقت أن Fly = primary، لكن لم يُتحقَّق من مكتمليّة النقل.

**الفحص الحي كشف:**
- Fly فيه 19 secret فقط، Railway فيه 39 (الفرق 20).
- `BASE_URL` غير مضبوط على Fly → عند restart يتخطّى webhook registration.
- `field` tenant غير محمَّل على Fly (TENANT_2_* مفقودة) — بوت الميدان @ekram_field_bot كان معطَّلاً على Fly.
- `REDIS_URL` مفقود على Fly → sessions تُفقد كل restart + log spam من analytics (incr/sadd/scard غير محققين في in-memory fallback).
- Disk على Fly ephemeral (لا `[mounts]` في `fly.toml`) → ticket-cache.json + arrival-alerts-state.json تُفقد كل restart.
- Railway فعلياً مضبوط `STANDBY_MODE=true` من قبل + لا `STANDBY_PUBLIC_URL` → صامت فعلاً (لا schedulers، لا setWebhook). race condition المخوَّف لا يحدث.
- `HajjBotServer/CLAUDE.md` كان لا يزال يقول "webhook يشير لـ Railway" — معلومة قديمة مُضلِّلة.

**الإجراءات المُنفَّذة:**
1. **نسخ 10 secrets من Railway لـ Fly عبر `flyctl secrets import`:** REDIS_URL، FIELD_BOT_TOKEN (legacy)، FIELD_ROLES_JSON، TENANT_2_*، NUSUK_VIDEO_DRIVE_ID + قيم URL مُعاد توجيهها لـ Fly (BASE_URL، ARRIVAL_ALERTS_BASE_URL = `https://hajjbot-standby.fly.dev`).
2. **بعد restart Fly تلقائياً:** keys_count قفز من 31 إلى 47، BASE_URL_set=true، tenants=[ikram, field] (بعدما كان [ikram] فقط).
3. **اكتشاف عند فحص logs:** `REDIS_URL` المنسوخ من Railway = `redis://default:***@redis.railway.internal:6379` — عنوان داخلي لشبكة Railway فقط، لا يصل من Fly. spam جديد `ENOTFOUND redis.railway.internal`.
4. **`flyctl secrets unset REDIS_URL`** → عاد الـ in-memory fallback. (التبعية: sessions لا تزال تُفقد عند restart، analytics لا يزال يَشِيب logs لكن أقل ضرراً من ENOTFOUND retry loop).
5. **تحقّق webhook:** كلا البوتين (ikram + field) موجَّهين لـ Fly عبر `getWebhookInfo`. Telegram registers تم على Fly عند restart.
6. **حُدِّث `HajjBotServer/CLAUDE.md`** بالكامل: Fly = primary فعلي، Railway = silent standby، روابط جديدة، تحذيرات Sessions/Spam/Disk، runbook التبديل بين الخادمين.
7. **نُظِّفت ملفات tmp** التي تحوي secrets.

**الوضع النهائي (متحقَّق عبر `/`):**
- Fly: uptime ~1د بعد آخر restart، tenants=[ikram, field]، schedulers تعمل (urlWatcher كل 5د، arrivalAlerts، trainTicketWatcher).
- Railway: HTTP 200، uptime 2.2h، STANDBY_MODE=true، tenants.field.exists=false، لا schedulers (مؤكَّد من env vars).

**الثغرات المتبقّية (لم تُحَل، تحتاج قرار المستخدم):**
1. **Redis حقيقي لـ Fly:** Upstash مجاناً يحلّ (sessions + analytics) — يحتاج المستخدم إنشاء حساب.
2. **Fly Volume:** إضافة `[mounts]` في fly.toml + `flyctl volumes create` يحلّ مشكلة ephemeral disk للـ ticket-cache.
3. **In-memory fallback ناقص:** يحتاج توسيع `src/redis.js` بـ incr/sadd/scard/hincrby/hgetall/expire — لو لم نختر Upstash.
4. **`SESSION_LOG` 02:00 PM ادّعى Railway قُتل** لكنه يردّ 200 — إما redeploy تلقائي رجَّعه، أو الادعاء كان خاطئاً. الذاكرة `project_fly_dr_standby.md` تحتاج تحديث (`STANDBY_MODE` لا يزال `true` على Railway فعلاً).

**ملفات معدَّلة:** `Projects/HajjBotServer/CLAUDE.md` (إعادة كتابة كاملة). صفر تعديل كود.

**Secrets المضافة لـ Fly (10):** REDIS_URL (ثم أُلغي), FIELD_BOT_TOKEN, FIELD_ROLES_JSON, TENANT_2_TOKEN, TENANT_2_NAME, TENANT_2_KIND, TENANT_2_ADMIN_IDS, NUSUK_VIDEO_DRIVE_ID, BASE_URL, ARRIVAL_ALERTS_BASE_URL.

---

## 2026-05-20 (02:00 PM) | HajjBotServer — Railway قُتل نهائياً + تصحيح ادعاء "Google Cloud block"

**التطوّر بعد توثيق 13:30:** المستخدم سأل عن سبب "Google Cloud block". الفحص الحي كشف الحقيقة:
- `railway.com` → HTTP 200 ✅ (الذاكرة قالت 525)
- `backboard.railway.com` → HTTP 200 ✅ (الذاكرة قالت SSL error)
- Railway deploys تنفّذ بشكل طبيعي (deployment `17a119d6` من اليوم 11:11)

**النتيجة:** الذاكرة كانت قديمة. "Google Cloud block" كان حقيقياً يوم 19 مايو لكن **رُفع منذ ذلك**. لم يُوثّق السبب الجذري — لم أعرفه ولا المستخدم استلم إيميل تفسير.

**قتل Railway النهائي:**
- `railway scale sfo=0` فشل بـ `Unknown region sfo` (CLI الجديد يستخدم أسماء عامة، لا region IDs)
- `railway scale us-west=0` نجح لكن لم يلمس sfo
- `railway down -y` بعد retry → "No deployments found"
- بعد 30 ثانية: `https://hajjbotserver-production.up.railway.app/` → **HTTP 404 "Application not found"**

**الوضع النهائي:**
- Railway: ميت تماماً (لا instances، لا deployments، لا تطبيق فعّال). env + كود + service ID + volume محفوظة كنسخة احتياطية.
- Fly: primary وحيد. uptime 47د، 6794 صف PD، صفر ازدواج.

**Runbook استيقاظ Railway (محدَّث):** يحتاج redeploy كامل (`railway up` أو deploy عبر CLI/GitHub) — ليس مجرد scale=1. سيستغرق دقائق بدل ثوانٍ.

**ملفات معدّلة:** صفر كود. فقط `railway down` + تحديث ذاكرة.

---

## 2026-05-20 (01:00 PM) | HajjBotServer — UrlWatcher + parity كامل لمتغيرات Fly

**السياق:** التقرير الساعي 1pm من UrlWatcher أظهر 12 دورة / 8 تغييرات / **6 فشل + 2 متخطّى** — كل الفشل بنفس الرسالة: `parse: GEMINI_API_KEY not set`.

**التشخيص الأولي:** `GEMINI_API_KEY` غير موجود في أسرار Fly رغم ذكره في التوثيق.

**الحل المبدئي:** `flyctl secrets set GEMINI_API_KEY=... -a hajjbot-standby` → rolling update نجح → failures cache مُمحاة.

### اكتشاف أكبر: متغيرات كثيرة مفقودة من Fly

بعد إصلاح GEMINI، فحصت كل `process.env.X` في الكود مقابل أسرار Fly. النتيجة: **٦ متغيرات حرجة مفقودة:**
- `TG_GROUP_MED=-4849598886`
- `TG_GROUP_JED_T1=-5220583519`
- `TG_GROUP_JED_NORTH=-5267173490`
- `TG_GROUP_OPS_MAKKAH=-4916619724`
- `TG_GROUP_OPS_MADINAH=-5284394785`
- `NUSUK_VIDEO_FILE_ID=BAACAgQAAxkDAAEC_fl...`

**المفارقة:** هذه المتغيرات كانت "تعمل" طوال فترة تعطّل Railway dashboard. التفسير: **Control plane ≠ Data plane.** Railway dashboard/deploys معطّلة (Google Cloud حظر)، لكن الـ container القديم لا يزال يعمل بـ env المحقونة وقت آخر start ناجح. التنبيهات كانت تخرج من Railway runtime لا من Fly.

### نقل المتغيرات وتفعيل خطة "Fly = أساس، Railway = طوارئ"

1. `railway login` (المستخدم) ← `railway variables --kv` → استخراج القيم
2. `flyctl secrets set TG_GROUP_MED=... TG_GROUP_JED_T1=... ... NUSUK_VIDEO_FILE_ID=... -a hajjbot-standby` (دفعة واحدة، rolling restart واحد)
3. على Railway: `railway variables --set "STANDBY_MODE=true" --skip-deploys` ← `railway scale sfo=0` ← **Queued (Google Cloud block يمنع التنفيذ)**

### الوضع النهائي

- **Fly:** مستقل 100% (13 secret أصلية + 7 جديدة = 20 secret). جميع schedulers عليه (`STANDBY_MODE=false`).
- **Railway:** runtime الـ container القديم لا يزال شغّال (لا نستطيع قتله الآن). env محدّث بـ `STANDBY_MODE=true` + `scale=0` معلّقان في الطابور، سيُنفَّذان تلقائياً عند رفع Google Cloud الحظر.
- **ازدواج مؤقت:** urlWatcher idempotent (ضرر صفر). arrival-alerts + train-ticket-watcher قد ترسل تنبيهات مكرّرة لساعات/أيام حتى يموت الـ container القديم تلقائياً.
- **ALERT_CHAT_ID** غير موجود حتى على Railway → fallback لـ `ADMIN_IDS[0]` يعمل في الجانبين.

### Runbook استيقاظ Railway (لو Fly سقط مستقبلاً)

```bash
# 1. تحويل Railway من standby لـ primary (env)
cd C:/Users/mubar/Ekram-Aldyf/Projects/HajjBotServer
railway variables --set "STANDBY_MODE=false" --skip-deploys

# 2. إعادة الـ container للحياة (لو سبق ومات بـ scale=0)
railway scale sfo=1

# 3. تحويل webhook لـ Railway
curl "https://api.telegram.org/bot$BOT_TOKEN/setWebhook?url=https://hajjbotserver-production.up.railway.app/webhook/$BOT_TOKEN"

# 4. (اختياري) تعطيل Fly مؤقتاً
flyctl secrets set STANDBY_MODE=true -a hajjbot-standby
```

**ملفات معدّلة:** صفر كود. فقط:
- `flyctl secrets set` ×2 (GEMINI + الستة)
- `railway variables --set STANDBY_MODE=true --skip-deploys`
- `railway scale sfo=0` (معلّق)
- تحديث `project_fly_dr_standby.md` + `MEMORY.md` (سطر الفهرس) + هذه الجلسة

**المعلَّق:**
- ازدواج تنبيهات مؤقت لساعات/أيام (مقبول، lid على نفسه عند رفع Google Cloud)
- التقرير الساعي 2pm: التحقق أن urlWatcher عاد طبيعي + متابعة الـ6 جوازات (`21AH28829`, `PA0926857`, `A31417381`, `PAX874893` ×2, `23AR29327`)
- تدوير `GEMINI_API_KEY` لاحقاً (ظهر بنص واضح في المحادثة)

---

## 2026-05-20 (06:00 AM) | HajjBotServer — Fly صار Primary كامل + حادثة 426 تذكرة قطار

**السياق:** بعد التصلّب الأمني، رفع المستخدم 426 تذكرة قطار جديدة في 05:42 AM. وصل تنبيه TrainTicketWatcher (من Railway). لكن البوت على Fly كان يقول "تذكرة القطار غير متوفرة" حتى لجوازات موجودة في القائمة (مثلاً `A09301911`).

**التشخيص:**
- اختبار مباشر على Fly: `trainTicketExists('A09301911') => true` (Drive يحوي الملف).
- المشكلة: `TrainTicketWatcher` معطّل على Fly (STANDBY_MODE=true). Railway watcher مسح cache **على Railway**. Fly cache بقي قديماً (TTL=1ساعة) → لا يعرف بالتذاكر الجديدة.

**الحل الفوري:** `flyctl machine restart 865139be693d18 --app hajjbot-standby` → cache يُفرَّغ → بناء فهرس جديد → التذاكر تظهر فوراً. اختبر المستخدم `A09301911` → ✅ التذكرة وصلت (99.8KB PDF).

**الحل الدائم:** `flyctl secrets set STANDBY_MODE=false --app hajjbot-standby`
- ✅ كل schedulers تعمل الآن على Fly: urlWatcher + arrivalAlerts + trainTicketWatcher + hourlyReport
- ✅ كل رفع تذاكر جديد سيُكتشف خلال 3د ويُمسح cache تلقائياً
- ⚠️ StandbyGuard معطّل (شرطه `STANDBY_MODE=true`). Railway عملياً ميت (Google Cloud حظر، deploys paused منذ 19 مايو) — لو رجع قد يحاول setWebhook ويسرق البوت.

**درس مُهم (أُضيف للذاكرة):** عند تشغيل watcher على عدة instances لنفس البوت، **كل instance له cache منفصل**. مسح cache على instance A لا يفيد instance B. القاعدة: schedulers + bot على نفس الـ instance، أو cache مشترك (Redis).

**فحص أمني سريع تم خلال الجلسة:** Railway dashboard لا يزال `SSL handshake failed (525)`، لكن لا يؤثر علينا.

**حالة الكود:** لم تُعدّل أي ملفات كود (فقط إعداد `STANDBY_MODE` عبر Fly secret + إعادة تشغيل machine).

---

## 2026-05-20 | HajjBotServer DR — Fly صار الأساسي + StandbyGuard مُصلَح

**السياق:** Railway تعطّل عالمياً يوم 19 مايو. أُنشئت نسخة Fly كنسخة احتياطية ثم صارت الأساسي لأن Railway ظلّ متذبذباً.

**العطل الحرج في الجلسة:** البوت سقط فجأة بعد ساعات من النشر. السبب: كود StandbyGuard الذي أُضيف بالأمس يستخدم متغيّر `tenants` غير معرّف → `ReferenceError` يكسر السيرفر كل دقيقتين حتى وصل `max restart count of 10`.

**الإصلاح (commit `6ec50e9` على HajjBotServer/main):**
```javascript
const assertWebhook = async () => {
  const tenants = getTenants();   // ← السطر المُضاف
  for (const tenant of Object.values(tenants)) { ... }
};
```

**الوضع الحالي:**
- Fly app: `hajjbot-standby` (bom, 1 machine, 512MB)
- URL: `https://hajjbot-standby.fly.dev`
- STANDBY_MODE=true → schedulers على Fly معطّلة (لتجنّب تكرار الكتابة في B2C_v2 لو Railway رجع)
- StandbyGuard كل 2د يستعيد الـ webhook لو Railway حاول سرقته
- Redis fallback in-memory (REDIS_URL غير مضبوط)

**أسئلة محمود في الجلسة (محسومة):**
1. هل نضيف Render كنسخة ثالثة؟ — لا، تعقيد بلا فائدة.
2. هل نستخدم نطاق `wafahajj.com`؟ — مفيد لو مملوك (تبديل DNS بين المزوّدين بدل تعديل webhook يدوياً)، لكن مؤجَّل حتى يتأكد من الملكية.
3. كم طلب/ث يتحمّل البوت؟ — مريح 30-50/ث، ذروة 100/ث، سقف تيليجرام 30 رسالة/ث. لـ6500 حاج: صفر تأثير عملي.

**Railway dashboard:** `backboard.railway.com` ردّ SSL 525 (Cloudflare). لا يؤثر — البوت لا يحتاج dashboard.

**الذاكرة:**
- جديد: `project_fly_dr_standby.md` — البنية الكاملة + متغيرات البيئة + أوامر flyctl + جدول الأعطال
- محدَّث: `MEMORY.md` — سطر في أعلى الـ index
- محدَّث: `HajjBotServer/SESSION_LOG.md` — جلسة كاملة

**تحديث ساعة لاحقاً — مراجعة أمنية كاملة (commits `fc492ec` + `0aac003`):**

- **اكتشاف:** `/arrivals*` و `/dashboard/*` كانت مفتوحة للعموم (تسريب أسماء + جوازات + فنادق الحجاج بـ Excel للجميع). `SYNC_TOKEN` له default `'change-me-please'`. لا rate limit. لا helmet. body limit 50MB كاسح.
- **الإصلاحات (دون لمس سرعة البوت — /webhook مستثنى من كل middleware):**
  1. SYNC_TOKEN إجباري (`process.exit(1)` لو غائب).
  2. DASH_TOKEN جديد يحمي اللوحات (`?key=...`). القيمة: `m37p8KW-OxXpX_tjdj5aDdArhScoUZst`.
  3. `express-rate-limit`: 120/د على /api، 60/د على اللوحات.
  4. `helmet`: HSTS + X-Frame=SAMEORIGIN + nosniff + Referrer-Policy=no-referrer.
  5. body limit 1MB افتراضي، 50MB حصراً لـ /api/sync/push.
  6. `npm audit`: 0 ثغرات.
- **سرعة البوت:** 0.49ث (لم تتأثر، بل أسرع قليلاً من قبل).

**روابط اللوحات الجديدة للعمليات:**
```
/arrivals-report?key=m37p8KW-OxXpX_tjdj5aDdArhScoUZst
/dashboard/early-arrivals?key=m37p8KW-OxXpX_tjdj5aDdArhScoUZst&date=YYYY-MM-DD
```

**معلَّق:**
- إكمال fallback Redis (sadd/scard/incr/expire/hincrby/hgetall) لإسكات أخطاء analytics
- تحقّق ملكية `wafahajj.com`
- مراقبة استقرار StandbyGuard لأسبوع
- ثغرتان غير قابلتين للإغلاق برمجياً: توكن البوت في URL (نمط تيليجرام)، Service Account JSON في Fly secret (لو سُرّب → صلاحيات شيت كاملة)

---

## 2026-05-17 | كشف السكن — أسماء الحجاج لكل فندق + إلغاء قسم "بدون تعيين"

**البلاغ:** نموذج `hotels_SM477_2026-05-17.xlsx` المرسَل لمجموعات تلجرام الـ 3 للمطارات يحوي ملخّصاً فقط بلا أسماء حجاج. المطلوب: قسم لكل فندق بأسماء الحجاج، 7 أعمدة (اسم EN، جواز، مدينة الهبوط، رقم الرحلة، تاريخ، ساعة، باقة).

**تصحيح حاسم أثناء التطوير:** المستخدم أكّد أنه **لا يوجد ولا يجوز وجود** حاج بدون سكن في الكشف — مصيبة إدارية بحتة. ممنوع ذكرهم في كشف المطارات حتى كمثال؛ يُبلَّغ عنهم عبر البوت إن وُجدوا.

**التنفيذ (HajjBotServer, commit 5ef449c على main):**
- `buildHotelDistribution`: `bucket.pilgrims[]` + `unassignedPilgrims[]` (للتنبيه فقط).
- `buildHotelsExcel`: قسم لكل فندق برأس أخضر #476831 + جدول 8 أعمدة (`# | Name EN | Passport | مدينة الهبوط | رقم الرحلة | تاريخ | ساعة | باقة`). الملخّص العلوي يبقى. **لا قسم/سطر "بدون تعيين فندق" نهائياً.**
- `appendHotelsBlock`: حُذف سطر `⚠️ بدون تعيين فندق: N` من رسائل تلجرام للمطارات.
- `notifyAdminUnassignedPilgrims`: دالة جديدة ترسل تنبيه `🚨 حاج بدون سكن — مصيبة إدارية` لـ `ALERT_CHAT_ID` فقط، تُستدعى من runCycle (T-0 و T-6h/T-1h) عند `dist.unassigned > 0`.
- `scripts/test-hotels-excel-2026-05-17.js`: سكريبت اختبار محلي يولّد عيّنة في `C:/tmp/`.

**Deploy:** Railway تلقائي على push لـ main.

**التحقّق الفعلي:** ينتظر أول رحلة T-6h القادمة في مجموعات المطارات.

---

## 2026-05-17 | إيقاف تطبيق Update-TKT-Email بالكامل

**البلاغ:** المستخدم أرسل صورة من مجموعة `UPDATE TKT EMAIL` على تلجرام تُظهر تنبيهات PendingMatch (Incident# 1202327 و 1215736 بسبب "إيميل غير موجود في PD") وطلب إيقاف هذه التقارير.

**التحقق:** ضمن الخيارات الثلاثة (إيقاف التنبيهات فقط، إيقاف التطبيق كاملاً، إيقاف PendingMatch فقط) اختار **إيقاف التطبيق بالكامل**.

**السياق من الذاكرة:**
- `session_update_tkt_email_2026-04-29_failures.md` — التطبيق متوقف عن التطوير من 17 يوماً: parser multi-pilgrim لم يُصلَح + زر "ألغى الحجز" معطّل + ترقيم البلاغات خاطئ. التنبيهات أصبحت إزعاجاً بلا فائدة عملية.

**التنفيذ:**
- scriptId: `1SFmbBd21hoVZQT7Pfujp0bGlzsH5ql9TlhHMLEbzEyy3yZg6nqKh22g3`
- `clasp run deleteAllTriggers` فشل (Apps Script API غير مفعّلة كـ executable).
- البديل الناجح: ClaudeAPI deployment @37 = `AKfycbzPKeOmd1DUkSZzI8D_5LQTIQ6UYXLhvVnMA4SvS8yPGQv8JlFGEHKpZFEMgu3Wny4_`.
- نداء `?action=run&fn=deleteAllTriggers&key=ekram2026claude` → `تم حذف 2 trigger.` (`scanEmails` + `pollTelegramUpdates`).
- التحقق: `listTriggers` → `لا توجد triggers.` ✓

**ما لم يُلمَس (محفوظ للاستئناف):**
- الكود الكامل (16 ملف) — لا تعديل.
- Script Properties (PD_SPREADSHEET_ID، LOG_SPREADSHEET_ID، DRIVE_FOLDER_ID، TG_BOT_TOKEN، TG_CHAT_ID).
- Gmail labels (TKT، TKT-v2-Done، TKT-PendingMatch، TKT-Cancelled، TKT-NoPDF).
- شيتات PendingMatches + ChangesLog (سجل تاريخي).
- 19 deployment على Apps Script (للرجوع لأي إصدار).

**الذاكرة:**
- `MEMORY.md` — أُضيف سطر `⛔ Update-TKT-Email موقوف بالكامل` في أعلى الـ index.
- `project_update_tkt_email_flow.md` — أُضيف رأس بحالة `STOPPED`.

**الاستئناف لاحقاً (إن لزم):**
```
curl ".../exec?action=run&fn=createTriggerHourly&key=ekram2026claude"
```
لكن لا يُستأنَف قبل إصلاح parser multi-pilgrim + زر "ألغى الحجز".

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

