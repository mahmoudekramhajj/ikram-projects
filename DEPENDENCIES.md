# خريطة التبعيات — مشروع إكرام الضيف

> **تعليمة لـ Claude:** اقرأ هذا الملف قبل أي تعديل. حدد الشيت المتأثر → اعرض التطبيقات المرتبطة → اطلب موافقة المستخدم.

---

## 1. مصفوفة الشيتات → التطبيقات

### رحلة الحاج (48 عمود، 5965+ صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | Airport Search, IkramHajjBot, Pilgrim App, Guide App, Hotel Management, Mina Camp Search, PilgrimStats |
| **يكتب** | JourneyMerger (Ikram), Reception Airport (أعمدة 48+) |
| **أعمدة حساسة** | H (Name), I (Passport), AK-AV (فنادق), Q-AI (طيران) |
| **⚠️ إذا تغيّر** | يتأثر 7 تطبيقات — أي تغيير في ترتيب الأعمدة يكسر الكل |

### رحلة الحاج 2 (42 عمود، 6527 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | JourneyMerger فقط |
| **يكتب** | منصة نسك (خارجي) |
| **⚠️ إذا تغيّر** | يتأثر JourneyMerger → ثم كل من يقرأ "رحلة الحاج" |

### Presonal Details (33 عمود، 6532 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | JourneyMerger, Pilgrim App, B2C Sync, TourGuide Manager, PilgrimStats |
| **يكتب** | منصة نسك (خارجي) |
| **أعمدة حساسة** | 15 (اسم المرشد), 5 (جواز), 18 (رقم الباقة), 28-31 (نقل) |
| **⚠️ إذا تغيّر** | يتأثر 5 تطبيقات + JourneyMerger يعيد دمج البيانات |

### الباقات (69 عمود، 60 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | Ikram, IkramAgent, IkramHajjBot, Hotel Management, Pilgrim App |
| **يكتب** | Ikram فقط |
| **أعمدة حساسة** | C (اسم الباقة), K (عدد الحجاج), L-BB (فنادق 1-3), BN (التنقل), BO-BQ (عقود نسك) |
| **⚠️ إذا تغيّر** | يتأثر 5 تطبيقات — تغيير هيكل الفنادق يكسر Hotel Management |

### الطيران (103 عمود، 191 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | Ikram, IkramAgent, GDS, B2C Sync, Sales Report |
| **يكتب** | Ikram, GDS, B2C Sync, PNR Sync |
| **أعمدة حساسة** | B (PNR), V-AW (رحلات ذهاب/عودة), AX-BQ (باقات مرتبطة) |
| **⚠️ إذا تغيّر** | يتأثر 5 تطبيقات — 4 منها تكتب أيضاً (خطر تعارض) |

### الفنادق (52 عمود، 993 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | Ikram, Hotel Management |
| **يكتب** | لا أحد (مرجعي فقط) |
| **⚠️ إذا تغيّر** | يتأثر 2 تطبيقات |

### Guide Rabih (33 عمود، 6535 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | IkramHajjBot, Hotel Management, Pilgrim App, Guide App, Ikram (Minacamp.js, Roomtype), SheetData |
| **يكتب** | يُحدَّث من نسك فقط |
| **أعمدة حساسة** | P (اسم المرشد — 88 مرشد) |
| **⚠️ إذا تغيّر** | يتأثر 6 تطبيقات — عمود P هو المفتاح الرئيسي للمرشدين |

### مخيم مني (28 عمود، 5966 صف)
| الدور | التطبيقات |
|---|---|
| **يقرأ** | Mina Camp Search |
| **يكتب** | Ikram (Minacamp.js) |
| **⚠️ إذا تغيّر** | يتأثر تطبيق واحد |

### BotSessions
| الدور | التطبيقات |
|---|---|
| **يقرأ/يكتب** | IkramHajjBot فقط |
| **⚠️ إذا تغيّر** | لا تأثير على تطبيقات أخرى |

---

## 2. مصفوفة التطبيقات → الشيتات

| التطبيق | يقرأ من | يكتب في |
|---|---|---|
| **Ikram** | الباقات, الطيران, الفنادق | الباقات, الطيران, مخيم مني, رحلة الحاج (JourneyMerger) |
| **IkramHajjBot** | رحلة الحاج, الباقات, Guide Rabih, BotSessions | BotSessions |
| **Airport Search** | رحلة الحاج | — |
| **Hotel Management** | رحلة الحاج, الباقات, الفنادق, Guide Rabih | — |
| **Pilgrim App** | رحلة الحاج, الباقات, Presonal Details, Guide Rabih | — |
| **Guide App** | رحلة الحاج, Guide Rabih | — |
| **Reception Airport** | رحلة الحاج | رحلة الحاج (أعمدة 48+) |
| **Sales Report** | الطيران | — |
| **Mina Camp Search** | رحلة الحاج, مخيم مني | — |
| **GDS** | الطيران | الطيران |
| **B2C Sync** | الطيران, Presonal Details | الطيران |
| **PNR Sync** | — | الطيران |
| **PilgrimStats** | Presonal Details, رحلة الحاج | — |
| **IkramAgent** | الباقات, الطيران | — (⏸️ مؤجّل) |
| **SheetData** | Guide Rabih + الكل (تصدير JSON) | — |
| **Transport Management** | — (مستقل) | — |
| **Report Builder** | — (مستقل) | — |
| **IkramAdmin** | — (إعدادات) | — |

---

## 3. قواعد التأثير السريعة

| إذا عدّلت... | تحقق من... |
|---|---|
| أي عمود في **رحلة الحاج** | Airport Search, IkramHajjBot, Pilgrim App, Guide App, Hotel Management, Mina Camp Search, PilgrimStats |
| أي عمود في **الباقات** | Ikram, IkramHajjBot, Hotel Management, Pilgrim App |
| أي عمود في **الطيران** | Ikram, GDS, B2C Sync, Sales Report, PNR Sync |
| أي عمود في **Presonal Details** | JourneyMerger → رحلة الحاج → 7 تطبيقات |
| أي عمود في **Guide Rabih** | IkramHajjBot, Hotel Management, Pilgrim App, Guide App, Ikram |
| **JourneyMerger** نفسه | كل من يقرأ رحلة الحاج (7 تطبيقات) |
| هيكل أعمدة أي شيت | **كل** التطبيقات التي تقرأ منه — فهارس الأعمدة hardcoded |

---

## 4. APIs الخارجية — نقاط الاتصال

| API | التطبيقات | إذا تغيّر |
|---|---|---|
| **AeroDataBox** | Ikram, Hotel Management, GDS | بيانات الرحلات تتوقف |
| **Telegram Bot API** | IkramHajjBot | البوت يتوقف |
| **Google Sheets API** | SheetData (Service Account) | تصدير JSON يتوقف |
| **Nusk Platform** | رحلة الحاج 2, Presonal Details | البيانات لا تتحدث — JourneyMerger يدمج بيانات قديمة |
| **Anthropic Claude API** | ClaudeBrain (مخصّص)، IkramAgent (مشترك) | المخ يتوقف + IkramAgent يتوقف |

---

## 5. ClaudeBrain — المخ المركزي (Orchestrator)

**الدور:** خدمة مركزية واحدة تستدعيها كل التطبيقات لاتخاذ قرارات ذكية. بدلاً من أن يكون لكل تطبيق Claude API الخاص، الكل يتصل بالمخ → الخبرة تتراكم مركزياً.

**المسار:** `Projects/ClaudeBrain/`
**Script ID:** `1ht-aJiQZE3JwgqUx563pbUykNLe--Xwfms2a-E1KhwEa8gtFnWReH98_`
**Web App URL:** `https://script.google.com/macros/s/AKfycbzi2Vft_DeQMSkVQGKBvbzmzO5yhjdnat498L5Dnsg2S4RPFKIyQCcqhhf5Rbobcjko/exec`

**المصادقة:**
- POST: `X-Signature` header = HMAC-SHA256(body, HMAC_SECRET) في hex
- ClaudeAPI: `?action=ping&key=ekram2026claude` (التحكم عن بُعد)

**الحالة:** M1 (هيكل فارغ) — لا tools ولا intents بعد.

**التطبيقات المخططة للربط:**
| # | التطبيق | الهدف | الحالة |
|---|---|---|---|
| 1 | FlightChangeProcessor | تحليل تغييرات نسك + تنسيق كامل (نقل/فندق/استقبال/حاج) | 🔜 M2-M7 |
| 2 | IkramHajjBot | صياغة ردود للحجاج | لاحقاً |
| 3 | معالج البلاغات | توجيه البلاغات وصياغة ردود | لاحقاً |
| 4 | Transport Management | قرارات توزيع الباصات | لاحقاً |

**⚠️ إذا تعطّل المخ:** التطبيقات المعتمدة عليه تفقد قدرة القرار الذكي، لكن لا تتعطّل (fallback مخطط للموظف البشري).

**ملاحظة مهمة:** المخ **لا يُفتي** في أمور شرعية. كل سؤال شرعي من حاج → تحويل لمرشد (Guide Rabih) + رقم 1966.
