# خطة العمل المحدثة: تنفيذ تدقيق النظام وترقية المحركات (P0 + P1)

خطة تنفيذية منقحة ومضبوطة بناءً على ملاحظات التدقيق وتوجيهات المستخدم الدقيقة، مع الحفاظ على الحياد التام وعدم ذكر أي منصات أو أنظمة خارجية، وضمان سلامة البيانات واختبار كافة المسارات الحيوية.

---

## القرارات المؤكدة والتعديلات المدخلة (Confirmed Decisions & Adjustments)

1. **نقطة دخول مستقلة وواضحة للتطبيع (`lib/field-values-normalizer.ts`):**
   - إنشاء ملف مركزي خفيف ومباشر يجمع بين `lib/symbol-map.ts` و`lib/notary-word-variables.ts` و`lib/tafqeet.ts`.
   - تصدير دالة رئيسية واحدة: `normalizeAllNotaryFieldValues(fieldValues, options?: { contractBodyHtml?: string; pageCount?: number })`.
   - استدعاء التطبيع إلزامياً في **3 محطات أساسية**:
     - **الحفظ (Auto-save / Manual save):** قبل كتابة العقد في IndexedDB.
     - **الفهرسة والختم (Indexing & SHA-256 Seal):** قبل تجميد أصل العقد وحفظ الفهرس.
     - **التصدير (Word Export):** قبل تنفيذ `mergePlaceholdersIntoHtml` وتوليد ملف `.docx`.

2. **علم الترحيل المزدوج بدون تغيير مخطط IndexedDB:**
   - استخدام معرّف ثابت `__meta_office_templates_migrated__` يتم تخزينه مباشرة داخل مخزن `STORE_OFFICE_TEMPLATES` القائم حالياً (`notary_office_templates_v1`).
   - لا حاجة لترقية نسخة قاعدة البيانات أو إنشاء مخازن جديدة؛ المخزن الحالي يستوعب هذا المعرف بسلاسة ويفحص بجانب `localStorage`.

3. **حساب دقيق لصفحات التفقيط (`pages2lettre_acte`):**
   - تمرير عدد الصفحات الناتج عن محرك الترقيم التوثيقي الفعلي (`pageCount` من `injectNotarialPaginationIntoHtml` أو المقدر من `docMetrics`) إلى دالة التطبيع لضمان تطابق التفقيط الحرفي مع الصفحات المطبوعة والمصدرة.

4. **تحديث README.md بحيادية كاملة:**
   - الاقتصار على مصطلحات المشروع الداخلية: الوسوم الذكية للمحرر `[…]`، الرموز القياسية `${…}`، والوسوم اللاتينية المعتمدة `[tag]`.
   - توثيق `DB_VERSION = 9` والشجرة الفعلية ومحركات القوالب الستة دون أي إشارة إلى أسماء أنظمة منافسة.

5. **شريط الحالة (Status Bar):**
   - عند عدم وجود أي وسوم مستخرجة بالعقد (`totalVariablesCount === 0`)، يُعرض النص التوضيحي الصريح:
     `لا توجد متغيرات مستخرجة بالعقد`

---

## 1. الهيكلية ومخطط تدفق البيانات (Architecture & Flow)

```
┌────────────────────────────────────────────────────────────────────────┐
│                   A4 Notary Editor & Inputs                            │
│  [Parties] + [Properties] + [Clerks] + [Editor Body HTML] + [Indexing] │
└──────────────────────────────────┬─────────────────────────────────────┘
                                   │
                                   ▼
┌────────────────────────────────────────────────────────────────────────┐
│                   lib/field-values-normalizer.ts                       │
│          normalizeAllNotaryFieldValues(fieldValues, options)           │
│                                                                        │
│  ├── 1. Arabic Keys Sync:      [الطرف_الأول_الاسم], [رقم_الفهرس]       │
│  ├── 2. Standard Symbols Sync: ${c_1}, ${num_acte}, ${act_value}       │
│  ├── 3. Standard Latin Sync:   [partie1_nom], [num_repertoire]        │
│  └── 4. Tafqeet Resolution:    Currency (DZD) + Dates + Pages Count    │
└──────────────┬───────────────────┬──────────────────────┬──────────────┘
               │                   │                      │
               ▼                   ▼                      ▼
┌──────────────────────┐ ┌───────────────────┐ ┌─────────────────────────┐
│     Auto-Save        │ │ Indexing & Seal   │ │    DOCX Export Engine   │
│ Before IDB write     │ │ Before SHA-256    │ │ Before merge & download │
└──────────────────────┘ └───────────────────┘ └─────────────────────────┘
```

---

## 2. خطوات التنفيذ الميدانية المحددة (Targeted Tasks)

### الخطوة 1: بناء وحدة التطبيع الموحدة `lib/field-values-normalizer.ts`
- **الملف المستهدف:** `lib/field-values-normalizer.ts` (جديد ومستقل).
- **السلوك المتوقع:**
  - استقبال أي قاموس `fieldValues` مع وسيطات اختيارية (`contractBodyHtml`، `pageCount`).
  - تحويل المفاتيح وتبادلها الكامل بين الصيغ الثلاث (العربية ↔ `${}` ↔ اللاتينية).
  - تقييم الحقول المالية بالتفقيط التوثيقي الجزائري (`act_value`، `act_value_lettre`، `totalLettre`).
  - تقييم عدد الصفحات بالتفقيط القانوني (`pages2lettre_acte`) بالاعتماد على `pageCount` الفعلي.
  - تقييم التواريخ القانونية بالأحرف (`date2lettre_acte` / `date_acte_lettres`).
- **معيار التحقق:** وحدة برمجية مستقلة قابلة للاستدعاء والاختبار في معزل تام.

### الخطوة 2: ربط التطبيع في محطات الحفظ، الفهرسة، والتصدير
- **الملفات المستهدفة:**
  - `lib/storage-browser.ts`: تطبيق `normalizeAllNotaryFieldValues` قبل تخزين العقد النشط في `saveActiveDocument`.
  - `lib/indexing-engine.ts` / `useNotaryLifecycle.ts`: استدعاء التطبيع عند إتمام الفهرسة وحقن رقم وتاريخ العقد في بيانات الحقول.
  - `lib/notary-word-templates.ts` و`lib/docx-engine.ts`: استدعاء التطبيع في بداية `exportContractWithStandardTemplate` و`downloadNotaryDocx` قبل دمج القوالب.
- **معيار التحقق:** ثبات واكتمال جميع المفاتيح في أي سجل محفوظ أو ملف مصدّر.

### الخطوة 3: تعزيز علم الترحيل المزدوج في IndexedDB
- **الملف المستهدف:** `lib/storage-browser.ts`.
- **السلوك المتوقع:**
  - عند نجاح الترحيل، كتابة سجل دائم في `STORE_OFFICE_TEMPLATES` بالمعرّف `__meta_office_templates_migrated__` بقيمة `{ id: '__meta_office_templates_migrated__', migrated: true, timestamp: ... }`.
  - تحديث `isOfficeTemplatesMigrated` ليفحص:
    1. المتغير في الذاكرة `memoryTemplatesMigrated`.
    2. `localStorage.getItem('office_templates_migrated_v1')`.
    3. استعلام السجل `__meta_office_templates_migrated__` من `STORE_OFFICE_TEMPLATES` في IndexedDB.
- **معيار التحقق:** محاكاة مسح `localStorage` والتأكد من عدم إعادة بذر القوالب الافتراضية إذا كانت IndexedDB تحتوي بيانات المستخدم.

### الخطوة 4: تصحيح نص شريط الحالة عند صفر متغيرات
- **الملف المستهدف:** `components/editor/EditorStatusBars.tsx`.
- **السلوك المتوقع:**
  - استبدال الصياغة المضللة `جميع متغيرات العقد مكتملة (0)` بالصياغة الصريحة: `لا توجد متغيرات مستخرجة بالعقد` عندما يكون `totalVariablesCount === 0`.
- **معيار التحقق:** فتح عقد فارغ في الواجهة وملاحظة ظهور العبارة الصحيحة دون أي لبس.

### الخطوة 5: تحديث وتصحيح `README.md` بحيادية تامة
- **الملف المستهدف:** `README.md`.
- **السلوك المتوقع:**
  - تصحيح إصدار التخزين إلى `IndexedDB v9`.
  - توثيق ملفات المحركات المركزية:
    - `lib/symbol-map.ts`: دليل الرموز التوثيقية المعتمدة والتفقيط القانوني.
    - `lib/notary-word-templates.ts`: مكتبة القوالب الرسمية الستة (.docx) والتخصيص.
    - `lib/notary-word-variables.ts`: قاموس الوسوم اللاتينية المعتمدة.
    - `lib/field-values-normalizer.ts`: محرك التطبيع ثنائي الاتجاه.
  - توثيق دعم حلقة البنود التكرارية `${block}...${/block}` في ملفات Word.
  - استبدال المراجع القديمة (`SmartModals`) بالنوافذ الموديلية الحالية (`EditorModalsHost`، `DocumentWorkflowModals`، `WordTemplatesManagerModal`).
  - الالتزام التام بعدم ذكر أي منصة أو نظام خارجي في التوثيق.
- **معيار التحقق:** مراجعة نص `README.md` ومطابقته الفعلية للكود.

### الخطوة 6: حزمة اختبارات وحدة موسعة في `tests/notarial-core.test.ts`
- **الملف المستهدف:** `tests/notarial-core.test.ts`.
- **الحالات المطلوب اختبارها:**
  1. **التحويل الثلاثي للرموز:** إدخال مفتاح عربي فقط (`رقم_الفهرس`) → توليد `${num_acte}` و`[num_repertoire]`؛ وإدخال لاتيني فقط → توليد العربي و`${}`.
  2. **حلقة البنود التكرارية (`expandBlockLoop`):** التحقق من تكرار البنود النشطة، والتأكد القطعي من أن البند المعطّل (`enabled: false`) لا يُدرج في الناتج.
  3. **تقييم التفقيط الديناميكي:** التحقق من مطابقة `pages2lettre_acte` لعدد الصفحات الممرر، وتفقيط المبالغ بالدينار الجزائري وتاريخ العقد بالحروف.
- **معيار التحقق:** تشغيل `npm test` واجتياز 100% من الاختبارات (15/15 اختباراً).

---

## 3. خطة التحقق والجاهزية (Verification Plan)

1. **اختبارات الوحدة الآلية:**
   - تنفيذ أمر: `npm test`
   - التأكد من نجاح كافة الاختبارات بدون أي تحذيرات أو أخطاء.
2. **فحص الـ Linting:**
   - تنفيذ أمر: `npx eslint .` (أو `lint_applet`)
3. **بناء المشروع للإنتاج:**
   - تنفيذ أمر: `compile_applet` للتأكد من خلو ملفات TypeScript والـ Next.js App Router من أي أخطاء تجميع.
4. **تأكيد الحياد التام للهوية:**
   - فحص شامل لكافة الملفات للتأكد من عدم وجود أي اسم أو إشارة لمنافسين خارجيين.
