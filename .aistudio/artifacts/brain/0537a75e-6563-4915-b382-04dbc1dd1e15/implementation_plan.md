# خطة استكمال التنظيف المعماري وتفكيك المحرر (محدّثة وفق مراجعة 26)

تم تحديث الخطة لتثبيت **التصحيحات الإلزامية الستة** وعكس ترتيب التنفيذ ليبدأ بطبقة البيانات والترحيل، ثم تنظيف الواجهات والـ props، ثم تفكيك/حذف الملفات المتوازية بعد التحقق من المستوردات، وأخيراً استخراج الـ 3 Hooks بحدود واضحة ونقية.

---

## الترتيب التنفيذي الآمن (من 1 إلى 5)

### المرحلة 1: البيانات والترحيل أولاً (`bodyHtml?` + ترحيل التنزيلات + توافق `BackupBundle`)

1. **جعل `bodyHtml` اختيارياً في `DocumentRevision` للأحداث الخفيفة (`lib/types.ts`)**:
   - تعديل `DocumentRevision` ليصبح:
     ```ts
     bodyHtml?: string; // مطلوب فقط عند eventType: 'save' | 'snapshot' | 'seal'
     ```
   - في `lib/document-io-workflows.ts` وعند تسجيل أحداث `eventType: 'export' | 'derived' | 'lifecycle'`، لا يتم تخزين `bodyHtml` داخل المراجعة لتجنب تضخم التخزين المحلي.
   - في نوافذ المقارنة والاسترجاع (`VersionDiffModal` / `SnapshotsHistoryModal`)، يقتصر زر «استرجاع اللقطة» أو «مقارنة النص» على المراجعات التي تملك `bodyHtml`.

2. **الترحيل التلقائي لسجلات `DownloadArchiveItem` القديمة (`lib/storage-browser.ts` و `lib/storage-desktop.ts`)**:
   - داخل `loadDocumentRevisions()`، فحص وجود بيانات قديمة في مفتاح `notaire_download_archive_v1` (أو ملف الأرشيف القديم)، وتحويل كل عنصر إلى `DocumentRevision` خفيف (`eventType: item.docTypeCode === 'acte' ? 'export' : 'derived'`, مع الاحتفاظ بـ `fileName`, `docTypeCode`, `docTypeLabel`, `createdAt`)، ثم دمجها لمرة واحدة وحذف المخزن المستقل القديم.
   - حذف دوال المخازن المستقلة المهجورة من `OfficeStore` و `lib/storage.ts` (`loadDownloadArchive`, `saveDownloadArchiveItem`, `loadDerivedDocTemplates`, `loadWordTemplates`, `loadCustomTemplates` ومشتقاتها).

3. **حماية استيراد النسخ الاحتياطية القديمة (`BackupBundle`)**:
   - الإبقاء على أنواع التوافق كأنواع أرشيفية اختيارية (`LegacyCustomTemplate`, `LegacyDerivedDocTemplate`, `LegacyWordTemplateDefinition`, `LegacyDownloadArchiveItem`) مخصصة حصرياً لقارئ الاستيراد والترحيل في `BackupBundle` و `storage-*`، بينما تعتمد كافة واجهات العمل النشطة على `OfficeTemplate` و `DocumentRevision` فقط.

---

### المرحلة 2: تنظيف الواجهات والـ Props من المسارات المزدوجة (الأولوية ب + هـ)

1. **`components/SidebarWorkspace.tsx` و `components/sidebar/PartiesTabPanel.tsx`**:
   - إزالة `partyFields`, `onUpdateFieldValue`, `onAddCustomPartyField`, `onUpdatePartyFieldMeta`, و `onDeletePartyField` من `SidebarWorkspaceProps`.
   - إزالة `derivedTemplates`, `onSaveDerivedTemplate`, و `onResetDerivedTemplates` من `SidebarWorkspaceProps` والاكتفاء بـ `templates: OfficeTemplate[]`.
   - في `PartiesTabPanel.tsx`: الاعتماد الحصري على `partyCards: ContractPartyCard[]` + دليل الأطراف والعقارات المحفوظة (`savedParties`, `savedProperties`) + `fieldValues` للوسوم الحرة الملتقطة من المتن (`extractedPlaceholders`).

2. **`components/editor/EditorModalsHost.tsx`**:
   - تنظيف `EditorModalsHostProps` من جميع الخصائص الميتة التي لم تعد تُمرر أو تُستخدم (`showSmartVarsModal`, `partyFields`, `downloads`, `showWordTemplatesModal`, `wordTemplates`, `onSaveWordTemplate`, `onDeleteWordTemplate`, إلخ).

---

### المرحلة 3: التحقق من المستوردات ونقل الدوال والملفات النشطة قبل الحذف (الأولوية ج)

بعد فحص المستوردات الفعلي في المستودع تبيّن ما يلي:
- `components/EditorModalsHost.tsx` (في الجذر): ملف ميت غير مستورد في أي مكان (المستورد الفعلي في `NotaryEditorApp.tsx` هو `./editor/EditorModalsHost`).
- `lib/docx-template-engine.ts`: مستورد فقط في `components/SmartModals.tsx` داخل نافذة `WordTemplatesManagerModal` المهجورة.
- `components/SmartModals.tsx`: يحتوي على نافذتين مهجورتين (`SmartVariablesModal`, `WordTemplatesManagerModal`) **بالإضافة إلى 6 نوافذ نشطة** يستوردها `components/editor/EditorModalsHost.tsx` (`VersionDiffModal`, `DocxImportPreviewModal`, `SaveAsTemplateModal`, `SnapshotsHistoryModal`, `OnboardingTourModal`, `MultiSourceStartModal`).

**خطوات التنفيذ الآمنة:**
1. نقل الدالة النقية `parseContractIntoClauses` إلى `lib/docx-engine.ts` (لإبقائها متاحة لمحرك الوثائق والاختبارات).
2. نقل النوافذ الست النشطة فقط (`VersionDiffModal`, `DocxImportPreviewModal`, `SaveAsTemplateModal`, `SnapshotsHistoryModal`, `OnboardingTourModal`, `MultiSourceStartModal`) إلى `components/editor/EditorModalsHost.tsx` بعد تنظيفها من أي إشارات لـ `CustomTemplate` أو `DownloadArchiveItem`.
3. بعد التأكد من عدم وجود أي `import` متبقٍ، حذف الملفات الثلاثة المهجورة:
   - `lib/docx-template-engine.ts`
   - `components/SmartModals.tsx`
   - `components/EditorModalsHost.tsx` (الملف المكرر في الجذر)

---

### المرحلة 4: استخراج الـ 3 Hooks بحدود واضحة وبدون حالة مشتركة مخفية (الأولوية أ)

يتم استخراج الـ Hooks الثلاثة بحيث يكون `NotaryEditorApp.tsx` هو الـ **Orchestrator** الذي يمرر المعاملات الصريحة بين الجلسة ودورة الحياة والكتالوج:

1. **`hooks/useDocumentSession.ts`**
   - يدير حالة جلسة التحرير الحالية فقط: `docId`, `docTitle`, `bodyEditorRef`, `headerEditorRef`, `footerEditorRef`, `showHeaderFooter`, `pageNumberingEnabled`, `zoom`, `previewMergedMode`, `mergedPreviewHtml`, `printPaginatedHtml`.
   - يدير حالة التحرير الفرعي المؤقت (`editingDerivedTpl`, `editingClauseObj`, `stashedMainDocRef`).
   - يدير سجل التراجع والإعادة (`undoStack`, `redoStack`, `recordUndoSnapshot`, `handleUndo`, `handleRedo`) وإحصائيات المستند (`metrics`).
   - يوفّر دالة قراءة لقطة الجلسة الحالية (`getSessionSnapshot()`) لتمريرها صراحةً إلى دوال الحفظ أو الفهرسة أو التصدير.

2. **`hooks/useOfficeCatalog.ts`**
   - يدير سجلات المكتب المرجعية فقط: `templates` (`OfficeTemplate[]`), `clauses` (`NotaryClause[]`), `estates` (`SubdivisionEstate[]`), `savedParties` (`SavedPartyRecord[]`), `savedProperties` (`SavedPropertyRecord[]`), `folders` (`ContractFolder[]`), `clerks` (`NotaryClerk[]`), وحالة ربط مجلد سطح المكتب (`localDirState`, `connectAndReloadLocalDirectory`).
   - **لا يتدخل** في ختم العقود أو قراءة الأصول المختومة.

3. **`hooks/useNotaryLifecycle.ts`**
   - يدير العقود المحفوظة (`documents`)، سجل المراجعات والأحداث الموحّد (`revisions`)، وحالة فك القفل الاستثنائي المؤقت (`unlockIndexedBody`).
   - يقدّم دوال دورة الحياة التي **تستقبل لقطة الجلسة بمعاملات صريحة** (مثل `bodyHtml`, `headerHtml`, `footerHtml`, `fieldValues`, `partyCards`) دون قراءة مخفية من DOM:
     - حفظ العقد وأخذ اللقطات (`saveContractSnapshot`, `recordRevisionEvent`).
     - الفهرسة الرسمية والختم وقراءة الأصل المختوم (`sealContractOriginal`, `readSealedOriginal`).
     - توليد الوثائق المشتقة وتصدير `DOCX` عبر استدعاء المحركات النقية (`docx-engine`, `document-io-workflows`) وتمرير लقطة الجلسة إليها.

---

### المرحلة 5: التحقق الشامل

- تشغيل `npm test` للتحقق من نجاح جميع الاختبارات الآلية التسعة (الفهرسة، الترقيم التوثيقي `../..`، قفل الأصول، تقسيم الفقرات، وجسر المكتب).
- تشغيل `lint_applet` و `compile_applet` لضمان خلو المشروع من أي أخطاء TypeScript أو ESLint.
