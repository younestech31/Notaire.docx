import {
  type BackupBundle,
  type ContractFolder,
  type CustomTemplate,
  type DerivedDocTemplate,
  type DocumentRevision,
  type DownloadArchiveItem,
  type NotaryClerk,
  type NotaryClause,
  type OfficeTemplate,
  type OfficeTemplateKind,
  type PartyField,
  type PublicationTrack,
  type RegistrationBatch,
  type SavedDocument,
  type SavedPartyRecord,
  type SavedPropertyRecord,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  type SubdivisionEstate,
  type WordTemplateDefinition,
} from './types';

export interface SealContractOriginalInput {
  documentId: string;
  year: number;
  indexNumber: string;
  docxBytes: Uint8Array;
}

export interface SealContractOriginalResult {
  relativePath: string;
  sha256: string;
}

export interface ReadSealedOriginalResult {
  relativePath: string;
  sha256: string;
  bytes: Uint8Array;
}

export interface OfficeStoreRuntimeInfo {
  mode: 'browser' | 'desktop';
  dataDirectory: string | null;
}

export interface OfficeStore {
  getRuntimeInfo(): Promise<OfficeStoreRuntimeInfo>;
  sealContractOriginal(
    input: SealContractOriginalInput
  ): Promise<SealContractOriginalResult>;
  readSealedOriginal(
    documentId: string
  ): Promise<ReadSealedOriginalResult | null>;

  // 1. Clauses
  loadNotaryClauses(): Promise<NotaryClause[]>;
  saveNotaryClause(clause: NotaryClause): Promise<void>;
  deleteNotaryClause(id: string): Promise<void>;

  // 2. Unified Office Templates
  loadOfficeTemplates(): Promise<OfficeTemplate[]>;
  saveOfficeTemplate(template: OfficeTemplate): Promise<void>;
  deleteOfficeTemplate(id: string): Promise<void>;
  exportOfficeTemplatesJson(): Promise<string>;
  importOfficeTemplatesJson(templates: unknown[]): Promise<number>;

  // 2a. Legacy Derived Document Templates (shimmed to OfficeTemplate)
  loadDerivedDocTemplates(): Promise<DerivedDocTemplate[]>;
  saveDerivedDocTemplate(tpl: DerivedDocTemplate): Promise<void>;
  deleteDerivedDocTemplate(id: string): Promise<void>;

  // 2b. Legacy Word Templates (shimmed to OfficeTemplate)
  loadWordTemplates(): Promise<WordTemplateDefinition[]>;
  saveWordTemplate(tpl: WordTemplateDefinition): Promise<void>;
  deleteWordTemplate(id: string): Promise<void>;

  // 3. Document Revisions
  loadDocumentRevisions(): Promise<DocumentRevision[]>;
  saveDocumentRevision(rev: DocumentRevision): Promise<void>;
  deleteDocumentRevision(id: string): Promise<void>;

  // 4. Download Archive (legacy compat)
  loadDownloadArchive(): Promise<DownloadArchiveItem[]>;
  saveDownloadArchiveItem(item: DownloadArchiveItem): Promise<void>;
  deleteDownloadArchiveItem(id: string): Promise<void>;
  clearDownloadArchive(): Promise<void>;

  // 5. Legacy Custom Templates (shimmed to OfficeTemplate)
  loadCustomTemplates(): Promise<CustomTemplate[]>;
  saveCustomTemplate(template: CustomTemplate): Promise<void>;
  deleteCustomTemplate(id: string): Promise<void>;
  exportCustomTemplatesJson(): Promise<string>;
  importCustomTemplatesJson(templates: CustomTemplate[]): Promise<number>;

  // 6. Saved Documents
  loadSavedDocuments(): Promise<SavedDocument[]>;
  saveDocumentRecord(doc: SavedDocument): Promise<void>;
  deleteDocumentRecord(id: string): Promise<void>;

  // 7. Subdivision Estates
  loadSubdivisionEstates(): Promise<SubdivisionEstate[]>;
  saveSubdivisionEstate(estate: SubdivisionEstate): Promise<void>;
  deleteSubdivisionEstate(id: string): Promise<void>;

  // 7b. Saved Parties & Properties Directory
  loadSavedParties(): Promise<SavedPartyRecord[]>;
  savePartyRecord(party: SavedPartyRecord): Promise<void>;
  deletePartyRecord(id: string): Promise<void>;
  loadSavedProperties(): Promise<SavedPropertyRecord[]>;
  savePropertyRecord(prop: SavedPropertyRecord): Promise<void>;
  deletePropertyRecord(id: string): Promise<void>;

  // 7c. Clerks & Folders
  loadNotaryClerks(): Promise<NotaryClerk[]>;
  saveNotaryClerk(clerk: NotaryClerk): Promise<void>;
  deleteNotaryClerk(id: string): Promise<void>;
  loadContractFolders(): Promise<ContractFolder[]>;
  saveContractFolder(folder: ContractFolder): Promise<void>;
  deleteContractFolder(id: string): Promise<void>;

  // 7d. Registration Batches & Land Publication Tracks
  loadRegistrationBatches(): Promise<RegistrationBatch[]>;
  saveRegistrationBatch(batch: RegistrationBatch): Promise<void>;
  deleteRegistrationBatch(id: string): Promise<void>;
  loadPublicationTracks(): Promise<PublicationTrack[]>;
  savePublicationTrack(track: PublicationTrack): Promise<void>;
  deletePublicationTrack(id: string): Promise<void>;

  // 8. Party Fields
  loadPartyFields(): Promise<PartyField[]>;
  savePartyFields(fields: PartyField[]): Promise<void>;

  // 9. Active Draft Session
  loadActiveDraftSession(): Promise<SavedDocument | null>;
  saveActiveDraftSession(doc: SavedDocument): Promise<void>;

  // 10. Full Backup Export & Import
  exportFullBackupBundle(): Promise<BackupBundle>;
  importFullBackupBundle(bundle: BackupBundle): Promise<{
    templatesCount: number;
    documentsCount: number;
    estatesCount: number;
    clausesCount: number;
  }>;
}

export const DEFAULT_PARTY_FIELDS: PartyField[] = [
  // 1. الأطراف (party1 & party2)
  { key: 'الطرف_الأول_الاسم', label: 'الطرف الأول (الاسم واللقب)', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_تاريخ_الميلاد', label: 'تاريخ ميلاد الطرف الأول', value: '', category: 'party1', inputType: 'date' },
  { key: 'الطرف_الأول_مكان_الميلاد', label: 'مكان ميلاد الطرف الأول', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_النسب', label: 'اسم الأب والأم (بن/بنت)', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_الهوية', label: 'رقم ونوع وثيقة الهوية (ط.ت.و / ر.س)', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_الإقامة', label: 'عنوان إقامة الطرف الأول', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الثاني_الاسم', label: 'الطرف الثاني (الاسم واللقب)', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_تاريخ_الميلاد', label: 'تاريخ ميلاد الطرف الثاني', value: '', category: 'party2', inputType: 'date' },
  { key: 'الطرف_الثاني_مكان_الميلاد', label: 'مكان ميلاد الطرف الثاني', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_النسب', label: 'اسم الأب والأم (بن/بنت)', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_الهوية', label: 'رقم ونوع وثيقة الهوية (ط.ت.و / ر.س)', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_الإقامة', label: 'عنوان إقامة الطرف الثاني', value: '', category: 'party2', inputType: 'text' },
  // 2. العقار (property)
  { key: 'رقم_الحصة', label: 'رقم الحصة العقارية', value: '', category: 'property', inputType: 'number' },
  { key: 'طبيعة_الحصة', label: 'طبيعة الحصة (شقة/محل/مرآب)', value: '', category: 'property', inputType: 'select', options: ['شقة سكنية', 'محل تجاري', 'مرآب', 'قطعة أرضية', 'مسكن فردي'] },
  { key: 'الطابق', label: 'الطابق والعمارة', value: '', category: 'property', inputType: 'text' },
  { key: 'المساحة', label: 'المساحة (م²)', value: '', category: 'property', inputType: 'number' },
  { key: 'الأجزاء_المشتركة', label: 'الحصة في الأجزاء المشتركة', value: '', category: 'property', inputType: 'text' },
  { key: 'تعيين_الحصة_الكامل', label: 'التعيين الكامل للحصة (من جدول الوصف)', value: '', category: 'property', inputType: 'text' },
  { key: 'مراجع_الوصف_التقسيمي', label: 'مراجع عقد الوصف التقسيمي', value: '', category: 'property', inputType: 'text' },
  // 3. الثمن والمالية (financial)
  { key: 'الثمن_بالأحرف', label: 'الثمن الإجمالي بالأحرف', value: '', category: 'financial', inputType: 'text' },
  { key: 'الثمن_بالأرقام', label: 'الثمن الإجمالي بالأرقام', value: '', category: 'financial', inputType: 'number' },
  { key: 'طريقة_الدفع', label: 'طريقة الدفع والقبض (نقداً / صك / خارج المكتب)', value: '', category: 'financial', inputType: 'text' },
  // 4. الوضعية الجبائية المنفصلة (fiscal)
  { key: 'تعيين_المبيع_الجبائي', label: 'تعيين المبيع (للوضعية الجبائية)', value: '', category: 'fiscal', inputType: 'text' },
  { key: 'أصل_الملكية_وثمن_الشراء_السابق', label: 'أصل الملكية وثمن الشراء السابق', value: '', category: 'fiscal', inputType: 'text' },
  { key: 'بيانات_البائع_الجبائية', label: 'بيانات البائع الجبائية (رقم التعريف الجبائي / المهنة)', value: '', category: 'fiscal', inputType: 'text' },
  { key: 'بيانات_المشتري_الجبائية', label: 'بيانات المشتري الجبائية (رقم التعريف الجبائي / المهنة)', value: '', category: 'fiscal', inputType: 'text' },
  { key: 'تكملة_خلف_الصفحة', label: 'تكملة خلف الصفحة (ملاحظات وتفاصيل جبائية إضافية)', value: '', category: 'fiscal', inputType: 'text' },
  // 5. الفهرس والتسجيل والشهر العقاري (index)
  { key: 'رقم_الإيداع', label: 'رقم ملف الإيداع السنوي (قبل الفهرسة)', value: '', category: 'index', inputType: 'text' },
  { key: 'تاريخ_العقد', label: 'تاريخ تحرير وتوقيع العقد', value: '', category: 'index', inputType: 'date' },
  { key: 'رقم_الفهرس', label: 'رقم الفهرس الرسمي السنوي', value: '', category: 'index', inputType: 'text' },
  { key: 'مفتشية_التسجيل', label: 'مفتشية التسجيل والطابع المختصة', value: '', category: 'index', inputType: 'text' },
  { key: 'مراجع_التسجيل', label: 'مراجع التسجيل (التاريخ، رقم الوصل، المفتشية)', value: '', category: 'index', inputType: 'text' },
  { key: 'المحافظة_العقارية', label: 'المحافظة العقارية المختصة', value: '', category: 'index', inputType: 'text' },
  { key: 'مراجع_الشهر', label: 'مراجع الشهر العقاري (التاريخ، المجلد، الرقم، المحافظة)', value: '', category: 'index', inputType: 'text' },
  { key: 'عنوان_المكتب', label: 'عنوان مكتب التوثيق الرسمي', value: '', category: 'index', inputType: 'text' },
];

const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;
const pCenterBold = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:center;font-weight:bold;`;

export const DEFAULT_DERIVED_DOC_TEMPLATES: DerivedDocTemplate[] = [
  {
    id: 'derived_extrait',
    code: 'extrait',
    name: 'المستخرج (Extrait)',
    description: 'مستخرج العقد الموجه لمفتشية التسجيل والطابع — قابل للتعديل الكامل داخل المحرر',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">مستخرج عقد توثيقي لأجل التسجيل</p>
      <p dir="rtl" style="${pStyle}">فهرس رقم: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> — بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> — مفتشية التسجيل: <span class="smart-tag" data-var="مفتشية_التسجيل">[مفتشية_التسجيل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الطرف الأول:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، المولود بتاريخ <span class="smart-tag" data-var="الطرف_الأول_تاريخ_الميلاد">[الطرف_الأول_تاريخ_الميلاد]</span> بـ <span class="smart-tag" data-var="الطرف_الأول_مكان_الميلاد">[الطرف_الأول_مكان_الميلاد]</span>، بن <span class="smart-tag" data-var="الطرف_الأول_النسب">[الطرف_الأول_النسب]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الأول_الإقامة">[الطرف_الأول_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الطرف الثاني:</span> <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، المولود بتاريخ <span class="smart-tag" data-var="الطرف_الثاني_تاريخ_الميلاد">[الطرف_الثاني_تاريخ_الميلاد]</span> بـ <span class="smart-tag" data-var="الطرف_الثاني_مكان_الميلاد">[الطرف_الثاني_مكان_الميلاد]</span>، بن <span class="smart-tag" data-var="الطرف_الثاني_النسب">[الطرف_الثاني_النسب]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الثاني_الإقامة">[الطرف_الثاني_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">تعيين العقار (الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span>):</span> <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الثمن الإجمالي:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">تأشيرة ومراجع التسجيل:</span> <span class="smart-tag" data-var="مراجع_التسجيل">[مراجع_التسجيل]</span></p>
    `.trim(),
  },
  {
    id: 'derived_publication',
    code: 'publication',
    name: 'إجراء الشهر العقاري (Publication)',
    description: 'وثيقة إجراء الشهر العقاري الموجهة للمحافظة العقارية — قابلة للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">قائمة إجراء الشهر العقاري</p>
      <p dir="rtl" style="${pStyle}">المحافظة العقارية المختصة: <span class="smart-tag" data-var="المحافظة_العقارية">[المحافظة_العقارية]</span></p>
      <p dir="rtl" style="${pStyle}">موجب العقد المحرر بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> تحت رقم الفهرس: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> — <span class="smart-tag" data-var="مراجع_التسجيل">[مراجع_التسجيل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">أولاً — تعيين الأطراف:</span></p>
      <p dir="rtl" style="${pStyle}">1. المتصرف (البائع): <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، تاريخ ومكان الميلاد: <span class="smart-tag" data-var="الطرف_الأول_تاريخ_الميلاد">[الطرف_الأول_تاريخ_الميلاد]</span> <span class="smart-tag" data-var="الطرف_الأول_مكان_الميلاد">[الطرف_الأول_مكان_الميلاد]</span>، النسب: <span class="smart-tag" data-var="الطرف_الأول_النسب">[الطرف_الأول_النسب]</span>.</p>
      <p dir="rtl" style="${pStyle}">2. المتصرف إليه (المشتري): <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، تاريخ ومكان الميلاد: <span class="smart-tag" data-var="الطرف_الثاني_تاريخ_الميلاد">[الطرف_الثاني_تاريخ_الميلاد]</span> <span class="smart-tag" data-var="الطرف_الثاني_مكان_الميلاد">[الطرف_الثاني_مكان_الميلاد]</span>، النسب: <span class="smart-tag" data-var="الطرف_الثاني_النسب">[الطرف_الثاني_النسب]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">ثانياً — التعيين الدقيق للعقار والوصف التقسيمي:</span></p>
      <p dir="rtl" style="${pStyle}">الحصة رقم: <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> — الطبيعة: <span class="smart-tag" data-var="طبيعة_الحصة">[طبيعة_الحصة]</span> — الطابق: <span class="smart-tag" data-var="الطابق">[الطابق]</span> — المساحة: <span class="smart-tag" data-var="المساحة">[المساحة]</span> م² — الأجزاء المشتركة: <span class="smart-tag" data-var="الأجزاء_المشتركة">[الأجزاء_المشتركة]</span>.</p>
      <p dir="rtl" style="${pStyle}">الوصف الكامل: <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}">مراجع الوصف التقسيمي: <span class="smart-tag" data-var="مراجع_الوصف_التقسيمي">[مراجع_الوصف_التقسيمي]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">ثالثاً — الثمن والشروط المالية:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">مراجع الشهر العقاري:</span> <span class="smart-tag" data-var="مراجع_الشهر">[مراجع_الشهر]</span></p>
    `.trim(),
  },
  {
    id: 'derived_attestation',
    code: 'attestation',
    name: 'شهادة توثيقية / شهادة بيع (Attestation)',
    description: 'شهادة تسلم للأطراف لإثبات إبرام العقد بمكتب الموثق — قابلة للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">شهادة توثيقية</p>
      <p dir="rtl" style="${pStyle}">يشهد الموثق الموقع أسفله أنه بموجب عقد محرر بمكتبنا بتاريخ <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> تحت رقم الفهرس <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> (<span class="smart-tag" data-var="مراجع_التسجيل">[مراجع_التسجيل]</span>):</p>
      <p dir="rtl" style="${pStyle}">قام السيد(ة): <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، الحامل لوثيقة الهوية: <span class="smart-tag" data-var="الطرف_الأول_الهوية">[الطرف_الأول_الهوية]</span>.</p>
      <p dir="rtl" style="${pStyle}">بالتصرف لفائدة السيد(ة): <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، الحامل لوثيقة الهوية: <span class="smart-tag" data-var="الطرف_الثاني_الهوية">[الطرف_الثاني_الهوية]</span>.</p>
      <p dir="rtl" style="${pStyle}">في العقار المعين كما يلي: <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}">سلمت هذه الشهادة للمعني بالأمر للإدلاء بها في حدود ما يسمح به القانون.</p>
    `.trim(),
  },
  {
    id: 'derived_proce_verbal',
    code: 'proce_verbal',
    name: 'الصيغة التنفيذية / النسخة (Grosse)',
    description: 'النسخة الممهورة بالصيغة التنفيذية الرسمية — قابلة للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenterBold}">باسم الشعب الجزائري — نسخة تنفيذية</p>
      <p dir="rtl" style="${pStyle}">موجب العقد المحرر بتاريخ <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> تحت رقم <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> (<span class="smart-tag" data-var="مراجع_التسجيل">[مراجع_التسجيل]</span> — <span class="smart-tag" data-var="مراجع_الشهر">[مراجع_الشهر]</span>) بين <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span> و <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span> بخصوص الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> (<span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span>) بمبلغ <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span>.</p>
      <p dir="rtl" style="${pStyle}">وبناءً على ذلك، فإن الجمهورية الجزائرية الديمقراطية الشعبية تدعو وتأمر جميع المحضرين وكذا كل الأعوان الذين طُلب منهم ذلك أن ينفذوا هذا العقد، وعلى النواب العامين ووكلاء الجمهورية لدى المحاكم أن يمدوا يد المساعدة اللازمة لتنفيذه، وعلى جميع قادة وضباط القوة العمومية أن يمدوا يد المعونة بقوة القانون متى طُلب منهم ذلك بصفة قانونية.</p>
    `.trim(),
  },
  {
    id: 'derived_extrait_donation',
    code: 'extrait_donation',
    name: 'مستخرج الهبات (Extrait Donation)',
    description: 'مستخرج خاص بعقود الهبة والتبرعات — قابل للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">مستخرج عقد هبة</p>
      <p dir="rtl" style="${pStyle}">بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> — فهرس رقم: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> — <span class="smart-tag" data-var="مراجع_التسجيل">[مراجع_التسجيل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الواهب:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الأول_الإقامة">[الطرف_الأول_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الموهوب له:</span> <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الثاني_الإقامة">[الطرف_الثاني_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">العقار الموهوب:</span> <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">التقويم المالي للهبة:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
    `.trim(),
  },
  {
    id: 'derived_br4bis',
    code: 'br4bis',
    name: 'طلب PR4BIS (طلب معلومات عقارية)',
    description: 'استمارة طلب معلومات على الإجراءات المشهرة من المحافظة العقارية — قابلة للتعديل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">طلب معلومات على الإجراءات المشهرة (PR4BIS)</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">المحافظة العقارية:</span> <span class="smart-tag" data-var="المحافظة_العقارية">[المحافظة_العقارية]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">المالك الحالي / الطرف الأول:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، تاريخ ومكان الميلاد: <span class="smart-tag" data-var="الطرف_الأول_تاريخ_الميلاد">[الطرف_الأول_تاريخ_الميلاد]</span> بـ <span class="smart-tag" data-var="الطرف_الأول_مكان_الميلاد">[الطرف_الأول_مكان_الميلاد]</span>، النسب: <span class="smart-tag" data-var="الطرف_الأول_النسب">[الطرف_الأول_النسب]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">تعيين العقار المطلوب بشأنه المعلومات:</span> الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> — <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">مراجع الوصف التقسيمي / السند:</span> <span class="smart-tag" data-var="مراجع_الوصف_التقسيمي">[مراجع_الوصف_التقسيمي]</span></p>
    `.trim(),
  },
  {
    id: 'derived_recu_depot',
    code: 'recu_depot',
    name: 'وصل استلام الوديعة / ملف الإيداع (Reçu de Dépôt)',
    description: 'وصل إيداع الملف أو استلام المبالغ والوثائق المودعة بالمكتب قبل الفهرسة — قابل للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">وصل إيداع ملف واستلام وديعة</p>
      <p dir="rtl" style="${pStyle}">رقم الإيداع: <span class="smart-tag" data-var="رقم_الإيداع">[رقم_الإيداع]</span> — بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span></p>
      <p dir="rtl" style="${pStyle}">استلمنا من السيد(ة): <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الثاني_الإقامة">[الطرف_الثاني_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}">مبلغاً قدره: <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
      <p dir="rtl" style="${pStyle}">وذلك بخصوص المعاملة المتعلقة بالحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> مع السيد(ة) <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>.</p>
    `.trim(),
  },
];

/**
 * Formats the full notarial registration reference string for [مراجع_التسجيل].
 */
export function formatRegistrationReferenceTag(batch: {
  inspectorateName?: string;
  registrationDate?: string;
  receiptNumber?: string;
  totalFees?: number;
  chequeRef?: string;
}): string {
  if (!batch.receiptNumber && !batch.registrationDate) {
    return batch.inspectorateName
      ? `قيد التسجيل لدى مفتشية التسجيل والطابع بـ ${batch.inspectorateName}`
      : 'غير مسجل بعد';
  }
  const parts: string[] = ['مسجل'];
  if (batch.inspectorateName) {
    parts.push(`لدى مفتشية التسجيل والطابع بـ ${batch.inspectorateName}`);
  }
  if (batch.registrationDate) {
    parts.push(`بتاريخ ${batch.registrationDate}`);
  }
  if (batch.receiptNumber) {
    parts.push(`حسب وصل رقم ${batch.receiptNumber}`);
  }
  if (typeof batch.totalFees === 'number' && batch.totalFees > 0) {
    parts.push(`بحقوق قدرها ${batch.totalFees.toLocaleString('ar-DZ')} دج`);
  }
  return parts.join(' ');
}

/**
 * Formats the full notarial land publication reference string for [مراجع_الشهر].
 */
export function formatPublicationReferenceTag(track: {
  conservationOffice?: string;
  publicationDate?: string;
  volume?: string;
  number?: string;
  depositDate?: string;
}): string {
  if (!track.volume && !track.number && !track.publicationDate) {
    if (track.depositDate && track.conservationOffice) {
      return `مودع للشهر بالمحافظة العقارية بـ ${track.conservationOffice} بتاريخ ${track.depositDate}`;
    }
    return 'غير مشهر بعد';
  }
  const parts: string[] = ['مشهر'];
  if (track.conservationOffice) {
    parts.push(`بالمحافظة العقارية بـ ${track.conservationOffice}`);
  }
  if (track.publicationDate) {
    parts.push(`بتاريخ ${track.publicationDate}`);
  }
  if (track.volume) {
    parts.push(`مجلد ${track.volume}`);
  }
  if (track.number) {
    parts.push(`رقم ${track.number}`);
  }
  return parts.join(' ');
}

export const DEFAULT_NOTARY_CLERKS: NotaryClerk[] = [
  { id: 'notary-head', name: 'الموثق الرئيسي (الأستاذ)', role: 'notary', color: '#1e3a8a', createdAt: '2026-01-01T00:00:00.000Z' },
  { id: 'clerk-1', name: 'الكاتب الأول (أمين)', role: 'clerk', color: '#047857', createdAt: '2026-01-01T00:00:00.000Z' },
  { id: 'clerk-2', name: 'الكاتب الثاني (فاطمة)', role: 'clerk', color: '#b91c1c', createdAt: '2026-01-01T00:00:00.000Z' },
  { id: 'clerk-3', name: 'الكاتب الثالث (سفيان)', role: 'clerk', color: '#d97706', createdAt: '2026-01-01T00:00:00.000Z' },
];

/**
 * Computes canonical relative archive path: archive/{year}/{index}.docx
 */
export function buildArchiveRelativePath(year: number, indexNumber: string): string {
  const safeYear = Number.isFinite(year) && year > 1900 ? Math.floor(year) : new Date().getFullYear();
  const rawIdx = String(indexNumber || '')
    .trim()
    .replace(/^20\d{2}[\/\-_]/, '') // Strip leading year prefix if present (e.g. 2026/0142 -> 0142)
    .replace(/[^a-zA-Z0-9\u0600-\u06FF_-]/g, '_');
  const paddedIdx = /^\d+$/.test(rawIdx) ? rawIdx.padStart(4, '0') : rawIdx || '0001';
  return `archive/${safeYear}/${paddedIdx}.docx`;
}

/**
 * Computes SHA-256 hex digest of a Uint8Array using Web Crypto API (with pure-JS fallback).
 */
export async function computeSha256Hex(bytes: Uint8Array): Promise<string> {
  if (typeof globalThis !== 'undefined' && globalThis.crypto?.subtle) {
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength
    ) as ArrayBuffer;
    const hashBuffer = await globalThis.crypto.subtle.digest('SHA-256', buffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
  }
  // Fallback deterministic 256-bit hash if subtle crypto is unavailable in non-secure context
  let h1 = 0xdeadbeef ^ bytes.byteLength;
  let h2 = 0x41c6ce57 ^ bytes.byteLength;
  for (let i = 0; i < bytes.byteLength; i++) {
    const ch = bytes[i];
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  const part = (h1 >>> 0).toString(16).padStart(8, '0') + (h2 >>> 0).toString(16).padStart(8, '0');
  return part.repeat(4);
}

export function uint8ArrayToBase64(bytes: Uint8Array): string {
  if (typeof window === 'undefined' || typeof window.btoa !== 'function') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return window.btoa(binary);
}

export function base64ToUint8Array(base64: string): Uint8Array {
  if (typeof window === 'undefined' || typeof window.atob !== 'function') {
    return new Uint8Array(Buffer.from(base64, 'base64'));
  }
  const binary = window.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

export function extractTemplatePlaceholders(
  ...htmlFragments: Array<string | undefined>
): string[] {
  const combined = htmlFragments.filter(Boolean).join('\n');
  const set = new Set<string>();
  const dataVarRegex = /data-var=["']([^"']+)["']/g;
  let m: RegExpExecArray | null;
  while ((m = dataVarRegex.exec(combined)) !== null) {
    if (m[1]?.trim()) set.add(m[1].trim());
  }
  const bracketRegex = /\[([^\[\]\n]{1,60})\]/g;
  while ((m = bracketRegex.exec(combined)) !== null) {
    const candidate = m[1]?.trim();
    if (candidate && !candidate.startsWith('فهرس:')) {
      set.add(candidate);
    }
  }
  return Array.from(set);
}

export function inferKindFromDerivedCode(code?: string): OfficeTemplateKind {
  if (!code) return 'extract';
  if (code === 'publication' || code === 'br4bis') return 'registration';
  if (code === 'fiscal') return 'fiscal';
  if (code === 'acte') return 'contract';
  if (code === 'recu_depot') return 'other';
  return 'extract';
}

export function migrateDerivedTemplateToOfficeTemplate(
  tpl: DerivedDocTemplate
): OfficeTemplate {
  const kind = inferKindFromDerivedCode(tpl.code);
  const category =
    kind === 'registration'
      ? 'استمارات التسجيل والشهر'
      : kind === 'fiscal'
      ? 'الوضعية الجبائية'
      : kind === 'other'
      ? 'وثائق المكتب'
      : 'مستخرجات ونسخ';
  return {
    id: tpl.id,
    name: tpl.name,
    kind,
    category,
    description: tpl.description || '',
    code: tpl.code,
    docTypeCode: tpl.code,
    bodyHtml: tpl.bodyHtml || '',
    headerHtml: tpl.headerHtml || '',
    footerHtml: tpl.footerHtml || '',
    pageNumberingEnabled: true,
    extractedPlaceholders: extractTemplatePlaceholders(
      tpl.headerHtml,
      tpl.bodyHtml,
      tpl.footerHtml
    ),
    isDefault: tpl.id.startsWith('derived_'),
    updatedAt: tpl.updatedAt || '2026-01-01T00:00:00.000Z',
  };
}

export function migrateCustomTemplateToOfficeTemplate(
  tpl: CustomTemplate
): OfficeTemplate {
  const cat = tpl.category || 'عقود رسمية';
  const kind: OfficeTemplateKind =
    cat.includes('مستخرج') || cat.includes('نسخ')
      ? 'extract'
      : cat.includes('شهر') || cat.includes('تسجيل')
      ? 'registration'
      : cat.includes('جبائ')
      ? 'fiscal'
      : 'contract';
  return {
    id: tpl.id,
    name: tpl.name,
    kind,
    category: cat,
    description: tpl.description || '',
    docTypeCode: kind === 'contract' ? 'acte' : undefined,
    bodyHtml: tpl.bodyHtml || '',
    headerHtml: tpl.headerHtml || '',
    footerHtml: tpl.footerHtml || '',
    pageNumberingEnabled: tpl.pageNumberingEnabled !== false,
    extractedPlaceholders:
      Array.isArray(tpl.extractedPlaceholders) &&
      tpl.extractedPlaceholders.length > 0
        ? tpl.extractedPlaceholders
        : extractTemplatePlaceholders(
            tpl.headerHtml,
            tpl.bodyHtml,
            tpl.footerHtml
          ),
    defaultFieldValues: tpl.defaultFieldValues,
    sourceFileName: tpl.sourceFileName,
    updatedAt: tpl.updatedAt || new Date().toISOString(),
  };
}

export function migrateWordTemplateToOfficeTemplate(
  tpl: WordTemplateDefinition
): OfficeTemplate {
  const kind: OfficeTemplateKind =
    tpl.type === 'original' || tpl.type === 'copy'
      ? 'contract'
      : tpl.type === 'extract'
      ? 'extract'
      : tpl.type === 'registration'
      ? 'registration'
      : tpl.type === 'fiscal'
      ? 'fiscal'
      : 'other';
  const category =
    kind === 'contract'
      ? 'عقود رسمية'
      : kind === 'registration'
      ? 'استمارات التسجيل والشهر'
      : kind === 'fiscal'
      ? 'الوضعية الجبائية'
      : 'مستخرجات ونسخ';
  return {
    id: tpl.id,
    name: tpl.name,
    kind,
    category,
    description: tpl.description || '',
    docTypeCode: tpl.type,
    bodyHtml: tpl.bodyHtml || '',
    headerHtml: tpl.headerHtml || '',
    footerHtml: tpl.footerHtml || '',
    pageNumberingEnabled: true,
    extractedPlaceholders: extractTemplatePlaceholders(
      tpl.headerHtml,
      tpl.bodyHtml,
      tpl.footerHtml
    ),
    isDefault: Boolean(tpl.isDefault),
    updatedAt: tpl.updatedAt || '2026-01-01T00:00:00.000Z',
  };
}

export function officeTemplateToCustomTemplate(
  tpl: OfficeTemplate
): CustomTemplate {
  return {
    id: tpl.id,
    name: tpl.name,
    category: tpl.category || 'عقود رسمية',
    description: tpl.description || '',
    bodyHtml: tpl.bodyHtml || '',
    headerHtml: tpl.headerHtml || '',
    footerHtml: tpl.footerHtml || '',
    pageNumberingEnabled: tpl.pageNumberingEnabled !== false,
    extractedPlaceholders: tpl.extractedPlaceholders || [],
    defaultFieldValues: tpl.defaultFieldValues,
    sourceFileName: tpl.sourceFileName,
    updatedAt: tpl.updatedAt,
  };
}

export function officeTemplateToDerivedDocTemplate(
  tpl: OfficeTemplate
): DerivedDocTemplate {
  return {
    id: tpl.id,
    code: tpl.docTypeCode || (tpl.kind === 'registration' ? 'publication' : 'extrait'),
    name: tpl.name,
    description: tpl.description || '',
    bodyHtml: tpl.bodyHtml || '',
    headerHtml: tpl.headerHtml || '',
    footerHtml: tpl.footerHtml || '',
    updatedAt: tpl.updatedAt,
  };
}

export function normalizeAnyTemplateToOfficeTemplate(
  raw: unknown
): OfficeTemplate | null {
  if (!raw || typeof raw !== 'object') return null;
  const obj = raw as Record<string, unknown>;
  if (typeof obj.name !== 'string' || !obj.name.trim()) return null;

  if (
    typeof obj.kind === 'string' &&
    ['contract', 'extract', 'registration', 'fiscal', 'other'].includes(obj.kind)
  ) {
    const t = obj as unknown as OfficeTemplate;
    return {
      id:
        t.id ||
        `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      name: t.name,
      kind: t.kind,
      category: t.category || 'عقود رسمية',
      description: t.description || '',
      docTypeCode: t.docTypeCode,
      bodyHtml: t.bodyHtml || '',
      headerHtml: t.headerHtml || '',
      footerHtml: t.footerHtml || '',
      pageNumberingEnabled: t.pageNumberingEnabled !== false,
      extractedPlaceholders: Array.isArray(t.extractedPlaceholders)
        ? t.extractedPlaceholders
        : extractTemplatePlaceholders(t.headerHtml, t.bodyHtml, t.footerHtml),
      defaultFieldValues: t.defaultFieldValues,
      sourceFileName: t.sourceFileName,
      isDefault: Boolean(t.isDefault),
      updatedAt: t.updatedAt || new Date().toISOString(),
    };
  }

  if (typeof obj.code === 'string') {
    return migrateDerivedTemplateToOfficeTemplate(
      obj as unknown as DerivedDocTemplate
    );
  }

  if (typeof obj.type === 'string') {
    return migrateWordTemplateToOfficeTemplate(
      obj as unknown as WordTemplateDefinition
    );
  }

  return migrateCustomTemplateToOfficeTemplate(
    obj as unknown as CustomTemplate
  );
}

export const DEFAULT_OFFICE_TEMPLATES: OfficeTemplate[] = [
  ...DEFAULT_DERIVED_DOC_TEMPLATES.map(migrateDerivedTemplateToOfficeTemplate),
  {
    id: 'derived_fiscal_card',
    name: 'بطاقة الوضعية الجبائية للعقد (Fiscal Position)',
    kind: 'fiscal',
    category: 'الوضعية الجبائية',
    docTypeCode: 'fiscal',
    description:
      'بطاقة تلخيص العناصر الجبائية الخمسة للعقد التوثيقي وأصل الملكية — قابلة للتعديل والتصدير',
    headerHtml: '',
    footerHtml: '',
    pageNumberingEnabled: true,
    isDefault: true,
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">بطاقة الوضعية الجبائية للعقد</p>
      <p dir="rtl" style="${pStyle}">فهرس رقم: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> — بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> — مفتشية التسجيل: <span class="smart-tag" data-var="مفتشية_التسجيل">[مفتشية_التسجيل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">1. تعيين المبيع الجبائي:</span> <span class="smart-tag" data-var="تعيين_المبيع_الجبائي">[تعيين_المبيع_الجبائي]</span> — <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">2. أصل الملكية وثمن الاقتناء السابق:</span> <span class="smart-tag" data-var="أصل_الملكية_وثمن_الشراء_السابق">[أصل_الملكية_وثمن_الشراء_السابق]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">3. البيانات الجبائية للبائع:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span> — <span class="smart-tag" data-var="بيانات_البائع_الجبائية">[بيانات_البائع_الجبائية]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">4. البيانات الجبائية للمشتري:</span> <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span> — <span class="smart-tag" data-var="بيانات_المشتري_الجبائية">[بيانات_المشتري_الجبائية]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">5. الثمن وطريقة الدفع:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>) — <span class="smart-tag" data-var="طريقة_الدفع">[طريقة_الدفع]</span>.</p>
    `.trim(),
    extractedPlaceholders: [
      'رقم_الفهرس',
      'تاريخ_العقد',
      'مفتشية_التسجيل',
      'تعيين_المبيع_الجبائي',
      'تعيين_الحصة_الكامل',
      'أصل_الملكية_وثمن_الشراء_السابق',
      'الطرف_الأول_الاسم',
      'بيانات_البائع_الجبائية',
      'الطرف_الثاني_الاسم',
      'بيانات_المشتري_الجبائية',
      'الثمن_بالأحرف',
      'الثمن_بالأرقام',
      'طريقة_الدفع',
    ],
  },
];

export function mergeLegacyStoresIntoOfficeTemplates(params: {
  existingOffice: OfficeTemplate[];
  legacyCustom?: CustomTemplate[];
  legacyDerived?: DerivedDocTemplate[];
  legacyWord?: WordTemplateDefinition[];
  seedDefaults?: boolean;
}): OfficeTemplate[] {
  const byId = new Map<string, OfficeTemplate>();

  if (params.seedDefaults !== false) {
    for (const defTpl of DEFAULT_OFFICE_TEMPLATES) {
      byId.set(defTpl.id, defTpl);
    }
  }

  for (const d of params.legacyDerived || []) {
    if (d && d.id && d.name) {
      byId.set(d.id, migrateDerivedTemplateToOfficeTemplate(d));
    }
  }

  for (const w of params.legacyWord || []) {
    if (w && w.id && w.name) {
      // Skip duplicate default word templates that overlap with derived_extrait / derived_publication
      if (
        w.isDefault &&
        (w.id === 'tpl_extract' ||
          w.id === 'tpl_registration' ||
          w.id === 'tpl_fiscal')
      ) {
        continue;
      }
      byId.set(w.id, migrateWordTemplateToOfficeTemplate(w));
    }
  }

  for (const c of params.legacyCustom || []) {
    if (c && c.id && c.name) {
      byId.set(c.id, migrateCustomTemplateToOfficeTemplate(c));
    }
  }

  for (const o of params.existingOffice || []) {
    if (o && o.id && o.name) {
      byId.set(o.id, o);
    }
  }

  return Array.from(byId.values()).sort((a, b) =>
    (b.updatedAt || '').localeCompare(a.updatedAt || '')
  );
}
