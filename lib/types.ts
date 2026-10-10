export const STRICT_FONT_FAMILY = 'Arial';
export const STRICT_FONT_SIZE_PT = 13;
export const STRICT_FONT_SIZE_HALF_PT = 26; // 13pt * 2 in OpenXML half-points
export const STRICT_LINE_SPACING_TWIPS = 240; // 1.0 single line spacing in 240ths of a line

// 1 cm = 566.929133858 twips
export const STRICT_MARGINS_CM = {
  top: 1,
  right: 7,
  bottom: 6,
  left: 2,
} as const;

export const STRICT_MARGINS_TWIPS = {
  top: 567, // 1 cm
  right: 3969, // 7 cm
  bottom: 3402, // 6 cm
  left: 1134, // 2 cm
  header: 284, // 0.5 cm
  footer: 1134, // 2 cm
} as const;

export const A4_SIZE_TWIPS = {
  width: 11906, // 21 cm
  height: 16838, // 29.7 cm
} as const;

// 11906 - (3969 + 1134) = 6803 twips = 12 cm net printable width
export const NET_CONTENT_WIDTH_TWIPS = 6803;
export const NET_CONTENT_WIDTH_PX = 453;

export type ListNumberingStyle = 'decimal' | 'arabic-alpha' | 'arabic-abjad' | 'bullet' | 'dash';

export type VariableInputType = 'text' | 'date' | 'number' | 'select';

export interface PartyField {
  key: string; // e.g., "الطرف_الأول_الاسم"
  label: string;
  value: string;
  category: 'party1' | 'party2' | 'property' | 'financial' | 'fiscal' | 'index' | 'custom';
  inputType?: VariableInputType;
  options?: string[];
}

export interface SubdivisionLot {
  id: string;
  lotNumber: string; // رقم الحصة
  building: string; // العمارة / المدخل
  floor: string; // الطابق
  nature: string; // طبيعة الحصة (شقة، محل، مرآب...)
  area: string; // المساحة بالمتر المربع
  commonShares: string; // الحصة في الأجزاء المشتركة (بالألفية / العشرية)
  boundaries: string; // الحدود
  fullDescription: string; // نص التعيين الكامل للحصة
}

export interface SubdivisionEstate {
  id: string;
  estateName: string; // اسم العقار أو الترقية العقارية
  commune: string; // البلدية
  wilaya: string; // الولاية
  cadastralSection: string; // القسم المساحي
  cadastralPlot: string; // مجموعة الملكية
  landArea: string; // المساحة الإجمالية للأرضية
  subdivisionDeedRef: string; // مراجع عقد الوصف التقسيمي وشهره
  lots: SubdivisionLot[];
  updatedAt: string;
}

export interface NotaryClause {
  id: string;
  title: string;
  category: string;
  contentHtml: string;
  enabled: boolean;
  order: number;
  updatedAt: string;
}

export type DerivedDocCode =
  | 'acte'
  | 'extrait'
  | 'publication'
  | 'attestation'
  | 'proce_verbal'
  | 'extrait_donation'
  | 'br4bis'
  | 'recu_depot'
  | string;

export interface DerivedDocTemplate {
  id: string;
  code: DerivedDocCode;
  name: string;
  description: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  updatedAt: string;
}

export type WordTemplateType = 'original' | 'copy' | 'extract' | 'registration' | 'fiscal' | string;

export interface WordTemplateDefinition {
  id: string;
  name: string;
  type: WordTemplateType;
  description: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  isDefault?: boolean;
  updatedAt: string;
}

export interface DocumentRevision {
  id: string;
  documentId: string;
  documentTitle: string;
  author: string;
  bodyHtml: string;
  fieldValues: Record<string, string>;
  createdAt: string;
  summary: string;
}

export interface DownloadArchiveItem {
  id: string;
  documentId: string;
  documentTitle: string;
  docTypeCode: DerivedDocCode;
  docTypeLabel: string;
  fileName: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  fieldValues: Record<string, string>;
  createdAt: string;
}

export interface CustomTemplate {
  id: string;
  name: string;
  category: string;
  description?: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  pageNumberingEnabled: boolean;
  extractedPlaceholders: string[];
  defaultFieldValues?: Record<string, string>;
  sourceFileName?: string;
  updatedAt: string;
}

export interface NotaryClerk {
  id: string;
  name: string;
  role: 'notary' | 'clerk';
  color: string;
  createdAt: string;
}

export interface ContractFolder {
  id: string;
  name: string;
  parentId: string | null; // null for root level
  clerkId: string;
  createdAt: string;
  updatedAt: string;
}

export interface SavedContractDerivedDoc {
  id: string;
  typeId: DerivedDocCode;
  title: string;
  content: string;
  headerHtml?: string;
  footerHtml?: string;
  fieldValues?: Record<string, string>;
  updatedAt: string;
}

export type NotaryPartyRole = 'بائع' | 'مشتري' | 'موكل' | 'وكيل' | 'واهب' | 'موهوب' | 'أخرى';

export interface ContractPartyCard {
  id: string;
  role: NotaryPartyRole;
  customRoleLabel?: string; // عند اختيار 'أخرى' (مثل: شريك، مقاسم، مؤجر)
  index: number; // 1, 2, 3... للسماح بـ بائع 1، بائع 2
  fullName: string;
  birthDate: string;
  birthPlace: string;
  filiation: string; // ابن فلان وفلانة
  nationalIdNin: string; // رقم التعريف الوطني NIN
  idCardDetails: string; // رقم وتاريخ وجهة صدور بطاقة الهوية
  address: string;
  legalRepresentative?: string;
}

export interface ClauseCondition {
  field: string; // اسم الشرط الذي يحدده المكتب (يترك فارغاً ليملأ حسب شروط المكتب)
  value: string | boolean;
}

export interface ContractOutlineClause {
  id: string;
  index: number;
  order: number;
  title: string;
  previewText: string;
  variables: string[];
  contentHtml: string;
  enabled: boolean;
  locked: boolean;
  domIndex: number;
  isContainer: boolean;
  condition?: ClauseCondition;
}

export interface SavedDocument {
  id: string;
  title: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  pageNumberingEnabled: boolean;
  showHeaderFooter: boolean;
  fieldValues: Record<string, string>;
  fieldInputTypes?: Record<string, VariableInputType>;
  outlineClauses?: ContractOutlineClause[];
  contractConditions?: Record<string, string | boolean>;
  partyCards?: ContractPartyCard[];
  selectedEstateId?: string;
  selectedLotNumber?: string;
  folderId?: string | null;
  clerkId?: string;
  clerkName?: string;
  contractNumber?: string;
  year?: number;
  clientName?: string;
  modelId?: string;
  modelTitle?: string;
  status?: 'draft' | 'editing' | 'ready_to_sign' | 'signed' | 'registered';
  derivedDocuments?: SavedContractDerivedDoc[];
  updatedAt: string;
  createdAt: string;
}

export interface SerializedSelectionPath {
  zone: 'body' | 'header' | 'footer';
  startPath: number[];
  startOffset: number;
  endPath: number[];
  endOffset: number;
}

export interface HistorySnapshot {
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  selection: SerializedSelectionPath | null;
  timestamp: number;
}

export interface ToolbarState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  align: 'right' | 'center' | 'left' | 'justify';
  dir: 'rtl' | 'ltr';
  textColor: string;
  highlightColor: string;
  listType: 'none' | ListNumberingStyle;
  inTable: boolean;
  canMergeCells: boolean;
  canSplitCell: boolean;
}

export interface SavedPartyRecord {
  id: string;
  fullName: string; // الاسم واللقب
  birthDate: string; // تاريخ الميلاد
  birthPlace: string; // مكان الميلاد
  parentage: string; // النسب (بن/بنت)
  idCardRef: string; // رقم ونوع وثيقة الهوية
  address: string; // عنوان الإقامة
  updatedAt: string;
}

export interface SavedPropertyRecord {
  id: string;
  label: string; // تسمية مختصرة للعقار/الحصة
  lotNumber: string; // رقم الحصة
  nature: string; // طبيعة الحصة
  floor: string; // الطابق والعمارة
  area: string; // المساحة
  commonShares: string; // الأجزاء المشتركة
  fullDescription: string; // التعيين الكامل
  deedRef: string; // مراجع الوصف التقسيمي / السند
  updatedAt: string;
}

export interface ClauseVariableGroup {
  clauseId: string;
  clauseTitle: string;
  variables: string[];
  clauseReferenceText?: string;
}

export interface SealedContractOriginal {
  documentId: string;
  relativePath: string; // مثال: archive/2026/0142.docx — مسار نسبي داخل مجلد المكتب، ليس مساراً مطلقاً
  sha256: string;
  sealedAt: string;
  byteLength: number;
}

export interface SerializedSealedOriginal extends SealedContractOriginal {
  base64Data?: string;
}

export interface BackupBundle {
  version: string;
  exportedAt: string;
  clerks?: NotaryClerk[];
  folders?: ContractFolder[];
  templates: CustomTemplate[];
  documents: SavedDocument[];
  estates: SubdivisionEstate[];
  clauses: NotaryClause[];
  derivedTemplates: DerivedDocTemplate[];
  revisions: DocumentRevision[];
  downloads: DownloadArchiveItem[];
  savedParties?: SavedPartyRecord[];
  savedProperties?: SavedPropertyRecord[];
  defaultFields: PartyField[];
  sealedOriginals?: SerializedSealedOriginal[];
}


