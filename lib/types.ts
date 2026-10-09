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
  category: 'party1' | 'party2' | 'property' | 'financial' | 'custom';
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
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  pageNumberingEnabled: boolean;
  extractedPlaceholders: string[];
  sourceFileName?: string;
  updatedAt: string;
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
  selectedEstateId?: string;
  selectedLotNumber?: string;
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

export interface BackupBundle {
  version: string;
  exportedAt: string;
  templates: CustomTemplate[];
  documents: SavedDocument[];
  estates: SubdivisionEstate[];
  clauses: NotaryClause[];
  derivedTemplates: DerivedDocTemplate[];
  revisions: DocumentRevision[];
  downloads: DownloadArchiveItem[];
  defaultFields: PartyField[];
}
