'use client';

import React, { useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  Building2,
  CheckSquare,
  Download,
  Edit3,
  FileOutput,
  FilePlus2,
  FileText,
  FileUp,
  FolderArchive,
  FormInput,
  GitCompare,
  History,
  ListChecks,
  Plus,
  RotateCcw,
  Save,
  Square,
  Trash2,
  Upload,
  UserCheck,
} from 'lucide-react';
import {
  CustomTemplate,
  DerivedDocTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  NotaryClause,
  PartyField,
  SavedDocument,
  SubdivisionEstate,
  SubdivisionLot,
} from '@/lib/types';

export type SidebarTab = 'clauses' | 'parties' | 'templates' | 'documents';

interface SidebarWorkspaceProps {
  activeTab: SidebarTab;
  onSelectTab: (tab: SidebarTab) => void;
  // 1. Ready Clauses (البنود الجاهزة — التبويب الأول الأكثر استخداماً)
  clauses: NotaryClause[];
  activeClauseIdsInDoc: string[];
  onToggleClauseInDoc: (clause: NotaryClause) => void;
  onInsertClauseAtCaret: (clause: NotaryClause) => void;
  onMoveClauseOrder: (clauseId: string, direction: 'up' | 'down') => Promise<void>;
  onSaveNewClause: (title: string, category: string, contentHtml?: string) => Promise<void>;
  onSaveSelectionAsClause: (title: string, category: string) => Promise<void>;
  onEditClauseInEditor: (clause: NotaryClause) => void;
  onDeleteClause: (id: string) => Promise<void>;
  // 2. Parties & Designations + Subdivision Table
  partyFields: PartyField[];
  extractedPlaceholders: string[];
  unfilledCount: number;
  fieldValues: Record<string, string>;
  onOpenSmartVariablesModal: (focusKey?: string) => void;
  onUpdateFieldValue: (key: string, value: string) => void;
  onAddCustomField: (key: string, label: string) => void;
  onInsertPlaceholderAtCaret: (key: string) => void;
  onInsertValueAtCaret: (value: string) => void;
  onBakeAllPlaceholdersIntoDocument: () => void;
  estates: SubdivisionEstate[];
  selectedEstateId: string;
  selectedLotNumber: string;
  onSelectEstateAndLot: (estateId: string, lotNumber: string) => void;
  onSaveEstate: (estate: SubdivisionEstate) => Promise<void>;
  onDeleteEstate: (id: string) => Promise<void>;
  onInsertLotClauseAtCaret: (estate: SubdivisionEstate, lot: SubdivisionLot) => void;
  onInsertFullSubdivisionTableAtCaret: (estate: SubdivisionEstate) => void;
  // 3. Office Templates & Editable Derived Document Templates
  templates: CustomTemplate[];
  onImportDocxAsTemplate: (files: FileList) => Promise<void>;
  onSaveCurrentAsTemplate: (name: string, category: string) => Promise<void>;
  onLoadTemplateFull: (tpl: CustomTemplate) => void;
  onInsertTemplateAtCaret: (tpl: CustomTemplate) => void;
  onExportTemplateDocx: (tpl: CustomTemplate) => Promise<void>;
  onDeleteTemplate: (id: string) => Promise<void>;
  derivedTemplates: DerivedDocTemplate[];
  activeEditingDerivedId: string | null;
  onGenerateDerivedDoc: (tpl: DerivedDocTemplate) => void;
  onEditDerivedTemplateInEditor: (tpl: DerivedDocTemplate) => void;
  onSaveCurrentAsDerivedTemplate: (name: string, code: string, description: string) => Promise<void>;
  onImportDocxForDerivedTemplate: (tpl: DerivedDocTemplate, file: File) => Promise<void>;
  onDeleteDerivedTemplate: (id: string) => Promise<void>;
  // 4. Saved Documents + Version History (Diff) + Download Archive
  documents: SavedDocument[];
  revisions: DocumentRevision[];
  downloads: DownloadArchiveItem[];
  activeDocumentId: string;
  onOpenDocument: (doc: SavedDocument) => void;
  onOpenMultiSourceNewModal: () => void;
  onSaveCurrentDocument: () => Promise<void>;
  onDeleteDocument: (id: string) => Promise<void>;
  onExportDocumentDocx: (doc: SavedDocument) => Promise<void>;
  onOpenVersionDiffModal: () => void;
  onRestoreRevision: (rev: DocumentRevision) => void;
  onOpenDownloadArchiveItemInEditor: (item: DownloadArchiveItem) => void;
  onRedownloadArchiveItemDocx: (item: DownloadArchiveItem) => Promise<void>;
  onDeleteDownloadArchiveItem: (id: string) => Promise<void>;
  onExportBackupJson: () => Promise<void>;
  onImportBackupJson: (file: File) => Promise<void>;
  onSaveSelectionBookmark: () => void;
}

export default function SidebarWorkspace({
  activeTab,
  onSelectTab,
  clauses,
  activeClauseIdsInDoc,
  onToggleClauseInDoc,
  onInsertClauseAtCaret,
  onMoveClauseOrder,
  onSaveNewClause,
  onSaveSelectionAsClause,
  onEditClauseInEditor,
  onDeleteClause,
  partyFields,
  extractedPlaceholders,
  unfilledCount,
  fieldValues,
  onOpenSmartVariablesModal,
  onUpdateFieldValue,
  onAddCustomField,
  onInsertPlaceholderAtCaret,
  onInsertValueAtCaret,
  onBakeAllPlaceholdersIntoDocument,
  estates,
  selectedEstateId,
  selectedLotNumber,
  onSelectEstateAndLot,
  onSaveEstate,
  onDeleteEstate,
  onInsertLotClauseAtCaret,
  onInsertFullSubdivisionTableAtCaret,
  templates,
  onImportDocxAsTemplate,
  onSaveCurrentAsTemplate,
  onLoadTemplateFull,
  onInsertTemplateAtCaret,
  onExportTemplateDocx,
  onDeleteTemplate,
  derivedTemplates,
  activeEditingDerivedId,
  onGenerateDerivedDoc,
  onEditDerivedTemplateInEditor,
  onSaveCurrentAsDerivedTemplate,
  onImportDocxForDerivedTemplate,
  onDeleteDerivedTemplate,
  documents,
  revisions,
  downloads,
  activeDocumentId,
  onOpenDocument,
  onOpenMultiSourceNewModal,
  onSaveCurrentDocument,
  onDeleteDocument,
  onExportDocumentDocx,
  onOpenVersionDiffModal,
  onRestoreRevision,
  onOpenDownloadArchiveItemInEditor,
  onRedownloadArchiveItemDocx,
  onDeleteDownloadArchiveItem,
  onExportBackupJson,
  onImportBackupJson,
  onSaveSelectionBookmark,
}: SidebarWorkspaceProps) {
  // Clause creation state
  const [newClauseTitle, setNewClauseTitle] = useState('');
  const [newClauseCategory, setNewClauseCategory] = useState('بنود عامة');
  const [newClauseText, setNewClauseText] = useState('');
  const [clauseSearch, setClauseSearch] = useState('');

  // Parties sub-view: 'fields' vs 'subdivision'
  const [partiesSubTab, setPartiesSubTab] = useState<'fields' | 'subdivision'>('fields');
  const [newFieldKey, setNewFieldKey] = useState('');

  // Subdivision Estate form state
  const [editingEstate, setEditingEstate] = useState<SubdivisionEstate | null>(null);
  const [newLotNumber, setNewLotNumber] = useState('');
  const [newLotBuilding, setNewLotBuilding] = useState('');
  const [newLotFloor, setNewLotFloor] = useState('');
  const [newLotNature, setNewLotNature] = useState('شقة سكنية');
  const [newLotArea, setNewLotArea] = useState('');
  const [newLotShares, setNewLotShares] = useState('');
  const [newLotDesc, setNewLotDesc] = useState('');

  // Templates sub-view: 'office' vs 'derived'
  const [templatesSubTab, setTemplatesSubTab] = useState<'office' | 'derived'>('derived');
  const [newTplName, setNewTplName] = useState('');
  const [newTplCategory, setNewTplCategory] = useState('عقود المكتب');
  const [newDerivedName, setNewDerivedName] = useState('');
  const [newDerivedCode, setNewDerivedCode] = useState('');
  const [newDerivedDesc, setNewDerivedDesc] = useState('');

  // Documents sub-view: 'saved' | 'revisions' | 'downloads'
  const [docsSubTab, setDocsSubTab] = useState<'saved' | 'revisions' | 'downloads'>('saved');
  const [docSearch, setDocSearch] = useState('');

  const selectedEstate = estates.find((e) => e.id === selectedEstateId) || estates[0] || null;
  const selectedLot =
    selectedEstate?.lots.find((l) => l.lotNumber === selectedLotNumber) || null;

  const knownKeys = new Set(partyFields.map((f) => f.key));
  const extraDocumentKeys = extractedPlaceholders.filter((k) => !knownKeys.has(k));

  const startNewEstate = () => {
    setEditingEstate({
      id: `est_${Date.now()}`,
      estateName: '',
      commune: '',
      wilaya: '',
      cadastralSection: '',
      cadastralPlot: '',
      landArea: '',
      subdivisionDeedRef: '',
      lots: [],
      updatedAt: new Date().toISOString(),
    });
  };

  const handleAddLotToEditingEstate = () => {
    if (!editingEstate || !newLotNumber.trim()) return;
    const autoDesc =
      newLotDesc.trim() ||
      `الحصة رقم (${newLotNumber.trim()}) المتمثلة في ${newLotNature.trim()}${
        newLotFloor.trim() ? ` الواقعة في ${newLotFloor.trim()}` : ''
      }${newLotBuilding.trim() ? ` بـ ${newLotBuilding.trim()}` : ''}${
        newLotArea.trim() ? `، تتربع على مساحة (${newLotArea.trim()} م²)` : ''
      }${
        newLotShares.trim()
          ? `، وما ينوبها في الأجزاء المشتركة والأرض بـ (${newLotShares.trim()})`
          : ''
      }.`;

    const lot: SubdivisionLot = {
      id: `lot_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      lotNumber: newLotNumber.trim(),
      building: newLotBuilding.trim(),
      floor: newLotFloor.trim(),
      nature: newLotNature.trim(),
      area: newLotArea.trim(),
      commonShares: newLotShares.trim(),
      boundaries: '',
      fullDescription: autoDesc,
    };

    setEditingEstate({
      ...editingEstate,
      lots: [...editingEstate.lots, lot],
      updatedAt: new Date().toISOString(),
    });

    setNewLotNumber('');
    setNewLotArea('');
    setNewLotShares('');
    setNewLotDesc('');
  };

  return (
    <aside className="w-full lg:w-88 xl:w-96 bg-white border-l border-slate-200 flex flex-col h-full shrink-0 select-none no-print">
      {/* Sidebar Segmented Navigation — Ordered by Daily Usage Frequency */}
      <div className="grid grid-cols-4 gap-1 p-2 bg-slate-50 border-b border-slate-200">
        {/* 1. البنود الجاهزة (الأكثر استخداماً أثناء التحرير) */}
        <button
          type="button"
          onClick={() => onSelectTab('clauses')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'clauses'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <ListChecks className="w-3.5 h-3.5 mb-0.5" />
          <span>البنود الجاهزة</span>
        </button>

        {/* 2. الأطراف والتعيينات */}
        <button
          type="button"
          onClick={() => onSelectTab('parties')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap relative ${
            activeTab === 'parties'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <UserCheck className="w-3.5 h-3.5 mb-0.5" />
          <span>الأطراف والتعيين</span>
          {unfilledCount > 0 && (
            <span className="absolute top-1 left-1 w-2 h-2 rounded-full bg-red-600" />
          )}
        </button>

        {/* 3. قوالب المكتب والمشتقات */}
        <button
          type="button"
          onClick={() => onSelectTab('templates')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'templates'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <FileOutput className="w-3.5 h-3.5 mb-0.5" />
          <span>القوالب والمشتقات</span>
        </button>

        {/* 4. المستندات المحفوظة والأرشيف */}
        <button
          type="button"
          onClick={() => onSelectTab('documents')}
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'documents'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <FolderArchive className="w-3.5 h-3.5 mb-0.5" />
          <span>المستندات والأرشيف</span>
        </button>
      </div>

      {/* Tab Content Area */}
      <div className="flex-1 overflow-y-auto p-3 space-y-4">
        {/* =========================================================
            TAB 1: READY CLAUSES SYSTEM (نظام البنود الجاهزة)
           ========================================================= */}
        {activeTab === 'clauses' && (
          <div className="space-y-4">
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-md space-y-2.5">
              <div className="text-xs font-bold text-slate-900">
                إضافة بند جديد إلى مكتبة بنود المكتب
              </div>
              <input
                type="text"
                value={newClauseTitle}
                onChange={(e) => setNewClauseTitle(e.target.value)}
                placeholder="عنوان البند (مثال: بند أصل الملكية / بند الحالة المدنية)"
                className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
              />
              <div className="flex items-center gap-2">
                <input
                  type="text"
                  value={newClauseCategory}
                  onChange={(e) => setNewClauseCategory(e.target.value)}
                  placeholder="التصنيف (بنود عامة، بيع، هبة...)"
                  className="flex-1 px-2.5 py-1 text-xs bg-white border border-slate-300 rounded"
                />
              </div>
              <textarea
                rows={3}
                value={newClauseText}
                onChange={(e) => setNewClauseText(e.target.value)}
                placeholder="اكتب نص البند هنا (يمكنك تضمين وسوم مثل {{الطرف_الأول_الاسم}}) أو اتركه فارغاً لحفظ النص المحدد حالياً داخل ورقة A4..."
                className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
              />
              <div className="grid grid-cols-2 gap-1.5">
                <button
                  type="button"
                  onClick={async () => {
                    if (!newClauseTitle.trim()) return;
                    const html = newClauseText.trim()
                      ? newClauseText
                          .trim()
                          .split(/\r?\n/)
                          .map(
                            (line) =>
                              `<p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;text-align:justify;">${
                                line.trim() || '<br>'
                              }</p>`
                          )
                          .join('')
                      : undefined;
                    await onSaveNewClause(
                      newClauseTitle.trim(),
                      newClauseCategory.trim() || 'عام',
                      html
                    );
                    setNewClauseTitle('');
                    setNewClauseText('');
                  }}
                  className="py-1.5 px-2 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800 transition-colors"
                >
                  + حفظ البند المكتوب
                </button>
                <button
                  type="button"
                  onMouseDown={(e) => {
                    e.preventDefault();
                    onSaveSelectionBookmark();
                  }}
                  onClick={async () => {
                    if (!newClauseTitle.trim()) return;
                    await onSaveSelectionAsClause(
                      newClauseTitle.trim(),
                      newClauseCategory.trim() || 'عام'
                    );
                    setNewClauseTitle('');
                  }}
                  className="py-1.5 px-2 bg-slate-900 text-white text-[11px] font-medium rounded hover:bg-slate-800 transition-colors"
                  title="يحفظ الفقرات أو الجداول المظللة حالياً داخل المحرر كبند جاهز"
                >
                  حفظ المحدد بالمحرر كبند
                </button>
              </div>
            </div>

            <input
              type="text"
              value={clauseSearch}
              onChange={(e) => setClauseSearch(e.target.value)}
              placeholder="بحث سريع في البنود الجاهزة..."
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded"
            />

            <div className="space-y-2">
              <div className="flex items-center justify-between border-b border-slate-200 pb-1.5">
                <span className="text-xs font-bold text-slate-900">
                  قائمة البنود الجاهزة ({clauses.length})
                </span>
                <span className="text-[10px] text-slate-500">
                  تفعيل ☑ يدرج البند في حاوية مرئية للتحرير فقط
                </span>
              </div>

              {clauses.length === 0 ? (
                <div className="border border-dashed border-slate-300 rounded-md p-5 text-center space-y-2 bg-slate-50/50">
                  <ListChecks className="w-7 h-7 text-slate-400 mx-auto" />
                  <div className="text-xs font-semibold text-slate-700">
                    لا توجد بنود محفوظة بعد
                  </div>
                  <p className="text-[11px] text-slate-500 leading-relaxed">
                    أضف بنود مكتبك المتكررة (مثل بند التكليف، أصل الملكية، التسليم، الضمان...) لتتمكن من تفعيلها أو تعطيلها وترتيبها بضغطة زر داخل أي عقد.
                  </p>
                </div>
              ) : (
                <div className="space-y-2">
                  {clauses
                    .filter(
                      (c) =>
                        !clauseSearch.trim() ||
                        c.title.includes(clauseSearch.trim()) ||
                        c.category.includes(clauseSearch.trim())
                    )
                    .map((clause, idx) => {
                      const isActiveInDoc = activeClauseIdsInDoc.includes(clause.id);
                      return (
                        <div
                          key={clause.id}
                          className={`border rounded-md p-2.5 transition-colors space-y-2 ${
                            isActiveInDoc
                              ? 'border-blue-400 bg-blue-50/30'
                              : 'border-slate-200 bg-white hover:border-slate-300'
                          }`}
                        >
                          <div className="flex items-start justify-between gap-2">
                            <button
                              type="button"
                              onMouseDown={(e) => {
                                e.preventDefault();
                                onSaveSelectionBookmark();
                              }}
                              onClick={() => onToggleClauseInDoc(clause)}
                              className="flex items-start gap-2 text-right flex-1"
                              title="تفعيل أو إزالة حاوية هذا البند داخل ورقة الـ A4"
                            >
                              {isActiveInDoc ? (
                                <CheckSquare className="w-4 h-4 text-blue-900 shrink-0 mt-0.5" />
                              ) : (
                                <Square className="w-4 h-4 text-slate-400 shrink-0 mt-0.5" />
                              )}
                              <div>
                                <div className="text-xs font-bold text-slate-900">
                                  {clause.title}
                                </div>
                                <div className="text-[10px] text-slate-500">
                                  {clause.category}
                                </div>
                              </div>
                            </button>

                            <div className="flex items-center gap-0.5 shrink-0">
                              <button
                                type="button"
                                onClick={() => onMoveClauseOrder(clause.id, 'up')}
                                disabled={idx === 0}
                                className="p-1 text-slate-400 hover:text-slate-800 disabled:opacity-30"
                                title="تحريك لأعلى"
                              >
                                <ArrowUp className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => onMoveClauseOrder(clause.id, 'down')}
                                disabled={idx === clauses.length - 1}
                                className="p-1 text-slate-400 hover:text-slate-800 disabled:opacity-30"
                                title="تحريك لأسفل"
                              >
                                <ArrowDown className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => onDeleteClause(clause.id)}
                                className="p-1 text-slate-400 hover:text-red-600"
                                title="حذف البند"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </div>

                          <div className="flex items-center gap-1.5 pt-0.5">
                            <button
                              type="button"
                              onMouseDown={(e) => {
                                e.preventDefault();
                                onSaveSelectionBookmark();
                              }}
                              onClick={() => onInsertClauseAtCaret(clause)}
                              className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-medium rounded transition-colors"
                              title="إدراج البند عند موضع المؤشر الحالي"
                            >
                              إدراج عند المؤشر
                            </button>
                            <button
                              type="button"
                              onClick={() => onEditClauseInEditor(clause)}
                              className="py-1 px-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 text-[11px] rounded inline-flex items-center gap-1"
                              title="فتح هذا البند داخل محرر A4 لتعديل صياغته"
                            >
                              <Edit3 className="w-3 h-3" />
                              <span>تعديل الصياغة</span>
                            </button>
                          </div>
                        </div>
                      );
                    })}
                </div>
              )}
            </div>
          </div>
        )}

        {/* =========================================================
            TAB 2: PARTIES, DESIGNATIONS & SUBDIVISION TABLE
           ========================================================= */}
        {activeTab === 'parties' && (
          <div className="space-y-4">
            {/* Sub-navigation between Variables & Subdivision Table */}
            <div className="grid grid-cols-2 gap-1 p-1 bg-slate-100 rounded text-xs font-medium">
              <button
                type="button"
                onClick={() => setPartiesSubTab('fields')}
                className={`py-1 rounded ${
                  partiesSubTab === 'fields'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600'
                }`}
              >
                الأطراف والمتغيرات
              </button>
              <button
                type="button"
                onClick={() => setPartiesSubTab('subdivision')}
                className={`py-1 rounded ${
                  partiesSubTab === 'subdivision'
                    ? 'bg-white text-slate-900 shadow-2xs'
                    : 'text-slate-600'
                }`}
              >
                جدول الوصف التقسيمي ({estates.length})
              </button>
            </div>

            {partiesSubTab === 'fields' ? (
              <div className="space-y-4">
                {/* Smart Variables Modal Launcher Card */}
                <div className="p-3 bg-pink-50/60 border border-pink-200 rounded-md space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-pink-950">
                      استمارة المتغيرات الموحدة
                    </span>
                    {unfilledCount > 0 ? (
                      <span className="px-1.5 py-0.5 text-[10px] font-bold bg-red-600 text-white rounded tabular-nums">
                        {unfilledCount} غير معبأة !
                      </span>
                    ) : (
                      <span className="px-1.5 py-0.5 text-[10px] font-bold bg-emerald-700 text-white rounded">
                        مكتمل ✓
                      </span>
                    )}
                  </div>
                  <p className="text-[11px] text-slate-600">
                    تُغذي هذه البيانات كلاً من العقد الأصلي وجميع الوثائق المشتقة (المستخرج، إجراء الشهر، شهادة البيع، الصيغة التنفيذية).
                  </p>
                  <button
                    type="button"
                    onClick={() => onOpenSmartVariablesModal()}
                    className="w-full py-1.5 px-3 bg-pink-800 text-white text-xs font-semibold rounded hover:bg-pink-900 inline-flex items-center justify-center gap-1.5 transition-colors"
                  >
                    <FormInput className="w-3.5 h-3.5" />
                    <span>فتح استمارة المتغيرات الكاملة</span>
                  </button>
                </div>

                {/* Quick Lot Selector from Subdivision Table */}
                <div className="border border-slate-200 rounded-md p-3 bg-slate-50/70">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-xs font-bold text-slate-900">
                      اختيار الحصة من جدول الوصف التقسيمي
                    </span>
                    <button
                      type="button"
                      onClick={() => setPartiesSubTab('subdivision')}
                      className="text-[11px] text-blue-800 hover:underline"
                    >
                      إدارة الجداول ({estates.length})
                    </button>
                  </div>

                  {estates.length === 0 ? (
                    <button
                      type="button"
                      onClick={() => {
                        setPartiesSubTab('subdivision');
                        startNewEstate();
                      }}
                      className="w-full py-1.5 px-2.5 bg-white border border-slate-300 text-slate-800 rounded text-xs font-medium hover:bg-slate-100 transition-colors"
                    >
                      + إضافة جدول وصف تقسيمي لعقار
                    </button>
                  ) : (
                    <div className="space-y-2">
                      <select
                        value={selectedEstate?.id || ''}
                        onChange={(e) => onSelectEstateAndLot(e.target.value, '')}
                        className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded"
                      >
                        {estates.map((est) => (
                          <option key={est.id} value={est.id}>
                            {est.estateName} ({est.lots.length} حصة)
                          </option>
                        ))}
                      </select>

                      {selectedEstate && (
                        <select
                          value={selectedLotNumber}
                          onChange={(e) =>
                            onSelectEstateAndLot(selectedEstate.id, e.target.value)
                          }
                          className="w-full px-2 py-1.5 text-xs bg-white border border-blue-300 rounded font-medium tabular-nums"
                        >
                          <option value="">-- اختر رقم الحصة لملء التعيين تلقائياً --</option>
                          {selectedEstate.lots.map((lot) => (
                            <option key={lot.id} value={lot.lotNumber}>
                              حصة رقم {lot.lotNumber} — {lot.nature} ({lot.floor}) — {lot.area} م²
                            </option>
                          ))}
                        </select>
                      )}

                      {selectedEstate && selectedLot && (
                        <div className="pt-1 space-y-1.5">
                          <div className="text-[11px] text-slate-700 bg-white p-2 rounded border border-slate-200 leading-relaxed">
                            {selectedLot.fullDescription}
                          </div>
                          <button
                            type="button"
                            onMouseDown={(e) => {
                              e.preventDefault();
                              onSaveSelectionBookmark();
                            }}
                            onClick={() => onInsertLotClauseAtCaret(selectedEstate, selectedLot)}
                            className="w-full py-1.5 px-2 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800 transition-colors"
                          >
                            إدراج فقرة تعيين الحصة رقم ({selectedLot.lotNumber}) عند المؤشر
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>

                {/* Detected Placeholders in Current Document */}
                {extraDocumentKeys.length > 0 && (
                  <div className="border border-amber-200 bg-amber-50/50 rounded-md p-3 space-y-2">
                    <div className="text-xs font-bold text-amber-950">
                      وسوم ذكية مستخرجة من العقد الحالي ({extraDocumentKeys.length})
                    </div>
                    {extraDocumentKeys.map((key) => (
                      <div key={key} className="space-y-1">
                        <div className="flex items-center justify-between">
                          <label className="text-[11px] font-medium text-slate-800">
                            {key}
                          </label>
                          <button
                            type="button"
                            onMouseDown={(e) => {
                              e.preventDefault();
                              onSaveSelectionBookmark();
                            }}
                            onClick={() => onInsertPlaceholderAtCaret(key)}
                            className="text-[10px] text-pink-800 font-semibold hover:underline"
                          >
                            + إدراج الوسم
                          </button>
                        </div>
                        <input
                          type="text"
                          value={fieldValues[key] || ''}
                          onChange={(e) => onUpdateFieldValue(key, e.target.value)}
                          placeholder={`أدخل قيمة ${key}...`}
                          className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded focus:border-blue-800 focus:outline-none"
                        />
                      </div>
                    ))}
                  </div>
                )}

                {/* Add Custom Variable */}
                <div className="border border-slate-200 rounded-md p-2.5 bg-white">
                  <div className="text-xs font-semibold text-slate-800 mb-1.5">
                    إنشاء وسم متغير جديد
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      value={newFieldKey}
                      onChange={(e) => setNewFieldKey(e.target.value)}
                      placeholder="مثال: اسم_الموثق أو رقم_الفهرس"
                      className="flex-1 px-2 py-1 text-xs border border-slate-300 rounded focus:border-blue-800 focus:outline-none"
                    />
                    <button
                      type="button"
                      onClick={() => {
                        const clean = newFieldKey.trim().replace(/\s+/g, '_');
                        if (clean) {
                          onAddCustomField(clean, newFieldKey.trim());
                          setNewFieldKey('');
                        }
                      }}
                      className="px-2.5 py-1 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800 shrink-0"
                    >
                      إضافة
                    </button>
                  </div>
                </div>

                {/* Standard Party & Property Fields */}
                <div className="space-y-3">
                  <div className="flex items-center justify-between border-b border-slate-200 pb-1.5">
                    <span className="text-xs font-bold text-slate-900">
                      حقول الأطراف والتعيينات
                    </span>
                    <button
                      type="button"
                      onClick={onBakeAllPlaceholdersIntoDocument}
                      className="text-[11px] text-pink-800 font-semibold hover:underline"
                    >
                      دمج الكل في النص
                    </button>
                  </div>
                  {partyFields.map((field) => (
                    <div key={field.key} className="space-y-1">
                      <div className="flex items-center justify-between gap-1">
                        <label className="text-[11px] font-medium text-slate-700 truncate">
                          {field.label}
                        </label>
                        <div className="flex items-center gap-1.5 shrink-0">
                          <button
                            type="button"
                            onMouseDown={(e) => {
                              e.preventDefault();
                              onSaveSelectionBookmark();
                            }}
                            onClick={() => onInsertPlaceholderAtCaret(field.key)}
                            className="text-[10px] text-pink-800 font-semibold hover:underline"
                            title={`إدراج الوسم {{${field.key}}} في موضع المؤشر`}
                          >
                            + وسم
                          </button>
                          {fieldValues[field.key] && (
                            <>
                              <span className="text-slate-300">·</span>
                              <button
                                type="button"
                                onMouseDown={(e) => {
                                  e.preventDefault();
                                  onSaveSelectionBookmark();
                                }}
                                onClick={() => onInsertValueAtCaret(fieldValues[field.key])}
                                className="text-[10px] text-emerald-700 hover:underline"
                              >
                                + قيمة
                              </button>
                            </>
                          )}
                        </div>
                      </div>
                      <input
                        type={field.inputType === 'date' ? 'date' : 'text'}
                        value={fieldValues[field.key] || ''}
                        onChange={(e) => onUpdateFieldValue(field.key, e.target.value)}
                        placeholder={`{{${field.key}}}`}
                        className="w-full px-2 py-1 text-xs bg-white border border-slate-200 rounded focus:border-blue-800 focus:outline-none"
                      />
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              /* Subdivision Estates Management Sub-tab */
              <div className="space-y-4">
                <div className="flex items-center justify-between">
                  <div className="text-xs font-bold text-slate-900">
                    جداول الوصف التقسيمي للعقارات
                  </div>
                  {!editingEstate && (
                    <button
                      type="button"
                      onClick={startNewEstate}
                      className="px-2.5 py-1 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800 inline-flex items-center gap-1"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>جدول عقار جديد</span>
                    </button>
                  )}
                </div>

                {editingEstate ? (
                  <div className="border border-blue-200 rounded-md p-3 bg-blue-50/30 space-y-3">
                    <div className="text-xs font-bold text-blue-950">
                      {editingEstate.estateName
                        ? `تحرير: ${editingEstate.estateName}`
                        : 'إنشاء جدول وصف تقسيمي جديد'}
                    </div>

                    <div className="space-y-2">
                      <input
                        type="text"
                        value={editingEstate.estateName}
                        onChange={(e) =>
                          setEditingEstate({ ...editingEstate, estateName: e.target.value })
                        }
                        placeholder="اسم العقار / الترقية العقارية *"
                        className="w-full px-2 py-1.5 text-xs bg-white border border-slate-300 rounded"
                      />
                      <div className="grid grid-cols-2 gap-2">
                        <input
                          type="text"
                          value={editingEstate.commune}
                          onChange={(e) =>
                            setEditingEstate({ ...editingEstate, commune: e.target.value })
                          }
                          placeholder="البلدية"
                          className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                        />
                        <input
                          type="text"
                          value={editingEstate.wilaya}
                          onChange={(e) =>
                            setEditingEstate({ ...editingEstate, wilaya: e.target.value })
                          }
                          placeholder="الولاية"
                          className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                        />
                      </div>
                      <input
                        type="text"
                        value={editingEstate.subdivisionDeedRef}
                        onChange={(e) =>
                          setEditingEstate({
                            ...editingEstate,
                            subdivisionDeedRef: e.target.value,
                          })
                        }
                        placeholder="مراجع عقد الوصف التقسيمي وشهره"
                        className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                      />
                    </div>

                    <div className="border-t border-slate-200 pt-2.5 space-y-2">
                      <div className="text-[11px] font-bold text-slate-800">
                        إضافة حصة إلى جدول الوصف التقسيمي
                      </div>
                      <div className="grid grid-cols-3 gap-1.5">
                        <input
                          type="text"
                          value={newLotNumber}
                          onChange={(e) => setNewLotNumber(e.target.value)}
                          placeholder="رقم الحصة *"
                          className="px-2 py-1 text-xs bg-white border border-slate-300 rounded tabular-nums"
                        />
                        <input
                          type="text"
                          value={newLotNature}
                          onChange={(e) => setNewLotNature(e.target.value)}
                          placeholder="الطبيعة"
                          className="px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                        />
                        <input
                          type="text"
                          value={newLotFloor}
                          onChange={(e) => setNewLotFloor(e.target.value)}
                          placeholder="الطابق"
                          className="px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                        />
                      </div>
                      <div className="grid grid-cols-3 gap-1.5">
                        <input
                          type="text"
                          value={newLotBuilding}
                          onChange={(e) => setNewLotBuilding(e.target.value)}
                          placeholder="العمارة/المدخل"
                          className="px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                        />
                        <input
                          type="text"
                          value={newLotArea}
                          onChange={(e) => setNewLotArea(e.target.value)}
                          placeholder="المساحة م²"
                          className="px-2 py-1 text-xs bg-white border border-slate-300 rounded tabular-nums"
                        />
                        <input
                          type="text"
                          value={newLotShares}
                          onChange={(e) => setNewLotShares(e.target.value)}
                          placeholder="الأجزاء المشتركة"
                          className="px-2 py-1 text-xs bg-white border border-slate-300 rounded tabular-nums"
                        />
                      </div>
                      <textarea
                        rows={2}
                        value={newLotDesc}
                        onChange={(e) => setNewLotDesc(e.target.value)}
                        placeholder="نص تعيين الحصة الكامل (يُولّد تلقائياً إن تُرك فارغاً)"
                        className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                      />
                      <button
                        type="button"
                        onClick={handleAddLotToEditingEstate}
                        className="w-full py-1.5 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800"
                      >
                        + إدراج الحصة في الجدول ({editingEstate.lots.length})
                      </button>
                    </div>

                    {editingEstate.lots.length > 0 && (
                      <div className="max-h-40 overflow-y-auto border border-slate-200 rounded bg-white divide-y divide-slate-100">
                        {editingEstate.lots.map((l) => (
                          <div
                            key={l.id}
                            className="flex items-center justify-between px-2 py-1.5 text-[11px]"
                          >
                            <div>
                              <span className="font-bold tabular-nums">حصة {l.lotNumber}</span>
                              <span> — {l.nature}</span>
                              {l.floor && <span> ({l.floor})</span>}
                              {l.area && <span className="tabular-nums"> · {l.area} م²</span>}
                            </div>
                            <button
                              type="button"
                              onClick={() =>
                                setEditingEstate({
                                  ...editingEstate,
                                  lots: editingEstate.lots.filter((x) => x.id !== l.id),
                                })
                              }
                              className="text-red-600 hover:underline"
                            >
                              حذف
                            </button>
                          </div>
                        ))}
                      </div>
                    )}

                    <div className="flex items-center gap-2 pt-1">
                      <button
                        type="button"
                        onClick={async () => {
                          if (!editingEstate.estateName.trim()) return;
                          await onSaveEstate(editingEstate);
                          setEditingEstate(null);
                        }}
                        className="flex-1 py-1.5 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800"
                      >
                        حفظ جدول العقار
                      </button>
                      <button
                        type="button"
                        onClick={() => setEditingEstate(null)}
                        className="px-3 py-1.5 bg-white border border-slate-300 text-slate-700 text-xs rounded hover:bg-slate-100"
                      >
                        إلغاء
                      </button>
                    </div>
                  </div>
                ) : (
                  <div className="space-y-2.5">
                    {estates.map((est) => (
                      <div
                        key={est.id}
                        className="border border-slate-200 rounded-md p-3 bg-white space-y-2"
                      >
                        <div className="flex items-start justify-between">
                          <div>
                            <div className="text-xs font-bold text-slate-900">
                              {est.estateName}
                            </div>
                            <div className="text-[11px] text-slate-500">
                              {est.commune} {est.wilaya && `· ${est.wilaya}`} ·{' '}
                              <span className="tabular-nums">{est.lots.length} حصة</span>
                            </div>
                          </div>
                          <div className="flex items-center gap-1">
                            <button
                              type="button"
                              onClick={() => setEditingEstate(est)}
                              className="px-2 py-0.5 text-[11px] text-blue-800 hover:bg-blue-50 rounded"
                            >
                              تعديل
                            </button>
                            <button
                              type="button"
                              onClick={() => onDeleteEstate(est.id)}
                              className="p-1 text-slate-400 hover:text-red-600"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                        <button
                          type="button"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            onSaveSelectionBookmark();
                          }}
                          onClick={() => onInsertFullSubdivisionTableAtCaret(est)}
                          className="w-full py-1.5 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-medium rounded transition-colors"
                        >
                          إدراج جدول الوصف التقسيمي الكامل عند المؤشر
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* =========================================================
            TAB 3: OFFICE TEMPLATES & EDITABLE DERIVED TEMPLATES
           ========================================================= */}
        {activeTab === 'templates' && (
          <div className="space-y-4">
            {/* Sub-tab selector: Derived Templates vs Office Contract Templates */}
            <div className="grid grid-cols-2 gap-1 p-1 bg-slate-100 rounded text-xs font-medium">
              <button
                type="button"
                onClick={() => setTemplatesSubTab('derived')}
                className={`py-1.5 rounded ${
                  templatesSubTab === 'derived'
                    ? 'bg-white text-slate-900 shadow-2xs font-bold'
                    : 'text-slate-600'
                }`}
              >
                قوالب الوثائق المشتقة ({derivedTemplates.length})
              </button>
              <button
                type="button"
                onClick={() => setTemplatesSubTab('office')}
                className={`py-1.5 rounded ${
                  templatesSubTab === 'office'
                    ? 'bg-white text-slate-900 shadow-2xs font-bold'
                    : 'text-slate-600'
                }`}
              >
                قوالب عقود المكتب ({templates.length})
              </button>
            </div>

            {templatesSubTab === 'derived' ? (
              <div className="space-y-4">
                <div className="p-3 bg-blue-50/50 border border-blue-200 rounded-md space-y-1.5">
                  <div className="text-xs font-bold text-blue-950">
                    مولد وقوالب الوثائق المشتقة القابلة للتعديل الكامل
                  </div>
                  <p className="text-[11px] text-slate-600 leading-relaxed">
                    كل قالب وثيقة مشتقة (المستخرج، إجراء الشهر، شهادة البيع، الصيغة التنفيذية...) قابل للفتح والتعديل داخل محرر الـ A4 أو استبداله بملف <code className="font-mono">.docx</code> من جهازك، ويعتمد عند التوليد حصرياً على قيم «استمارة المتغيرات» والأطراف والتعيينات.
                  </p>
                </div>

                {/* List of Editable Derived Document Templates */}
                <div className="space-y-2.5">
                  {derivedTemplates.map((dt) => {
                    const isCurrentlyEditing = activeEditingDerivedId === dt.id;
                    return (
                      <div
                        key={dt.id}
                        className={`border rounded-md p-3 space-y-2 transition-colors ${
                          isCurrentlyEditing
                            ? 'border-amber-500 bg-amber-50/40'
                            : 'border-slate-200 bg-white hover:border-slate-300'
                        }`}
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <div className="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                              <span>{dt.name}</span>
                              {isCurrentlyEditing && (
                                <span className="text-[10px] bg-amber-200 text-amber-950 px-1.5 py-0.5 rounded">
                                  قيد التعديل بالمحرر
                                </span>
                              )}
                            </div>
                            <div className="text-[11px] text-slate-500 mt-0.5">
                              {dt.description}
                            </div>
                          </div>
                          {!dt.id.startsWith('derived_') && (
                            <button
                              type="button"
                              onClick={() => onDeleteDerivedTemplate(dt.id)}
                              className="p-1 text-slate-400 hover:text-red-600"
                              title="حذف القالب المشتق"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>

                        <div className="flex flex-wrap items-center gap-1.5 pt-1">
                          <button
                            type="button"
                            onClick={() => onGenerateDerivedDoc(dt)}
                            className="flex-1 py-1.5 px-2 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800 inline-flex items-center justify-center gap-1"
                            title="توليد وتصدير هذه الوثيقة المشتقة فوراً من بيانات استمارة المتغيرات الحالية"
                          >
                            <Download className="w-3.5 h-3.5" />
                            <span>توليد وتصدير (.docx)</span>
                          </button>

                          <button
                            type="button"
                            onClick={() => onEditDerivedTemplateInEditor(dt)}
                            className="py-1.5 px-2.5 bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-200 text-[11px] font-medium rounded inline-flex items-center gap-1"
                            title="فتح قالب هذه الوثيقة المشتقة داخل محرر A4 لتعديل صياغته أو جدوله أو وسومه"
                          >
                            <Edit3 className="w-3 h-3" />
                            <span>تعديل القالب بالمحرر</span>
                          </button>

                          <label
                            className="py-1.5 px-2 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 text-[11px] rounded cursor-pointer inline-flex items-center gap-1"
                            title="استيراد ملف Word (.docx) من جهازك ليصبح هو القالب المعتمد لهذه الوثيقة المشتقة"
                          >
                            <Upload className="w-3 h-3" />
                            <span>.docx</span>
                            <input
                              type="file"
                              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                              className="hidden"
                              onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) {
                                  onImportDocxForDerivedTemplate(dt, file);
                                  e.target.value = '';
                                }
                              }}
                            />
                          </label>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Add a new custom Derived Document Template from current Editor */}
                <div className="border border-slate-200 rounded-md p-3 bg-slate-50 space-y-2">
                  <div className="text-xs font-bold text-slate-900">
                    إضافة قالب وثيقة مشتقة جديد من المحرر
                  </div>
                  <input
                    type="text"
                    value={newDerivedName}
                    onChange={(e) => setNewDerivedName(e.target.value)}
                    placeholder="اسم الوثيقة المشتقة (مثال: إشعار بالتسجيل / جدول إرسال)"
                    className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded"
                  />
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      value={newDerivedCode}
                      onChange={(e) => setNewDerivedCode(e.target.value)}
                      placeholder="الرمز (مثال: bordereau)"
                      className="w-28 px-2 py-1 text-xs bg-white border border-slate-300 rounded font-mono"
                    />
                    <input
                      type="text"
                      value={newDerivedDesc}
                      onChange={(e) => setNewDerivedDesc(e.target.value)}
                      placeholder="وصف مختصر..."
                      className="flex-1 px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={async () => {
                      if (!newDerivedName.trim()) return;
                      await onSaveCurrentAsDerivedTemplate(
                        newDerivedName.trim(),
                        newDerivedCode.trim() || `custom_${Date.now()}`,
                        newDerivedDesc.trim() || 'قالب وثيقة مشتقة مخصص للمكتب'
                      );
                      setNewDerivedName('');
                      setNewDerivedCode('');
                      setNewDerivedDesc('');
                    }}
                    className="w-full py-1.5 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800"
                  >
                    + حفظ محتوى المحرر الحالي كقالب وثيقة مشتقة
                  </button>
                </div>
              </div>
            ) : (
              /* Office Contract Templates Sub-tab */
              <div className="space-y-4">
                <div className="border border-blue-200 bg-blue-50/40 rounded-md p-3 space-y-2.5">
                  <div className="text-xs font-bold text-slate-900">
                    استيراد نماذج العقود الخاصة بمكتبكم (.docx)
                  </div>
                  <label className="w-full flex items-center justify-center gap-2 py-2 px-3 bg-blue-900 text-white rounded text-xs font-medium hover:bg-blue-800 transition-colors cursor-pointer">
                    <Upload className="w-4 h-4" />
                    <span>استيراد قالب Word (.docx) من الجهاز</span>
                    <input
                      type="file"
                      accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                      multiple
                      className="hidden"
                      onChange={(e) => {
                        if (e.target.files && e.target.files.length > 0) {
                          onImportDocxAsTemplate(e.target.files);
                          e.target.value = '';
                        }
                      }}
                    />
                  </label>
                </div>

                <div className="border border-slate-200 rounded-md p-3 space-y-2 bg-white">
                  <div className="text-xs font-bold text-slate-900">
                    حفظ العقد الحالي كقالب في مكتبة المكتب
                  </div>
                  <input
                    type="text"
                    value={newTplName}
                    onChange={(e) => setNewTplName(e.target.value)}
                    placeholder="اسم القالب (مثال: عقد بيع شقة - مكتبنا)"
                    className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded"
                  />
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      value={newTplCategory}
                      onChange={(e) => setNewTplCategory(e.target.value)}
                      placeholder="التصنيف"
                      className="flex-1 px-2.5 py-1.5 text-xs border border-slate-300 rounded"
                    />
                    <button
                      type="button"
                      onClick={async () => {
                        if (!newTplName.trim()) return;
                        await onSaveCurrentAsTemplate(
                          newTplName.trim(),
                          newTplCategory.trim() || 'عام'
                        );
                        setNewTplName('');
                      }}
                      className="px-3 py-1.5 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800"
                    >
                      حفظ كقالب
                    </button>
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="text-xs font-bold text-slate-900 border-b border-slate-200 pb-1.5">
                    مكتبة قوالب العقود المحفوظة ({templates.length})
                  </div>
                  {templates.length === 0 ? (
                    <div className="border border-dashed border-slate-300 rounded-md p-5 text-center space-y-2 bg-slate-50/50">
                      <FileText className="w-7 h-7 text-slate-400 mx-auto" />
                      <div className="text-xs font-semibold text-slate-700">
                        مكتبة قوالب العقود فارغة حالياً
                      </div>
                      <p className="text-[11px] text-slate-500">
                        استوردوا ملفات Word (.docx) الخاصة بمكتبكم أو احفظوا أي عقد من المحرر.
                      </p>
                    </div>
                  ) : (
                    templates.map((tpl) => (
                      <div
                        key={tpl.id}
                        className="border border-slate-200 rounded-md p-2.5 bg-white space-y-2"
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <div className="text-xs font-bold text-slate-900">{tpl.name}</div>
                            <div className="text-[11px] text-slate-500">
                              {tpl.category} · {tpl.extractedPlaceholders.length} وسم ذكي
                            </div>
                          </div>
                          <button
                            type="button"
                            onClick={() => onDeleteTemplate(tpl.id)}
                            className="p-1 text-slate-400 hover:text-red-600"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => onLoadTemplateFull(tpl)}
                            className="flex-1 py-1 px-2 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800"
                          >
                            فتح كعقد جديد
                          </button>
                          <button
                            type="button"
                            onMouseDown={(e) => {
                              e.preventDefault();
                              onSaveSelectionBookmark();
                            }}
                            onClick={() => onInsertTemplateAtCaret(tpl)}
                            className="flex-1 py-1 px-2 bg-slate-100 text-slate-800 border border-slate-200 text-[11px] font-medium rounded hover:bg-slate-200"
                          >
                            إدراج عند المؤشر
                          </button>
                          <button
                            type="button"
                            onClick={() => onExportTemplateDocx(tpl)}
                            className="py-1 px-2 bg-white border border-slate-200 text-slate-700 text-[11px] rounded hover:bg-slate-50"
                          >
                            <Download className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            )}
          </div>
        )}

        {/* =========================================================
            TAB 4: SAVED DOCUMENTS, VERSION DIFF & DOWNLOAD ARCHIVE
           ========================================================= */}
        {activeTab === 'documents' && (
          <div className="space-y-4">
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={onOpenMultiSourceNewModal}
                className="flex-1 py-1.5 px-2.5 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800 inline-flex items-center justify-center gap-1.5"
              >
                <FilePlus2 className="w-3.5 h-3.5" />
                <span>عقد جديد (متعدد المصادر)</span>
              </button>
              <button
                type="button"
                onClick={onSaveCurrentDocument}
                className="py-1.5 px-3 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800 inline-flex items-center gap-1.5"
              >
                <Save className="w-3.5 h-3.5" />
                <span>حفظ</span>
              </button>
            </div>

            {/* Sub-tabs: Saved Contracts | Revisions Diff | Download Archive */}
            <div className="grid grid-cols-3 gap-1 p-1 bg-slate-100 rounded text-[11px] font-medium">
              <button
                type="button"
                onClick={() => setDocsSubTab('saved')}
                className={`py-1 rounded ${
                  docsSubTab === 'saved' ? 'bg-white text-slate-900 shadow-2xs font-bold' : 'text-slate-600'
                }`}
              >
                المسودات ({documents.length})
              </button>
              <button
                type="button"
                onClick={() => setDocsSubTab('revisions')}
                className={`py-1 rounded ${
                  docsSubTab === 'revisions'
                    ? 'bg-white text-slate-900 shadow-2xs font-bold'
                    : 'text-slate-600'
                }`}
              >
                سجل النسخ ({revisions.length})
              </button>
              <button
                type="button"
                onClick={() => setDocsSubTab('downloads')}
                className={`py-1 rounded ${
                  docsSubTab === 'downloads'
                    ? 'bg-white text-slate-900 shadow-2xs font-bold'
                    : 'text-slate-600'
                }`}
              >
                أرشيف التحميلات ({downloads.length})
              </button>
            </div>

            {docsSubTab === 'saved' && (
              <div className="space-y-2.5">
                <input
                  type="text"
                  value={docSearch}
                  onChange={(e) => setDocSearch(e.target.value)}
                  placeholder="بحث في العقود والمسودات المحفوظة..."
                  className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded"
                />

                {documents
                  .filter((d) => !docSearch.trim() || d.title.includes(docSearch.trim()))
                  .map((doc) => {
                    const isCurrent = doc.id === activeDocumentId;
                    return (
                      <div
                        key={doc.id}
                        className={`border rounded-md p-2.5 space-y-2 ${
                          isCurrent
                            ? 'border-blue-400 bg-blue-50/40'
                            : 'border-slate-200 bg-white'
                        }`}
                      >
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <div className="text-xs font-bold text-slate-900">{doc.title}</div>
                            <div className="text-[11px] text-slate-500 tabular-nums">
                              {new Date(doc.updatedAt).toLocaleString('ar-DZ')}
                            </div>
                          </div>
                          <button
                            type="button"
                            onClick={() => onDeleteDocument(doc.id)}
                            className="p-1 text-slate-400 hover:text-red-600"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <button
                            type="button"
                            onClick={() => onOpenDocument(doc)}
                            className="flex-1 py-1 px-2 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800"
                          >
                            فتح في المحرر
                          </button>
                          <button
                            type="button"
                            onClick={() => onExportDocumentDocx(doc)}
                            className="py-1 px-2.5 bg-slate-100 text-slate-800 text-[11px] font-medium rounded hover:bg-slate-200 inline-flex items-center gap-1"
                          >
                            <Download className="w-3 h-3" />
                            <span>.docx</span>
                          </button>
                        </div>
                      </div>
                    );
                  })}
              </div>
            )}

            {docsSubTab === 'revisions' && (
              <div className="space-y-2.5">
                <button
                  type="button"
                  onClick={onOpenVersionDiffModal}
                  className="w-full py-2 px-3 bg-blue-900 text-white text-xs font-semibold rounded hover:bg-blue-800 inline-flex items-center justify-center gap-1.5"
                >
                  <GitCompare className="w-4 h-4" />
                  <span>فتح مقارن النسخ التفصيلي (Diff بالأخضر والأحمر)</span>
                </button>

                {revisions.length === 0 ? (
                  <div className="text-center py-6 text-xs text-slate-500">
                    لا توجد نسخ مسجلة بعد. يتم تسجيل النسخ تلقائياً مع الحفظ.
                  </div>
                ) : (
                  revisions.map((rev) => (
                    <div
                      key={rev.id}
                      className="border border-slate-200 rounded-md p-2.5 bg-white space-y-1.5"
                    >
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold text-slate-900">
                          {rev.documentTitle}
                        </span>
                        <span className="text-[10px] text-slate-500 tabular-nums">
                          {new Date(rev.createdAt).toLocaleString('ar-DZ')}
                        </span>
                      </div>
                      <div className="text-[11px] text-slate-600">
                        {rev.summary} · ({rev.author})
                      </div>
                      <div className="flex items-center gap-1.5 pt-1">
                        <button
                          type="button"
                          onClick={onOpenVersionDiffModal}
                          className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-medium rounded"
                        >
                          مقارنة مع الحالية
                        </button>
                        <button
                          type="button"
                          onClick={() => onRestoreRevision(rev)}
                          className="py-1 px-2 bg-blue-50 text-blue-900 hover:bg-blue-100 text-[11px] font-medium rounded inline-flex items-center gap-1"
                        >
                          <RotateCcw className="w-3 h-3" />
                          <span>استعادة</span>
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            )}

            {docsSubTab === 'downloads' && (
              <div className="space-y-2.5">
                <div className="text-[11px] text-slate-600 bg-slate-50 p-2 rounded border border-slate-200">
                  يُحفظ هنا كل ملف تم تصديره (سواء العقد الأصلي أو المستخرج أو إجراء الشهر أو شهادة البيع أو الصيغة التنفيذية) لإعادة فتحه أو تحميله فوراً.
                </div>

                {downloads.length === 0 ? (
                  <div className="text-center py-6 text-xs text-slate-500">
                    أرشيف التحميلات فارغ حالياً.
                  </div>
                ) : (
                  downloads.map((item) => (
                    <div
                      key={item.id}
                      className="border border-slate-200 rounded-md p-2.5 bg-white space-y-2"
                    >
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <div className="text-xs font-bold text-slate-900">
                            {item.documentTitle}
                          </div>
                          <div className="text-[11px] text-blue-900 font-semibold">
                            نوع الوثيقة: {item.docTypeLabel}
                          </div>
                          <div className="text-[10px] text-slate-500 tabular-nums">
                            {new Date(item.createdAt).toLocaleString('ar-DZ')}
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={() => onDeleteDownloadArchiveItem(item.id)}
                          className="p-1 text-slate-400 hover:text-red-600"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <button
                          type="button"
                          onClick={() => onOpenDownloadArchiveItemInEditor(item)}
                          className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-medium rounded"
                        >
                          فتح في المحرر
                        </button>
                        <button
                          type="button"
                          onClick={() => onRedownloadArchiveItemDocx(item)}
                          className="flex-1 py-1 px-2 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800 inline-flex items-center justify-center gap-1"
                        >
                          <Download className="w-3 h-3" />
                          <span>إعادة تحميل .docx</span>
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
            )}

            {/* Full Offline Backup Export / Import (JSON) */}
            <div className="border-t border-slate-200 pt-3 space-y-2">
              <div className="text-xs font-bold text-slate-900">
                النسخ الاحتياطي الشامل للمكتب (JSON)
              </div>
              <div className="grid grid-cols-2 gap-2">
                <button
                  type="button"
                  onClick={onExportBackupJson}
                  className="py-1.5 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded inline-flex items-center justify-center gap-1"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>تصدير قاعدة المكتب</span>
                </button>
                <label className="py-1.5 px-2 bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-medium rounded inline-flex items-center justify-center gap-1 cursor-pointer">
                  <Upload className="w-3.5 h-3.5" />
                  <span>استيراد نسخة</span>
                  <input
                    type="file"
                    accept=".json,application/json"
                    className="hidden"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) {
                        onImportBackupJson(f);
                        e.target.value = '';
                      }
                    }}
                  />
                </label>
              </div>
            </div>
          </div>
        )}
      </div>
    </aside>
  );
}
