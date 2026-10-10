'use client';

import React, { useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  BookmarkPlus,
  Building2,
  Check,
  CheckSquare,
  ChevronDown,
  ChevronUp,
  Download,
  Edit3,
  Eye,
  FileOutput,
  FilePlus2,
  FileText,
  FileUp,
  FolderArchive,
  FormInput,
  GitCompare,
  GripVertical,
  Hash,
  History,
  ListChecks,
  Lock,
  PanelRightClose,
  Plus,
  RotateCcw,
  Save,
  Sparkles,
  Square,
  Trash2,
  Unlock,
  Upload,
  UserCheck,
  Users,
  X,
} from 'lucide-react';
import {
  ClauseVariableGroup,
  ContractFolder,
  ContractOutlineClause,
  ContractPartyCard,
  CustomTemplate,
  DerivedDocTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  NotaryClerk,
  NotaryClause,
  NotaryPartyRole,
  PartyField,
  SavedContractDerivedDoc,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
  SubdivisionLot,
  VariableInputType,
} from '@/lib/types';
import FolderTreeExplorer from './FolderTreeExplorer';

export type SidebarTab = 'clauses' | 'parties' | 'templates' | 'documents';

interface SidebarWorkspaceProps {
  activeTab: SidebarTab;
  onSelectTab: (tab: SidebarTab) => void;
  // 1. Live Contract Clauses (Single Source of Truth) + Optional Ready Clauses Library
  liveContractClauses: ContractOutlineClause[];
  onScrollToLiveClause: (domIndexOrId: number | string) => void;
  onMoveLiveClauseInDoc: (domIndex: number, direction: 'up' | 'down') => void;
  onToggleLiveClauseEnabled?: (clauseId: string) => void;
  onToggleLiveClauseLocked?: (clauseId: string) => void;
  onReorderLiveClauses?: (sourceClauseId: string, targetClauseId: string) => void;
  onDeleteLiveClauseFromDoc: (domIndexOrId: number | string) => void;
  onRenameLiveClauseInDoc: (domIndexOrId: number | string, newTitle: string) => void;
  onInsertNewClauseHeadingInDoc: (title?: string) => void;
  onImportStandardClausesPack?: () => Promise<void>;
  clauses: NotaryClause[];
  activeClauseIdsInDoc: string[];
  onToggleClauseInDoc: (clause: NotaryClause) => void;
  onInsertClauseAtCaret: (clause: NotaryClause) => void;
  onMoveClauseOrder: (clauseId: string, direction: 'up' | 'down') => Promise<void>;
  onSaveNewClause: (title: string, category: string, contentHtml?: string) => Promise<void>;
  onSaveSelectionAsClause: (title: string, category: string) => Promise<void>;
  onEditClauseInEditor: (clause: NotaryClause) => void;
  onDeleteClause: (id: string) => Promise<void>;
  // 2. Parties & Designations + Structured Party Cards + Subdivision Table
  partyCards?: ContractPartyCard[];
  onUpdatePartyCards?: (nextCards: ContractPartyCard[]) => void;
  onSavePartyCardToDirectory?: (card: ContractPartyCard) => void;
  partyFields: PartyField[];
  extractedPlaceholders: string[];
  clauseGroups: ClauseVariableGroup[];
  unfilledCount: number;
  fieldValues: Record<string, string>;
  fieldInputTypes?: Record<string, VariableInputType>;
  onOpenSmartVariablesModal: (focusKey?: string) => void;
  onUpdateFieldValue: (key: string, value: string) => void;
  onChangeFieldInputType: (key: string, inputType: 'text' | 'number' | 'date') => void;
  onAddCustomField: (key: string, label: string) => void;
  onInsertPlaceholderAtCaret: (key: string) => void;
  onInsertValueAtCaret: (value: string) => void;
  onBakeAllPlaceholdersIntoDocument: (explicitValues?: Record<string, string>) => void;
  estates: SubdivisionEstate[];
  selectedEstateId: string;
  selectedLotNumber: string;
  onSelectEstateAndLot: (estateId: string, lotNumber: string) => void;
  onSaveEstate: (estate: SubdivisionEstate) => Promise<void>;
  onDeleteEstate: (id: string) => Promise<void>;
  onInsertLotClauseAtCaret: (estate: SubdivisionEstate, lot: SubdivisionLot) => void;
  onInsertFullSubdivisionTableAtCaret: (estate: SubdivisionEstate) => void;
  // Saved Parties & Saved Properties Directory (v2.5)
  savedParties: SavedPartyRecord[];
  savedProperties: SavedPropertyRecord[];
  onRecallPartyToRole: (party: SavedPartyRecord, role: 'party1' | 'party2') => void;
  onSaveCurrentPartyToDirectory: (role: 'party1' | 'party2') => void;
  onDeleteSavedParty: (id: string) => void;
  onRecallPropertyRecord: (prop: SavedPropertyRecord) => void;
  onSaveCurrentPropertyToDirectory: () => void;
  onDeleteSavedProperty: (id: string) => void;
  // 3. Office Templates & Editable Derived Document Templates
  templates: CustomTemplate[];
  onImportDocxAsTemplate: (files: FileList) => Promise<void>;
  onSaveCurrentAsTemplate: (name: string, category: string) => Promise<void>;
  onOpenSaveAsTemplateModal?: () => void;
  onExportTemplatesJson?: () => void;
  onImportTemplatesJson?: (file: File) => void;
  onLoadTemplateFull: (tpl: CustomTemplate, clearPreviousClauses?: boolean) => void;
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
  currentEditorTitle?: string;
  onOpenDocument: (doc: SavedDocument) => void;
  onOpenMultiSourceNewModal: () => void;
  onSaveCurrentDocument: (targetFolderId?: string | null) => Promise<void>;
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
  onToggleCollapse?: () => void;
  folders: ContractFolder[];
  clerks: NotaryClerk[];
  activeClerk: NotaryClerk;
  activeDocument: SavedDocument | null;
  activeDerivedDocId?: string | null;
  onCreateFolder: (parentId: string | null, name: string) => void;
  onRenameFolder: (folderId: string, newName: string) => void;
  onDeleteFolder: (folderId: string) => void;
  onMoveDocument: (docId: string, targetFolderId: string | null) => void;
  onMoveFolder: (folderId: string, targetParentId: string | null) => void;
  onCreateContractInFolder: (folderId: string | null) => void;
  onUpdateDocumentMeta: (updated: SavedDocument) => void;
  onOpenDerivedModal: (templateId?: string) => void;
  onSelectDerivedDoc: (derivedDoc: SavedContractDerivedDoc) => void;
}

export default function SidebarWorkspace({
  activeTab,
  onSelectTab,
  liveContractClauses,
  onScrollToLiveClause,
  onMoveLiveClauseInDoc,
  onToggleLiveClauseEnabled,
  onToggleLiveClauseLocked,
  onReorderLiveClauses,
  onDeleteLiveClauseFromDoc,
  onRenameLiveClauseInDoc,
  onInsertNewClauseHeadingInDoc,
  onImportStandardClausesPack,
  clauses,
  activeClauseIdsInDoc,
  onToggleClauseInDoc,
  onInsertClauseAtCaret,
  onMoveClauseOrder,
  onSaveNewClause,
  onSaveSelectionAsClause,
  onEditClauseInEditor,
  onDeleteClause,
  partyCards = [],
  onUpdatePartyCards,
  onSavePartyCardToDirectory,
  partyFields,
  extractedPlaceholders,
  clauseGroups,
  unfilledCount,
  fieldValues,
  fieldInputTypes = {},
  onOpenSmartVariablesModal,
  onUpdateFieldValue,
  onChangeFieldInputType,
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
  savedParties,
  savedProperties,
  onRecallPartyToRole,
  onSaveCurrentPartyToDirectory,
  onDeleteSavedParty,
  onRecallPropertyRecord,
  onSaveCurrentPropertyToDirectory,
  onDeleteSavedProperty,
  templates,
  onImportDocxAsTemplate,
  onSaveCurrentAsTemplate,
  onOpenSaveAsTemplateModal,
  onExportTemplatesJson,
  onImportTemplatesJson,
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
  currentEditorTitle,
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
  onToggleCollapse,
  folders,
  clerks,
  activeClerk,
  activeDocument,
  activeDerivedDocId,
  onCreateFolder,
  onRenameFolder,
  onDeleteFolder,
  onMoveDocument,
  onMoveFolder,
  onCreateContractInFolder,
  onUpdateDocumentMeta,
  onOpenDerivedModal,
  onSelectDerivedDoc,
}: SidebarWorkspaceProps) {
  // Clause creation & live contract outline state
  const [newClauseTitle, setNewClauseTitle] = useState('');
  const [newClauseCategory, setNewClauseCategory] = useState('بنود عامة');
  const [newClauseText, setNewClauseText] = useState('');
  const [clauseSearch, setClauseSearch] = useState('');
  const [quickDocClauseTitle, setQuickDocClauseTitle] = useState('');
  const [renamingClauseId, setRenamingClauseId] = useState<string | null>(null);
  const [renamingClauseVal, setRenamingClauseVal] = useState('');
  const [draggedClauseId, setDraggedClauseId] = useState<string | null>(null);
  const [dragOverClauseId, setDragOverClauseId] = useState<string | null>(null);
  const [showSuggestionsLibrary, setShowSuggestionsLibrary] = useState<boolean>(true);
  const [showAddLibraryForm, setShowAddLibraryForm] = useState<boolean>(false);

  // Parties sub-view: 'fields' vs 'subdivision'
  const [partiesSubTab, setPartiesSubTab] = useState<'fields' | 'subdivision'>('fields');
  const [newFieldKey, setNewFieldKey] = useState('');
  const [newPartyRole, setNewPartyRole] = useState<NotaryPartyRole>('بائع');
  const [expandedPartyCardId, setExpandedPartyCardId] = useState<string | null>(null);

  // Subdivision Estate form state
  const [editingEstate, setEditingEstate] = useState<SubdivisionEstate | null>(null);
  const [newLotNumber, setNewLotNumber] = useState('');
  const [newLotBuilding, setNewLotBuilding] = useState('');
  const [newLotFloor, setNewLotFloor] = useState('');
  const [newLotNature, setNewLotNature] = useState('شقة سكنية');
  const [newLotArea, setNewLotArea] = useState('');
  const [newLotShares, setNewLotShares] = useState('');
  const [newLotDesc, setNewLotDesc] = useState('');

  // Office Contract Templates form state
  const [newTplName, setNewTplName] = useState('');
  const [newTplCategory, setNewTplCategory] = useState('عقود المكتب');
  const [clearClausesOnReplace, setClearClausesOnReplace] = useState<boolean>(true);
  const [selectedPartyIdSidebar, setSelectedPartyIdSidebar] = useState<string>('');
  const [selectedPropIdSidebar, setSelectedPropIdSidebar] = useState<string>('');

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
    <aside className="w-80 md:w-84 lg:w-88 xl:w-96 max-w-[85vw] bg-white border-l border-slate-200 flex flex-col h-full min-h-0 overflow-hidden shrink-0 select-none no-print">
      {/* Sidebar Header & Collapse Toggle */}
      <div className="flex items-center justify-between px-3 py-2 bg-white border-b border-slate-200 shrink-0">
        <div className="flex items-center gap-2">
          <span className="w-2 h-2 rounded-full bg-blue-900" />
          <h3 className="text-xs font-bold text-slate-900">
            {activeTab === 'clauses' && 'بنود وصياغات العقود'}
            {activeTab === 'parties' && 'استمارة الأطراف والتعيينات'}
            {activeTab === 'templates' && 'قوالب عقود المكتب'}
            {activeTab === 'documents' && 'المستندات والأرشيف المحلي'}
          </h3>
        </div>
        {onToggleCollapse && (
          <button
            type="button"
            onClick={onToggleCollapse}
            className="px-2 py-1 text-slate-600 hover:text-blue-900 hover:bg-slate-100 rounded border border-transparent hover:border-slate-200 transition-colors inline-flex items-center gap-1 text-[11px] font-medium"
            title="طي اللوحة الجانبية إلى شريط أيقونات نحيف لتوسيع مساحة ورقة العقد"
          >
            <span>طي</span>
            <PanelRightClose className="w-3.5 h-3.5" />
          </button>
        )}
      </div>

      {/* Sidebar Segmented Navigation — Single Source of Section Navigation */}
      <div className="grid grid-cols-4 gap-1 p-2 bg-slate-50 border-b border-slate-200 shrink-0">
        {/* 1. البنود */}
        <button
          type="button"
          onClick={() =>
            activeTab === 'clauses' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('clauses')
          }
          title="البنود الجاهزة (انقر مجدداً لطي الشريط)"
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'clauses'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <ListChecks className="w-3.5 h-3.5 mb-0.5" />
          <span>البنود</span>
        </button>

        {/* 2. الأطراف والتعيينات */}
        <button
          type="button"
          onClick={() =>
            activeTab === 'parties' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('parties')
          }
          title="الأطراف والتعيينات (انقر مجدداً لطي الشريط)"
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap relative ${
            activeTab === 'parties'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <UserCheck className="w-3.5 h-3.5 mb-0.5" />
          <span>الأطراف</span>
          {unfilledCount > 0 && (
            <span className="absolute top-1 left-1 w-2 h-2 rounded-full bg-red-600" />
          )}
        </button>

        {/* 3. قوالب العقود (مفصولة عن الوثائق المشتقة) */}
        <button
          type="button"
          onClick={() =>
            activeTab === 'templates' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('templates')
          }
          title="قوالب عقود المكتب (انقر مجدداً لطي الشريط)"
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'templates'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <FileText className="w-3.5 h-3.5 mb-0.5" />
          <span>القوالب</span>
        </button>

        {/* 4. الأرشيف والمستندات */}
        <button
          type="button"
          onClick={() =>
            activeTab === 'documents' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('documents')
          }
          title="المستندات والأرشيف (انقر مجدداً لطي الشريط)"
          className={`flex flex-col items-center justify-center py-1.5 px-1 rounded text-[11px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'documents'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <FolderArchive className="w-3.5 h-3.5 mb-0.5" />
          <span>الأرشيف</span>
        </button>
      </div>

      {/* Tab Content Area */}
      <div className="flex-1 min-h-0 overflow-y-auto p-3 space-y-4">
        {/* =========================================================
            TAB 1: LIVE CONTRACT CLAUSES (#) + OPTIONAL CLAUSE LIBRARY
           ========================================================= */}
        {activeTab === 'clauses' && (
          <div className="space-y-4">
            {/* 1. LIVE CONTRACT CLAUSES OUTLINE (بنود هذا العقد الحالية من عناوين # داخل الورقة) */}
            <div className="border border-blue-200 bg-blue-50/30 rounded-md p-3 space-y-2.5">
              <div className="flex items-center justify-between border-b border-blue-200/80 pb-2">
                <div className="flex items-center gap-1.5">
                  <Hash className="w-3.5 h-3.5 text-blue-900" />
                  <span className="text-xs font-bold text-slate-900">
                    بنود العقد المفتوح حالياً ({liveContractClauses.length})
                  </span>
                </div>
                <span className="text-[10px] text-blue-900 font-medium">
                  تُقرأ حياً من `#` في الورقة
                </span>
              </div>

              {/* Quick add `# Clause` directly into the open contract */}
              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={quickDocClauseTitle}
                  onChange={(e) => setQuickDocClauseTitle(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && quickDocClauseTitle.trim()) {
                      onInsertNewClauseHeadingInDoc(quickDocClauseTitle.trim());
                      setQuickDocClauseTitle('');
                    }
                  }}
                  placeholder="عنوان بند جديد في هذا العقد (مثال: التعيين)..."
                  className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
                />
                <button
                  type="button"
                  onMouseDown={(e) => {
                    e.preventDefault();
                    onSaveSelectionBookmark();
                  }}
                  onClick={() => {
                    onInsertNewClauseHeadingInDoc(
                      quickDocClauseTitle.trim() || undefined
                    );
                    setQuickDocClauseTitle('');
                  }}
                  className="px-2.5 py-1.5 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800 shrink-0 inline-flex items-center gap-1 transition-colors"
                  title="إدراج عنوان بند جديد (#) عند المؤشر في العقد الحالي"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span># بند بالعقد</span>
                </button>
              </div>

              {liveContractClauses.length === 0 ? (
                <div className="border border-dashed border-blue-200 rounded p-3 text-center bg-white/80 space-y-1">
                  <div className="text-xs font-semibold text-slate-700">
                    لا توجد بنود مقسّمة بعلامة `#` في هذا العقد بعد
                  </div>
                  <p className="text-[11px] text-slate-500 leading-relaxed">
                    اكتب <code className="font-mono font-bold text-blue-900"># عنوان البند</code> في بداية أي سطر داخل الورقة أو اضغط زر <strong>«# بند بالعقد»</strong> أعلاه لتقسيم هذا العقد إلى بنود مستقلة قابلة للتفعيل والتعطيل والسحب.
                  </p>
                </div>
              ) : (
                <div className="space-y-2 max-h-80 overflow-y-auto pr-0.5">
                  {liveContractClauses.map((c, idx) => {
                    const isRenaming = renamingClauseId === c.id;
                    const isEnabled = c.enabled !== false;
                    const isLocked = !!c.locked;
                    const isDragOver = dragOverClauseId === c.id;
                    return (
                      <div
                        key={c.id}
                        draggable={!isRenaming}
                        onDragStart={(e) => {
                          setDraggedClauseId(c.id);
                          e.dataTransfer.effectAllowed = 'move';
                          e.dataTransfer.setData('text/plain', c.id);
                        }}
                        onDragOver={(e) => {
                          e.preventDefault();
                          e.dataTransfer.dropEffect = 'move';
                          if (draggedClauseId && draggedClauseId !== c.id) {
                            setDragOverClauseId(c.id);
                          }
                        }}
                        onDragLeave={() => {
                          if (dragOverClauseId === c.id) {
                            setDragOverClauseId(null);
                          }
                        }}
                        onDrop={(e) => {
                          e.preventDefault();
                          const srcId = draggedClauseId || e.dataTransfer.getData('text/plain');
                          setDraggedClauseId(null);
                          setDragOverClauseId(null);
                          if (srcId && srcId !== c.id && onReorderLiveClauses) {
                            onReorderLiveClauses(srcId, c.id);
                          }
                        }}
                        onDragEnd={() => {
                          setDraggedClauseId(null);
                          setDragOverClauseId(null);
                        }}
                        className={`border rounded-md p-2.5 space-y-1.5 transition-all ${
                          isDragOver
                            ? 'border-blue-600 bg-blue-50/70 ring-2 ring-blue-500/20'
                            : !isEnabled
                            ? 'border-slate-200 bg-slate-100/80 opacity-75'
                            : isLocked
                            ? 'border-amber-300 bg-amber-50/20'
                            : 'border-slate-200 bg-white hover:border-blue-400'
                        }`}
                      >
                        <div className="flex items-start justify-between gap-1.5">
                          {/* Drag Handle + Checkbox Toggle */}
                          <div className="flex items-center gap-1 pt-0.5 shrink-0">
                            <span
                              className="cursor-grab active:cursor-grabbing text-slate-400 hover:text-slate-700"
                              title="اسحب لإعادة ترتيب هذا البند في العقد"
                            >
                              <GripVertical className="w-3.5 h-3.5" />
                            </span>
                            <button
                              type="button"
                              onClick={() =>
                                onToggleLiveClauseEnabled && onToggleLiveClauseEnabled(c.id)
                              }
                              className={`p-0.5 rounded transition-colors ${
                                isEnabled
                                  ? 'text-blue-900 hover:text-blue-700'
                                  : 'text-slate-400 hover:text-slate-700'
                              }`}
                              title={
                                isEnabled
                                  ? 'البند مفعل وظاهر في ورقة A4 والتصدير — انقر لإخفائه مؤقتاً مع حفظ نصه وموضعه'
                                  : 'البند معطل ومخفي من ورقة A4 — انقر لإعادته فوراً إلى موضعه الأصلي'
                              }
                            >
                              {isEnabled ? (
                                <CheckSquare className="w-4 h-4" />
                              ) : (
                                <Square className="w-4 h-4" />
                              )}
                            </button>
                          </div>

                          {isRenaming ? (
                            <div className="flex items-center gap-1 flex-1">
                              <input
                                type="text"
                                value={renamingClauseVal}
                                onChange={(e) => setRenamingClauseVal(e.target.value)}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' && renamingClauseVal.trim()) {
                                    onRenameLiveClauseInDoc(
                                      c.id,
                                      renamingClauseVal.trim()
                                    );
                                    setRenamingClauseId(null);
                                  } else if (e.key === 'Escape') {
                                    setRenamingClauseId(null);
                                  }
                                }}
                                autoFocus
                                className="flex-1 px-2 py-0.5 text-xs border border-blue-800 rounded focus:outline-none bg-white"
                              />
                              <button
                                type="button"
                                onClick={() => {
                                  if (renamingClauseVal.trim()) {
                                    onRenameLiveClauseInDoc(
                                      c.id,
                                      renamingClauseVal.trim()
                                    );
                                  }
                                  setRenamingClauseId(null);
                                }}
                                className="p-1 text-emerald-700 hover:bg-emerald-50 rounded"
                              >
                                <Check className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() => setRenamingClauseId(null)}
                                className="p-1 text-slate-400 hover:bg-slate-100 rounded"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => {
                                if (isEnabled) {
                                  onScrollToLiveClause(c.id);
                                } else if (onToggleLiveClauseEnabled) {
                                  onToggleLiveClauseEnabled(c.id);
                                }
                              }}
                              className="text-right flex-1 group min-w-0"
                              title={
                                isEnabled
                                  ? 'انقر للقفز المباشر إلى موضع هذا البند داخل ورقة العقد'
                                  : 'هذا البند معطل ومخفي — انقر لإعادة تفعيله في موضعه'
                              }
                            >
                              <div className="text-xs font-bold text-slate-900 group-hover:text-blue-900 flex items-center gap-1">
                                <span className="text-blue-900 font-mono tabular-nums shrink-0">
                                  #{c.index}
                                </span>
                                <span
                                  className={`truncate ${
                                    !isEnabled ? 'line-through text-slate-500' : ''
                                  }`}
                                >
                                  {c.title}
                                </span>
                                {!isEnabled && (
                                  <span className="text-[9px] font-normal bg-slate-200 text-slate-700 px-1.5 py-0.5 rounded shrink-0">
                                    مخفي
                                  </span>
                                )}
                                {isLocked && isEnabled && (
                                  <span className="text-[9px] font-normal bg-amber-100 text-amber-900 px-1.5 py-0.5 rounded shrink-0">
                                    مقفل
                                  </span>
                                )}
                              </div>
                              {c.previewText && (
                                <p className="text-[11px] text-slate-500 line-clamp-1 mt-0.5">
                                  {c.previewText}
                                </p>
                              )}
                            </button>
                          )}

                          {/* Live Clause Controls inside this Contract */}
                          <div className="flex items-center gap-0.5 shrink-0">
                            <button
                              type="button"
                              onClick={() =>
                                onToggleLiveClauseLocked && onToggleLiveClauseLocked(c.id)
                              }
                              className={`p-1 rounded ${
                                isLocked
                                  ? 'text-amber-700 bg-amber-100/80 hover:bg-amber-200/70'
                                  : 'text-slate-400 hover:text-amber-700'
                              }`}
                              title={
                                isLocked
                                  ? 'البند مقفل ضد التعديل بالخطأ — انقر لفك القفل'
                                  : 'قفل البند لمنع تعديله بالخطأ في الورقة'
                              }
                            >
                              {isLocked ? (
                                <Lock className="w-3.5 h-3.5" />
                              ) : (
                                <Unlock className="w-3.5 h-3.5" />
                              )}
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                if (idx > 0 && onReorderLiveClauses) {
                                  onReorderLiveClauses(c.id, liveContractClauses[idx - 1].id);
                                } else {
                                  onMoveLiveClauseInDoc(c.domIndex, 'up');
                                }
                              }}
                              disabled={idx === 0}
                              className="p-1 text-slate-400 hover:text-slate-900 disabled:opacity-30"
                              title="تحريك البند لأعلى داخل هذا العقد"
                            >
                              <ArrowUp className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                if (
                                  idx < liveContractClauses.length - 1 &&
                                  onReorderLiveClauses
                                ) {
                                  onReorderLiveClauses(c.id, liveContractClauses[idx + 1].id);
                                } else {
                                  onMoveLiveClauseInDoc(c.domIndex, 'down');
                                }
                              }}
                              disabled={idx === liveContractClauses.length - 1}
                              className="p-1 text-slate-400 hover:text-slate-900 disabled:opacity-30"
                              title="تحريك البند لأسفل داخل هذا العقد"
                            >
                              <ArrowDown className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => {
                                setRenamingClauseId(c.id);
                                setRenamingClauseVal(c.title);
                              }}
                              className="p-1 text-slate-400 hover:text-blue-900"
                              title="تعديل عنوان البند في الورقة"
                            >
                              <Edit3 className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() =>
                                onSaveNewClause(c.title, 'بنود مستخرجة', c.contentHtml)
                              }
                              className="p-1 text-slate-400 hover:text-emerald-700"
                              title="حفظ نسخة من هذا البند في مكتبة المكتب لإعادة استخدامه في عقود أخرى"
                            >
                              <BookmarkPlus className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => onDeleteLiveClauseFromDoc(c.id)}
                              className="p-1 text-slate-400 hover:text-red-600"
                              title="حذف هذا البند نهائياً من العقد الحالي"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>

                        {c.variables.length > 0 && (
                          <div className="flex items-center flex-wrap gap-1 pt-1 border-t border-slate-100">
                            <span className="text-[10px] text-slate-400">المتغيرات:</span>
                            {c.variables.map((v) => (
                              <button
                                key={v}
                                type="button"
                                onClick={() => onOpenSmartVariablesModal(v)}
                                className={`text-[10px] font-mono px-1.5 py-0.5 rounded border ${
                                  fieldValues[v]?.trim()
                                    ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
                                    : 'bg-pink-50 text-pink-900 border-pink-200'
                                }`}
                              >
                                [{v}]
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* 2. OPTIONAL CLAUSES SUGGESTIONS LIBRARY (مكتبة البنود المقترحة للإدراج فقط) */}
            <div className="border border-slate-200 rounded-md bg-white overflow-hidden">
              <button
                type="button"
                onClick={() => setShowSuggestionsLibrary((v) => !v)}
                className="w-full px-3 py-2.5 bg-slate-50 hover:bg-slate-100 border-b border-slate-200 flex items-center justify-between text-right transition-colors"
              >
                <div>
                  <div className="text-xs font-bold text-slate-900">
                    مكتبة البنود المقترحة للإدراج ({clauses.length})
                  </div>
                  <div className="text-[10px] text-slate-500">
                    مكتبة مساعدة فارغة افتراضياً — لا تدخل في العقد إلا عند ضغط «+ إدراج في العقد»
                  </div>
                </div>
                {showSuggestionsLibrary ? (
                  <ChevronUp className="w-4 h-4 text-slate-500 shrink-0" />
                ) : (
                  <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
                )}
              </button>

              {showSuggestionsLibrary && (
                <div className="p-3 space-y-3">
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      value={clauseSearch}
                      onChange={(e) => setClauseSearch(e.target.value)}
                      placeholder="بحث في مكتبة البنود المقترحة..."
                      className="flex-1 px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded focus:bg-white focus:border-blue-900 focus:outline-none"
                    />
                    <button
                      type="button"
                      onClick={() => setShowAddLibraryForm((v) => !v)}
                      className="px-2.5 py-1.5 bg-slate-900 text-white text-[11px] font-medium rounded hover:bg-slate-800 shrink-0"
                    >
                      {showAddLibraryForm ? 'إغلاق' : '+ بند للمكتبة'}
                    </button>
                  </div>

                  {onImportStandardClausesPack && (
                    <button
                      type="button"
                      onClick={() => onImportStandardClausesPack()}
                      className="w-full py-1.5 px-2.5 bg-blue-50 hover:bg-blue-100/80 border border-blue-200 text-blue-950 text-[11px] font-semibold rounded inline-flex items-center justify-center gap-1.5 transition-colors"
                      title="تحميل حزمة البنود التوثيقية القياسية (التعيين، أصل الملكية، الثمن، التصاريح الجبائية، الوكالة، الهبة) عند الطلب"
                    >
                      <Sparkles className="w-3.5 h-3.5 text-blue-800" />
                      <span>استيراد الحزمة القياسية الاختيارية للمكتب</span>
                    </button>
                  )}

                  {showAddLibraryForm && (
                    <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-md space-y-2">
                      <div className="text-[11px] font-bold text-slate-900">
                        حفظ قالب بند جديد في مكتبة المكتب العامة
                      </div>
                      <input
                        type="text"
                        value={newClauseTitle}
                        onChange={(e) => setNewClauseTitle(e.target.value)}
                        placeholder="عنوان البند (مثال: بند أصل الملكية / بند الضمان)"
                        className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                      />
                      <input
                        type="text"
                        value={newClauseCategory}
                        onChange={(e) => setNewClauseCategory(e.target.value)}
                        placeholder="التصنيف (بيع، تأسيس شركة، إيجار...)"
                        className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                      />
                      <textarea
                        rows={3}
                        value={newClauseText}
                        onChange={(e) => setNewClauseText(e.target.value)}
                        placeholder="نص البند (يمكن تضمين وسوم مثل [البائع]) أو اتركه فارغاً لحفظ النص المحدد حالياً في الورقة..."
                        className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                      />
                      <div className="grid grid-cols-2 gap-1.5">
                        <button
                          type="button"
                          onClick={async () => {
                            if (!newClauseTitle.trim()) return;
                            const html = newClauseText.trim()
                              ? `<p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;text-align:justify;font-weight:bold;"># ${newClauseTitle.trim()}</p>` +
                                newClauseText
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
                            setShowAddLibraryForm(false);
                          }}
                          className="py-1.5 px-2 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800"
                        >
                          حفظ في المكتبة
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
                            setShowAddLibraryForm(false);
                          }}
                          className="py-1.5 px-2 bg-slate-800 text-white text-[11px] font-medium rounded hover:bg-slate-700"
                        >
                          حفظ المحدد بالورقة
                        </button>
                      </div>
                    </div>
                  )}

                  {clauses.length === 0 ? (
                    <div className="border border-dashed border-slate-200 rounded p-4 text-center text-xs text-slate-500 space-y-2">
                      <p>
                        مكتبة البنود المقترحة فارغة افتراضياً للحفاظ على خفة النظام. يمكنك حفظ أي بند من العقد بالضغط على أيقونة الحفظ بجانب البند، أو استيراد الحزمة القياسية أعلاه.
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-2 max-h-72 overflow-y-auto pr-0.5">
                      {clauses
                        .filter(
                          (c) =>
                            !clauseSearch.trim() ||
                            c.title.includes(clauseSearch.trim()) ||
                            c.category.includes(clauseSearch.trim())
                        )
                        .map((clause) => (
                          <div
                            key={clause.id}
                            className="border border-slate-200 rounded p-2.5 bg-slate-50/40 hover:bg-white transition-colors space-y-2"
                          >
                            <div className="flex items-start justify-between gap-2">
                              <div>
                                <div className="text-xs font-bold text-slate-900">
                                  {clause.title}
                                </div>
                                <div className="text-[10px] text-slate-500">
                                  {clause.category}
                                </div>
                              </div>
                              <div className="flex items-center gap-1 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => onEditClauseInEditor(clause)}
                                  className="p-1 text-slate-400 hover:text-blue-900"
                                  title="تعديل نص هذا البند في المكتبة"
                                >
                                  <Edit3 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => onDeleteClause(clause.id)}
                                  className="p-1 text-slate-400 hover:text-red-600"
                                  title="حذف من المكتبة"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </div>

                            <div className="flex items-center gap-1.5">
                              <button
                                type="button"
                                onMouseDown={(e) => {
                                  e.preventDefault();
                                  onSaveSelectionBookmark();
                                }}
                                onClick={() => onInsertClauseAtCaret(clause)}
                                className="w-full py-1 px-2 bg-white border border-slate-300 hover:border-blue-900 hover:bg-blue-50/40 text-slate-800 text-[11px] font-semibold rounded transition-colors"
                              >
                                + إدراج نسخة من البند عند المؤشر في العقد الحالي
                              </button>
                            </div>
                          </div>
                        ))}
                    </div>
                  )}
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

                {/* PHASE 4: STRUCTURED MULTI-PARTY CARDS (بطاقات الأطراف المهيكلة بالصفة والرقم) */}
                {onUpdatePartyCards && (
                  <div className="border border-blue-200 bg-blue-50/25 rounded-md p-3 space-y-2.5">
                    <div className="flex items-center justify-between border-b border-blue-200/80 pb-2">
                      <div className="flex items-center gap-1.5">
                        <Users className="w-3.5 h-3.5 text-blue-900" />
                        <span className="text-xs font-bold text-slate-900">
                          بطاقات أطراف العقد ({partyCards.length})
                        </span>
                      </div>
                      <span className="text-[10px] text-blue-900 font-medium">
                        ربط تلقائي بالوسوم [...]
                      </span>
                    </div>

                    {/* Add New Party Card Bar */}
                    <div className="flex items-center gap-1.5">
                      <select
                        value={newPartyRole}
                        onChange={(e) => setNewPartyRole(e.target.value as NotaryPartyRole)}
                        className="flex-1 px-2 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
                      >
                        <option value="بائع">صفة الطرف: بائع</option>
                        <option value="مشتري">صفة الطرف: مشتري</option>
                        <option value="موكل">صفة الطرف: موكل</option>
                        <option value="وكيل">صفة الطرف: وكيل</option>
                        <option value="واهب">صفة الطرف: واهب</option>
                        <option value="موهوب">صفة الطرف: موهوب له</option>
                        <option value="أخرى">صفة أخرى مخصصة...</option>
                      </select>
                      <button
                        type="button"
                        onClick={() => {
                          const sameRoleCount = partyCards.filter(
                            (p) => p.role === newPartyRole
                          ).length;
                          const newCard: ContractPartyCard = {
                            id: `pcard_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                            role: newPartyRole,
                            customRoleLabel: newPartyRole === 'أخرى' ? 'شريك' : undefined,
                            index: sameRoleCount + 1,
                            fullName: '',
                            birthDate: '',
                            birthPlace: '',
                            filiation: '',
                            nationalIdNin: '',
                            idCardDetails: '',
                            address: '',
                            legalRepresentative: '',
                          };
                          onUpdatePartyCards([...partyCards, newCard]);
                          setExpandedPartyCardId(newCard.id);
                        }}
                        className="px-2.5 py-1.5 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800 inline-flex items-center gap-1 shrink-0"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>+ إضافة طرف</span>
                      </button>
                    </div>

                    {partyCards.length === 0 ? (
                      <div className="text-[11px] text-slate-500 bg-white/80 border border-dashed border-blue-200 rounded p-2.5 text-center">
                        أضف بطاقات الأطراف (بائع 1، بائع 2، مشتري 1...) لترتبط تلقائياً بجميع الوسوم مثل <code className="font-mono text-blue-900">[البائع_1_الاسم]</code> و<code className="font-mono text-blue-900">[البائع.الاسم]</code>.
                      </div>
                    ) : (
                      <div className="space-y-2">
                        {partyCards.map((card) => {
                          const isExpanded = expandedPartyCardId === card.id;
                          const roleLabel =
                            card.role === 'أخرى'
                              ? card.customRoleLabel || 'طرف'
                              : card.role;
                          const defRole = roleLabel.startsWith('ال')
                            ? roleLabel
                            : `ال${roleLabel}`;
                          const tagPrefix = `${defRole}_${card.index}`;

                          const updateThisCard = (patch: Partial<ContractPartyCard>) => {
                            const updated = partyCards.map((c) =>
                              c.id === card.id ? { ...c, ...patch } : c
                            );
                            // Recompute sequential index per role
                            const roleCounters: Record<string, number> = {};
                            const reindexed = updated.map((c) => {
                              const rKey =
                                c.role === 'أخرى' ? c.customRoleLabel || 'أخرى' : c.role;
                              roleCounters[rKey] = (roleCounters[rKey] || 0) + 1;
                              return { ...c, index: roleCounters[rKey] };
                            });
                            onUpdatePartyCards(reindexed);
                          };

                          return (
                            <div
                              key={card.id}
                              className="bg-white border border-slate-200 rounded-md overflow-hidden"
                            >
                              <div className="px-2.5 py-2 bg-slate-50 border-b border-slate-200 flex items-center justify-between gap-1.5">
                                <button
                                  type="button"
                                  onClick={() =>
                                    setExpandedPartyCardId(isExpanded ? null : card.id)
                                  }
                                  className="flex items-center gap-1.5 text-right flex-1 min-w-0"
                                >
                                  <span className="px-1.5 py-0.5 text-[10px] font-bold bg-blue-900 text-white rounded shrink-0">
                                    {roleLabel} {card.index}
                                  </span>
                                  <span className="text-xs font-bold text-slate-900 truncate">
                                    {card.fullName || 'بدون اسم بعد...'}
                                  </span>
                                </button>
                                <div className="flex items-center gap-1 shrink-0">
                                  {onSavePartyCardToDirectory && card.fullName.trim() && (
                                    <button
                                      type="button"
                                      onClick={() => onSavePartyCardToDirectory(card)}
                                      className="p-1 text-slate-500 hover:text-blue-900"
                                      title="حفظ هذا الطرف في دفتر الأطراف المحفوظين"
                                    >
                                      <Save className="w-3.5 h-3.5" />
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const filtered = partyCards.filter(
                                        (c) => c.id !== card.id
                                      );
                                      const roleCounters: Record<string, number> = {};
                                      const reindexed = filtered.map((c) => {
                                        const rKey =
                                          c.role === 'أخرى'
                                            ? c.customRoleLabel || 'أخرى'
                                            : c.role;
                                        roleCounters[rKey] = (roleCounters[rKey] || 0) + 1;
                                        return { ...c, index: roleCounters[rKey] };
                                      });
                                      onUpdatePartyCards(reindexed);
                                    }}
                                    className="p-1 text-slate-400 hover:text-red-600"
                                    title="حذف بطاقة هذا الطرف"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setExpandedPartyCardId(isExpanded ? null : card.id)
                                    }
                                    className="p-1 text-slate-500"
                                  >
                                    {isExpanded ? (
                                      <ChevronUp className="w-3.5 h-3.5" />
                                    ) : (
                                      <ChevronDown className="w-3.5 h-3.5" />
                                    )}
                                  </button>
                                </div>
                              </div>

                              {isExpanded && (
                                <div className="p-2.5 space-y-2 text-xs">
                                  {/* Recall from Saved Parties directly into this card */}
                                  {savedParties.length > 0 && (
                                    <div className="flex items-center gap-1 pb-1.5 border-b border-slate-100">
                                      <select
                                        value={
                                          savedParties.find(
                                            (p) => p.id === selectedPartyIdSidebar
                                          )?.id ||
                                          savedParties[0]?.id ||
                                          ''
                                        }
                                        onChange={(e) =>
                                          setSelectedPartyIdSidebar(e.target.value)
                                        }
                                        className="flex-1 px-2 py-1 text-[11px] bg-slate-50 border border-slate-300 rounded"
                                      >
                                        {savedParties.map((p) => (
                                          <option key={p.id} value={p.id}>
                                            {p.fullName}
                                          </option>
                                        ))}
                                      </select>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const sp =
                                            savedParties.find(
                                              (p) => p.id === selectedPartyIdSidebar
                                            ) || savedParties[0];
                                          if (!sp) return;
                                          updateThisCard({
                                            fullName: sp.fullName,
                                            birthDate: sp.birthDate,
                                            birthPlace: sp.birthPlace,
                                            filiation: sp.parentage,
                                            idCardDetails: sp.idCardRef,
                                            address: sp.address,
                                          });
                                        }}
                                        className="px-2 py-1 bg-slate-800 text-white text-[10px] font-semibold rounded hover:bg-slate-700 shrink-0"
                                      >
                                        تعبئة من الدفتر
                                      </button>
                                    </div>
                                  )}

                                  <div className="grid grid-cols-2 gap-1.5">
                                    <div>
                                      <label className="text-[10px] text-slate-500 block mb-0.5">
                                        الصفة في العقد
                                      </label>
                                      <select
                                        value={card.role}
                                        onChange={(e) =>
                                          updateThisCard({
                                            role: e.target.value as NotaryPartyRole,
                                          })
                                        }
                                        className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                      >
                                        <option value="بائع">بائع</option>
                                        <option value="مشتري">مشتري</option>
                                        <option value="موكل">موكل</option>
                                        <option value="وكيل">وكيل</option>
                                        <option value="واهب">واهب</option>
                                        <option value="موهوب">موهوب له</option>
                                        <option value="أخرى">أخرى</option>
                                      </select>
                                    </div>
                                    {card.role === 'أخرى' && (
                                      <div>
                                        <label className="text-[10px] text-slate-500 block mb-0.5">
                                          تسمية الصفة
                                        </label>
                                        <input
                                          type="text"
                                          value={card.customRoleLabel || ''}
                                          onChange={(e) =>
                                            updateThisCard({
                                              customRoleLabel: e.target.value,
                                            })
                                          }
                                          placeholder="مثال: شريك / مؤجر"
                                          className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                        />
                                      </div>
                                    )}
                                  </div>

                                  <div>
                                    <label className="text-[10px] text-slate-600 font-semibold block mb-0.5">
                                      الاسم واللقب الكامل [{tagPrefix}_الاسم]
                                    </label>
                                    <input
                                      type="text"
                                      value={card.fullName}
                                      onChange={(e) =>
                                        updateThisCard({ fullName: e.target.value })
                                      }
                                      placeholder="الاسم واللقب الكامل..."
                                      className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded focus:bg-white focus:border-blue-900 focus:outline-none"
                                    />
                                  </div>

                                  <div className="grid grid-cols-2 gap-1.5">
                                    <div>
                                      <label className="text-[10px] text-slate-500 block mb-0.5">
                                        تاريخ الميلاد
                                      </label>
                                      <input
                                        type="date"
                                        value={card.birthDate}
                                        onChange={(e) =>
                                          updateThisCard({ birthDate: e.target.value })
                                        }
                                        className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                      />
                                    </div>
                                    <div>
                                      <label className="text-[10px] text-slate-500 block mb-0.5">
                                        مكان الميلاد
                                      </label>
                                      <input
                                        type="text"
                                        value={card.birthPlace}
                                        onChange={(e) =>
                                          updateThisCard({ birthPlace: e.target.value })
                                        }
                                        placeholder="البلدية والولاية..."
                                        className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                      />
                                    </div>
                                  </div>

                                  <div>
                                    <label className="text-[10px] text-slate-500 block mb-0.5">
                                      النسب (ابن / ابنة فلان وفلانة)
                                    </label>
                                    <input
                                      type="text"
                                      value={card.filiation}
                                      onChange={(e) =>
                                        updateThisCard({ filiation: e.target.value })
                                      }
                                      placeholder="بن ... وأمه ..."
                                      className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                    />
                                  </div>

                                  <div className="grid grid-cols-2 gap-1.5">
                                    <div>
                                      <label className="text-[10px] text-slate-500 block mb-0.5">
                                        الرقم الوطني (NIN)
                                      </label>
                                      <input
                                        type="text"
                                        value={card.nationalIdNin}
                                        onChange={(e) =>
                                          updateThisCard({
                                            nationalIdNin: e.target.value,
                                          })
                                        }
                                        placeholder="18 رقماً..."
                                        className="w-full px-2 py-1 text-xs font-mono bg-slate-50 border border-slate-300 rounded"
                                      />
                                    </div>
                                    <div>
                                      <label className="text-[10px] text-slate-500 block mb-0.5">
                                        بطاقة الهوية / ر.س
                                      </label>
                                      <input
                                        type="text"
                                        value={card.idCardDetails}
                                        onChange={(e) =>
                                          updateThisCard({
                                            idCardDetails: e.target.value,
                                          })
                                        }
                                        placeholder="رقم وتاريخ الصدور..."
                                        className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                      />
                                    </div>
                                  </div>

                                  <div>
                                    <label className="text-[10px] text-slate-500 block mb-0.5">
                                      عنوان الإقامة الكامل
                                    </label>
                                    <input
                                      type="text"
                                      value={card.address}
                                      onChange={(e) =>
                                        updateThisCard({ address: e.target.value })
                                      }
                                      placeholder="الحي، البلدية، الولاية..."
                                      className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                                    />
                                  </div>

                                  {/* Quick insert party tags at caret */}
                                  <div className="pt-1 border-t border-slate-100 flex items-center gap-1 flex-wrap">
                                    <button
                                      type="button"
                                      onMouseDown={(e) => {
                                        e.preventDefault();
                                        onSaveSelectionBookmark();
                                      }}
                                      onClick={() =>
                                        onInsertPlaceholderAtCaret(`${tagPrefix}_الاسم`)
                                      }
                                      className="px-2 py-0.5 bg-pink-50 hover:bg-pink-100 text-pink-900 border border-pink-200 rounded text-[10px] font-mono"
                                    >
                                      + [{tagPrefix}_الاسم]
                                    </button>
                                    <button
                                      type="button"
                                      onMouseDown={(e) => {
                                        e.preventDefault();
                                        onSaveSelectionBookmark();
                                      }}
                                      onClick={() =>
                                        onInsertPlaceholderAtCaret(
                                          `${tagPrefix}_الرقم_الوطني`
                                        )
                                      }
                                      className="px-2 py-0.5 bg-pink-50 hover:bg-pink-100 text-pink-900 border border-pink-200 rounded text-[10px] font-mono"
                                    >
                                      + [{tagPrefix}_الرقم_الوطني]
                                    </button>
                                    <button
                                      type="button"
                                      onMouseDown={(e) => {
                                        e.preventDefault();
                                        onSaveSelectionBookmark();
                                      }}
                                      onClick={() =>
                                        onInsertPlaceholderAtCaret(`${tagPrefix}_العنوان`)
                                      }
                                      className="px-2 py-0.5 bg-pink-50 hover:bg-pink-100 text-pink-900 border border-pink-200 rounded text-[10px] font-mono"
                                    >
                                      + [{tagPrefix}_العنوان]
                                    </button>
                                  </div>
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                )}

                {/* Conditional Clauses Management (إدارة شروط البنود) */}
                <div className="border border-slate-200 rounded-md p-3 bg-white space-y-2.5">
                  <span className="text-xs font-bold text-slate-900">
                    شروط البنود والخصائص (تُضاف حسب الحاجة)
                  </span>
                  <div className="text-[11px] text-slate-500 italic">
                    لم يتم تعريف أي شروط بعد. أضف شروطاً خاصة بعقود مكتبك هنا.
                  </div>
                </div>

                {/* Saved Parties & Properties Directory (دفتر الأطراف والعقارات القابل للاستدعاء) */}
                <div className="border border-slate-200 rounded-md p-3 bg-white space-y-2.5">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-900">
                      دفتر الأطراف والعقارات (استدعاء وحفظ سريع)
                    </span>
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => onSaveCurrentPartyToDirectory('party1')}
                        className="px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-[10px] font-semibold text-slate-800 rounded"
                        title="حفظ بيانات الطرف الأول الحالية في الدفتر"
                      >
                        + حفظ طرف 1
                      </button>
                      <button
                        type="button"
                        onClick={() => onSaveCurrentPartyToDirectory('party2')}
                        className="px-1.5 py-0.5 bg-slate-100 hover:bg-slate-200 text-[10px] font-semibold text-slate-800 rounded"
                        title="حفظ بيانات الطرف الثاني الحالية في الدفتر"
                      >
                        + حفظ طرف 2
                      </button>
                    </div>
                  </div>

                  {/* Recall Saved Party */}
                  {savedParties.length > 0 ? (
                    <div className="space-y-1.5">
                      <select
                        value={
                          savedParties.find((p) => p.id === selectedPartyIdSidebar)?.id ||
                          savedParties[0]?.id ||
                          ''
                        }
                        onChange={(e) => setSelectedPartyIdSidebar(e.target.value)}
                        className="w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                      >
                        {savedParties.map((p) => (
                          <option key={p.id} value={p.id}>
                            {p.fullName} {p.birthDate ? `(${p.birthDate})` : ''}
                          </option>
                        ))}
                      </select>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => {
                            const target =
                              savedParties.find((p) => p.id === selectedPartyIdSidebar) ||
                              savedParties[0];
                            if (target) onRecallPartyToRole(target, 'party1');
                          }}
                          className="flex-1 py-1 px-2 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800"
                        >
                          استدعاء كطرف أول
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            const target =
                              savedParties.find((p) => p.id === selectedPartyIdSidebar) ||
                              savedParties[0];
                            if (target) onRecallPartyToRole(target, 'party2');
                          }}
                          className="flex-1 py-1 px-2 bg-slate-800 text-white text-[11px] font-medium rounded hover:bg-slate-700"
                        >
                          استدعاء كطرف ثانٍ
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div className="text-[11px] text-slate-500">
                      لا يوجد أطراف محفوظون بعد. املأ بيانات طرف ثم اضغط «+ حفظ طرف 1/2».
                    </div>
                  )}

                  {/* Recall / Save Property */}
                  <div className="border-t border-slate-100 pt-2 space-y-1.5">
                    <div className="flex items-center justify-between">
                      <span className="text-[11px] font-semibold text-slate-700">
                        العقارات المحفوظة ({savedProperties.length})
                      </span>
                      <button
                        type="button"
                        onClick={onSaveCurrentPropertyToDirectory}
                        className="px-1.5 py-0.5 bg-blue-50 hover:bg-blue-100 text-[10px] font-semibold text-blue-900 rounded"
                      >
                        + حفظ العقار الحالي
                      </button>
                    </div>
                    {savedProperties.length > 0 && (
                      <div className="flex items-center gap-1">
                        <select
                          value={
                            savedProperties.find((p) => p.id === selectedPropIdSidebar)?.id ||
                            savedProperties[0]?.id ||
                            ''
                          }
                          onChange={(e) => setSelectedPropIdSidebar(e.target.value)}
                          className="flex-1 px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded"
                        >
                          {savedProperties.map((pr) => (
                            <option key={pr.id} value={pr.id}>
                              {pr.label} {pr.lotNumber ? `(حصة ${pr.lotNumber})` : ''}
                            </option>
                          ))}
                        </select>
                        <button
                          type="button"
                          onClick={() => {
                            const target =
                              savedProperties.find((p) => p.id === selectedPropIdSidebar) ||
                              savedProperties[0];
                            if (target) onRecallPropertyRecord(target);
                          }}
                          className="px-2.5 py-1 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800 shrink-0"
                        >
                          استدعاء
                        </button>
                      </div>
                    )}
                  </div>
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

                {/* Dynamic Variables Grouped by Clause in Current Document (STRICTLY EXTRACTED FROM [...] IN OPEN CONTRACT) */}
                <div className="space-y-2.5">
                  <div className="text-xs font-bold text-slate-900 flex items-center justify-between border-b border-slate-200 pb-1.5">
                    <span>
                      متغيرات هذا العقد المستخرجة من `[...]` ({extractedPlaceholders.length})
                    </span>
                    {extractedPlaceholders.length > 0 && (
                      <button
                        type="button"
                        onClick={() => onBakeAllPlaceholdersIntoDocument(fieldValues)}
                        className="text-[11px] text-pink-800 font-semibold hover:underline"
                      >
                        دمج الكل في النص
                      </button>
                    )}
                  </div>

                  {clauseGroups.length > 0 ? (
                    clauseGroups.map((group) => (
                      <div
                        key={group.clauseId}
                        className="border border-pink-200 bg-pink-50/30 rounded-md p-2.5 space-y-2"
                      >
                        <div className="text-[11px] font-bold text-pink-950 border-b border-pink-200/70 pb-1 flex items-center justify-between">
                          <span className="truncate">{group.clauseTitle}</span>
                          <span className="text-[10px] font-mono bg-white px-1.5 py-0.5 rounded border border-pink-200 shrink-0 tabular-nums">
                            {group.variables.length} متغير
                          </span>
                        </div>
                        {group.variables.map((key) => {
                          const meta = partyFields.find((f) => f.key === key);
                          const label = meta?.label || key.replace(/_/g, ' ');
                          const perContractType = fieldInputTypes[key];
                          const resolvedType = perContractType || meta?.inputType;
                          const currentType: 'text' | 'number' | 'date' =
                            resolvedType === 'number'
                              ? 'number'
                              : resolvedType === 'date'
                              ? 'date'
                              : 'text';
                          return (
                            <div
                              key={key}
                              className="space-y-1 bg-white p-2 rounded border border-slate-200"
                            >
                              <div className="flex items-center justify-between gap-1">
                                <label className="text-[11px] font-semibold text-slate-800 truncate">
                                  {label}
                                </label>
                                <div className="flex items-center gap-1 shrink-0">
                                  <div className="inline-flex items-center bg-slate-100 rounded p-0.5 text-[9px]">
                                    <button
                                      type="button"
                                      onClick={() => onChangeFieldInputType(key, 'text')}
                                      className={`px-1 rounded ${
                                        currentType === 'text'
                                          ? 'bg-white text-slate-900 font-bold shadow-2xs'
                                          : 'text-slate-500'
                                      }`}
                                    >
                                      نص
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => onChangeFieldInputType(key, 'number')}
                                      className={`px-1 rounded ${
                                        currentType === 'number'
                                          ? 'bg-white text-blue-900 font-bold shadow-2xs'
                                          : 'text-slate-500'
                                      }`}
                                    >
                                      رقم
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => onChangeFieldInputType(key, 'date')}
                                      className={`px-1 rounded ${
                                        currentType === 'date'
                                          ? 'bg-white text-blue-900 font-bold shadow-2xs'
                                          : 'text-slate-500'
                                      }`}
                                    >
                                      تاريخ
                                    </button>
                                  </div>
                                  <button
                                    type="button"
                                    onMouseDown={(e) => {
                                      e.preventDefault();
                                      onSaveSelectionBookmark();
                                    }}
                                    onClick={() => onInsertPlaceholderAtCaret(key)}
                                    className="text-[10px] text-pink-800 font-semibold hover:underline"
                                    title={`إدراج الوسم [${key}] مرة أخرى عند المؤشر`}
                                  >
                                    +وسم
                                  </button>
                                </div>
                              </div>
                              <input
                                type={
                                  currentType === 'date'
                                    ? 'date'
                                    : currentType === 'number'
                                    ? 'number'
                                    : 'text'
                                }
                                value={fieldValues[key] || ''}
                                onChange={(e) => onUpdateFieldValue(key, e.target.value)}
                                placeholder={`أدخل ${label} [${key}]...`}
                                className={`w-full px-2 py-1 text-xs bg-slate-50 border border-slate-300 rounded focus:bg-white focus:border-blue-800 focus:outline-none ${
                                  currentType === 'number' ? 'tabular-nums font-mono' : ''
                                }`}
                              />
                            </div>
                          );
                        })}
                      </div>
                    ))
                  ) : (
                    <div className="border border-dashed border-pink-200 bg-pink-50/20 rounded-md p-3.5 text-center space-y-1.5">
                      <div className="text-xs font-bold text-slate-800">
                        لا توجد وسوم `[...]` مكتوبة في هذا العقد حالياً
                      </div>
                      <p className="text-[11px] text-slate-500 leading-relaxed">
                        تُستخرج حقول هذه الاستمارة حصرياً من الوسوم المكتوبة بصيغة <code className="font-mono font-bold text-pink-800">[اسم_المتغير]</code> داخل ورقة العقد المفتوح حالياً لضمان عدم ظهور متغيرات لعقود سابقة.
                      </p>
                    </div>
                  )}
                </div>

                {/* Insert New Variable Directly Into Current Contract at Caret */}
                <div className="border border-slate-200 rounded-md p-2.5 bg-white">
                  <div className="text-xs font-semibold text-slate-800 mb-1.5">
                    إدراج وسم متغير جديد `[...]` في العقد عند المؤشر
                  </div>
                  <div className="flex items-center gap-1.5">
                    <input
                      type="text"
                      value={newFieldKey}
                      onChange={(e) => setNewFieldKey(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && newFieldKey.trim()) {
                          const clean = newFieldKey.trim().replace(/\s+/g, '_');
                          onAddCustomField(clean, newFieldKey.trim());
                          onInsertPlaceholderAtCaret(clean);
                          setNewFieldKey('');
                        }
                      }}
                      placeholder="مثال: تسمية_الشركة أو البائع أو الثمن"
                      className="flex-1 px-2 py-1 text-xs border border-slate-300 rounded focus:border-blue-800 focus:outline-none"
                    />
                    <button
                      type="button"
                      onMouseDown={(e) => {
                        e.preventDefault();
                        onSaveSelectionBookmark();
                      }}
                      onClick={() => {
                        const clean = newFieldKey.trim().replace(/\s+/g, '_');
                        if (clean) {
                          onAddCustomField(clean, newFieldKey.trim());
                          onInsertPlaceholderAtCaret(clean);
                          setNewFieldKey('');
                        }
                      }}
                      className="px-2.5 py-1 bg-pink-800 text-white text-xs font-medium rounded hover:bg-pink-900 shrink-0"
                    >
                      + إدراج بالعقد
                    </button>
                  </div>
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
            TAB 3: OFFICE CONTRACT TEMPLATES ONLY (قوالب عقود المكتب فقط)
           ========================================================= */}
        {activeTab === 'templates' && (
          <div className="space-y-4">
            <div className="border border-blue-200 bg-blue-50/40 rounded-md p-3 space-y-2.5">
              <div className="text-xs font-bold text-slate-900">
                استيراد نماذج العقود الخاصة بمكتبكم (.docx)
              </div>
              <label className="w-full flex items-center justify-center gap-2 py-2 px-3 bg-blue-900 text-white rounded text-xs font-medium hover:bg-blue-800 transition-colors cursor-pointer">
                <Upload className="w-4 h-4" />
                <span>استيراد قالب عقد Word (.docx) من الجهاز</span>
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

            {/* Save Current Contract as Template Box */}
            <div className="border border-slate-200 rounded-md p-3 space-y-2 bg-white">
              <div className="flex items-center justify-between">
                <span className="text-xs font-bold text-slate-900">
                  حفظ العقد الحالي كقالب مكتب
                </span>
                {onOpenSaveAsTemplateModal && (
                  <button
                    type="button"
                    onClick={onOpenSaveAsTemplateModal}
                    className="text-[11px] text-blue-900 hover:underline font-semibold"
                    title="فتح نافذة الخيارات المتقدمة وتفريغ القيم المعبأة"
                  >
                    خيارات متقدمة...
                  </button>
                )}
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
                  placeholder="التصنيف (مثال: عقود البيع)"
                  className="flex-1 px-2.5 py-1.5 text-xs border border-slate-300 rounded"
                />
                <button
                  type="button"
                  onClick={async () => {
                    if (!newTplName.trim()) return;
                    if (onOpenSaveAsTemplateModal) {
                      onOpenSaveAsTemplateModal();
                    } else {
                      await onSaveCurrentAsTemplate(
                        newTplName.trim(),
                        newTplCategory.trim() || 'عام'
                      );
                      setNewTplName('');
                    }
                  }}
                  className="px-3 py-1.5 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800 shrink-0"
                >
                  حفظ كقالب
                </button>
              </div>
            </div>

            {/* Templates JSON Export / Import Strip */}
            <div className="flex items-center gap-2">
              {onExportTemplatesJson && (
                <button
                  type="button"
                  onClick={onExportTemplatesJson}
                  disabled={templates.length === 0}
                  className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-[11px] font-medium rounded inline-flex items-center justify-center gap-1 disabled:opacity-40"
                  title="تصدير جميع قوالب مكتبكم المحفوظة إلى ملف JSON"
                >
                  <Download className="w-3 h-3 text-blue-900" />
                  <span>تصدير القوالب (JSON)</span>
                </button>
              )}

              {onImportTemplatesJson && (
                <label
                  className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-[11px] font-medium rounded inline-flex items-center justify-center gap-1 cursor-pointer"
                  title="استيراد قوالب من ملف JSON تم تصديره مسبقاً"
                >
                  <Upload className="w-3 h-3 text-blue-900" />
                  <span>استيراد قوالب (JSON)</span>
                  <input
                    type="file"
                    accept=".json,application/json"
                    className="hidden"
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        onImportTemplatesJson(file);
                        e.target.value = '';
                      }
                    }}
                  />
                </label>
              )}
            </div>

            <div className="space-y-2">
              <div className="flex items-center justify-between border-b border-slate-200 pb-1.5">
                <span className="text-xs font-bold text-slate-900">
                  مكتبة قوالب عقود المكتب ({templates.length})
                </span>
              </div>

              {/* Checkbox: Clear Previous Clauses when Replacing Full Content */}
              <label className="flex items-center gap-2 p-2 bg-blue-50/60 border border-blue-200 rounded text-[11px] text-blue-950 cursor-pointer">
                <input
                  type="checkbox"
                  checked={clearClausesOnReplace}
                  onChange={(e) => setClearClausesOnReplace(e.target.checked)}
                  className="rounded border-slate-300 text-blue-900"
                />
                <span className="font-semibold">
                  حذف البنود السابقة المفعّلة عند استبدال المحتوى بالقالب
                </span>
              </label>

              {templates.length === 0 ? (
                <div className="border border-dashed border-slate-300 rounded-md p-5 text-center space-y-2 bg-slate-50/50">
                  <FileText className="w-7 h-7 text-slate-400 mx-auto" />
                  <div className="text-xs font-semibold text-slate-700">
                    مكتبة قوالب العقود فارغة حالياً
                  </div>
                  <p className="text-[11px] text-slate-500">
                    استوردوا ملفات Word (.docx) الخاصة بمكتبكم أو احفظوا أي عقد من المحرر. (الوثائق المشتقة متوفرة في شريط التوثيق العلوي).
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
                        onClick={() => onLoadTemplateFull(tpl, clearClausesOnReplace)}
                        className="flex-1 py-1 px-2 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800"
                        title="يمسح ورقة الـ A4 ويضع القالب كاملاً"
                      >
                        استبدال المحتوى الحالي
                      </button>
                      <button
                        type="button"
                        onMouseDown={(e) => {
                          e.preventDefault();
                          onSaveSelectionBookmark();
                        }}
                        onClick={() => onInsertTemplateAtCaret(tpl)}
                        className="flex-1 py-1 px-2 bg-slate-100 text-slate-800 border border-slate-200 text-[11px] font-medium rounded hover:bg-slate-200"
                        title="يُبقي النص الحالي ويدرج القالب عند موضع المؤشر"
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

        {/* =========================================================
            TAB 4: SAVED DOCUMENTS, VERSION DIFF & DOWNLOAD ARCHIVE
           ========================================================= */}
        {activeTab === 'documents' && (
          <FolderTreeExplorer
            folders={folders}
            documents={documents}
            clerks={clerks}
            activeClerk={activeClerk}
            activeDocument={activeDocument}
            activeDocumentId={activeDocumentId}
            currentEditorTitle={currentEditorTitle || 'عقد توثيقي جديد'}
            onSelectDocument={onOpenDocument}
            onSaveCurrentToFolder={(targetFolderId) => onSaveCurrentDocument(targetFolderId)}
            onExportDocumentDocx={onExportDocumentDocx}
            onCreateFolder={onCreateFolder}
            onRenameFolder={onRenameFolder}
            onDeleteFolder={onDeleteFolder}
            onMoveDocument={onMoveDocument}
            onMoveFolder={onMoveFolder}
            onCreateContractInFolder={onCreateContractInFolder}
            onDeleteDocument={onDeleteDocument}
            onUpdateDocumentMeta={onUpdateDocumentMeta}
            derivedTemplates={derivedTemplates}
            onOpenDerivedModal={onOpenDerivedModal}
            onSelectDerivedDoc={onSelectDerivedDoc}
            activeDerivedDocId={activeDerivedDocId}
          />
        )}
      </div>
    </aside>
  );
}
