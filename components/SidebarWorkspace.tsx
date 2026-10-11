'use client';

import React from 'react';
import {
  FileText,
  FolderArchive,
  ListChecks,
  PanelRightClose,
  ShieldCheck,
  UserCheck,
} from 'lucide-react';
import {
  ClauseCondition,
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
import ClausesPanel from './ClausesPanel';
import NotaryHubPanel from './NotaryHubPanel';
import PartiesTabPanel from './sidebar/PartiesTabPanel';
import TemplatesTabPanel from './sidebar/TemplatesTabPanel';

export type SidebarTab = 'clauses' | 'parties' | 'templates' | 'documents' | 'notary_hub';

interface SidebarWorkspaceProps {
  activeTab: SidebarTab;
  onSelectTab: (tab: SidebarTab) => void;
  // 1. Live Contract Clauses + Optional Ready Clauses Library
  liveContractClauses: ContractOutlineClause[];
  contractConditions?: Record<string, string | boolean>;
  onUpdateContractCondition?: (key: string, value: string | boolean | undefined) => void;
  onSetClauseCondition?: (clauseId: string, condition: ClauseCondition | undefined) => void;
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
  onContractIndexedOrSynced?: (updatedDoc?: SavedDocument) => Promise<void> | void;
  showToast: (msg: string) => void;
}

export default function SidebarWorkspace({
  activeTab,
  onSelectTab,
  liveContractClauses,
  contractConditions,
  onUpdateContractCondition,
  onSetClauseCondition,
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
  onInsertClauseAtCaret,
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
  onRecallPropertyRecord,
  onSaveCurrentPropertyToDirectory,
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
  documents,
  activeDocumentId,
  currentEditorTitle,
  onOpenDocument,
  onSaveCurrentDocument,
  onDeleteDocument,
  onExportDocumentDocx,
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
  onContractIndexedOrSynced,
  showToast,
}: SidebarWorkspaceProps) {
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
            {activeTab === 'notary_hub' && 'الفهارس ومنظومة الموثق'}
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

      {/* Sidebar Segmented Navigation */}
      <div className="grid grid-cols-5 gap-1 p-2 bg-slate-50 border-b border-slate-200 shrink-0">
        <button
          type="button"
          onClick={() =>
            activeTab === 'clauses' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('clauses')
          }
          title="البنود الجاهزة"
          className={`flex flex-col items-center justify-center py-1.5 px-0.5 rounded text-[10px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'clauses'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <ListChecks className="w-3.5 h-3.5 mb-0.5" />
          <span>البنود</span>
        </button>

        <button
          type="button"
          onClick={() =>
            activeTab === 'parties' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('parties')
          }
          title="الأطراف والتعيينات"
          className={`flex flex-col items-center justify-center py-1.5 px-0.5 rounded text-[10px] font-medium transition-colors whitespace-nowrap relative ${
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

        <button
          type="button"
          onClick={() =>
            activeTab === 'templates' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('templates')
          }
          title="قوالب عقود المكتب"
          className={`flex flex-col items-center justify-center py-1.5 px-0.5 rounded text-[10px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'templates'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <FileText className="w-3.5 h-3.5 mb-0.5" />
          <span>القوالب</span>
        </button>

        <button
          type="button"
          onClick={() =>
            activeTab === 'documents' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('documents')
          }
          title="المستندات والأرشيف المحلي"
          className={`flex flex-col items-center justify-center py-1.5 px-0.5 rounded text-[10px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'documents'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <FolderArchive className="w-3.5 h-3.5 mb-0.5" />
          <span>الأرشيف</span>
        </button>

        <button
          type="button"
          onClick={() =>
            activeTab === 'notary_hub' && onToggleCollapse
              ? onToggleCollapse()
              : onSelectTab('notary_hub')
          }
          title="الفهارس، المفتشية والشهر العقاري"
          className={`flex flex-col items-center justify-center py-1.5 px-0.5 rounded text-[10px] font-medium transition-colors whitespace-nowrap ${
            activeTab === 'notary_hub'
              ? 'bg-blue-900 text-white shadow-2xs'
              : 'text-slate-600 hover:bg-slate-200/70 hover:text-slate-900'
          }`}
        >
          <ShieldCheck className="w-3.5 h-3.5 mb-0.5" />
          <span>الفهارس</span>
        </button>
      </div>

      {/* Tab Content Area */}
      <div className="flex-1 min-h-0 overflow-y-auto p-3 space-y-4">
        {activeTab === 'clauses' && (
          <ClausesPanel
            liveContractClauses={liveContractClauses}
            contractConditions={contractConditions}
            onUpdateContractCondition={onUpdateContractCondition}
            onSetClauseCondition={onSetClauseCondition}
            onScrollToLiveClause={(id) => onScrollToLiveClause(id)}
            onToggleLiveClauseEnabled={onToggleLiveClauseEnabled}
            onToggleLiveClauseLocked={onToggleLiveClauseLocked}
            onReorderLiveClauses={onReorderLiveClauses}
            onMoveLiveClauseInDoc={onMoveLiveClauseInDoc}
            onDeleteLiveClauseFromDoc={(id) => onDeleteLiveClauseFromDoc(id)}
            onRenameLiveClauseInDoc={(id, title) => onRenameLiveClauseInDoc(id, title)}
            onInsertNewClauseHeadingInDoc={onInsertNewClauseHeadingInDoc}
            onImportStandardClausesPack={onImportStandardClausesPack}
            clauses={clauses}
            onInsertClauseAtCaret={onInsertClauseAtCaret}
            onSaveNewClause={onSaveNewClause}
            onSaveSelectionAsClause={onSaveSelectionAsClause}
            onEditClauseInEditor={onEditClauseInEditor}
            onDeleteClause={onDeleteClause}
            fieldValues={fieldValues}
            onOpenSmartVariablesModal={onOpenSmartVariablesModal}
            onSaveSelectionBookmark={onSaveSelectionBookmark}
          />
        )}

        {activeTab === 'parties' && (
          <PartiesTabPanel
            partyCards={partyCards}
            onUpdatePartyCards={onUpdatePartyCards}
            onSavePartyCardToDirectory={onSavePartyCardToDirectory}
            partyFields={partyFields}
            extractedPlaceholders={extractedPlaceholders}
            clauseGroups={clauseGroups}
            unfilledCount={unfilledCount}
            fieldValues={fieldValues}
            fieldInputTypes={fieldInputTypes}
            onOpenSmartVariablesModal={onOpenSmartVariablesModal}
            onUpdateFieldValue={onUpdateFieldValue}
            onChangeFieldInputType={onChangeFieldInputType}
            onAddCustomField={onAddCustomField}
            onInsertPlaceholderAtCaret={onInsertPlaceholderAtCaret}
            onBakeAllPlaceholdersIntoDocument={onBakeAllPlaceholdersIntoDocument}
            estates={estates}
            selectedEstateId={selectedEstateId}
            selectedLotNumber={selectedLotNumber}
            onSelectEstateAndLot={onSelectEstateAndLot}
            onSaveEstate={onSaveEstate}
            onDeleteEstate={onDeleteEstate}
            onInsertLotClauseAtCaret={onInsertLotClauseAtCaret}
            onInsertFullSubdivisionTableAtCaret={onInsertFullSubdivisionTableAtCaret}
            savedParties={savedParties}
            savedProperties={savedProperties}
            onRecallPartyToRole={onRecallPartyToRole}
            onSaveCurrentPartyToDirectory={onSaveCurrentPartyToDirectory}
            onRecallPropertyRecord={onRecallPropertyRecord}
            onSaveCurrentPropertyToDirectory={onSaveCurrentPropertyToDirectory}
            onSaveSelectionBookmark={onSaveSelectionBookmark}
          />
        )}

        {activeTab === 'templates' && (
          <TemplatesTabPanel
            templates={templates}
            onImportDocxAsTemplate={onImportDocxAsTemplate}
            onSaveCurrentAsTemplate={onSaveCurrentAsTemplate}
            onOpenSaveAsTemplateModal={onOpenSaveAsTemplateModal}
            onExportTemplatesJson={onExportTemplatesJson}
            onImportTemplatesJson={onImportTemplatesJson}
            onLoadTemplateFull={onLoadTemplateFull}
            onInsertTemplateAtCaret={onInsertTemplateAtCaret}
            onExportTemplateDocx={onExportTemplateDocx}
            onDeleteTemplate={onDeleteTemplate}
            onSaveSelectionBookmark={onSaveSelectionBookmark}
          />
        )}

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

        {activeTab === 'notary_hub' && (
          <NotaryHubPanel
            savedDocuments={documents}
            activeDocumentId={activeDocumentId}
            documentTitle={currentEditorTitle || activeDocument?.title || 'عقد توثيقي'}
            fieldValues={fieldValues}
            onUpdateFieldValue={onUpdateFieldValue}
            onSaveCurrentDocument={() => onSaveCurrentDocument()}
            onContractIndexedOrSynced={onContractIndexedOrSynced}
            showToast={showToast}
          />
        )}
      </div>
    </aside>
  );
}
