'use client';

import React from 'react';
import {
  FileText,
  FolderArchive,
  ListChecks,
  ShieldCheck,
  UserCheck,
} from 'lucide-react';
import InlineVariablePopover, {
  InlineVariablePopoverState,
} from '@/components/InlineVariablePopover';
import { SidebarTab } from '@/components/SidebarWorkspace';
import {
  DocxImportPreviewModal,
  MultiSourceStartModal,
  OnboardingTourModal,
  SaveAsTemplateModal,
  SnapshotsHistoryModal,
  VersionDiffModal,
} from '@/components/SmartModals';
import {
  ClauseVariableGroup,
  ContractOutlineClause,
  CustomTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  NotaryClerk,
  OfficeTemplate,
  OfficeTemplateKind,
  PartyField,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
  VariableInputType,
  WordTemplateDefinition,
} from '@/lib/types';
import { DocxImportResult } from '@/lib/docx-engine';

interface EditorActivityRailProps {
  unfilledCount: number;
  onSelectTab: (tab: SidebarTab) => void;
}

export function EditorActivityRail({
  unfilledCount,
  onSelectTab,
}: EditorActivityRailProps) {
  return (
    <aside className="flex w-14 bg-white border-l border-[#C4A46A]/35 flex-col items-center py-2.5 gap-2 h-full min-h-0 shrink-0 select-none no-print transition-all duration-150 shadow-2xs">
      <button
        type="button"
        onClick={() => onSelectTab('clauses')}
        className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-[#5C6B7A] hover:bg-[#F7F5F0] hover:text-[#0F2744]"
        title="فتح قسم البنود الجاهزة"
      >
        <ListChecks className="w-4 h-4" />
        <span className="text-[10px] font-medium">البنود</span>
      </button>

      <button
        type="button"
        onClick={() => onSelectTab('parties')}
        className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-[#5C6B7A] hover:bg-[#F7F5F0] hover:text-[#0F2744] relative"
        title="فتح قسم الأطراف والتعيينات"
      >
        <UserCheck className="w-4 h-4" />
        <span className="text-[10px] font-medium">الأطراف</span>
        {unfilledCount > 0 && (
          <span className="absolute top-1.5 left-1.5 w-2 h-2 rounded-full bg-[#8C1D2C]" />
        )}
      </button>

      <button
        type="button"
        onClick={() => onSelectTab('templates')}
        className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-[#5C6B7A] hover:bg-[#F7F5F0] hover:text-[#0F2744]"
        title="فتح قسم قوالب عقود المكتب"
      >
        <FileText className="w-4 h-4" />
        <span className="text-[10px] font-medium">القوالب</span>
      </button>

      <button
        type="button"
        onClick={() => onSelectTab('documents')}
        className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-[#5C6B7A] hover:bg-[#F7F5F0] hover:text-[#0F2744]"
        title="فتح قسم المستندات والأرشيف"
      >
        <FolderArchive className="w-4 h-4" />
        <span className="text-[10px] font-medium">الأرشيف</span>
      </button>

      <button
        type="button"
        onClick={() => onSelectTab('notary_hub')}
        className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-[#5C6B7A] hover:bg-[#F7F5F0] hover:text-[#0F2744]"
        title="الفهارس، المفتشية والشهر العقاري"
      >
        <ShieldCheck className="w-4 h-4" />
        <span className="text-[10px] font-medium">الفهارس</span>
      </button>
    </aside>
  );
}

interface EditorModalsHostProps {
  docId: string;
  docTitle: string;
  activeClerk?: NotaryClerk;
  inlinePopover: InlineVariablePopoverState | null;
  onCloseInlinePopover: () => void;
  showSmartVarsModal?: boolean;
  onCloseSmartVarsModal?: () => void;
  focusedVarKey?: string | null;
  onOpenSmartVarsModalForKey: (varKey: string) => void;
  extractedPlaceholders: string[];
  clauseGroups?: ClauseVariableGroup[];
  partyFields?: PartyField[];
  fieldValues: Record<string, string>;
  fieldInputTypes: Record<string, VariableInputType>;
  estates?: SubdivisionEstate[];
  selectedEstateId?: string;
  selectedLotNumber?: string;
  savedParties?: SavedPartyRecord[];
  savedProperties?: SavedPropertyRecord[];
  onSelectEstateAndLot?: (estateId: string, lotNumber: string) => void;
  onRecallPartyToRole?: (party: SavedPartyRecord, role: 'party1' | 'party2') => void;
  onSaveCurrentPartyToDirectory?: (role: 'party1' | 'party2') => void;
  onDeleteSavedParty?: (id: string) => void;
  onRecallPropertyRecord?: (prop: SavedPropertyRecord) => void;
  onSaveCurrentPropertyToDirectory?: () => void;
  onDeleteSavedProperty?: (id: string) => void;
  onUpdateFieldValue: (key: string, value: string) => void;
  onChangeFieldInputType: (key: string, inputType: 'text' | 'number' | 'date') => void;
  onBakeAllIntoDocument?: (explicitValues?: Record<string, string>) => void;
  showVersionDiffModal: boolean;
  onCloseVersionDiffModal: () => void;
  revisions: DocumentRevision[];
  diffComparisonRevision: DocumentRevision | null;
  modalActiveBodyHtml: string;
  onRestoreSnapshot: (rev: DocumentRevision) => void;
  showDocxPreviewModal: boolean;
  onCloseDocxPreviewModal: () => void;
  pendingDocxImport: DocxImportResult | null;
  onConfirmApplyDocxImport: (
    mode: 'replace' | 'insert',
    customSegmentedHtml?: string
  ) => void;
  showSaveAsTemplateModal: boolean;
  onCloseSaveAsTemplateModal: () => void;
  onConfirmSaveCustomTemplate: (data: {
    name: string;
    kind?: OfficeTemplateKind;
    code?: string;
    category: string;
    description: string;
    clearFilledValues: boolean;
  }) => Promise<void>;
  showSnapshotsHistoryModal: boolean;
  onCloseSnapshotsHistoryModal: () => void;
  onTakeManualSnapshot: (label: string) => Promise<void>;
  onDeleteSnapshot: (id: string) => Promise<void>;
  onOpenDiffComparison: (rev: DocumentRevision) => void;
  showOnboardingTour: boolean;
  onCloseOnboardingTour: () => void;
  onFinishOnboardingTour: () => void;
  showMultiSourceModal: boolean;
  onCloseMultiSourceModal: () => void;
  templates: (OfficeTemplate | CustomTemplate)[];
  documents: SavedDocument[];
  downloads?: DownloadArchiveItem[];
  onStartBlank: (clearPreviousClauses?: boolean) => void;
  onSelectTemplateFromMultiSource: (
    tpl: OfficeTemplate | CustomTemplate,
    mode: 'replace' | 'insert',
    clearPreviousClauses?: boolean
  ) => void;
  onSelectSavedDocFromMultiSource: (doc: SavedDocument) => void;
  onImportDocxFileFromMultiSource: (file: File) => Promise<void>;
  onSelectDownloadArchiveItemFromMultiSource?: (item: DownloadArchiveItem) => void;
  showWordTemplatesModal?: boolean;
  onCloseWordTemplatesModal?: () => void;
  wordTemplates?: WordTemplateDefinition[];
  outlineClauses?: ContractOutlineClause[];
  onSaveWordTemplate?: (tpl: WordTemplateDefinition) => Promise<void>;
  onDeleteWordTemplate?: (id: string) => Promise<void>;
}

export function EditorModalsHost({
  docId,
  docTitle,
  inlinePopover,
  onCloseInlinePopover,
  onOpenSmartVarsModalForKey,
  extractedPlaceholders,
  fieldValues,
  fieldInputTypes,
  onUpdateFieldValue,
  onChangeFieldInputType,
  showVersionDiffModal,
  onCloseVersionDiffModal,
  revisions,
  diffComparisonRevision,
  modalActiveBodyHtml,
  onRestoreSnapshot,
  showDocxPreviewModal,
  onCloseDocxPreviewModal,
  pendingDocxImport,
  onConfirmApplyDocxImport,
  showSaveAsTemplateModal,
  onCloseSaveAsTemplateModal,
  onConfirmSaveCustomTemplate,
  showSnapshotsHistoryModal,
  onCloseSnapshotsHistoryModal,
  onTakeManualSnapshot,
  onDeleteSnapshot,
  onOpenDiffComparison,
  showOnboardingTour,
  onCloseOnboardingTour,
  onFinishOnboardingTour,
  showMultiSourceModal,
  onCloseMultiSourceModal,
  templates,
  documents,
  onStartBlank,
  onSelectTemplateFromMultiSource,
  onSelectSavedDocFromMultiSource,
  onImportDocxFileFromMultiSource,
}: EditorModalsHostProps) {
  const contractRevs = revisions.filter((r) => r.documentId === docId);
  const diffRevs = diffComparisonRevision
    ? [
        diffComparisonRevision,
        ...contractRevs.filter((r) => r.id !== diffComparisonRevision.id && r.bodyHtml),
      ]
    : contractRevs.filter((r) => Boolean(r.bodyHtml && r.bodyHtml.trim()));

  return (
    <>
      <InlineVariablePopover
        popover={inlinePopover}
        value={inlinePopover ? fieldValues[inlinePopover.varKey] || '' : ''}
        inputType={
          inlinePopover ? fieldInputTypes[inlinePopover.varKey] || 'text' : 'text'
        }
        onChangeValue={onUpdateFieldValue}
        onChangeInputType={onChangeFieldInputType}
        onOpenFullModal={onOpenSmartVarsModalForKey}
        onClose={onCloseInlinePopover}
      />

      <VersionDiffModal
        isOpen={showVersionDiffModal}
        onClose={onCloseVersionDiffModal}
        revisions={diffRevs}
        currentBodyHtml={modalActiveBodyHtml}
        currentTitle={docTitle}
        onRestoreRevision={onRestoreSnapshot}
      />

      <DocxImportPreviewModal
        isOpen={showDocxPreviewModal}
        onClose={onCloseDocxPreviewModal}
        importResult={pendingDocxImport}
        onConfirmApply={onConfirmApplyDocxImport}
      />

      <SaveAsTemplateModal
        isOpen={showSaveAsTemplateModal}
        onClose={onCloseSaveAsTemplateModal}
        currentDocTitle={docTitle}
        extractedPlaceholders={extractedPlaceholders}
        currentFieldValues={fieldValues}
        onConfirmSave={onConfirmSaveCustomTemplate}
      />

      <SnapshotsHistoryModal
        isOpen={showSnapshotsHistoryModal}
        onClose={onCloseSnapshotsHistoryModal}
        revisions={contractRevs}
        currentBodyHtml={modalActiveBodyHtml}
        currentTitle={docTitle}
        onTakeManualSnapshot={onTakeManualSnapshot}
        onRestoreSnapshot={onRestoreSnapshot}
        onDeleteSnapshot={onDeleteSnapshot}
        onOpenDiffComparison={onOpenDiffComparison}
      />

      <OnboardingTourModal
        isOpen={showOnboardingTour}
        onClose={onCloseOnboardingTour}
        onFinishTour={onFinishOnboardingTour}
      />

      <MultiSourceStartModal
        isOpen={showMultiSourceModal}
        onClose={onCloseMultiSourceModal}
        templates={templates}
        documents={documents}
        onStartBlank={onStartBlank}
        onSelectTemplate={onSelectTemplateFromMultiSource}
        onSelectSavedDoc={onSelectSavedDocFromMultiSource}
        onImportDocxFile={onImportDocxFileFromMultiSource}
      />
    </>
  );
}
