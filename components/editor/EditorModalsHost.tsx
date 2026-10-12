'use client';

import React, { useState } from 'react';
import {
  FilePlus2,
  FileText,
  FileUp,
  FolderArchive,
  ListChecks,
  ShieldCheck,
  UserCheck,
  X,
} from 'lucide-react';
import InlineVariablePopover, {
  InlineVariablePopoverState,
} from '@/components/InlineVariablePopover';
import { SidebarTab } from '@/components/SidebarWorkspace';
import {
  DocxImportPreviewModal,
  OnboardingTourModal,
  SaveAsTemplateModal,
} from '@/components/modals/DocumentWorkflowModals';
import {
  SnapshotsHistoryModal,
  VersionDiffModal,
} from '@/components/RevisionsDiffModal';
import {
  DocumentRevision,
  NotaryClerk,
  OfficeTemplate,
  OfficeTemplateKind,
  SavedDocument,
  VariableInputType,
} from '@/lib/types';
import { DocxImportResult } from '@/lib/docx-engine';

interface MultiSourceStartModalProps {
  isOpen: boolean;
  onClose: () => void;
  templates: OfficeTemplate[];
  documents: SavedDocument[];
  onStartBlank: (clearPreviousClauses?: boolean) => void;
  onSelectTemplate: (
    tpl: OfficeTemplate,
    mode: 'replace' | 'insert',
    clearPreviousClauses?: boolean
  ) => void;
  onSelectSavedDoc: (doc: SavedDocument) => void;
  onImportDocxFile: (file: File) => void;
}

export function MultiSourceStartModal({
  isOpen,
  onClose,
  templates,
  documents,
  onStartBlank,
  onSelectTemplate,
  onSelectSavedDoc,
  onImportDocxFile,
}: MultiSourceStartModalProps) {
  const [sourceTab, setSourceTab] = useState<'blank' | 'templates' | 'docs'>('blank');
  const [templateLoadMode, setTemplateLoadMode] = useState<'replace' | 'insert'>('replace');
  const [clearPreviousClauses, setClearPreviousClauses] = useState<boolean>(true);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-[#0F2744]/55 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-[#C4A46A]/45 shadow-2xl w-full max-w-2xl max-h-[85vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-[#F7F5F0] border-b border-[#C4A46A]/35 flex items-center justify-between gap-3">
          <div className="flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-3">
            <img
              src="/brand/logo-lockup.svg"
              alt="الموثق الرقمي — محرر العقود التوثيقية"
              className="h-9 w-auto select-none shrink-0"
            />
            <span className="hidden sm:inline-block h-4 w-px bg-[#C4A46A]/50" />
            <h2 className="text-xs sm:text-sm font-bold text-[#0F2744]">
              بدء تحرير عقد جديد / فتح قالب
            </h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-[#5C6B7A] hover:text-[#0F2744] rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="grid grid-cols-3 gap-1 p-2 bg-slate-100 border-b border-slate-200 text-xs font-medium">
          <button
            type="button"
            onClick={() => setSourceTab('blank')}
            className={`py-1.5 rounded ${
              sourceTab === 'blank' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
            }`}
          >
            صفحة فارغة / ملف .docx
          </button>
          <button
            type="button"
            onClick={() => setSourceTab('templates')}
            className={`py-1.5 rounded ${
              sourceTab === 'templates' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
            }`}
          >
            قوالب المكتب ({templates.length})
          </button>
          <button
            type="button"
            onClick={() => setSourceTab('docs')}
            className={`py-1.5 rounded ${
              sourceTab === 'docs' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
            }`}
          >
            العقود المحفوظة ({documents.length})
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-5">
          {sourceTab === 'blank' && (
            <div className="space-y-4">
              <label className="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 p-2.5 rounded border border-slate-200 cursor-pointer">
                <input
                  type="checkbox"
                  checked={clearPreviousClauses}
                  onChange={(e) => setClearPreviousClauses(e.target.checked)}
                  className="rounded border-slate-300 text-blue-900"
                />
                <span className="font-medium">
                  حذف البنود السابقة المفعّلة عند بدء ورقة عقد جديدة
                </span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <button
                  type="button"
                  onClick={() => {
                    onStartBlank(clearPreviousClauses);
                    onClose();
                  }}
                  className="p-5 border border-slate-200 rounded-lg hover:border-blue-800 hover:bg-blue-50/30 text-right transition-colors space-y-2"
                >
                  <FilePlus2 className="w-6 h-6 text-blue-900" />
                  <div className="text-xs font-bold text-slate-900">
                    ورقة عقد توثيقي فارغة (A4)
                  </div>
                  <p className="text-[11px] text-slate-500">
                    بدء عقد جديد بخط Arial 13pt وتباعد 1.0 والهوامش التوثيقية الثابتة (7/2/1/6 سم).
                  </p>
                </button>

                <label className="p-5 border border-slate-200 rounded-lg hover:border-blue-800 hover:bg-blue-50/30 text-right transition-colors space-y-2 cursor-pointer block">
                  <FileUp className="w-6 h-6 text-blue-900" />
                  <div className="text-xs font-bold text-slate-900">
                    استيراد ملف Word (.docx) من الجهاز
                  </div>
                  <p className="text-[11px] text-slate-500">
                    فتح أي عقد وورد موجود على جهازك وتحويله فوراً لمعيار المكتب مع استخراج الوسوم.
                  </p>
                  <input
                    type="file"
                    accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    className="hidden"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      if (f) {
                        onImportDocxFile(f);
                        onClose();
                      }
                    }}
                  />
                </label>
              </div>
            </div>
          )}

          {sourceTab === 'templates' && (
            <div className="space-y-3">
              <div className="p-3 bg-blue-50/60 border border-blue-200 rounded-md space-y-2">
                <div className="text-xs font-bold text-blue-950">
                  خيار فتح القالب / النموذج:
                </div>
                <div className="flex flex-wrap items-center gap-4 text-xs">
                  <label className="inline-flex items-center gap-1.5 cursor-pointer font-medium text-slate-800">
                    <input
                      type="radio"
                      name="tplLoadMode"
                      checked={templateLoadMode === 'replace'}
                      onChange={() => setTemplateLoadMode('replace')}
                    />
                    <span>استبدال المحتوى الحالي (يمسح ورقة الـ A4 ويضع القالب كاملاً)</span>
                  </label>
                  <label className="inline-flex items-center gap-1.5 cursor-pointer font-medium text-slate-800">
                    <input
                      type="radio"
                      name="tplLoadMode"
                      checked={templateLoadMode === 'insert'}
                      onChange={() => setTemplateLoadMode('insert')}
                    />
                    <span>إدراج عند موضع المؤشر (يُبقي النص الحالي)</span>
                  </label>
                </div>

                {templateLoadMode === 'replace' && (
                  <label className="inline-flex items-center gap-2 text-xs text-slate-700 pt-1 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={clearPreviousClauses}
                      onChange={(e) => setClearPreviousClauses(e.target.checked)}
                      className="rounded border-slate-300 text-blue-900"
                    />
                    <span className="font-semibold text-blue-950">
                      حذف البنود السابقة (تفريغ قائمة البنود المفعّلة عند الاستبدال الكامل)
                    </span>
                  </label>
                )}
              </div>

              {templates.length === 0 ? (
                <div className="text-center py-8 text-xs text-slate-500">
                  لا توجد قوالب محفوظة في مكتبة المكتب بعد.
                </div>
              ) : (
                templates.map((tpl) => (
                  <div
                    key={tpl.id}
                    className="flex items-center justify-between p-3 border border-slate-200 rounded hover:bg-slate-50"
                  >
                    <div>
                      <div className="text-xs font-bold text-slate-900">{tpl.name}</div>
                      <div className="text-[11px] text-slate-500">
                        {tpl.category} · {(tpl.extractedPlaceholders || []).length} وسم ذكي
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        onSelectTemplate(tpl, templateLoadMode, clearPreviousClauses);
                        onClose();
                      }}
                      className="px-3 py-1.5 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800"
                    >
                      {templateLoadMode === 'replace'
                        ? 'استبدال وفتح القالب'
                        : 'إدراج القالب عند المؤشر'}
                    </button>
                  </div>
                ))
              )}
            </div>
          )}

          {sourceTab === 'docs' && (
            <div className="space-y-2">
              {documents.length === 0 ? (
                <div className="text-center py-8 text-xs text-slate-500">
                  لا توجد مسودات عقود محفوظة بعد.
                </div>
              ) : (
                documents.map((doc) => (
                  <div
                    key={doc.id}
                    className="flex items-center justify-between p-3 border border-slate-200 rounded hover:bg-slate-50"
                  >
                    <div>
                      <div className="text-xs font-bold text-slate-900">{doc.title}</div>
                      <div className="text-[11px] text-slate-500 tabular-nums">
                        {new Date(doc.updatedAt).toLocaleString('ar-DZ')}
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        onSelectSavedDoc(doc);
                        onClose();
                      }}
                      className="px-3 py-1 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800"
                    >
                      فتح العقد
                    </button>
                  </div>
                ))
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

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
  onOpenSmartVarsModalForKey: (varKey: string) => void;
  extractedPlaceholders: string[];
  fieldValues: Record<string, string>;
  fieldInputTypes: Record<string, VariableInputType>;
  onUpdateFieldValue: (key: string, value: string) => void;
  onChangeFieldInputType: (key: string, inputType: 'text' | 'number' | 'date') => void;
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
  templates: OfficeTemplate[];
  documents: SavedDocument[];
  onStartBlank: (clearPreviousClauses?: boolean) => void;
  onSelectTemplateFromMultiSource: (
    tpl: OfficeTemplate,
    mode: 'replace' | 'insert',
    clearPreviousClauses?: boolean
  ) => void;
  onSelectSavedDocFromMultiSource: (doc: SavedDocument) => void;
  onImportDocxFileFromMultiSource: (file: File) => Promise<void>;
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

