'use client';

import React from 'react';
import {
  Download,
  FileEdit,
  FilePlus2,
  FileUp,
  History,
  Printer,
  Settings,
} from 'lucide-react';
import type { LocalOfficeDirectoryState } from '@/lib/desktop-bridge';

import ExportMenuDropdown from './editor/ExportMenuDropdown';
import { StandardNotaryTemplateId } from '@/lib/notary-word-templates';

interface EditorTopBarProps {
  docTitle: string;
  onChangeDocTitle: (title: string) => void;
  autoSaveState: 'saved' | 'saving' | 'error';
  previewMergedMode: boolean;
  revisionsCount: number;
  localDirState?: LocalOfficeDirectoryState;
  onOpenNewContractModal: () => void;
  onImportWordFile: (file: File) => void;
  onOpenSnapshotsHistoryModal: () => void;
  onPrint: () => void;
  onExportCurrentToWord: () => void;
  onExportWithTemplate?: (templateId: StandardNotaryTemplateId) => void;
  onOpenTemplatesManager?: () => void;
  onOpenOfficeSettings: () => void;
}

export default function EditorTopBar({
  docTitle,
  onChangeDocTitle,
  autoSaveState,
  revisionsCount,
  localDirState,
  onOpenNewContractModal,
  onImportWordFile,
  onOpenSnapshotsHistoryModal,
  onPrint,
  onExportCurrentToWord,
  onExportWithTemplate,
  onOpenTemplatesManager,
  onOpenOfficeSettings,
}: EditorTopBarProps) {
  const isDiskConnected =
    localDirState?.status === 'connected' ||
    localDirState?.status === 'native_desktop';

  return (
    <header className="flex flex-wrap sm:flex-nowrap items-center justify-between gap-2 px-2 sm:px-4 py-2 bg-white border-b border-[#C4A46A]/35 shrink-0 no-print select-none shadow-2xs">
      {/* Right: Brand Mark & Editable Contract Title Input with Auto-Save Badge */}
      <div className="flex items-center gap-2.5 shrink-0">
        <div className="flex items-center gap-2.5">
          <img
            src="/brand/mark.svg"
            alt="الموثق الرقمي"
            width={34}
            height={34}
            className="w-8.5 h-8.5 rounded-lg shadow-2xs shrink-0 select-none"
          />
          <div className="hidden sm:flex flex-col leading-tight">
            <span className="text-sm font-bold tracking-tight text-[#0F2744]">
              الموثق الرقمي
            </span>
            <span className="text-[10px] font-medium text-[#5C6B7A]">
              محرر العقود التوثيقية
            </span>
          </div>
        </div>

        <div className="h-6 w-px bg-slate-200 hidden sm:block" />

        <div className="flex items-center gap-1.5 bg-[#F1F5F9] hover:bg-slate-100 border border-slate-200 focus-within:border-[#1E3A8A] focus-within:bg-white rounded-lg px-2.5 py-1 transition-all">
          <FileEdit className="w-3.5 h-3.5 text-[#5C6B7A] shrink-0" />
          <input
            type="text"
            value={docTitle}
            onChange={(e) => onChangeDocTitle(e.target.value)}
            title="انقر لتعديل عنوان العقد مباشرة"
            className="bg-transparent text-xs font-bold text-[#0F2744] focus:outline-none w-36 sm:w-48 md:w-64 truncate"
            placeholder="عنوان العقد..."
          />
          <span className="text-[10px] shrink-0">
            {autoSaveState === 'saving' ? (
              <span className="text-amber-600 font-medium">جاري الحفظ...</span>
            ) : autoSaveState === 'error' ? (
              <span className="text-[#8C1D2C] font-bold">تعذر الحفظ !</span>
            ) : (
              <span className="text-emerald-700 font-medium">محفوظ ✓</span>
            )}
          </span>
        </div>
      </div>

      {/* Left: Streamlined Contract File Actions + Office Settings Button */}
      <div className="flex flex-wrap items-center gap-1.5 shrink-0">
        {/* New Contract */}
        <button
          type="button"
          onClick={onOpenNewContractModal}
          className="px-2.5 py-1.5 bg-white border border-slate-300 hover:bg-[#F7F5F0] text-[#0F2744] rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
          title="بدء عقد جديد من قالب أو مسودة أو ملف وورد"
        >
          <FilePlus2 className="w-3.5 h-3.5 text-[#1E3A8A]" />
          <span className="hidden md:inline">عقد جديد</span>
        </button>

        {/* Import Word (.docx) */}
        <label
          className="px-2.5 py-1.5 bg-white border border-slate-300 text-[#0F2744] hover:bg-[#F7F5F0] rounded-lg text-xs font-medium inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap shrink-0 transition-colors"
          title="فتح واستيراد ملف Word (.docx)"
        >
          <FileUp className="w-3.5 h-3.5 text-[#1E3A8A]" />
          <span className="hidden md:inline">فتح Word</span>
          <input
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            className="hidden"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) {
                onImportWordFile(file);
                e.target.value = '';
              }
            }}
          />
        </label>

        {/* Revisions / History */}
        <button
          type="button"
          onClick={onOpenSnapshotsHistoryModal}
          className="px-2.5 py-1.5 bg-white border border-slate-300 hover:bg-[#F7F5F0] text-[#0F2744] rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors relative"
          title="سجل مراجعات وتعديلات العقد"
        >
          <History className="w-3.5 h-3.5 text-[#1E3A8A]" />
          <span className="hidden md:inline">السجل</span>
          {revisionsCount > 0 && (
            <span className="absolute -top-1 -right-1 bg-[#1E3A8A] text-white text-[9px] w-4 h-4 rounded-full flex items-center justify-center font-bold">
              {revisionsCount}
            </span>
          )}
        </button>

        <div className="h-5 w-px bg-slate-200 mx-0.5 hidden sm:block" />

        {/* Print / Preview */}
        <button
          type="button"
          onClick={onPrint}
          className="px-2.5 py-1.5 bg-white border border-slate-300 hover:bg-[#F7F5F0] text-[#0F2744] rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
          title="معاينة الطباعة وطباعة الوثيقة"
        >
          <Printer className="w-3.5 h-3.5 text-[#1E3A8A]" />
          <span className="hidden md:inline">طباعة</span>
        </button>

        {/* Export Word Dropdown Menu (Standard Templates) */}
        <ExportMenuDropdown
          onExportDirectDocx={onExportCurrentToWord}
          onExportWithTemplate={(tplId) => {
            if (onExportWithTemplate) {
              onExportWithTemplate(tplId);
            } else {
              onExportCurrentToWord();
            }
          }}
          onOpenTemplatesManager={() => {
            if (onOpenTemplatesManager) {
              onOpenTemplatesManager();
            }
          }}
        />

        <div className="h-5 w-px bg-slate-200 mx-0.5 hidden sm:block" />

        {/* Office Settings Button */}
        <button
          type="button"
          onClick={onOpenOfficeSettings}
          className={`px-3 py-1.5 border rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 transition-colors ${
            isDiskConnected
              ? 'bg-emerald-50 border-emerald-300 text-emerald-950 hover:bg-emerald-100'
              : 'bg-slate-50 border-slate-300 text-[#0F2744] hover:bg-slate-100'
          }`}
          title="إعدادات المكتب والبرنامج، التخزين والنسخ الاحتياطي"
        >
          <Settings className="w-3.5 h-3.5 text-[#1E3A8A]" />
          <span>إعدادات المكتب</span>
          {isDiskConnected && (
            <span className="w-1.5 h-1.5 rounded-full bg-emerald-600 shrink-0" />
          )}
        </button>
      </div>
    </header>
  );
}
