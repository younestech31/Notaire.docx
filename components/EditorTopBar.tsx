'use client';

import React from 'react';
import {
  Clock,
  Download,
  Eye,
  FileEdit,
  FilePlus2,
  FileUp,
  HelpCircle,
  History,
  Printer,
  Upload,
} from 'lucide-react';

interface EditorTopBarProps {
  docTitle: string;
  onChangeDocTitle: (title: string) => void;
  autoSaveState: 'saved' | 'saving' | 'error';
  previewMergedMode: boolean;
  revisionsCount: number;
  lastBackupAt: string | null;
  onOpenNewContractModal: () => void;
  onImportWordFile: (file: File) => void;
  onTogglePreviewMergedMode: () => void;
  onOpenSnapshotsHistoryModal: () => void;
  onPrint: () => void;
  onExportCurrentToWord: () => void;
  onExportBackupJson: () => void;
  onImportBackupJson: (file: File) => void;
  onOpenOnboardingTour: () => void;
}

export default function EditorTopBar({
  docTitle,
  onChangeDocTitle,
  autoSaveState,
  previewMergedMode,
  revisionsCount,
  lastBackupAt,
  onOpenNewContractModal,
  onImportWordFile,
  onTogglePreviewMergedMode,
  onOpenSnapshotsHistoryModal,
  onPrint,
  onExportCurrentToWord,
  onExportBackupJson,
  onImportBackupJson,
  onOpenOnboardingTour,
}: EditorTopBarProps) {
  const formattedLastBackup = lastBackupAt
    ? new Date(lastBackupAt).toLocaleString('ar-DZ', {
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
      })
    : null;

  return (
    <header className="flex flex-wrap items-center justify-between gap-2 px-3 sm:px-4 py-2 bg-white border-b border-slate-200 shrink-0 no-print select-none">
      {/* Right: Brand & Editable Contract Title Input with Auto-Save Badge */}
      <div className="flex items-center gap-2.5 shrink-0">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-blue-900 text-white flex items-center justify-center font-bold text-sm shadow-2xs">
            م
          </div>
          <span className="text-sm font-bold tracking-tight text-slate-900 hidden sm:inline">
            الموثق الرقمي
          </span>
        </div>

        <div className="h-5 w-px bg-slate-200 hidden sm:block" />

        <div className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100/90 border border-slate-200 focus-within:border-blue-900 focus-within:bg-white rounded-lg px-2.5 py-1 transition-all">
          <FileEdit className="w-3.5 h-3.5 text-slate-400 shrink-0" />
          <input
            type="text"
            value={docTitle}
            onChange={(e) => onChangeDocTitle(e.target.value)}
            title="انقر لتعديل عنوان العقد مباشرة"
            className="bg-transparent text-xs font-bold text-slate-900 focus:outline-none w-36 sm:w-48 md:w-60 truncate"
            placeholder="عنوان العقد..."
          />
          <span className="text-[10px] shrink-0">
            {autoSaveState === 'saving' ? (
              <span className="text-amber-600 font-medium">جاري الحفظ...</span>
            ) : autoSaveState === 'error' ? (
              <span className="text-red-600 font-bold">تعذر الحفظ !</span>
            ) : (
              <span className="text-emerald-600 font-medium">محفوظ ✓</span>
            )}
          </span>
        </div>
      </div>

      {/* Left: GROUP 1 — ملف (File & Office Backup Controls) */}
      <div className="flex flex-wrap items-center gap-1.5 shrink-0">
        <span className="text-[10px] font-bold text-slate-400 px-1.5 py-0.5 rounded bg-slate-100 border border-slate-200 select-none hidden xl:inline">
          ملف
        </span>

        {/* New Contract Modal */}
        <button
          type="button"
          onClick={onOpenNewContractModal}
          className="px-2.5 py-1.5 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
          title="بدء عقد جديد من قالب أو مسودة أو ملف وورد أو أرشيف التحميلات"
        >
          <FilePlus2 className="w-3.5 h-3.5 text-blue-900" />
          <span className="hidden md:inline">عقد جديد</span>
        </button>

        {/* Import Word (.docx) */}
        <label
          className="px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap shrink-0 transition-colors"
          title="فتح واستيراد ملف Word (.docx)"
        >
          <FileUp className="w-3.5 h-3.5 text-blue-900" />
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

        {/* Side-by-Side Merged Preview Toggle */}
        <button
          type="button"
          onClick={onTogglePreviewMergedMode}
          className={`px-2.5 py-1.5 border rounded-lg text-xs font-medium inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 transition-colors ${
            previewMergedMode
              ? 'bg-amber-100 border-amber-300 text-amber-950 font-bold'
              : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50'
          }`}
          title="معاينة دمج الحقول جنباً إلى جنب مع الاستمارة قبل التصدير"
        >
          <Eye className="w-3.5 h-3.5" />
          <span className="hidden sm:inline">
            {previewMergedMode ? 'وضع التحرير' : 'معاينة'}
          </span>
        </button>

        {/* Snapshots & Revisions History Modal */}
        <button
          type="button"
          onClick={onOpenSnapshotsHistoryModal}
          className="px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 transition-colors"
          title="سجل التعديلات واللقطات الزمنية والمقارنة الخاصة بهذا العقد"
        >
          <History className="w-3.5 h-3.5 text-blue-900" />
          <span className="hidden lg:inline">السجل ({revisionsCount})</span>
        </button>

        {/* Print */}
        <button
          type="button"
          onClick={onPrint}
          className="px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
          title="طباعة العقد A4"
        >
          <Printer className="w-3.5 h-3.5" />
          <span className="hidden xl:inline">طباعة</span>
        </button>

        {/* Export Word (.docx) - Primary CTA */}
        <button
          type="button"
          onClick={onExportCurrentToWord}
          className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-900 rounded-lg hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 shadow-2xs"
          title="تصدير العقد التوثيقي كاملاً إلى ملف Word (.docx)"
        >
          <Download className="w-3.5 h-3.5" />
          <span>تصدير Word</span>
        </button>

        <div className="h-5 w-px bg-slate-200 mx-0.5 hidden sm:block" />

        {/* OFFICE BACKUP SECTION (ضمن مجموعة ملف مع مؤشر آخر نسخ احتياطي) */}
        <div className="flex items-center gap-1 bg-slate-50 border border-slate-200 rounded-lg px-2 py-1">
          <button
            type="button"
            onClick={onExportBackupJson}
            className="px-2 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded text-[11px] font-semibold inline-flex items-center gap-1 transition-colors"
            title="تصدير نسخة احتياطية شاملة لكل عقود وقوالب وبنود المكتب (JSON)"
          >
            <Download className="w-3 h-3" />
            <span>تصدير نسخة المكتب</span>
          </button>

          <label
            className="px-2 py-1 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded text-[11px] font-medium inline-flex items-center gap-1 cursor-pointer transition-colors"
            title="استعادة نسخة المكتب الاحتياطية من ملف JSON"
          >
            <Upload className="w-3 h-3 text-blue-900" />
            <span className="hidden sm:inline">استعادة</span>
            <input
              type="file"
              accept=".json,application/json"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  onImportBackupJson(file);
                  e.target.value = '';
                }
              }}
            />
          </label>

          <span
            className={`hidden md:inline-flex items-center gap-1 text-[10px] px-1.5 py-0.5 rounded ${
              formattedLastBackup
                ? 'text-emerald-800 bg-emerald-50 border border-emerald-200'
                : 'text-amber-800 bg-amber-50 border border-amber-200'
            }`}
            title={
              formattedLastBackup
                ? `تاريخ آخر تصدير لنسخة المكتب: ${formattedLastBackup}`
                : 'البيانات محفوظة محلياً في المتصفح — يُنصح بتصدير نسخة المكتب دورياً'
            }
          >
            <Clock className="w-2.5 h-2.5" />
            {formattedLastBackup ? `آخر نسخ: ${formattedLastBackup}` : 'لم يُنسخ بعد'}
          </span>
        </div>

        {/* Editor Tour Guide */}
        <button
          type="button"
          onClick={onOpenOnboardingTour}
          className="p-1.5 text-slate-500 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-colors shrink-0"
          title="فتح الجولة الإرشادية التفاعلية للمحرر"
        >
          <HelpCircle className="w-4 h-4" />
        </button>
      </div>
    </header>
  );
}
