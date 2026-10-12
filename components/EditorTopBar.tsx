'use client';

import React, { useState, useRef, useEffect } from 'react';
import {
  ChevronDown,
  Clock,
  Download,
  Eye,
  FileEdit,
  FilePlus2,
  FileUp,
  HardDrive,
  HelpCircle,
  History,
  Printer,
  Settings,
  Upload,
} from 'lucide-react';
import type { LocalOfficeDirectoryState } from '@/lib/desktop-bridge';

interface EditorTopBarProps {
  docTitle: string;
  onChangeDocTitle: (title: string) => void;
  autoSaveState: 'saved' | 'saving' | 'error';
  previewMergedMode: boolean;
  revisionsCount: number;
  lastBackupAt: string | null;
  localDirState?: LocalOfficeDirectoryState;
  onConnectLocalDirectory?: () => void;
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
  localDirState,
  onConnectLocalDirectory,
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
  const [showMaintenanceMenu, setShowMaintenanceMenu] = useState(false);
  const menuRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!showMaintenanceMenu) return;
    const handleClickOutside = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        setShowMaintenanceMenu(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [showMaintenanceMenu]);

  const formattedLastBackup = lastBackupAt
    ? new Date(lastBackupAt).toLocaleString('ar-DZ', {
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
      })
    : null;

  const isDiskConnected =
    localDirState?.status === 'connected' ||
    localDirState?.status === 'native_desktop';

  return (
    <header className="flex flex-wrap sm:flex-nowrap items-center justify-between gap-2 px-2 sm:px-4 py-2 bg-white border-b border-[#C4A46A]/35 shrink-0 no-print select-none shadow-2xs">
      {/* Right: Brand Mark (/brand/mark.svg) & Editable Contract Title Input with Auto-Save Badge */}
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

      {/* Left: Streamlined File Actions + Office Maintenance Dropdown */}
      <div className="flex flex-wrap items-center gap-1.5 shrink-0">
        {/* New Contract Modal */}
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

        {/* Unified Edit / Print Preview Toggle */}
        <button
          type="button"
          onClick={onTogglePreviewMergedMode}
          className={`px-2.5 py-1.5 border rounded-lg text-xs font-medium inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 transition-colors ${
            previewMergedMode
              ? 'bg-[#0F2744] border-[#0F2744] text-white font-bold'
              : 'bg-white border-slate-300 text-[#0F2744] hover:bg-[#F7F5F0]'
          }`}
          title="التبديل بين وضع التحرير الحي ومعاينة الطباعة النهائية (دمج الوسوم وترقيم ../..)"
        >
          <Eye className="w-3.5 h-3.5" />
          <span className="hidden sm:inline">
            {previewMergedMode ? 'العودة للتحرير' : 'معاينة الطباعة'}
          </span>
        </button>

        {/* Unified Revisions & Exports History Modal */}
        <button
          type="button"
          onClick={onOpenSnapshotsHistoryModal}
          className="px-2.5 py-1.5 bg-white border border-slate-300 text-[#0F2744] hover:bg-[#F7F5F0] rounded-lg text-xs font-medium inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 transition-colors"
          title="سجل مراجعات العقد واللقطات والتصديرات"
        >
          <History className="w-3.5 h-3.5 text-[#1E3A8A]" />
          <span className="hidden lg:inline tabular-nums">
            السجل ({revisionsCount})
          </span>
        </button>

        {/* Print */}
        <button
          type="button"
          onClick={onPrint}
          className="px-2.5 py-1.5 bg-white border border-slate-300 text-[#0F2744] hover:bg-[#F7F5F0] rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
          title="طباعة العقد A4"
        >
          <Printer className="w-3.5 h-3.5 text-[#0F2744]" />
          <span className="hidden xl:inline">طباعة</span>
        </button>

        {/* Export Word (.docx) - Primary CTA */}
        <button
          type="button"
          onClick={onExportCurrentToWord}
          className="px-3 py-1.5 text-xs font-semibold text-white bg-[#1E3A8A] rounded-lg hover:bg-[#0F2744] transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 shadow-2xs"
          title="تصدير العقد بصيغة Word (.docx) وتسجيله في سجل العقد"
        >
          <Download className="w-3.5 h-3.5" />
          <span>تصدير Word</span>
        </button>

        <div className="h-5 w-px bg-slate-200 mx-0.5 hidden sm:block" />

        {/* Office Maintenance Dropdown (صيانة المكتب) */}
        <div className="relative" ref={menuRef}>
          <button
            type="button"
            onClick={() => setShowMaintenanceMenu((prev) => !prev)}
            className={`px-2.5 py-1.5 border rounded-lg text-xs font-semibold inline-flex items-center gap-1.5 transition-colors ${
              showMaintenanceMenu
                ? 'bg-slate-900 border-slate-900 text-white'
                : isDiskConnected
                ? 'bg-emerald-50 border-emerald-300 text-emerald-950 hover:bg-emerald-100'
                : 'bg-slate-50 border-slate-300 text-[#0F2744] hover:bg-slate-100'
            }`}
            title="صيانة المكتب: ربط المجلد المحلي على القرص، تصدير واستعادة النسخة الاحتياطية"
          >
            <Settings className="w-3.5 h-3.5" />
            <span className="hidden md:inline">صيانة المكتب</span>
            {isDiskConnected && (
              <span className="w-1.5 h-1.5 rounded-full bg-emerald-600 shrink-0" />
            )}
            <ChevronDown className="w-3 h-3" />
          </button>

          {showMaintenanceMenu && (
            <div className="absolute left-0 mt-1.5 w-80 bg-white border border-slate-200 rounded-lg shadow-xl p-3 z-50 space-y-2.5 text-xs">
              <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                <span className="font-bold text-[#0F2744]">
                  صيانة المكتب والتخزين المحلي
                </span>
                <span
                  className={`inline-flex items-center gap-1 text-[10px] px-1.5 py-0.5 rounded tabular-nums ${
                    formattedLastBackup
                      ? 'text-emerald-800 bg-emerald-50 border border-emerald-200'
                      : 'text-amber-800 bg-amber-50 border border-amber-200'
                  }`}
                >
                  <Clock className="w-2.5 h-2.5" />
                  {formattedLastBackup
                    ? `آخر نسخ: ${formattedLastBackup}`
                    : 'لم يُنسخ بعد'}
                </span>
              </div>

              {onConnectLocalDirectory && (
                <button
                  type="button"
                  onClick={() => {
                    onConnectLocalDirectory();
                    setShowMaintenanceMenu(false);
                  }}
                  className={`w-full px-2.5 py-2 rounded-md text-xs font-semibold flex items-center justify-between transition-colors border ${
                    isDiskConnected
                      ? 'bg-emerald-50 border-emerald-300 text-emerald-950 hover:bg-emerald-100'
                      : localDirState?.status === 'needs_permission'
                      ? 'bg-amber-50 border-amber-300 text-amber-950 hover:bg-amber-100'
                      : 'bg-slate-50 border-slate-200 text-[#0F2744] hover:bg-slate-100'
                  }`}
                >
                  <span className="inline-flex items-center gap-1.5">
                    <HardDrive className="w-3.5 h-3.5 text-[#1E3A8A]" />
                    <span>
                      {isDiskConnected
                        ? `مجلد المكتب: ${localDirState?.folderName || 'متصل'}`
                        : localDirState?.status === 'needs_permission'
                        ? `تفعيل صلاحية: ${localDirState.folderName}`
                        : 'ربط مجلد المكتب على القرص المحلي'}
                    </span>
                  </span>
                  <span className="text-[10px] font-bold">
                    {isDiskConnected ? 'متصل ✓' : 'ربط'}
                  </span>
                </button>
              )}

              <div className="grid grid-cols-2 gap-1.5 pt-1">
                <button
                  type="button"
                  onClick={() => {
                    onExportBackupJson();
                    setShowMaintenanceMenu(false);
                  }}
                  className="py-2 px-2.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-md text-xs font-semibold inline-flex items-center justify-center gap-1.5 transition-colors"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>تصدير نسخة المكتب</span>
                </button>

                <label className="py-2 px-2.5 bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 rounded-md text-xs font-semibold inline-flex items-center justify-center gap-1.5 cursor-pointer transition-colors">
                  <Upload className="w-3.5 h-3.5 text-blue-900" />
                  <span>استعادة نسخة</span>
                  <input
                    type="file"
                    accept=".json,application/json"
                    className="hidden"
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (file) {
                        onImportBackupJson(file);
                        e.target.value = '';
                        setShowMaintenanceMenu(false);
                      }
                    }}
                  />
                </label>
              </div>

              <div className="pt-2 border-t border-slate-100 flex items-center justify-between">
                <button
                  type="button"
                  onClick={() => {
                    onOpenOnboardingTour();
                    setShowMaintenanceMenu(false);
                  }}
                  className="w-full py-1.5 px-2 text-[11px] font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-50 rounded flex items-center justify-center gap-1.5 transition-colors"
                >
                  <HelpCircle className="w-3.5 h-3.5 text-blue-900" />
                  <span>فتح الدليل الإرشادي للمحرر</span>
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
    </header>
  );
}
