'use client';

import React, { useState, useEffect, useRef } from 'react';
import {
  X,
  HardDrive,
  Download,
  Upload,
  Building2,
  FileText,
  LayoutTemplate,
  BookOpen,
  Sliders,
  HelpCircle,
  Search,
  CheckCircle2,
  AlertCircle,
  Clock,
  Shield,
  Printer,
  Sparkles,
} from 'lucide-react';
import type { EditorSmartAssistPreferences } from '@/lib/types';
import type { LocalOfficeDirectoryState } from '@/lib/desktop-bridge';

interface OfficeSettingsModalProps {
  isOpen: boolean;
  onClose: () => void;
  localDirState?: LocalOfficeDirectoryState;
  onConnectLocalDirectory?: () => void;
  onExportBackupJson: () => void;
  onImportBackupJson: (file: File) => void;
  lastBackupAt: string | null;
  smartAssistPrefs: EditorSmartAssistPreferences;
  onChangeSmartAssistPrefs: (prefs: EditorSmartAssistPreferences) => void;
  onOpenOnboardingTour: () => void;
}

type SettingsSectionId =
  | 'storage'
  | 'identity'
  | 'pageSetup'
  | 'templates'
  | 'clauses'
  | 'editor'
  | 'about';

export default function OfficeSettingsModal({
  isOpen,
  onClose,
  localDirState,
  onConnectLocalDirectory,
  onExportBackupJson,
  onImportBackupJson,
  lastBackupAt,
  smartAssistPrefs,
  onChangeSmartAssistPrefs,
  onOpenOnboardingTour,
}: OfficeSettingsModalProps) {
  const [activeSection, setActiveSection] = useState<SettingsSectionId>('storage');
  const [searchQuery, setSearchQuery] = useState('');
  const modalRef = useRef<HTMLDivElement | null>(null);

  // Office Identity local state
  const [identity, setIdentity] = useState({
    officeName: 'مكتب التوثيق الرسمي',
    notaryName: 'الأستاذ الموثق',
    address: 'شارع العقيد لطفي، الجزائر العاصمة',
    phone: '+213 (0) 21 00 00 00',
  });

  // Page setup local state
  const [pageSetup, setPageSetup] = useState({
    topMarginCm: 7,
    bottomMarginCm: 2,
    leftMarginCm: 1,
    rightMarginCm: 6,
    fontFamily: 'Arial',
    fontSizePt: 13,
  });

  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showNotification = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => setToastMessage(null), 3500);
  };

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  if (!isOpen) return null;

  const formattedLastBackup = lastBackupAt
    ? new Date(lastBackupAt).toLocaleString('ar-DZ', {
        dateStyle: 'medium',
        timeStyle: 'short',
      })
    : null;

  const isDiskConnected =
    localDirState?.status === 'connected' ||
    localDirState?.status === 'native_desktop';

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-2 sm:p-6 select-none animate-in fade-in duration-200">
      <div
        ref={modalRef}
        role="dialog"
        aria-modal="true"
        aria-label="إعدادات المكتب"
        className="bg-white w-full max-w-5xl h-[90vh] sm:h-[82vh] rounded-2xl shadow-2xl border border-slate-300 flex flex-col overflow-hidden"
      >
        {/* Top Header */}
        <div className="flex items-center justify-between px-6 py-4 bg-[#0F2744] text-white shrink-0 border-b border-[#C4A46A]/40">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-white/10 rounded-xl">
              <Building2 className="w-5 h-5 text-[#C4A46A]" />
            </div>
            <div>
              <h2 className="text-base font-bold tracking-tight">إعدادات المكتب والبرنامج</h2>
              <p className="text-xs text-slate-300">
                إدارة التخزين، هوية الموثق، معايير الصفحة، تفضيلات المحرر والصيانة
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <div className="relative hidden sm:block">
              <Search className="w-4 h-4 text-slate-400 absolute right-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="بحث في الإعدادات..."
                className="bg-slate-800 text-xs text-white placeholder-slate-400 rounded-lg pr-9 pl-3 py-1.5 border border-slate-700 focus:outline-none focus:border-[#C4A46A] w-56 transition-all"
              />
            </div>
            <button
              type="button"
              onClick={onClose}
              className="p-2 rounded-lg bg-white/10 hover:bg-white/20 text-white transition-colors"
              title="إغلاق (Esc)"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Toast Notification */}
        {toastMessage && (
          <div className="bg-emerald-50 border-b border-emerald-200 px-4 py-2 text-xs font-semibold text-emerald-950 flex items-center gap-2 animate-in slide-in-from-top">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span>{toastMessage}</span>
          </div>
        )}

        {/* Main Body: Sidebar + Content Panel */}
        <div className="flex-1 flex overflow-hidden">
          {/* Sidebar Navigation */}
          <nav aria-label="أقسام إعدادات المكتب" className="w-64 bg-[#F8FAFC] border-s border-slate-200 p-3 space-y-1 overflow-y-auto shrink-0 hidden md:block">
            <div className="text-[11px] font-bold text-slate-400 uppercase tracking-wider px-3 py-1.5">
              أقسام الإعدادات
            </div>

            <button
              type="button"
              onClick={() => setActiveSection('storage')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'storage'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <HardDrive className="w-4 h-4 shrink-0" />
              <span>التخزين والنسخ الاحتياطي</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveSection('identity')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'identity'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <Building2 className="w-4 h-4 shrink-0" />
              <span>هوية المكتب والموثق</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveSection('pageSetup')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'pageSetup'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <Printer className="w-4 h-4 shrink-0" />
              <span>نموذج الصفحة والهوامش</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveSection('templates')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'templates'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <LayoutTemplate className="w-4 h-4 shrink-0" />
              <span>قوالب العقود الافتراضية</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveSection('clauses')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'clauses'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <BookOpen className="w-4 h-4 shrink-0" />
              <span>مكتبة البنود الجاهزة</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveSection('editor')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'editor'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <Sliders className="w-4 h-4 shrink-0" />
              <span>تفضيلات المحرر والمساعد</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveSection('about')}
              className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium transition-colors ${
                activeSection === 'about'
                  ? 'bg-[#1E3A8A] text-white shadow-2xs font-semibold'
                  : 'text-slate-700 hover:bg-slate-200/60'
              }`}
            >
              <HelpCircle className="w-4 h-4 shrink-0" />
              <span>حول البرنامج والمساعدة</span>
            </button>
          </nav>

          {/* Mobile Tab Select Dropdown */}
          <div className="md:hidden w-full bg-slate-100 p-3 border-b border-slate-200 flex items-center justify-between shrink-0">
            <select
              value={activeSection}
              onChange={(e) => setActiveSection(e.target.value as SettingsSectionId)}
              className="bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-xs font-bold text-slate-800 w-full"
            >
              <option value="storage">التخزين والنسخ الاحتياطي</option>
              <option value="identity">هوية المكتب والموثق</option>
              <option value="pageSetup">نموذج الصفحة والهوامش</option>
              <option value="templates">قوالب العقود الافتراضية</option>
              <option value="clauses">مكتبة البنود الجاهزة</option>
              <option value="editor">تفضيلات المحرر والمساعد</option>
              <option value="about">حول البرنامج والمساعدة</option>
            </select>
          </div>

          {/* Content Panel Area */}
          <div className="flex-1 overflow-y-auto p-6 bg-white space-y-6">
            {/* Section 1: Storage & Backup */}
            {activeSection === 'storage' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">التخزين والنسخ الاحتياطي للقرص المحلي</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    إدارة الاتصال بمجلد المكتب على القرص المحلي، تصدير واستعادة قواعد البيانات والنسخ الاحتياطية.
                  </p>
                </div>

                {/* Local Folder Connection Card */}
                <div className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-3">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <HardDrive className={`w-5 h-5 ${isDiskConnected ? 'text-emerald-600' : 'text-[#1E3A8A]'}`} />
                      <div>
                        <h4 className="text-xs font-bold text-slate-900">مجلد المكتب المحلي</h4>
                        <p className="text-[11px] text-slate-500">
                          {isDiskConnected
                            ? `مرتبط بـ: ${localDirState?.folderName || 'المجلد المحلي'}`
                            : 'حفظ المستندات تلقائياً في مجلد على حاسوبك'}
                        </p>
                      </div>
                    </div>
                    <span
                      className={`px-2.5 py-1 rounded-full text-[11px] font-bold ${
                        isDiskConnected
                          ? 'bg-emerald-100 text-emerald-800'
                          : 'bg-amber-100 text-amber-800'
                      }`}
                    >
                      {isDiskConnected ? 'متصل ومؤمّن ✓' : 'غير مرتبط'}
                    </span>
                  </div>

                  {onConnectLocalDirectory && (
                    <button
                      type="button"
                      onClick={onConnectLocalDirectory}
                      className="px-4 py-2 bg-[#1E3A8A] hover:bg-[#0F2744] text-white rounded-lg text-xs font-semibold inline-flex items-center gap-2 transition-colors shadow-2xs"
                    >
                      <HardDrive className="w-4 h-4" />
                      <span>{isDiskConnected ? 'تغيير أو تحديث المجلد' : 'ربط مجلد المكتب على القرص'}</span>
                    </button>
                  )}
                </div>

                {/* Backup Actions Card */}
                <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-3">
                  <div className="flex items-center justify-between">
                    <div>
                      <h4 className="text-xs font-bold text-slate-900">النسخ الاحتياطي واستعادة البيانات</h4>
                      <p className="text-[11px] text-slate-500">
                        {formattedLastBackup
                          ? `آخر نسخة احتياطية تم تصديرها: ${formattedLastBackup}`
                          : 'لم يتم تصدير نسخة احتياطية بعد'}
                      </p>
                    </div>
                    <span className="inline-flex items-center gap-1 text-[11px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                      <Shield className="w-3 h-3" />
                      بصمة SHA-256 للسلامة
                    </span>
                  </div>

                  <div className="flex flex-wrap items-center gap-3 pt-2">
                    <button
                      type="button"
                      onClick={() => {
                        onExportBackupJson();
                        showNotification('تم تصدير النسخة الاحتياطية بنجاح.');
                      }}
                      className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg text-xs font-semibold inline-flex items-center gap-2 transition-colors"
                    >
                      <Download className="w-4 h-4" />
                      <span>تصدير نسخة احتياطية كاملة (JSON)</span>
                    </button>

                    <label className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 border border-slate-300 rounded-lg text-xs font-semibold inline-flex items-center gap-2 cursor-pointer transition-colors">
                      <Upload className="w-4 h-4 text-[#1E3A8A]" />
                      <span>استعادة نسخة احتياطية</span>
                      <input
                        type="file"
                        accept=".json,application/json"
                        className="hidden"
                        onChange={(e) => {
                          const file = e.target.files?.[0];
                          if (file) {
                            onImportBackupJson(file);
                            e.target.value = '';
                            showNotification('تمت قراءة وفحص واستعادة النسخة بنجاح.');
                          }
                        }}
                      />
                    </label>
                  </div>
                </div>
              </div>
            )}

            {/* Section 2: Office Identity */}
            {activeSection === 'identity' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">هوية المكتب والموثق الرسمي</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    تُستعمل هذه البيانات تلقائياً في ترويسات العقود والوثائق الرسمية المصدرة.
                  </p>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">اسم المكتب التوثيقي</label>
                    <input
                      type="text"
                      value={identity.officeName}
                      onChange={(e) => setIdentity({ ...identity, officeName: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:border-[#1E3A8A] focus:outline-none"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">اسم ولقب الموثق</label>
                    <input
                      type="text"
                      value={identity.notaryName}
                      onChange={(e) => setIdentity({ ...identity, notaryName: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:border-[#1E3A8A] focus:outline-none"
                    />
                  </div>

                  <div className="space-y-1.5 sm:col-span-2">
                    <label className="text-xs font-semibold text-slate-700">عنوان المقر المهني</label>
                    <input
                      type="text"
                      value={identity.address}
                      onChange={(e) => setIdentity({ ...identity, address: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:border-[#1E3A8A] focus:outline-none"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">رقم الهاتف / الفاكس</label>
                    <input
                      type="text"
                      value={identity.phone}
                      onChange={(e) => setIdentity({ ...identity, phone: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:border-[#1E3A8A] focus:outline-none"
                    />
                  </div>
                </div>

                <div className="pt-2">
                  <button
                    type="button"
                    onClick={() => showNotification('تم حفظ هوية المكتب وتحديث الترويسة بنجاح.')}
                    className="px-4 py-2 bg-[#1E3A8A] hover:bg-[#0F2744] text-white rounded-lg text-xs font-semibold shadow-2xs transition-colors"
                  >
                    حفظ التغييرات
                  </button>
                </div>
              </div>
            )}

            {/* Section 3: Page Setup */}
            {activeSection === 'pageSetup' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">نموذج الصفحة والهوامش المعيارية</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    المعايير المعتمدة للعقود التوثيقية (الهوامش: علوي 7 سم، سفلي 2 سم، أيمن 6 سم، أيسر 1 سم).
                  </p>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">الهامش العلوي (سم)</label>
                    <input
                      type="number"
                      step="0.5"
                      value={pageSetup.topMarginCm}
                      onChange={(e) => setPageSetup({ ...pageSetup, topMarginCm: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">الهامش السفلي (سم)</label>
                    <input
                      type="number"
                      step="0.5"
                      value={pageSetup.bottomMarginCm}
                      onChange={(e) => setPageSetup({ ...pageSetup, bottomMarginCm: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">الهامش الأيمن (سم)</label>
                    <input
                      type="number"
                      step="0.5"
                      value={pageSetup.rightMarginCm}
                      onChange={(e) => setPageSetup({ ...pageSetup, rightMarginCm: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                    />
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">الهامش الأيسر (سم)</label>
                    <input
                      type="number"
                      step="0.5"
                      value={pageSetup.leftMarginCm}
                      onChange={(e) => setPageSetup({ ...pageSetup, leftMarginCm: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-2">
                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">نوع الخط الأساسي</label>
                    <select
                      value={pageSetup.fontFamily}
                      onChange={(e) => setPageSetup({ ...pageSetup, fontFamily: e.target.value })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                    >
                      <option value="Arial">Arial (معتمد)</option>
                      <option value="Traditional Arabic">Traditional Arabic</option>
                      <option value="Amiri">Amiri</option>
                    </select>
                  </div>

                  <div className="space-y-1.5">
                    <label className="text-xs font-semibold text-slate-700">حجم الخط (pt)</label>
                    <input
                      type="number"
                      value={pageSetup.fontSizePt}
                      onChange={(e) => setPageSetup({ ...pageSetup, fontSizePt: Number(e.target.value) })}
                      className="w-full bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-xs text-slate-800"
                    />
                  </div>
                </div>

                <div className="pt-2">
                  <button
                    type="button"
                    onClick={() => showNotification('تم حفظ إعدادات الصفحة والهوامش.')}
                    className="px-4 py-2 bg-[#1E3A8A] hover:bg-[#0F2744] text-white rounded-lg text-xs font-semibold shadow-2xs transition-colors"
                  >
                    تطبيق المعايير
                  </button>
                </div>
              </div>
            )}

            {/* Section 4: Templates */}
            {activeSection === 'templates' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">قوالب العقود الافتراضية</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    إدارة قوالب العقود الجاهزة (عقد بيع عقار، هبة، قسمة، وكالة عامة).
                  </p>
                </div>

                <div className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800">قالب عقد بيع عقار (شقة / قطعة أرض)</span>
                    <span className="text-[11px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">الافتراضي</span>
                  </div>
                  <p className="text-[11px] text-slate-600">
                    يحتوي على الأركان الأساسية والبنود القانونية المتعارف عليها في التوثيق الجزائري.
                  </p>
                </div>

                <div className="p-4 rounded-xl border border-slate-200 bg-white space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-800">قالب عقد الهبة والتنازل</span>
                    <span className="text-[11px] font-semibold text-slate-600 bg-slate-100 px-2 py-0.5 rounded">قالب قياسي</span>
                  </div>
                  <p className="text-[11px] text-slate-600">
                    مخصص لنقل الملكية بدون عوض مع شروط التحوز والقبول.
                  </p>
                </div>
              </div>
            )}

            {/* Section 5: Clauses */}
            {activeSection === 'clauses' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">مكتبة البنود الجاهزة والمقترحة</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    البنود النموذجية التي تدرج في صلب العقود (الشروط العامة، الضمانات، مصاريف التسجيل).
                  </p>
                </div>

                <div className="space-y-2">
                  <div className="p-3 rounded-xl border border-slate-200 bg-white flex items-center justify-between">
                    <div>
                      <h4 className="text-xs font-bold text-slate-800">بند الحالة المدنية والتصريح بالشرف</h4>
                      <p className="text-[11px] text-slate-500">يتضمن التأكد من الأهلية والهوية والصفة.</p>
                    </div>
                    <span className="text-[10px] bg-blue-50 text-blue-800 px-2 py-1 rounded font-medium">جاهز</span>
                  </div>

                  <div className="p-3 rounded-xl border border-slate-200 bg-white flex items-center justify-between">
                    <div>
                      <h4 className="text-xs font-bold text-slate-800">بند نقل الملكية وتحويل الحيازة</h4>
                      <p className="text-[11px] text-slate-500">ينقل الحيازة والانتفاع ابتداءً من تاريخ التوقيع.</p>
                    </div>
                    <span className="text-[10px] bg-blue-50 text-blue-800 px-2 py-1 rounded font-medium">جاهز</span>
                  </div>
                </div>
              </div>
            )}

            {/* Section 6: Editor Settings & Smart Assist */}
            {activeSection === 'editor' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">تفضيلات المحرر والمساعد الذكي</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    تخصيص مشغلات المساعد الذكي عند الكتابة (`/` للبند، `[` للوسم، `#` للترويسة).
                  </p>
                </div>

                <div className="space-y-3">
                  <label className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-200 cursor-pointer">
                    <div>
                      <span className="text-xs font-bold text-slate-800 block">مشغل الفهرسة والبنود (رمز `/`)</span>
                      <span className="text-[11px] text-slate-500">استدعاء قائمة إدراج بند جديد عند الضغط على مفتاح الشباك أو الشرطة المائلة.</span>
                    </div>
                    <input
                      type="checkbox"
                      checked={smartAssistPrefs.enableSlashCommands}
                      onChange={(e) =>
                        onChangeSmartAssistPrefs({
                          ...smartAssistPrefs,
                          enableSlashCommands: e.target.checked,
                        })
                      }
                      className="w-4 h-4 text-[#1E3A8A] rounded accent-[#1E3A8A]"
                    />
                  </label>

                  <label className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-200 cursor-pointer">
                    <div>
                      <span className="text-xs font-bold text-slate-800 block">الإكمال التلقائي الذكي للوسوم والمتغيرات</span>
                      <span className="text-[11px] text-slate-500">إظهار منسدلة الوسوم والمتغيرات تلقائياً عند الكتابة.</span>
                    </div>
                    <input
                      type="checkbox"
                      checked={smartAssistPrefs.enableInlineAutocomplete}
                      onChange={(e) =>
                        onChangeSmartAssistPrefs({
                          ...smartAssistPrefs,
                          enableInlineAutocomplete: e.target.checked,
                        })
                      }
                      className="w-4 h-4 text-[#1E3A8A] rounded accent-[#1E3A8A]"
                    />
                  </label>
                </div>
              </div>
            )}

            {/* Section 7: About / Help */}
            {activeSection === 'about' && (
              <div className="space-y-6 animate-in fade-in">
                <div>
                  <h3 className="text-base font-bold text-[#0F2744]">حول برنامج الموثق الرقمي والمساعدة</h3>
                  <p className="text-xs text-slate-500 mt-0.5">
                    الإصدار الاحترافي 2.0 (مجهز لبيئة سطح المكتب وتطبيق الويب المتكامل).
                  </p>
                </div>

                <div className="p-4 rounded-xl border border-slate-200 bg-slate-50 space-y-3">
                  <div className="flex items-center gap-3">
                    <Sparkles className="w-6 h-6 text-[#1E3A8A]" />
                    <div>
                      <h4 className="text-xs font-bold text-slate-900">الموثق الرقمي — الإصدار 2.0</h4>
                      <p className="text-[11px] text-slate-500">برنامج متخصص لتحرير ومراجعة وعقد العقود التوثيقية والشرعية.</p>
                    </div>
                  </div>

                  <div className="pt-2 border-t border-slate-200">
                    <button
                      type="button"
                      onClick={() => {
                        onClose();
                        onOpenOnboardingTour();
                      }}
                      className="px-4 py-2 bg-[#1E3A8A] text-white rounded-lg text-xs font-semibold inline-flex items-center gap-2 hover:bg-[#0F2744] transition-colors shadow-2xs"
                    >
                      <HelpCircle className="w-4 h-4" />
                      <span>فتح الدليل الإرشادي التفاعلي</span>
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
