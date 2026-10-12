'use client';

import React, { useState } from 'react';
import {
  Download,
  Edit3,
  FileOutput,
  FileText,
  Trash2,
  Upload,
} from 'lucide-react';
import {
  OfficeTemplate,
  OfficeTemplateKind,
} from '@/lib/types';

interface TemplatesTabPanelProps {
  templates: OfficeTemplate[];
  onImportDocxAsTemplate: (files: FileList) => Promise<void>;
  onSaveCurrentAsTemplate: (
    name: string,
    category: string,
    kind?: OfficeTemplateKind,
    code?: string
  ) => Promise<void>;
  onOpenSaveAsTemplateModal?: () => void;
  onExportTemplatesJson?: () => void;
  onImportTemplatesJson?: (file: File) => void;
  onLoadTemplateFull: (
    tpl: OfficeTemplate,
    clearPreviousClauses?: boolean
  ) => void;
  onInsertTemplateAtCaret: (tpl: OfficeTemplate) => void;
  onExportTemplateDocx: (tpl: OfficeTemplate) => Promise<void>;
  onGenerateDerivedFromTemplate?: (tpl: OfficeTemplate) => Promise<void>;
  onGenerateDerivedDoc?: (tpl: OfficeTemplate) => void | Promise<void>;
  onEditTemplateInEditor?: (tpl: OfficeTemplate) => void;
  activeEditingTemplateId?: string | null;
  onDeleteTemplate: (id: string) => Promise<void>;
  onSaveSelectionBookmark: () => void;
}

const KIND_LABELS: Record<OfficeTemplateKind, string> = {
  contract: 'عقد رسمي',
  extract: 'مستخرج / نسخة',
  registration: 'تسجيل وشهر',
  fiscal: 'وضعية جبائية',
  other: 'وثيقة مكتب',
};

export default function TemplatesTabPanel({
  templates = [],
  onImportDocxAsTemplate,
  onSaveCurrentAsTemplate,
  onOpenSaveAsTemplateModal,
  onExportTemplatesJson,
  onImportTemplatesJson,
  onLoadTemplateFull,
  onInsertTemplateAtCaret,
  onExportTemplateDocx,
  onGenerateDerivedFromTemplate,
  onGenerateDerivedDoc,
  onEditTemplateInEditor,
  activeEditingTemplateId,
  onDeleteTemplate,
  onSaveSelectionBookmark,
}: TemplatesTabPanelProps) {
  const [newTplName, setNewTplName] = useState('');
  const [newTplKind, setNewTplKind] = useState<OfficeTemplateKind>('contract');
  const [filterKind, setFilterKind] = useState<
    'all' | 'contract' | 'extract' | 'registration_fiscal'
  >('all');
  const [clearClausesOnReplace, setClearClausesOnReplace] =
    useState<boolean>(true);

  const unifiedList: OfficeTemplate[] = templates;

  const filteredList = unifiedList.filter((tpl) => {
    if (filterKind === 'all') return true;
    if (filterKind === 'contract') return tpl.kind === 'contract';
    if (filterKind === 'extract')
      return tpl.kind === 'extract' || tpl.kind === 'other';
    if (filterKind === 'registration_fiscal')
      return tpl.kind === 'registration' || tpl.kind === 'fiscal';
    return true;
  });

  return (
    <div className="space-y-4">
      {/* Import .docx as OfficeTemplate */}
      <div className="border border-blue-200 bg-blue-50/40 rounded-md p-3 space-y-2">
        <div className="text-xs font-bold text-slate-900">
          استيراد نماذج العقود والوثائق (.docx)
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

      {/* Save Current Document as Unified OfficeTemplate */}
      <div className="border border-slate-200 rounded-md p-3 space-y-2 bg-white">
        <div className="flex items-center justify-between">
          <span className="text-xs font-bold text-slate-900">
            حفظ المحرر الحالي كقالب مكتب موحّد
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
          placeholder="اسم القالب (مثال: عقد بيع شقة / مستخرج تسجيل)"
          className="w-full px-2.5 py-1.5 text-xs border border-slate-300 rounded"
        />
        <div className="flex items-center gap-1.5">
          <select
            value={newTplKind}
            onChange={(e) =>
              setNewTplKind(e.target.value as OfficeTemplateKind)
            }
            className="flex-1 px-2 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded"
          >
            <option value="contract">النوع: عقد رسمي</option>
            <option value="extract">النوع: مستخرج / نسخة</option>
            <option value="registration">النوع: استمارة شهر / تسجيل</option>
            <option value="fiscal">النوع: وضعية جبائية</option>
            <option value="other">النوع: وثيقة مكتب أخرى</option>
          </select>
          <button
            type="button"
            onClick={async () => {
              if (!newTplName.trim()) return;
              const categoryLabel = KIND_LABELS[newTplKind] || 'عقود رسمية';
              await onSaveCurrentAsTemplate(
                newTplName.trim(),
                categoryLabel,
                newTplKind
              );
              setNewTplName('');
            }}
            className="px-3 py-1.5 bg-slate-900 text-white text-xs font-medium rounded hover:bg-slate-800 shrink-0"
          >
            حفظ القالب
          </button>
        </div>
      </div>

      {/* Templates JSON Export / Import Strip */}
      <div className="flex items-center gap-2">
        {onExportTemplatesJson && (
          <button
            type="button"
            onClick={onExportTemplatesJson}
            disabled={unifiedList.length === 0}
            className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-[11px] font-medium rounded inline-flex items-center justify-center gap-1 disabled:opacity-40"
            title="تصدير جميع قوالب المكتب الموحّدة إلى ملف JSON"
          >
            <Download className="w-3 h-3 text-blue-900" />
            <span>تصدير القوالب (JSON)</span>
          </button>
        )}

        {onImportTemplatesJson && (
          <label
            className="flex-1 py-1 px-2 bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 text-[11px] font-medium rounded inline-flex items-center justify-center gap-1 cursor-pointer"
            title="استيراد قوالب من ملف JSON"
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

      {/* Unified Templates Catalog with Kind Filters */}
      <div className="space-y-2.5">
        <div className="flex items-center justify-between border-b border-slate-200 pb-1.5">
          <span className="text-xs font-bold text-slate-900">
            قوالب المكتب الموحّدة (
            <span className="tabular-nums">{filteredList.length}</span>)
          </span>
        </div>

        {/* Filter Bar */}
        <div className="grid grid-cols-2 gap-1 p-1 bg-slate-100 rounded text-[11px] font-semibold">
          <button
            type="button"
            onClick={() => setFilterKind('all')}
            className={`py-1 rounded transition-colors ${
              filterKind === 'all'
                ? 'bg-white text-slate-900 shadow-2xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            الكل ({unifiedList.length})
          </button>
          <button
            type="button"
            onClick={() => setFilterKind('contract')}
            className={`py-1 rounded transition-colors ${
              filterKind === 'contract'
                ? 'bg-white text-blue-950 shadow-2xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            عقود رسمية ({unifiedList.filter((t) => t.kind === 'contract').length})
          </button>
          <button
            type="button"
            onClick={() => setFilterKind('extract')}
            className={`py-1 rounded transition-colors ${
              filterKind === 'extract'
                ? 'bg-white text-emerald-950 shadow-2xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            مستخرجات ونسخ (
            {
              unifiedList.filter(
                (t) => t.kind === 'extract' || t.kind === 'other'
              ).length
            }
            )
          </button>
          <button
            type="button"
            onClick={() => setFilterKind('registration_fiscal')}
            className={`py-1 rounded transition-colors ${
              filterKind === 'registration_fiscal'
                ? 'bg-white text-amber-950 shadow-2xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            شهر وجبائية (
            {
              unifiedList.filter(
                (t) => t.kind === 'registration' || t.kind === 'fiscal'
              ).length
            }
            )
          </button>
        </div>

        <label className="flex items-center gap-2 p-2 bg-blue-50/60 border border-blue-200 rounded text-[11px] text-blue-950 cursor-pointer">
          <input
            type="checkbox"
            checked={clearClausesOnReplace}
            onChange={(e) => setClearClausesOnReplace(e.target.checked)}
            className="rounded border-slate-300 text-blue-900"
          />
          <span className="font-semibold">
            حذف البنود السابقة المفعّلة عند فتح القالب كعقد كامل
          </span>
        </label>

        {filteredList.length === 0 ? (
          <div className="border border-dashed border-slate-300 rounded-md p-5 text-center space-y-2 bg-slate-50/50">
            <FileText className="w-7 h-7 text-slate-400 mx-auto" />
            <div className="text-xs font-semibold text-slate-700">
              لا توجد قوالب مطابقة لهذا التصنيف
            </div>
            <p className="text-[11px] text-slate-500">
              استوردوا ملفات Word (.docx) أو احفظوا العقد الحالي كقالب مكتب.
            </p>
          </div>
        ) : (
          filteredList.map((tpl) => {
            const isDerivedKind = tpl.kind !== 'contract';
            return (
              <div
                key={tpl.id}
                className="border border-slate-200 rounded-md p-2.5 bg-white space-y-2"
              >
                <div className="flex items-start justify-between gap-2">
                  <div className="min-w-0">
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <span
                        className={`px-1.5 py-0.5 text-[10px] font-bold rounded ${
                          tpl.kind === 'contract'
                            ? 'bg-blue-100 text-blue-900'
                            : tpl.kind === 'registration'
                            ? 'bg-amber-100 text-amber-900'
                            : tpl.kind === 'fiscal'
                            ? 'bg-purple-100 text-purple-900'
                            : 'bg-emerald-100 text-emerald-900'
                        }`}
                      >
                        {KIND_LABELS[tpl.kind] || tpl.category}
                      </span>
                      <span className="text-xs font-bold text-slate-900 truncate">
                        {tpl.name}
                      </span>
                    </div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      {tpl.description || tpl.category} ·{' '}
                      <span className="tabular-nums font-mono">
                        {tpl.extractedPlaceholders.length}
                      </span>{' '}
                      وسم ذكي
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => onDeleteTemplate(tpl.id)}
                    className="p-1 text-slate-400 hover:text-red-600 shrink-0"
                    title="حذف هذا القالب من المكتب"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>

                {isDerivedKind ? (
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() =>
                        onGenerateDerivedDoc
                          ? onGenerateDerivedDoc(tpl)
                          : onGenerateDerivedFromTemplate
                          ? onGenerateDerivedFromTemplate(tpl)
                          : onExportTemplateDocx(tpl)
                      }
                      className="flex-1 py-1.5 px-2 bg-emerald-800 text-white text-[11px] font-semibold rounded hover:bg-emerald-900 inline-flex items-center justify-center gap-1"
                      title="توليد هذه الوثيقة المشتقة ببيانات العقد الحالي وتنزيلها بصيغة Word (.docx)"
                    >
                      <FileOutput className="w-3.5 h-3.5" />
                      <span>توليد وتصدير (.docx)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() =>
                        onEditTemplateInEditor
                          ? onEditTemplateInEditor(tpl)
                          : onLoadTemplateFull(tpl, clearClausesOnReplace)
                      }
                      className={`py-1.5 px-2.5 border text-[11px] font-medium rounded inline-flex items-center gap-1 ${
                        activeEditingTemplateId === tpl.id
                          ? 'bg-amber-900 text-white border-amber-900'
                          : 'bg-slate-100 text-slate-800 border-slate-200 hover:bg-slate-200'
                      }`}
                      title="فتح هذا القالب في المحرر لتعديله"
                    >
                      <Edit3 className="w-3 h-3" />
                      <span>
                        {activeEditingTemplateId === tpl.id
                          ? 'قيد التعديل'
                          : 'فتح بالمحرر'}
                      </span>
                    </button>
                  </div>
                ) : (
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() =>
                        onLoadTemplateFull(tpl, clearClausesOnReplace)
                      }
                      className="flex-1 py-1 px-2 bg-blue-900 text-white text-[11px] font-medium rounded hover:bg-blue-800"
                      title="يفتح هذا القالب كعقد كامل على ورقة الـ A4"
                    >
                      فتح كعقد كامل
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
                      title="تصدير القالب بصيغة Word (.docx)"
                    >
                      <Download className="w-3.5 h-3.5" />
                    </button>
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}
