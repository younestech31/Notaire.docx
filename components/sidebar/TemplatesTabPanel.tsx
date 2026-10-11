'use client';

import React, { useState } from 'react';
import { Download, FileText, Trash2, Upload } from 'lucide-react';
import { CustomTemplate } from '@/lib/types';

interface TemplatesTabPanelProps {
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
  onSaveSelectionBookmark: () => void;
}

export default function TemplatesTabPanel({
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
  onSaveSelectionBookmark,
}: TemplatesTabPanelProps) {
  const [newTplName, setNewTplName] = useState('');
  const [newTplCategory, setNewTplCategory] = useState('عقود المكتب');
  const [clearClausesOnReplace, setClearClausesOnReplace] = useState<boolean>(true);

  return (
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
                    {tpl.category} · <span className="tabular-nums">{tpl.extractedPlaceholders.length}</span> وسم ذكي
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
  );
}
