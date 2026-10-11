'use client';

import React, { useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  BookUser,
  Check,
  ChevronDown,
  ChevronUp,
  Download,
  Eye,
  FileText,
  FileUp,
  Merge,
  Save,
  Scissors,
  Search,
  Sparkles,
  Tag,
  Trash2,
  X,
} from 'lucide-react';
import { DocxImportResult } from '@/lib/docx-engine';
import {
  CandidateImportedClause,
  compileCandidateClausesToHtml,
  segmentHtmlIntoCandidateClauses,
} from '@/lib/editor-utils';
import {
  exportTemplateAsDocx,
  renderTemplateWithContractData,
} from '@/lib/docx-template-engine';
import { WordTemplateDefinition } from '@/lib/types';

// =========================================================
// 1. WORD IMPORT PREVIEW MODAL (معاينة استيراد ملف Word)
// =========================================================
interface DocxImportPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  importResult: DocxImportResult | null;
  onConfirmApply: (mode: 'replace' | 'insert', compiledBodyHtml?: string) => void;
}

export function DocxImportPreviewModal({
  isOpen,
  onClose,
  importResult,
  onConfirmApply,
}: DocxImportPreviewModalProps) {
  const [importMode, setImportMode] = useState<'replace' | 'insert'>('replace');
  const [lastResultRef, setLastResultRef] = useState<DocxImportResult | null>(null);
  const [segmentIntoClauses, setSegmentIntoClauses] = useState<boolean>(true);
  const [candidateClauses, setCandidateClauses] = useState<CandidateImportedClause[]>([]);
  const [expandedClauseId, setExpandedClauseId] = useState<string | null>(null);

  if (isOpen && importResult && importResult !== lastResultRef) {
    const detected = segmentHtmlIntoCandidateClauses(importResult.bodyHtml);
    setLastResultRef(importResult);
    setCandidateClauses(detected);
    setSegmentIntoClauses(detected.length > 0);
    setExpandedClauseId(detected[0]?.id || null);
  }

  if (!isOpen || !importResult) return null;

  const plainSnippet = importResult.bodyHtml
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 480);

  const handleRenameClause = (clauseId: string, newTitle: string) => {
    setCandidateClauses((prev) =>
      prev.map((c) => (c.id === clauseId ? { ...c, title: newTitle } : c))
    );
  };

  const handleMergeWithPrevious = (idx: number) => {
    if (idx <= 0) return;
    setCandidateClauses((prev) => {
      const next = [...prev];
      const prevClause = next[idx - 1];
      const currClause = next[idx];
      const mergedParagraphs = [
        ...prevClause.paragraphsHtml,
        ...currClause.paragraphsHtml,
      ];
      const tmp = document.createElement('div');
      tmp.innerHTML = mergedParagraphs.join(' ');
      next[idx - 1] = {
        ...prevClause,
        paragraphsHtml: mergedParagraphs,
        previewText: (tmp.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120),
      };
      next.splice(idx, 1);
      return next;
    });
  };

  const handleSplitAtParagraph = (clauseIdx: number, paraIdx: number) => {
    if (paraIdx <= 0) return;
    setCandidateClauses((prev) => {
      const next = [...prev];
      const target = next[clauseIdx];
      if (!target || paraIdx >= target.paragraphsHtml.length) return prev;

      const firstParas = target.paragraphsHtml.slice(0, paraIdx);
      const secondParas = target.paragraphsHtml.slice(paraIdx);

      const tmp1 = document.createElement('div');
      tmp1.innerHTML = firstParas.join(' ');
      const tmp2 = document.createElement('div');
      tmp2.innerHTML = secondParas.join(' ');

      const firstParaText = (secondParas[0] || '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      const suggestedTitle =
        firstParaText.length >= 3 && firstParaText.length <= 65
          ? firstParaText.replace(/^#+\s*/, '').replace(/:$/, '').trim()
          : `بند مستخرج ${clauseIdx + 2}`;

      const updatedFirst: CandidateImportedClause = {
        ...target,
        paragraphsHtml: firstParas,
        previewText: (tmp1.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120),
      };

      const newSecond: CandidateImportedClause = {
        id: `imp_split_${Date.now()}_${clauseIdx}_${paraIdx}`,
        title: suggestedTitle,
        paragraphsHtml: secondParas,
        previewText: (tmp2.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120),
      };

      next.splice(clauseIdx, 1, updatedFirst, newSecond);
      setExpandedClauseId(newSecond.id);
      return next;
    });
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in-50 duration-200">
        {/* Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-blue-100 text-blue-900 rounded-lg">
              <FileUp className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                معاينة ملف Word (.docx) ومراجعة تقسيم البنود قبل التطبيق
              </h2>
              <div className="text-[11px] text-slate-500 font-mono">
                {importResult.fileName}
              </div>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* Quick Metrics Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
            <div className="p-2.5 rounded-lg border border-slate-200 bg-slate-50/60 text-center">
              <div className="text-[11px] text-slate-500 font-medium">الفقرات</div>
              <div className="text-base font-bold text-slate-900 tabular-nums">
                {importResult.stats.paragraphCount}
              </div>
            </div>
            <div className="p-2.5 rounded-lg border border-slate-200 bg-slate-50/60 text-center">
              <div className="text-[11px] text-slate-500 font-medium">الكلمات</div>
              <div className="text-base font-bold text-slate-900 tabular-nums">
                {importResult.stats.wordCount}
              </div>
            </div>
            <div className="p-2.5 rounded-lg border border-blue-200 bg-blue-50/50 text-center">
              <div className="text-[11px] text-blue-900 font-medium">البنود المكتشفة</div>
              <div className="text-base font-bold text-blue-900 tabular-nums">
                {candidateClauses.length}
              </div>
            </div>
            <div className="p-2.5 rounded-lg border border-pink-200 bg-pink-50/50 text-center">
              <div className="text-[11px] text-pink-900 font-medium">الوسوم الذكية</div>
              <div className="text-base font-bold text-pink-900 tabular-nums">
                {importResult.extractedPlaceholders.length}
              </div>
            </div>
          </div>

          {/* PHASE 2: INTERACTIVE CLAUSE SEGMENTATION REVIEW */}
          <div className="border border-blue-200 rounded-lg p-3.5 bg-blue-50/30 space-y-3">
            <div className="flex items-center justify-between gap-2 flex-wrap">
              <label className="flex items-center gap-2 cursor-pointer">
                <input
                  type="checkbox"
                  checked={segmentIntoClauses}
                  onChange={(e) => setSegmentIntoClauses(e.target.checked)}
                  className="rounded border-slate-300 text-blue-900"
                />
                <span className="text-xs font-bold text-slate-900">
                  هيكلة العقد المستورد إلى بنود حية قابلة للتفعيل والتعطيل والسحب ({candidateClauses.length} بند)
                </span>
              </label>
              <span className="text-[11px] text-blue-900">
                يمكنك تعديل عنوان أي بند أو دمجه أو فصله يدوياً قبل الإدراج
              </span>
            </div>

            {segmentIntoClauses && candidateClauses.length > 0 && (
              <div className="space-y-2 max-h-64 overflow-y-auto pr-1">
                {candidateClauses.map((clause, idx) => {
                  const isExpanded = expandedClauseId === clause.id;
                  return (
                    <div
                      key={clause.id}
                      className="bg-white border border-slate-200 rounded-md p-2.5 space-y-2"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <div className="flex items-center gap-1.5 flex-1">
                          <span className="px-1.5 py-0.5 text-[10px] font-mono font-bold bg-blue-900 text-white rounded shrink-0">
                            #{idx + 1}
                          </span>
                          <input
                            type="text"
                            value={clause.title}
                            onChange={(e) => handleRenameClause(clause.id, e.target.value)}
                            className="flex-1 px-2 py-1 text-xs font-bold text-slate-900 bg-slate-50 border border-slate-300 rounded focus:bg-white focus:border-blue-900 focus:outline-none"
                            placeholder="عنوان البند..."
                          />
                        </div>

                        <div className="flex items-center gap-1 shrink-0">
                          {idx > 0 && (
                            <button
                              type="button"
                              onClick={() => handleMergeWithPrevious(idx)}
                              className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-[10px] font-semibold rounded inline-flex items-center gap-1"
                              title="دمج هذا البند مع البند السابق إذا لم يكن بنداً مستقلاً"
                            >
                              <Merge className="w-3 h-3" />
                              <span>دمج مع السابق</span>
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() =>
                              setExpandedClauseId(isExpanded ? null : clause.id)
                            }
                            className="px-2 py-1 bg-blue-50 hover:bg-blue-100 text-blue-900 text-[10px] font-semibold rounded inline-flex items-center gap-1"
                          >
                            <span>فقرات ({clause.paragraphsHtml.length})</span>
                            {isExpanded ? (
                              <ChevronUp className="w-3 h-3" />
                            ) : (
                              <ChevronDown className="w-3 h-3" />
                            )}
                          </button>
                        </div>
                      </div>

                      {!isExpanded && clause.previewText && (
                        <p className="text-[11px] text-slate-500 line-clamp-1 pr-6">
                          {clause.previewText}
                        </p>
                      )}

                      {isExpanded && (
                        <div className="space-y-1.5 pt-1 border-t border-slate-100">
                          {clause.paragraphsHtml.map((pHtml, pIdx) => {
                            const pText = pHtml
                              .replace(/<[^>]+>/g, ' ')
                              .replace(/&nbsp;/g, ' ')
                              .replace(/\s+/g, ' ')
                              .trim();
                            return (
                              <div
                                key={pIdx}
                                className="flex items-start justify-between gap-2 p-1.5 rounded bg-slate-50 border border-slate-200/70 text-[11px]"
                              >
                                <span className="text-slate-700 line-clamp-2 flex-1">
                                  {pText || '(فقرة فارغة / فاصل)'}
                                </span>
                                {pIdx > 0 && (
                                  <button
                                    type="button"
                                    onClick={() => handleSplitAtParagraph(idx, pIdx)}
                                    className="px-2 py-0.5 bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-200 rounded text-[10px] font-semibold shrink-0 inline-flex items-center gap-1"
                                    title="فصل المستند ابتداءً من هذه الفقرة كبند جديد مستقل"
                                  >
                                    <Scissors className="w-2.5 h-2.5" />
                                    <span>فصل كبند جديد</span>
                                  </button>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Extracted Smart Tags */}
          {importResult.extractedPlaceholders.length > 0 && (
            <div className="border border-slate-200 rounded-lg p-3 bg-white space-y-1.5">
              <div className="text-xs font-bold text-slate-800 flex items-center justify-between">
                <span>الوسوم المكتشفة تلقائياً من المعقوفات `[...]`:</span>
                <span className="text-[11px] text-slate-500">
                  {importResult.extractedPlaceholders.length} وسم
                </span>
              </div>
              <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto pt-1">
                {importResult.extractedPlaceholders.map((k) => (
                  <span
                    key={k}
                    className="px-2 py-0.5 rounded text-[11px] font-mono font-medium bg-pink-50 text-pink-900 border border-pink-200"
                  >
                    {`[${k}]`}
                  </span>
                ))}
              </div>
            </div>
          )}

          {/* Content Snippet Preview */}
          {!segmentIntoClauses && (
            <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/50 space-y-1.5">
              <div className="text-xs font-bold text-slate-700">مقتطف من نص الوثيقة:</div>
              <p
                dir="rtl"
                className="text-xs text-slate-600 leading-relaxed font-arabic line-clamp-4 bg-white p-2.5 rounded border border-slate-200/80"
              >
                {plainSnippet || 'وثيقة وورد خالية من النصوص المباشرة.'}
                {plainSnippet.length >= 480 ? '...' : ''}
              </p>
            </div>
          )}

          {/* Import Destination Options */}
          <div className="border border-slate-200 rounded-lg p-3.5 bg-white space-y-2.5">
            <div className="text-xs font-bold text-slate-900">
              طريقة تطبيق المحتوى المستورد على ورقة المحرر:
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <label
                className={`flex items-start gap-2.5 p-2.5 rounded-lg border cursor-pointer transition-colors ${
                  importMode === 'replace'
                    ? 'border-blue-800 bg-blue-50/40'
                    : 'border-slate-200 hover:bg-slate-50'
                }`}
              >
                <input
                  type="radio"
                  name="importMode"
                  value="replace"
                  checked={importMode === 'replace'}
                  onChange={() => setImportMode('replace')}
                  className="mt-0.5 text-blue-900"
                />
                <div>
                  <div className="text-xs font-bold text-slate-900">
                    استبدال محتوى الورقة الحالية بالكامل
                  </div>
                  <div className="text-[10px] text-slate-500 mt-0.5">
                    فتح العقد المستورد مع توحيده لخط Arial 13pt وهوامش التوثيق.
                  </div>
                </div>
              </label>

              <label
                className={`flex items-start gap-2.5 p-2.5 rounded-lg border cursor-pointer transition-colors ${
                  importMode === 'insert'
                    ? 'border-blue-800 bg-blue-50/40'
                    : 'border-slate-200 hover:bg-slate-50'
                }`}
              >
                <input
                  type="radio"
                  name="importMode"
                  value="insert"
                  checked={importMode === 'insert'}
                  onChange={() => setImportMode('insert')}
                  className="mt-0.5 text-blue-900"
                />
                <div>
                  <div className="text-xs font-bold text-slate-900">
                    إدراج المحتوى عند موضع المؤشر الحالي
                  </div>
                  <div className="text-[10px] text-slate-500 mt-0.5">
                    دمج فقرات وبنود الملف المستورد في موضع المؤشر.
                  </div>
                </div>
              </label>
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between gap-3">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-1.5 text-xs text-slate-600 hover:text-slate-800 rounded-lg hover:bg-slate-200/60 font-medium transition-colors"
          >
            إلغاء
          </button>
          <button
            type="button"
            onClick={() => {
              const finalHtml =
                segmentIntoClauses && candidateClauses.length > 0
                  ? compileCandidateClausesToHtml(candidateClauses)
                  : undefined;
              onConfirmApply(importMode, finalHtml);
              onClose();
            }}
            className="px-4 py-2 bg-blue-900 text-white text-xs font-bold rounded-lg hover:bg-blue-800 transition-colors shadow-2xs inline-flex items-center gap-1.5"
          >
            <Check className="w-4 h-4" />
            <span>
              {importMode === 'replace'
                ? 'اعتماد البنود وفتح العقد بالمحرر'
                : 'إدراج البنود المعتمدة عند المؤشر'}
            </span>
          </button>
        </div>
      </div>
    </div>
  );
}

// =========================================================
// 2. SAVE AS OFFICE TEMPLATE MODAL (حفظ كقالب مكتب مخصص)
// =========================================================
interface SaveAsTemplateModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentDocTitle: string;
  extractedPlaceholders: string[];
  currentFieldValues: Record<string, string>;
  onConfirmSave: (data: {
    name: string;
    category: string;
    description: string;
    clearFilledValues: boolean;
  }) => void;
}

export function SaveAsTemplateModal({
  isOpen,
  onClose,
  currentDocTitle,
  extractedPlaceholders,
  currentFieldValues,
  onConfirmSave,
}: SaveAsTemplateModalProps) {
  const [enteredName, setEnteredName] = useState<string | null>(null);
  const [category, setCategory] = useState<string>('عقود البيع');
  const [customCategory, setCustomCategory] = useState<string>('');
  const [description, setDescription] = useState<string>('');
  const [clearFilledValues, setClearFilledValues] = useState<boolean>(true);

  if (!isOpen) return null;

  const templateName =
    enteredName !== null
      ? enteredName
      : currentDocTitle
      ? `قالب: ${currentDocTitle}`
      : 'قالب عقد توثيقي مخصص';

  const filledCount = extractedPlaceholders.filter(
    (k) => (currentFieldValues[k] || '').trim() !== ''
  ).length;

  const categoryPresets = [
    'عقود البيع',
    'عقود التنازل',
    'عقود الإيجار',
    'وكالات وتفويضات',
    'شركات وتجارة',
    'قروض ورهون',
    'عام',
  ];

  const handleSave = () => {
    if (!templateName.trim()) return;
    const finalCategory = customCategory.trim() || category;
    onConfirmSave({
      name: templateName.trim(),
      category: finalCategory,
      description: description.trim(),
      clearFilledValues,
    });
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xl w-full max-w-lg max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in-50 duration-200">
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-blue-100 text-blue-900 rounded-lg">
              <Save className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                حفظ العقد الحالي كقالب في مكتبة المكتب
              </h2>
              <div className="text-[11px] text-slate-500">
                يُخزن محلياً في قاعدة بيانات المكتب للرجوع إليه وتطبيقه على أي عقد جديد
              </div>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded-lg hover:bg-slate-100 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700">اسم القالب *</label>
            <input
              type="text"
              value={templateName}
              onChange={(e) => setEnteredName(e.target.value)}
              placeholder="مثال: عقد بيع شقة بنظام التمليك المشترك - مكتبنا"
              className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:border-blue-900 focus:outline-none"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700">التصنيف</label>
            <div className="grid grid-cols-3 gap-1.5">
              {categoryPresets.map((cat) => (
                <button
                  key={cat}
                  type="button"
                  onClick={() => {
                    setCategory(cat);
                    setCustomCategory('');
                  }}
                  className={`py-1.5 px-2 rounded text-[11px] font-medium border text-center transition-colors ${
                    category === cat && !customCategory
                      ? 'bg-blue-900 text-white border-blue-900'
                      : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
                  }`}
                >
                  {cat}
                </button>
              ))}
            </div>
            <input
              type="text"
              value={customCategory}
              onChange={(e) => setCustomCategory(e.target.value)}
              placeholder="أو اكتب تصنيفاً خاصاً..."
              className="w-full px-3 py-1.5 text-xs border border-slate-300 rounded-lg focus:border-blue-900 focus:outline-none mt-1"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-xs font-bold text-slate-700">
              وصف أو ملاحظات للاستخدام (اختياري)
            </label>
            <textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              placeholder="ملاحظات حول متى يستخدم هذا القالب والوثائق المرفقة به..."
              rows={2}
              className="w-full px-3 py-2 text-xs border border-slate-300 rounded-lg focus:border-blue-900 focus:outline-none resize-none"
            />
          </div>

          {/* Options for field values */}
          <div className="border border-slate-200 rounded-lg p-3 bg-slate-50 space-y-2">
            <div className="text-xs font-bold text-slate-800">
              معالجة قيم الحقول المعبأة حالياً ({filledCount} من {extractedPlaceholders.length} وسم):
            </div>

            <label className="flex items-start gap-2.5 p-2 bg-white rounded border border-slate-200 cursor-pointer">
              <input
                type="radio"
                name="clearValues"
                checked={clearFilledValues}
                onChange={() => setClearFilledValues(true)}
                className="mt-0.5 text-blue-900"
              />
              <div>
                <div className="text-xs font-semibold text-slate-900">
                  تفريغ القيم المعبأة والاحتفاظ بالوسوم فقط (موصى به)
                </div>
                <div className="text-[11px] text-slate-500">
                  ينتج قالباً عاماً ونظيفاً جاهزاً لإعادة استخدامه مع أي أطراف وعقارات جديدة.
                </div>
              </div>
            </label>

            <label className="flex items-start gap-2.5 p-2 bg-white rounded border border-slate-200 cursor-pointer">
              <input
                type="radio"
                name="clearValues"
                checked={!clearFilledValues}
                onChange={() => setClearFilledValues(false)}
                className="mt-0.5 text-blue-900"
              />
              <div>
                <div className="text-xs font-semibold text-slate-900">
                  الاحتفاظ بالقيم الحالية كقيم افتراضية للقالب
                </div>
                <div className="text-[11px] text-slate-500">
                  يُحفظ القالب مع القيم المكتوبة حالياً لاسترجاعها عند فتح هذا القالب مجدداً.
                </div>
              </div>
            </label>
          </div>
        </div>

        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between gap-3">
          <button
            type="button"
            onClick={onClose}
            className="px-3.5 py-1.5 text-xs text-slate-600 hover:text-slate-800 rounded-lg font-medium"
          >
            إلغاء
          </button>
          <button
            type="button"
            onClick={handleSave}
            disabled={!templateName.trim()}
            className="px-4 py-2 bg-blue-900 text-white text-xs font-bold rounded-lg hover:bg-blue-800 disabled:opacity-50 inline-flex items-center gap-1.5"
          >
            <Save className="w-4 h-4" />
            <span>حفظ القالب في مكتبة المكتب</span>
          </button>
        </div>
      </div>
    </div>
  );
}

// =========================================================
// 3. ONBOARDING TOUR MODAL (جولة إرشادية خفيفة من 5 خطوات للمحرر)
// =========================================================
interface OnboardingTourModalProps {
  isOpen: boolean;
  onClose: () => void;
  onFinishTour: () => void;
}

export function OnboardingTourModal({
  isOpen,
  onClose,
  onFinishTour,
}: OnboardingTourModalProps) {
  const [currentStep, setCurrentStep] = useState<number>(0);

  if (!isOpen) return null;

  const tourSteps = [
    {
      stepNumber: 1,
      title: 'ورقة A4 التوثيقية واستيراد ملفات Word (.docx)',
      description:
        'صُمم المحرر ليعمل تماماً كمكتب توثيق معتمد: الورقة تطابق معايير A4 الدقيقة بهوامش ثابتة (7 سم يمين، 2 سم يسار، 1 سم أعلى، 6 سم أسفل) مع خط Arial 13pt وتباعد 1.0. يمكنك بدء كتابة عقدك فوراً أو استيراد أي ملف Word قديم (.docx) ليُوزّع تلقائياً على الورقة مع الاحتفاظ بالاتجاه العربي والجداول.',
      badgeText: 'هوامش 7/2/1/6 سم + Arial 13pt',
      icon: <FileText className="w-7 h-7 text-blue-900" />,
      tips: [
        'زر «فتح ملف Word (.docx)» يحلل الوثيقة القديمة ويستخرج نصوصها وجداولها بدقة.',
        'يمكنك اختيار استبدال المحتوى بالكامل أو الإدراج عند موضع المؤشر.',
      ],
    },
    {
      stepNumber: 2,
      title: 'إنشاء الوسوم الذكية فوراً والنقر المزدوج (Double Click)',
      description:
        'حوّل أي كلمة أو جملة في العقد إلى وسم ذكي تفاعلي بضغطة واحدة دون فتح أي نافذة: حدد النص واضغط «[ ] تحويل المحدد لوسم» في الشريط أو اختصار Alt+V. وعند قراءة العقد، اضغط نقراً مزدوجاً (Double Click) على أي وسم [...] للانتقال مباشرة إلى حقله المخصص في الاستمارة وتعديل قيمته فوراً.',
      badgeText: 'تحويل سريع بـ 1-Click + dblclick',
      icon: <Tag className="w-7 h-7 text-pink-700" />,
      tips: [
        'اختصار لوحة المفاتيح: Alt+V لتحويل النص المحدد إلى وسم ذكي فوراً.',
        'نقر مزدوج على الوسم يفتح الاستمارة الذكية ويركز المؤشر على الحقل مباشرة.',
      ],
    },
    {
      stepNumber: 3,
      title: 'الاستمارة الديناميكية واستبدال الوسوم نهائياً',
      description:
        'تُجمّع المتغيرات تلقائياً حسب بنود العقد الرسمية لتسهيل تعبئتها خطوة بخطوة. كما يتوفر دفتر محفوظ للأطراف والعقارات لاستدعائها بضغطة زر. وعند الانتهاء، يتيح لك زر «استبدال الوسوم نهائياً» دمج جميع القيم داخل نص العقد مع الحفاظ على التنسيقات والخطوط الرسمية.',
      badgeText: 'تجميع حسب البنود + تفقيط مالي آلي',
      icon: <BookUser className="w-7 h-7 text-emerald-800" />,
      tips: [
        'المبالغ المالية تُحوّل تلقائياً إلى أحرف عربية (تفقيط ديناري جزائري معتمد).',
        'التواريخ تُحوّل تلقائياً إلى الصيغة العربية التوثيقية الكاملة.',
      ],
    },
    {
      stepNumber: 4,
      title: 'البحث العربي الذكي (Ctrl+F) وسجل اللقطات الزمنية',
      description:
        'يدعم شريط البحث والاستبدال المتقدم عبر Ctrl+F تطبيع الأحرف العربية (تجاهل الهمزات أ/إ/آ وتجاهل التشكيل والتاء المربوطة ة/ه)، مع التنقل السلس بين النتائج عبر مختلف الفقرات. كما يسجل النظام لقطات زمنية تلقائية قبل أي عملية استبدال جذري للرجوع لأي نسخة سابقة في أي لحظة.',
      badgeText: 'تطبيع العربية + لقطات أمان تلقائية',
      icon: <Search className="w-7 h-7 text-amber-700" />,
      tips: [
        'شريط البحث يستبدل النصوص حتى عبر الحدود المعقدة للفقرات.',
        'سجل اللقطات (Diff) يتيح مقارنة التعديلات كلمة بكلمة واستعادة أي نسخة سابقة.',
      ],
    },
    {
      stepNumber: 5,
      title: 'تصدير العقد الرسمي وتوليد الوثائق المشتقة (7)',
      description:
        'بمجرد اكتمال الصياغة، صدّر العقد بملف Word (.docx) مهيأ فوراً للطباعة أو الأرشفة، أو ولّد الوثائق المشتقة السبعة التابعة للعقد (المستخرج للتسجيل، قائمة الشهر العقاري، الشهادة التوثيقية، ...) بضغطة زر واحدة من نفس البيانات المعبأة في الاستمارة دون أي إعادة كتابة.',
      badgeText: 'تصدير Word + 7 وثائق مشتقة آلية',
      icon: <Sparkles className="w-7 h-7 text-indigo-900" />,
      tips: [
        'قوالب الوثائق المشتقة قابلة للتعديل والتخصيص الكامل داخل نفس محرر الـ A4.',
        'يمكن حفظ العقد الحالي كقالب مكتب مخصص في أي وقت من القائمة الجانبية.',
      ],
    },
  ];

  const current = tourSteps[currentStep];
  const isFirst = currentStep === 0;
  const isLast = currentStep === tourSteps.length - 1;

  const handleNext = () => {
    if (isLast) {
      onFinishTour();
      onClose();
    } else {
      setCurrentStep((v) => v + 1);
    }
  };

  const handlePrev = () => {
    if (!isFirst) {
      setCurrentStep((v) => v - 1);
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-[#0F2744]/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-2xl border border-[#C4A46A]/45 shadow-2xl w-full max-w-xl overflow-hidden animate-in fade-in-50 duration-200 flex flex-col">
        <div className="px-6 py-3.5 bg-[#F7F5F0] border-b border-[#C4A46A]/35 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <img
              src="/brand/logo-lockup.svg"
              alt="الموثق الرقمي — محرر العقود التوثيقية"
              className="h-9 w-auto select-none shrink-0"
            />
            <span className="text-[11px] px-2.5 py-0.5 bg-[#0F2744] text-[#F7F5F0] border border-[#C4A46A]/50 rounded-full font-bold tabular-nums shrink-0">
              خطوة {currentStep + 1} من {tourSteps.length}
            </span>
          </div>

          <button
            type="button"
            onClick={() => {
              onFinishTour();
              onClose();
            }}
            className="text-xs text-[#5C6B7A] hover:text-[#0F2744] font-medium shrink-0"
          >
            تخطي الجولة
          </button>
        </div>

        <div className="w-full bg-[#F1F5F9] h-1">
          <div
            className="bg-[#C4A46A] h-1 transition-all duration-300"
            style={{ width: `${((currentStep + 1) / tourSteps.length) * 100}%` }}
          />
        </div>

        <div className="p-6 space-y-4">
          <div className="flex items-center gap-3">
            <div className="p-3 bg-[#F7F5F0] border border-[#C4A46A]/35 rounded-xl shrink-0">
              {current.icon}
            </div>
            <div>
              <span className="inline-block text-[11px] font-bold text-[#1E3A8A] bg-blue-50 px-2 py-0.5 rounded mb-1">
                {current.badgeText}
              </span>
              <h3 className="text-base font-bold text-[#0F2744] leading-snug">
                {current.title}
              </h3>
            </div>
          </div>

          <p className="text-xs text-[#5C6B7A] leading-relaxed text-justify">
            {current.description}
          </p>

          <div className="bg-[#F7F5F0] border border-[#C4A46A]/35 rounded-lg p-3 space-y-1.5">
            <div className="text-[11px] font-bold text-[#0F2744]">
              إضاءات وميزات سريعة:
            </div>
            {current.tips.map((tip, idx) => (
              <div key={idx} className="text-[11px] text-[#5C6B7A] flex items-start gap-1.5">
                <span className="text-[#1E3A8A] font-bold shrink-0">✓</span>
                <span>{tip}</span>
              </div>
            ))}
          </div>

          <div className="flex items-center justify-center gap-1.5 pt-2">
            {tourSteps.map((_, idx) => (
              <button
                key={idx}
                type="button"
                onClick={() => setCurrentStep(idx)}
                className={`h-2 rounded-full transition-all ${
                  idx === currentStep ? 'w-6 bg-[#1E3A8A]' : 'w-2 bg-slate-200 hover:bg-slate-300'
                }`}
                title={`الخطوة ${idx + 1}`}
              />
            ))}
          </div>
        </div>

        <div className="px-6 py-3.5 bg-[#F7F5F0] border-t border-[#C4A46A]/35 flex items-center justify-between">
          <button
            type="button"
            onClick={handlePrev}
            disabled={isFirst}
            className="px-3.5 py-1.5 text-xs text-[#5C6B7A] hover:text-[#0F2744] font-medium disabled:opacity-30 inline-flex items-center gap-1"
          >
            <ArrowRight className="w-3.5 h-3.5" />
            <span>السابق</span>
          </button>

          <button
            type="button"
            onClick={handleNext}
            className="px-5 py-2 bg-[#1E3A8A] text-white text-xs font-bold rounded-lg hover:bg-[#0F2744] transition-colors shadow-2xs inline-flex items-center gap-1.5"
          >
            <span>{isLast ? 'إنهاء وبدء التحرير' : 'التالي'}</span>
            {!isLast && <ArrowLeft className="w-3.5 h-3.5" />}
          </button>
        </div>
      </div>
    </div>
  );
}

// =========================================================
// 4. WORD TEMPLATES MODAL (إدارة قوالب Word الرسمية)
// =========================================================
interface WordTemplatesModalProps {
  isOpen: boolean;
  onClose: () => void;
  wordTemplates: WordTemplateDefinition[];
  onSaveWordTemplate: (tpl: WordTemplateDefinition) => void;
  onDeleteWordTemplate: (id: string) => void;
  contractData: {
    officeName?: string;
    officeAddr?: string;
    typeActe?: string;
    client1?: string;
    client2?: string;
    dateActe?: string;
    dateLettre?: string;
    clauses?: { title: string; contentHtml: string }[];
    rawContractBodyHtml?: string;
    fieldValues?: Record<string, string>;
  };
}

export function WordTemplatesModal({
  isOpen,
  onClose,
  wordTemplates,
  onDeleteWordTemplate,
  contractData,
}: WordTemplatesModalProps) {
  const [selectedTpl, setSelectedTpl] = useState<WordTemplateDefinition | null>(null);
  const [previewHtml, setPreviewHtml] = useState<string>('');
  const [showPreviewModal, setShowPreviewModal] = useState<boolean>(false);

  if (!isOpen) return null;

  const handleOpenPreview = (tpl: WordTemplateDefinition) => {
    setSelectedTpl(tpl);
    const rendered = renderTemplateWithContractData(tpl, contractData);
    setPreviewHtml(rendered.bodyHtml);
    setShowPreviewModal(true);
  };

  const handleExport = async (tpl: WordTemplateDefinition) => {
    await exportTemplateAsDocx(tpl, contractData);
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-4xl max-h-[88vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2 className="text-sm font-bold text-slate-900">
              إدارة قوالب Word الرسمية وتوليد الوثائق المشتقة (أصل، مستخرج، شهر، وضعية جبائية)
            </h2>
            <p className="text-[11px] text-slate-500 mt-0.5">
              فصل تام بين نص العقد (المحتوى والبنود) وشكل وثيقة Word الرسمية (القوالب، الهوامش، العناوين).
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {wordTemplates.map((tpl) => (
              <div
                key={tpl.id}
                className="border border-slate-200 rounded-lg p-4 bg-slate-50/50 flex flex-col justify-between space-y-3 hover:border-blue-800 transition-colors"
              >
                <div>
                  <div className="flex items-center justify-between mb-1">
                    <span className="text-xs font-bold text-slate-900">{tpl.name}</span>
                    <span className="text-[10px] font-mono px-2 py-0.5 bg-blue-100 text-blue-900 rounded font-semibold uppercase">
                      {tpl.type}
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-600 leading-relaxed">
                    {tpl.description}
                  </p>
                </div>

                <div className="flex items-center justify-between pt-2 border-t border-slate-200">
                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={() => handleOpenPreview(tpl)}
                      className="px-2.5 py-1 bg-white border border-slate-300 text-xs font-semibold text-slate-800 rounded hover:bg-slate-100 inline-flex items-center gap-1"
                    >
                      <Eye className="w-3.5 h-3.5 text-blue-900" />
                      <span>معاينة سريعة</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => handleExport(tpl)}
                      className="px-2.5 py-1 bg-blue-900 text-white text-xs font-semibold rounded hover:bg-blue-800 inline-flex items-center gap-1"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>تصدير Word (.docx)</span>
                    </button>
                  </div>

                  {!tpl.isDefault && (
                    <button
                      type="button"
                      onClick={() => onDeleteWordTemplate(tpl.id)}
                      className="p-1 text-slate-400 hover:text-red-600 rounded"
                      title="حذف القالب"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  )}
                </div>
              </div>
            ))}
          </div>
        </div>

        {showPreviewModal && selectedTpl && (
          <div className="fixed inset-0 z-60 bg-slate-900/60 flex items-center justify-center p-4">
            <div className="bg-white rounded-lg shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col overflow-hidden border border-slate-300">
              <div className="px-4 py-3 bg-slate-100 border-b border-slate-200 flex items-center justify-between">
                <span className="text-xs font-bold text-slate-900">
                  معاينة وثيقة: {selectedTpl.name}
                </span>
                <button
                  type="button"
                  onClick={() => setShowPreviewModal(false)}
                  className="p-1 text-slate-400 hover:text-slate-700"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
              <div
                dir="rtl"
                className="flex-1 overflow-y-auto p-6 bg-slate-50 text-sm leading-relaxed space-y-3"
                style={{ fontFamily: 'Arial, sans-serif' }}
                dangerouslySetInnerHTML={{ __html: previewHtml }}
              />
              <div className="px-4 py-3 bg-slate-100 border-t border-slate-200 flex items-center justify-end gap-2">
                <button
                  type="button"
                  onClick={() => setShowPreviewModal(false)}
                  className="px-3 py-1.5 bg-slate-200 text-slate-800 text-xs font-medium rounded hover:bg-slate-300"
                >
                  إغلاق المعاينة
                </button>
                <button
                  type="button"
                  onClick={() => {
                    handleExport(selectedTpl);
                    setShowPreviewModal(false);
                  }}
                  className="px-4 py-1.5 bg-blue-900 text-white text-xs font-semibold rounded hover:bg-blue-800 inline-flex items-center gap-1.5"
                >
                  <Download className="w-4 h-4" />
                  <span>تصدير فوري كملف Word</span>
                </button>
              </div>
            </div>
          </div>
        )}

        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-blue-900 text-white text-xs font-semibold rounded hover:bg-blue-800"
          >
            إغلاق النافذة
          </button>
        </div>
      </div>
    </div>
  );
}
