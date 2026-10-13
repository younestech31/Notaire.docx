'use client';

import React, { useState } from 'react';
import {
  Check,
  Copy,
  Download,
  FileCheck,
  FileCode,
  FileText,
  HelpCircle,
  RotateCcw,
  Sparkles,
  Upload,
  X,
} from 'lucide-react';
import {
  BUILTIN_STANDARD_WORD_TEMPLATES,
  downloadDefaultNotaryDocxTemplate,
  getCustomizedNotaryTemplateMeta,
  hasCustomizedNotaryTemplate,
  resetNotaryTemplateToDefault,
  saveCustomizedNotaryTemplateFile,
  StandardNotaryTemplateConfig,
  StandardNotaryTemplateId,
} from '@/lib/notary-word-templates';
import {
  STANDARD_NOTARY_TAG_DEFINITIONS,
  StandardNotaryTagDefinition,
} from '@/lib/notary-word-variables';

interface WordTemplatesManagerModalProps {
  isOpen: boolean;
  onClose: () => void;
  onShowToast: (message: string) => void;
  onSelectExportTemplate?: (templateId: StandardNotaryTemplateId) => void;
}

export default function WordTemplatesManagerModal({
  isOpen,
  onClose,
  onShowToast,
  onSelectExportTemplate,
}: WordTemplatesManagerModalProps) {
  const [activeTab, setActiveTab] = useState<'templates' | 'variables'>('templates');
  const [selectedTagCategory, setSelectedTagCategory] = useState<string>('all');
  const [copiedTag, setCopiedTag] = useState<string | null>(null);
  const [refreshKey, setRefreshKey] = useState<number>(0);

  if (!isOpen) return null;

  const templateList: StandardNotaryTemplateConfig[] = Object.values(
    BUILTIN_STANDARD_WORD_TEMPLATES
  );

  const handleDownloadTemplate = async (tplId: StandardNotaryTemplateId) => {
    try {
      await downloadDefaultNotaryDocxTemplate(tplId);
      const name = BUILTIN_STANDARD_WORD_TEMPLATES[tplId].name;
      onShowToast(`تم تنزيل قالب Word «${name}» بنجاح بصيغة (.docx)`);
    } catch {
      onShowToast('حدث خطأ أثناء تنزيل القالب');
    }
  };

  const handleUploadCustomDocx = async (
    tplId: StandardNotaryTemplateId,
    file: File
  ) => {
    try {
      await saveCustomizedNotaryTemplateFile(tplId, file);
      setRefreshKey((k) => k + 1);
      const name = BUILTIN_STANDARD_WORD_TEMPLATES[tplId].name;
      onShowToast(`تم تخصيص وحفظ قالب Word «${name}» للمكتب بنجاح`);
    } catch {
      onShowToast('تعذر حفظ ملف القالب المخصص');
    }
  };

  const handleResetToDefault = (tplId: StandardNotaryTemplateId) => {
    resetNotaryTemplateToDefault(tplId);
    setRefreshKey((k) => k + 1);
    const name = BUILTIN_STANDARD_WORD_TEMPLATES[tplId].name;
    onShowToast(`تمت استعادة القالب الافتراضي لـ «${name}»`);
  };

  const handleCopyTag = (tag: string) => {
    const formatted = `[${tag}]`;
    if (navigator.clipboard) {
      navigator.clipboard.writeText(formatted);
      setCopiedTag(tag);
      setTimeout(() => setCopiedTag((c) => (c === tag ? null : c)), 2000);
      onShowToast(`تم نسخ الوسم ${formatted}`);
    }
  };

  const filteredTags = STANDARD_NOTARY_TAG_DEFINITIONS.filter((t) => {
    if (selectedTagCategory === 'all') return true;
    return t.category === selectedTagCategory;
  });

  return (
    <div
      key={refreshKey}
      className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-3 sm:p-4 select-none"
    >
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xl w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden animate-in fade-in-50 duration-200">
        {/* Modal Header */}
        <div className="px-5 py-3.5 bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 text-white flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-blue-600/30 border border-blue-400/30 rounded-lg">
              <FileCode className="w-5 h-5 text-blue-300" />
            </div>
            <div>
              <h2 className="text-sm font-bold flex items-center gap-2">
                <span>مكتبة وتخصيص قوالب Word (.docx) التوثيقية</span>
                <span className="text-[10px] px-2 py-0.5 rounded-full bg-blue-500/20 text-blue-200 border border-blue-400/30 font-mono">
                  Word (.docx)
                </span>
              </h2>
              <p className="text-[11px] text-slate-300 mt-0.5">
                القوالب الستة الرسمية للمكتب: تنزيل ملفات .docx، تخصيص الهوية البصرية، واعتماد وسوم التوثيق اللاتينية القياسية.
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Modal Nav Tabs */}
        <div className="px-5 pt-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveTab('templates')}
              className={`pb-2.5 px-3 text-xs font-bold border-b-2 transition-colors flex items-center gap-1.5 ${
                activeTab === 'templates'
                  ? 'border-blue-900 text-blue-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileText className="w-3.5 h-3.5" />
              <span>القوالب التوثيقية الستة (6)</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('variables')}
              className={`pb-2.5 px-3 text-xs font-bold border-b-2 transition-colors flex items-center gap-1.5 ${
                activeTab === 'variables'
                  ? 'border-blue-900 text-blue-900'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileCode className="w-3.5 h-3.5" />
              <span>دليل الرموز والمتغيرات اللاتينية ([tags])</span>
            </button>
          </div>

          <div className="text-[11px] text-slate-500 hidden sm:block pb-2">
            صيغة المتغيرات داخل ملفات الوورد: <code className="bg-slate-200 px-1 py-0.5 rounded text-blue-950 font-mono font-bold">[tag_name]</code>
          </div>
        </div>

        {/* Modal Body */}
        <div className="p-5 overflow-y-auto flex-1 space-y-4">
          {activeTab === 'templates' && (
            <div className="space-y-4">
              <div className="bg-blue-50/70 border border-blue-200 rounded-lg p-3 text-xs text-blue-950 flex items-start gap-2.5">
                <Sparkles className="w-4 h-4 text-blue-700 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <span className="font-bold">كيف يعمل نظام القوالب؟</span>
                  <p className="text-[11px] leading-relaxed text-blue-900">
                    قوالب النظام هي ملفات Word جاهزة تمثل التنسيق النهائي والهوية البصرية للمكتب. يمكنكم تنزيل أي قالب كملف <code className="font-mono bg-blue-100 px-1 rounded">.docx</code>، تعديل ترويسة المكتب وخاتمه في برنامج Word، ثم رفع الملف المخصص ليصبح هو المعتمد تلقائياً عند التصدير.
                  </p>
                </div>
              </div>

              {/* Grid of 6 Templates */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                {templateList.map((tpl) => {
                  const isCustom = hasCustomizedNotaryTemplate(tpl.id);
                  const meta = getCustomizedNotaryTemplateMeta(tpl.id);

                  return (
                    <div
                      key={tpl.id}
                      className="border border-slate-200 rounded-lg p-3.5 bg-white hover:border-slate-300 transition-shadow shadow-2xs flex flex-col justify-between gap-3"
                    >
                      <div className="space-y-1.5">
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-1.5">
                            <span className="text-xs font-bold text-slate-900">
                              {tpl.name}
                            </span>
                            <span className="text-[10px] text-slate-500 font-sans">
                              ({tpl.frenchTitle})
                            </span>
                          </div>
                          <span
                            className={`px-2 py-0.5 text-[10px] font-bold rounded-full ${
                              isCustom
                                ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                : 'bg-emerald-100 text-emerald-900 border border-emerald-300'
                            }`}
                          >
                            {isCustom ? 'قالب مخصص للمكتب' : 'قالب افتراضي'}
                          </span>
                        </div>

                        <p className="text-[11px] text-slate-600 leading-relaxed">
                          {tpl.description}
                        </p>

                        {isCustom && meta && (
                          <div className="text-[10px] text-amber-900 bg-amber-50 p-1.5 rounded border border-amber-200">
                            الملف المخصص: <span className="font-bold">{meta.fileName}</span> (
                            {new Date(meta.updatedAt).toLocaleDateString('ar-DZ')})
                          </div>
                        )}
                      </div>

                      {/* Action buttons */}
                      <div className="pt-2 border-t border-slate-100 flex items-center justify-between gap-1.5 flex-wrap">
                        <div className="flex items-center gap-1.5">
                          {/* Download default docx */}
                          <button
                            type="button"
                            onClick={() => handleDownloadTemplate(tpl.id)}
                            className="px-2.5 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-800 text-[11px] font-medium rounded inline-flex items-center gap-1 transition-colors"
                            title="تنزيل ملف القالب بصيغة Word (.docx) لمعاينته أو تعديله"
                          >
                            <Download className="w-3 h-3 text-blue-900" />
                            <span>تنزيل (.docx)</span>
                          </button>

                          {/* Upload custom docx */}
                          <label
                            className="px-2.5 py-1.5 bg-blue-900 hover:bg-blue-800 text-white text-[11px] font-medium rounded inline-flex items-center gap-1 cursor-pointer transition-colors"
                            title="رفع ملف Word معدل (.docx) ليصبح هو القالب المعتمد للمكتب"
                          >
                            <Upload className="w-3 h-3" />
                            <span>تخصيص ورفع</span>
                            <input
                              type="file"
                              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                              className="hidden"
                              onChange={(e) => {
                                const f = e.target.files?.[0];
                                if (f) {
                                  handleUploadCustomDocx(tpl.id, f);
                                  e.target.value = '';
                                }
                              }}
                            />
                          </label>

                          {/* Reset to default */}
                          {isCustom && (
                            <button
                              type="button"
                              onClick={() => handleResetToDefault(tpl.id)}
                              className="p-1.5 text-slate-500 hover:text-red-700 hover:bg-slate-100 rounded"
                              title="استعادة القالب الافتراضي المدمج"
                            >
                              <RotateCcw className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>

                        {onSelectExportTemplate && (
                          <button
                            type="button"
                            onClick={() => {
                              onSelectExportTemplate(tpl.id);
                              onClose();
                            }}
                            className="px-2 py-1 bg-emerald-700 hover:bg-emerald-800 text-white text-[11px] font-semibold rounded inline-flex items-center gap-1 transition-colors"
                            title="تصدير العقد الحالي وحقن بياناته في هذا القالب فوراً"
                          >
                            <FileCheck className="w-3 h-3" />
                            <span>تصدير الآن</span>
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {activeTab === 'variables' && (
            <div className="space-y-3.5">
              <div className="flex items-center justify-between gap-2 flex-wrap">
                <div className="text-xs font-bold text-slate-900">
                  قائمة الرموز والمتغيرات اللاتينية المعتمدة في ملفات Word
                </div>

                {/* Categories Filter */}
                <div className="flex items-center gap-1 bg-slate-100 p-1 rounded text-[11px]">
                  <button
                    type="button"
                    onClick={() => setSelectedTagCategory('all')}
                    className={`px-2 py-0.5 rounded transition-colors ${
                      selectedTagCategory === 'all'
                        ? 'bg-white text-slate-900 font-bold shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    الكل ({STANDARD_NOTARY_TAG_DEFINITIONS.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedTagCategory('parties')}
                    className={`px-2 py-0.5 rounded transition-colors ${
                      selectedTagCategory === 'parties'
                        ? 'bg-white text-blue-900 font-bold shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    الأطراف
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedTagCategory('acte')}
                    className={`px-2 py-0.5 rounded transition-colors ${
                      selectedTagCategory === 'acte'
                        ? 'bg-white text-blue-900 font-bold shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    العقد والمكتب
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedTagCategory('immeuble')}
                    className={`px-2 py-0.5 rounded transition-colors ${
                      selectedTagCategory === 'immeuble'
                        ? 'bg-white text-blue-900 font-bold shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    العقار والتقسيم
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedTagCategory('finance')}
                    className={`px-2 py-0.5 rounded transition-colors ${
                      selectedTagCategory === 'finance'
                        ? 'bg-white text-blue-900 font-bold shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    المالية والثمن
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedTagCategory('texte')}
                    className={`px-2 py-0.5 rounded transition-colors ${
                      selectedTagCategory === 'texte'
                        ? 'bg-white text-blue-900 font-bold shadow-2xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    متن العقد
                  </button>
                </div>
              </div>

              {/* Tags Reference Table */}
              <div className="border border-slate-200 rounded-lg overflow-hidden">
                <table className="w-full text-xs text-right border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                      <th className="p-2.5">الوسم اللاتيني (Word Tag)</th>
                      <th className="p-2.5">البيان والحقل المعادل</th>
                      <th className="p-2.5">قيمة توضيحية</th>
                      <th className="p-2.5 text-center">نسخ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredTags.map((def: StandardNotaryTagDefinition) => {
                      const isCopied = copiedTag === def.tag;
                      return (
                        <tr
                          key={def.tag}
                          className="hover:bg-slate-50/70 transition-colors"
                        >
                          <td className="p-2.5 font-mono text-[11px] text-blue-950 font-semibold dir-ltr text-left">
                            <span className="bg-blue-50 text-blue-900 px-1.5 py-0.5 rounded border border-blue-200">
                              [{def.tag}]
                            </span>
                          </td>
                          <td className="p-2.5">
                            <div className="font-semibold text-slate-900">
                              {def.label}
                            </div>
                            <div className="text-[10px] text-slate-500 font-mono mt-0.5">
                              {def.mappedFieldKeys.join(' / ')}
                            </div>
                          </td>
                          <td className="p-2.5 text-slate-600 text-[11px] truncate max-w-xs">
                            {def.sampleValue}
                          </td>
                          <td className="p-2.5 text-center">
                            <button
                              type="button"
                              onClick={() => handleCopyTag(def.tag)}
                              className={`p-1.5 rounded transition-colors ${
                                isCopied
                                  ? 'bg-emerald-100 text-emerald-800'
                                  : 'text-slate-500 hover:bg-slate-200 hover:text-slate-900'
                              }`}
                              title={`نسخ [${def.tag}]`}
                            >
                              {isCopied ? (
                                <Check className="w-3.5 h-3.5 text-emerald-700" />
                              ) : (
                                <Copy className="w-3.5 h-3.5" />
                              )}
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-xs">
          <div className="text-slate-500 flex items-center gap-1.5">
            <HelpCircle className="w-4 h-4 text-slate-400" />
            <span>
              يتم حفظ أي قالب مخصص للمكتب تلقائياً، وتستبدل الوسوم اللاتينية عند كل عملية تصدير.
            </span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-900 text-white font-medium rounded-lg hover:bg-slate-800 transition-colors"
          >
            إغلاق
          </button>
        </div>
      </div>
    </div>
  );
}
