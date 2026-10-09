'use client';

import React, { useEffect, useRef, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  Download,
  FilePlus2,
  FileText,
  FileUp,
  GitCompare,
  History,
  RotateCcw,
  Sparkles,
  X,
} from 'lucide-react';
import {
  CustomTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  PartyField,
  SavedDocument,
  SubdivisionEstate,
} from '@/lib/types';
import { computeTextDiff } from '@/lib/editor-utils';

interface SmartVariablesModalProps {
  isOpen: boolean;
  onClose: () => void;
  focusedVarKey: string | null;
  extractedPlaceholders: string[];
  partyFields: PartyField[];
  fieldValues: Record<string, string>;
  estates: SubdivisionEstate[];
  selectedEstateId: string;
  selectedLotNumber: string;
  onSelectEstateAndLot: (estateId: string, lotNumber: string) => void;
  onUpdateFieldValue: (key: string, value: string) => void;
  onBakeAllIntoDocument: () => void;
}

export function SmartVariablesModal({
  isOpen,
  onClose,
  focusedVarKey,
  extractedPlaceholders,
  partyFields,
  fieldValues,
  estates,
  selectedEstateId,
  selectedLotNumber,
  onSelectEstateAndLot,
  onUpdateFieldValue,
  onBakeAllIntoDocument,
}: SmartVariablesModalProps) {
  const inputRefs = useRef<Record<string, HTMLInputElement | HTMLSelectElement | null>>({});

  useEffect(() => {
    if (isOpen && focusedVarKey) {
      setTimeout(() => {
        const el = inputRefs.current[focusedVarKey];
        if (el) {
          el.focus();
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }, 80);
    }
  }, [isOpen, focusedVarKey]);

  if (!isOpen) return null;

  // Build unified variable list: document-extracted variables first, then standard party/property fields
  const knownMap = new Map(partyFields.map((f) => [f.key, f]));
  const allKeys = Array.from(
    new Set([...extractedPlaceholders, ...partyFields.map((f) => f.key)])
  );

  const unfilledInDoc = extractedPlaceholders.filter(
    (k) => !fieldValues[k] || !fieldValues[k].trim()
  );

  const selectedEstate = estates.find((e) => e.id === selectedEstateId) || estates[0] || null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-3xl max-h-[88vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-bold text-slate-900">
                استمارة المتغيرات الذكية (المصدر الموحد للعقد والوثائق المشتقة)
              </h2>
              {unfilledInDoc.length > 0 ? (
                <span className="text-xs font-semibold text-amber-800 bg-amber-100 px-2 py-0.5 rounded tabular-nums">
                  متبقي {unfilledInDoc.length} غير معبأ
                </span>
              ) : (
                <span className="text-xs font-semibold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded">
                  جميع متغيرات العقد مكتملة ✓
                </span>
              )}
            </div>
            <p className="text-[11px] text-slate-500 mt-0.5">
              القيم المدخلة هنا تُعوض تلقائياً في العقد الأصلي وفي جميع الوثائق المشتقة (المستخرج، إجراء الشهر، شهادة البيع، الصيغة التنفيذية).
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

        {/* Body */}
        <div className="flex-1 overflow-y-auto p-5 space-y-5">
          {/* Quick Subdivision Lot Selector inside Smart Form */}
          {estates.length > 0 && (
            <div className="p-3 bg-blue-50/60 border border-blue-200 rounded-md flex flex-wrap items-center gap-3">
              <div className="text-xs font-bold text-blue-950 shrink-0">
                ربط تلقائي من جدول الوصف التقسيمي:
              </div>
              <select
                value={selectedEstate?.id || ''}
                onChange={(e) => onSelectEstateAndLot(e.target.value, '')}
                className="px-2.5 py-1 text-xs bg-white border border-slate-300 rounded"
              >
                {estates.map((est) => (
                  <option key={est.id} value={est.id}>
                    {est.estateName}
                  </option>
                ))}
              </select>
              {selectedEstate && (
                <select
                  value={selectedLotNumber}
                  onChange={(e) => onSelectEstateAndLot(selectedEstate.id, e.target.value)}
                  className="px-2.5 py-1 text-xs bg-white border border-blue-400 rounded font-semibold tabular-nums"
                >
                  <option value="">-- اختر رقم الحصة لملء التعيين --</option>
                  {selectedEstate.lots.map((l) => (
                    <option key={l.id} value={l.lotNumber}>
                      حصة رقم {l.lotNumber} — {l.nature} ({l.floor}) — {l.area} م²
                    </option>
                  ))}
                </select>
              )}
            </div>
          )}

          {/* Variables Grid */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
            {allKeys.map((key) => {
              const meta = knownMap.get(key);
              const label = meta?.label || key.replace(/_/g, ' ');
              const inCurrentDoc = extractedPlaceholders.includes(key);
              const val = fieldValues[key] || '';
              const isEmptyInDoc = inCurrentDoc && !val.trim();
              const isFocused = focusedVarKey === key;

              return (
                <div
                  key={key}
                  className={`p-2.5 rounded-md border transition-colors ${
                    isFocused
                      ? 'border-pink-600 bg-pink-50/40 ring-2 ring-pink-500/20'
                      : isEmptyInDoc
                      ? 'border-amber-300 bg-amber-50/30'
                      : 'border-slate-200 bg-white'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2 mb-1">
                    <label className="text-xs font-bold text-slate-800 truncate">
                      {label}
                    </label>
                    <span className="text-[10px] font-mono text-pink-800 bg-pink-50 px-1.5 py-0.5 rounded shrink-0">
                      {`{{${key}}}`}
                    </span>
                  </div>

                  {meta?.inputType === 'select' && meta.options ? (
                    <select
                      ref={(el) => {
                        inputRefs.current[key] = el;
                      }}
                      value={val}
                      onChange={(e) => onUpdateFieldValue(key, e.target.value)}
                      className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
                    >
                      <option value="">-- اختر أو اكتب أدناه --</option>
                      {meta.options.map((opt) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  ) : (
                    <input
                      ref={(el) => {
                        inputRefs.current[key] = el;
                      }}
                      type={meta?.inputType === 'date' ? 'date' : 'text'}
                      value={val}
                      onChange={(e) => onUpdateFieldValue(key, e.target.value)}
                      placeholder={`أدخل ${label}...`}
                      className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
                    />
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Footer Actions */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
          <button
            type="button"
            onClick={() => {
              onBakeAllIntoDocument();
              onClose();
            }}
            className="px-3.5 py-1.5 bg-pink-800 text-white text-xs font-medium rounded hover:bg-pink-900 transition-colors"
            title="يستبدل كل وسم {{...}} داخل ورقة الـ A4 بقيمته المكتوبة مع الحفاظ على التنسيق المحيط"
          >
            استبدال الوسوم نهائياً داخل نص العقد
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1.5 bg-blue-900 text-white text-xs font-semibold rounded hover:bg-blue-800 transition-colors"
            >
              حفظ واعتماد القيم
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

interface VersionDiffModalProps {
  isOpen: boolean;
  onClose: () => void;
  revisions: DocumentRevision[];
  currentBodyHtml: string;
  currentTitle: string;
  onRestoreRevision: (rev: DocumentRevision) => void;
}

export function VersionDiffModal({
  isOpen,
  onClose,
  revisions,
  currentBodyHtml,
  currentTitle,
  onRestoreRevision,
}: VersionDiffModalProps) {
  const [leftRevId, setLeftRevId] = useState<string>('');
  const [rightRevId, setRightRevId] = useState<string>('CURRENT');

  useEffect(() => {
    if (isOpen && revisions.length > 0 && !leftRevId) {
      setLeftRevId(revisions[0].id);
    }
  }, [isOpen, revisions, leftRevId]);

  if (!isOpen) return null;

  const leftRev = revisions.find((r) => r.id === leftRevId) || revisions[0] || null;
  const rightRev =
    rightRevId === 'CURRENT'
      ? null
      : revisions.find((r) => r.id === rightRevId) || null;

  const oldHtml = leftRev ? leftRev.bodyHtml : '';
  const newHtml = rightRev ? rightRev.bodyHtml : currentBodyHtml;

  const diffSegments = leftRev ? computeTextDiff(oldHtml, newHtml) : [];
  const addedCount = diffSegments.filter((s) => s.type === 'added').length;
  const removedCount = diffSegments.filter((s) => s.type === 'removed').length;

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-4xl max-h-[88vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <GitCompare className="w-4 h-4 text-blue-900" />
            <h2 className="text-sm font-bold text-slate-900">
              نظام مقارنة النسخ والتعديلات (Diff) — {currentTitle}
            </h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="p-4 border-b border-slate-200 bg-slate-50/50 grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">
              النسخة المرجعية (الأقدم):
            </label>
            <select
              value={leftRev?.id || ''}
              onChange={(e) => setLeftRevId(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded tabular-nums"
            >
              {revisions.length === 0 ? (
                <option value="">لا توجد نسخ محفوظة بعد</option>
              ) : (
                revisions.map((r) => (
                  <option key={r.id} value={r.id}>
                    {new Date(r.createdAt).toLocaleString('ar-DZ')} — {r.summary} ({r.author})
                  </option>
                ))
              )}
            </select>
          </div>

          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">
              المقارنة مع النسخة (الأحدث):
            </label>
            <select
              value={rightRevId}
              onChange={(e) => setRightRevId(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded tabular-nums"
            >
              <option value="CURRENT">النسخة الحالية قيد التحرير الآن</option>
              {revisions.map((r) => (
                <option key={r.id} value={r.id}>
                  {new Date(r.createdAt).toLocaleString('ar-DZ')} — {r.summary} ({r.author})
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Diff Legend & Content */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {revisions.length === 0 ? (
            <div className="text-center py-12 text-xs text-slate-500">
              لم يتم تسجيل نسخ سابقة بعد. يتم تسجيل النسخ تلقائياً عند الحفظ أو التعديل لعرض الفروقات هنا.
            </div>
          ) : (
            <>
              <div className="flex items-center justify-between text-xs border-b border-slate-200 pb-2">
                <div className="flex items-center gap-4">
                  <span className="inline-flex items-center gap-1.5 text-emerald-800 font-medium">
                    <span className="w-3 h-3 rounded-xs bg-emerald-200 inline-block" />
                    نص مضاف ({addedCount})
                  </span>
                  <span className="inline-flex items-center gap-1.5 text-red-800 font-medium">
                    <span className="w-3 h-3 rounded-xs bg-red-200 inline-block" />
                    نص محذوف ({removedCount})
                  </span>
                </div>
                {leftRev && (
                  <button
                    type="button"
                    onClick={() => {
                      onRestoreRevision(leftRev);
                      onClose();
                    }}
                    className="px-3 py-1 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800 inline-flex items-center gap-1.5"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>استعادة النسخة المرجعية المختارة</span>
                  </button>
                )}
              </div>

              <div
                dir="rtl"
                className="p-4 bg-slate-50 border border-slate-200 rounded-md text-sm leading-relaxed whitespace-pre-wrap"
                style={{ fontFamily: 'Arial, sans-serif' }}
              >
                {diffSegments.map((seg, idx) => {
                  if (seg.type === 'added') {
                    return (
                      <span
                        key={idx}
                        className="bg-emerald-100 text-emerald-950 font-semibold px-0.5 rounded"
                      >
                        {seg.text}
                      </span>
                    );
                  }
                  if (seg.type === 'removed') {
                    return (
                      <span
                        key={idx}
                        className="bg-red-100 text-red-900 line-through px-0.5 rounded"
                      >
                        {seg.text}
                      </span>
                    );
                  }
                  return <span key={idx}>{seg.text}</span>;
                })}
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

interface MultiSourceStartModalProps {
  isOpen: boolean;
  onClose: () => void;
  templates: CustomTemplate[];
  documents: SavedDocument[];
  downloads: DownloadArchiveItem[];
  onStartBlank: () => void;
  onSelectTemplate: (tpl: CustomTemplate) => void;
  onSelectSavedDoc: (doc: SavedDocument) => void;
  onImportDocxFile: (file: File) => void;
  onSelectDownloadArchiveItem: (item: DownloadArchiveItem) => void;
}

export function MultiSourceStartModal({
  isOpen,
  onClose,
  templates,
  documents,
  downloads,
  onStartBlank,
  onSelectTemplate,
  onSelectSavedDoc,
  onImportDocxFile,
  onSelectDownloadArchiveItem,
}: MultiSourceStartModalProps) {
  const [sourceTab, setSourceTab] = useState<'blank' | 'templates' | 'docs' | 'archive'>('blank');

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[85vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <h2 className="text-sm font-bold text-slate-900">
            بدء تحرير عقد جديد (مصادر البدء المتعددة)
          </h2>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-slate-400 hover:text-slate-700 rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Source Tabs */}
        <div className="grid grid-cols-4 gap-1 p-2 bg-slate-100 border-b border-slate-200 text-xs font-medium">
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
            المسودات السابقة ({documents.length})
          </button>
          <button
            type="button"
            onClick={() => setSourceTab('archive')}
            className={`py-1.5 rounded ${
              sourceTab === 'archive' ? 'bg-white text-slate-900 shadow-2xs' : 'text-slate-600'
            }`}
          >
            من أرشيف التحميلات ({downloads.length})
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-5">
          {sourceTab === 'blank' && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              <button
                type="button"
                onClick={() => {
                  onStartBlank();
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
          )}

          {sourceTab === 'templates' && (
            <div className="space-y-2">
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
                        {tpl.category} · {tpl.extractedPlaceholders.length} وسم ذكي
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        onSelectTemplate(tpl);
                        onClose();
                      }}
                      className="px-3 py-1 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800"
                    >
                      البدء بهذا القالب
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

          {sourceTab === 'archive' && (
            <div className="space-y-2">
              {downloads.length === 0 ? (
                <div className="text-center py-8 text-xs text-slate-500">
                  أرشيف التحميلات فارغ حالياً.
                </div>
              ) : (
                downloads.map((item) => (
                  <div
                    key={item.id}
                    className="flex items-center justify-between p-3 border border-slate-200 rounded hover:bg-slate-50"
                  >
                    <div>
                      <div className="text-xs font-bold text-slate-900">
                        {item.documentTitle} — ({item.docTypeLabel})
                      </div>
                      <div className="text-[11px] text-slate-500 tabular-nums">
                        {new Date(item.createdAt).toLocaleString('ar-DZ')} · {item.fileName}
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => {
                        onSelectDownloadArchiveItem(item);
                        onClose();
                      }}
                      className="px-3 py-1 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800"
                    >
                      استيراد للمحرر
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
