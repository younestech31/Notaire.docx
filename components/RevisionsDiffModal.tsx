'use client';

import React, { useState } from 'react';
import {
  Clock,
  Columns2,
  FileText,
  GitCompare,
  History,
  RotateCcw,
  Trash2,
  UserCheck,
  X,
} from 'lucide-react';
import { DocumentRevision } from '@/lib/types';
import { computeDocumentMetrics, computeTextDiff } from '@/lib/editor-utils';

export interface VersionDiffModalProps {
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
  const [selectedLeftRevId, setSelectedLeftRevId] = useState<string>('');
  const [rightRevId, setRightRevId] = useState<string>('CURRENT');
  const [diffViewMode, setDiffViewMode] = useState<'side-by-side' | 'unified'>('side-by-side');

  if (!isOpen) return null;

  const leftRevId = selectedLeftRevId || revisions[0]?.id || '';
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
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-5xl max-h-[88vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <GitCompare className="w-4 h-4 text-blue-900" />
            <h2 className="text-sm font-bold text-slate-900">
              نظام مقارنة النسخ والتعديلات جنباً إلى جنب (Diff) — {currentTitle}
            </h2>
          </div>

          <div className="flex items-center gap-2">
            <div className="inline-flex items-center bg-slate-200/80 rounded p-0.5 text-xs">
              <button
                type="button"
                onClick={() => setDiffViewMode('side-by-side')}
                className={`px-2.5 py-1 rounded inline-flex items-center gap-1 transition-colors ${
                  diffViewMode === 'side-by-side'
                    ? 'bg-white text-blue-950 font-bold shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Columns2 className="w-3.5 h-3.5" />
                <span>جنباً إلى جنب (Side-by-Side)</span>
              </button>
              <button
                type="button"
                onClick={() => setDiffViewMode('unified')}
                className={`px-2.5 py-1 rounded inline-flex items-center gap-1 transition-colors ${
                  diffViewMode === 'unified'
                    ? 'bg-white text-blue-950 font-bold shadow-2xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <FileText className="w-3.5 h-3.5" />
                <span>عرض مدمج موحد</span>
              </button>
            </div>

            <button
              type="button"
              onClick={onClose}
              className="p-1.5 text-slate-400 hover:text-slate-700 rounded"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>

        <div className="p-4 border-b border-slate-200 bg-slate-50/50 grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">
              النسخة المرجعية (الأقدم — العمود الأيمن):
            </label>
            <select
              value={leftRev?.id || ''}
              onChange={(e) => setSelectedLeftRevId(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded tabular-nums"
            >
              {revisions.length === 0 ? (
                <option value="">لا توجد نسخ محفوظة بعد</option>
              ) : (
                revisions.map((r) => (
                  <option key={r.id} value={r.id}>
                    {new Date(r.createdAt).toLocaleString('ar-DZ')} — {r.summary} ({r.author || 'محرر المكتب'})
                  </option>
                ))
              )}
            </select>
          </div>

          <div>
            <label className="block text-[11px] font-semibold text-slate-600 mb-1">
              المقارنة مع النسخة (الأحدث — العمود الأيسر):
            </label>
            <select
              value={rightRevId}
              onChange={(e) => setRightRevId(e.target.value)}
              className="w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded tabular-nums"
            >
              <option value="CURRENT">النسخة الحالية قيد التحرير الآن</option>
              {revisions.map((r) => (
                <option key={r.id} value={r.id}>
                  {new Date(r.createdAt).toLocaleString('ar-DZ')} — {r.summary} ({r.author || 'محرر المكتب'})
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Diff Legend & Content */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {revisions.length === 0 ? (
            <div className="text-center py-12 text-xs text-slate-500">
              لم يتم تسجيل نسخ سابقة بعد. يتم تسجيل النسخ تلقائياً عند الحفظ أو تعديل البنود لعرض الفروقات هنا.
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
                  {leftRev && (
                    <span className="inline-flex items-center gap-1 text-[11px] text-slate-600 bg-slate-100 px-2 py-0.5 rounded">
                      <UserCheck className="w-3 h-3 text-blue-900" />
                      <span>الكاتب: {leftRev.author || 'محرر المكتب'}</span>
                    </span>
                  )}
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

              {diffViewMode === 'side-by-side' ? (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="border border-slate-200 rounded-md overflow-hidden flex flex-col bg-white">
                    <div className="px-3 py-2 bg-red-50/60 border-b border-slate-200 flex items-center justify-between text-xs font-bold text-slate-800">
                      <span>النسخة المرجعية (قبل التعديل)</span>
                      <span className="text-[11px] font-normal text-red-800">
                        المحذوفات: {removedCount}
                      </span>
                    </div>
                    <div
                      dir="rtl"
                      className="p-4 text-sm leading-relaxed whitespace-pre-wrap flex-1 bg-slate-50/30"
                      style={{ fontFamily: 'Arial, sans-serif' }}
                    >
                      {diffSegments
                        .filter((seg) => seg.type !== 'added')
                        .map((seg, idx) =>
                          seg.type === 'removed' ? (
                            <span
                              key={idx}
                              className="bg-red-100 text-red-900 line-through px-0.5 rounded font-semibold"
                            >
                              {seg.text}
                            </span>
                          ) : (
                            <span key={idx}>{seg.text}</span>
                          )
                        )}
                    </div>
                  </div>

                  <div className="border border-slate-200 rounded-md overflow-hidden flex flex-col bg-white">
                    <div className="px-3 py-2 bg-emerald-50/60 border-b border-slate-200 flex items-center justify-between text-xs font-bold text-slate-800">
                      <span>النسخة المقارنة (بعد التعديل)</span>
                      <span className="text-[11px] font-normal text-emerald-800">
                        الإضافات: {addedCount}
                      </span>
                    </div>
                    <div
                      dir="rtl"
                      className="p-4 text-sm leading-relaxed whitespace-pre-wrap flex-1 bg-slate-50/30"
                      style={{ fontFamily: 'Arial, sans-serif' }}
                    >
                      {diffSegments
                        .filter((seg) => seg.type !== 'removed')
                        .map((seg, idx) =>
                          seg.type === 'added' ? (
                            <span
                              key={idx}
                              className="bg-emerald-100 text-emerald-950 font-semibold px-0.5 rounded"
                            >
                              {seg.text}
                            </span>
                          ) : (
                            <span key={idx}>{seg.text}</span>
                          )
                        )}
                    </div>
                  </div>
                </div>
              ) : (
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
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

export interface SnapshotsHistoryModalProps {
  isOpen: boolean;
  onClose: () => void;
  revisions: DocumentRevision[];
  currentBodyHtml: string;
  currentTitle: string;
  onTakeManualSnapshot: (label: string) => void;
  onRestoreSnapshot: (rev: DocumentRevision) => void;
  onDeleteSnapshot: (id: string) => void;
  onOpenDiffComparison: (rev: DocumentRevision) => void;
}

export function SnapshotsHistoryModal({
  isOpen,
  onClose,
  revisions,
  currentBodyHtml,
  currentTitle,
  onTakeManualSnapshot,
  onRestoreSnapshot,
  onDeleteSnapshot,
  onOpenDiffComparison,
}: SnapshotsHistoryModalProps) {
  const [snapshotLabel, setSnapshotLabel] = useState<string>('');

  if (!isOpen) return null;

  const currentMetrics = computeDocumentMetrics(currentBodyHtml);

  const handleTakeSnapshot = () => {
    const label = snapshotLabel.trim() || 'حفظ مراجعة يدوية من المحرر';
    onTakeManualSnapshot(label);
    setSnapshotLabel('');
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in-50 duration-200">
        {/* Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-blue-100 text-blue-900 rounded-lg">
              <History className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                سجل المراجعات والتعديلات المفصّل — {currentTitle}
              </h2>
              <div className="text-[11px] text-slate-500">
                جدول المراجعات التلقائية واليدوية (التاريخ · الكاتب · تفاصيل التعديل · مقارنة واستعادة)
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

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-5 space-y-4">
          {/* Quick manual snapshot bar */}
          <div className="border border-blue-200 bg-blue-50/40 rounded-lg p-3 space-y-2">
            <div className="text-xs font-bold text-blue-950 flex items-center justify-between">
              <span>تسجيل مراجعة جديدة في سجل العقد الآن:</span>
              <span className="text-[11px] text-blue-900 tabular-nums font-mono">
                {currentMetrics.wordCount} كلمة · ~{currentMetrics.estimatedPages} صفحة A4
              </span>
            </div>
            <div className="flex items-center gap-2">
              <input
                type="text"
                value={snapshotLabel}
                onChange={(e) => setSnapshotLabel(e.target.value)}
                placeholder="وصف التعديل (مثال: تعديل بند الثمن / مراجعة بيانات البائع)..."
                className="flex-1 px-3 py-1.5 text-xs bg-white border border-slate-300 rounded-lg focus:border-blue-900 focus:outline-none"
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleTakeSnapshot();
                }}
              />
              <button
                type="button"
                onClick={handleTakeSnapshot}
                className="px-3.5 py-1.5 bg-blue-900 text-white text-xs font-semibold rounded-lg hover:bg-blue-800 transition-colors shrink-0 inline-flex items-center gap-1.5"
              >
                <Clock className="w-3.5 h-3.5" />
                <span>تسجيل مراجعة</span>
              </button>
            </div>
          </div>

          {/* Structured Revision History Table (LocalNotaire style: التاريخ | الكاتب | تفاصيل التعديل | إجراءات) */}
          <div className="space-y-2">
            <div className="text-xs font-bold text-slate-800 flex items-center justify-between">
              <span>جدول سجل المراجعات ({revisions.length}):</span>
              <span className="text-[11px] text-slate-500">
                يُسجَّل تلقائياً عند الحفظ وتعديل البنود ودمج المتغيرات
              </span>
            </div>

            {revisions.length === 0 ? (
              <div className="border border-dashed border-slate-200 rounded-lg p-8 text-center text-xs text-slate-500 space-y-1">
                <Clock className="w-6 h-6 text-slate-300 mx-auto mb-1" />
                <div>لا توجد مراجعات مسجلة لهذا العقد بعد.</div>
                <div className="text-[11px] text-slate-400">
                  اضغط «تسجيل مراجعة» بالأعلى أو قم بحفظ/تعديل البنود ليتم توثيق السجل تلقائياً مع اسم الكاتب والتاريخ.
                </div>
              </div>
            ) : (
              <div className="border border-slate-200 rounded-lg overflow-hidden bg-white">
                <table className="w-full text-right border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-100 border-b border-slate-200 text-slate-700 font-bold">
                      <th className="py-2 px-3">التاريخ والوقت</th>
                      <th className="py-2 px-3">الكاتب / المستخدم</th>
                      <th className="py-2 px-3">تفاصيل التعديل</th>
                      <th className="py-2 px-3 text-left">إجراءات المقارنة</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {revisions.map((rev) => {
                      const revMetrics = computeDocumentMetrics(rev.bodyHtml);
                      return (
                        <tr
                          key={rev.id}
                          className="hover:bg-slate-50/80 transition-colors"
                        >
                          <td className="py-2.5 px-3 font-mono text-[11px] text-slate-600 tabular-nums whitespace-nowrap">
                            {new Date(rev.createdAt).toLocaleString('ar-DZ')}
                          </td>
                          <td className="py-2.5 px-3 whitespace-nowrap">
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-blue-50 text-blue-900 border border-blue-200 text-[11px] font-semibold">
                              <UserCheck className="w-3 h-3" />
                              <span>{rev.author || 'محرر المكتب'}</span>
                            </span>
                          </td>
                          <td className="py-2.5 px-3">
                            <div className="font-bold text-slate-900">
                              {rev.summary || 'مراجعة محفوظة'}
                            </div>
                            <div className="text-[10px] text-slate-500 font-mono">
                              {revMetrics.wordCount} كلمة
                            </div>
                          </td>
                          <td className="py-2.5 px-3">
                            <div className="flex items-center justify-end gap-1.5">
                              <button
                                type="button"
                                onClick={() => {
                                  onOpenDiffComparison(rev);
                                  onClose();
                                }}
                                className="px-2.5 py-1 bg-white border border-slate-300 hover:bg-blue-50 hover:border-blue-300 text-blue-950 text-[11px] font-semibold rounded inline-flex items-center gap-1"
                                title="مقارنة هذه النسخة مع النسخة الحالية جنباً إلى جنب"
                              >
                                <GitCompare className="w-3.5 h-3.5 text-blue-900" />
                                <span>مقارنة</span>
                              </button>

                              <button
                                type="button"
                                onClick={() => {
                                  onRestoreSnapshot(rev);
                                  onClose();
                                }}
                                className="px-2 py-1 bg-blue-900 hover:bg-blue-800 text-white text-[11px] font-medium rounded inline-flex items-center gap-1"
                                title="استعادة هذه النسخة إلى المحرر"
                              >
                                <RotateCcw className="w-3 h-3" />
                                <span>استعادة</span>
                              </button>

                              <button
                                type="button"
                                onClick={() => onDeleteSnapshot(rev.id)}
                                className="p-1 text-slate-400 hover:text-red-600 rounded"
                                title="حذف هذه المراجعة"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-1.5 bg-slate-900 text-white text-xs font-medium rounded-lg hover:bg-slate-800 transition-colors"
          >
            إغلاق
          </button>
        </div>
      </div>
    </div>
  );
}
