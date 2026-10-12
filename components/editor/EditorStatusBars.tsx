'use client';

import React from 'react';
import {
  ChevronDown,
  ChevronUp,
  Lock,
  Replace,
  Search,
  ShieldCheck,
  Unlock,
  X,
} from 'lucide-react';
import {
  NotaryClause,
  OfficeTemplate,
  SavedDocument,
} from '@/lib/types';
import { isContractOfficiallyIndexed } from '@/lib/indexing-engine';

export function EditorModeAndSealBanners({
  editingDerivedTpl,
  editingClauseObj,
  activeDocument,
  unlockIndexedBody,
  onToggleUnlockIndexedBody,
  onFinishEditingDerivedTemplate,
  onFinishEditingClause,
}: {
  editingDerivedTpl: OfficeTemplate | null;
  editingClauseObj: NotaryClause | null;
  activeDocument: SavedDocument | null;
  unlockIndexedBody: boolean;
  onToggleUnlockIndexedBody: () => void;
  onFinishEditingDerivedTemplate: (saveChanges: boolean) => void;
  onFinishEditingClause: (saveChanges: boolean) => void;
}) {
  return (
    <>
      {/* Contextual Banner when Editing an Office Template or Clause inside the A4 Editor */}
      {(editingDerivedTpl || editingClauseObj) && (
        <div className="bg-amber-100 border-b border-amber-300 px-6 py-2 flex flex-wrap items-center justify-between gap-3 no-print">
          <div className="text-xs font-bold text-amber-950">
            {editingDerivedTpl
              ? `وضع تعديل قالب المكتب: "${editingDerivedTpl.name}" — يمكنك تعديل النص والجداول ووسوم [...] بحرية تامة.`
              : `وضع تعديل صياغة البند الجاهز: "${editingClauseObj?.title}"`}
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() =>
                editingDerivedTpl
                  ? onFinishEditingDerivedTemplate(true)
                  : onFinishEditingClause(true)
              }
              className="px-3 py-1 bg-amber-900 text-white text-xs font-semibold rounded hover:bg-amber-950"
            >
              حفظ التعديلات والعودة للعقد
            </button>
            <button
              type="button"
              onClick={() =>
                editingDerivedTpl
                  ? onFinishEditingDerivedTemplate(false)
                  : onFinishEditingClause(false)
              }
              className="px-3 py-1 bg-white border border-amber-300 text-amber-950 text-xs font-medium rounded hover:bg-amber-50"
            >
              إلغاء والعودة للعقد
            </button>
          </div>
        </div>
      )}

      {/* Official Seal & Indexing Protection Banner */}
      {isContractOfficiallyIndexed(activeDocument) &&
        !editingDerivedTpl &&
        !editingClauseObj && (
          <div className="bg-emerald-950 text-white border-b border-emerald-800 px-6 py-2 flex flex-wrap items-center justify-between gap-3 no-print">
            <div className="flex items-center gap-2 text-xs">
              <ShieldCheck className="w-4 h-4 text-emerald-400 shrink-0" />
              <span className="font-bold">
                عقد مفهرس ومختوم رسمياً — فهرس رقم:{' '}
                <span className="font-mono tabular-nums text-emerald-300">
                  {activeDocument?.indexNumber}
                </span>
                {activeDocument?.depositNumber ? (
                  <>
                    {' '}
                    · إيداع رقم:{' '}
                    <span className="font-mono tabular-nums text-emerald-300">
                      {activeDocument.depositNumber}
                    </span>
                  </>
                ) : null}
              </span>
              <span className="text-emerald-200/90 hidden md:inline">
                — المتن الأساسي مقفول ومحمي بعد الختم (الكتابة واللصق وتعديل البنود معطلة)، بينما تظل مراجع التسجيل والشهر والوثائق المشتقة متاحة للتحديث.
              </span>
            </div>
            <button
              type="button"
              onClick={onToggleUnlockIndexedBody}
              className={`px-2.5 py-1 rounded text-[11px] font-semibold inline-flex items-center gap-1.5 transition-colors ${
                unlockIndexedBody
                  ? 'bg-amber-500 text-slate-950 hover:bg-amber-400'
                  : 'bg-emerald-900 hover:bg-emerald-800 text-emerald-100 border border-emerald-700'
              }`}
            >
              {unlockIndexedBody ? (
                <>
                  <Lock className="w-3.5 h-3.5" />
                  <span>إعادة قفل المتن المختوم</span>
                </>
              ) : (
                <>
                  <Unlock className="w-3.5 h-3.5" />
                  <span>فتح قفل المتن استثنائياً</span>
                </>
              )}
            </button>
          </div>
        )}
    </>
  );
}

export function EditorFindReplaceBar({
  findQuery,
  replaceQuery,
  ignoreArabicHamzaAndDiacritics,
  matchCount,
  activeMatchIdx,
  isBodyLocked,
  onChangeFindQuery,
  onChangeReplaceQuery,
  onToggleIgnoreHamza,
  onPrevMatch,
  onNextMatch,
  onRunReplace,
  onClose,
}: {
  findQuery: string;
  replaceQuery: string;
  ignoreArabicHamzaAndDiacritics: boolean;
  matchCount: number;
  activeMatchIdx: number;
  isBodyLocked: boolean;
  onChangeFindQuery: (val: string) => void;
  onChangeReplaceQuery: (val: string) => void;
  onToggleIgnoreHamza: (checked: boolean) => void;
  onPrevMatch: () => void;
  onNextMatch: () => void;
  onRunReplace: (replaceAll: boolean) => void;
  onClose: () => void;
}) {
  return (
    <div className="bg-amber-50/95 border-b border-amber-200 px-4 py-2 flex flex-wrap items-center gap-2.5 text-xs no-print">
      <div className="flex items-center gap-1.5 font-semibold text-amber-950 shrink-0">
        <Search className="w-3.5 h-3.5 text-amber-900" />
        <span>بحث واستبدال عربي:</span>
      </div>
      <input
        type="text"
        value={findQuery}
        onChange={(e) => onChangeFindQuery(e.target.value)}
        placeholder="ابحث عن نص (مثال: البائع)..."
        className="px-2.5 py-1 bg-white border border-slate-300 rounded w-44 focus:border-blue-900 focus:outline-none"
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            if (e.shiftKey) onPrevMatch();
            else onNextMatch();
          } else if (e.key === 'Escape') {
            onClose();
          }
        }}
      />
      <input
        type="text"
        value={replaceQuery}
        onChange={(e) => onChangeReplaceQuery(e.target.value)}
        placeholder="استبدال بـ..."
        disabled={isBodyLocked}
        className="px-2.5 py-1 bg-white border border-slate-300 rounded w-44 focus:border-blue-900 focus:outline-none disabled:opacity-50"
        onKeyDown={(e) => {
          if (e.key === 'Enter' && !isBodyLocked) onRunReplace(false);
          else if (e.key === 'Escape') onClose();
        }}
      />

      <label className="flex items-center gap-1.5 text-[11px] text-amber-950 bg-white/70 px-2 py-1 rounded border border-amber-200 cursor-pointer select-none">
        <input
          type="checkbox"
          checked={ignoreArabicHamzaAndDiacritics}
          onChange={(e) => onToggleIgnoreHamza(e.target.checked)}
          className="rounded border-slate-300 text-blue-900"
        />
        <span>تجاهل الهمزات والتشكيل (أ/إ/آ، ة/ه)</span>
      </label>

      <span className="text-slate-600 tabular-nums text-xs">
        {matchCount > 0 ? `${activeMatchIdx + 1} من ${matchCount}` : '0 نتائج'}
      </span>

      <div className="flex items-center gap-1">
        <button
          type="button"
          onClick={onPrevMatch}
          disabled={matchCount === 0}
          className="p-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40"
          title="التطابق السابق (Shift+Enter)"
        >
          <ChevronUp className="w-3.5 h-3.5" />
        </button>
        <button
          type="button"
          onClick={onNextMatch}
          disabled={matchCount === 0}
          className="p-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40"
          title="التطابق التالي (Enter)"
        >
          <ChevronDown className="w-3.5 h-3.5" />
        </button>
      </div>

      <button
        type="button"
        onClick={() => onRunReplace(false)}
        disabled={matchCount === 0 || isBodyLocked}
        className="px-2.5 py-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 inline-flex items-center gap-1"
      >
        <Replace className="w-3.5 h-3.5" />
        <span>استبدال الحالي</span>
      </button>
      <button
        type="button"
        onClick={() => onRunReplace(true)}
        disabled={matchCount === 0 || isBodyLocked}
        className="px-2.5 py-1 bg-blue-900 text-white rounded hover:bg-blue-800 disabled:opacity-40 font-medium"
        title="استبدال كل المطابقات مع أخذ لقطة أمان تلقائية"
      >
        استبدال الكل ({matchCount})
      </button>
      <button
        type="button"
        onClick={onClose}
        className="mr-auto p-1 text-slate-500 hover:text-slate-800"
        title="إغلاق (Esc)"
      >
        <X className="w-4 h-4" />
      </button>
    </div>
  );
}

export function EditorBottomStatusBar({
  autoSaveState,
  unfilledCount,
  totalVariablesCount,
  docMetrics,
  onOpenSmartVariablesModal,
}: {
  autoSaveState: 'saved' | 'saving' | 'error';
  unfilledCount: number;
  totalVariablesCount: number;
  docMetrics: {
    wordCount: number;
    charCount: number;
    estimatedPages: number;
  };
  onOpenSmartVariablesModal: () => void;
}) {
  return (
    <footer className="bg-white border-t border-[#C4A46A]/35 px-4 py-1.5 text-[11px] text-[#5C6B7A] flex flex-wrap items-center justify-between gap-4 select-none no-print">
      <div className="flex items-center gap-3">
        <span className="inline-flex items-center gap-1.5 font-medium text-[#0F2744]">
          <span
            className={`w-2 h-2 rounded-full ${
              autoSaveState === 'saved'
                ? 'bg-emerald-500'
                : autoSaveState === 'error'
                ? 'bg-[#8C1D2C]'
                : 'bg-[#C4A46A] animate-pulse'
            }`}
          />
          {autoSaveState === 'saved'
            ? 'تم الحفظ تلقائياً في قاعدة المكتب'
            : autoSaveState === 'error'
            ? 'تعذر الحفظ في مخزن المكتب — راجع التنبيه'
            : 'جارٍ الحفظ التلقائي...'}
        </span>

        <span className="text-slate-300">|</span>

        <button
          type="button"
          onClick={onOpenSmartVariablesModal}
          className={`font-semibold hover:underline ${
            unfilledCount > 0 ? 'text-[#8C1D2C]' : 'text-emerald-700'
          }`}
        >
          {unfilledCount > 0
            ? `المتغيرات غير المعبأة: ${unfilledCount} (اضغط للتعبئة)`
            : `جميع متغيرات العقد مكتملة (${totalVariablesCount})`}
        </button>
      </div>

      <div className="flex items-center gap-3 tabular-nums">
        <span>الكلمات: {docMetrics.wordCount}</span>
        <span className="text-slate-300">·</span>
        <span>الأحرف: {docMetrics.charCount}</span>
        <span className="text-slate-300">·</span>
        <span>الصفحات المقدرة: {docMetrics.estimatedPages}</span>
        <span className="text-slate-300">·</span>
        <span className="font-semibold text-[#0F2744]">
          Arial 13pt · تباعد 1.0 · هوامش 7/2/1/6 سم
        </span>
      </div>
    </footer>
  );
}
