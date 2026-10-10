import React, { useEffect, useRef, useState } from 'react';
import {
  ArrowLeft,
  ArrowRight,
  BookUser,
  Building2,
  Check,
  CheckCircle2,
  Clock,
  Download,
  Eye,
  FilePlus2,
  FileText,
  FileUp,
  GitCompare,
  HelpCircle,
  History,
  Layers,
  RotateCcw,
  Save,
  Search,
  Sparkles,
  Tag,
  Trash2,
  UserCheck,
  X,
} from 'lucide-react';
import {
  ClauseVariableGroup,
  CustomTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  PartyField,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
  WordTemplateDefinition,
} from '@/lib/types';
import { computeDocumentMetrics, computeTextDiff } from '@/lib/editor-utils';
import { DocxImportResult } from '@/lib/docx-engine';
import { exportTemplateAsDocx, renderTemplateWithContractData } from '@/lib/docx-template-engine';

interface SmartVariablesModalProps {
  isOpen: boolean;
  onClose: () => void;
  focusedVarKey: string | null;
  extractedPlaceholders: string[];
  clauseGroups: ClauseVariableGroup[];
  partyFields: PartyField[];
  fieldValues: Record<string, string>;
  estates: SubdivisionEstate[];
  selectedEstateId: string;
  selectedLotNumber: string;
  savedParties: SavedPartyRecord[];
  savedProperties: SavedPropertyRecord[];
  onSelectEstateAndLot: (estateId: string, lotNumber: string) => void;
  onRecallPartyToRole: (party: SavedPartyRecord, role: 'party1' | 'party2') => void;
  onSaveCurrentPartyToDirectory: (role: 'party1' | 'party2') => void;
  onDeleteSavedParty: (id: string) => void;
  onRecallPropertyRecord: (prop: SavedPropertyRecord) => void;
  onSaveCurrentPropertyToDirectory: () => void;
  onDeleteSavedProperty: (id: string) => void;
  onUpdateFieldValue: (key: string, value: string) => void;
  onChangeFieldInputType: (key: string, inputType: 'text' | 'number' | 'date') => void;
  onBakeAllIntoDocument: (values: Record<string, string>) => void;
}

export function SmartVariablesModal({
  isOpen,
  onClose,
  focusedVarKey,
  extractedPlaceholders,
  clauseGroups,
  partyFields,
  fieldValues,
  estates,
  selectedEstateId,
  selectedLotNumber,
  savedParties,
  savedProperties,
  onSelectEstateAndLot,
  onRecallPartyToRole,
  onSaveCurrentPartyToDirectory,
  onDeleteSavedParty,
  onRecallPropertyRecord,
  onSaveCurrentPropertyToDirectory,
  onDeleteSavedProperty,
  onUpdateFieldValue,
  onChangeFieldInputType,
  onBakeAllIntoDocument,
}: SmartVariablesModalProps) {
  const inputRefs = useRef<Record<string, HTMLInputElement | HTMLSelectElement | null>>({});
  const [selectedPartyIdForRecall, setSelectedPartyIdForRecall] = useState<string>('');
  const [selectedPropIdForRecall, setSelectedPropIdForRecall] = useState<string>('');
  const [showStandardUnusedFields, setShowStandardUnusedFields] = useState<boolean>(true);

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

  const knownMap = new Map(partyFields.map((f) => [f.key, f]));
  const unfilledInDoc = extractedPlaceholders.filter(
    (k) => !fieldValues[k] || !fieldValues[k].trim()
  );

  // Build Clause-Grouped sections from the live contract + optional standard office fields
  const activeDocGroups: ClauseVariableGroup[] =
    clauseGroups.length > 0
      ? clauseGroups
      : extractedPlaceholders.length > 0
      ? [
          {
            clauseId: 'doc_vars',
            clauseTitle: '1. متغيرات العقد الحالي',
            variables: extractedPlaceholders,
          },
        ]
      : [];

  const inDocSet = new Set(extractedPlaceholders);
  const unusedStandardKeys = partyFields
    .map((f) => f.key)
    .filter((k) => !inDocSet.has(k));

  const selectedEstate = estates.find((e) => e.id === selectedEstateId) || estates[0] || null;
  const activePartyToRecall =
    savedParties.find((p) => p.id === selectedPartyIdForRecall) || savedParties[0] || null;
  const activePropToRecall =
    savedProperties.find((p) => p.id === selectedPropIdForRecall) || savedProperties[0] || null;

  const renderVariableCard = (key: string) => {
    const meta = knownMap.get(key);
    const label = meta?.label || key.replace(/_/g, ' ');
    const inCurrentDoc = inDocSet.has(key);
    const val = fieldValues[key] || '';
    const isFilled = val.trim() !== '';
    const isEmptyInDoc = inCurrentDoc && !isFilled;
    const isFocused = focusedVarKey === key;
    const currentType: 'text' | 'number' | 'date' =
      meta?.inputType === 'number'
        ? 'number'
        : meta?.inputType === 'date'
        ? 'date'
        : 'text';

    return (
      <div
        key={key}
        className={`p-2.5 rounded-md border transition-colors ${
          isFocused
            ? 'border-pink-600 bg-pink-50/40 ring-2 ring-pink-500/20'
            : isEmptyInDoc
            ? 'border-amber-300 bg-amber-50/30'
            : isFilled
            ? 'border-emerald-300 bg-emerald-50/20'
            : 'border-slate-200 bg-white'
        }`}
      >
        <div className="flex items-center justify-between gap-1.5 mb-1.5">
          <label className="text-xs font-bold text-slate-800 truncate flex items-center gap-1">
            <span>{label}</span>
            {isFilled && (
              <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 shrink-0" />
            )}
          </label>

          <div className="flex items-center gap-1 shrink-0">
            {/* Field Type Selector: نص / رقم / تاريخ */}
            <div
              className="inline-flex items-center bg-slate-100 border border-slate-200 rounded p-0.5 text-[10px]"
              title="نوع الحقل: نص / رقم / تاريخ"
            >
              <button
                type="button"
                onClick={() => onChangeFieldInputType(key, 'text')}
                className={`px-1.5 py-0.5 rounded transition-colors ${
                  currentType === 'text'
                    ? 'bg-white text-slate-900 font-bold shadow-2xs'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                نص
              </button>
              <button
                type="button"
                onClick={() => onChangeFieldInputType(key, 'number')}
                className={`px-1.5 py-0.5 rounded transition-colors ${
                  currentType === 'number'
                    ? 'bg-white text-blue-900 font-bold shadow-2xs'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                رقم
              </button>
              <button
                type="button"
                onClick={() => onChangeFieldInputType(key, 'date')}
                className={`px-1.5 py-0.5 rounded transition-colors ${
                  currentType === 'date'
                    ? 'bg-white text-blue-900 font-bold shadow-2xs'
                    : 'text-slate-500 hover:text-slate-800'
                }`}
              >
                تاريخ
              </button>
            </div>

            <span
              className={`text-[10px] font-mono px-1.5 py-0.5 rounded shrink-0 ${
                isFilled
                  ? 'text-emerald-800 bg-emerald-100/80'
                  : 'text-pink-800 bg-pink-50'
              }`}
            >
              {`[${key}]`}
            </span>
          </div>
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
            type={
              currentType === 'date'
                ? 'date'
                : currentType === 'number'
                ? 'number'
                : 'text'
            }
            value={val}
            onChange={(e) => onUpdateFieldValue(key, e.target.value)}
            placeholder={
              currentType === 'number'
                ? `أدخل رقماً (${label})...`
                : `أدخل ${label}...`
            }
            className={`w-full px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none ${
              currentType === 'number' ? 'tabular-nums font-mono' : ''
            }`}
          />
        )}
      </div>
    );
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-bold text-slate-900">
                الاستمارة الديناميكية للفقرات والبنود (مرتبة حسب بنود العقد)
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
              تُبنى هذه الاستمارة تلقائياً من الوسوم الموجودة في النص وتُجمّع تحت عنوان كل بند، مع إمكانية اختيار نوع الحقل (نص / رقم / تاريخ).
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
          {/* Quick Directory Bar: Recall / Save Parties & Properties */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {/* 1. Saved Parties Directory Box */}
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-md space-y-2">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-1.5 text-xs font-bold text-slate-900">
                  <BookUser className="w-4 h-4 text-blue-900" />
                  <span>دفتر الأطراف المحفوظين ({savedParties.length})</span>
                </div>
                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => onSaveCurrentPartyToDirectory('party1')}
                    className="px-2 py-0.5 bg-white border border-slate-300 hover:bg-slate-100 text-[10px] font-semibold text-slate-800 rounded inline-flex items-center gap-1"
                    title="حفظ بيانات الطرف الأول الحالية في دفتر الأطراف لاستدعائها لاحقاً"
                  >
                    <Save className="w-3 h-3 text-blue-800" />
                    <span>حفظ الطرف 1</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => onSaveCurrentPartyToDirectory('party2')}
                    className="px-2 py-0.5 bg-white border border-slate-300 hover:bg-slate-100 text-[10px] font-semibold text-slate-800 rounded inline-flex items-center gap-1"
                    title="حفظ بيانات الطرف الثاني الحالية في دفتر الأطراف لاستدعائها لاحقاً"
                  >
                    <Save className="w-3 h-3 text-blue-800" />
                    <span>حفظ الطرف 2</span>
                  </button>
                </div>
              </div>

              {savedParties.length === 0 ? (
                <div className="text-[11px] text-slate-500">
                  املأ بيانات الطرف الأول أو الثاني بالأسفل ثم اضغط «حفظ الطرف» لإضافته للدفتر.
                </div>
              ) : (
                <div className="flex flex-wrap items-center gap-1.5">
                  <select
                    value={activePartyToRecall?.id || ''}
                    onChange={(e) => setSelectedPartyIdForRecall(e.target.value)}
                    className="flex-1 min-w-[140px] px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                  >
                    {savedParties.map((p) => (
                      <option key={p.id} value={p.id}>
                        {p.fullName} {p.birthDate ? `(${p.birthDate})` : ''}
                      </option>
                    ))}
                  </select>
                  {activePartyToRecall && (
                    <>
                      <button
                        type="button"
                        onClick={() => onRecallPartyToRole(activePartyToRecall, 'party1')}
                        className="px-2 py-1 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800"
                      >
                        استدعاء كطرف أول
                      </button>
                      <button
                        type="button"
                        onClick={() => onRecallPartyToRole(activePartyToRecall, 'party2')}
                        className="px-2 py-1 bg-slate-800 text-white text-[11px] font-semibold rounded hover:bg-slate-700"
                      >
                        استدعاء كطرف ثانٍ
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteSavedParty(activePartyToRecall.id)}
                        className="p-1 text-slate-400 hover:text-red-600"
                        title="حذف من دفتر الأطراف"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </>
                  )}
                </div>
              )}
            </div>

            {/* 2. Saved Properties & Subdivision Lots Box */}
            <div className="p-3 bg-blue-50/50 border border-blue-200 rounded-md space-y-2">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-1.5 text-xs font-bold text-blue-950">
                  <Building2 className="w-4 h-4 text-blue-900" />
                  <span>دفتر العقارات وجدول الوصف التقسيمي</span>
                </div>
                <button
                  type="button"
                  onClick={onSaveCurrentPropertyToDirectory}
                  className="px-2 py-0.5 bg-white border border-blue-300 hover:bg-blue-50 text-[10px] font-semibold text-blue-950 rounded inline-flex items-center gap-1"
                  title="حفظ بيانات العقار/الحصة الحالية في دفتر العقارات"
                >
                  <Save className="w-3 h-3 text-blue-800" />
                  <span>حفظ العقار الحالي</span>
                </button>
              </div>

              {savedProperties.length > 0 && (
                <div className="flex items-center gap-1.5">
                  <select
                    value={activePropToRecall?.id || ''}
                    onChange={(e) => setSelectedPropIdForRecall(e.target.value)}
                    className="flex-1 px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                  >
                    {savedProperties.map((pr) => (
                      <option key={pr.id} value={pr.id}>
                        {pr.label} {pr.lotNumber ? `(حصة ${pr.lotNumber})` : ''}
                      </option>
                    ))}
                  </select>
                  {activePropToRecall && (
                    <>
                      <button
                        type="button"
                        onClick={() => onRecallPropertyRecord(activePropToRecall)}
                        className="px-2.5 py-1 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800"
                      >
                        استدعاء العقار
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteSavedProperty(activePropToRecall.id)}
                        className="p-1 text-slate-400 hover:text-red-600"
                        title="حذف من دفتر العقارات"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </>
                  )}
                </div>
              )}

              {estates.length > 0 && (
                <div className="flex flex-wrap items-center gap-1.5 pt-1">
                  <select
                    value={selectedEstate?.id || ''}
                    onChange={(e) => onSelectEstateAndLot(e.target.value, '')}
                    className="px-2 py-1 text-xs bg-white border border-slate-300 rounded"
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
                      className="flex-1 px-2 py-1 text-xs bg-white border border-blue-400 rounded font-semibold tabular-nums"
                    >
                      <option value="">-- اختر حصة من جدول الوصف التقسيمي --</option>
                      {selectedEstate.lots.map((l) => (
                        <option key={l.id} value={l.lotNumber}>
                          حصة {l.lotNumber} — {l.nature} ({l.floor}) — {l.area} م²
                        </option>
                      ))}
                    </select>
                  )}
                </div>
              )}
            </div>
          </div>

          {/* SECTION 1: DYNAMIC VARIABLES GROUPED BY CLAUSE IN THE ACTIVE CONTRACT */}
          {activeDocGroups.length > 0 ? (
            <div className="space-y-4">
              {activeDocGroups.map((group) => {
                const unfilledInClause = group.variables.filter(
                  (v) => !fieldValues[v] || !fieldValues[v].trim()
                ).length;
                return (
                  <div
                    key={group.clauseId}
                    className="border border-slate-200 rounded-lg overflow-hidden bg-slate-50/40"
                  >
                    <div className="px-3.5 py-2 bg-slate-100 border-b border-slate-200 flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-900">
                        {group.clauseTitle}
                      </span>
                      <span
                        className={`text-[10px] font-semibold px-2 py-0.5 rounded tabular-nums ${
                          unfilledInClause > 0
                            ? 'bg-amber-100 text-amber-900'
                            : 'bg-emerald-100 text-emerald-900'
                        }`}
                      >
                        {unfilledInClause > 0
                          ? `متبقي ${unfilledInClause} من ${group.variables.length}`
                          : `مكتمل (${group.variables.length}) ✓`}
                      </span>
                    </div>
                    <div className="p-3 grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {group.variables.map((vKey) => renderVariableCard(vKey))}
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="p-4 bg-amber-50/60 border border-amber-200 rounded-md text-xs text-amber-950">
              لا توجد وسوم <code className="font-mono">{`[...]`}</code> مدرجة في نص العقد الحالي بعد. يمكنك تحديد أي كلمة في المحرر والضغط على زر <strong>«[ ] تحويل المحدد لوسم»</strong> أو تعبئة حقول المكتب القياسية أدناه لتوليد الوثائق المشتقة.
            </div>
          )}

          {/* SECTION 2: STANDARD OFFICE FIELDS FOR DERIVED DOCUMENTS */}
          {unusedStandardKeys.length > 0 && (
            <div className="border border-slate-200 rounded-lg overflow-hidden">
              <button
                type="button"
                onClick={() => setShowStandardUnusedFields((v) => !v)}
                className="w-full px-3.5 py-2 bg-slate-50 hover:bg-slate-100 border-b border-slate-200 flex items-center justify-between text-xs font-bold text-slate-700"
              >
                <span>
                  حقول المكتب القياسية الإضافية للوثائق المشتقة ({unusedStandardKeys.length})
                </span>
                <span className="text-[11px] text-blue-900">
                  {showStandardUnusedFields ? 'إخفاء ▲' : 'إظهار ▼'}
                </span>
              </button>
              {showStandardUnusedFields && (
                <div className="p-3 grid grid-cols-1 sm:grid-cols-2 gap-3 bg-white">
                  {unusedStandardKeys.map((k) => renderVariableCard(k))}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Footer Actions */}
        <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-2">
          <button
            type="button"
            onClick={() => {
              const liveValues: Record<string, string> = { ...fieldValues };
              for (const [k, el] of Object.entries(inputRefs.current)) {
                if (el && typeof el.value === 'string' && el.value.trim() !== '') {
                  liveValues[k] = el.value;
                  onUpdateFieldValue(k, el.value);
                }
              }
              onBakeAllIntoDocument(liveValues);
              onClose();
            }}
            className="px-3.5 py-1.5 bg-pink-800 text-white text-xs font-medium rounded hover:bg-pink-900 transition-colors"
            title="يستبدل كل وسم [...] داخل ورقة الـ A4 بقيمته المكتوبة مع الحفاظ على التنسيق المحيط"
          >
            استبدال الوسوم نهائياً داخل نص العقد
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-1.5 bg-blue-900 text-white text-xs font-semibold rounded hover:bg-blue-800 transition-colors"
            >
              حفظ واعتماد القيم (مع التلوين الأخضر للوسوم المعبأة)
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
  const [selectedLeftRevId, setSelectedLeftRevId] = useState<string>('');
  const [rightRevId, setRightRevId] = useState<string>('CURRENT');

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
              onChange={(e) => setSelectedLeftRevId(e.target.value)}
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
  onStartBlank: (clearPreviousClauses?: boolean) => void;
  onSelectTemplate: (
    tpl: CustomTemplate,
    mode: 'replace' | 'insert',
    clearPreviousClauses: boolean
  ) => void;
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
  const [templateLoadMode, setTemplateLoadMode] = useState<'replace' | 'insert'>('replace');
  const [clearPreviousClauses, setClearPreviousClauses] = useState<boolean>(true);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-lg border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[85vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <h2 className="text-sm font-bold text-slate-900">
            بدء تحرير عقد جديد / فتح قالب (مع خيارات الإدراج والبنود)
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
            <div className="space-y-4">
              <label className="flex items-center gap-2 text-xs text-slate-700 bg-slate-50 p-2.5 rounded border border-slate-200 cursor-pointer">
                <input
                  type="checkbox"
                  checked={clearPreviousClauses}
                  onChange={(e) => setClearPreviousClauses(e.target.checked)}
                  className="rounded border-slate-300 text-blue-900"
                />
                <span className="font-medium">
                  حذف البنود السابقة المفعّلة عند بدء ورقة عقد جديدة
                </span>
              </label>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <button
                  type="button"
                  onClick={() => {
                    onStartBlank(clearPreviousClauses);
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
            </div>
          )}

          {sourceTab === 'templates' && (
            <div className="space-y-3">
              {/* Template Load Options Bar */}
              <div className="p-3 bg-blue-50/60 border border-blue-200 rounded-md space-y-2">
                <div className="text-xs font-bold text-blue-950">
                  خيار فتح القالب / النموذج:
                </div>
                <div className="flex flex-wrap items-center gap-4 text-xs">
                  <label className="inline-flex items-center gap-1.5 cursor-pointer font-medium text-slate-800">
                    <input
                      type="radio"
                      name="tplLoadMode"
                      checked={templateLoadMode === 'replace'}
                      onChange={() => setTemplateLoadMode('replace')}
                    />
                    <span>استبدال المحتوى الحالي (يمسح ورقة الـ A4 ويضع القالب كاملاً)</span>
                  </label>
                  <label className="inline-flex items-center gap-1.5 cursor-pointer font-medium text-slate-800">
                    <input
                      type="radio"
                      name="tplLoadMode"
                      checked={templateLoadMode === 'insert'}
                      onChange={() => setTemplateLoadMode('insert')}
                    />
                    <span>إدراج عند موضع المؤشر (يُبقي النص الحالي)</span>
                  </label>
                </div>

                {templateLoadMode === 'replace' && (
                  <label className="inline-flex items-center gap-2 text-xs text-slate-700 pt-1 cursor-pointer">
                    <input
                      type="checkbox"
                      checked={clearPreviousClauses}
                      onChange={(e) => setClearPreviousClauses(e.target.checked)}
                      className="rounded border-slate-300 text-blue-900"
                    />
                    <span className="font-semibold text-blue-950">
                      حذف البنود السابقة (تفريغ قائمة البنود المفعّلة عند الاستبدال الكامل)
                    </span>
                  </label>
                )}
              </div>

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
                        onSelectTemplate(tpl, templateLoadMode, clearPreviousClauses);
                        onClose();
                      }}
                      className="px-3 py-1.5 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800"
                    >
                      {templateLoadMode === 'replace'
                        ? 'استبدال وفتح القالب'
                        : 'إدراج القالب عند المؤشر'}
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

// =========================================================
// 1. WORD IMPORT PREVIEW MODAL (معاينة استيراد ملف Word)
// =========================================================
interface DocxImportPreviewModalProps {
  isOpen: boolean;
  onClose: () => void;
  importResult: DocxImportResult | null;
  onConfirmApply: (mode: 'replace' | 'insert') => void;
}

export function DocxImportPreviewModal({
  isOpen,
  onClose,
  importResult,
  onConfirmApply,
}: DocxImportPreviewModalProps) {
  const [importMode, setImportMode] = useState<'replace' | 'insert'>('replace');

  if (!isOpen || !importResult) return null;

  const plainSnippet = importResult.bodyHtml
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 480);

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in-50 duration-200">
        {/* Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-blue-100 text-blue-900 rounded-lg">
              <FileUp className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                معاينة ملف Word (.docx) قبل التطبيق
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
            <div className="p-2.5 rounded-lg border border-slate-200 bg-slate-50/60 text-center">
              <div className="text-[11px] text-slate-500 font-medium">الجداول</div>
              <div className="text-base font-bold text-slate-900 tabular-nums">
                {importResult.stats.tableCount}
              </div>
            </div>
            <div className="p-2.5 rounded-lg border border-blue-200 bg-blue-50/50 text-center">
              <div className="text-[11px] text-blue-900 font-medium">الوسوم الذكية</div>
              <div className="text-base font-bold text-blue-900 tabular-nums">
                {importResult.extractedPlaceholders.length}
              </div>
            </div>
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

          {/* Import Destination Options */}
          <div className="border border-slate-200 rounded-lg p-3.5 bg-white space-y-2.5">
            <div className="text-xs font-bold text-slate-900">
              طريقة تطبيق المحتوى المستورد على ورقة المحرر:
            </div>

            <div className="space-y-2">
              <label
                className={`flex items-start gap-3 p-3 rounded-lg border cursor-pointer transition-colors ${
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
                    استبدال محتوى الورقة الحالية بالكامل (الافتراضي الموصى به)
                  </div>
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    يتم أخذ لقطة زمنية تلقائية (Snapshot) من العقد الحالي لحمايته من الفقدان، وتوحيد العقد المستورد إلى خط Arial 13pt وهوامش 7/2/1/6 سم.
                  </div>
                </div>
              </label>

              <label
                className={`flex items-start gap-3 p-3 rounded-lg border cursor-pointer transition-colors ${
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
                  <div className="text-[11px] text-slate-500 mt-0.5">
                    يُبقي على محتوى العقد المفتوح ويدمج فقرات وجداول الملف المستورد في الموضع الذي يقف عنده المؤشر.
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
              onConfirmApply(importMode);
              onClose();
            }}
            className="px-4 py-2 bg-blue-900 text-white text-xs font-bold rounded-lg hover:bg-blue-800 transition-colors shadow-2xs inline-flex items-center gap-1.5"
          >
            <Check className="w-4 h-4" />
            <span>
              {importMode === 'replace'
                ? 'استبدال وفتح العقد بالمحرر'
                : 'إدراج المستورد عند المؤشر'}
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
                يُخزن محلياً في IndexedDB للرجوع إليه وتطبيقه على أي عقد جديد
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
            <label className="text-xs font-bold text-slate-700">وصف أو ملاحظات للاستخدام (اختياري)</label>
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
// 3. SNAPSHOTS HISTORY MODAL (سجل اللقطات الزمنية والمقارنة)
// =========================================================
interface SnapshotsHistoryModalProps {
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
    const label = snapshotLabel.trim() || 'لقطة يدوية مسجلة من الكاتب';
    onTakeManualSnapshot(label);
    setSnapshotLabel('');
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-slate-200 shadow-2xl w-full max-w-2xl max-h-[90vh] flex flex-col overflow-hidden animate-in fade-in-50 duration-200">
        {/* Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-blue-100 text-blue-900 rounded-lg">
              <History className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-slate-900">
                سجل اللقطات الزمنية والنسخ المرجعية للعقد
              </h2>
              <div className="text-[11px] text-slate-500">
                اللقطات التلقائية المحفوظة قبل الاستبدال والتحويل + اللقطات اليدوية
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
              <span>أخذ لقطة زمنية فورية للعقد الحالي الآن:</span>
              <span className="text-[11px] text-blue-900 tabular-nums font-mono">
                {currentMetrics.wordCount} كلمة · ~{currentMetrics.estimatedPages} صفحة A4
              </span>
            </div>
            <div className="flex items-center gap-2">
              <input
                type="text"
                value={snapshotLabel}
                onChange={(e) => setSnapshotLabel(e.target.value)}
                placeholder="تسمية اللقطة (مثال: قبل مراجعة الثمن / بعد إضافة الأطراف)..."
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
                <span>تسجيل اللقطة</span>
              </button>
            </div>
          </div>

          {/* List of snapshots */}
          <div className="space-y-2">
            <div className="text-xs font-bold text-slate-800 flex items-center justify-between">
              <span>سجل اللقطات المحفوظة ({revisions.length}):</span>
              <span className="text-[11px] text-slate-500">
                الأحدث أولاً (تخزين محلي IndexedDB)
              </span>
            </div>

            {revisions.length === 0 ? (
              <div className="border border-dashed border-slate-200 rounded-lg p-8 text-center text-xs text-slate-500 space-y-1">
                <Clock className="w-6 h-6 text-slate-300 mx-auto mb-1" />
                <div>لا توجد لقطات زمنية مسجلة بعد.</div>
                <div className="text-[11px] text-slate-400">
                  يتم أخذ لقطات تلقائية عند استبدال الوسوم أو استيراد ملفات Word، كما يمكنك تسجيل لقطات يدوياً في أي وقت.
                </div>
              </div>
            ) : (
              <div className="space-y-2">
                {revisions.map((rev) => {
                  const revMetrics = computeDocumentMetrics(rev.bodyHtml);
                  return (
                    <div
                      key={rev.id}
                      className="p-3 border border-slate-200 rounded-lg bg-white hover:bg-slate-50/60 transition-colors flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3"
                    >
                      <div className="space-y-0.5">
                        <div className="text-xs font-bold text-slate-900 flex items-center gap-2">
                          <span>{rev.summary || 'نسخة مرجعية'}</span>
                          <span className="text-[10px] px-1.5 py-0.2 bg-slate-100 text-slate-600 rounded font-mono">
                            {revMetrics.wordCount} كلمة
                          </span>
                        </div>
                        <div className="text-[11px] text-slate-500 tabular-nums">
                          {new Date(rev.createdAt).toLocaleString('ar-DZ')} · {rev.documentTitle || currentTitle}
                        </div>
                      </div>

                      <div className="flex items-center gap-1.5 shrink-0 self-end sm:self-auto">
                        <button
                          type="button"
                          onClick={() => {
                            onOpenDiffComparison(rev);
                            onClose();
                          }}
                          className="px-2.5 py-1 bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 text-xs font-medium rounded inline-flex items-center gap-1"
                          title="مقارنة الفروق بين هذه النسخة والمسودة الحالية"
                        >
                          <GitCompare className="w-3.5 h-3.5 text-blue-900" />
                          <span>مقارنة (Diff)</span>
                        </button>

                        <button
                          type="button"
                          onClick={() => {
                            if (window.confirm('هل أنت متأكد من استعادة هذه النسخة؟ سيتم أخذ لقطة تلقائية للمسودة الحالية أولاً.')) {
                              onRestoreSnapshot(rev);
                              onClose();
                            }
                          }}
                          className="px-2.5 py-1 bg-blue-900 hover:bg-blue-800 text-white text-xs font-medium rounded inline-flex items-center gap-1"
                          title="استرجاع نص هذه النسخة إلى الورقة مباشرة"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                          <span>استعادة</span>
                        </button>

                        <button
                          type="button"
                          onClick={() => onDeleteSnapshot(rev.id)}
                          className="p-1 text-slate-400 hover:text-red-600 rounded"
                          title="حذف هذه النسخة من السجل"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  );
                })}
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

// =========================================================
// 4. ONBOARDING TOUR MODAL (جولة إرشادية خفيفة من 5 خطوات للمحرر)
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
    <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-2xs flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-2xl border border-slate-200 shadow-2xl w-full max-w-xl overflow-hidden animate-in fade-in-50 duration-200 flex flex-col">
        {/* Header Strip with step indicators */}
        <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="text-xs font-bold text-slate-900">
              دليل محرر العقود التوثيقية
            </div>
            <span className="text-[11px] px-2 py-0.5 bg-blue-100 text-blue-900 rounded-full font-bold tabular-nums">
              خطوة {currentStep + 1} من {tourSteps.length}
            </span>
          </div>

          <button
            type="button"
            onClick={() => {
              onFinishTour();
              onClose();
            }}
            className="text-xs text-slate-500 hover:text-slate-800 font-medium"
          >
            تخطي الجولة
          </button>
        </div>

        {/* Step Progress Line */}
        <div className="w-full bg-slate-100 h-1">
          <div
            className="bg-blue-900 h-1 transition-all duration-300"
            style={{ width: `${((currentStep + 1) / tourSteps.length) * 100}%` }}
          />
        </div>

        {/* Step Content */}
        <div className="p-6 space-y-4">
          <div className="flex items-center gap-3">
            <div className="p-3 bg-slate-100 rounded-xl shrink-0">
              {current.icon}
            </div>
            <div>
              <span className="inline-block text-[11px] font-bold text-blue-900 bg-blue-50 px-2 py-0.5 rounded mb-1">
                {current.badgeText}
              </span>
              <h3 className="text-base font-bold text-slate-900 leading-snug">
                {current.title}
              </h3>
            </div>
          </div>

          <p className="text-xs text-slate-600 leading-relaxed text-justify">
            {current.description}
          </p>

          <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 space-y-1.5">
            <div className="text-[11px] font-bold text-slate-800">
              إضاءات وميزات سريعة:
            </div>
            {current.tips.map((tip, idx) => (
              <div key={idx} className="text-[11px] text-slate-600 flex items-start gap-1.5">
                <span className="text-blue-900 font-bold shrink-0">✓</span>
                <span>{tip}</span>
              </div>
            ))}
          </div>

          {/* Stepper Dots */}
          <div className="flex items-center justify-center gap-1.5 pt-2">
            {tourSteps.map((_, idx) => (
              <button
                key={idx}
                type="button"
                onClick={() => setCurrentStep(idx)}
                className={`h-2 rounded-full transition-all ${
                  idx === currentStep ? 'w-6 bg-blue-900' : 'w-2 bg-slate-200 hover:bg-slate-300'
                }`}
                title={`الخطوة ${idx + 1}`}
              />
            ))}
          </div>
        </div>

        {/* Footer Actions */}
        <div className="px-6 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
          <button
            type="button"
            onClick={handlePrev}
            disabled={isFirst}
            className="px-3.5 py-1.5 text-xs text-slate-700 hover:text-slate-900 font-medium disabled:opacity-30 inline-flex items-center gap-1"
          >
            <ArrowRight className="w-3.5 h-3.5" />
            <span>السابق</span>
          </button>

          <button
            type="button"
            onClick={handleNext}
            className="px-5 py-2 bg-blue-900 text-white text-xs font-bold rounded-lg hover:bg-blue-800 transition-colors shadow-2xs inline-flex items-center gap-1.5"
          >
            <span>{isLast ? 'إنهاء وبدء التحرير' : 'التالي'}</span>
            {!isLast && <ArrowLeft className="w-3.5 h-3.5" />}
          </button>
        </div>
      </div>
    </div>
  );
}

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
    fieldValues?: Record<string, string>;
  };
}

export function WordTemplatesModal({
  isOpen,
  onClose,
  wordTemplates,
  onSaveWordTemplate,
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
        {/* Header */}
        <div className="px-5 py-3.5 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div>
            <h2 className="text-sm font-bold text-slate-900">
              إدارة قوالب Word الرسمية وتوليد الوثائق المشتقة (أصل، مستخرج، شهر، وضعية جبائية)
            </h2>
            <p className="text-[11px] text-slate-500 mt-0.5">
              فصل تامة بين نص العقد (المحتوى والبنود) وشكل وثيقة Word الرسمية (القوالب، الهوامش، العناوين).
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

        {/* Templates List Grid */}
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
                      <span>معيـنة سريعة</span>
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

        {/* Quick Preview Sub-Modal */}
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

        {/* Footer */}
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
