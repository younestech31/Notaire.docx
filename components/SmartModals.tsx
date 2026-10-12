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
  Columns2,
  Merge,
  Scissors,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import {
  ClauseVariableGroup,
  CustomTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  OfficeTemplate,
  PartyField,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
  VariableInputType,
  WordTemplateDefinition,
} from '@/lib/types';
import { computeTextDiff } from '@/lib/editor-utils';
import { computeDocumentMetrics } from '@/lib/notarial-pagination';
import {
  CandidateImportedClause,
  compileCandidateClausesToHtml,
  segmentHtmlIntoCandidateClauses,
} from '@/lib/clause-outline-engine';
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
  fieldInputTypes?: Record<string, VariableInputType>;
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
  fieldInputTypes = {},
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

  // Build Clause-Grouped sections strictly from the live contract's [...] placeholders
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
    const perContractType = fieldInputTypes[key];
    const resolvedType = perContractType || meta?.inputType;
    const currentType: 'text' | 'number' | 'date' =
      resolvedType === 'number'
        ? 'number'
        : resolvedType === 'date'
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
                    <div className="p-3 space-y-3">
                      {group.clauseReferenceText && (
                        <div
                          dir="rtl"
                          className="p-2.5 bg-white border border-slate-200/90 rounded-md text-xs text-slate-700 leading-relaxed"
                          style={{ fontFamily: 'Arial, sans-serif' }}
                        >
                          <div className="text-[10px] font-bold text-blue-900 mb-1 flex items-center gap-1">
                            <FileText className="w-3 h-3" />
                            <span>النص المرجعي الكامل للبند:</span>
                          </div>
                          <p className="text-justify">{group.clauseReferenceText}</p>
                        </div>
                      )}
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        {group.variables.map((vKey) => renderVariableCard(vKey))}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="p-4 bg-amber-50/60 border border-amber-200 rounded-md text-xs text-amber-950">
              لا توجد وسوم <code className="font-mono">{`[...]`}</code> مدرجة في نص العقد المفتوح حالياً بعد. يمكنك كتابة أي وسم مثل <code className="font-mono">{`[البائع]`}</code> أو تحديد أي كلمة في المحرر والضغط على زر <strong>«[ ] تحويل المحدد لوسم»</strong> ليظهر الحقل هنا فوراً دون أي حقول دخيلة من عقود أخرى.
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

export { VersionDiffModal, SnapshotsHistoryModal } from './RevisionsDiffModal';


interface MultiSourceStartModalProps {
  isOpen: boolean;
  onClose: () => void;
  templates: (OfficeTemplate | CustomTemplate)[];
  documents: SavedDocument[];
  downloads?: DownloadArchiveItem[];
  onStartBlank: (clearPreviousClauses?: boolean) => void;
  onSelectTemplate: (
    tpl: OfficeTemplate | CustomTemplate,
    mode: 'replace' | 'insert',
    clearPreviousClauses: boolean
  ) => void;
  onSelectSavedDoc: (doc: SavedDocument) => void;
  onImportDocxFile: (file: File) => void;
  onSelectDownloadArchiveItem?: (item: DownloadArchiveItem) => void;
}

export function MultiSourceStartModal({
  isOpen,
  onClose,
  templates,
  documents,
  onStartBlank,
  onSelectTemplate,
  onSelectSavedDoc,
  onImportDocxFile,
}: MultiSourceStartModalProps) {
  const [sourceTab, setSourceTab] = useState<'blank' | 'templates' | 'docs'>('blank');
  const [templateLoadMode, setTemplateLoadMode] = useState<'replace' | 'insert'>('replace');
  const [clearPreviousClauses, setClearPreviousClauses] = useState<boolean>(true);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-[#0F2744]/55 flex items-center justify-center p-4 select-none">
      <div className="bg-white rounded-xl border border-[#C4A46A]/45 shadow-2xl w-full max-w-2xl max-h-[85vh] flex flex-col overflow-hidden">
        <div className="px-5 py-3.5 bg-[#F7F5F0] border-b border-[#C4A46A]/35 flex items-center justify-between gap-3">
          <div className="flex flex-col sm:flex-row sm:items-center gap-2 sm:gap-3">
            <img
              src="/brand/logo-lockup.svg"
              alt="الموثق الرقمي — محرر العقود التوثيقية"
              className="h-9 w-auto select-none shrink-0"
            />
            <span className="hidden sm:inline-block h-4 w-px bg-[#C4A46A]/50" />
            <h2 className="text-xs sm:text-sm font-bold text-[#0F2744]">
              بدء تحرير عقد جديد / فتح قالب
            </h2>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="p-1.5 text-[#5C6B7A] hover:text-[#0F2744] rounded"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Source Tabs */}
        <div className="grid grid-cols-3 gap-1 p-2 bg-slate-100 border-b border-slate-200 text-xs font-medium">
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
            العقود المحفوظة ({documents.length})
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
                        {tpl.category} · {(tpl.extractedPlaceholders || []).length} وسم ذكي
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
        </div>
      </div>
    </div>
  );
}

export {
  DocxImportPreviewModal,
  SaveAsTemplateModal,
  OnboardingTourModal,
  WordTemplatesModal,
} from './modals/DocumentWorkflowModals';

