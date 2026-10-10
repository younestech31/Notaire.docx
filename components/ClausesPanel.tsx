'use client';

import React, { useState } from 'react';
import {
  ArrowDown,
  ArrowUp,
  BookmarkPlus,
  Check,
  CheckSquare,
  ChevronDown,
  ChevronUp,
  Edit3,
  GripVertical,
  Hash,
  Lock,
  Plus,
  Sparkles,
  Square,
  Trash2,
  Unlock,
  X,
  Zap,
} from 'lucide-react';
import {
  ClauseCondition,
  ContractOutlineClause,
  NotaryClause,
} from '@/lib/types';

export interface ClausesPanelProps {
  liveContractClauses: ContractOutlineClause[];
  contractConditions?: Record<string, string | boolean>;
  onUpdateContractCondition?: (key: string, value: string | boolean | undefined) => void;
  onSetClauseCondition?: (clauseId: string, condition: ClauseCondition | undefined) => void;
  onScrollToLiveClause: (clauseId: string) => void;
  onToggleLiveClauseEnabled?: (clauseId: string) => void;
  onToggleLiveClauseLocked?: (clauseId: string) => void;
  onReorderLiveClauses?: (sourceId: string, targetId: string) => void;
  onMoveLiveClauseInDoc: (domIndex: number, direction: 'up' | 'down') => void;
  onDeleteLiveClauseFromDoc: (clauseId: string) => void;
  onRenameLiveClauseInDoc: (clauseId: string, newTitle: string) => void;
  onInsertNewClauseHeadingInDoc: (customTitle?: string) => void;
  onImportStandardClausesPack?: () => Promise<void>;
  clauses: NotaryClause[];
  onInsertClauseAtCaret: (clause: NotaryClause) => void;
  onSaveNewClause: (
    title: string,
    category: string,
    contentHtml?: string
  ) => Promise<void>;
  onSaveSelectionAsClause: (title: string, category: string) => Promise<void>;
  onEditClauseInEditor: (clause: NotaryClause) => void;
  onDeleteClause: (id: string) => Promise<void>;
  fieldValues: Record<string, string>;
  onOpenSmartVariablesModal: (focusKey?: string) => void;
  onSaveSelectionBookmark: () => void;
}

export default function ClausesPanel({
  liveContractClauses,
  contractConditions = {},
  onUpdateContractCondition,
  onSetClauseCondition,
  onScrollToLiveClause,
  onToggleLiveClauseEnabled,
  onToggleLiveClauseLocked,
  onReorderLiveClauses,
  onMoveLiveClauseInDoc,
  onDeleteLiveClauseFromDoc,
  onRenameLiveClauseInDoc,
  onInsertNewClauseHeadingInDoc,
  onImportStandardClausesPack,
  clauses,
  onInsertClauseAtCaret,
  onSaveNewClause,
  onSaveSelectionAsClause,
  onEditClauseInEditor,
  onDeleteClause,
  fieldValues,
  onOpenSmartVariablesModal,
  onSaveSelectionBookmark,
}: ClausesPanelProps) {
  const [quickDocClauseTitle, setQuickDocClauseTitle] = useState('');
  const [renamingClauseId, setRenamingClauseId] = useState<string | null>(null);
  const [renamingClauseVal, setRenamingClauseVal] = useState('');
  const [draggedClauseId, setDraggedClauseId] = useState<string | null>(null);
  const [dragOverClauseId, setDragOverClauseId] = useState<string | null>(null);

  // Condition editor state
  const [newConditionName, setNewConditionName] = useState('');
  const [editingConditionClauseId, setEditingConditionClauseId] = useState<string | null>(null);
  const [conditionFieldDraft, setConditionFieldDraft] = useState('');
  const [conditionValueDraft, setConditionValueDraft] = useState<string>('true');

  // Ready clauses library state
  const [showSuggestionsLibrary, setShowSuggestionsLibrary] = useState(false);
  const [clauseSearch, setClauseSearch] = useState('');
  const [showAddLibraryForm, setShowAddLibraryForm] = useState(false);
  const [newClauseTitle, setNewClauseTitle] = useState('');
  const [newClauseCategory, setNewClauseCategory] = useState('عام');
  const [newClauseText, setNewClauseText] = useState('');

  const conditionKeys = Object.keys(contractConditions);

  const handleAddCondition = () => {
    const cleanKey = newConditionName.trim();
    if (!cleanKey || !onUpdateContractCondition) return;
    onUpdateContractCondition(cleanKey, true);
    setNewConditionName('');
  };

  return (
    <div className="space-y-4">
      {/* 1. LIVE CONTRACT CLAUSES OUTLINE */}
      <div className="border border-blue-200 bg-blue-50/30 rounded-md p-3 space-y-2.5">
        <div className="flex items-center justify-between border-b border-blue-200/80 pb-2">
          <div className="flex items-center gap-1.5">
            <Hash className="w-3.5 h-3.5 text-blue-900" />
            <span className="text-xs font-bold text-slate-900">
              بنود العقد المفتوح حالياً ({liveContractClauses.length})
            </span>
          </div>
          <span className="text-[10px] text-blue-900 font-medium">
            تفعيل · سحب · قفل · شروط
          </span>
        </div>

        {/* Quick add `# Clause` directly into the open contract */}
        <div className="flex items-center gap-1.5">
          <input
            type="text"
            value={quickDocClauseTitle}
            onChange={(e) => setQuickDocClauseTitle(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && quickDocClauseTitle.trim()) {
                onInsertNewClauseHeadingInDoc(quickDocClauseTitle.trim());
                setQuickDocClauseTitle('');
              }
            }}
            placeholder="عنوان بند جديد في هذا العقد (مثال: التعيين)..."
            className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
          />
          <button
            type="button"
            onMouseDown={(e) => {
              e.preventDefault();
              onSaveSelectionBookmark();
            }}
            onClick={() => {
              onInsertNewClauseHeadingInDoc(
                quickDocClauseTitle.trim() || undefined
              );
              setQuickDocClauseTitle('');
            }}
            className="px-2.5 py-1.5 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800 shrink-0 inline-flex items-center gap-1 transition-colors"
            title="إدراج عنوان بند جديد (#) عند المؤشر في العقد الحالي"
          >
            <Plus className="w-3.5 h-3.5" />
            <span># بند بالعقد</span>
          </button>
        </div>

        {liveContractClauses.length === 0 ? (
          <div className="border border-dashed border-blue-200 rounded p-3 text-center bg-white/80 space-y-1">
            <div className="text-xs font-semibold text-slate-700">
              لا توجد بنود مقسّمة بعلامة `#` في هذا العقد بعد
            </div>
            <p className="text-[11px] text-slate-500 leading-relaxed">
              اكتب <code className="font-mono font-bold text-blue-900"># عنوان البند</code> في بداية أي سطر داخل الورقة أو اضغط زر <strong>«# بند بالعقد»</strong> أعلاه لتقسيم هذا العقد إلى بنود مستقلة قابلة للتفعيل والتعطيل والسحب.
            </p>
          </div>
        ) : (
          <div className="space-y-2 max-h-85 overflow-y-auto pr-0.5">
            {liveContractClauses.map((c, idx) => {
              const isRenaming = renamingClauseId === c.id;
              const isEnabled = c.enabled !== false;
              const isLocked = !!c.locked;
              const isDragOver = dragOverClauseId === c.id;
              const isEditingCond = editingConditionClauseId === c.id;

              return (
                <div
                  key={c.id}
                  draggable={!isRenaming && !isEditingCond}
                  onDragStart={(e) => {
                    setDraggedClauseId(c.id);
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', c.id);
                  }}
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                    if (draggedClauseId && draggedClauseId !== c.id) {
                      setDragOverClauseId(c.id);
                    }
                  }}
                  onDragLeave={() => {
                    if (dragOverClauseId === c.id) {
                      setDragOverClauseId(null);
                    }
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    const srcId =
                      draggedClauseId || e.dataTransfer.getData('text/plain');
                    setDraggedClauseId(null);
                    setDragOverClauseId(null);
                    if (srcId && srcId !== c.id && onReorderLiveClauses) {
                      onReorderLiveClauses(srcId, c.id);
                    }
                  }}
                  onDragEnd={() => {
                    setDraggedClauseId(null);
                    setDragOverClauseId(null);
                  }}
                  className={`border rounded-md p-2.5 space-y-1.5 transition-all ${
                    isDragOver
                      ? 'border-blue-600 bg-blue-50/70 ring-2 ring-blue-500/20'
                      : !isEnabled
                      ? 'border-slate-200 bg-slate-100/80 opacity-75'
                      : isLocked
                      ? 'border-amber-300 bg-amber-50/20'
                      : 'border-slate-200 bg-white hover:border-blue-400'
                  }`}
                >
                  <div className="flex items-start justify-between gap-1.5">
                    {/* Drag Handle + Checkbox Toggle */}
                    <div className="flex items-center gap-1 pt-0.5 shrink-0">
                      <span
                        className="cursor-grab active:cursor-grabbing text-slate-400 hover:text-slate-700"
                        title="اسحب لإعادة ترتيب هذا البند في العقد"
                      >
                        <GripVertical className="w-3.5 h-3.5" />
                      </span>
                      <button
                        type="button"
                        onClick={() =>
                          onToggleLiveClauseEnabled &&
                          onToggleLiveClauseEnabled(c.id)
                        }
                        className={`p-0.5 rounded transition-colors ${
                          isEnabled
                            ? 'text-blue-900 hover:text-blue-700'
                            : 'text-slate-400 hover:text-slate-700'
                        }`}
                        title={
                          isEnabled
                            ? 'البند مفعل وظاهر في ورقة A4 والتصدير — انقر لإخفائه مؤقتاً مع حفظ نصه وموضعه'
                            : 'البند معطل ومخفي من ورقة A4 — انقر لإعادته فوراً إلى موضعه الأصلي'
                        }
                      >
                        {isEnabled ? (
                          <CheckSquare className="w-4 h-4" />
                        ) : (
                          <Square className="w-4 h-4" />
                        )}
                      </button>
                    </div>

                    {isRenaming ? (
                      <div className="flex items-center gap-1 flex-1">
                        <input
                          type="text"
                          value={renamingClauseVal}
                          onChange={(e) => setRenamingClauseVal(e.target.value)}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter' && renamingClauseVal.trim()) {
                              onRenameLiveClauseInDoc(
                                c.id,
                                renamingClauseVal.trim()
                              );
                              setRenamingClauseId(null);
                            } else if (e.key === 'Escape') {
                              setRenamingClauseId(null);
                            }
                          }}
                          autoFocus
                          className="flex-1 px-2 py-0.5 text-xs border border-blue-800 rounded focus:outline-none bg-white"
                        />
                        <button
                          type="button"
                          onClick={() => {
                            if (renamingClauseVal.trim()) {
                              onRenameLiveClauseInDoc(
                                c.id,
                                renamingClauseVal.trim()
                              );
                            }
                            setRenamingClauseId(null);
                          }}
                          className="p-1 text-emerald-700 hover:bg-emerald-50 rounded"
                        >
                          <Check className="w-3.5 h-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => setRenamingClauseId(null)}
                          className="p-1 text-slate-400 hover:bg-slate-100 rounded"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    ) : (
                      <button
                        type="button"
                        onClick={() => {
                          if (isEnabled) {
                            onScrollToLiveClause(c.id);
                          } else if (onToggleLiveClauseEnabled) {
                            onToggleLiveClauseEnabled(c.id);
                          }
                        }}
                        onDoubleClick={() => {
                          setRenamingClauseId(c.id);
                          setRenamingClauseVal(c.title);
                        }}
                        className="text-right flex-1 group min-w-0"
                        title={
                          isEnabled
                            ? 'انقر للقفز إلى موضع البند (أو نقر مزدوج لتعديل العنوان)'
                            : 'هذا البند معطل ومخفي — انقر لإعادة تفعيله في موضعه'
                        }
                      >
                        <div className="text-xs font-bold text-slate-900 group-hover:text-blue-900 flex items-center gap-1 flex-wrap">
                          <span className="text-blue-900 font-mono tabular-nums shrink-0">
                            #{c.index}
                          </span>
                          <span
                            className={`truncate ${
                              !isEnabled ? 'line-through text-slate-500' : ''
                            }`}
                          >
                            {c.title}
                          </span>
                          {!isEnabled && (
                            <span className="text-[9px] font-normal bg-slate-200 text-slate-700 px-1.5 py-0.5 rounded shrink-0">
                              مخفي
                            </span>
                          )}
                          {isLocked && isEnabled && (
                            <span className="text-[9px] font-normal bg-amber-100 text-amber-900 px-1.5 py-0.5 rounded shrink-0">
                              مقفل
                            </span>
                          )}
                          {c.condition && (
                            <span className="text-[9px] font-semibold bg-indigo-100 text-indigo-900 px-1.5 py-0.5 rounded shrink-0 inline-flex items-center gap-0.5">
                              <Zap className="w-2.5 h-2.5" />
                              <span>شرط: {c.condition.field}</span>
                            </span>
                          )}
                        </div>
                        {c.previewText && (
                          <p className="text-[11px] text-slate-500 line-clamp-1 mt-0.5">
                            {c.previewText}
                          </p>
                        )}
                      </button>
                    )}

                    {/* Live Clause Controls inside this Contract */}
                    <div className="flex items-center gap-0.5 shrink-0">
                      {/* Link/Edit Condition Button */}
                      {onSetClauseCondition && (
                        <button
                          type="button"
                          onClick={() => {
                            if (isEditingCond) {
                              setEditingConditionClauseId(null);
                            } else {
                              setEditingConditionClauseId(c.id);
                              setConditionFieldDraft(
                                c.condition?.field || conditionKeys[0] || ''
                              );
                              setConditionValueDraft(
                                c.condition?.value !== undefined
                                  ? String(c.condition.value)
                                  : 'true'
                              );
                            }
                          }}
                          className={`p-1 rounded ${
                            c.condition
                              ? 'text-indigo-700 bg-indigo-100/80 hover:bg-indigo-200/70'
                              : 'text-slate-400 hover:text-indigo-700'
                          }`}
                          title="ربط هذا البند بشرط مخصص ليظهر أو يختفي تلقائياً"
                        >
                          <Zap className="w-3.5 h-3.5" />
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() =>
                          onToggleLiveClauseLocked &&
                          onToggleLiveClauseLocked(c.id)
                        }
                        className={`p-1 rounded ${
                          isLocked
                            ? 'text-amber-700 bg-amber-100/80 hover:bg-amber-200/70'
                            : 'text-slate-400 hover:text-amber-700'
                        }`}
                        title={
                          isLocked
                            ? 'البند مقفل ضد التعديل بالخطأ — انقر لفك القفل'
                            : 'قفل البند لمنع تعديله بالخطأ في الورقة'
                        }
                      >
                        {isLocked ? (
                          <Lock className="w-3.5 h-3.5" />
                        ) : (
                          <Unlock className="w-3.5 h-3.5" />
                        )}
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          if (idx > 0 && onReorderLiveClauses) {
                            onReorderLiveClauses(
                              c.id,
                              liveContractClauses[idx - 1].id
                            );
                          } else {
                            onMoveLiveClauseInDoc(c.domIndex, 'up');
                          }
                        }}
                        disabled={idx === 0}
                        className="p-1 text-slate-400 hover:text-slate-900 disabled:opacity-30"
                        title="تحريك البند لأعلى داخل هذا العقد"
                      >
                        <ArrowUp className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          if (
                            idx < liveContractClauses.length - 1 &&
                            onReorderLiveClauses
                          ) {
                            onReorderLiveClauses(
                              c.id,
                              liveContractClauses[idx + 1].id
                            );
                          } else {
                            onMoveLiveClauseInDoc(c.domIndex, 'down');
                          }
                        }}
                        disabled={idx === liveContractClauses.length - 1}
                        className="p-1 text-slate-400 hover:text-slate-900 disabled:opacity-30"
                        title="تحريك البند لأسفل داخل هذا العقد"
                      >
                        <ArrowDown className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setRenamingClauseId(c.id);
                          setRenamingClauseVal(c.title);
                        }}
                        className="p-1 text-slate-400 hover:text-blue-900"
                        title="تعديل عنوان البند في الورقة"
                      >
                        <Edit3 className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          onSaveNewClause(
                            c.title,
                            'بنود مستخرجة',
                            c.contentHtml
                          )
                        }
                        className="p-1 text-slate-400 hover:text-emerald-700"
                        title="حفظ نسخة من هذا البند في مكتبة المكتب لإعادة استخدامه في عقود أخرى"
                      >
                        <BookmarkPlus className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteLiveClauseFromDoc(c.id)}
                        className="p-1 text-slate-400 hover:text-red-600"
                        title="حذف هذا البند نهائياً من العقد الحالي"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>

                  {/* Inline Condition Linker Popover for this Clause */}
                  {isEditingCond && onSetClauseCondition && (
                    <div className="p-2 bg-indigo-50/70 border border-indigo-200 rounded space-y-1.5 text-[11px]">
                      <div className="font-bold text-indigo-950 flex items-center justify-between">
                        <span>ربط البند بشرط مخصص (يظهر البند عند تحقق الشرط):</span>
                        <button
                          type="button"
                          onClick={() => setEditingConditionClauseId(null)}
                          className="text-slate-400 hover:text-slate-700"
                        >
                          <X className="w-3.5 h-3.5" />
                        </button>
                      </div>
                      <div className="flex items-center gap-1.5">
                        {conditionKeys.length > 0 ? (
                          <select
                            value={conditionFieldDraft}
                            onChange={(e) => setConditionFieldDraft(e.target.value)}
                            className="flex-1 px-2 py-1 bg-white border border-indigo-300 rounded text-xs"
                          >
                            <option value="">-- اختر شرطاً معرفاً --</option>
                            {conditionKeys.map((k) => (
                              <option key={k} value={k}>
                                {k}
                              </option>
                            ))}
                          </select>
                        ) : null}
                        <input
                          type="text"
                          value={conditionFieldDraft}
                          onChange={(e) => setConditionFieldDraft(e.target.value)}
                          placeholder="أو اكتب اسم شرط جديد..."
                          className="flex-1 px-2 py-1 bg-white border border-indigo-300 rounded text-xs"
                        />
                      </div>
                      <div className="flex items-center justify-between gap-1.5">
                        <div className="flex items-center gap-1">
                          <span className="text-slate-600">يتفعل عندما تكون القيمة:</span>
                          <select
                            value={conditionValueDraft}
                            onChange={(e) => setConditionValueDraft(e.target.value)}
                            className="px-2 py-0.5 bg-white border border-indigo-300 rounded text-xs font-semibold"
                          >
                            <option value="true">مفعل (نعم ✓)</option>
                            <option value="false">غير مفعل (لا)</option>
                          </select>
                        </div>
                        <div className="flex items-center gap-1">
                          {c.condition && (
                            <button
                              type="button"
                              onClick={() => {
                                onSetClauseCondition(c.id, undefined);
                                setEditingConditionClauseId(null);
                              }}
                              className="px-2 py-0.5 bg-white border border-red-200 text-red-700 rounded hover:bg-red-50"
                            >
                              إلغاء الشرط
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={() => {
                              const cleanField = conditionFieldDraft.trim();
                              if (!cleanField) return;
                              const boolVal =
                                conditionValueDraft === 'true'
                                  ? true
                                  : conditionValueDraft === 'false'
                                  ? false
                                  : conditionValueDraft;
                              if (
                                onUpdateContractCondition &&
                                contractConditions[cleanField] === undefined
                              ) {
                                onUpdateContractCondition(cleanField, boolVal);
                              }
                              onSetClauseCondition(c.id, {
                                field: cleanField,
                                value: boolVal,
                              });
                              setEditingConditionClauseId(null);
                            }}
                            className="px-2.5 py-0.5 bg-indigo-900 text-white font-semibold rounded hover:bg-indigo-800"
                          >
                            حفظ وربط
                          </button>
                        </div>
                      </div>
                    </div>
                  )}

                  {c.variables.length > 0 && (
                    <div className="flex items-center flex-wrap gap-1 pt-1 border-t border-slate-100">
                      <span className="text-[10px] text-slate-400">المتغيرات:</span>
                      {c.variables.map((v) => (
                        <button
                          key={v}
                          type="button"
                          onClick={() => onOpenSmartVariablesModal(v)}
                          className={`text-[10px] font-mono px-1.5 py-0.5 rounded border ${
                            fieldValues[v]?.trim()
                              ? 'bg-emerald-50 text-emerald-900 border-emerald-200'
                              : 'bg-pink-50 text-pink-900 border-pink-200'
                          }`}
                        >
                          [{v}]
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* 2. CUSTOM OFFICE CONDITIONAL CLAUSES MANAGER (شروط البنود المخصصة للمكتب — تبدأ فارغة افتراضياً) */}
      {onUpdateContractCondition && (
        <div className="border border-indigo-200 bg-indigo-50/25 rounded-md p-3 space-y-2.5">
          <div className="flex items-center justify-between border-b border-indigo-200/70 pb-1.5">
            <div className="flex items-center gap-1.5">
              <Zap className="w-3.5 h-3.5 text-indigo-900" />
              <span className="text-xs font-bold text-slate-900">
                شروط العقد والبنود الشرطية ({conditionKeys.length})
              </span>
            </div>
            <span className="text-[10px] text-indigo-900 font-medium">
              إظهار/إخفاء تلقائي
            </span>
          </div>

          <div className="flex items-center gap-1.5">
            <input
              type="text"
              value={newConditionName}
              onChange={(e) => setNewConditionName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleAddCondition();
              }}
              placeholder="أضف شرطاً خاصاً بمكتبك (تبدأ فارغة)..."
              className="flex-1 px-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded focus:border-indigo-800 focus:outline-none"
            />
            <button
              type="button"
              onClick={handleAddCondition}
              className="px-2.5 py-1.5 bg-indigo-900 text-white text-[11px] font-semibold rounded hover:bg-indigo-800 shrink-0 inline-flex items-center gap-1"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>+ إضافة شرط</span>
            </button>
          </div>

          {conditionKeys.length === 0 ? (
            <div className="text-[11px] text-slate-500 bg-white/80 border border-dashed border-indigo-200 rounded p-2.5 text-center">
              قائمة الشروط فارغة افتراضياً. أنشئ أي شرط يخص عقود مكتبك هنا ثم اربط به أي بند عبر أيقونة <Zap className="w-3 h-3 inline text-indigo-800" /> ليظهر أو يختفي البند تلقائياً.
            </div>
          ) : (
            <div className="space-y-1.5">
              {conditionKeys.map((key) => {
                const val = contractConditions[key];
                const isActive = val === true || val === 'true' || val === 'نعم';
                const linkedCount = liveContractClauses.filter(
                  (c) => c.condition?.field === key
                ).length;
                return (
                  <div
                    key={key}
                    className="flex items-center justify-between gap-2 px-2.5 py-1.5 bg-white border border-slate-200 rounded"
                  >
                    <label className="flex items-center gap-2 cursor-pointer flex-1 min-w-0">
                      <input
                        type="checkbox"
                        checked={isActive}
                        onChange={(e) =>
                          onUpdateContractCondition(key, e.target.checked)
                        }
                        className="rounded border-slate-300 text-indigo-900"
                      />
                      <span className="text-xs font-semibold text-slate-800 truncate">
                        {key}
                      </span>
                      {linkedCount > 0 && (
                        <span className="text-[10px] bg-indigo-50 text-indigo-900 border border-indigo-200 px-1.5 py-0.2 rounded shrink-0">
                          {linkedCount} بند مرتبط
                        </span>
                      )}
                    </label>
                    <button
                      type="button"
                      onClick={() => onUpdateContractCondition(key, undefined)}
                      className="p-1 text-slate-400 hover:text-red-600"
                      title="حذف هذا الشرط"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* 3. OPTIONAL CLAUSES SUGGESTIONS LIBRARY */}
      <div className="border border-slate-200 rounded-md bg-white overflow-hidden">
        <button
          type="button"
          onClick={() => setShowSuggestionsLibrary((v) => !v)}
          className="w-full px-3 py-2.5 bg-slate-50 hover:bg-slate-100 border-b border-slate-200 flex items-center justify-between text-right transition-colors"
        >
          <div>
            <div className="text-xs font-bold text-slate-900">
              مكتبة البنود المقترحة للإدراج ({clauses.length})
            </div>
            <div className="text-[10px] text-slate-500">
              مكتبة مساعدة فارغة افتراضياً — لا تدخل في العقد إلا عند ضغط «+ إدراج في العقد»
            </div>
          </div>
          {showSuggestionsLibrary ? (
            <ChevronUp className="w-4 h-4 text-slate-500 shrink-0" />
          ) : (
            <ChevronDown className="w-4 h-4 text-slate-500 shrink-0" />
          )}
        </button>

        {showSuggestionsLibrary && (
          <div className="p-3 space-y-3">
            <div className="flex items-center gap-1.5">
              <input
                type="text"
                value={clauseSearch}
                onChange={(e) => setClauseSearch(e.target.value)}
                placeholder="بحث في مكتبة البنود المقترحة..."
                className="flex-1 px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded focus:bg-white focus:border-blue-900 focus:outline-none"
              />
              <button
                type="button"
                onClick={() => setShowAddLibraryForm((v) => !v)}
                className="px-2.5 py-1.5 bg-slate-900 text-white text-[11px] font-medium rounded hover:bg-slate-800 shrink-0"
              >
                {showAddLibraryForm ? 'إغلاق' : '+ بند للمكتبة'}
              </button>
            </div>

            {onImportStandardClausesPack && (
              <button
                type="button"
                onClick={() => onImportStandardClausesPack()}
                className="w-full py-1.5 px-2.5 bg-blue-50 hover:bg-blue-100/80 border border-blue-200 text-blue-950 text-[11px] font-semibold rounded inline-flex items-center justify-center gap-1.5 transition-colors"
                title="تحميل حزمة البنود التوثيقية القياسية عند الطلب"
              >
                <Sparkles className="w-3.5 h-3.5 text-blue-800" />
                <span>استيراد الحزمة القياسية الاختيارية للمكتب</span>
              </button>
            )}

            {showAddLibraryForm && (
              <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-md space-y-2">
                <div className="text-[11px] font-bold text-slate-900">
                  حفظ قالب بند جديد في مكتبة المكتب العامة
                </div>
                <input
                  type="text"
                  value={newClauseTitle}
                  onChange={(e) => setNewClauseTitle(e.target.value)}
                  placeholder="عنوان البند (مثال: بند أصل الملكية / بند الضمان)"
                  className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                />
                <input
                  type="text"
                  value={newClauseCategory}
                  onChange={(e) => setNewClauseCategory(e.target.value)}
                  placeholder="التصنيف (بيع، تأسيس شركة، إيجار...)"
                  className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                />
                <textarea
                  rows={3}
                  value={newClauseText}
                  onChange={(e) => setNewClauseText(e.target.value)}
                  placeholder="نص البند (يمكن تضمين وسوم مثل [البائع]) أو اتركه فارغاً لحفظ النص المحدد حالياً في الورقة..."
                  className="w-full px-2 py-1 text-xs bg-white border border-slate-300 rounded"
                />
                <div className="grid grid-cols-2 gap-1.5">
                  <button
                    type="button"
                    onClick={async () => {
                      if (!newClauseTitle.trim()) return;
                      const html = newClauseText.trim()
                        ? `<p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;text-align:justify;font-weight:bold;"># ${newClauseTitle.trim()}</p>` +
                          newClauseText
                            .trim()
                            .split(/\r?\n/)
                            .map(
                              (line) =>
                                `<p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;text-align:justify;">${
                                  line.trim() || '<br>'
                                }</p>`
                            )
                            .join('')
                        : undefined;
                      await onSaveNewClause(
                        newClauseTitle.trim(),
                        newClauseCategory.trim() || 'عام',
                        html
                      );
                      setNewClauseTitle('');
                      setNewClauseText('');
                      setShowAddLibraryForm(false);
                    }}
                    className="py-1.5 px-2 bg-blue-900 text-white text-[11px] font-semibold rounded hover:bg-blue-800"
                  >
                    حفظ في المكتبة
                  </button>
                  <button
                    type="button"
                    onMouseDown={(e) => {
                      e.preventDefault();
                      onSaveSelectionBookmark();
                    }}
                    onClick={async () => {
                      if (!newClauseTitle.trim()) return;
                      await onSaveSelectionAsClause(
                        newClauseTitle.trim(),
                        newClauseCategory.trim() || 'عام'
                      );
                      setNewClauseTitle('');
                      setShowAddLibraryForm(false);
                    }}
                    className="py-1.5 px-2 bg-slate-800 text-white text-[11px] font-medium rounded hover:bg-slate-700"
                  >
                    حفظ المحدد بالورقة
                  </button>
                </div>
              </div>
            )}

            {clauses.length === 0 ? (
              <div className="border border-dashed border-slate-200 rounded p-4 text-center text-xs text-slate-500 space-y-2">
                <p>
                  مكتبة البنود المقترحة فارغة افتراضياً للحفاظ على خفة النظام. يمكنك حفظ أي بند من العقد بالضغط على أيقونة الحفظ بجانب البند، أو استيراد الحزمة القياسية أعلاه.
                </p>
              </div>
            ) : (
              <div className="space-y-2 max-h-72 overflow-y-auto pr-0.5">
                {clauses
                  .filter(
                    (c) =>
                      !clauseSearch.trim() ||
                      c.title.includes(clauseSearch.trim()) ||
                      c.category.includes(clauseSearch.trim())
                  )
                  .map((clause) => (
                    <div
                      key={clause.id}
                      className="border border-slate-200 rounded p-2.5 bg-slate-50/40 hover:bg-white transition-colors space-y-2"
                    >
                      <div className="flex items-start justify-between gap-2">
                        <div>
                          <div className="text-xs font-bold text-slate-900">
                            {clause.title}
                          </div>
                          <div className="text-[10px] text-slate-500">
                            {clause.category}
                          </div>
                        </div>
                        <div className="flex items-center gap-1 shrink-0">
                          <button
                            type="button"
                            onClick={() => onEditClauseInEditor(clause)}
                            className="p-1 text-slate-400 hover:text-blue-900"
                            title="تعديل نص هذا البند في المكتبة"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => onDeleteClause(clause.id)}
                            className="p-1 text-slate-400 hover:text-red-600"
                            title="حذف من المكتبة"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </div>

                      <div className="flex items-center gap-1.5">
                        <button
                          type="button"
                          onMouseDown={(e) => {
                            e.preventDefault();
                            onSaveSelectionBookmark();
                          }}
                          onClick={() => onInsertClauseAtCaret(clause)}
                          className="w-full py-1 px-2 bg-white border border-slate-300 hover:border-blue-900 hover:bg-blue-50/40 text-slate-800 text-[11px] font-semibold rounded transition-colors"
                        >
                          + إدراج نسخة من البند عند المؤشر في العقد الحالي
                        </button>
                      </div>
                    </div>
                  ))}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
