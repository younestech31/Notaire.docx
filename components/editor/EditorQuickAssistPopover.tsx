'use client';

import React from 'react';
import {
  FileText,
  Hash,
  SeparatorHorizontal,
  Sparkles,
  Table as TableIcon,
  UserPlus,
} from 'lucide-react';

export interface QuickAssistOption {
  id: string;
  kind: 'tag' | 'clause' | 'slash';
  label: string;
  subtitle?: string;
  insertValue: string;
  slashAction?: 'clause' | 'tag' | 'table' | 'party' | 'page-break';
  isFilled?: boolean;
}

interface EditorQuickAssistPopoverProps {
  open: boolean;
  triggerType: 'tag' | 'clause' | 'slash';
  query: string;
  position: { top: number; left: number; bottom: number };
  options: QuickAssistOption[];
  selectedIndex: number;
  onSelectOption: (option: QuickAssistOption) => void;
  onClose: () => void;
  onDisableAssistType?: (type: 'autocomplete' | 'slash') => void;
}

export default function EditorQuickAssistPopover({
  open,
  triggerType,
  query,
  position,
  options,
  selectedIndex,
  onSelectOption,
  onClose,
  onDisableAssistType,
}: EditorQuickAssistPopoverProps) {
  if (!open || options.length === 0) return null;

  const viewportHeight =
    typeof window !== 'undefined' ? window.innerHeight : 800;
  const viewportWidth =
    typeof window !== 'undefined' ? window.innerWidth : 1200;

  const popoverHeight = 270;
  const popoverWidth = 300;

  const placeAbove = position.bottom + popoverHeight > viewportHeight - 16;
  const top = placeAbove
    ? Math.max(12, position.top - popoverHeight - 6)
    : Math.min(viewportHeight - popoverHeight - 12, position.bottom + 6);

  const left = Math.min(
    Math.max(16, position.left - 120),
    viewportWidth - popoverWidth - 16
  );

  const headerTitle =
    triggerType === 'tag'
      ? 'إكمال البيان [...]'
      : triggerType === 'clause'
      ? 'إدراج عنوان بند توثيقي (#)'
      : 'أوامر الإدراج السريع (/)';

  return (
    <div
      dir="rtl"
      style={{ top: `${top}px`, left: `${left}px`, width: `${popoverWidth}px` }}
      onMouseDown={(e) => {
        // Keep focus inside editor
        e.preventDefault();
      }}
      className="fixed z-50 bg-white border border-[#C4A46A]/70 rounded-lg shadow-2xl overflow-hidden text-right no-print"
    >
      {/* Header */}
      <div className="px-3 py-1.5 bg-[#0F2744] text-white flex items-center justify-between text-[11px]">
        <div className="flex items-center gap-1.5 font-bold">
          {triggerType === 'tag' && (
            <span className="font-mono text-[10px] bg-pink-500/30 px-1 rounded text-pink-200">
              [ ]
            </span>
          )}
          {triggerType === 'clause' && (
            <Hash className="w-3.5 h-3.5 text-amber-300" />
          )}
          {triggerType === 'slash' && (
            <Sparkles className="w-3.5 h-3.5 text-amber-300" />
          )}
          <span>{headerTitle}</span>
        </div>
        {query && (
          <span className="font-mono text-[10px] bg-white/15 px-1.5 py-0.5 rounded truncate max-w-[100px]">
            {query}
          </span>
        )}
      </div>

      {/* Options List */}
      <div className="max-h-52 overflow-y-auto divide-y divide-slate-100 py-1">
        {options.map((opt, idx) => {
          const isSelected = idx === selectedIndex;
          return (
            <button
              key={opt.id}
              type="button"
              onMouseDown={(e) => {
                e.preventDefault();
                onSelectOption(opt);
              }}
              className={`w-full px-3 py-1.5 flex items-center justify-between gap-2 text-right transition-colors ${
                isSelected
                  ? 'bg-blue-50 text-[#0F2744] font-bold'
                  : 'text-slate-700 hover:bg-slate-50'
              }`}
            >
              <div className="min-w-0 flex-1">
                <div className="text-xs truncate flex items-center gap-1.5">
                  {opt.kind === 'slash' && opt.slashAction === 'clause' && (
                    <Hash className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                  )}
                  {opt.kind === 'slash' && opt.slashAction === 'tag' && (
                    <FileText className="w-3.5 h-3.5 text-pink-700 shrink-0" />
                  )}
                  {opt.kind === 'slash' && opt.slashAction === 'table' && (
                    <TableIcon className="w-3.5 h-3.5 text-blue-800 shrink-0" />
                  )}
                  {opt.kind === 'slash' && opt.slashAction === 'party' && (
                    <UserPlus className="w-3.5 h-3.5 text-emerald-700 shrink-0" />
                  )}
                  {opt.kind === 'slash' && opt.slashAction === 'page-break' && (
                    <SeparatorHorizontal className="w-3.5 h-3.5 text-slate-600 shrink-0" />
                  )}
                  <span className="truncate">{opt.label}</span>
                </div>
                {opt.subtitle && (
                  <div className="text-[10px] text-slate-400 truncate mt-0.5">
                    {opt.subtitle}
                  </div>
                )}
              </div>

              {opt.kind === 'tag' && (
                <span
                  className={`text-[10px] font-mono px-1.5 py-0.5 rounded shrink-0 ${
                    opt.isFilled
                      ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                      : 'bg-pink-50 text-pink-800 border border-pink-200'
                  }`}
                >
                  {`[${opt.insertValue}]`}
                </span>
              )}

              {opt.kind === 'clause' && (
                <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-amber-50 text-amber-900 border border-amber-200 shrink-0">
                  {opt.insertValue}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* Footer Hint + Quick Disable */}
      <div className="px-2.5 py-1 bg-slate-50 border-t border-slate-200/80 flex items-center justify-between text-[10px] text-slate-500">
        <span>↑↓ للتنقل · Enter للإدراج · Esc للإغلاق</span>
        {onDisableAssistType && (
          <button
            type="button"
            onMouseDown={(e) => {
              e.preventDefault();
              onDisableAssistType(
                triggerType === 'slash' ? 'slash' : 'autocomplete'
              );
              onClose();
            }}
            className="text-slate-400 hover:text-red-700 underline"
            title="إيقاف هذه القائمة من إعدادات المساعدة الذكية"
          >
            إيقاف التلقائي
          </button>
        )}
      </div>
    </div>
  );
}
