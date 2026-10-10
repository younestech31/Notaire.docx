'use client';

import React, { useEffect, useRef } from 'react';
import { Calendar, Check, ExternalLink, Hash, Sparkles, Type, X } from 'lucide-react';
import { VariableInputType } from '@/lib/types';

export interface InlineVariablePopoverState {
  varKey: string;
  rect: {
    top: number;
    left: number;
    bottom: number;
    width: number;
  };
}

interface InlineVariablePopoverProps {
  popover: InlineVariablePopoverState | null;
  value: string;
  inputType: VariableInputType;
  onChangeValue: (key: string, val: string) => void;
  onChangeInputType: (key: string, type: 'text' | 'number' | 'date') => void;
  onOpenFullModal: (key: string) => void;
  onClose: () => void;
}

/**
 * Converts an integer number into standard Algerian Arabic legal words (التفقيط الجزائري).
 */
export function numberToAlgerianLegalWords(rawNumber: string | number): string {
  const cleaned = String(rawNumber ?? '').replace(/[^\d]/g, '');
  if (!cleaned) return '';
  const num = parseInt(cleaned, 10);
  if (Number.isNaN(num) || num < 0) return '';
  if (num === 0) return 'صفر دينار جزائري';

  const ones = [
    '',
    'واحد',
    'اثنان',
    'ثلاثة',
    'أربعة',
    'خمسة',
    'ستة',
    'سبعة',
    'ثمانية',
    'تسعة',
    'عشرة',
    'أحد عشر',
    'اثنا عشر',
    'ثلاثة عشر',
    'أربعة عشر',
    'خمسة عشر',
    'ستة عشر',
    'سبعة عشر',
    'ثمانية عشر',
    'تسعة عشر',
  ];

  const tens = [
    '',
    'عشرة',
    'عشرون',
    'ثلاثون',
    'أربعون',
    'خمسون',
    'ستون',
    'سبعون',
    'ثمانون',
    'تسعون',
  ];

  const hundreds = [
    '',
    'مائة',
    'مائتان',
    'ثلاثمائة',
    'أربعمائة',
    'خمسمائة',
    'ستمائة',
    'سبعمائة',
    'ثمانمائة',
    'تسعمائة',
  ];

  const convertBelow1000 = (n: number): string => {
    const parts: string[] = [];
    const h = Math.floor(n / 100);
    const rem = n % 100;
    if (h > 0) parts.push(hundreds[h]);
    if (rem > 0) {
      if (rem < 20) {
        parts.push(ones[rem]);
      } else {
        const o = rem % 10;
        const t = Math.floor(rem / 10);
        if (o > 0) {
          parts.push(`${ones[o]} و${tens[t]}`);
        } else {
          parts.push(tens[t]);
        }
      }
    }
    return parts.join(' و');
  };

  const parts: string[] = [];
  const billions = Math.floor(num / 1_000_000_000);
  const millions = Math.floor((num % 1_000_000_000) / 1_000_000);
  const thousands = Math.floor((num % 1_000_000) / 1000);
  const remainder = num % 1000;

  if (billions > 0) {
    if (billions === 1) parts.push('مليار واحد');
    else if (billions === 2) parts.push('ملياران');
    else if (billions >= 3 && billions <= 10)
      parts.push(`${convertBelow1000(billions)} ملايير`);
    else parts.push(`${convertBelow1000(billions)} مليار`);
  }

  if (millions > 0) {
    if (millions === 1) parts.push('مليون واحد');
    else if (millions === 2) parts.push('مليونان');
    else if (millions >= 3 && millions <= 10)
      parts.push(`${convertBelow1000(millions)} ملايين`);
    else parts.push(`${convertBelow1000(millions)} مليون`);
  }

  if (thousands > 0) {
    if (thousands === 1) parts.push('ألف');
    else if (thousands === 2) parts.push('ألفان');
    else if (thousands >= 3 && thousands <= 10)
      parts.push(`${convertBelow1000(thousands)} آلاف`);
    else parts.push(`${convertBelow1000(thousands)} ألف`);
  }

  if (remainder > 0) {
    parts.push(convertBelow1000(remainder));
  }

  return `${parts.join(' و')} دينار جزائري`;
}

/**
 * Converts an ISO date YYYY-MM-DD into formal Algerian notary date phrasing.
 */
export function dateToAlgerianLegalWords(isoDate: string): string {
  if (!isoDate || !/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return '';
  const [yStr, mStr, dStr] = isoDate.split('-');
  const y = parseInt(yStr, 10);
  const m = parseInt(mStr, 10);
  const d = parseInt(dStr, 10);
  const months = [
    '',
    'جانفي',
    'فيفري',
    'مارس',
    'أفريل',
    'ماي',
    'جوان',
    'جويلية',
    'أوت',
    'سبتمبر',
    'أكتوبر',
    'نوفمبر',
    'ديسمبر',
  ];
  if (!months[m]) return '';
  return `${d} ${months[m]} سنة ${y}`;
}

export default function InlineVariablePopover({
  popover,
  value,
  inputType,
  onChangeValue,
  onChangeInputType,
  onOpenFullModal,
  onClose,
}: InlineVariablePopoverProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (popover) {
      setTimeout(() => {
        inputRef.current?.focus();
        inputRef.current?.select();
      }, 25);
    }
  }, [popover?.varKey]);

  useEffect(() => {
    if (!popover) return;
    const handleOutsideClick = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null;
      if (!target) return;
      if (containerRef.current?.contains(target)) return;
      if (target.closest('.smart-tag, .smart-placeholder')) return;
      onClose();
    };
    window.addEventListener('mousedown', handleOutsideClick);
    return () => window.removeEventListener('mousedown', handleOutsideClick);
  }, [popover, onClose]);

  if (!popover) return null;

  const currentType: 'text' | 'number' | 'date' =
    inputType === 'number' ? 'number' : inputType === 'date' ? 'date' : 'text';

  const tafqeetPreview =
    currentType === 'number' && value.trim()
      ? numberToAlgerianLegalWords(value)
      : currentType === 'date' && value.trim()
      ? dateToAlgerianLegalWords(value)
      : '';

  // Compute safe viewport position above (or below) the clicked smart tag
  const popoverWidth = 310;
  const viewportW = typeof window !== 'undefined' ? window.innerWidth : 1200;
  const clampedLeft = Math.min(
    Math.max(12, popover.rect.left + popover.rect.width / 2 - popoverWidth / 2),
    viewportW - popoverWidth - 12
  );
  const placeBelow = popover.rect.top < 175;
  const topPos = placeBelow ? popover.rect.bottom + 8 : Math.max(12, popover.rect.top - 152);

  return (
    <div
      ref={containerRef}
      dir="rtl"
      style={{
        position: 'fixed',
        top: `${topPos}px`,
        left: `${clampedLeft}px`,
        width: `${popoverWidth}px`,
        zIndex: 9999,
      }}
      className="bg-white rounded-lg border border-slate-300 shadow-2xl p-2.5 select-none animate-in fade-in zoom-in-95 duration-150"
    >
      {/* Header: Variable Key + Type Switcher + Close */}
      <div className="flex items-center justify-between gap-1.5 mb-2">
        <span className="text-[11px] font-mono font-bold text-pink-900 bg-pink-50 border border-pink-200 px-1.5 py-0.5 rounded truncate max-w-[145px]">
          [{popover.varKey}]
        </span>

        <div className="flex items-center gap-1">
          <div className="inline-flex items-center bg-slate-100 border border-slate-200 rounded p-0.5 text-[10px]">
            <button
              type="button"
              onClick={() => onChangeInputType(popover.varKey, 'text')}
              className={`px-1.5 py-0.5 rounded inline-flex items-center gap-0.5 ${
                currentType === 'text'
                  ? 'bg-white text-slate-900 font-bold shadow-2xs'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
              title="حقل نصي"
            >
              <Type className="w-2.5 h-2.5" />
              <span>نص</span>
            </button>
            <button
              type="button"
              onClick={() => onChangeInputType(popover.varKey, 'number')}
              className={`px-1.5 py-0.5 rounded inline-flex items-center gap-0.5 ${
                currentType === 'number'
                  ? 'bg-white text-blue-900 font-bold shadow-2xs'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
              title="حقل رقمي مع تفقيط"
            >
              <Hash className="w-2.5 h-2.5" />
              <span>رقم</span>
            </button>
            <button
              type="button"
              onClick={() => onChangeInputType(popover.varKey, 'date')}
              className={`px-1.5 py-0.5 rounded inline-flex items-center gap-0.5 ${
                currentType === 'date'
                  ? 'bg-white text-blue-900 font-bold shadow-2xs'
                  : 'text-slate-500 hover:text-slate-800'
              }`}
              title="حقل تاريخ"
            >
              <Calendar className="w-2.5 h-2.5" />
              <span>تاريخ</span>
            </button>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="p-1 text-slate-400 hover:text-slate-700 rounded"
            title="إغلاق (Esc)"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>

      {/* Instant Input */}
      <div className="flex items-center gap-1.5">
        <input
          ref={inputRef}
          type={
            currentType === 'date'
              ? 'date'
              : currentType === 'number'
              ? 'number'
              : 'text'
          }
          value={value}
          onChange={(e) => onChangeValue(popover.varKey, e.target.value)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              onClose();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              onClose();
            }
          }}
          placeholder={`أدخل قيمة [${popover.varKey.replace(/_/g, ' ')}]...`}
          className="flex-1 px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-300 rounded focus:bg-white focus:border-blue-900 focus:outline-none text-slate-900"
        />
        <button
          type="button"
          onClick={onClose}
          className="px-2.5 py-1.5 bg-blue-950 hover:bg-blue-900 text-white text-[11px] font-semibold rounded inline-flex items-center gap-1 shrink-0"
          title="تأكيد القيمة (Enter)"
        >
          <Check className="w-3.5 h-3.5" />
          <span>حفظ</span>
        </button>
      </div>

      {/* Tafqeet Helper Bar (for number / date) */}
      {tafqeetPreview && (
        <div className="mt-1.5 p-1.5 bg-amber-50/70 border border-amber-200 rounded flex items-center justify-between gap-1.5">
          <div className="text-[10px] text-amber-950 truncate flex items-center gap-1">
            <Sparkles className="w-3 h-3 text-amber-700 shrink-0" />
            <span className="truncate" title={tafqeetPreview}>
              {tafqeetPreview}
            </span>
          </div>
          <button
            type="button"
            onClick={() => onChangeValue(popover.varKey, tafqeetPreview)}
            className="px-1.5 py-0.5 bg-amber-700 hover:bg-amber-800 text-white text-[9px] font-bold rounded shrink-0"
            title="استبدال القيمة بالصيغة الحرفية التوثيقية"
          >
            اعتماد بالحروف
          </button>
        </div>
      )}

      {/* Footer: Open Full Variables Form */}
      <div className="mt-2 pt-1.5 border-t border-slate-100 flex items-center justify-between">
        <span className="text-[10px] text-slate-400">
          {value.trim() ? 'معبأ في جميع مواضع العقد ✓' : 'غير معبأ بعد'}
        </span>
        <button
          type="button"
          onClick={() => {
            const key = popover.varKey;
            onClose();
            onOpenFullModal(key);
          }}
          className="text-[10px] font-bold text-blue-900 hover:text-blue-700 inline-flex items-center gap-1"
        >
          <span>فتح الاستمارة الكاملة</span>
          <ExternalLink className="w-2.5 h-2.5" />
        </button>
      </div>
    </div>
  );
}
