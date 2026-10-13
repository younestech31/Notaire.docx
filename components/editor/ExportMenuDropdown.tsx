'use client';

import React, { useEffect, useRef, useState } from 'react';
import {
  ChevronDown,
  Download,
  FileCheck,
  FileCode,
  FileSpreadsheet,
  FileText,
  Layers,
  Send,
  Settings,
  Sparkles,
} from 'lucide-react';
import {
  BUILTIN_STANDARD_WORD_TEMPLATES,
  hasCustomizedNotaryTemplate,
  StandardNotaryTemplateId,
} from '@/lib/notary-word-templates';

interface ExportMenuDropdownProps {
  onExportDirectDocx: () => void;
  onExportWithTemplate: (templateId: StandardNotaryTemplateId) => void;
  onOpenTemplatesManager: () => void;
  disabled?: boolean;
}

export default function ExportMenuDropdown({
  onExportDirectDocx,
  onExportWithTemplate,
  onOpenTemplatesManager,
  disabled = false,
}: ExportMenuDropdownProps) {
  const [isOpen, setIsOpen] = useState<boolean>(false);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdown on click outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(event.target as Node)
      ) {
        setIsOpen(false);
      }
    }
    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  const items: Array<{
    id: StandardNotaryTemplateId;
    title: string;
    subtitle: string;
    icon: React.ReactNode;
  }> = [
    {
      id: 'acte_original',
      title: 'أصل العقد (Acte Original / Minute)',
      subtitle: 'العقد الأصلي المحفوظ بالأرشيف مع هوامش التوثيق الرسمية',
      icon: <FileText className="w-4 h-4 text-blue-800" />,
    },
    {
      id: 'contrat_reference',
      title: 'ملف العقد المرجعي (Contrat de Référence)',
      subtitle: 'الملف النموذجي الشامل الذي ينسخ تنسيقه لباقي العقود',
      icon: <Layers className="w-4 h-4 text-indigo-800" />,
    },
    {
      id: 'expedition',
      title: 'نسخة العقد (Expédition / Copie)',
      subtitle: 'النسخة الرسمية المسلمة للأطراف مع شهادة المطابقة',
      icon: <FileCheck className="w-4 h-4 text-emerald-800" />,
    },
    {
      id: 'extrait',
      title: "المستخرج (Extrait d'Acte)",
      subtitle: 'مستخرج العقد الموجه لمفتشية التسجيل والضرائب',
      icon: <FileSpreadsheet className="w-4 h-4 text-amber-800" />,
    },
    {
      id: 'certificat_vente',
      title: 'شهادة البيع (Certificat de Vente)',
      subtitle: 'شهادة إدارية رسمية تثبت واقعة البيع أو نقل الملكية',
      icon: <Sparkles className="w-4 h-4 text-purple-800" />,
    },
    {
      id: 'bordereau_envoi',
      title: "جدول الإرسال (Bordereau d'Envoi)",
      subtitle: 'جدول إيداع وإرسال المعاملات لمصالح الضرائب والشهر',
      icon: <Send className="w-4 h-4 text-cyan-800" />,
    },
  ];

  return (
    <div className="relative inline-block text-right" ref={dropdownRef}>
      {/* Trigger Button */}
      <button
        type="button"
        disabled={disabled}
        onClick={() => setIsOpen((prev) => !prev)}
        className="px-3 py-1.5 text-xs font-semibold text-white bg-[#1E3A8A] rounded-lg hover:bg-[#0F2744] transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 shadow-2xs disabled:opacity-50"
        title="تصدير وثائق العقد بصيغة Word (.docx) بحقن بيانات العقد في القوالب التوثيقية المعتمدة"
      >
        <Download className="w-3.5 h-3.5" />
        <span>تصدير الوثائق (Word)</span>
        <ChevronDown
          className={`w-3.5 h-3.5 transition-transform duration-200 ${
            isOpen ? 'rotate-180' : ''
          }`}
        />
      </button>

      {/* Dropdown Menu */}
      {isOpen && (
        <div className="absolute left-0 sm:right-auto sm:left-0 mt-1.5 w-80 sm:w-96 bg-white rounded-xl shadow-2xl border border-slate-200 z-50 overflow-hidden animate-in fade-in-50 zoom-in-95 duration-150">
          {/* Header */}
          <div className="px-3.5 py-2.5 bg-slate-50 border-b border-slate-100 flex items-center justify-between">
            <span className="text-[11px] font-bold text-slate-800">
              اختر نوع الوثيقة وقالب التصدير:
            </span>
            <span className="text-[10px] text-blue-900 bg-blue-50 px-2 py-0.5 rounded font-mono font-semibold">
              الوسوم المعتمدة
            </span>
          </div>

          {/* Section 1: 6 Built-in / Customized Standard Templates */}
          <div className="p-1.5 space-y-0.5 max-h-[60vh] overflow-y-auto">
            <div className="px-2.5 py-1 text-[10px] font-bold text-slate-400">
              القوالب التوثيقية الستة (بحقن المتغيرات الآلي)
            </div>

            {items.map((item) => {
              const isCustom = hasCustomizedNotaryTemplate(item.id);
              return (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => {
                    setIsOpen(false);
                    onExportWithTemplate(item.id);
                  }}
                  className="w-full text-right px-2.5 py-2 hover:bg-slate-50 rounded-lg flex items-start gap-2.5 transition-colors group"
                >
                  <div className="p-1.5 bg-slate-100 group-hover:bg-blue-50 rounded-md shrink-0 mt-0.5 transition-colors">
                    {item.icon}
                  </div>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center justify-between gap-1">
                      <span className="text-xs font-bold text-slate-900 group-hover:text-blue-900 truncate">
                        {item.title}
                      </span>
                      {isCustom && (
                        <span className="text-[9px] font-bold px-1.5 py-0.2 bg-amber-100 text-amber-900 rounded shrink-0">
                          مخصص
                        </span>
                      )}
                    </div>
                    <p className="text-[11px] text-slate-500 line-clamp-1 mt-0.5">
                      {item.subtitle}
                    </p>
                  </div>
                </button>
              );
            })}

            {/* Direct Export Option */}
            <div className="my-1 border-t border-slate-100" />
            <div className="px-2.5 py-1 text-[10px] font-bold text-slate-400">
              تصدير مباشر
            </div>
            <button
              type="button"
              onClick={() => {
                setIsOpen(false);
                onExportDirectDocx();
              }}
              className="w-full text-right px-2.5 py-2 hover:bg-slate-50 rounded-lg flex items-start gap-2.5 transition-colors group"
            >
              <div className="p-1.5 bg-slate-100 group-hover:bg-slate-200 rounded-md shrink-0 mt-0.5 transition-colors">
                <FileCode className="w-4 h-4 text-slate-700" />
              </div>
              <div className="flex-1 min-w-0">
                <span className="text-xs font-bold text-slate-900 group-hover:text-blue-900">
                  تصدير متن العقد الحالي مباشرة (.docx)
                </span>
                <p className="text-[11px] text-slate-500 mt-0.5">
                  تصدير محتوى المحرر كما يظهر على ورقة A4 بدون غلاف قالب
                </p>
              </div>
            </button>
          </div>

          {/* Footer with Manage Templates link */}
          <div className="p-2 bg-slate-50 border-t border-slate-100 flex items-center justify-between">
            <button
              type="button"
              onClick={() => {
                setIsOpen(false);
                onOpenTemplatesManager();
              }}
              className="w-full py-1.5 px-3 bg-white hover:bg-slate-100 text-slate-700 border border-slate-200 rounded-lg text-xs font-semibold inline-flex items-center justify-center gap-1.5 transition-colors"
            >
              <Settings className="w-3.5 h-3.5 text-blue-900" />
              <span>إدارة وتخصيص قوالب الوورد (.docx)...</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
