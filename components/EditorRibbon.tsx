'use client';

import React, { useState } from 'react';
import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  Bold,
  ChevronDown,
  Columns,
  Download,
  Edit3,
  FileOutput,
  FormInput,
  GitCompare,
  Highlighter,
  Image as ImageIcon,
  IndentDecrease,
  IndentIncrease,
  Italic,
  ListOrdered,
  Merge,
  Palette,
  PanelTopBottomDashed,
  PilcrowLeft,
  PilcrowRight,
  Redo2,
  Rows,
  Search,
  SeparatorHorizontal,
  Split,
  Table as TableIcon,
  Trash2,
  Underline,
  Undo2,
  Upload,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import {
  DerivedDocTemplate,
  ListNumberingStyle,
  PartyField,
  ToolbarState,
} from '@/lib/types';

interface EditorRibbonProps {
  toolbarState: ToolbarState;
  canUndo: boolean;
  canRedo: boolean;
  zoom: number;
  showHeaderFooter: boolean;
  pageNumberingEnabled: boolean;
  showFindReplace: boolean;
  unfilledCount: number;
  totalVariablesCount: number;
  partyFields: PartyField[];
  extractedPlaceholders: string[];
  fieldValues: Record<string, string>;
  derivedTemplates: DerivedDocTemplate[];
  activeEditingDerivedId?: string | null;
  onOpenSmartVariablesModal: () => void;
  onConvertSelectionToSmartTag: () => void;
  onInsertSmartTagAtCaret: (varKey: string) => void;
  onInsertNewClauseHeadingAtCaret: () => void;
  onOpenVersionDiffModal: () => void;
  onGenerateDerivedDoc: (tpl: DerivedDocTemplate) => void;
  onEditDerivedTemplateInEditor: (tpl: DerivedDocTemplate) => void;
  onSaveCurrentAsDerivedTemplate?: (
    name: string,
    code: string,
    description: string
  ) => Promise<void>;
  onImportDocxForDerivedTemplate?: (
    tpl: DerivedDocTemplate,
    file: File
  ) => Promise<void>;
  onDeleteDerivedTemplate?: (id: string) => Promise<void>;
  onOpenWordTemplatesModal?: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onToggleBold: () => void;
  onToggleItalic: () => void;
  onToggleUnderline: () => void;
  onSetTextColor: (color: string) => void;
  onSetHighlightColor: (color: string) => void;
  onSetAlign: (align: 'right' | 'center' | 'left' | 'justify') => void;
  onSetDirection: (dir: 'rtl' | 'ltr') => void;
  onAdjustIndent: (delta: 'increase' | 'decrease') => void;
  onToggleList: (style: ListNumberingStyle) => void;
  onInsertTable: (rows: number, cols: number) => void;
  onTableAction: (
    action:
      | 'row-above'
      | 'row-below'
      | 'delete-row'
      | 'col-before'
      | 'col-after'
      | 'delete-col'
      | 'merge-cells'
      | 'split-cell'
      | 'toggle-border'
      | 'shade-cell'
  ) => void;
  onInsertPageBreak: () => void;
  onInsertImageFile: (file: File) => void;
  onToggleHeaderFooter: () => void;
  onTogglePageNumbering: () => void;
  onToggleFindReplace: () => void;
  onChangeZoom: (newZoom: number) => void;
  onSaveSelectionBookmark: () => void;
}

const TEXT_COLORS = [
  { label: 'أسود رسمي', value: '#000000' },
  { label: 'كحلي توثيقي', value: '#1E3A8A' },
  { label: 'أحمر قانوني', value: '#991B1B' },
  { label: 'أخضر داكن', value: '#065F46' },
  { label: 'رمادي داكن', value: '#334155' },
];

const HIGHLIGHT_COLORS = [
  { label: 'بدون إبراز', value: 'transparent' },
  { label: 'أصفر فاتح', value: '#FEF08A' },
  { label: 'أخضر فاتح', value: '#BBF7D0' },
  { label: 'سماوي فاتح', value: '#BAE6FD' },
  { label: 'وردي فاتح', value: '#FBCFE8' },
];

export default function EditorRibbon({
  toolbarState,
  canUndo,
  canRedo,
  zoom,
  showHeaderFooter,
  pageNumberingEnabled,
  showFindReplace,
  unfilledCount,
  totalVariablesCount,
  partyFields,
  extractedPlaceholders,
  fieldValues,
  derivedTemplates,
  activeEditingDerivedId,
  onOpenSmartVariablesModal,
  onConvertSelectionToSmartTag,
  onInsertSmartTagAtCaret,
  onInsertNewClauseHeadingAtCaret,
  onOpenVersionDiffModal,
  onGenerateDerivedDoc,
  onEditDerivedTemplateInEditor,
  onSaveCurrentAsDerivedTemplate,
  onImportDocxForDerivedTemplate,
  onDeleteDerivedTemplate,
  onOpenWordTemplatesModal,
  onUndo,
  onRedo,
  onToggleBold,
  onToggleItalic,
  onToggleUnderline,
  onSetTextColor,
  onSetHighlightColor,
  onSetAlign,
  onSetDirection,
  onAdjustIndent,
  onToggleList,
  onInsertTable,
  onTableAction,
  onInsertPageBreak,
  onInsertImageFile,
  onToggleHeaderFooter,
  onTogglePageNumbering,
  onToggleFindReplace,
  onChangeZoom,
  onSaveSelectionBookmark,
}: EditorRibbonProps) {
  const [showTablePicker, setShowTablePicker] = useState(false);
  const [tableRows, setTableRows] = useState(3);
  const [tableCols, setTableCols] = useState(4);
  const [showColorMenu, setShowColorMenu] = useState(false);
  const [showHighlightMenu, setShowHighlightMenu] = useState(false);
  const [showListMenu, setShowListMenu] = useState(false);
  const [showDerivedMenu, setShowDerivedMenu] = useState(false);
  const [showTagInserterMenu, setShowTagInserterMenu] = useState(false);
  const [newDerivedName, setNewDerivedName] = useState('');
  const [tagFilterQuery, setTagFilterQuery] = useState('');
  const [isFormattingBarVisible, setIsFormattingBarVisible] = useState(true);
  const [showMoreMenu, setShowMoreMenu] = useState(false);

  const preventFocusLoss = (e: React.MouseEvent) => {
    e.preventDefault();
  };

  const closeAllMenus = () => {
    setShowTablePicker(false);
    setShowColorMenu(false);
    setShowHighlightMenu(false);
    setShowListMenu(false);
    setShowDerivedMenu(false);
    setShowTagInserterMenu(false);
  };

  const btnClass = (active: boolean = false, disabled: boolean = false) =>
    `inline-flex items-center justify-center h-7 px-2 rounded-md text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
      disabled
        ? 'text-slate-300 cursor-not-allowed'
        : active
        ? 'bg-blue-900 text-white shadow-2xs font-bold'
        : 'text-slate-700 hover:bg-slate-200/80 hover:text-slate-900'
    }`;

  return (
    <div className="bg-white border-b border-slate-200 select-none no-print">
      {/* المستوى 1: شريط إجراءات الملف */}
      <div className="flex items-center justify-between px-3 py-2 border-b border-slate-100">
        <div className="flex items-center gap-1.5">
          <button className="text-sm font-bold text-slate-900 px-2">الموثق الرقمي</button>
          <button className="px-2 py-1 text-xs hover:bg-slate-100 rounded">عقد جديد</button>
          <button className="px-2 py-1 text-xs hover:bg-slate-100 rounded">فتح Word</button>
          <button className="px-2 py-1 text-xs hover:bg-slate-100 rounded">معاينة</button>
          <button className="px-2 py-1 text-xs hover:bg-slate-100 rounded">طباعة</button>
          <button className="px-2 py-1 text-xs hover:bg-slate-100 rounded">تصدير Word</button>
        </div>
      </div>

      {/* المستوى 2: شريط التنسيق (قابل للطي) */}
      <div className="border-b border-slate-100 bg-slate-50">
        <button 
            onClick={() => setIsFormattingBarVisible(!isFormattingBarVisible)}
            className="px-3 py-1 text-[10px] text-slate-500 hover:text-slate-800"
        >
            {isFormattingBarVisible ? 'إخفاء التنسيق' : 'إظهار التنسيق'}
        </button>
        {isFormattingBarVisible && (
            <div className="flex flex-wrap items-center gap-1 p-2">
                {/* Formatting tools here... (B I U, Colors, Align, List, Table, etc.) */}
            </div>
        )}
      </div>

      {/* المستوى 3: شريط الأدوات التوثيقية */}
      <div className="flex flex-wrap items-center gap-1.5 px-3 py-2 bg-white">
        {/* Doc tools here... (Variables, Tags, Clauses, etc.) */}
      </div>
          {/* Strict Standard Indicator (Arial 13pt, 1.0 spacing, 7/2/1/6cm margins) */}
          <div
            className="flex items-center gap-1.5 px-2 py-0.5 bg-slate-50 border border-slate-200 rounded text-[11px] text-slate-700 shrink-0"
            title="معيار المكتب الثابت: الخط Arial بحجم 13pt، تباعد مفرد 1.0 بدون مسافات فقرات، الهوامش 7سم يمين / 2سم يسار / 1سم أعلى / 6سم أسفل"
          >
            <span className="font-semibold text-slate-900">Arial · 13pt</span>
            <span aria-hidden="true" className="text-slate-300 hidden md:inline">
              |
            </span>
            <span className="tabular-nums text-slate-500 hidden md:inline">
              7/2/1/6 سم
            </span>
          </div>

          <div className="h-4 w-px bg-slate-200 mx-0.5" />

          {/* Undo / Redo */}
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onUndo}
            disabled={!canUndo}
            className={btnClass(false, !canUndo)}
            title="تراجع (Ctrl+Z)"
          >
            <Undo2 className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onRedo}
            disabled={!canRedo}
            className={btnClass(false, !canRedo)}
            title="إعادة (Ctrl+Y)"
          >
            <Redo2 className="w-4 h-4" />
          </button>

          <div className="h-4 w-px bg-slate-200 mx-0.5" />

          {/* Bold / Italic / Underline */}
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onToggleBold}
            className={btnClass(toolbarState.bold)}
            title="خط عريض (Ctrl+B)"
          >
            <Bold className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onToggleItalic}
            className={btnClass(toolbarState.italic)}
            title="خط مائل (Ctrl+I)"
          >
            <Italic className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onToggleUnderline}
            className={btnClass(toolbarState.underline)}
            title="تسطير (Ctrl+U)"
          >
            <Underline className="w-4 h-4" />
          </button>

          {/* Text Color Picker */}
          <div className="relative">
            <button
              type="button"
              onMouseDown={(e) => {
                e.preventDefault();
                onSaveSelectionBookmark();
              }}
              onClick={() => {
                const next = !showColorMenu;
                closeAllMenus();
                setShowColorMenu(next);
              }}
              className={btnClass(showColorMenu)}
              title="لون النص"
            >
              <Palette className="w-4 h-4 ml-1" />
              <span
                className="w-2.5 h-2.5 rounded-sm border border-slate-300"
                style={{ backgroundColor: toolbarState.textColor || '#000000' }}
              />
            </button>
            {showColorMenu && (
              <div className="absolute right-0 mt-1 w-40 bg-white border border-slate-200 rounded-md shadow-lg p-1.5 z-50">
                {TEXT_COLORS.map((c) => (
                  <button
                    key={c.value}
                    type="button"
                    onMouseDown={preventFocusLoss}
                    onClick={() => {
                      onSetTextColor(c.value);
                      setShowColorMenu(false);
                    }}
                    className="w-full flex items-center gap-2 px-2 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded text-right"
                  >
                    <span
                      className="w-3.5 h-3.5 rounded-sm border border-slate-300 shrink-0"
                      style={{ backgroundColor: c.value }}
                    />
                    <span>{c.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>

          {/* Highlight Color Picker */}
          <div className="relative">
            <button
              type="button"
              onMouseDown={(e) => {
                e.preventDefault();
                onSaveSelectionBookmark();
              }}
              onClick={() => {
                const next = !showHighlightMenu;
                closeAllMenus();
                setShowHighlightMenu(next);
              }}
              className={btnClass(showHighlightMenu)}
              title="تمييز خلفية النص"
            >
              <Highlighter className="w-4 h-4" />
            </button>
            {showHighlightMenu && (
              <div className="absolute right-0 mt-1 w-40 bg-white border border-slate-200 rounded-md shadow-lg p-1.5 z-50">
                {HIGHLIGHT_COLORS.map((c) => (
                  <button
                    key={c.value}
                    type="button"
                    onMouseDown={preventFocusLoss}
                    onClick={() => {
                      onSetHighlightColor(c.value);
                      setShowHighlightMenu(false);
                    }}
                    className="w-full flex items-center gap-2 px-2 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded text-right"
                  >
                    <span
                      className="w-3.5 h-3.5 rounded-sm border border-slate-300 shrink-0"
                      style={{
                        backgroundColor:
                          c.value === 'transparent' ? '#ffffff' : c.value,
                      }}
                    />
                    <span>{c.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className="h-4 w-px bg-slate-200 mx-0.5" />

          {/* Paragraph Alignments */}
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onSetAlign('right')}
            className={btnClass(toolbarState.align === 'right')}
            title="محاذاة إلى اليمين"
          >
            <AlignRight className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onSetAlign('center')}
            className={btnClass(toolbarState.align === 'center')}
            title="توسيط"
          >
            <AlignCenter className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onSetAlign('left')}
            className={btnClass(toolbarState.align === 'left')}
            title="محاذاة إلى اليسار"
          >
            <AlignLeft className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onSetAlign('justify')}
            className={btnClass(toolbarState.align === 'justify')}
            title="ضبط كلي للفقرة (Justify)"
          >
            <AlignJustify className="w-4 h-4" />
          </button>

          {/* Paragraph Direction RTL / LTR */}
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onSetDirection('rtl')}
            className={btnClass(toolbarState.dir === 'rtl')}
            title="اتجاه الفقرة من اليمين إلى اليسار (RTL)"
          >
            <PilcrowLeft className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onSetDirection('ltr')}
            className={btnClass(toolbarState.dir === 'ltr')}
            title="اتجاه الفقرة من اليسار إلى اليمين (LTR)"
          >
            <PilcrowRight className="w-4 h-4" />
          </button>

          {/* Indentation */}
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onAdjustIndent('increase')}
            className={btnClass(false)}
            title="زيادة المسافة البادئة"
          >
            <IndentIncrease className="w-4 h-4" />
          </button>
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => onAdjustIndent('decrease')}
            className={btnClass(false)}
            title="إنقاص المسافة البادئة"
          >
            <IndentDecrease className="w-4 h-4" />
          </button>

          <div className="h-4 w-px bg-slate-200 mx-0.5" />

          {/* Lists Menu — Renamed to "قائمة" to prevent confusion with Page Numbering */}
          <div className="relative">
            <button
              type="button"
              onMouseDown={(e) => {
                e.preventDefault();
                onSaveSelectionBookmark();
              }}
              onClick={() => {
                const next = !showListMenu;
                closeAllMenus();
                setShowListMenu(next);
              }}
              className={btnClass(toolbarState.listType !== 'none' || showListMenu)}
              title="إدراج قائمة مرقمة أو أبجدية عربية أو نقطية"
            >
              <ListOrdered className="w-4 h-4 ml-1" />
              <span>قائمة</span>
              <ChevronDown className="w-3 h-3 mr-0.5" />
            </button>
            {showListMenu && (
              <div className="absolute right-0 mt-1 w-48 bg-white border border-slate-200 rounded-md shadow-lg p-1.5 z-50">
                <button
                  type="button"
                  onMouseDown={preventFocusLoss}
                  onClick={() => {
                    onToggleList('decimal');
                    setShowListMenu(false);
                  }}
                  className="w-full flex items-center justify-between px-2.5 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded"
                >
                  <span>قائمة رقمية</span>
                  <span className="font-mono text-slate-500">1. 2. 3.</span>
                </button>
                <button
                  type="button"
                  onMouseDown={preventFocusLoss}
                  onClick={() => {
                    onToggleList('arabic-alpha');
                    setShowListMenu(false);
                  }}
                  className="w-full flex items-center justify-between px-2.5 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded"
                >
                  <span>أبجدي عربي</span>
                  <span className="text-slate-500">أ- ب- ج-</span>
                </button>
                <button
                  type="button"
                  onMouseDown={preventFocusLoss}
                  onClick={() => {
                    onToggleList('arabic-abjad');
                    setShowListMenu(false);
                  }}
                  className="w-full flex items-center justify-between px-2.5 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded"
                >
                  <span>ترقيم أبجد هوز</span>
                  <span className="text-slate-500">أبجد- هوز-</span>
                </button>
                <button
                  type="button"
                  onMouseDown={preventFocusLoss}
                  onClick={() => {
                    onToggleList('bullet');
                    setShowListMenu(false);
                  }}
                  className="w-full flex items-center justify-between px-2.5 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded"
                >
                  <span>قائمة نقطية</span>
                  <span className="text-slate-500">• • •</span>
                </button>
                <button
                  type="button"
                  onMouseDown={preventFocusLoss}
                  onClick={() => {
                    onToggleList('dash');
                    setShowListMenu(false);
                  }}
                  className="w-full flex items-center justify-between px-2.5 py-1.5 text-xs text-slate-700 hover:bg-slate-100 rounded"
                >
                  <span>قائمة بعلامة مطة</span>
                  <span className="text-slate-500">- - -</span>
                </button>
              </div>
            )}
          </div>

          {/* Table Creator */}
          <div className="relative">
            <button
              type="button"
              onMouseDown={(e) => {
                e.preventDefault();
                onSaveSelectionBookmark();
              }}
              onClick={() => {
                const next = !showTablePicker;
                closeAllMenus();
                setShowTablePicker(next);
              }}
              className={btnClass(toolbarState.inTable || showTablePicker)}
              title="إدراج وإدارة الجداول"
            >
              <TableIcon className="w-4 h-4 ml-1" />
              <span>جدول</span>
            </button>
            {showTablePicker && (
              <div className="absolute right-0 mt-1 w-64 bg-white border border-slate-200 rounded-md shadow-lg p-3 z-50">
                <div className="text-xs font-semibold text-slate-800 mb-2">
                  إدراج جدول توثيقي (بعرض صافي 12 سم)
                </div>
                <div className="grid grid-cols-2 gap-2 mb-2.5">
                  <div>
                    <label className="block text-[11px] text-slate-500 mb-1">
                      عدد الصفوف
                    </label>
                    <input
                      type="number"
                      min={1}
                      max={50}
                      value={tableRows}
                      onChange={(e) =>
                        setTableRows(Math.max(1, parseInt(e.target.value, 10) || 1))
                      }
                      className="w-full px-2 py-1 text-xs border border-slate-200 rounded tabular-nums"
                    />
                  </div>
                  <div>
                    <label className="block text-[11px] text-slate-500 mb-1">
                      عدد الأعمدة
                    </label>
                    <input
                      type="number"
                      min={1}
                      max={12}
                      value={tableCols}
                      onChange={(e) =>
                        setTableCols(Math.max(1, parseInt(e.target.value, 10) || 1))
                      }
                      className="w-full px-2 py-1 text-xs border border-slate-200 rounded tabular-nums"
                    />
                  </div>
                </div>
                <button
                  type="button"
                  onMouseDown={preventFocusLoss}
                  onClick={() => {
                    onInsertTable(tableRows, tableCols);
                    setShowTablePicker(false);
                  }}
                  className="w-full py-1.5 bg-blue-900 text-white text-xs font-medium rounded hover:bg-blue-800 transition-colors"
                >
                  إدراج الجدول عند المؤشر
                </button>
              </div>
            )}
          </div>

          {/* Contextual Table Tools when Caret is inside a Table */}
          {toolbarState.inTable && (
            <div className="flex items-center gap-1 bg-blue-50/70 border border-blue-200 rounded px-1.5 py-0.5">
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('row-below')}
                className="px-1.5 py-0.5 text-[11px] text-blue-950 hover:bg-blue-100 rounded inline-flex items-center gap-1"
                title="إضافة صف لأسفل"
              >
                <Rows className="w-3.5 h-3.5" />
                <span>+صف</span>
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('col-after')}
                className="px-1.5 py-0.5 text-[11px] text-blue-950 hover:bg-blue-100 rounded inline-flex items-center gap-1"
                title="إضافة عمود"
              >
                <Columns className="w-3.5 h-3.5" />
                <span>+عمود</span>
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('merge-cells')}
                className="px-1.5 py-0.5 text-[11px] text-blue-950 hover:bg-blue-100 rounded inline-flex items-center gap-1"
                title="دمج الخلايا المحددة"
              >
                <Merge className="w-3.5 h-3.5" />
                <span>دمج</span>
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('split-cell')}
                className="px-1.5 py-0.5 text-[11px] text-blue-950 hover:bg-blue-100 rounded inline-flex items-center gap-1"
                title="فك دمج الخلية الحالية"
              >
                <Split className="w-3.5 h-3.5" />
                <span>فك</span>
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('toggle-border')}
                className="px-1.5 py-0.5 text-[11px] text-blue-950 hover:bg-blue-100 rounded"
                title="إظهار/إخفاء حدود الخلية"
              >
                الحدود
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('shade-cell')}
                className="px-1.5 py-0.5 text-[11px] text-blue-950 hover:bg-blue-100 rounded"
                title="تظليل رمادي للخلية"
              >
                تظليل
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('delete-row')}
                className="px-1.5 py-0.5 text-[11px] text-red-700 hover:bg-red-100 rounded inline-flex items-center gap-0.5"
                title="حذف الصف الحالي"
              >
                <Trash2 className="w-3 h-3" />
                <span>صف</span>
              </button>
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => onTableAction('delete-col')}
                className="px-1.5 py-0.5 text-[11px] text-red-700 hover:bg-red-100 rounded inline-flex items-center gap-0.5"
                title="حذف العمود الحالي"
              >
                <Trash2 className="w-3 h-3" />
                <span>عمود</span>
              </button>
            </div>
          )}

          <div className="h-4 w-px bg-slate-200 mx-0.5" />

          {/* Page Break, Image, Header/Footer, Page Numbering ("رقم الصفحة") */}
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onInsertPageBreak}
            className={btnClass(false)}
            title="إدراج فاصل صفحات"
          >
            <SeparatorHorizontal className="w-4 h-4 ml-1" />
            <span className="hidden xl:inline">فاصل</span>
          </button>

          <label
            className={btnClass(false) + ' cursor-pointer'}
            title="إدراج صورة أو ختم مضمن"
            onMouseDown={onSaveSelectionBookmark}
          >
            <ImageIcon className="w-4 h-4" />
            <input
              type="file"
              accept="image/png,image/jpeg,image/webp"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  onInsertImageFile(file);
                  e.target.value = '';
                }
              }}
            />
          </label>

          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onToggleHeaderFooter}
            className={btnClass(showHeaderFooter)}
            title="إظهار/إخفاء الترويسة والتذييل"
          >
            <PanelTopBottomDashed className="w-4 h-4 ml-1" />
            <span className="hidden 2xl:inline">ترويسة</span>
          </button>

          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onTogglePageNumbering}
            className={btnClass(pageNumberingEnabled)}
            title="تفعيل أو إيقاف ترقيم الصفحات في أسفل الورقة"
          >
            <span>رقم الصفحة</span>
          </button>

        <div className="flex items-center gap-1.5 shrink-0">
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={onToggleFindReplace}
            className={btnClass(showFindReplace)}
            title="بحث واستبدال داخل العقد (Ctrl+F)"
          >
            <Search className="w-3.5 h-3.5 ml-1" />
            <span className="hidden sm:inline">بحث</span>
          </button>

          <div className="flex items-center gap-0.5 border border-slate-200 rounded px-1 py-0.5 bg-slate-50">
            <button
              type="button"
              onMouseDown={preventFocusLoss}
              onClick={() => onChangeZoom(Math.max(60, zoom - 10))}
              className="p-1 text-slate-600 hover:text-slate-900"
              title="تصغير العرض"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="text-[11px] font-mono tabular-nums text-slate-700 w-9 text-center">
              {zoom}%
            </span>
            <button
              type="button"
              onMouseDown={preventFocusLoss}
              onClick={() => onChangeZoom(Math.min(140, zoom + 10))}
              className="p-1 text-slate-600 hover:text-slate-900"
              title="تكبير العرض"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

      {/* =====================================================================
          السطر الثاني: شريط الأوامر التوثيقية المخصص (الاستمارة · الوسوم · البنود · المشتقات · المقارنة)
         ===================================================================== */}
      <div className="flex flex-wrap items-center gap-1.5 px-3 py-1 bg-slate-50/80">
        {/* SMART VARIABLES FORM BUTTON WITH UNFILLED COUNTER BADGE */}
        <button
          type="button"
          onMouseDown={preventFocusLoss}
          onClick={onOpenSmartVariablesModal}
          className={`inline-flex items-center gap-1.5 h-7 px-2.5 rounded text-xs font-semibold border transition-colors whitespace-nowrap shrink-0 ${
            unfilledCount > 0
              ? 'bg-amber-50 border-amber-300 text-amber-950 hover:bg-amber-100'
              : 'bg-emerald-50 border-emerald-300 text-emerald-950 hover:bg-emerald-100'
          }`}
          title="فتح استمارة المتغيرات الموحدة لتعبئة بيانات العقد والوثائق المشتقة"
        >
          <FormInput className="w-3.5 h-3.5 text-pink-800" />
          <span>استمارة المتغيرات</span>
          {unfilledCount > 0 ? (
            <span className="px-1.5 py-0.2 text-[10px] font-bold bg-red-600 text-white rounded tabular-nums">
              {unfilledCount} غير معبأة
            </span>
          ) : totalVariablesCount > 0 ? (
            <span className="px-1.5 py-0.2 text-[10px] font-bold bg-emerald-700 text-white rounded">
              مكتمل ✓
            </span>
          ) : null}
        </button>

        {/* 1-CLICK SELECTION TO SMART TAG CONVERTER */}
        <button
          type="button"
          onMouseDown={(e) => {
            e.preventDefault();
            onSaveSelectionBookmark();
          }}
          onClick={onConvertSelectionToSmartTag}
          className="inline-flex items-center gap-1 h-7 px-2.5 rounded text-xs font-bold bg-pink-800 text-white hover:bg-pink-900 transition-colors whitespace-nowrap shrink-0 shadow-2xs"
          title="حدد أي كلمة أو جملة داخل العقد ثم اضغط هنا لتحويلها فوراً إلى وسم متغير [...] دون فتح أي نافذة (Alt+V)"
        >
          <span className="font-mono text-[10px] bg-pink-950/40 px-1 rounded">[ ]</span>
          <span>تحويل المحدد لوسم</span>
        </button>

        {/* QUICK SMART TAG INSERTER DROPDOWN */}
        <div className="relative">
          <button
            type="button"
            onMouseDown={(e) => {
              e.preventDefault();
              onSaveSelectionBookmark();
            }}
            onClick={() => {
              const sel = typeof window !== 'undefined' ? window.getSelection() : null;
              if (sel && !sel.isCollapsed && sel.toString().trim().length > 0) {
                closeAllMenus();
                onConvertSelectionToSmartTag();
                return;
              }
              const next = !showTagInserterMenu;
              closeAllMenus();
              setShowTagInserterMenu(next);
            }}
            className="inline-flex items-center gap-1 h-7 px-2 rounded text-xs font-semibold bg-white border border-pink-200 text-pink-900 hover:bg-pink-50 transition-colors whitespace-nowrap shrink-0"
            title="إدراج وسم متغير جاهز [...] أو إنشاء وسم جديد عند موضع المؤشر"
          >
            <span className="font-mono font-bold text-[11px]">[ ]</span>
            <span>إدراج وسم</span>
            <ChevronDown className="w-3.5 h-3.5" />
          </button>

          {showTagInserterMenu && (
            <div className="absolute right-0 mt-1 w-72 bg-white border border-slate-200 rounded-md shadow-xl p-2.5 z-50 space-y-2">
              <button
                type="button"
                onMouseDown={preventFocusLoss}
                onClick={() => {
                  onConvertSelectionToSmartTag();
                  setShowTagInserterMenu(false);
                }}
                className="w-full py-1.5 px-2.5 bg-pink-800 hover:bg-pink-900 text-white text-xs font-bold rounded flex items-center justify-between"
              >
                <span>تحويل الكلمة المحددة/الحالية إلى وسم</span>
                <span className="font-mono text-[10px] bg-pink-950/40 px-1.5 py-0.5 rounded">
                  [ ]
                </span>
              </button>
              <div className="text-[11px] font-bold text-slate-900 pt-1 border-t border-slate-100">
                أو أدرج وسماً جاهزاً عند موضع المؤشر:
              </div>
              <div className="flex items-center gap-1.5">
                <input
                  type="text"
                  value={tagFilterQuery}
                  onChange={(e) => setTagFilterQuery(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && tagFilterQuery.trim()) {
                      e.preventDefault();
                      const cleanKey = tagFilterQuery.trim().replace(/\s+/g, '_');
                      onInsertSmartTagAtCaret(cleanKey);
                      setTagFilterQuery('');
                      setShowTagInserterMenu(false);
                    }
                  }}
                  placeholder="ابحث أو اكتب اسم وسم جديد..."
                  className="flex-1 px-2 py-1 text-xs border border-slate-300 rounded focus:border-pink-700 focus:outline-none"
                />
                {tagFilterQuery.trim() && (
                  <button
                    type="button"
                    onMouseDown={preventFocusLoss}
                    onClick={() => {
                      const cleanKey = tagFilterQuery.trim().replace(/\s+/g, '_');
                      onInsertSmartTagAtCaret(cleanKey);
                      setTagFilterQuery('');
                      setShowTagInserterMenu(false);
                    }}
                    className="px-2 py-1 bg-pink-800 text-white text-[11px] font-semibold rounded hover:bg-pink-900 shrink-0"
                  >
                    + إدراج
                  </button>
                )}
              </div>

              <div className="max-h-56 overflow-y-auto divide-y divide-slate-100 border border-slate-100 rounded">
                {Array.from(
                  new Set([
                    ...extractedPlaceholders,
                    ...partyFields.map((f) => f.key),
                  ])
                )
                  .filter((k) => {
                    if (!tagFilterQuery.trim()) return true;
                    const q = tagFilterQuery.trim();
                    const meta = partyFields.find((f) => f.key === k);
                    return k.includes(q) || (meta?.label && meta.label.includes(q));
                  })
                  .map((key) => {
                    const meta = partyFields.find((f) => f.key === key);
                    const label = meta?.label || key.replace(/_/g, ' ');
                    const isFilled = Boolean(
                      fieldValues[key] && fieldValues[key].trim()
                    );
                    return (
                      <button
                        key={key}
                        type="button"
                        onMouseDown={preventFocusLoss}
                        onClick={() => {
                          onInsertSmartTagAtCaret(key);
                          setShowTagInserterMenu(false);
                        }}
                        className="w-full flex items-center justify-between gap-2 px-2 py-1.5 text-right hover:bg-slate-50 transition-colors"
                      >
                        <span className="text-xs font-medium text-slate-800 truncate">
                          {label}
                        </span>
                        <span
                          className={`text-[10px] font-mono px-1.5 py-0.5 rounded shrink-0 ${
                            isFilled
                              ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                              : 'bg-pink-50 text-pink-800 border border-pink-200'
                          }`}
                        >
                          {`[${key}]`}
                        </span>
                      </button>
                    );
                  })}
              </div>
            </div>
          )}
        </div>

        {/* # NEW CLAUSE HEADING & CONTAINER AT CARET */}
        <button
          type="button"
          onMouseDown={(e) => {
            e.preventDefault();
            onSaveSelectionBookmark();
          }}
          onClick={onInsertNewClauseHeadingAtCaret}
          className="inline-flex items-center gap-1 h-7 px-2.5 rounded text-xs font-semibold bg-slate-800 text-white hover:bg-slate-700 transition-colors whitespace-nowrap shrink-0"
          title="تحويل السطر/التحديد الحالي إلى عنوان بند جديد أو إدراج حاوية بند جديدة"
        >
          <span className="font-mono font-bold text-amber-300">#</span>
          <span>بند جديد</span>
        </button>

        {/* DERIVED DOCUMENTS GENERATOR & TEMPLATE MANAGER DROPDOWN */}
        <div className="relative">
          <button
            type="button"
            onMouseDown={preventFocusLoss}
            onClick={() => {
              const next = !showDerivedMenu;
              closeAllMenus();
              setShowDerivedMenu(next);
            }}
            className="inline-flex items-center gap-1.5 h-7 px-2.5 rounded text-xs font-semibold bg-blue-900 text-white hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0"
            title="توليد وإدارة الوثائق المشتقة (المستخرج، الشهر العقاري، شهادة البيع، الصيغة التنفيذية...)"
          >
            <FileOutput className="w-3.5 h-3.5" />
            <span>الوثائق المشتقة ({derivedTemplates.length})</span>
            <ChevronDown className="w-3.5 h-3.5" />
          </button>

          {showDerivedMenu && (
            <div className="absolute right-0 mt-1 w-96 max-w-[92vw] bg-white border border-slate-200 rounded-md shadow-xl p-2.5 z-50 space-y-2">
              {onOpenWordTemplatesModal && (
                <button
                  type="button"
                  onClick={() => {
                    onOpenWordTemplatesModal();
                    setShowDerivedMenu(false);
                  }}
                  className="w-full mb-2 py-1.5 px-2.5 bg-blue-50 hover:bg-blue-100 text-blue-950 text-xs font-bold rounded flex items-center justify-between border border-blue-200"
                >
                  <span>قوالب Word الرسمية والمشتقة (Original, Extract...)</span>
                  <span className="text-[10px] font-mono">⚙️ إدارة وقوالب</span>
                </button>
              )}

              <div className="px-1.5 py-1 text-[11px] font-bold text-slate-800 border-b border-slate-100 flex items-center justify-between">
                <span>توليد فوري (.docx) أو تعديل قوالب المشتقات:</span>
                <span className="text-[10px] font-normal text-slate-500">
                  تعتمد على استمارة المتغيرات
                </span>
              </div>

              <div className="max-h-72 overflow-y-auto divide-y divide-slate-100">
                {derivedTemplates.map((dt) => {
                  const isEditing = activeEditingDerivedId === dt.id;
                  return (
                    <div
                      key={dt.id}
                      className={`flex items-center justify-between gap-1.5 px-2 py-1.5 rounded transition-colors ${
                        isEditing ? 'bg-amber-50' : 'hover:bg-slate-50'
                      }`}
                    >
                      <button
                        type="button"
                        onClick={() => {
                          onGenerateDerivedDoc(dt);
                          setShowDerivedMenu(false);
                        }}
                        className="flex-1 text-right text-xs font-semibold text-slate-800 hover:text-blue-900 flex items-center gap-1.5 min-w-0"
                        title={`توليد وتصدير "${dt.name}" بصيغة Word (.docx)`}
                      >
                        <Download className="w-3.5 h-3.5 text-blue-800 shrink-0" />
                        <span className="truncate">{dt.name}</span>
                      </button>

                      <div className="flex items-center gap-1 shrink-0">
                        <button
                          type="button"
                          onClick={() => {
                            onEditDerivedTemplateInEditor(dt);
                            setShowDerivedMenu(false);
                          }}
                          className="px-1.5 py-0.5 text-[10px] text-slate-600 hover:text-blue-900 bg-slate-100 hover:bg-blue-50 rounded inline-flex items-center gap-0.5"
                          title="فتح وتعديل قالب هذه الوثيقة المشتقة داخل محرر A4"
                        >
                          <Edit3 className="w-3 h-3" />
                          <span>تعديل</span>
                        </button>

                        {onImportDocxForDerivedTemplate && (
                          <label
                            className="px-1.5 py-0.5 text-[10px] text-slate-600 hover:text-blue-900 bg-slate-100 hover:bg-blue-50 rounded cursor-pointer inline-flex items-center gap-0.5"
                            title="استبدال هذا القالب المشتق بملف Word (.docx) من جهازك"
                          >
                            <Upload className="w-3 h-3" />
                            <span>.docx</span>
                            <input
                              type="file"
                              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                              className="hidden"
                              onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) {
                                  onImportDocxForDerivedTemplate(dt, file);
                                  e.target.value = '';
                                  setShowDerivedMenu(false);
                                }
                              }}
                            />
                          </label>
                        )}

                        {onDeleteDerivedTemplate && !dt.id.startsWith('derived_') && (
                          <button
                            type="button"
                            onClick={() => onDeleteDerivedTemplate(dt.id)}
                            className="p-1 text-slate-400 hover:text-red-600"
                            title="حذف القالب المشتق المخصص"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>

              {onSaveCurrentAsDerivedTemplate && (
                <div className="pt-2 border-t border-slate-200 flex items-center gap-1.5">
                  <input
                    type="text"
                    value={newDerivedName}
                    onChange={(e) => setNewDerivedName(e.target.value)}
                    placeholder="اسم مشتق جديد لحفظ المحرر الحالي..."
                    className="flex-1 px-2 py-1 text-xs border border-slate-300 rounded focus:border-blue-900 focus:outline-none"
                  />
                  <button
                    type="button"
                    onClick={async () => {
                      if (!newDerivedName.trim()) return;
                      await onSaveCurrentAsDerivedTemplate(
                        newDerivedName.trim(),
                        `custom_${Date.now()}`,
                        'قالب وثيقة مشتقة مخصص للمكتب'
                      );
                      setNewDerivedName('');
                    }}
                    className="px-2.5 py-1 bg-slate-900 text-white text-[11px] font-medium rounded hover:bg-slate-800 shrink-0"
                  >
                    + حفظ كمشتق
                  </button>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Version Diff Button */}
        <button
          type="button"
          onMouseDown={preventFocusLoss}
          onClick={onOpenVersionDiffModal}
          className="inline-flex items-center gap-1 h-7 px-2 rounded text-xs font-medium text-slate-700 hover:bg-slate-200/80 transition-colors whitespace-nowrap shrink-0"
          title="مقارنة النسخ والتعديلات (Diff)"
        >
          <GitCompare className="w-3.5 h-3.5 text-blue-900" />
          <span>مقارنة النسخ</span>
        </button>
      </div>
    </div>
  );
}
