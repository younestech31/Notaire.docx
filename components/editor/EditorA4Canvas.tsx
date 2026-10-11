'use client';

import React from 'react';
import { Eye } from 'lucide-react';
import { STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from '@/lib/types';
import { buildNotarialPaginationClosingText } from '@/lib/editor-utils';

interface EditorA4CanvasProps {
  zoom: number;
  showHeaderFooter: boolean;
  pageNumberingEnabled: boolean;
  previewMergedMode: boolean;
  mergedPreviewHtml: string;
  printPaginatedHtml: string;
  isBodyLocked: boolean;
  estimatedPages: number;
  headerEditorRef: React.RefObject<HTMLDivElement | null>;
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  footerEditorRef: React.RefObject<HTMLDivElement | null>;
  onSyncPlaceholdersAndDraft: () => void;
  onDecorateEditorZonesPreservingSelection: () => void;
  onTogglePreviewMergedMode?: () => void;
  onEditorKeyDown: (e: React.KeyboardEvent<HTMLDivElement>) => void;
  onEditorPaste: (e: React.ClipboardEvent<HTMLDivElement>) => void;
  onEditorMouseDown: (e: React.MouseEvent<HTMLDivElement>) => void;
  onEditorClick: (e: React.MouseEvent<HTMLDivElement>) => void;
  onEditorDoubleClick: (e: React.MouseEvent<HTMLDivElement>) => void;
}

export default function EditorA4Canvas({
  zoom,
  showHeaderFooter,
  pageNumberingEnabled,
  previewMergedMode,
  mergedPreviewHtml,
  printPaginatedHtml,
  isBodyLocked,
  estimatedPages,
  headerEditorRef,
  bodyEditorRef,
  footerEditorRef,
  onSyncPlaceholdersAndDraft,
  onDecorateEditorZonesPreservingSelection,
  onTogglePreviewMergedMode,
  onEditorKeyDown,
  onEditorPaste,
  onEditorMouseDown,
  onEditorClick,
  onEditorDoubleClick,
}: EditorA4CanvasProps) {
  return (
    <main className="flex-1 min-w-0 min-h-0 overflow-auto print:overflow-visible p-4 sm:p-8 flex flex-col items-center bg-[#F7F5F0]">
      {/* Top Horizontal A4 Margin Ruler Indicator (21cm total: 7cm Right | 12cm Content | 2cm Left) */}
      <div
        className="mb-2 bg-white border border-[#C4A46A]/45 rounded-t shadow-2xs text-[10px] font-mono text-[#5C6B7A] flex items-center select-none overflow-hidden no-print"
        style={{
          width: '210mm',
          transform: `scale(${zoom / 100})`,
          transformOrigin: 'top center',
        }}
        dir="rtl"
      >
        <div
          style={{ width: '70mm' }}
          className="bg-[#F1F5F9] border-l border-dashed border-[#C4A46A] py-1 text-center text-[#5C6B7A] tabular-nums"
          title="هامش أيمن توثيقي ثابت: 7 سنتيمتر"
        >
          هامش يمين: 7 سم
        </div>
        <div
          style={{ width: '120mm' }}
          className="py-1 text-center font-semibold text-[#0F2744] tabular-nums flex items-center justify-between px-2"
        >
          <span>0</span>
          <span>المساحة النصية الصافية (12 سم · Arial 13pt · تباعد 1.0)</span>
          <span>12</span>
        </div>
        <div
          style={{ width: '20mm' }}
          className="bg-[#F1F5F9] border-r border-dashed border-[#C4A46A] py-1 text-center text-[#5C6B7A] tabular-nums"
          title="هامش أيسر توثيقي ثابت: 2 سنتيمتر"
        >
          2 سم
        </div>
      </div>

      {/* A4 Sheet + Vertical Margin Ruler Wrapper */}
      <div
        className="flex items-stretch gap-1.5"
        style={{
          transform: `scale(${zoom / 100})`,
          transformOrigin: 'top center',
          marginBottom: zoom < 100 ? `-${(100 - zoom) * 2.2}px` : '2rem',
        }}
      >
        {/* Vertical Ruler (29.7cm height: 1cm Top | 22.7cm Net Height | 6cm Bottom) */}
        <div
          className="hidden xl:flex flex-col justify-between bg-white border border-slate-300 rounded-r text-[9px] font-mono text-slate-500 w-6 select-none overflow-hidden no-print"
          style={{ minHeight: '297mm' }}
        >
          <div
            style={{ height: '10mm' }}
            className="bg-slate-100 border-b border-dashed border-slate-400 flex items-center justify-center tabular-nums"
            title="هامش علوي ثابت: 1 سم"
          >
            1سم
          </div>
          <div className="flex-1 flex flex-col items-center justify-around py-2 text-slate-400 tabular-nums">
            <span>2</span>
            <span>6</span>
            <span>10</span>
            <span>14</span>
            <span>18</span>
            <span>22</span>
          </div>
          <div
            style={{ height: '60mm' }}
            className="bg-slate-100 border-t border-dashed border-slate-400 flex items-center justify-center tabular-nums [writing-mode:vertical-rl]"
            title="هامش سفلي توثيقي ثابت: 6 سم"
          >
            هامش سفلي: 6 سم
          </div>
        </div>

        {/* TRUE A4 SHEET CONTAINER */}
        <div className="notary-a4-sheet shadow-xl border border-slate-300 transition-transform">
          {/* Visual Subtle Margin Guides on the A4 Sheet (Hidden on Print) */}
          <div
            aria-hidden="true"
            className="pointer-events-none absolute inset-0 no-print"
          >
            <div
              style={{ right: '70mm' }}
              className="absolute top-0 bottom-0 border-l border-dashed border-slate-200/80"
            />
            <div
              style={{ left: '20mm' }}
              className="absolute top-0 bottom-0 border-r border-dashed border-slate-200/80"
            />
            <div
              style={{ top: '10mm' }}
              className="absolute right-0 left-0 border-b border-dashed border-slate-200/80"
            />
            <div
              style={{ bottom: '60mm' }}
              className="absolute right-0 left-0 border-t border-dashed border-slate-200/80"
            />
          </div>

          {/* Optional Document Header Zone */}
          {showHeaderFooter && (
            <div className="mb-2 pb-1 border-b border-dashed border-slate-300 relative z-10">
              <div className="text-[9px] text-slate-400 mb-0.5 select-none no-print">
                ترويسة العقد (Header)
              </div>
              <div
                ref={headerEditorRef}
                contentEditable={!previewMergedMode && !isBodyLocked}
                suppressContentEditableWarning
                dir="rtl"
                onInput={onSyncPlaceholdersAndDraft}
                onKeyUp={(e) => {
                  if (e.key === ']' || e.key === '}') {
                    onDecorateEditorZonesPreservingSelection();
                  }
                }}
                onBlur={onDecorateEditorZonesPreservingSelection}
                onKeyDown={onEditorKeyDown}
                onPaste={onEditorPaste}
                onClick={onEditorClick}
                className="notary-editor-zone outline-none min-h-[18px]"
              />
            </div>
          )}

          {/* MAIN CONTRACT BODY EDITABLE ZONE OR PAGINATED MERGED PREVIEW */}
          {previewMergedMode ? (
            <div
              dir="rtl"
              className="notary-editor-zone relative z-10 min-h-[210mm]"
              dangerouslySetInnerHTML={{
                __html: mergedPreviewHtml,
              }}
            />
          ) : (
            <>
              <div
                ref={bodyEditorRef}
                contentEditable={!isBodyLocked}
                suppressContentEditableWarning
                dir="rtl"
                onInput={onSyncPlaceholdersAndDraft}
                onKeyUp={(e) => {
                  if (e.key === ']' || e.key === '}') {
                    onDecorateEditorZonesPreservingSelection();
                  }
                }}
                onBlur={onDecorateEditorZonesPreservingSelection}
                onKeyDown={onEditorKeyDown}
                onPaste={onEditorPaste}
                onMouseDown={onEditorMouseDown}
                onClick={onEditorClick}
                onDoubleClick={onEditorDoubleClick}
                className={`notary-editor-zone relative z-10 outline-none min-h-[210mm] print:hidden ${
                  isBodyLocked ? 'cursor-default select-text' : ''
                }`}
              />
              {/* Print-Only Paginated Layer with exact ../.. on non-final pages and الصفحة … والأخيرة on final page */}
              <div
                dir="rtl"
                className="hidden print:block notary-editor-zone relative z-10 min-h-[210mm]"
                dangerouslySetInnerHTML={{
                  __html: printPaginatedHtml,
                }}
              />
            </>
          )}

          {/* Clean Notarial Closing Line Marker at end of body text (no artificial explanation on screen) */}
          {pageNumberingEnabled && !previewMergedMode && (
            <div
              dir="rtl"
              className="text-left text-slate-900 font-bold select-none mt-1 relative z-10 no-print flex items-center justify-between gap-2"
              style={{
                fontFamily: STRICT_FONT_FAMILY,
                fontSize: `${STRICT_FONT_SIZE_PT}pt`,
                lineHeight: 1,
              }}
            >
              {estimatedPages > 1 && onTogglePreviewMergedMode ? (
                <button
                  type="button"
                  onClick={onTogglePreviewMergedMode}
                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-slate-100 hover:bg-blue-50 text-slate-600 hover:text-blue-900 border border-slate-200 font-sans font-semibold text-[10px] transition"
                  title="معاينة تقسيم الصفحات ومواضع علامة ../.. تحت آخر سطر لكل صفحة كما ستخرج في الطباعة والوورد"
                >
                  <Eye className="w-3 h-3" />
                  <span>معاينة مواضع (../..) عبر {estimatedPages} صفحات</span>
                </button>
              ) : (
                <span />
              )}
              <div>{buildNotarialPaginationClosingText(estimatedPages)}</div>
            </div>
          )}

          {/* Optional Document Footer */}
          {showHeaderFooter && (
            <div className="mt-4 pt-1 border-t border-dashed border-slate-300 relative z-10">
              <div className="text-[9px] text-slate-400 mb-0.5 select-none no-print">
                تذييل العقد (Footer)
              </div>
              <div
                ref={footerEditorRef}
                contentEditable={!previewMergedMode && !isBodyLocked}
                suppressContentEditableWarning
                dir="rtl"
                onInput={onSyncPlaceholdersAndDraft}
                onKeyUp={(e) => {
                  if (e.key === ']' || e.key === '}') {
                    onDecorateEditorZonesPreservingSelection();
                  }
                }}
                onBlur={onDecorateEditorZonesPreservingSelection}
                onKeyDown={onEditorKeyDown}
                onPaste={onEditorPaste}
                onClick={onEditorClick}
                className="notary-editor-zone outline-none min-h-[18px]"
              />
            </div>
          )}
        </div>
      </div>
    </main>
  );
}
