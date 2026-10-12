'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  DocumentRevision,
  HistorySnapshot,
  ListNumberingStyle,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  ToolbarState,
} from '@/lib/types';
import { loadDocumentRevisions, saveDocumentRevision } from '@/lib/storage';
import {
  applyModernInlineFormattingToRange,
  applyModernListToggleToRange,
  escapeHtml,
  getIntersectingBlockElements,
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
  restoreSerializedSelection,
  serializeCurrentSelection,
} from '@/lib/editor-utils';
import {
  createNotaryTableHTML,
  deleteTableColumnAtCell,
  deleteTableRowAtCell,
  insertTableColumnAtCell,
  insertTableRowAtCell,
  mergeSelectedTableCells,
  splitMergedTableCell,
} from '@/lib/table-grid';

interface UseEditorDomCommandsParams {
  docId: string;
  docTitle: string;
  authorName: string;
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  headerEditorRef: React.RefObject<HTMLDivElement | null>;
  footerEditorRef: React.RefObject<HTMLDivElement | null>;
  fieldValuesRef: React.RefObject<Record<string, string>>;
  ensureBodyEditable: () => boolean;
  getEditorZonesHtml: () => {
    bodyHtml: string;
    headerHtml: string;
    footerHtml: string;
  };
  syncPlaceholdersAndDraft: () => void;
  setRevisions: (revs: DocumentRevision[]) => void;
  setAutoSaveState: (state: 'saved' | 'saving' | 'error') => void;
  showToast: (msg: string) => void;
}

export function useEditorDomCommands({
  docId,
  docTitle,
  authorName,
  bodyEditorRef,
  headerEditorRef,
  footerEditorRef,
  fieldValuesRef,
  ensureBodyEditable,
  getEditorZonesHtml,
  syncPlaceholdersAndDraft,
  setRevisions,
  setAutoSaveState,
  showToast,
}: UseEditorDomCommandsParams) {
  const savedRangeRef = useRef<Range | null>(null);
  const activeTableCellRef = useRef<HTMLTableCellElement | null>(null);
  const selectedTableCellsRef = useRef<HTMLTableCellElement[]>([]);

  const [undoStack, setUndoStack] = useState<HistorySnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<HistorySnapshot[]>([]);
  const lastTypingSnapshotTimeRef = useRef<number>(0);

  const [toolbarState, setToolbarState] = useState<ToolbarState>({
    bold: false,
    italic: false,
    underline: false,
    align: 'justify',
    dir: 'rtl',
    textColor: '#000000',
    highlightColor: 'transparent',
    listType: 'none',
    inTable: false,
    canMergeCells: false,
    canSplitCell: false,
  });

  const saveSelectionBookmark = useCallback(() => {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    if (
      bodyEditorRef.current?.contains(range.commonAncestorContainer) ||
      headerEditorRef.current?.contains(range.commonAncestorContainer) ||
      footerEditorRef.current?.contains(range.commonAncestorContainer)
    ) {
      savedRangeRef.current = range.cloneRange();
    }
  }, [bodyEditorRef, footerEditorRef, headerEditorRef]);

  const restoreSelectionBookmarkIfNeeded = useCallback(() => {
    const sel = window.getSelection();
    if (!sel) return;
    if (
      sel.rangeCount > 0 &&
      (bodyEditorRef.current?.contains(
        sel.getRangeAt(0).commonAncestorContainer
      ) ||
        headerEditorRef.current?.contains(
          sel.getRangeAt(0).commonAncestorContainer
        ) ||
        footerEditorRef.current?.contains(
          sel.getRangeAt(0).commonAncestorContainer
        ))
    ) {
      return;
    }
    if (savedRangeRef.current) {
      try {
        sel.removeAllRanges();
        sel.addRange(savedRangeRef.current);
      } catch {
        // Ignore stale range
      }
    }
  }, [bodyEditorRef, footerEditorRef, headerEditorRef]);

  const syncToolbarWithSelection = useCallback(() => {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const range = sel.getRangeAt(0);
    const container = range.commonAncestorContainer;

    const isInsideEditor =
      bodyEditorRef.current?.contains(container) ||
      headerEditorRef.current?.contains(container) ||
      footerEditorRef.current?.contains(container);

    if (!isInsideEditor) return;

    savedRangeRef.current = range.cloneRange();

    let el: HTMLElement | null =
      container.nodeType === Node.ELEMENT_NODE
        ? (container as HTMLElement)
        : container.parentElement;

    let bold = false;
    let italic = false;
    let underline = false;
    let align: 'right' | 'center' | 'left' | 'justify' = 'justify';
    let dir: 'rtl' | 'ltr' = 'rtl';
    let textColor = '#000000';
    let highlightColor = 'transparent';
    let listType: 'none' | ListNumberingStyle = 'none';
    let foundCell: HTMLTableCellElement | null = null;

    while (
      el &&
      el !== bodyEditorRef.current &&
      el !== headerEditorRef.current &&
      el !== footerEditorRef.current
    ) {
      const tag = el.tagName.toUpperCase();
      const style = window.getComputedStyle(el);

      if (
        tag === 'B' ||
        tag === 'STRONG' ||
        style.fontWeight === 'bold' ||
        Number(style.fontWeight) >= 600
      ) {
        bold = true;
      }
      if (tag === 'I' || tag === 'EM' || style.fontStyle === 'italic') {
        italic = true;
      }
      if (tag === 'U' || style.textDecorationLine.includes('underline')) {
        underline = true;
      }
      if (el.style?.color) {
        textColor = el.style.color;
      }
      if (
        el.style?.backgroundColor &&
        el.style.backgroundColor !== 'transparent'
      ) {
        highlightColor = el.style.backgroundColor;
      }
      if (['P', 'LI', 'DIV', 'H1', 'H2', 'H3', 'H4'].includes(tag)) {
        const rawAlign = (
          el.style?.textAlign ||
          style.textAlign ||
          'justify'
        ).toLowerCase();
        if (rawAlign === 'center') align = 'center';
        else if (rawAlign === 'left') align = 'left';
        else if (rawAlign === 'right' || rawAlign === 'start') align = 'right';
        else align = 'justify';

        const rawDir = (
          el.getAttribute('dir') ||
          style.direction ||
          'rtl'
        ).toLowerCase();
        dir = rawDir === 'ltr' ? 'ltr' : 'rtl';
      }
      if (tag === 'OL' || tag === 'UL') {
        const lType = el.getAttribute(
          'data-list-type'
        ) as ListNumberingStyle | null;
        listType = lType || (tag === 'UL' ? 'bullet' : 'decimal');
      }
      if ((tag === 'TD' || tag === 'TH') && !foundCell) {
        foundCell = el as HTMLTableCellElement;
      }

      el = el.parentElement;
    }

    activeTableCellRef.current = foundCell;

    setToolbarState({
      bold,
      italic,
      underline,
      align,
      dir,
      textColor,
      highlightColor,
      listType,
      inTable: Boolean(foundCell),
      canMergeCells: Boolean(foundCell),
      canSplitCell: Boolean(
        foundCell &&
          ((foundCell.rowSpan || 1) > 1 || (foundCell.colSpan || 1) > 1)
      ),
    });
  }, [bodyEditorRef, footerEditorRef, headerEditorRef]);

  useEffect(() => {
    const handler = () => syncToolbarWithSelection();
    document.addEventListener('selectionchange', handler);
    return () => document.removeEventListener('selectionchange', handler);
  }, [syncToolbarWithSelection]);

  const recordHistorySnapshot = useCallback(
    (label?: string) => {
      lastTypingSnapshotTimeRef.current = 0;
      const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
      const selection = serializeCurrentSelection({
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      });
      setUndoStack((prev) => {
        if (prev.length > 0 && prev[prev.length - 1].bodyHtml === bodyHtml) {
          return prev;
        }
        return [
          ...prev.slice(-49),
          {
            bodyHtml,
            headerHtml,
            footerHtml,
            selection,
            timestamp: Date.now(),
          },
        ];
      });
      setRedoStack([]);

      if (label) {
        const rev: DocumentRevision = {
          id: `rev_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
          documentId: docId,
          documentTitle: docTitle || 'عقد توثيقي',
          author: authorName || 'كاتب المكتب',
          bodyHtml,
          fieldValues: { ...fieldValuesRef.current },
          createdAt: new Date().toISOString(),
          summary: label,
          eventType: 'save',
        };
        saveDocumentRevision(rev)
          .then(() => {
            loadDocumentRevisions().then((revs) => setRevisions(revs));
          })
          .catch((err) => {
            setAutoSaveState('error');
            showToast(
              err instanceof Error
                ? `تعذر حفظ مراجعة العقد: ${err.message}`
                : 'تعذر حفظ مراجعة العقد'
            );
          });
      }
    },
    [
      authorName,
      bodyEditorRef,
      docId,
      docTitle,
      fieldValuesRef,
      footerEditorRef,
      getEditorZonesHtml,
      headerEditorRef,
      setAutoSaveState,
      setRevisions,
      showToast,
    ]
  );

  /**
   * Coalesced Typing History Snapshot:
   * Records a pre-typing snapshot only when starting a new typing burst (> 900ms since last keystroke),
   * so Ctrl+Z undoes a logical sentence/phrase rather than every single letter.
   */
  const recordTypingHistorySnapshot = useCallback(() => {
    const now = Date.now();
    if (now - lastTypingSnapshotTimeRef.current < 900) {
      lastTypingSnapshotTimeRef.current = now;
      return;
    }
    lastTypingSnapshotTimeRef.current = now;
    const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
    const selection = serializeCurrentSelection({
      body: bodyEditorRef.current,
      header: headerEditorRef.current,
      footer: footerEditorRef.current,
    });
    setUndoStack((prev) => {
      if (prev.length > 0 && prev[prev.length - 1].bodyHtml === bodyHtml) {
        return prev;
      }
      return [
        ...prev.slice(-49),
        {
          bodyHtml,
          headerHtml,
          footerHtml,
          selection,
          timestamp: now,
        },
      ];
    });
    setRedoStack([]);
  }, [bodyEditorRef, footerEditorRef, getEditorZonesHtml, headerEditorRef]);

  const handleUndo = useCallback(() => {
    if (!ensureBodyEditable() || undoStack.length === 0) return;
    lastTypingSnapshotTimeRef.current = 0;
    const currentSnapshot: HistorySnapshot = {
      ...getEditorZonesHtml(),
      selection: serializeCurrentSelection({
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      }),
      timestamp: Date.now(),
    };
    const previous = undoStack[undoStack.length - 1];
    setUndoStack((prev) => prev.slice(0, -1));
    setRedoStack((prev) => [...prev, currentSnapshot]);

    if (bodyEditorRef.current) {
      bodyEditorRef.current.innerHTML = previous.bodyHtml;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
    }
    if (headerEditorRef.current)
      headerEditorRef.current.innerHTML = previous.headerHtml;
    if (footerEditorRef.current)
      footerEditorRef.current.innerHTML = previous.footerHtml;

    restoreSerializedSelection(
      {
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      },
      previous.selection
    );
    syncPlaceholdersAndDraft();
  }, [
    bodyEditorRef,
    ensureBodyEditable,
    footerEditorRef,
    getEditorZonesHtml,
    headerEditorRef,
    syncPlaceholdersAndDraft,
    undoStack,
  ]);

  const handleRedo = useCallback(() => {
    if (!ensureBodyEditable() || redoStack.length === 0) return;
    lastTypingSnapshotTimeRef.current = 0;
    const currentSnapshot: HistorySnapshot = {
      ...getEditorZonesHtml(),
      selection: serializeCurrentSelection({
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      }),
      timestamp: Date.now(),
    };
    const next = redoStack[redoStack.length - 1];
    setRedoStack((prev) => prev.slice(0, -1));
    setUndoStack((prev) => [...prev, currentSnapshot]);

    if (bodyEditorRef.current) {
      bodyEditorRef.current.innerHTML = next.bodyHtml;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
    }
    if (headerEditorRef.current)
      headerEditorRef.current.innerHTML = next.headerHtml;
    if (footerEditorRef.current)
      footerEditorRef.current.innerHTML = next.footerHtml;

    restoreSerializedSelection(
      {
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      },
      next.selection
    );
    syncPlaceholdersAndDraft();
  }, [
    bodyEditorRef,
    ensureBodyEditable,
    footerEditorRef,
    getEditorZonesHtml,
    headerEditorRef,
    redoStack,
    syncPlaceholdersAndDraft,
  ]);

  const applyInlineCommand = useCallback(
    (command: string, value?: string) => {
      if (!ensureBodyEditable()) return;
      restoreSelectionBookmarkIfNeeded();
      recordHistorySnapshot();

      const sel = window.getSelection();
      const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
      const activeZone =
        range && headerEditorRef.current?.contains(range.commonAncestorContainer)
          ? headerEditorRef.current
          : range &&
            footerEditorRef.current?.contains(range.commonAncestorContainer)
          ? footerEditorRef.current
          : bodyEditorRef.current;

      let handledByRangeEngine = false;
      if (
        activeZone &&
        (command === 'bold' ||
          command === 'italic' ||
          command === 'underline' ||
          command === 'foreColor' ||
          command === 'hiliteColor')
      ) {
        handledByRangeEngine = applyModernInlineFormattingToRange(
          activeZone,
          command,
          value
        );
      }

      if (!handledByRangeEngine) {
        document.execCommand(command, false, value);
      }

      if (bodyEditorRef.current)
        normalizeNotaryContainerDOM(bodyEditorRef.current);
      syncToolbarWithSelection();
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      footerEditorRef,
      headerEditorRef,
      recordHistorySnapshot,
      restoreSelectionBookmarkIfNeeded,
      syncPlaceholdersAndDraft,
      syncToolbarWithSelection,
    ]
  );

  const handleSetAlign = useCallback(
    (align: 'right' | 'center' | 'left' | 'justify') => {
      if (!ensureBodyEditable()) return;
      restoreSelectionBookmarkIfNeeded();
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0 || !bodyEditorRef.current) return;
      recordHistorySnapshot();

      const range = sel.getRangeAt(0);
      const activeZone = headerEditorRef.current?.contains(
        range.commonAncestorContainer
      )
        ? headerEditorRef.current
        : footerEditorRef.current?.contains(range.commonAncestorContainer)
        ? footerEditorRef.current
        : bodyEditorRef.current;

      const blocks = getIntersectingBlockElements(activeZone, range);
      for (const b of blocks) {
        b.style.textAlign = align;
      }
      syncToolbarWithSelection();
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      footerEditorRef,
      headerEditorRef,
      recordHistorySnapshot,
      restoreSelectionBookmarkIfNeeded,
      syncPlaceholdersAndDraft,
      syncToolbarWithSelection,
    ]
  );

  const handleSetDirection = useCallback(
    (dir: 'rtl' | 'ltr') => {
      if (!ensureBodyEditable()) return;
      restoreSelectionBookmarkIfNeeded();
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0 || !bodyEditorRef.current) return;
      recordHistorySnapshot();

      const range = sel.getRangeAt(0);
      const activeZone = headerEditorRef.current?.contains(
        range.commonAncestorContainer
      )
        ? headerEditorRef.current
        : footerEditorRef.current?.contains(range.commonAncestorContainer)
        ? footerEditorRef.current
        : bodyEditorRef.current;

      const blocks = getIntersectingBlockElements(activeZone, range);
      for (const b of blocks) {
        const oldDir = (b.getAttribute('dir') || 'rtl').toLowerCase();
        b.setAttribute('dir', dir);
        if (oldDir !== dir) {
          if (b.style.textAlign === 'right') b.style.textAlign = 'left';
          else if (b.style.textAlign === 'left') b.style.textAlign = 'right';
        }
      }
      syncToolbarWithSelection();
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      footerEditorRef,
      headerEditorRef,
      recordHistorySnapshot,
      restoreSelectionBookmarkIfNeeded,
      syncPlaceholdersAndDraft,
      syncToolbarWithSelection,
    ]
  );

  const handleAdjustIndent = useCallback(
    (delta: 'increase' | 'decrease') => {
      if (!ensureBodyEditable()) return;
      restoreSelectionBookmarkIfNeeded();
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0 || !bodyEditorRef.current) return;
      recordHistorySnapshot();

      const range = sel.getRangeAt(0);
      const blocks = getIntersectingBlockElements(bodyEditorRef.current, range);
      for (const b of blocks) {
        const isRtl = (b.getAttribute('dir') || 'rtl').toLowerCase() !== 'ltr';
        const prop = isRtl ? 'paddingRight' : 'paddingLeft';
        const currentPx = parseInt(b.style[prop] || '0', 10) || 0;
        const nextPx =
          delta === 'increase' ? currentPx + 24 : Math.max(0, currentPx - 24);
        b.style[prop] = nextPx > 0 ? `${nextPx}px` : '';
      }
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      restoreSelectionBookmarkIfNeeded,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleToggleList = useCallback(
    (style: ListNumberingStyle) => {
      if (!ensureBodyEditable()) return;
      restoreSelectionBookmarkIfNeeded();
      if (!bodyEditorRef.current) return;
      recordHistorySnapshot();

      const handledByRangeEngine = applyModernListToggleToRange(
        bodyEditorRef.current,
        style
      );

      if (!handledByRangeEngine) {
        const isOrdered =
          style === 'decimal' ||
          style === 'arabic-alpha' ||
          style === 'arabic-abjad';
        document.execCommand(
          isOrdered ? 'insertOrderedList' : 'insertUnorderedList',
          false
        );

        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0) {
          let node: Node | null = sel.getRangeAt(0).commonAncestorContainer;
          while (node && node !== bodyEditorRef.current) {
            if (
              node.nodeType === Node.ELEMENT_NODE &&
              ['OL', 'UL'].includes((node as HTMLElement).tagName)
            ) {
              const listEl = node as HTMLElement;
              listEl.setAttribute('data-list-type', style);
              listEl.setAttribute('dir', 'rtl');
              break;
            }
            node = node.parentNode;
          }
        }
      }

      normalizeNotaryContainerDOM(bodyEditorRef.current);
      syncToolbarWithSelection();
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      restoreSelectionBookmarkIfNeeded,
      syncPlaceholdersAndDraft,
      syncToolbarWithSelection,
    ]
  );

  const handleInsertTable = useCallback(
    (rows: number, cols: number) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      const html = createNotaryTableHTML(rows, cols);
      insertHtmlAtSelection(
        bodyEditorRef.current,
        html,
        savedRangeRef.current
      );
      syncToolbarWithSelection();
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      syncPlaceholdersAndDraft,
      syncToolbarWithSelection,
    ]
  );

  const handleTableAction = useCallback(
    (
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
    ) => {
      if (!ensureBodyEditable()) return;
      const cell = activeTableCellRef.current;
      if (!cell || !bodyEditorRef.current) return;
      recordHistorySnapshot();

      if (action === 'row-above') insertTableRowAtCell(cell, 'above');
      else if (action === 'row-below') insertTableRowAtCell(cell, 'below');
      else if (action === 'delete-row') deleteTableRowAtCell(cell);
      else if (action === 'col-before') insertTableColumnAtCell(cell, 'before');
      else if (action === 'col-after') insertTableColumnAtCell(cell, 'after');
      else if (action === 'delete-col') deleteTableColumnAtCell(cell);
      else if (action === 'merge-cells') {
        const cellsToMerge =
          selectedTableCellsRef.current.length > 1
            ? selectedTableCellsRef.current
            : [cell];
        mergeSelectedTableCells(cellsToMerge);
        selectedTableCellsRef.current.forEach((c) =>
          c.classList.remove('selected-table-cell')
        );
        selectedTableCellsRef.current = [];
      } else if (action === 'split-cell') {
        splitMergedTableCell(cell);
      } else if (action === 'toggle-border') {
        const isHidden = cell.getAttribute('data-border') === 'none';
        if (isHidden) {
          cell.removeAttribute('data-border');
          cell.style.border = '1px solid #000000';
        } else {
          cell.setAttribute('data-border', 'none');
          cell.style.border = 'none';
        }
      } else if (action === 'shade-cell') {
        const hasShade =
          cell.style.backgroundColor &&
          cell.style.backgroundColor !== 'transparent' &&
          cell.style.backgroundColor !== 'rgb(255, 255, 255)';
        cell.style.backgroundColor = hasShade ? '' : '#F1F5F9';
      }

      normalizeNotaryContainerDOM(bodyEditorRef.current);
      syncToolbarWithSelection();
      syncPlaceholdersAndDraft();
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      syncPlaceholdersAndDraft,
      syncToolbarWithSelection,
    ]
  );

  const handleInsertPageBreak = useCallback(() => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();
    const breakHtml = `<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div><p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><br></p>`;
    insertHtmlAtSelection(
      bodyEditorRef.current,
      breakHtml,
      savedRangeRef.current
    );
    syncPlaceholdersAndDraft();
  }, [
    bodyEditorRef,
    ensureBodyEditable,
    recordHistorySnapshot,
    syncPlaceholdersAndDraft,
  ]);

  const handleInsertImageFile = useCallback(
    (file: File) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result === 'string' && bodyEditorRef.current) {
          recordHistorySnapshot();
          const imgHtml = `<img src="${reader.result}" alt="${escapeHtml(
            file.name
          )}" style="max-width:100%;height:auto;" />`;
          insertHtmlAtSelection(
            bodyEditorRef.current,
            imgHtml,
            savedRangeRef.current
          );
          syncPlaceholdersAndDraft();
        }
      };
      reader.readAsDataURL(file);
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      syncPlaceholdersAndDraft,
    ]
  );

  return {
    savedRangeRef,
    activeTableCellRef,
    selectedTableCellsRef,
    undoStack,
    setUndoStack,
    redoStack,
    setRedoStack,
    toolbarState,
    saveSelectionBookmark,
    restoreSelectionBookmarkIfNeeded,
    syncToolbarWithSelection,
    recordHistorySnapshot,
    recordTypingHistorySnapshot,
    handleUndo,
    handleRedo,
    applyInlineCommand,
    handleSetAlign,
    handleSetDirection,
    handleAdjustIndent,
    handleToggleList,
    handleInsertTable,
    handleTableAction,
    handleInsertPageBreak,
    handleInsertImageFile,
  };
}
