'use client';

import React, { useCallback, useMemo, useRef, useState } from 'react';
import { QuickAssistOption } from '@/components/editor/EditorQuickAssistPopover';
import { InlineVariablePopoverState } from '@/components/InlineVariablePopover';
import { SidebarTab } from '@/components/SidebarWorkspace';
import { EditorSmartAssistPreferences } from '@/lib/types';
import {
  CaretInlineTrigger,
  detectCaretInlineTrigger,
  findAdjacentSmartTagElement,
  insertHtmlAtSelection,
  moveCaretOutsideSmartTag,
  sanitizePastedWordHTML,
} from '@/lib/editor-utils';
import { insertTableRowAtCell } from '@/lib/table-grid';

interface UseEditorSmartAssistParams {
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  savedRangeRef: React.MutableRefObject<Range | null>;
  activeTableCellRef: React.MutableRefObject<HTMLTableCellElement | null>;
  selectedTableCellsRef: React.MutableRefObject<HTMLTableCellElement[]>;
  debouncedSyncTimerRef: React.MutableRefObject<
    ReturnType<typeof setTimeout> | null
  >;
  isBodyLocked: boolean;
  extractedPlaceholders: string[];
  fieldValues: Record<string, string>;
  ensureBodyEditable: () => boolean;
  recordHistorySnapshot: (label?: string) => void;
  recordTypingHistorySnapshot: () => void;
  applyInlineCommand: (
    command: 'bold' | 'italic' | 'underline' | 'foreColor' | 'hiliteColor',
    value?: string
  ) => void;
  handleUndo: () => void;
  handleRedo: () => void;
  handleInsertTable: (rows: number, cols: number) => void;
  handleInsertPageBreak: () => void;
  handleConvertSelectionToSmartTag: () => void;
  syncPlaceholdersAndDraft: () => void;
  flushSyncPlaceholdersAndDraft: () => void;
  triggerFastClauseScan?: () => void;
  setShowFindReplace: React.Dispatch<React.SetStateAction<boolean>>;
  setInlinePopover: React.Dispatch<
    React.SetStateAction<InlineVariablePopoverState | null>
  >;
  setSidebarTab: React.Dispatch<React.SetStateAction<SidebarTab>>;
  setIsSidebarOpen: React.Dispatch<React.SetStateAction<boolean>>;
}

export function useEditorSmartAssist({
  bodyEditorRef,
  savedRangeRef,
  activeTableCellRef,
  selectedTableCellsRef,
  debouncedSyncTimerRef,
  isBodyLocked,
  extractedPlaceholders,
  fieldValues,
  ensureBodyEditable,
  recordHistorySnapshot,
  recordTypingHistorySnapshot,
  applyInlineCommand,
  handleUndo,
  handleRedo,
  handleInsertTable,
  handleInsertPageBreak,
  handleConvertSelectionToSmartTag,
  syncPlaceholdersAndDraft,
  flushSyncPlaceholdersAndDraft,
  triggerFastClauseScan,
  setShowFindReplace,
  setInlinePopover,
  setSidebarTab,
  setIsSidebarOpen,
}: UseEditorSmartAssistParams) {
  const selectedAtomicTagRef = useRef<HTMLElement | null>(null);

  const [smartAssistPrefs, setSmartAssistPrefs] =
    useState<EditorSmartAssistPreferences>({
      enableInlineAutocomplete: true,
      enableSlashCommands: true,
    });
  const [quickAssistTrigger, setQuickAssistTrigger] =
    useState<CaretInlineTrigger | null>(null);
  const [quickAssistSelectedIdx, setQuickAssistSelectedIdx] =
    useState<number>(0);

  const clearAtomicSelectedTag = useCallback(() => {
    if (selectedAtomicTagRef.current) {
      selectedAtomicTagRef.current.classList.remove(
        'smart-tag-selected-atomic'
      );
      selectedAtomicTagRef.current = null;
    }
    if (bodyEditorRef.current) {
      bodyEditorRef.current
        .querySelectorAll('.smart-tag-selected-atomic')
        .forEach((el) => el.classList.remove('smart-tag-selected-atomic'));
    }
  }, [bodyEditorRef]);

  const handleUpdateSmartAssistPrefs = useCallback(
    (next: EditorSmartAssistPreferences) => {
      setSmartAssistPrefs(next);
      if (typeof window !== 'undefined') {
        localStorage.setItem(
          'notary_smart_assist_prefs_v1',
          JSON.stringify(next)
        );
      }
      if (!next.enableInlineAutocomplete && !next.enableSlashCommands) {
        setQuickAssistTrigger(null);
      }
    },
    []
  );

  const refreshInlineQuickAssistTrigger = useCallback(() => {
    if (!bodyEditorRef.current || isBodyLocked) {
      setQuickAssistTrigger(null);
      return;
    }
    const detected = detectCaretInlineTrigger(bodyEditorRef.current);
    if (!detected) {
      setQuickAssistTrigger(null);
      return;
    }
    if (
      (detected.type === 'tag' || detected.type === 'clause') &&
      !smartAssistPrefs.enableInlineAutocomplete
    ) {
      setQuickAssistTrigger(null);
      return;
    }
    if (detected.type === 'slash' && !smartAssistPrefs.enableSlashCommands) {
      setQuickAssistTrigger(null);
      return;
    }
    setQuickAssistTrigger(detected);
    setQuickAssistSelectedIdx(0);
  }, [bodyEditorRef, isBodyLocked, smartAssistPrefs]);

  const scheduleSyncPlaceholdersAndDraft = useCallback(() => {
    clearAtomicSelectedTag();
    refreshInlineQuickAssistTrigger();
    triggerFastClauseScan?.();
    if (debouncedSyncTimerRef.current) {
      clearTimeout(debouncedSyncTimerRef.current);
    }
    debouncedSyncTimerRef.current = setTimeout(() => {
      debouncedSyncTimerRef.current = null;
      syncPlaceholdersAndDraft();
    }, 160);
  }, [
    clearAtomicSelectedTag,
    debouncedSyncTimerRef,
    refreshInlineQuickAssistTrigger,
    syncPlaceholdersAndDraft,
    triggerFastClauseScan,
  ]);

  const quickAssistOptions = useMemo<QuickAssistOption[]>(() => {
    if (!quickAssistTrigger) return [];
    const q = quickAssistTrigger.query.trim();

    if (quickAssistTrigger.type === 'tag') {
      const commonTags = [
        { key: 'الطرف_الأول_الاسم', label: 'اسم ولقب الطرف الأول' },
        { key: 'الطرف_الأول_تاريخ_الميلاد', label: 'تاريخ ميلاد الطرف الأول' },
        { key: 'الطرف_الأول_مكان_الميلاد', label: 'مكان ميلاد الطرف الأول' },
        { key: 'الطرف_الأول_النسب', label: 'نسب الطرف الأول' },
        { key: 'الطرف_الأول_الهوية', label: 'مرجع هوية الطرف الأول' },
        { key: 'الطرف_الأول_الإقامة', label: 'إقامة الطرف الأول' },
        { key: 'الطرف_الثاني_الاسم', label: 'اسم ولقب الطرف الثاني' },
        { key: 'الطرف_الثاني_تاريخ_الميلاد', label: 'تاريخ ميلاد الطرف الثاني' },
        { key: 'الطرف_الثاني_مكان_الميلاد', label: 'مكان ميلاد الطرف الثاني' },
        { key: 'الطرف_الثاني_النسب', label: 'نسب الطرف الثاني' },
        { key: 'الطرف_الثاني_الهوية', label: 'مرجع هوية الطرف الثاني' },
        { key: 'الطرف_الثاني_الإقامة', label: 'إقامة الطرف الثاني' },
        { key: 'رقم_الفهرس', label: 'رقم الفهرس الرسمي' },
        { key: 'تاريخ_العقد', label: 'تاريخ تحرير العقد' },
        { key: 'تعيين_الحصة_الكامل', label: 'التعيين الكامل للعقار/الحصة' },
        { key: 'المساحة', label: 'المساحة (م²)' },
        { key: 'الثمن_بالأحرف', label: 'الثمن الإجمالي بالأحرف' },
        { key: 'الثمن_بالأرقام', label: 'الثمن الإجمالي بالأرقام' },
        { key: 'مراجع_التسجيل', label: 'مراجع التسجيل والطابع' },
        { key: 'المحافظة_العقارية', label: 'المحافظة العقارية' },
      ];
      const allKeys = Array.from(
        new Set([...extractedPlaceholders, ...commonTags.map((t) => t.key)])
      );
      const list: QuickAssistOption[] = [];
      if (q && !allKeys.includes(q.replace(/\s+/g, '_'))) {
        const customKey = q.replace(/\s+/g, '_');
        list.push({
          id: `custom_tag_${customKey}`,
          kind: 'tag',
          label: `إنشاء وسم جديد: ${q}`,
          insertValue: customKey,
          isFilled: Boolean(fieldValues[customKey]?.trim()),
        });
      }
      for (const k of allKeys) {
        const meta = commonTags.find((t) => t.key === k);
        const label = meta?.label || k.replace(/_/g, ' ');
        if (!q || k.includes(q) || label.includes(q)) {
          list.push({
            id: `tag_${k}`,
            kind: 'tag',
            label,
            insertValue: k,
            isFilled: Boolean(fieldValues[k]?.trim()),
          });
        }
      }
      return list.slice(0, 10);
    }

    if (quickAssistTrigger.type === 'clause') {
      const standardHeadings = [
        { title: 'حضر', subtitle: 'ديباجة حضور الأطراف أمام الموثق' },
        { title: 'التعيين', subtitle: 'وصف وتعيين العقار أو محل العقد' },
        { title: 'أصل الملكية', subtitle: 'تسلسل سندات الملكية السابقة' },
        {
          title: 'الملكية والانتفاع',
          subtitle: 'انتقال الملكية والحيازة القانونية',
        },
        { title: 'التكاليف والشروط', subtitle: 'الالتزامات العامة والخاصة' },
        { title: 'الثمن', subtitle: 'مبلغ المعاملة وطريقة الدفع والقبض' },
        {
          title: 'التسجيل والشهر العقاري',
          subtitle: 'الإجراءات الجبائية والعقارية',
        },
        {
          title: 'الحالة المدنية والأهلية',
          subtitle: 'تصريحات الأطراف والأهلية',
        },
        {
          title: 'الموطن المختار',
          subtitle: 'تحديد الموطن القانوني لتنفيذ العقد',
        },
        { title: 'إثباتاً لما ذكر', subtitle: 'خاتمة العقد والتلاوة والتوقيع' },
      ];
      const list: QuickAssistOption[] = [];
      if (q && !standardHeadings.some((h) => h.title === q)) {
        list.push({
          id: `custom_clause_${q}`,
          kind: 'clause',
          label: `عنوان بند حر: # ${q}`,
          subtitle: 'إدراج عنوان البند كما كتبته',
          insertValue: `# ${q}`,
        });
      }
      for (const h of standardHeadings) {
        if (!q || h.title.includes(q) || h.subtitle.includes(q)) {
          list.push({
            id: `clause_${h.title}`,
            kind: 'clause',
            label: h.title,
            subtitle: h.subtitle,
            insertValue: `# ${h.title}`,
          });
        }
      }
      return list;
    }

    const slashItems: QuickAssistOption[] = [
      {
        id: 'slash_clause',
        kind: 'slash',
        slashAction: 'clause',
        label: 'إدراج عنوان بند جديد (#)',
        subtitle: 'يضيف سطر عنوان بند مرن (# عنوان) في الموضع الحالي',
        insertValue: '# ',
      },
      {
        id: 'slash_tag',
        kind: 'slash',
        slashAction: 'tag',
        label: 'إدراج وسم متغير ذكي [...]',
        subtitle: 'يفتح قائمة المتغيرات لإدراج وسم ذكي مرتبط بالعقد',
        insertValue: '[',
      },
      {
        id: 'slash_table',
        kind: 'slash',
        slashAction: 'table',
        label: 'إدراج جدول توثيقي (3 × 4)',
        subtitle: 'جدول رسمي بعرض المساحة النصية الصافية (12 سم)',
        insertValue: '',
      },
      {
        id: 'slash_party',
        kind: 'slash',
        slashAction: 'party',
        label: 'فتح بطاقات الأطراف والمتغيرات',
        subtitle: 'استدعاء بيانات طرف من الدليل أو تعبئة الحقول',
        insertValue: '',
      },
      {
        id: 'slash_page_break',
        kind: 'slash',
        slashAction: 'page-break',
        label: 'إدراج فاصل صفحات',
        subtitle: 'بدء صفحة جديدة عند الطباعة وتصدير الوورد',
        insertValue: '',
      },
    ];
    return slashItems.filter(
      (item) =>
        !q ||
        item.label.includes(q) ||
        (item.subtitle && item.subtitle.includes(q))
    );
  }, [extractedPlaceholders, fieldValues, quickAssistTrigger]);

  const handleSelectQuickAssistOption = useCallback(
    (opt: QuickAssistOption) => {
      if (
        !quickAssistTrigger ||
        !bodyEditorRef.current ||
        !ensureBodyEditable()
      ) {
        setQuickAssistTrigger(null);
        return;
      }

      const { textNode, triggerOffset, caretOffset } = quickAssistTrigger;
      recordHistorySnapshot();

      try {
        const range = document.createRange();
        range.setStart(textNode, triggerOffset);
        range.setEnd(textNode, caretOffset);
        range.deleteContents();

        const sel = window.getSelection();
        if (sel) {
          sel.removeAllRanges();
          sel.addRange(range);
        }

        if (opt.kind === 'tag') {
          const span = document.createElement('span');
          span.className = 'smart-tag';
          span.setAttribute('data-var', opt.insertValue);
          span.textContent = `[${opt.insertValue}]`;
          const space = document.createTextNode('\u00A0');
          range.insertNode(space);
          range.insertNode(span);
          const afterRange = document.createRange();
          afterRange.setStartAfter(space);
          afterRange.collapse(true);
          sel?.removeAllRanges();
          sel?.addRange(afterRange);
          setQuickAssistTrigger(null);
          flushSyncPlaceholdersAndDraft();
          return;
        }

        if (opt.kind === 'clause') {
          const block = textNode.parentElement?.closest(
            'p, div, li'
          ) as HTMLElement | null;
          const boldSpan = document.createElement('span');
          boldSpan.style.fontWeight = 'bold';
          boldSpan.textContent = opt.insertValue;
          range.insertNode(boldSpan);
          if (block) {
            block.style.fontWeight = 'bold';
          }
          const afterRange = document.createRange();
          afterRange.setStartAfter(boldSpan);
          afterRange.collapse(true);
          sel?.removeAllRanges();
          sel?.addRange(afterRange);
          setQuickAssistTrigger(null);
          flushSyncPlaceholdersAndDraft();
          return;
        }

        if (opt.kind === 'slash') {
          setQuickAssistTrigger(null);
          if (opt.slashAction === 'clause') {
            const hashNode = document.createTextNode('# ');
            range.insertNode(hashNode);
            const afterRange = document.createRange();
            afterRange.setStartAfter(hashNode);
            afterRange.collapse(true);
            sel?.removeAllRanges();
            sel?.addRange(afterRange);
            setTimeout(() => refreshInlineQuickAssistTrigger(), 10);
          } else if (opt.slashAction === 'tag') {
            const bracketNode = document.createTextNode('[');
            range.insertNode(bracketNode);
            const afterRange = document.createRange();
            afterRange.setStartAfter(bracketNode);
            afterRange.collapse(true);
            sel?.removeAllRanges();
            sel?.addRange(afterRange);
            setTimeout(() => refreshInlineQuickAssistTrigger(), 10);
          } else if (opt.slashAction === 'table') {
            savedRangeRef.current = range.cloneRange();
            handleInsertTable(3, 4);
          } else if (opt.slashAction === 'party') {
            setSidebarTab('parties');
            setIsSidebarOpen(true);
          } else if (opt.slashAction === 'page-break') {
            savedRangeRef.current = range.cloneRange();
            handleInsertPageBreak();
          }
          flushSyncPlaceholdersAndDraft();
        }
      } catch {
        setQuickAssistTrigger(null);
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      flushSyncPlaceholdersAndDraft,
      handleInsertPageBreak,
      handleInsertTable,
      quickAssistTrigger,
      recordHistorySnapshot,
      refreshInlineQuickAssistTrigger,
      savedRangeRef,
      setIsSidebarOpen,
      setSidebarTab,
    ]
  );

  const handleEditorBeforeInput = useCallback(
    (e: React.FormEvent<HTMLDivElement>) => {
      if (isBodyLocked) {
        e.preventDefault();
        return;
      }
      const nativeEvent = e.nativeEvent as InputEvent;
      const inputType = nativeEvent.inputType || '';

      if (inputType === 'formatBold') {
        e.preventDefault();
        applyInlineCommand('bold');
        return;
      }
      if (inputType === 'formatItalic') {
        e.preventDefault();
        applyInlineCommand('italic');
        return;
      }
      if (inputType === 'formatUnderline') {
        e.preventDefault();
        applyInlineCommand('underline');
        return;
      }

      if (
        inputType.startsWith('insertText') ||
        inputType.startsWith('deleteContent') ||
        inputType === 'insertParagraph'
      ) {
        recordTypingHistorySnapshot();
      }
    },
    [applyInlineCommand, isBodyLocked, recordTypingHistorySnapshot]
  );

  const handleEditorKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      if (quickAssistTrigger && quickAssistOptions.length > 0) {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          setQuickAssistSelectedIdx(
            (prev) => (prev + 1) % quickAssistOptions.length
          );
          return;
        }
        if (e.key === 'ArrowUp') {
          e.preventDefault();
          setQuickAssistSelectedIdx(
            (prev) =>
              (prev - 1 + quickAssistOptions.length) % quickAssistOptions.length
          );
          return;
        }
        if (e.key === 'Enter' || e.key === 'Tab') {
          e.preventDefault();
          const chosen =
            quickAssistOptions[quickAssistSelectedIdx] || quickAssistOptions[0];
          if (chosen) {
            handleSelectQuickAssistOption(chosen);
          }
          return;
        }
        if (e.key === 'Escape') {
          e.preventDefault();
          setQuickAssistTrigger(null);
          return;
        }
      }

      if (
        (e.ctrlKey || e.metaKey) &&
        e.key.toLowerCase() === 'z' &&
        !e.shiftKey
      ) {
        e.preventDefault();
        clearAtomicSelectedTag();
        handleUndo();
        return;
      }
      if (
        ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') ||
        ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'z')
      ) {
        e.preventDefault();
        clearAtomicSelectedTag();
        handleRedo();
        return;
      }
      if (
        (e.ctrlKey || e.metaKey) &&
        !e.shiftKey &&
        e.key.toLowerCase() === 'b'
      ) {
        e.preventDefault();
        applyInlineCommand('bold');
        return;
      }
      if (
        (e.ctrlKey || e.metaKey) &&
        !e.shiftKey &&
        e.key.toLowerCase() === 'i'
      ) {
        e.preventDefault();
        applyInlineCommand('italic');
        return;
      }
      if (
        (e.ctrlKey || e.metaKey) &&
        !e.shiftKey &&
        e.key.toLowerCase() === 'u'
      ) {
        e.preventDefault();
        applyInlineCommand('underline');
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        setShowFindReplace(true);
        return;
      }
      if (
        (e.altKey && e.key.toLowerCase() === 'v') ||
        ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'x')
      ) {
        e.preventDefault();
        handleConvertSelectionToSmartTag();
        return;
      }

      if (e.key === 'Backspace' || e.key === 'Delete') {
        if (isBodyLocked) return;
        if (
          selectedAtomicTagRef.current &&
          document.body.contains(selectedAtomicTagRef.current)
        ) {
          e.preventDefault();
          recordHistorySnapshot();
          const tagToRemove = selectedAtomicTagRef.current;
          selectedAtomicTagRef.current = null;
          const parent = tagToRemove.parentNode;
          const nextSibling = tagToRemove.nextSibling;
          tagToRemove.remove();
          if (parent) {
            const sel = window.getSelection();
            const newRange = document.createRange();
            if (nextSibling) {
              newRange.setStartBefore(nextSibling);
            } else {
              newRange.selectNodeContents(parent);
              newRange.collapse(false);
            }
            sel?.removeAllRanges();
            sel?.addRange(newRange);
          }
          flushSyncPlaceholdersAndDraft();
          return;
        }

        const direction = e.key === 'Backspace' ? 'backward' : 'forward';
        const adjacentTag = findAdjacentSmartTagElement(direction);
        if (adjacentTag) {
          e.preventDefault();
          clearAtomicSelectedTag();
          adjacentTag.classList.add('smart-tag-selected-atomic');
          selectedAtomicTagRef.current = adjacentTag;
          moveCaretOutsideSmartTag(
            adjacentTag,
            direction === 'backward' ? 'after' : 'before'
          );
          return;
        }
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
        clearAtomicSelectedTag();
        if (!e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) {
          const direction = e.key === 'ArrowLeft' ? 'forward' : 'backward';
          const adjacentTag = findAdjacentSmartTagElement(direction);
          if (adjacentTag) {
            e.preventDefault();
            moveCaretOutsideSmartTag(
              adjacentTag,
              direction === 'forward' ? 'after' : 'before'
            );
            return;
          }
        }
      } else {
        clearAtomicSelectedTag();
      }

      if (e.key === 'Tab') {
        if (isBodyLocked) return;
        e.preventDefault();
        const cell = activeTableCellRef.current;
        if (cell) {
          const table = cell.closest('table') as HTMLTableElement | null;
          if (table) {
            const allCells = Array.from(
              table.querySelectorAll('td, th')
            ) as HTMLTableCellElement[];
            const idx = allCells.indexOf(cell);
            if (!e.shiftKey) {
              if (idx >= 0 && idx < allCells.length - 1) {
                const nextCell = allCells[idx + 1];
                const range = document.createRange();
                range.selectNodeContents(nextCell);
                range.collapse(true);
                const sel = window.getSelection();
                sel?.removeAllRanges();
                sel?.addRange(range);
              } else if (idx === allCells.length - 1) {
                recordHistorySnapshot();
                const newFirstCell = insertTableRowAtCell(cell, 'below');
                if (newFirstCell) {
                  const range = document.createRange();
                  range.selectNodeContents(newFirstCell);
                  range.collapse(true);
                  const sel = window.getSelection();
                  sel?.removeAllRanges();
                  sel?.addRange(range);
                }
              }
            } else if (idx > 0) {
              const prevCell = allCells[idx - 1];
              const range = document.createRange();
              range.selectNodeContents(prevCell);
              range.collapse(true);
              const sel = window.getSelection();
              sel?.removeAllRanges();
              sel?.addRange(range);
            }
            return;
          }
        }

        recordHistorySnapshot();
        if (bodyEditorRef.current) {
          insertHtmlAtSelection(
            bodyEditorRef.current,
            '&nbsp;&nbsp;&nbsp;&nbsp;',
            savedRangeRef.current
          );
        }
        flushSyncPlaceholdersAndDraft();
      }
    },
    [
      activeTableCellRef,
      applyInlineCommand,
      bodyEditorRef,
      clearAtomicSelectedTag,
      flushSyncPlaceholdersAndDraft,
      handleConvertSelectionToSmartTag,
      handleRedo,
      handleSelectQuickAssistOption,
      handleUndo,
      isBodyLocked,
      quickAssistOptions,
      quickAssistSelectedIdx,
      quickAssistTrigger,
      recordHistorySnapshot,
      savedRangeRef,
      setShowFindReplace,
    ]
  );

  const handleEditorPaste = useCallback(
    (e: React.ClipboardEvent<HTMLDivElement>) => {
      e.preventDefault();
      if (!ensureBodyEditable()) return;
      recordHistorySnapshot();
      const rawHtml = e.clipboardData.getData('text/html');
      const plainText = e.clipboardData.getData('text/plain');
      const cleanHtml = sanitizePastedWordHTML(rawHtml, plainText);
      if (bodyEditorRef.current) {
        const sel = window.getSelection();
        const range =
          sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : savedRangeRef.current;
        insertHtmlAtSelection(bodyEditorRef.current, cleanHtml, range);
        syncPlaceholdersAndDraft();
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      savedRangeRef,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleEditorMouseDown = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      clearAtomicSelectedTag();
      setQuickAssistTrigger(null);
      const target = e.target as HTMLElement;
      const clickedCell = target.closest(
        'td, th'
      ) as HTMLTableCellElement | null;
      if (clickedCell && e.shiftKey && activeTableCellRef.current) {
        e.preventDefault();
        const firstCell = activeTableCellRef.current;
        if (firstCell.closest('table') === clickedCell.closest('table')) {
          const list = selectedTableCellsRef.current.includes(firstCell)
            ? [...selectedTableCellsRef.current]
            : [firstCell];
          if (!list.includes(clickedCell)) list.push(clickedCell);
          selectedTableCellsRef.current = list;
          list.forEach((c) => c.classList.add('selected-table-cell'));
        }
      } else if (selectedTableCellsRef.current.length > 0) {
        selectedTableCellsRef.current.forEach((c) =>
          c.classList.remove('selected-table-cell')
        );
        selectedTableCellsRef.current = [];
      }
    },
    [activeTableCellRef, clearAtomicSelectedTag, selectedTableCellsRef]
  );

  const handleEditorClick = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      const target = e.target as HTMLElement;
      const smartTag = target.closest(
        '.smart-tag, .smart-placeholder'
      ) as HTMLElement | null;
      if (smartTag) {
        e.stopPropagation();
        const varKey =
          smartTag.getAttribute('data-var') ||
          (smartTag.textContent || '').replace(/[\[\]{}]/g, '').trim();
        if (varKey) {
          const domRect = smartTag.getBoundingClientRect();
          setInlinePopover({
            varKey,
            rect: {
              top: domRect.top,
              left: domRect.left,
              bottom: domRect.bottom,
              width: domRect.width,
            },
          });
        }
      }
    },
    [setInlinePopover]
  );

  const handleEditorDoubleClick = useCallback(
    (e: React.MouseEvent<HTMLDivElement>) => {
      const target = e.target as HTMLElement;
      const smartTag = target.closest(
        '.smart-tag, .smart-placeholder'
      ) as HTMLElement | null;
      if (smartTag) {
        e.preventDefault();
        setInlinePopover(null);
        setSidebarTab('parties');
        setIsSidebarOpen(true);
      }
    },
    [setInlinePopover, setIsSidebarOpen, setSidebarTab]
  );

  return {
    smartAssistPrefs,
    setSmartAssistPrefs,
    quickAssistTrigger,
    setQuickAssistTrigger,
    quickAssistSelectedIdx,
    quickAssistOptions,
    clearAtomicSelectedTag,
    handleUpdateSmartAssistPrefs,
    refreshInlineQuickAssistTrigger,
    scheduleSyncPlaceholdersAndDraft,
    handleSelectQuickAssistOption,
    handleEditorBeforeInput,
    handleEditorKeyDown,
    handleEditorPaste,
    handleEditorMouseDown,
    handleEditorClick,
    handleEditorDoubleClick,
  };
}
