'use client';

import React, { useCallback } from 'react';
import {
  ClauseCondition,
  ContractOutlineClause,
  ContractPartyCard,
  NotaryClause,
  OfficeTemplate,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  VariableInputType,
} from '@/lib/types';
import {
  loadNotaryClauses,
  loadOfficeTemplates,
  saveNotaryClause,
  saveOfficeTemplate,
} from '@/lib/storage';
import {
  deleteContractClauseFromDOM,
  extractLiveContractClausesFromDOM,
  insertOrWrapNewClauseAtSelection,
  moveContractClauseInDOM,
  renameContractClauseInDOM,
  reorderLiveContractClausesInDOM,
  scrollToContractClauseInDOM,
  setContractClauseLockedInDOM,
  toggleContractClauseEnabledInDOM,
  upgradeTopLevelHashSegmentsToContainers,
} from '@/lib/clause-outline-engine';
import {
  escapeHtml,
  extractPlaceholdersFromHtml,
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
} from '@/lib/editor-utils';
import { INITIAL_EMPTY_PARAGRAPH } from './useDocumentSession';

interface UseEditorClausesParams {
  docTitle: string;
  setDocTitle: React.Dispatch<React.SetStateAction<string>>;
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  headerEditorRef: React.RefObject<HTMLDivElement | null>;
  footerEditorRef: React.RefObject<HTMLDivElement | null>;
  savedRangeRef: React.MutableRefObject<Range | null>;
  outlineClauses: ContractOutlineClause[];
  setOutlineClauses: React.Dispatch<
    React.SetStateAction<ContractOutlineClause[]>
  >;
  outlineClausesRef: React.MutableRefObject<ContractOutlineClause[]>;
  contractConditionsRef: React.MutableRefObject<
    Record<string, string | boolean>
  >;
  setContractConditions: React.Dispatch<
    React.SetStateAction<Record<string, string | boolean>>
  >;
  editingDerivedTpl: OfficeTemplate | null;
  setEditingDerivedTpl: React.Dispatch<
    React.SetStateAction<OfficeTemplate | null>
  >;
  editingClauseObj: NotaryClause | null;
  setEditingClauseObj: React.Dispatch<React.SetStateAction<NotaryClause | null>>;
  stashedMainDocRef: React.MutableRefObject<{
    bodyHtml: string;
    headerHtml: string;
    footerHtml: string;
    title: string;
  } | null>;
  setClauses: React.Dispatch<React.SetStateAction<NotaryClause[]>>;
  setTemplates: React.Dispatch<React.SetStateAction<OfficeTemplate[]>>;
  ensureBodyEditable: () => boolean;
  recordHistorySnapshot: (label?: string) => void;
  restoreSelectionBookmarkIfNeeded: () => void;
  getEditorZonesHtml: () => {
    bodyHtml: string;
    headerHtml: string;
    footerHtml: string;
  };
  syncPlaceholdersAndDraft: (
    customFieldValues?: Record<string, string>,
    customFieldInputTypes?: Record<string, VariableInputType>,
    customOutlineClauses?: ContractOutlineClause[],
    customPartyCards?: ContractPartyCard[]
  ) => void;
  saveNewClauseToCatalog: (
    title: string,
    category: string,
    contentHtml: string
  ) => Promise<void>;
  showToast: (msg: string) => void;
}

export function useEditorClauses({
  docTitle,
  setDocTitle,
  bodyEditorRef,
  headerEditorRef,
  footerEditorRef,
  savedRangeRef,
  outlineClauses,
  setOutlineClauses,
  outlineClausesRef,
  contractConditionsRef,
  setContractConditions,
  editingDerivedTpl,
  setEditingDerivedTpl,
  editingClauseObj,
  setEditingClauseObj,
  stashedMainDocRef,
  setClauses,
  setTemplates,
  ensureBodyEditable,
  recordHistorySnapshot,
  restoreSelectionBookmarkIfNeeded,
  getEditorZonesHtml,
  syncPlaceholdersAndDraft,
  saveNewClauseToCatalog,
  showToast,
}: UseEditorClausesParams) {
  const handleScrollToLiveClause = useCallback(
    (domIndexOrId: number | string) => {
      if (!bodyEditorRef.current) return;
      scrollToContractClauseInDOM(bodyEditorRef.current, domIndexOrId);
    },
    [bodyEditorRef]
  );

  const handleToggleLiveClauseEnabled = useCallback(
    (clauseId: string) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      upgradeTopLevelHashSegmentsToContainers(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      const currentOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      const targetClause = currentOutline.find((c) => c.id === clauseId);
      if (targetClause?.locked) {
        showToast(
          'هذا البند مقفول لحمايته من التعطيل أو الحذف — افتح القفل أولاً'
        );
        return;
      }
      recordHistorySnapshot(
        targetClause
          ? `${targetClause.enabled ? 'تعطيل' : 'تفعيل'} البند: ${targetClause.title}`
          : 'تغيير حالة تفعيل بند'
      );
      const updatedOutline = toggleContractClauseEnabledInDOM(
        bodyEditorRef.current,
        currentOutline,
        clauseId
      );
      outlineClausesRef.current = updatedOutline;
      setOutlineClauses(updatedOutline);
      syncPlaceholdersAndDraft(undefined, undefined, updatedOutline);
      const afterClause = updatedOutline.find((c) => c.id === clauseId);
      if (afterClause) {
        showToast(
          afterClause.enabled
            ? `تم إعادة تفعيل البند "${afterClause.title}" في موضعه الدقيق داخل العقد`
            : `تم إخفاء البند "${afterClause.title}" من ورقة العقد والتصدير مؤقتاً`
        );
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      outlineClausesRef,
      recordHistorySnapshot,
      setOutlineClauses,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleUpdateContractCondition = useCallback(
    (key: string, value: string | boolean | null | undefined) => {
      if (!ensureBodyEditable()) return;
      const nextConds = { ...contractConditionsRef.current };
      if (value === null || value === undefined) {
        delete nextConds[key];
      } else {
        nextConds[key] = value;
      }
      contractConditionsRef.current = nextConds;
      setContractConditions(nextConds);

      if (!bodyEditorRef.current) return;
      upgradeTopLevelHashSegmentsToContainers(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      let workingOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        outlineClausesRef.current
      );

      let toggledCount = 0;
      if (value !== null && value !== undefined) {
        for (const clause of workingOutline) {
          if (
            clause.condition &&
            clause.condition.field === key &&
            !clause.locked
          ) {
            const condMatch =
              String(clause.condition.value).trim() === String(value).trim() ||
              (clause.condition.value === true &&
                (value === true || value === 'نعم')) ||
              (clause.condition.value === false &&
                (value === false || value === 'لا'));
            if (clause.enabled !== condMatch) {
              workingOutline = toggleContractClauseEnabledInDOM(
                bodyEditorRef.current,
                workingOutline,
                clause.id
              );
              toggledCount++;
            }
          }
        }
      }

      outlineClausesRef.current = workingOutline;
      setOutlineClauses(workingOutline);
      if (toggledCount > 0) {
        recordHistorySnapshot(
          `تحديث الشرط "${key}" وتبديل (${toggledCount}) بند مرتبط تلقائياً`
        );
        showToast(
          `تم تحديث الشرط "${key}" وتفعيل/إخفاء (${toggledCount}) بند مرتبط تلقائياً`
        );
      }
      syncPlaceholdersAndDraft(undefined, undefined, workingOutline);
    },
    [
      bodyEditorRef,
      contractConditionsRef,
      ensureBodyEditable,
      outlineClausesRef,
      recordHistorySnapshot,
      setContractConditions,
      setOutlineClauses,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleSetClauseCondition = useCallback(
    (clauseId: string, condition?: ClauseCondition) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      upgradeTopLevelHashSegmentsToContainers(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      let workingOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        outlineClausesRef.current
      ).map((c) => (c.id === clauseId ? { ...c, condition } : c));

      if (condition && condition.field in contractConditionsRef.current) {
        const currentVal = contractConditionsRef.current[condition.field];
        const target = workingOutline.find((c) => c.id === clauseId);
        if (target && !target.locked) {
          const condMatch =
            String(condition.value).trim() === String(currentVal).trim() ||
            (condition.value === true &&
              (currentVal === true || currentVal === 'نعم')) ||
            (condition.value === false &&
              (currentVal === false || currentVal === 'لا'));
          if (target.enabled !== condMatch) {
            workingOutline = toggleContractClauseEnabledInDOM(
              bodyEditorRef.current,
              workingOutline,
              clauseId
            );
          }
        }
      }

      outlineClausesRef.current = workingOutline;
      setOutlineClauses(workingOutline);
      syncPlaceholdersAndDraft(undefined, undefined, workingOutline);
      showToast(
        condition
          ? `تم ربط البند بالشرط "${condition.field}"`
          : 'تم إلغاء الشرط المرتبط بالبند'
      );
    },
    [
      bodyEditorRef,
      contractConditionsRef,
      ensureBodyEditable,
      outlineClausesRef,
      setOutlineClauses,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleToggleLiveClauseLocked = useCallback(
    (clauseId: string) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      upgradeTopLevelHashSegmentsToContainers(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      const currentOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      const target = currentOutline.find((c) => c.id === clauseId);
      if (!target) return;
      const nextLocked = !target.locked;
      const updatedOutline = setContractClauseLockedInDOM(
        bodyEditorRef.current,
        currentOutline,
        clauseId,
        nextLocked
      );
      outlineClausesRef.current = updatedOutline;
      setOutlineClauses(updatedOutline);
      syncPlaceholdersAndDraft(undefined, undefined, updatedOutline);
      showToast(
        nextLocked
          ? `تم قفل البند "${target.title}" لحمايته من التعديل أو الحذف العرضي`
          : `تم فتح قفل البند "${target.title}"`
      );
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      outlineClausesRef,
      setOutlineClauses,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleReorderLiveClauses = useCallback(
    (draggedClauseId: string, targetClauseId: string) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot('قبل إعادة ترتيب بنود العقد بالسحب والإفلات');
      upgradeTopLevelHashSegmentsToContainers(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      const currentOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        outlineClausesRef.current
      );
      const updatedOutline = reorderLiveContractClausesInDOM(
        bodyEditorRef.current,
        currentOutline,
        draggedClauseId,
        targetClauseId
      );
      outlineClausesRef.current = updatedOutline;
      setOutlineClauses(updatedOutline);
      syncPlaceholdersAndDraft(undefined, undefined, updatedOutline);
      showToast('تم تحديث ترتيب البنود داخل العقد بنجاح');
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      outlineClausesRef,
      recordHistorySnapshot,
      setOutlineClauses,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleMoveLiveClauseInDoc = useCallback(
    (domIndex: number, direction: 'up' | 'down') => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      const moved = moveContractClauseInDOM(
        bodyEditorRef.current,
        domIndex,
        direction
      );
      if (moved) {
        syncPlaceholdersAndDraft();
        showToast(
          `تم تحريك البند ${direction === 'up' ? 'للأعلى' : 'للأسفل'} داخل العقد`
        );
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleDeleteLiveClauseFromDoc = useCallback(
    (domIndexOrId: number | string) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      const targetClause = outlineClausesRef.current.find(
        (c) => c.domIndex === domIndexOrId || c.id === domIndexOrId
      );
      if (targetClause?.locked) {
        showToast('لا يمكن حذف بند مقفول — افتح القفل أولاً');
        return;
      }
      recordHistorySnapshot('قبل حذف بند من العقد');
      const deleted = deleteContractClauseFromDOM(
        bodyEditorRef.current,
        domIndexOrId
      );
      if (deleted) {
        const filteredOutline = outlineClausesRef.current.filter(
          (c) => c.domIndex !== domIndexOrId && c.id !== domIndexOrId
        );
        outlineClausesRef.current = filteredOutline;
        syncPlaceholdersAndDraft(undefined, undefined, filteredOutline);
        showToast('تم حذف البند من العقد الحالي');
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      outlineClausesRef,
      recordHistorySnapshot,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleRenameLiveClauseInDoc = useCallback(
    (domIndexOrId: number | string, newTitle: string) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      const targetClause = outlineClausesRef.current.find(
        (c) => c.domIndex === domIndexOrId || c.id === domIndexOrId
      );
      if (targetClause?.locked) {
        showToast('لا يمكن تعديل عنوان بند مقفول — افتح القفل أولاً');
        return;
      }
      recordHistorySnapshot('قبل تعديل عنوان بند في العقد');
      const renamed = renameContractClauseInDOM(
        bodyEditorRef.current,
        domIndexOrId,
        newTitle
      );
      if (renamed) {
        syncPlaceholdersAndDraft();
        showToast(`تم تحديث عنوان البند إلى "# ${newTitle}"`);
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      outlineClausesRef,
      recordHistorySnapshot,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleToggleClauseInDoc = useCallback(
    (clause: NotaryClause) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();

      const existing = bodyEditorRef.current.querySelector(
        `[data-clause-id="${clause.id}"]`
      ) as HTMLElement | null;

      if (existing) {
        existing.remove();
        normalizeNotaryContainerDOM(bodyEditorRef.current);
        syncPlaceholdersAndDraft();
        showToast(`تم إزالة البند "${clause.title}" من ورقة العقد`);
      } else {
        const containerHtml = `<div class="clause-container" data-clause-id="${
          clause.id
        }" data-clause-title="${escapeHtml(clause.title)}">${clause.contentHtml}</div>`;
        insertHtmlAtSelection(
          bodyEditorRef.current,
          containerHtml,
          savedRangeRef.current
        );
        syncPlaceholdersAndDraft();
        showToast(`تم إدراج البند "${clause.title}" في العقد الحالي`);
      }
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      savedRangeRef,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleInsertClauseAtCaret = useCallback(
    (clause: NotaryClause) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;
      const cleanTitle = clause.title.replace(/^#+\s*/, '').trim();
      const rawContent = (clause.contentHtml || '').trim();
      const plainContent = rawContent.replace(/<[^>]+>/g, '').trim();
      const startsWithHash = plainContent.startsWith('#');
      const headingLine = startsWithHash
        ? ''
        : `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># ${escapeHtml(
            cleanTitle
          )}</span></p>`;
      const bodyPart = rawContent || `<p dir="rtl" style="${pStyle}"><br></p>`;
      insertHtmlAtSelection(
        bodyEditorRef.current,
        `${headingLine}${bodyPart}`,
        savedRangeRef.current
      );
      syncPlaceholdersAndDraft();
      showToast(`تم إدراج البند "${cleanTitle}" عند المؤشر`);
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      savedRangeRef,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleInsertNewClauseHeadingAtCaret = useCallback(
    async (customTitle?: string) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      const { title } = insertOrWrapNewClauseAtSelection(
        bodyEditorRef.current,
        savedRangeRef.current,
        outlineClauses.length + 1,
        customTitle
      );
      syncPlaceholdersAndDraft();
      showToast(
        `تم إدراج عنوان البند "# ${title}" — اكتب محتواه بحرية في الورقة`
      );
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      outlineClauses.length,
      recordHistorySnapshot,
      savedRangeRef,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleSaveNewClause = useCallback(
    async (title: string, category: string, contentHtml?: string) => {
      const bodyHtml =
        contentHtml ||
        bodyEditorRef.current?.innerHTML ||
        INITIAL_EMPTY_PARAGRAPH;
      await saveNewClauseToCatalog(title, category, bodyHtml);
    },
    [bodyEditorRef, saveNewClauseToCatalog]
  );

  const handleSaveSelectionAsClause = useCallback(
    async (title: string, category: string) => {
      restoreSelectionBookmarkIfNeeded();
      const sel = window.getSelection();
      let htmlSnippet = '';
      if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
        const frag = sel.getRangeAt(0).cloneContents();
        const div = document.createElement('div');
        div.appendChild(frag);
        normalizeNotaryContainerDOM(div);
        htmlSnippet = div.innerHTML;
      } else if (bodyEditorRef.current) {
        htmlSnippet = bodyEditorRef.current.innerHTML;
      }

      await handleSaveNewClause(title, category, htmlSnippet);
    },
    [
      bodyEditorRef,
      handleSaveNewClause,
      restoreSelectionBookmarkIfNeeded,
    ]
  );

  const handleEditClauseInEditor = useCallback(
    (clause: NotaryClause) => {
      if (!bodyEditorRef.current) return;
      if (!editingDerivedTpl && !editingClauseObj) {
        stashedMainDocRef.current = {
          ...getEditorZonesHtml(),
          title: docTitle,
        };
      }
      setEditingDerivedTpl(null);
      setEditingClauseObj(clause);
      bodyEditorRef.current.innerHTML =
        clause.contentHtml || INITIAL_EMPTY_PARAGRAPH;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
      showToast(`أنت الآن تعدل صياغة البند "${clause.title}" داخل ورقة الـ A4`);
    },
    [
      bodyEditorRef,
      docTitle,
      editingClauseObj,
      editingDerivedTpl,
      getEditorZonesHtml,
      setEditingClauseObj,
      setEditingDerivedTpl,
      showToast,
      stashedMainDocRef,
    ]
  );

  const handleFinishEditingClause = useCallback(
    async (saveChanges: boolean) => {
      if (!editingClauseObj || !bodyEditorRef.current) return;
      if (saveChanges) {
        const updatedClause: NotaryClause = {
          ...editingClauseObj,
          contentHtml: bodyEditorRef.current.innerHTML,
          updatedAt: new Date().toISOString(),
        };
        await saveNotaryClause(updatedClause);
        setClauses(await loadNotaryClauses());
        showToast(`تم حفظ التعديلات على البند "${updatedClause.title}"`);
      }
      setEditingClauseObj(null);
      if (stashedMainDocRef.current && bodyEditorRef.current) {
        bodyEditorRef.current.innerHTML = stashedMainDocRef.current.bodyHtml;
        if (headerEditorRef.current)
          headerEditorRef.current.innerHTML =
            stashedMainDocRef.current.headerHtml;
        if (footerEditorRef.current)
          footerEditorRef.current.innerHTML =
            stashedMainDocRef.current.footerHtml;
        setDocTitle(stashedMainDocRef.current.title);
        stashedMainDocRef.current = null;
        normalizeNotaryContainerDOM(bodyEditorRef.current);
      }
    },
    [
      bodyEditorRef,
      editingClauseObj,
      footerEditorRef,
      headerEditorRef,
      setClauses,
      setDocTitle,
      setEditingClauseObj,
      showToast,
      stashedMainDocRef,
    ]
  );

  const handleEditDerivedTemplateInEditor = useCallback(
    (tpl: OfficeTemplate) => {
      if (!bodyEditorRef.current) return;
      if (!editingDerivedTpl && !editingClauseObj) {
        stashedMainDocRef.current = {
          ...getEditorZonesHtml(),
          title: docTitle,
        };
      }
      setEditingClauseObj(null);
      setEditingDerivedTpl(tpl);
      bodyEditorRef.current.innerHTML = tpl.bodyHtml || INITIAL_EMPTY_PARAGRAPH;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
      if (headerEditorRef.current)
        headerEditorRef.current.innerHTML = tpl.headerHtml || '';
      if (footerEditorRef.current)
        footerEditorRef.current.innerHTML = tpl.footerHtml || '';
      showToast(`أنت الآن تعدل قالب المكتب "${tpl.name}" داخل ورقة الـ A4`);
    },
    [
      bodyEditorRef,
      docTitle,
      editingClauseObj,
      editingDerivedTpl,
      footerEditorRef,
      getEditorZonesHtml,
      headerEditorRef,
      setEditingClauseObj,
      setEditingDerivedTpl,
      showToast,
      stashedMainDocRef,
    ]
  );

  const handleFinishEditingDerivedTemplate = useCallback(
    async (saveChanges: boolean) => {
      if (!editingDerivedTpl || !bodyEditorRef.current) return;
      if (saveChanges) {
        const bodyHtml = bodyEditorRef.current.innerHTML;
        const headerHtml = headerEditorRef.current?.innerHTML || '';
        const footerHtml = footerEditorRef.current?.innerHTML || '';
        const placeholders = extractPlaceholdersFromHtml(
          headerHtml,
          bodyHtml,
          footerHtml
        );
        const updatedOfficeTpl: OfficeTemplate = {
          ...editingDerivedTpl,
          bodyHtml,
          headerHtml,
          footerHtml,
          extractedPlaceholders: placeholders,
          updatedAt: new Date().toISOString(),
        };
        await saveOfficeTemplate(updatedOfficeTpl);
        setTemplates(await loadOfficeTemplates());
        showToast(
          `تم حفظ تعديلات قالب المكتب "${updatedOfficeTpl.name}" بنجاح`
        );
      }
      setEditingDerivedTpl(null);
      if (stashedMainDocRef.current && bodyEditorRef.current) {
        bodyEditorRef.current.innerHTML = stashedMainDocRef.current.bodyHtml;
        if (headerEditorRef.current)
          headerEditorRef.current.innerHTML =
            stashedMainDocRef.current.headerHtml;
        if (footerEditorRef.current)
          footerEditorRef.current.innerHTML =
            stashedMainDocRef.current.footerHtml;
        setDocTitle(stashedMainDocRef.current.title);
        stashedMainDocRef.current = null;
        normalizeNotaryContainerDOM(bodyEditorRef.current);
      }
    },
    [
      bodyEditorRef,
      editingDerivedTpl,
      footerEditorRef,
      headerEditorRef,
      setDocTitle,
      setEditingDerivedTpl,
      setTemplates,
      showToast,
      stashedMainDocRef,
    ]
  );

  return {
    handleScrollToLiveClause,
    handleToggleLiveClauseEnabled,
    handleUpdateContractCondition,
    handleSetClauseCondition,
    handleToggleLiveClauseLocked,
    handleReorderLiveClauses,
    handleMoveLiveClauseInDoc,
    handleDeleteLiveClauseFromDoc,
    handleRenameLiveClauseInDoc,
    handleToggleClauseInDoc,
    handleInsertClauseAtCaret,
    handleInsertNewClauseHeadingAtCaret,
    handleSaveNewClause,
    handleSaveSelectionAsClause,
    handleEditClauseInEditor,
    handleFinishEditingClause,
    handleEditDerivedTemplateInEditor,
    handleFinishEditingDerivedTemplate,
  };
}
