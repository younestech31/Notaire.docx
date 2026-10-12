'use client';

import React, { useCallback, useRef } from 'react';
import {
  ClauseVariableGroup,
  ContractOutlineClause,
  ContractPartyCard,
  NotaryClause,
  NotaryClerk,
  OfficeTemplate,
  SavedDocument,
  VariableInputType,
} from '@/lib/types';
import {
  loadSavedDocuments,
  saveActiveDraftSession,
  saveDocumentRecord,
} from '@/lib/storage';
import {
  computeDocumentMetrics,
  injectNotarialPaginationIntoHtml,
} from '@/lib/notarial-pagination';
import {
  applyPartyCardsToFieldValues,
  extractLiveContractClausesFromDOM,
} from '@/lib/clause-outline-engine';
import {
  decorateSmartTagsInDOM,
  extractPlaceholdersFromHtml,
  extractPlaceholdersGroupedByClause,
  restoreSerializedSelection,
  serializeCurrentSelection,
  syncSmartTagsFilledStateInDOM,
} from '@/lib/editor-utils';
import { mergePlaceholdersIntoHtml } from '@/lib/docx-engine';

interface UseEditorSyncParams {
  docId: string;
  docTitle: string;
  pageNumberingEnabled: boolean;
  showHeaderFooter: boolean;
  previewMergedMode: boolean;
  editingDerivedTpl: OfficeTemplate | null;
  editingClauseObj: NotaryClause | null;
  selectedEstateId: string;
  selectedLotNumber: string;
  activeClerk: NotaryClerk;
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  headerEditorRef: React.RefObject<HTMLDivElement | null>;
  footerEditorRef: React.RefObject<HTMLDivElement | null>;
  outlineClausesRef: React.MutableRefObject<ContractOutlineClause[]>;
  contractConditionsRef: React.MutableRefObject<
    Record<string, string | boolean>
  >;
  partyCards: ContractPartyCard[];
  partyCardsRef: React.MutableRefObject<ContractPartyCard[]>;
  fieldValues: Record<string, string>;
  fieldValuesRef: React.MutableRefObject<Record<string, string>>;
  fieldInputTypes: Record<string, VariableInputType>;
  fieldInputTypesRef: React.MutableRefObject<Record<string, VariableInputType>>;
  documentsRef: React.MutableRefObject<SavedDocument[]>;
  activeDocumentRef: React.MutableRefObject<SavedDocument | null>;
  setActiveClauseIdsInDoc: React.Dispatch<React.SetStateAction<string[]>>;
  setOutlineClauses: React.Dispatch<
    React.SetStateAction<ContractOutlineClause[]>
  >;
  setExtractedPlaceholders: React.Dispatch<React.SetStateAction<string[]>>;
  setClauseGroups: React.Dispatch<React.SetStateAction<ClauseVariableGroup[]>>;
  setDocMetrics: React.Dispatch<
    React.SetStateAction<ReturnType<typeof computeDocumentMetrics>>
  >;
  setMergedPreviewHtml: React.Dispatch<React.SetStateAction<string>>;
  setPrintPaginatedHtml: React.Dispatch<React.SetStateAction<string>>;
  setAutoSaveState: React.Dispatch<
    React.SetStateAction<'saved' | 'saving' | 'error'>
  >;
  setDocuments: (nextDocs: SavedDocument[]) => void;
  setActiveDocument: React.Dispatch<React.SetStateAction<SavedDocument | null>>;
  getEditorZonesHtml: () => {
    bodyHtml: string;
    headerHtml: string;
    footerHtml: string;
  };
  showToast: (msg: string) => void;
}

export function useEditorSync({
  docId,
  docTitle,
  pageNumberingEnabled,
  showHeaderFooter,
  previewMergedMode,
  editingDerivedTpl,
  editingClauseObj,
  selectedEstateId,
  selectedLotNumber,
  activeClerk,
  bodyEditorRef,
  headerEditorRef,
  footerEditorRef,
  outlineClausesRef,
  contractConditionsRef,
  partyCards,
  partyCardsRef,
  fieldValues,
  fieldValuesRef,
  fieldInputTypes,
  fieldInputTypesRef,
  documentsRef,
  activeDocumentRef,
  setActiveClauseIdsInDoc,
  setOutlineClauses,
  setExtractedPlaceholders,
  setClauseGroups,
  setDocMetrics,
  setMergedPreviewHtml,
  setPrintPaginatedHtml,
  setAutoSaveState,
  setDocuments,
  setActiveDocument,
  getEditorZonesHtml,
  showToast,
}: UseEditorSyncParams) {
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const debouncedSyncTimerRef = useRef<ReturnType<typeof setTimeout> | null>(
    null
  );

  const clearAutoSaveTimer = useCallback(() => {
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }
  }, []);

  const refreshActiveClausesInDOM = useCallback(
    (existingOutline?: ContractOutlineClause[]) => {
      if (!bodyEditorRef.current) return;
      const containers = Array.from(
        bodyEditorRef.current.querySelectorAll('[data-clause-id]')
      ) as HTMLElement[];
      const ids = containers
        .map((el) => el.getAttribute('data-clause-id') || '')
        .filter(Boolean);
      setActiveClauseIdsInDoc(ids);
      const liveOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        existingOutline ?? outlineClausesRef.current
      );
      outlineClausesRef.current = liveOutline;
      setOutlineClauses(liveOutline);
    },
    [bodyEditorRef, outlineClausesRef, setActiveClauseIdsInDoc, setOutlineClauses]
  );

  const syncPlaceholdersAndDraft = useCallback(
    (
      customFieldValues?: Record<string, string>,
      customFieldInputTypes?: Record<string, VariableInputType>,
      customOutlineClauses?: ContractOutlineClause[],
      customPartyCards?: ContractPartyCard[]
    ) => {
      const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
      const found = extractPlaceholdersFromHtml(headerHtml, bodyHtml, footerHtml);
      setExtractedPlaceholders(found);
      setClauseGroups(
        extractPlaceholdersGroupedByClause(bodyEditorRef.current, found)
      );
      setDocMetrics(computeDocumentMetrics(bodyHtml, false));
      const baseOutline = customOutlineClauses ?? outlineClausesRef.current;
      const liveOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current,
        baseOutline
      );
      outlineClausesRef.current = liveOutline;
      setOutlineClauses(liveOutline);
      refreshActiveClausesInDOM(liveOutline);

      const activeCards =
        customPartyCards ?? partyCardsRef.current ?? partyCards;
      let activeValues =
        customFieldValues ?? fieldValuesRef.current ?? fieldValues;
      if (activeCards.length > 0) {
        activeValues = applyPartyCardsToFieldValues(
          activeCards,
          activeValues,
          found
        );
      }
      const activeInputTypes =
        customFieldInputTypes ?? fieldInputTypesRef.current ?? fieldInputTypes;
      syncSmartTagsFilledStateInDOM(bodyEditorRef.current, activeValues);
      syncSmartTagsFilledStateInDOM(headerEditorRef.current, activeValues);
      syncSmartTagsFilledStateInDOM(footerEditorRef.current, activeValues);

      const rawMergedForPagination = mergePlaceholdersIntoHtml(
        bodyHtml,
        activeValues
      );
      if (previewMergedMode) {
        setMergedPreviewHtml(
          injectNotarialPaginationIntoHtml(
            rawMergedForPagination,
            pageNumberingEnabled,
            true
          )
        );
      }

      if (editingDerivedTpl || editingClauseObj) {
        return;
      }

      const existingDoc =
        documentsRef.current.find((d) => d.id === docId) ||
        (activeDocumentRef.current?.id === docId
          ? activeDocumentRef.current
          : null);

      const inferredClient =
        existingDoc?.clientName ||
        activeCards[0]?.fullName ||
        activeValues['الطرف_الأول_الاسم'] ||
        activeValues['البائع'] ||
        activeValues['المؤجر'] ||
        activeValues['الموكل'] ||
        '';

      const nowIso = new Date().toISOString();
      const draft: SavedDocument = {
        ...(existingDoc || {}),
        id: docId,
        title: docTitle,
        bodyHtml,
        headerHtml,
        footerHtml,
        pageNumberingEnabled,
        showHeaderFooter,
        fieldValues: activeValues,
        fieldInputTypes: activeInputTypes,
        outlineClauses: liveOutline,
        contractConditions: contractConditionsRef.current,
        partyCards: activeCards,
        selectedEstateId,
        selectedLotNumber,
        folderId: existingDoc?.folderId ?? null,
        clerkId: existingDoc?.clerkId || activeClerk.id,
        clerkName: existingDoc?.clerkName || activeClerk.name,
        contractNumber:
          existingDoc?.contractNumber ||
          activeValues['رقم_الفهرس'] ||
          `2026/${Math.floor(100 + Math.random() * 900)}`,
        year: existingDoc?.year || new Date().getFullYear(),
        clientName: inferredClient,
        status: existingDoc?.status || 'editing',
        derivedDocuments: existingDoc?.derivedDocuments || [],
        updatedAt: nowIso,
        createdAt: existingDoc?.createdAt || nowIso,
      };
      setAutoSaveState('saving');
      if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = setTimeout(async () => {
        try {
          setDocMetrics(computeDocumentMetrics(bodyHtml, true));
          setPrintPaginatedHtml(
            injectNotarialPaginationIntoHtml(
              rawMergedForPagination,
              pageNumberingEnabled,
              true
            )
          );
          await saveActiveDraftSession(draft);
          await saveDocumentRecord(draft);
          const updatedDocs = await loadSavedDocuments();
          setDocuments(updatedDocs);
          setActiveDocument((prev) => (prev?.id === draft.id ? draft : prev));
          if (activeDocumentRef.current?.id === draft.id) {
            activeDocumentRef.current = draft;
          }
          setAutoSaveState('saved');
        } catch (err) {
          setAutoSaveState('error');
          showToast(
            err instanceof Error
              ? `تعذر حفظ العقد: ${err.message}`
              : 'تعذر حفظ العقد في مخزن المكتب'
          );
        }
      }, 450);
    },
    [
      activeClerk.id,
      activeClerk.name,
      activeDocumentRef,
      bodyEditorRef,
      contractConditionsRef,
      docId,
      docTitle,
      documentsRef,
      editingClauseObj,
      editingDerivedTpl,
      fieldInputTypes,
      fieldInputTypesRef,
      fieldValues,
      fieldValuesRef,
      footerEditorRef,
      getEditorZonesHtml,
      headerEditorRef,
      outlineClausesRef,
      pageNumberingEnabled,
      partyCards,
      partyCardsRef,
      previewMergedMode,
      refreshActiveClausesInDOM,
      selectedEstateId,
      selectedLotNumber,
      setActiveDocument,
      setAutoSaveState,
      setClauseGroups,
      setDocMetrics,
      setDocuments,
      setExtractedPlaceholders,
      setMergedPreviewHtml,
      setOutlineClauses,
      setPrintPaginatedHtml,
      showHeaderFooter,
      showToast,
    ]
  );

  const flushSyncPlaceholdersAndDraft = useCallback(() => {
    if (debouncedSyncTimerRef.current) {
      clearTimeout(debouncedSyncTimerRef.current);
      debouncedSyncTimerRef.current = null;
    }
    syncPlaceholdersAndDraft();
  }, [syncPlaceholdersAndDraft]);

  const decorateEditorZonesPreservingSelection = useCallback(() => {
    const zones = {
      body: bodyEditorRef.current,
      header: headerEditorRef.current,
      footer: footerEditorRef.current,
    };
    const savedSel = serializeCurrentSelection(zones);
    if (bodyEditorRef.current) decorateSmartTagsInDOM(bodyEditorRef.current);
    if (headerEditorRef.current) decorateSmartTagsInDOM(headerEditorRef.current);
    if (footerEditorRef.current) decorateSmartTagsInDOM(footerEditorRef.current);
    restoreSerializedSelection(zones, savedSel);
    syncPlaceholdersAndDraft();
  }, [
    bodyEditorRef,
    footerEditorRef,
    headerEditorRef,
    syncPlaceholdersAndDraft,
  ]);

  return {
    autoSaveTimerRef,
    debouncedSyncTimerRef,
    clearAutoSaveTimer,
    refreshActiveClausesInDOM,
    syncPlaceholdersAndDraft,
    flushSyncPlaceholdersAndDraft,
    decorateEditorZonesPreservingSelection,
  };
}
