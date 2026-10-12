'use client';

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { CheckCircle2 } from 'lucide-react';
import EditorRibbon from './EditorRibbon';
import EditorTopBar from './EditorTopBar';
import EditorA4Canvas from './editor/EditorA4Canvas';
import {
  EditorActivityRail,
  EditorModalsHost,
} from './editor/EditorModalsHost';
import {
  EditorBottomStatusBar,
  EditorFindReplaceBar,
  EditorModeAndSealBanners,
} from './editor/EditorStatusBars';
import { InlineVariablePopoverState } from './InlineVariablePopover';
import SidebarWorkspace, { SidebarTab } from './SidebarWorkspace';
import {
  ClauseCondition,
  ContractFolder,
  ContractOutlineClause,
  ContractPartyCard,
  DocumentRevision,
  NotaryClause,
  OfficeTemplate,
  OfficeTemplateKind,
  SavedContractDerivedDoc,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  SubdivisionEstate,
  SubdivisionLot,
  VariableInputType,
} from '@/lib/types';
import {
  getRuntimeInfo,
  importFullBackupBundle,
  loadDocumentRevisions,
  loadNotaryClauses,
  loadOfficeTemplates,
  loadSavedDocuments,
  saveActiveDraftSession,
  saveDocumentRecord,
  saveNotaryClause,
  saveOfficeTemplate,
} from '@/lib/storage';
import {
  computeDocumentMetrics,
  injectNotarialPaginationIntoHtml,
} from '@/lib/notarial-pagination';
import {
  applyPartyCardsToFieldValues,
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
  convertSelectionToSmartTag,
  decorateSmartTagsInDOM,
  escapeHtml,
  extractPlaceholdersFromHtml,
  extractPlaceholdersGroupedByClause,
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
  restoreSerializedSelection,
  sanitizePastedWordHTML,
  serializeCurrentSelection,
  syncSmartTagsFilledStateInDOM,
} from '@/lib/editor-utils';
import { insertTableRowAtCell } from '@/lib/table-grid';
import {
  DocxImportResult,
  downloadNotaryDocx,
  importNotaryDocxFile,
  lookupPlaceholderValue,
  mergePlaceholdersIntoHtml,
  normalizePlaceholderKey,
} from '@/lib/docx-engine';
import {
  buildFullSubdivisionTableHtml,
  buildSubdivisionLotClauseHtml,
  downloadOfficeBackupBundleFile,
  downloadTemplatesJsonFile,
  exportContractToWord,
  generateAndDownloadDerivedDoc,
  importDocxFilesAsTemplates,
  importTemplatesJsonFile,
} from '@/lib/document-io-workflows';
import { tryRestorePersistedOfficeDirectory } from '@/lib/desktop-bridge';
import {
  INITIAL_EMPTY_PARAGRAPH,
  useDocumentSession,
} from '@/hooks/useDocumentSession';
import { useOfficeCatalog } from '@/hooks/useOfficeCatalog';
import { useNotaryLifecycle } from '@/hooks/useNotaryLifecycle';
import { useEditorDomCommands } from '@/hooks/useEditorDomCommands';
import { useFindAndReplace } from '@/hooks/useFindAndReplace';

export default function NotaryEditorApp() {
  // Status Toast Notification
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const showToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((prev) => (prev === msg ? null : prev));
    }, 3400);
  }, []);

  // 1. Document Editing Session Hook
  const {
    bodyEditorRef,
    headerEditorRef,
    footerEditorRef,
    docId,
    setDocId,
    docTitle,
    setDocTitle,
    showHeaderFooter,
    setShowHeaderFooter,
    pageNumberingEnabled,
    setPageNumberingEnabled,
    previewMergedMode,
    setPreviewMergedMode,
    mergedPreviewHtml,
    setMergedPreviewHtml,
    printPaginatedHtml,
    setPrintPaginatedHtml,
    zoom,
    setZoom,
    editingDerivedTpl,
    setEditingDerivedTpl,
    editingClauseObj,
    setEditingClauseObj,
    stashedMainDocRef,
    activeClauseIdsInDoc,
    setActiveClauseIdsInDoc,
    outlineClauses,
    setOutlineClauses,
    outlineClausesRef,
    contractConditions,
    setContractConditions,
    contractConditionsRef,
    partyCards,
    setPartyCards,
    partyCardsRef,
    fieldValues,
    setFieldValues,
    fieldValuesRef,
    fieldInputTypes,
    setFieldInputTypes,
    fieldInputTypesRef,
    extractedPlaceholders,
    setExtractedPlaceholders,
    setClauseGroups,
    selectedEstateId,
    setSelectedEstateId,
    selectedLotNumber,
    setSelectedLotNumber,
    autoSaveState,
    setAutoSaveState,
    docMetrics,
    setDocMetrics,
    getEditorZonesHtml,
    getSessionSnapshot,
  } = useDocumentSession();

  // 2. Office Reference Catalogs Hook
  const {
    templates,
    setTemplates,
    clauses,
    setClauses,
    estates,
    savedParties,
    savedProperties,
    clerks,
    activeClerk,
    folders,
    setFolders,
    lastBackupAt,
    setLastBackupAt,
    localDirState,
    setLocalDirState,
    reloadCatalogRecords,
    connectAndReloadLocalDirectory,
    handleCreateFolder,
    handleRenameFolder,
    handleMoveFolder,
    handleSaveEstate,
    handleDeleteEstate,
    handleDeleteClause,
    handleImportStandardClausesPack,
    handleSaveNewClause: saveNewClauseToCatalog,
    handleDeleteTemplate,
    handleSaveTemplateFromSnapshot,
    handleSaveCurrentPartyToDirectory: saveRolePartyToCatalog,
    handleSavePartyCardToDirectory,
    handleSaveCurrentPropertyToDirectory: savePropertyToCatalog,
    handleExportTemplatesJson,
    handleImportTemplatesJson,
    handleImportDocxAsTemplate,
  } = useOfficeCatalog(showToast);

  // 3. Notary Contract Lifecycle & Revisions Hook
  const {
    documents,
    setDocuments,
    documentsRef,
    revisions,
    setRevisions,
    activeDocument,
    setActiveDocument,
    activeDocumentRef,
    activeDerivedDoc,
    setActiveDerivedDoc,
    isContractLocked,
    setUnlockIndexedBodyForDoc,
    reloadLifecycleRecords,
    handleMoveDocument: moveDocumentInLifecycle,
    handleDeleteFolderAndMigrateDocs,
    handleUpdateDocumentMeta,
    handleSaveDocumentAndRevision,
    handleDeleteSnapshot,
    handleDeleteDocument,
    handleExportContractWord,
    handleGenerateDerivedDoc: runGenerateDerivedDoc,
  } = useNotaryLifecycle(activeClerk, showToast);

  const { unlockIndexedBody, isBodyLocked } = isContractLocked(
    docId,
    Boolean(editingDerivedTpl || editingClauseObj)
  );
  const setUnlockIndexedBody = useCallback(
    (value: React.SetStateAction<boolean>) => {
      setUnlockIndexedBodyForDoc(docId, value);
    },
    [docId, setUnlockIndexedBodyForDoc]
  );

  const reloadAllOfficeRecords = useCallback(async () => {
    const [catalogData, lifecycleData] = await Promise.all([
      reloadCatalogRecords(),
      reloadLifecycleRecords(),
    ]);
    return {
      ...catalogData,
      ...lifecycleData,
    };
  }, [reloadCatalogRecords, reloadLifecycleRecords]);

  // Sidebar & UI state
  const [sidebarTab, setSidebarTab] = useState<SidebarTab>('clauses');
  const [isSidebarOpen, setIsSidebarOpen] = useState<boolean>(false);
  const [inlinePopover, setInlinePopover] =
    useState<InlineVariablePopoverState | null>(null);

  // Modals state
  const [showVersionDiffModal, setShowVersionDiffModal] = useState<boolean>(false);
  const [diffComparisonRevision, setDiffComparisonRevision] =
    useState<DocumentRevision | null>(null);
  const [showMultiSourceModal, setShowMultiSourceModal] = useState<boolean>(false);
  const [showDocxPreviewModal, setShowDocxPreviewModal] = useState<boolean>(false);
  const [pendingDocxImport, setPendingDocxImport] =
    useState<DocxImportResult | null>(null);
  const [showSaveAsTemplateModal, setShowSaveAsTemplateModal] =
    useState<boolean>(false);
  const [showSnapshotsHistoryModal, setShowSnapshotsHistoryModal] =
    useState<boolean>(false);
  const [modalActiveBodyHtml, setModalActiveBodyHtml] = useState<string>(
    INITIAL_EMPTY_PARAGRAPH
  );
  const [showOnboardingTour, setShowOnboardingTour] = useState<boolean>(false);
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const ensureBodyEditable = useCallback(() => {
    if (isBodyLocked) {
      showToast(
        'متن هذا العقد مقفول رسمياً بعد الفهرسة والختم — اضغط «فتح قفل المتن استثنائياً» في الشريط العلوي إن لزم التعديل'
      );
      return false;
    }
    return true;
  }, [isBodyLocked, showToast]);

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
    []
  );

  // Refresh extracted [...] placeholders, live clauses, metrics, and trigger atomic per-contract auto-save
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

      const activeCards = customPartyCards ?? partyCardsRef.current ?? partyCards;
      let activeValues = customFieldValues ?? fieldValuesRef.current ?? fieldValues;
      if (activeCards.length > 0) {
        activeValues = applyPartyCardsToFieldValues(activeCards, activeValues, found);
      }
      const activeInputTypes =
        customFieldInputTypes ?? fieldInputTypesRef.current ?? fieldInputTypes;
      syncSmartTagsFilledStateInDOM(bodyEditorRef.current, activeValues);
      syncSmartTagsFilledStateInDOM(headerEditorRef.current, activeValues);
      syncSmartTagsFilledStateInDOM(footerEditorRef.current, activeValues);

      const rawMergedForPagination = mergePlaceholdersIntoHtml(bodyHtml, activeValues);
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
        (activeDocumentRef.current?.id === docId ? activeDocumentRef.current : null);

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
      docId,
      docTitle,
      documentsRef,
      editingClauseObj,
      editingDerivedTpl,
      fieldInputTypes,
      fieldValues,
      getEditorZonesHtml,
      pageNumberingEnabled,
      partyCards,
      previewMergedMode,
      refreshActiveClausesInDOM,
      selectedEstateId,
      selectedLotNumber,
      setDocuments,
      showHeaderFooter,
      showToast,
    ]
  );

  // Editor DOM Formatting & Undo/Redo Hook
  const {
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
    recordHistorySnapshot,
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
  } = useEditorDomCommands({
    docId,
    docTitle,
    authorName: activeClerk.name,
    bodyEditorRef,
    headerEditorRef,
    footerEditorRef,
    fieldValuesRef,
    ensureBodyEditable,
    getEditorZonesHtml,
    syncPlaceholdersAndDraft: () => syncPlaceholdersAndDraft(),
    setRevisions,
    setAutoSaveState,
    showToast,
  });

  // Arabic Find & Replace Hook
  const {
    showFindReplace,
    setShowFindReplace,
    findQuery,
    setFindQuery,
    replaceQuery,
    setReplaceQuery,
    ignoreArabicHamzaAndDiacritics,
    setIgnoreArabicHamzaAndDiacritics,
    activeMatchIdx,
    matchCount,
    handleNextMatch,
    handlePrevMatch,
    handleRunReplace,
  } = useFindAndReplace({
    bodyEditorRef,
    ensureBodyEditable,
    recordHistorySnapshot,
    syncPlaceholdersAndDraft: () => syncPlaceholdersAndDraft(),
    showToast,
  });

  const handleDeleteFolder = useCallback(
    async (folderId: string) => {
      await handleDeleteFolderAndMigrateDocs(folderId, folders, setFolders);
    },
    [folders, handleDeleteFolderAndMigrateDocs, setFolders]
  );

  const handleMoveDocument = useCallback(
    async (targetDocId: string, targetFolderId: string | null) => {
      await moveDocumentInLifecycle(targetDocId, targetFolderId, folders);
    },
    [folders, moveDocumentInLifecycle]
  );

  const handleSelectDocument = useCallback(
    (doc: SavedDocument) => {
      if (!bodyEditorRef.current) return;
      if (autoSaveTimerRef.current) {
        clearTimeout(autoSaveTimerRef.current);
        autoSaveTimerRef.current = null;
      }
      setUndoStack([]);
      setRedoStack([]);
      setDocId(doc.id);
      setDocTitle(doc.title);
      setShowHeaderFooter(Boolean(doc.showHeaderFooter));
      setPageNumberingEnabled(
        doc.pageNumberingEnabled !== undefined ? doc.pageNumberingEnabled : true
      );
      const isolatedValues = { ...(doc.fieldValues || {}) };
      const isolatedInputTypes = { ...(doc.fieldInputTypes || {}) };
      const isolatedOutline = [...(doc.outlineClauses || [])];
      const isolatedConditions = { ...(doc.contractConditions || {}) };
      const isolatedPartyCards = [...(doc.partyCards || [])];
      fieldValuesRef.current = isolatedValues;
      fieldInputTypesRef.current = isolatedInputTypes;
      outlineClausesRef.current = isolatedOutline;
      contractConditionsRef.current = isolatedConditions;
      partyCardsRef.current = isolatedPartyCards;
      setFieldValues(isolatedValues);
      setFieldInputTypes(isolatedInputTypes);
      setOutlineClauses(isolatedOutline);
      setContractConditions(isolatedConditions);
      setPartyCards(isolatedPartyCards);
      setInlinePopover(null);
      setSelectedEstateId(doc.selectedEstateId || '');
      setSelectedLotNumber(doc.selectedLotNumber || '');
      bodyEditorRef.current.innerHTML = doc.bodyHtml || INITIAL_EMPTY_PARAGRAPH;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
      if (headerEditorRef.current)
        headerEditorRef.current.innerHTML = doc.headerHtml || '';
      if (footerEditorRef.current)
        footerEditorRef.current.innerHTML = doc.footerHtml || '';
      setActiveDocument(doc);
      activeDocumentRef.current = doc;
      setActiveDerivedDoc(null);
      syncPlaceholdersAndDraft(
        isolatedValues,
        isolatedInputTypes,
        isolatedOutline,
        isolatedPartyCards
      );
    },
    [setRedoStack, setUndoStack, syncPlaceholdersAndDraft]
  );

  const handleCreateContractInFolder = useCallback(
    async (folderId: string | null) => {
      const newDocId = `doc_${Date.now()}`;
      const newDoc: SavedDocument = {
        id: newDocId,
        title: 'عقد توثيقي جديد',
        bodyHtml: INITIAL_EMPTY_PARAGRAPH,
        headerHtml: '',
        footerHtml: '',
        pageNumberingEnabled: true,
        showHeaderFooter: false,
        fieldValues: {},
        fieldInputTypes: {},
        outlineClauses: [],
        selectedEstateId: '',
        selectedLotNumber: '',
        folderId,
        clerkId: activeClerk.id,
        clerkName: activeClerk.name,
        contractNumber: `2026/${Math.floor(100 + Math.random() * 900)}`,
        status: 'draft',
        updatedAt: new Date().toISOString(),
        createdAt: new Date().toISOString(),
        derivedDocuments: [],
      };
      await saveDocumentRecord(newDoc);
      const refreshed = await loadSavedDocuments();
      setDocuments(refreshed);
      handleSelectDocument(newDoc);
      const folderName = folderId
        ? folders.find((f) => f.id === folderId)?.name || 'المجلد المختار'
        : 'الجذر الرئيسي';
      showToast(`تم إنشاء عقد جديد معزول داخل "${folderName}"`);
    },
    [activeClerk, folders, handleSelectDocument, setDocuments, showToast]
  );

  const handleSelectDerivedDoc = useCallback(
    (derivedDoc: SavedContractDerivedDoc) => {
      if (!bodyEditorRef.current) return;
      recordHistorySnapshot();
      setActiveDerivedDoc(derivedDoc);
      setDocTitle(derivedDoc.title);
      bodyEditorRef.current.innerHTML = derivedDoc.content || INITIAL_EMPTY_PARAGRAPH;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
      if (headerEditorRef.current)
        headerEditorRef.current.innerHTML = derivedDoc.headerHtml || '';
      if (footerEditorRef.current)
        footerEditorRef.current.innerHTML = derivedDoc.footerHtml || '';
      if (derivedDoc.fieldValues) {
        fieldValuesRef.current = { ...derivedDoc.fieldValues };
        setFieldValues({ ...derivedDoc.fieldValues });
      }
      syncPlaceholdersAndDraft(derivedDoc.fieldValues);
      showToast(`تم فتح وثيقة المشتق: ${derivedDoc.title}`);
    },
    [recordHistorySnapshot, syncPlaceholdersAndDraft, showToast]
  );

  // Initial Load from OfficeStore
  useEffect(() => {
    let mounted = true;
    async function initWorkspace() {
      try {
        const restoredDir = await tryRestorePersistedOfficeDirectory();
        if (mounted) {
          setLocalDirState(restoredDir);
        }
        await getRuntimeInfo();
      } catch (err) {
        if (mounted) {
          setAutoSaveState('error');
          showToast(
            err instanceof Error
              ? err.message
              : 'تعذر التحقق من مجلد بيانات المكتب'
          );
        }
      }

      try {
        const {
          loadedDocs,
          loadedEstates,
          activeDraft,
        } = await reloadAllOfficeRecords();
        if (!mounted) return;

        if (loadedEstates.length > 0) {
          setSelectedEstateId(loadedEstates[0].id);
        }

        if (activeDraft && bodyEditorRef.current) {
          setDocId(activeDraft.id || `doc_${Date.now()}`);
          setDocTitle(activeDraft.title || 'عقد توثيقي جديد');
          setShowHeaderFooter(Boolean(activeDraft.showHeaderFooter));
          setPageNumberingEnabled(
            activeDraft.pageNumberingEnabled !== undefined
              ? activeDraft.pageNumberingEnabled
              : true
          );
          const initialVals = activeDraft.fieldValues || {};
          const initialTypes = activeDraft.fieldInputTypes || {};
          const initialOutline = activeDraft.outlineClauses || [];
          const initialConditions = activeDraft.contractConditions || {};
          const initialPartyCards = activeDraft.partyCards || [];
          fieldValuesRef.current = initialVals;
          fieldInputTypesRef.current = initialTypes;
          outlineClausesRef.current = initialOutline;
          contractConditionsRef.current = initialConditions;
          partyCardsRef.current = initialPartyCards;
          setFieldValues(initialVals);
          setFieldInputTypes(initialTypes);
          setOutlineClauses(initialOutline);
          setContractConditions(initialConditions);
          setPartyCards(initialPartyCards);
          setSelectedEstateId(activeDraft.selectedEstateId || '');
          setSelectedLotNumber(activeDraft.selectedLotNumber || '');

          bodyEditorRef.current.innerHTML = activeDraft.bodyHtml || INITIAL_EMPTY_PARAGRAPH;
          normalizeNotaryContainerDOM(bodyEditorRef.current);
          if (headerEditorRef.current && activeDraft.headerHtml) {
            headerEditorRef.current.innerHTML = activeDraft.headerHtml;
          }
          if (footerEditorRef.current && activeDraft.footerHtml) {
            footerEditorRef.current.innerHTML = activeDraft.footerHtml;
          }
          const found = extractPlaceholdersFromHtml(
            activeDraft.headerHtml || '',
            activeDraft.bodyHtml || '',
            activeDraft.footerHtml || ''
          );
          setExtractedPlaceholders(found);
          setClauseGroups(
            extractPlaceholdersGroupedByClause(bodyEditorRef.current, found)
          );
          const liveOutline = extractLiveContractClausesFromDOM(
            bodyEditorRef.current,
            initialOutline
          );
          outlineClausesRef.current = liveOutline;
          setOutlineClauses(liveOutline);
          setDocMetrics(computeDocumentMetrics(activeDraft.bodyHtml || ''));
          refreshActiveClausesInDOM(liveOutline);
          syncSmartTagsFilledStateInDOM(bodyEditorRef.current, initialVals);
          const matchingSaved = loadedDocs.find((d) => d.id === activeDraft.id);
          if (matchingSaved) {
            setActiveDocument(matchingSaved);
            activeDocumentRef.current = matchingSaved;
          }
        } else if (bodyEditorRef.current) {
          bodyEditorRef.current.innerHTML = INITIAL_EMPTY_PARAGRAPH;
        }
      } catch (err) {
        if (mounted && bodyEditorRef.current) {
          bodyEditorRef.current.innerHTML = INITIAL_EMPTY_PARAGRAPH;
        }
        if (mounted) {
          showToast(
            err instanceof Error
              ? err.message
              : 'تعذر تحميل بعض السجلات المحفوظة'
          );
        }
      }

      if (typeof window !== 'undefined') {
        const storedBackup = localStorage.getItem('notary_last_backup_at_v1');
        if (storedBackup) setLastBackupAt(storedBackup);
        const tourDone = localStorage.getItem('notary_editor_tour_v27');
        if (!tourDone) {
          setShowOnboardingTour(true);
        }
      }
    }
    initWorkspace();
    return () => {
      mounted = false;
    };
  }, [refreshActiveClausesInDOM, reloadAllOfficeRecords, setLastBackupAt, setLocalDirState, showToast]);

  // Single-click on a .smart-tag inside the A4 Editor -> Open InlineVariablePopover right above the tag
  const handleEditorClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    const smartTag = target.closest('.smart-tag, .smart-placeholder') as HTMLElement | null;
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
  };

  // Double-click on a .smart-tag inside the A4 Editor -> Open Parties & Variables Sidebar Tab
  const handleEditorDoubleClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    const smartTag = target.closest('.smart-tag, .smart-placeholder') as HTMLElement | null;
    if (smartTag) {
      e.preventDefault();
      setInlinePopover(null);
      setSidebarTab('parties');
      setIsSidebarOpen(true);
    }
  };

  // 1-Click Selection-to-Smart-Tag Converter
  const handleConvertSelectionToSmartTag = () => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();
    const createdVar = convertSelectionToSmartTag(
      bodyEditorRef.current,
      savedRangeRef.current
    );
    if (!createdVar) {
      showToast(
        'حدد أي كلمة أو جملة داخل ورقة العقد أولاً ثم اضغط «[ ] تحويل المحدد لوسم»'
      );
      return;
    }
    syncPlaceholdersAndDraft();
    showToast(`تم تحويل النص المحدد إلى وسم ذكي [${createdVar}] مرتبط بهذا العقد`);
  };

  // Keyboard shortcuts inside Editor
  const handleEditorKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) {
      e.preventDefault();
      handleUndo();
      return;
    }
    if (
      ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') ||
      ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'z')
    ) {
      e.preventDefault();
      handleRedo();
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
      document.execCommand('insertHTML', false, '&nbsp;&nbsp;&nbsp;&nbsp;');
      syncPlaceholdersAndDraft();
    }
  };

  // Word Paste Sanitizer
  const handleEditorPaste = (e: React.ClipboardEvent<HTMLDivElement>) => {
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
  };

  const handleEditorMouseDown = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    const clickedCell = target.closest('td, th') as HTMLTableCellElement | null;
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
  };

  // =========================================================
  // LIVE CONTRACT OUTLINE CLAUSES & LIBRARY HANDLERS
  // =========================================================
  const handleScrollToLiveClause = (domIndexOrId: number | string) => {
    if (!bodyEditorRef.current) return;
    scrollToContractClauseInDOM(bodyEditorRef.current, domIndexOrId);
  };

  const handleToggleLiveClauseEnabled = (clauseId: string) => {
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
      showToast('هذا البند مقفول لحمايته من التعطيل أو الحذف — افتح القفل أولاً');
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
  };

  const handleUpdateContractCondition = (
    key: string,
    value: string | boolean | null | undefined
  ) => {
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
        if (clause.condition && clause.condition.field === key && !clause.locked) {
          const condMatch =
            String(clause.condition.value).trim() === String(value).trim() ||
            (clause.condition.value === true && (value === true || value === 'نعم')) ||
            (clause.condition.value === false && (value === false || value === 'لا'));
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
  };

  const handleSetClauseCondition = (
    clauseId: string,
    condition?: ClauseCondition
  ) => {
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
          (condition.value === true && (currentVal === true || currentVal === 'نعم')) ||
          (condition.value === false && (currentVal === false || currentVal === 'لا'));
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
  };

  const handleToggleLiveClauseLocked = (clauseId: string) => {
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
  };

  const handleReorderLiveClauses = (
    draggedClauseId: string,
    targetClauseId: string
  ) => {
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
  };

  const handleMoveLiveClauseInDoc = (
    domIndex: number,
    direction: 'up' | 'down'
  ) => {
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
  };

  const handleDeleteLiveClauseFromDoc = (domIndexOrId: number | string) => {
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
  };

  const handleRenameLiveClauseInDoc = (
    domIndexOrId: number | string,
    newTitle: string
  ) => {
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
  };

  const handleToggleClauseInDoc = (clause: NotaryClause) => {
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
  };

  const handleInsertClauseAtCaret = (clause: NotaryClause) => {
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
  };

  const handleSaveNewClause = async (
    title: string,
    category: string,
    contentHtml?: string
  ) => {
    const bodyHtml =
      contentHtml ||
      bodyEditorRef.current?.innerHTML ||
      INITIAL_EMPTY_PARAGRAPH;
    await saveNewClauseToCatalog(title, category, bodyHtml);
  };

  const handleSaveSelectionAsClause = async (title: string, category: string) => {
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
  };

  const handleEditClauseInEditor = (clause: NotaryClause) => {
    if (!bodyEditorRef.current) return;
    if (!editingDerivedTpl && !editingClauseObj) {
      stashedMainDocRef.current = {
        ...getEditorZonesHtml(),
        title: docTitle,
      };
    }
    setEditingDerivedTpl(null);
    setEditingClauseObj(clause);
    bodyEditorRef.current.innerHTML = clause.contentHtml || INITIAL_EMPTY_PARAGRAPH;
    normalizeNotaryContainerDOM(bodyEditorRef.current);
    showToast(`أنت الآن تعدل صياغة البند "${clause.title}" داخل ورقة الـ A4`);
  };

  const handleFinishEditingClause = async (saveChanges: boolean) => {
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
        headerEditorRef.current.innerHTML = stashedMainDocRef.current.headerHtml;
      if (footerEditorRef.current)
        footerEditorRef.current.innerHTML = stashedMainDocRef.current.footerHtml;
      setDocTitle(stashedMainDocRef.current.title);
      stashedMainDocRef.current = null;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
    }
  };

  // =========================================================
  // DERIVED DOCUMENTS GENERATOR & IN-APP UNIFIED TEMPLATE EDITING
  // =========================================================
  const handleGenerateDerivedDoc = async (tpl: OfficeTemplate) => {
    await runGenerateDerivedDoc(tpl, getSessionSnapshot());
  };

  const handleEditDerivedTemplateInEditor = (tpl: OfficeTemplate) => {
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
    if (headerEditorRef.current) headerEditorRef.current.innerHTML = tpl.headerHtml || '';
    if (footerEditorRef.current) footerEditorRef.current.innerHTML = tpl.footerHtml || '';
    showToast(`أنت الآن تعدل قالب المكتب "${tpl.name}" داخل ورقة الـ A4`);
  };

  const handleFinishEditingDerivedTemplate = async (saveChanges: boolean) => {
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
      showToast(`تم حفظ تعديلات قالب المكتب "${updatedOfficeTpl.name}" بنجاح`);
    }
    setEditingDerivedTpl(null);
    if (stashedMainDocRef.current && bodyEditorRef.current) {
      bodyEditorRef.current.innerHTML = stashedMainDocRef.current.bodyHtml;
      if (headerEditorRef.current)
        headerEditorRef.current.innerHTML = stashedMainDocRef.current.headerHtml;
      if (footerEditorRef.current)
        footerEditorRef.current.innerHTML = stashedMainDocRef.current.footerHtml;
      setDocTitle(stashedMainDocRef.current.title);
      stashedMainDocRef.current = null;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
    }
  };

  // =========================================================
  // SUBDIVISION LOT SELECTION
  // =========================================================
  const handleSelectEstateAndLot = (estateId: string, lotNumber: string) => {
    setSelectedEstateId(estateId);
    setSelectedLotNumber(lotNumber);

    const estate = estates.find((e) => e.id === estateId);
    const lot = estate?.lots.find((l) => l.lotNumber === lotNumber);

    if (estate && lot) {
      const nextValues: Record<string, string> = {
        ...fieldValuesRef.current,
        رقم_الحصة: lot.lotNumber,
        طبيعة_الحصة: lot.nature,
        الطابق: `${lot.floor}${lot.building ? ` (${lot.building})` : ''}`,
        المساحة: lot.area,
        الأجزاء_المشتركة: lot.commonShares,
        تعيين_الحصة_الكامل: lot.fullDescription,
        مراجع_الوصف_التقسيمي: estate.subdivisionDeedRef,
      };
      fieldValuesRef.current = nextValues;
      setFieldValues(nextValues);
      syncPlaceholdersAndDraft(nextValues);
      showToast(`تم ربط بيانات الحصة رقم (${lot.lotNumber}) بالتعيين العقاري تلقائياً`);
    }
  };

  const handleInsertLotClauseAtCaret = (
    estate: SubdivisionEstate,
    lot: SubdivisionLot
  ) => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();
    const clauseHtml = buildSubdivisionLotClauseHtml(estate, lot);
    insertHtmlAtSelection(bodyEditorRef.current, clauseHtml, savedRangeRef.current);
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج فقرة تعيين الحصة رقم (${lot.lotNumber}) في العقد`);
  };

  const handleInsertFullSubdivisionTableAtCaret = (estate: SubdivisionEstate) => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();
    const tableHtml = buildFullSubdivisionTableHtml(estate);
    insertHtmlAtSelection(bodyEditorRef.current, tableHtml, savedRangeRef.current);
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج جدول الوصف التقسيمي (${estate.estateName}) في العقد`);
  };

  // =========================================================
  // DAILY EXPORT TO WORD (.docx) + LOG TO CONTRACT REVISION HISTORY
  // =========================================================
  const handleExportCurrentToWord = async () => {
    await handleExportContractWord(getSessionSnapshot());
  };

  // Word (.docx) Importer with Preview & Option to Replace or Insert at Caret
  const handleParseAndPreviewDocx = async (file: File) => {
    try {
      const imported = await importNotaryDocxFile(file);
      setPendingDocxImport(imported);
      setShowDocxPreviewModal(true);
    } catch (err) {
      showToast(
        err instanceof Error ? err.message : 'تعذر قراءة ملف الوورد المرفوع'
      );
    }
  };

  const handleApplyDocxImport = (
    mode: 'replace' | 'insert',
    customSegmentedHtml?: string
  ) => {
    if (!pendingDocxImport || !ensureBodyEditable()) return;
    const htmlToUse = customSegmentedHtml || pendingDocxImport.bodyHtml;

    if (mode === 'replace') {
      recordHistorySnapshot('قبل استبدال العقد بملف Word مستورد');
      setDocTitle(pendingDocxImport.title);
      outlineClausesRef.current = [];
      setOutlineClauses([]);
      if (bodyEditorRef.current) {
        bodyEditorRef.current.innerHTML = htmlToUse;
        normalizeNotaryContainerDOM(bodyEditorRef.current);
      }
      if (pendingDocxImport.headerHtml || pendingDocxImport.footerHtml) {
        setShowHeaderFooter(true);
        if (headerEditorRef.current)
          headerEditorRef.current.innerHTML = pendingDocxImport.headerHtml;
        if (footerEditorRef.current)
          footerEditorRef.current.innerHTML = pendingDocxImport.footerHtml;
      }
      setExtractedPlaceholders(pendingDocxImport.extractedPlaceholders);
      syncPlaceholdersAndDraft(undefined, undefined, []);
      showToast(
        `تم استبدال العقد بمحتوى ملف "${pendingDocxImport.title}" بنجاح`
      );
    } else {
      if (!bodyEditorRef.current) return;
      recordHistorySnapshot('قبل إدراج ملف Word عند المؤشر');
      insertHtmlAtSelection(
        bodyEditorRef.current,
        htmlToUse,
        savedRangeRef.current
      );
      normalizeNotaryContainerDOM(bodyEditorRef.current);
      syncPlaceholdersAndDraft();
      showToast(
        `تم إدراج محتوى ملف "${pendingDocxImport.title}" عند موضع المؤشر بنجاح`
      );
    }

    setPendingDocxImport(null);
  };

  const handleImportDocxDirectlyToEditor = async (file: File) => {
    await handleParseAndPreviewDocx(file);
  };

  const handleSaveCurrentAsTemplate = async (
    name: string,
    category: string,
    kind: OfficeTemplateKind = 'contract',
    code?: string,
    description?: string,
    clearFilledValues: boolean = true
  ) => {
    await handleSaveTemplateFromSnapshot({
      name,
      category,
      kind,
      code,
      description,
      clearFilledValues,
      snapshot: getSessionSnapshot(),
    });
  };

  const handleSaveCustomTemplateAdvanced = async (data: {
    name: string;
    kind?: OfficeTemplateKind;
    code?: string;
    category: string;
    description: string;
    clearFilledValues: boolean;
  }) => {
    await handleSaveCurrentAsTemplate(
      data.name,
      data.category,
      data.kind || 'contract',
      data.code,
      data.description,
      data.clearFilledValues
    );
  };

  const handleSaveCurrentDocumentAndRevision = async (
    summaryLabel?: string,
    explicitFolderId?: string | null
  ): Promise<SavedDocument> => {
    if (autoSaveTimerRef.current) {
      clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = null;
    }
    const liveOutline = extractLiveContractClausesFromDOM(bodyEditorRef.current);
    const savedDoc = await handleSaveDocumentAndRevision({
      snapshot: getSessionSnapshot(),
      liveOutline,
      folders,
      summaryLabel,
      explicitFolderId,
    });
    setAutoSaveState('saved');
    return savedDoc;
  };

  const handleTakeManualSnapshot = async (label: string) => {
    recordHistorySnapshot(label);
    await handleSaveCurrentDocumentAndRevision(label);
  };

  const handleRestoreSnapshot = async (rev: DocumentRevision) => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    if (!rev.bodyHtml || !rev.bodyHtml.trim()) {
      showToast('هذه المراجعة سجل حدث ولا تحتوي على نسخة متن نصي للاستعادة');
      return;
    }
    recordHistorySnapshot('قبل استعادة النسخة: ' + (rev.summary || rev.documentTitle));
    bodyEditorRef.current.innerHTML = rev.bodyHtml;
    normalizeNotaryContainerDOM(bodyEditorRef.current);
    const nextVals = rev.fieldValues || {};
    fieldValuesRef.current = nextVals;
    setFieldValues(nextVals);
    syncPlaceholdersAndDraft(nextVals);
    showToast(`تم استعادة النسخة: "${rev.summary || 'نسخة مرجعية'}"`);
  };

  const handleOpenSnapshotsHistoryModal = () => {
    setModalActiveBodyHtml(bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH);
    setShowSnapshotsHistoryModal(true);
  };

  const handleOpenDiffComparison = (rev: DocumentRevision) => {
    setModalActiveBodyHtml(bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH);
    setDiffComparisonRevision(rev);
    setShowVersionDiffModal(true);
  };

  // Saved Parties & Saved Properties Directory Handlers
  const handleRecallPartyToRole = (party: SavedPartyRecord, role: 'party1' | 'party2') => {
    const prefix = role === 'party1' ? 'الطرف_الأول' : 'الطرف_الثاني';
    const next: Record<string, string> = {
      ...fieldValuesRef.current,
      [`${prefix}_الاسم`]: party.fullName || '',
      [`${prefix}_تاريخ_الميلاد`]: party.birthDate || '',
      [`${prefix}_مكان_الميلاد`]: party.birthPlace || '',
      [`${prefix}_النسب`]: party.parentage || '',
      [`${prefix}_الهوية`]: party.idCardRef || '',
      [`${prefix}_الإقامة`]: party.address || '',
    };
    fieldValuesRef.current = next;
    setFieldValues(next);
    syncPlaceholdersAndDraft(next);
    showToast(
      `تم استدعاء بيانات "${party.fullName}" في حقول ${
        role === 'party1' ? 'الطرف الأول' : 'الطرف الثاني'
      }`
    );
  };

  const handleSaveCurrentPartyToDirectory = async (role: 'party1' | 'party2') => {
    await saveRolePartyToCatalog(role, fieldValuesRef.current);
  };

  // Structured Multi-Party Cards Handlers
  const handleUpdatePartyCards = useCallback(
    (nextCards: ContractPartyCard[]) => {
      partyCardsRef.current = nextCards;
      setPartyCards(nextCards);
      const nextValues = applyPartyCardsToFieldValues(
        nextCards,
        fieldValuesRef.current,
        extractedPlaceholders
      );
      fieldValuesRef.current = nextValues;
      setFieldValues(nextValues);
      syncPlaceholdersAndDraft(nextValues, undefined, undefined, nextCards);
    },
    [extractedPlaceholders, syncPlaceholdersAndDraft]
  );

  const handleRecallPropertyRecord = (prop: SavedPropertyRecord) => {
    const next: Record<string, string> = {
      ...fieldValuesRef.current,
      رقم_الحصة: prop.lotNumber || '',
      طبيعة_الحصة: prop.nature || '',
      الطابق: prop.floor || '',
      المساحة: prop.area || '',
      الأجزاء_المشتركة: prop.commonShares || '',
      تعيين_الحصة_الكامل: prop.fullDescription || '',
      مراجع_الوصف_التقسيمي: prop.deedRef || '',
    };
    fieldValuesRef.current = next;
    setFieldValues(next);
    syncPlaceholdersAndDraft(next);
    showToast(`تم استدعاء بيانات العقار "${prop.label}" في حقول التعيين`);
  };

  const handleSaveCurrentPropertyToDirectory = async () => {
    await savePropertyToCatalog(fieldValuesRef.current);
  };

  const handleChangeFieldInputType = (
    key: string,
    inputType: 'text' | 'number' | 'date'
  ) => {
    const nextTypes: Record<string, VariableInputType> = {
      ...fieldInputTypesRef.current,
      [key]: inputType,
    };
    fieldInputTypesRef.current = nextTypes;
    setFieldInputTypes(nextTypes);
    syncPlaceholdersAndDraft(fieldValuesRef.current, nextTypes);
  };

  // # New Clause Toolbar Action (Flexible # Heading without container or mandatory body)
  const handleInsertNewClauseHeadingAtCaret = async (customTitle?: string) => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();
    const { title } = insertOrWrapNewClauseAtSelection(
      bodyEditorRef.current,
      savedRangeRef.current,
      outlineClauses.length + 1,
      customTitle
    );
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج عنوان البند "# ${title}" — اكتب محتواه بحرية في الورقة`);
  };

  const handleInsertSmartTagAtCaret = (varKey: string) => {
    if (!bodyEditorRef.current || !varKey.trim() || !ensureBodyEditable()) return;
    const cleanKey = normalizePlaceholderKey(varKey).replace(/\s+/g, '_');
    if (!cleanKey) return;
    recordHistorySnapshot();
    const tagHtml = `<span class="smart-tag" data-var="${escapeHtml(
      cleanKey
    )}">[${escapeHtml(cleanKey)}]</span>&nbsp;`;
    insertHtmlAtSelection(bodyEditorRef.current, tagHtml, savedRangeRef.current);
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج الوسم [${cleanKey}] عند موضع المؤشر`);
  };

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
  }, [syncPlaceholdersAndDraft]);

  const handleBakeAllPlaceholdersIntoDocument = (
    explicitValues?: Record<string, string>
  ) => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();

    const activeValues: Record<string, string> = {
      ...fieldValues,
      ...fieldValuesRef.current,
      ...(explicitValues || {}),
    };
    fieldValuesRef.current = activeValues;
    setFieldValues(activeValues);

    const replaceLiveZoneTags = (zoneEl: HTMLElement | null) => {
      if (!zoneEl) return;
      decorateSmartTagsInDOM(zoneEl);
      const liveSpans = Array.from(
        zoneEl.querySelectorAll('.smart-tag, .smart-placeholder, [data-var]')
      ) as HTMLElement[];
      for (const span of liveSpans) {
        const rawKey =
          span.getAttribute('data-var') || (span.textContent || '');
        const val = lookupPlaceholderValue(activeValues, rawKey);
        if (val !== undefined && val.trim() !== '') {
          const textNode = document.createTextNode(val.trim());
          span.parentNode?.replaceChild(textNode, span);
        }
      }
      zoneEl.innerHTML = mergePlaceholdersIntoHtml(zoneEl.innerHTML, activeValues);
    };

    replaceLiveZoneTags(bodyEditorRef.current);
    normalizeNotaryContainerDOM(bodyEditorRef.current);

    if (headerEditorRef.current && headerEditorRef.current.innerHTML) {
      replaceLiveZoneTags(headerEditorRef.current);
    }
    if (footerEditorRef.current && footerEditorRef.current.innerHTML) {
      replaceLiveZoneTags(footerEditorRef.current);
    }

    syncPlaceholdersAndDraft(activeValues);
    showToast('تم دمج واستبدال قيم الحقول داخل نص العقد بنجاح');
  };

  // Compute unfilled variables count for the badge
  const unfilledCount = extractedPlaceholders.filter((k) => {
    const val = lookupPlaceholderValue(fieldValues, k);
    return !val || !val.trim();
  }).length;

  return (
    <div className="h-screen print:h-auto overflow-hidden print:overflow-visible flex flex-col bg-slate-100 text-slate-900">
      <EditorTopBar
        docTitle={docTitle}
        onChangeDocTitle={setDocTitle}
        autoSaveState={autoSaveState}
        previewMergedMode={previewMergedMode}
        revisionsCount={revisions.filter((r) => r.documentId === docId).length}
        lastBackupAt={lastBackupAt}
        localDirState={localDirState}
        onConnectLocalDirectory={() =>
          connectAndReloadLocalDirectory(async () => {
            await reloadLifecycleRecords();
          })
        }
        onOpenNewContractModal={() => setShowMultiSourceModal(true)}
        onImportWordFile={handleImportDocxDirectlyToEditor}
        onTogglePreviewMergedMode={() => {
          const next = !previewMergedMode;
          if (next) {
            const currentHtml =
              bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH;
            const activeVals = fieldValuesRef.current || fieldValues;
            const rawMerged = mergePlaceholdersIntoHtml(currentHtml, activeVals);
            setMergedPreviewHtml(
              injectNotarialPaginationIntoHtml(rawMerged, pageNumberingEnabled)
            );
            setPreviewMergedMode(true);
            setSidebarTab('parties');
            setIsSidebarOpen(true);
            showToast(
              'وضع معاينة الطباعة والدمج: عدّل أي قيمة في اليمين لترى العقد المدمج وعلامات الترقيم (../..) مباشرة'
            );
          } else {
            setPreviewMergedMode(false);
          }
        }}
        onOpenSnapshotsHistoryModal={handleOpenSnapshotsHistoryModal}
        onPrint={() => {
          const currentHtml =
            bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH;
          const activeVals = fieldValuesRef.current || fieldValues;
          const rawMerged = mergePlaceholdersIntoHtml(currentHtml, activeVals);
          setPrintPaginatedHtml(
            injectNotarialPaginationIntoHtml(rawMerged, pageNumberingEnabled)
          );
          setTimeout(() => window.print(), 60);
        }}
        onExportCurrentToWord={handleExportCurrentToWord}
        onExportBackupJson={async () => {
          const nowIso = await downloadOfficeBackupBundleFile();
          setLastBackupAt(nowIso);
          if (typeof window !== 'undefined') {
            localStorage.setItem('notary_last_backup_at_v1', nowIso);
          }
          showToast('تم تصدير النسخة الاحتياطية الشاملة للمكتب (JSON)');
        }}
        onImportBackupJson={async (file) => {
          try {
            const text = await file.text();
            const parsed = JSON.parse(text);
            const counts = await importFullBackupBundle(parsed);
            await reloadAllOfficeRecords();
            const nowIso = new Date().toISOString();
            setLastBackupAt(nowIso);
            if (typeof window !== 'undefined') {
              localStorage.setItem('notary_last_backup_at_v1', nowIso);
            }
            showToast(
              `تم استعادة (${counts.templatesCount}) قالب و (${counts.clausesCount}) بند و (${counts.documentsCount}) عقد`
            );
          } catch {
            showToast('تعذر قراءة ملف النسخة الاحتياطية');
          }
        }}
        onOpenOnboardingTour={() => setShowOnboardingTour(true)}
      />

      <EditorRibbon
        toolbarState={toolbarState}
        isBodyLocked={isBodyLocked}
        canUndo={undoStack.length > 0}
        canRedo={redoStack.length > 0}
        zoom={zoom}
        showHeaderFooter={showHeaderFooter}
        pageNumberingEnabled={pageNumberingEnabled}
        showFindReplace={showFindReplace}
        extractedPlaceholders={extractedPlaceholders}
        fieldValues={fieldValues}
        onConvertSelectionToSmartTag={handleConvertSelectionToSmartTag}
        onInsertSmartTagAtCaret={handleInsertSmartTagAtCaret}
        onInsertNewClauseHeadingAtCaret={handleInsertNewClauseHeadingAtCaret}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onToggleBold={() => applyInlineCommand('bold')}
        onToggleItalic={() => applyInlineCommand('italic')}
        onToggleUnderline={() => applyInlineCommand('underline')}
        onSetTextColor={(color) => applyInlineCommand('foreColor', color)}
        onSetHighlightColor={(color) => applyInlineCommand('hiliteColor', color)}
        onSetAlign={handleSetAlign}
        onSetDirection={handleSetDirection}
        onAdjustIndent={handleAdjustIndent}
        onToggleList={handleToggleList}
        onInsertTable={handleInsertTable}
        onTableAction={handleTableAction}
        onInsertPageBreak={handleInsertPageBreak}
        onInsertImageFile={handleInsertImageFile}
        onToggleHeaderFooter={() => setShowHeaderFooter((v) => !v)}
        onTogglePageNumbering={() => setPageNumberingEnabled((v) => !v)}
        onToggleFindReplace={() => setShowFindReplace((v) => !v)}
        onChangeZoom={setZoom}
        onSaveSelectionBookmark={saveSelectionBookmark}
      />

      <EditorModeAndSealBanners
        editingDerivedTpl={editingDerivedTpl}
        editingClauseObj={editingClauseObj}
        activeDocument={activeDocument}
        unlockIndexedBody={unlockIndexedBody}
        onToggleUnlockIndexedBody={() => setUnlockIndexedBody((v) => !v)}
        onFinishEditingDerivedTemplate={handleFinishEditingDerivedTemplate}
        onFinishEditingClause={handleFinishEditingClause}
      />

      {showFindReplace && (
        <EditorFindReplaceBar
          findQuery={findQuery}
          replaceQuery={replaceQuery}
          ignoreArabicHamzaAndDiacritics={ignoreArabicHamzaAndDiacritics}
          matchCount={matchCount}
          activeMatchIdx={activeMatchIdx}
          isBodyLocked={isBodyLocked}
          onChangeFindQuery={setFindQuery}
          onChangeReplaceQuery={setReplaceQuery}
          onToggleIgnoreHamza={setIgnoreArabicHamzaAndDiacritics}
          onPrevMatch={handlePrevMatch}
          onNextMatch={handleNextMatch}
          onRunReplace={handleRunReplace}
          onClose={() => setShowFindReplace(false)}
        />
      )}

      {toastMessage && (
        <div className="fixed bottom-12 left-6 z-50 bg-slate-900 text-white px-4 py-2.5 rounded-lg shadow-lg text-xs font-medium flex items-center gap-2 no-print">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}

      <div
        id="workspace"
        className="flex-1 flex flex-row min-h-0 overflow-hidden print:overflow-visible relative"
      >
        {isSidebarOpen ? (
          <div className="block h-full min-h-0 shrink-0 transition-all duration-150">
            <SidebarWorkspace
              onToggleCollapse={() => setIsSidebarOpen(false)}
              activeTab={sidebarTab}
              onSelectTab={(tab) => {
                setSidebarTab(tab);
                setIsSidebarOpen(true);
              }}
              liveContractClauses={outlineClauses}
              contractConditions={contractConditions}
              onUpdateContractCondition={handleUpdateContractCondition}
              onSetClauseCondition={handleSetClauseCondition}
              onScrollToLiveClause={handleScrollToLiveClause}
              onToggleLiveClauseEnabled={handleToggleLiveClauseEnabled}
              onToggleLiveClauseLocked={handleToggleLiveClauseLocked}
              onReorderLiveClauses={handleReorderLiveClauses}
              onMoveLiveClauseInDoc={handleMoveLiveClauseInDoc}
              onDeleteLiveClauseFromDoc={handleDeleteLiveClauseFromDoc}
              onRenameLiveClauseInDoc={handleRenameLiveClauseInDoc}
              onInsertNewClauseHeadingInDoc={handleInsertNewClauseHeadingAtCaret}
              onImportStandardClausesPack={handleImportStandardClausesPack}
              clauses={clauses}
              onInsertClauseAtCaret={handleInsertClauseAtCaret}
              onSaveNewClause={handleSaveNewClause}
              onSaveSelectionAsClause={handleSaveSelectionAsClause}
              onEditClauseInEditor={handleEditClauseInEditor}
              onDeleteClause={handleDeleteClause}
              partyCards={partyCards}
              onUpdatePartyCards={handleUpdatePartyCards}
              onSavePartyCardToDirectory={handleSavePartyCardToDirectory}
              extractedPlaceholders={extractedPlaceholders}
              unfilledCount={unfilledCount}
              fieldValues={fieldValues}
              fieldInputTypes={fieldInputTypes}
              onUpdateFieldValue={(key, value) => {
                const next = { ...fieldValuesRef.current, [key]: value };
                fieldValuesRef.current = next;
                setFieldValues(next);
                syncPlaceholdersAndDraft(next);
              }}
              onInsertPlaceholderAtCaret={(key) => {
                if (!bodyEditorRef.current || !ensureBodyEditable()) return;
                recordHistorySnapshot();
                const tagHtml = `<span class="smart-tag" data-var="${escapeHtml(
                  key
                )}">[${escapeHtml(key)}]</span>&nbsp;`;
                insertHtmlAtSelection(
                  bodyEditorRef.current,
                  tagHtml,
                  savedRangeRef.current
                );
                syncPlaceholdersAndDraft();
              }}
              onBakeAllPlaceholdersIntoDocument={handleBakeAllPlaceholdersIntoDocument}
              estates={estates}
              selectedEstateId={selectedEstateId}
              selectedLotNumber={selectedLotNumber}
              onSelectEstateAndLot={handleSelectEstateAndLot}
              onSaveEstate={async (estate) => {
                await handleSaveEstate(estate);
                setSelectedEstateId(estate.id);
              }}
              onDeleteEstate={handleDeleteEstate}
              onInsertLotClauseAtCaret={handleInsertLotClauseAtCaret}
              onInsertFullSubdivisionTableAtCaret={
                handleInsertFullSubdivisionTableAtCaret
              }
              savedParties={savedParties}
              savedProperties={savedProperties}
              onRecallPartyToRole={handleRecallPartyToRole}
              onSaveCurrentPartyToDirectory={handleSaveCurrentPartyToDirectory}
              onRecallPropertyRecord={handleRecallPropertyRecord}
              onSaveCurrentPropertyToDirectory={handleSaveCurrentPropertyToDirectory}
              templates={templates}
              onImportDocxAsTemplate={handleImportDocxAsTemplate}
              onSaveCurrentAsTemplate={handleSaveCurrentAsTemplate}
              onOpenSaveAsTemplateModal={() => setShowSaveAsTemplateModal(true)}
              onExportTemplatesJson={handleExportTemplatesJson}
              onImportTemplatesJson={handleImportTemplatesJson}
              onLoadTemplateFull={(tpl, clearPreviousClauses = true) => {
                if (!bodyEditorRef.current || !ensureBodyEditable()) return;
                recordHistorySnapshot();
                setDocTitle(tpl.name);
                let nextHtml = tpl.bodyHtml;
                if (clearPreviousClauses) {
                  const temp = document.createElement('div');
                  temp.innerHTML = nextHtml;
                  const clauseContainers = Array.from(
                    temp.querySelectorAll('.clause-container, [data-clause-id]')
                  );
                  for (const c of clauseContainers) {
                    c.remove();
                  }
                  nextHtml = temp.innerHTML || INITIAL_EMPTY_PARAGRAPH;
                  setActiveClauseIdsInDoc([]);
                }
                bodyEditorRef.current.innerHTML = nextHtml;
                normalizeNotaryContainerDOM(bodyEditorRef.current);
                if (tpl.headerHtml || tpl.footerHtml) {
                  setShowHeaderFooter(true);
                  if (headerEditorRef.current)
                    headerEditorRef.current.innerHTML = tpl.headerHtml;
                  if (footerEditorRef.current)
                    footerEditorRef.current.innerHTML = tpl.footerHtml;
                }
                const nextVals = { ...(tpl.defaultFieldValues || {}) };
                fieldValuesRef.current = nextVals;
                setFieldValues(nextVals);
                syncPlaceholdersAndDraft(nextVals);
                showToast(
                  `تم استبدال المحتوى بالقالب "${tpl.name}" وضبط متغيراته المعزولة`
                );
              }}
              onInsertTemplateAtCaret={(tpl) => {
                if (!bodyEditorRef.current || !ensureBodyEditable()) return;
                recordHistorySnapshot();
                insertHtmlAtSelection(
                  bodyEditorRef.current,
                  tpl.bodyHtml,
                  savedRangeRef.current
                );
                syncPlaceholdersAndDraft();
                showToast(`تم إدراج محتوى القالب "${tpl.name}" عند موضع المؤشر`);
              }}
              onExportTemplateDocx={async (tpl) => {
                await downloadNotaryDocx({
                  title: tpl.name,
                  bodyHtml: tpl.bodyHtml,
                  headerHtml: tpl.headerHtml,
                  footerHtml: tpl.footerHtml,
                  pageNumberingEnabled: tpl.pageNumberingEnabled,
                  fieldValues,
                });
                showToast(`تم تصدير القالب "${tpl.name}" بصيغة Word (.docx)`);
              }}
              onDeleteTemplate={handleDeleteTemplate}
              activeEditingDerivedId={editingDerivedTpl?.id || null}
              onGenerateDerivedDoc={handleGenerateDerivedDoc}
              onEditDerivedTemplateInEditor={handleEditDerivedTemplateInEditor}
              folders={folders}
              clerks={clerks}
              activeClerk={activeClerk}
              activeDocument={activeDocument}
              activeDerivedDocId={activeDerivedDoc?.id || null}
              onCreateFolder={handleCreateFolder}
              onRenameFolder={handleRenameFolder}
              onDeleteFolder={handleDeleteFolder}
              onMoveDocument={handleMoveDocument}
              onMoveFolder={handleMoveFolder}
              onCreateContractInFolder={handleCreateContractInFolder}
              onUpdateDocumentMeta={handleUpdateDocumentMeta}
              onOpenDerivedModal={() => setShowMultiSourceModal(true)}
              onSelectDerivedDoc={handleSelectDerivedDoc}
              documents={documents}
              activeDocumentId={docId}
              currentEditorTitle={docTitle}
              onOpenDocument={handleSelectDocument}
              onSaveCurrentDocument={(targetFolderId?: string | null) =>
                handleSaveCurrentDocumentAndRevision(undefined, targetFolderId)
              }
              onDeleteDocument={handleDeleteDocument}
              onExportDocumentDocx={async (doc) => {
                await downloadNotaryDocx({
                  title: doc.title,
                  bodyHtml: doc.bodyHtml,
                  headerHtml: doc.showHeaderFooter ? doc.headerHtml : '',
                  footerHtml: doc.showHeaderFooter ? doc.footerHtml : '',
                  pageNumberingEnabled: doc.pageNumberingEnabled,
                  fieldValues: doc.fieldValues,
                });
              }}
              onSaveSelectionBookmark={saveSelectionBookmark}
              onContractIndexedOrSynced={async (updatedDoc) => {
                const [freshDocs, freshRevs] = await Promise.all([
                  loadSavedDocuments(),
                  loadDocumentRevisions(),
                ]);
                setDocuments(freshDocs);
                setRevisions(freshRevs);
                const target =
                  updatedDoc || freshDocs.find((d) => d.id === docId) || null;
                if (target && target.id === docId) {
                  setActiveDocument(target);
                  activeDocumentRef.current = target;
                  setDocTitle(target.title);
                  const mergedVals = {
                    ...fieldValuesRef.current,
                    ...(target.fieldValues || {}),
                  };
                  fieldValuesRef.current = mergedVals;
                  setFieldValues(mergedVals);
                  syncSmartTagsFilledStateInDOM(bodyEditorRef.current, mergedVals);
                  if (previewMergedMode && bodyEditorRef.current) {
                    const rawMerged = mergePlaceholdersIntoHtml(
                      bodyEditorRef.current.innerHTML,
                      mergedVals
                    );
                    setMergedPreviewHtml(
                      injectNotarialPaginationIntoHtml(
                        rawMerged,
                        pageNumberingEnabled
                      )
                    );
                  }
                }
              }}
              showToast={showToast}
            />
          </div>
        ) : (
          <EditorActivityRail
            unfilledCount={unfilledCount}
            onSelectTab={(tab) => {
              setSidebarTab(tab);
              setIsSidebarOpen(true);
            }}
          />
        )}

        <EditorA4Canvas
          zoom={zoom}
          showHeaderFooter={showHeaderFooter}
          pageNumberingEnabled={pageNumberingEnabled}
          previewMergedMode={previewMergedMode}
          mergedPreviewHtml={mergedPreviewHtml}
          printPaginatedHtml={printPaginatedHtml}
          isBodyLocked={isBodyLocked}
          estimatedPages={docMetrics.estimatedPages}
          headerEditorRef={headerEditorRef}
          bodyEditorRef={bodyEditorRef}
          footerEditorRef={footerEditorRef}
          onSyncPlaceholdersAndDraft={() => syncPlaceholdersAndDraft()}
          onDecorateEditorZonesPreservingSelection={
            decorateEditorZonesPreservingSelection
          }
          onTogglePreviewMergedMode={() => {
            const next = !previewMergedMode;
            if (next) {
              const currentHtml =
                bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH;
              const activeVals = fieldValuesRef.current || fieldValues;
              const rawMerged = mergePlaceholdersIntoHtml(currentHtml, activeVals);
              setMergedPreviewHtml(
                injectNotarialPaginationIntoHtml(
                  rawMerged,
                  pageNumberingEnabled,
                  true
                )
              );
              setPreviewMergedMode(true);
            } else {
              setPreviewMergedMode(false);
            }
          }}
          onEditorKeyDown={handleEditorKeyDown}
          onEditorPaste={handleEditorPaste}
          onEditorMouseDown={handleEditorMouseDown}
          onEditorClick={handleEditorClick}
          onEditorDoubleClick={handleEditorDoubleClick}
        />
      </div>

      <EditorBottomStatusBar
        autoSaveState={autoSaveState}
        unfilledCount={unfilledCount}
        totalVariablesCount={extractedPlaceholders.length}
        docMetrics={docMetrics}
        onOpenSmartVariablesModal={() => {
          setSidebarTab('parties');
          setIsSidebarOpen(true);
        }}
      />

      <EditorModalsHost
        docId={docId}
        docTitle={docTitle}
        inlinePopover={inlinePopover}
        onCloseInlinePopover={() => setInlinePopover(null)}
        onOpenSmartVarsModalForKey={() => {
          decorateEditorZonesPreservingSelection();
          setSidebarTab('parties');
          setIsSidebarOpen(true);
        }}
        extractedPlaceholders={extractedPlaceholders}
        fieldValues={fieldValues}
        fieldInputTypes={fieldInputTypes}
        onUpdateFieldValue={(key, value) => {
          const next = { ...fieldValuesRef.current, [key]: value };
          fieldValuesRef.current = next;
          setFieldValues(next);
          syncPlaceholdersAndDraft(next);
        }}
        onChangeFieldInputType={handleChangeFieldInputType}
        showVersionDiffModal={showVersionDiffModal}
        onCloseVersionDiffModal={() => {
          setShowVersionDiffModal(false);
          setDiffComparisonRevision(null);
        }}
        revisions={revisions}
        diffComparisonRevision={diffComparisonRevision}
        modalActiveBodyHtml={modalActiveBodyHtml}
        onRestoreSnapshot={handleRestoreSnapshot}
        showDocxPreviewModal={showDocxPreviewModal}
        onCloseDocxPreviewModal={() => {
          setShowDocxPreviewModal(false);
          setPendingDocxImport(null);
        }}
        pendingDocxImport={pendingDocxImport}
        onConfirmApplyDocxImport={handleApplyDocxImport}
        showSaveAsTemplateModal={showSaveAsTemplateModal}
        onCloseSaveAsTemplateModal={() => setShowSaveAsTemplateModal(false)}
        onConfirmSaveCustomTemplate={handleSaveCustomTemplateAdvanced}
        showSnapshotsHistoryModal={showSnapshotsHistoryModal}
        onCloseSnapshotsHistoryModal={() => setShowSnapshotsHistoryModal(false)}
        onTakeManualSnapshot={handleTakeManualSnapshot}
        onDeleteSnapshot={handleDeleteSnapshot}
        onOpenDiffComparison={handleOpenDiffComparison}
        showOnboardingTour={showOnboardingTour}
        onCloseOnboardingTour={() => setShowOnboardingTour(false)}
        onFinishOnboardingTour={() => {
          if (typeof window !== 'undefined') {
            localStorage.setItem('notary_editor_tour_v27', 'true');
          }
          setShowOnboardingTour(false);
        }}
        showMultiSourceModal={showMultiSourceModal}
        onCloseMultiSourceModal={() => setShowMultiSourceModal(false)}
        templates={templates}
        documents={documents}
        onStartBlank={(clearPreviousClauses = true) => {
          if (!bodyEditorRef.current) return;
          setUndoStack([]);
          setRedoStack([]);
          const newId = `doc_${Date.now()}`;
          setDocId(newId);
          setDocTitle('عقد توثيقي جديد');
          setActiveDocument(null);
          activeDocumentRef.current = null;
          fieldValuesRef.current = {};
          fieldInputTypesRef.current = {};
          setFieldValues({});
          setFieldInputTypes({});
          setSelectedEstateId('');
          setSelectedLotNumber('');
          bodyEditorRef.current.innerHTML = INITIAL_EMPTY_PARAGRAPH;
          if (headerEditorRef.current) headerEditorRef.current.innerHTML = '';
          if (footerEditorRef.current) footerEditorRef.current.innerHTML = '';
          if (clearPreviousClauses) {
            setActiveClauseIdsInDoc([]);
          }
          syncPlaceholdersAndDraft({}, {});
          showToast('تم فتح ورقة عقد توثيقي جديد معزول تماماً');
        }}
        onSelectTemplateFromMultiSource={(tpl: OfficeTemplate, mode, clearPreviousClauses) => {
          if (!bodyEditorRef.current) return;
          if (mode === 'insert') {
            if (!ensureBodyEditable()) return;
            recordHistorySnapshot(`قبل إدراج القالب "${tpl.name}" عند المؤشر`);
            insertHtmlAtSelection(
              bodyEditorRef.current,
              tpl.bodyHtml,
              savedRangeRef.current
            );
            syncPlaceholdersAndDraft();
            showToast(`تم إدراج القالب "${tpl.name}" عند موضع المؤشر`);
            return;
          }
          setUndoStack([]);
          setRedoStack([]);
          setDocId(`doc_${Date.now()}`);
          setDocTitle(tpl.name);
          setActiveDocument(null);
          activeDocumentRef.current = null;
          let nextHtml = tpl.bodyHtml;
          if (clearPreviousClauses) {
            const temp = document.createElement('div');
            temp.innerHTML = nextHtml;
            const clauseContainers = Array.from(
              temp.querySelectorAll('.clause-container, [data-clause-id]')
            );
            for (const c of clauseContainers) {
              c.remove();
            }
            nextHtml = temp.innerHTML || INITIAL_EMPTY_PARAGRAPH;
            setActiveClauseIdsInDoc([]);
          }
          bodyEditorRef.current.innerHTML = nextHtml;
          normalizeNotaryContainerDOM(bodyEditorRef.current);
          const nextVals = { ...(tpl.defaultFieldValues || {}) };
          fieldValuesRef.current = nextVals;
          fieldInputTypesRef.current = {};
          setFieldValues(nextVals);
          setFieldInputTypes({});
          setSelectedEstateId('');
          setSelectedLotNumber('');
          syncPlaceholdersAndDraft(nextVals, {});
          showToast(
            `تم فتح عقد جديد من القالب "${tpl.name}" بمتغيرات وبنود معزولة`
          );
        }}
        onSelectSavedDocFromMultiSource={(doc) => {
          handleSelectDocument(doc);
          showToast(`تم فتح العقد "${doc.title}" بجميع خصائصه ومتغيراته المحفوظة`);
        }}
        onImportDocxFileFromMultiSource={handleImportDocxDirectlyToEditor}
      />
    </div>
  );
}
