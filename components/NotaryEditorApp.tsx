'use client';

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { CheckCircle2 } from 'lucide-react';
import EditorRibbon from './EditorRibbon';
import EditorTopBar from './EditorTopBar';
import EditorA4Canvas from './editor/EditorA4Canvas';
import EditorQuickAssistPopover from './editor/EditorQuickAssistPopover';
import {
  EditorActivityRail,
  EditorModalsHost,
} from './editor/EditorModalsHost';
import OfficeSettingsModal from './modals/OfficeSettingsModal';
import {
  EditorBottomStatusBar,
  EditorFindReplaceBar,
  EditorModeAndSealBanners,
} from './editor/EditorStatusBars';
import { InlineVariablePopoverState } from './InlineVariablePopover';
import SidebarWorkspace, { SidebarTab } from './SidebarWorkspace';
import WordTemplatesManagerModal from './modals/WordTemplatesManagerModal';
import { DocumentRevision, OfficeTemplate } from '@/lib/types';
import {
  getRuntimeInfo,
  importFullBackupBundle,
  loadDocumentRevisions,
  saveDocumentRevision,
} from '@/lib/storage';
import {
  exportContractWithStandardTemplate,
  StandardNotaryTemplateId,
} from '@/lib/notary-word-templates';
import {
  computeDocumentMetrics,
  injectNotarialPaginationIntoHtml,
} from '@/lib/notarial-pagination';
import { extractLiveContractClausesFromDOM } from '@/lib/clause-outline-engine';
import {
  escapeHtml,
  extractPlaceholdersFromHtml,
  extractPlaceholdersGroupedByClause,
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
  syncSmartTagsFilledStateInDOM,
} from '@/lib/editor-utils';
import {
  DocxImportResult,
  downloadNotaryDocx,
  mergePlaceholdersIntoHtml,
} from '@/lib/docx-engine';
import { downloadOfficeBackupBundleFile } from '@/lib/document-io-workflows';
import { tryRestorePersistedOfficeDirectory } from '@/lib/desktop-bridge';
import {
  INITIAL_EMPTY_PARAGRAPH,
  useDocumentSession,
} from '@/hooks/useDocumentSession';
import { useOfficeCatalog } from '@/hooks/useOfficeCatalog';
import { useNotaryLifecycle } from '@/hooks/useNotaryLifecycle';
import { useEditorSync } from '@/hooks/useEditorSync';
import { useEditorDomCommands } from '@/hooks/useEditorDomCommands';
import { useFindAndReplace } from '@/hooks/useFindAndReplace';
import { useEditorClauses } from '@/hooks/useEditorClauses';
import { useEditorParties } from '@/hooks/useEditorParties';
import { useEditorSmartAssist } from '@/hooks/useEditorSmartAssist';
import { useEditorDocumentActions } from '@/hooks/useEditorDocumentActions';

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

  const ensureBodyEditable = useCallback(() => {
    if (isBodyLocked) {
      showToast(
        'متن هذا العقد مقفول رسمياً بعد الفهرسة والختم — اضغط «فتح قفل المتن استثنائياً» في الشريط العلوي إن لزم التعديل'
      );
      return false;
    }
    return true;
  }, [isBodyLocked, showToast]);

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
  const [editorInteractionMode, setEditorInteractionMode] = useState<
    'desktop-pro' | 'adaptive-touch'
  >('desktop-pro');

  useEffect(() => {
    const handleResize = () => {
      if (window.innerWidth < 1024) {
        setEditorInteractionMode('adaptive-touch');
      } else {
        setEditorInteractionMode('desktop-pro');
      }
    };
    handleResize();
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);
  const [inlinePopover, setInlinePopover] =
    useState<InlineVariablePopoverState | null>(null);

  // Modals state
  const [showVersionDiffModal, setShowVersionDiffModal] =
    useState<boolean>(false);
  const [diffComparisonRevision, setDiffComparisonRevision] =
    useState<DocumentRevision | null>(null);
  const [showMultiSourceModal, setShowMultiSourceModal] =
    useState<boolean>(false);
  const [showDocxPreviewModal, setShowDocxPreviewModal] =
    useState<boolean>(false);
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
  const [showOfficeSettingsModal, setShowOfficeSettingsModal] = useState<boolean>(false);
  const [showWordTemplatesManager, setShowWordTemplatesManager] = useState<boolean>(false);
  const previewBackupHtmlRef = useRef<string>('');

  // 4. Editor Sync & Auto-Save Hook
  const {
    debouncedSyncTimerRef,
    clearAutoSaveTimer,
    refreshActiveClausesInDOM,
    triggerFastClauseScan,
    syncPlaceholdersAndDraft,
    flushSyncPlaceholdersAndDraft,
    decorateEditorZonesPreservingSelection,
  } = useEditorSync({
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
  });

  // 5. Editor DOM Formatting & Undo/Redo Hook
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

  // 6. Arabic Find & Replace Hook
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

  // 7. Live Contract Clauses & Library Hook
  const {
    handleScrollToLiveClause,
    handleToggleLiveClauseEnabled,
    handleUpdateContractCondition,
    handleSetClauseCondition,
    handleToggleLiveClauseLocked,
    handleReorderLiveClauses,
    handleMoveLiveClauseInDoc,
    handleDeleteLiveClauseFromDoc,
    handleRenameLiveClauseInDoc,
    handleInsertClauseAtCaret,
    handleInsertNewClauseHeadingAtCaret,
    handleSaveNewClause,
    handleSaveSelectionAsClause,
    handleEditClauseInEditor,
    handleFinishEditingClause,
    handleEditDerivedTemplateInEditor,
    handleFinishEditingDerivedTemplate,
  } = useEditorClauses({
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
  });

  // 8. Parties, Properties, Subdivision Lots & Smart Tags Hook
  const {
    unfilledCount,
    handleUpdateFieldValue,
    handleChangeFieldInputType,
    handleUpdatePartyCards,
    handleRecallPartyToRole,
    handleSaveCurrentPartyToDirectory,
    handleRecallPropertyRecord,
    handleSaveCurrentPropertyToDirectory,
    handleSelectEstateAndLot,
    handleInsertLotClauseAtCaret,
    handleInsertFullSubdivisionTableAtCaret,
    handleConvertSelectionToSmartTag,
    handleInsertSmartTagAtCaret,
    handleBakeAllPlaceholdersIntoDocument,
  } = useEditorParties({
    bodyEditorRef,
    headerEditorRef,
    footerEditorRef,
    savedRangeRef,
    estates,
    extractedPlaceholders,
    partyCardsRef,
    setPartyCards,
    fieldValues,
    fieldValuesRef,
    setFieldValues,
    fieldInputTypesRef,
    setFieldInputTypes,
    setSelectedEstateId,
    setSelectedLotNumber,
    ensureBodyEditable,
    recordHistorySnapshot,
    syncPlaceholdersAndDraft,
    saveRolePartyToCatalog,
    savePropertyToCatalog,
    showToast,
  });

  // 9. Smart Assist ([ # /) & Canvas Events Hook
  const {
    smartAssistPrefs,
    setSmartAssistPrefs,
    quickAssistTrigger,
    setQuickAssistTrigger,
    quickAssistSelectedIdx,
    quickAssistOptions,
    handleUpdateSmartAssistPrefs,
    scheduleSyncPlaceholdersAndDraft,
    handleSelectQuickAssistOption,
    handleEditorBeforeInput,
    handleEditorKeyDown,
    handleEditorPaste,
    handleEditorMouseDown,
    handleEditorClick,
    handleEditorDoubleClick,
  } = useEditorSmartAssist({
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
    syncPlaceholdersAndDraft: () => syncPlaceholdersAndDraft(),
    flushSyncPlaceholdersAndDraft,
    triggerFastClauseScan,
    setShowFindReplace,
    setInlinePopover,
    setSidebarTab,
    setIsSidebarOpen,
    showToast,
  });

  // 10. Document Switching, Word Import/Export, Snapshots & Template Workflows Hook
  const {
    handleSelectDocument,
    handleCreateContractInFolder,
    handleSelectDerivedDoc,
    handleParseAndPreviewDocx,
    handleApplyDocxImport,
    handleSaveCurrentAsTemplate,
    handleSaveCustomTemplateAdvanced,
    handleSaveCurrentDocumentAndRevision,
    handleTakeManualSnapshot,
    handleRestoreSnapshot,
    handleLoadTemplateFull,
    handleInsertTemplateAtCaret,
    handleStartBlank,
    handleSelectTemplateFromMultiSource,
    handleContractIndexedOrSynced,
  } = useEditorDocumentActions({
    docId,
    setDocId,
    docTitle,
    setDocTitle,
    showHeaderFooter,
    setShowHeaderFooter,
    pageNumberingEnabled,
    setPageNumberingEnabled,
    previewMergedMode,
    setMergedPreviewHtml,
    bodyEditorRef,
    headerEditorRef,
    footerEditorRef,
    savedRangeRef,
    setUndoStack,
    setRedoStack,
    fieldValuesRef,
    setFieldValues,
    fieldInputTypesRef,
    setFieldInputTypes,
    outlineClausesRef,
    setOutlineClauses,
    contractConditionsRef,
    setContractConditions,
    partyCardsRef,
    setPartyCards,
    setActiveClauseIdsInDoc,
    setExtractedPlaceholders,
    setSelectedEstateId,
    setSelectedLotNumber,
    setInlinePopover,
    activeClerk,
    folders,
    setDocuments,
    setRevisions,
    setActiveDocument,
    activeDocumentRef,
    setActiveDerivedDoc,
    setAutoSaveState,
    pendingDocxImport,
    setPendingDocxImport,
    setShowDocxPreviewModal,
    clearAutoSaveTimer,
    ensureBodyEditable,
    recordHistorySnapshot,
    getSessionSnapshot,
    syncPlaceholdersAndDraft,
    handleSaveTemplateFromSnapshot,
    handleSaveDocumentAndRevision,
    showToast,
  });

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
        const { loadedDocs, loadedEstates, activeDraft } =
          await reloadAllOfficeRecords();
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

          bodyEditorRef.current.innerHTML =
            activeDraft.bodyHtml || INITIAL_EMPTY_PARAGRAPH;
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
        const storedAssist = localStorage.getItem(
          'notary_smart_assist_prefs_v1'
        );
        if (storedAssist) {
          try {
            const parsed = JSON.parse(storedAssist);
            setSmartAssistPrefs({
              enableInlineAutocomplete:
                parsed.enableInlineAutocomplete !== false,
              enableSlashCommands: parsed.enableSlashCommands !== false,
            });
          } catch {
            // Ignore malformed storage
          }
        }
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
  }, [
    activeDocumentRef,
    bodyEditorRef,
    contractConditionsRef,
    fieldInputTypesRef,
    fieldValuesRef,
    footerEditorRef,
    headerEditorRef,
    outlineClausesRef,
    partyCardsRef,
    refreshActiveClausesInDOM,
    reloadAllOfficeRecords,
    setActiveDocument,
    setAutoSaveState,
    setClauseGroups,
    setContractConditions,
    setDocId,
    setDocMetrics,
    setDocTitle,
    setExtractedPlaceholders,
    setFieldInputTypes,
    setFieldValues,
    setLastBackupAt,
    setLocalDirState,
    setOutlineClauses,
    setPageNumberingEnabled,
    setPartyCards,
    setSelectedEstateId,
    setSelectedLotNumber,
    setSmartAssistPrefs,
    showToast,
  ]);

  const handleTogglePreviewMergedMode = useCallback(() => {
    flushSyncPlaceholdersAndDraft();
    const next = !previewMergedMode;
    if (next) {
      const currentHtml =
        bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH;
      previewBackupHtmlRef.current = currentHtml;
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
      setSidebarTab('parties');
      setIsSidebarOpen(true);
      showToast(
        'وضع معاينة الطباعة والدمج: عدّل أي قيمة في اليمين لترى العقد المدمج وعلامات الترقيم (../..) مباشرة'
      );
    } else {
      setPreviewMergedMode(false);
      if (
        bodyEditorRef.current &&
        previewBackupHtmlRef.current &&
        (bodyEditorRef.current.textContent || '').trim().length <
          (previewBackupHtmlRef.current.replace(/<[^>]+>/g, '') || '').trim().length / 2
      ) {
        bodyEditorRef.current.innerHTML = previewBackupHtmlRef.current;
        normalizeNotaryContainerDOM(bodyEditorRef.current, false);
      }
      setTimeout(() => {
        flushSyncPlaceholdersAndDraft();
      }, 50);
    }
  }, [
    bodyEditorRef,
    fieldValues,
    fieldValuesRef,
    flushSyncPlaceholdersAndDraft,
    pageNumberingEnabled,
    previewMergedMode,
    setIsSidebarOpen,
    setMergedPreviewHtml,
    setPreviewMergedMode,
    setSidebarTab,
    showToast,
  ]);

  const handleExportWithStandardTemplate = useCallback(
    async (templateId: StandardNotaryTemplateId) => {
      try {
        const snapshot = getSessionSnapshot();
        const { fileName, templateName } =
          await exportContractWithStandardTemplate({
            templateId,
            contractTitle: snapshot.docTitle,
            bodyHtml: snapshot.bodyHtml,
            fieldValues: snapshot.fieldValues,
          });

        const exportRevision: DocumentRevision = {
          id: `rev_exp_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
          documentId: snapshot.docId,
          documentTitle: snapshot.docTitle || 'عقد_توثيقي',
          author: activeClerk.name || 'الموثق الرئيسي',
          fieldValues: { ...snapshot.fieldValues },
          createdAt: new Date().toISOString(),
          eventType: 'export',
          fileName,
          docTypeCode: templateId,
          summary: `تصدير بقالب (${templateName}): ${fileName}`,
        };
        await saveDocumentRevision(exportRevision);
        const nextRevs = await loadDocumentRevisions();
        setRevisions(nextRevs);

        showToast(
          `تم تصدير «${templateName}» بنجاح بصيغة Word (${fileName}) وحقن بيانات العقد`
        );
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'حدث خطأ أثناء تصدير القالب'
        );
      }
    },
    [activeClerk.name, getSessionSnapshot, setRevisions, showToast]
  );

  return (
    <div className="h-screen print:h-auto overflow-hidden print:overflow-visible flex flex-col bg-slate-100 text-slate-900">
      <EditorTopBar
        docTitle={docTitle}
        onChangeDocTitle={setDocTitle}
        autoSaveState={autoSaveState}
        previewMergedMode={previewMergedMode}
        revisionsCount={revisions.filter((r) => r.documentId === docId).length}
        localDirState={localDirState}
        onOpenNewContractModal={() => setShowMultiSourceModal(true)}
        onImportWordFile={handleParseAndPreviewDocx}
        onOpenSnapshotsHistoryModal={() => {
          setModalActiveBodyHtml(
            bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH
          );
          setShowSnapshotsHistoryModal(true);
        }}
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
        onExportCurrentToWord={() =>
          handleExportContractWord(getSessionSnapshot())
        }
        onExportWithTemplate={handleExportWithStandardTemplate}
        onOpenTemplatesManager={() => setShowWordTemplatesManager(true)}
        onOpenOfficeSettings={() => setShowOfficeSettingsModal(true)}
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
        onSetHighlightColor={(color) =>
          applyInlineCommand('hiliteColor', color)
        }
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
        smartAssistPrefs={smartAssistPrefs}
        onChangeSmartAssistPrefs={handleUpdateSmartAssistPrefs}
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
              onInsertNewClauseHeadingInDoc={
                handleInsertNewClauseHeadingAtCaret
              }
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
              onUpdateFieldValue={handleUpdateFieldValue}
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
              onBakeAllPlaceholdersIntoDocument={
                handleBakeAllPlaceholdersIntoDocument
              }
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
              onSaveCurrentPropertyToDirectory={
                handleSaveCurrentPropertyToDirectory
              }
              templates={templates}
              onImportDocxAsTemplate={handleImportDocxAsTemplate}
              onSaveCurrentAsTemplate={handleSaveCurrentAsTemplate}
              onOpenSaveAsTemplateModal={() => setShowSaveAsTemplateModal(true)}
              onExportTemplatesJson={handleExportTemplatesJson}
              onImportTemplatesJson={handleImportTemplatesJson}
              onLoadTemplateFull={handleLoadTemplateFull}
              onInsertTemplateAtCaret={handleInsertTemplateAtCaret}
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
              onGenerateDerivedDoc={(tpl: OfficeTemplate) =>
                runGenerateDerivedDoc(tpl, getSessionSnapshot())
              }
              onEditDerivedTemplateInEditor={handleEditDerivedTemplateInEditor}
              onOpenWordTemplatesManager={() =>
                setShowWordTemplatesManager(true)
              }
              folders={folders}
              clerks={clerks}
              activeClerk={activeClerk}
              activeDocument={activeDocument}
              activeDerivedDocId={activeDerivedDoc?.id || null}
              onCreateFolder={handleCreateFolder}
              onRenameFolder={handleRenameFolder}
              onDeleteFolder={(folderId) =>
                handleDeleteFolderAndMigrateDocs(folderId, folders, setFolders)
              }
              onMoveDocument={(targetDocId, targetFolderId) =>
                moveDocumentInLifecycle(targetDocId, targetFolderId, folders)
              }
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
              onContractIndexedOrSynced={handleContractIndexedOrSynced}
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
          onSyncPlaceholdersAndDraft={scheduleSyncPlaceholdersAndDraft}
          onFlushSyncPlaceholdersAndDraft={flushSyncPlaceholdersAndDraft}
          onDecorateEditorZonesPreservingSelection={
            decorateEditorZonesPreservingSelection
          }
          onTogglePreviewMergedMode={handleTogglePreviewMergedMode}
          onEditorBeforeInput={handleEditorBeforeInput}
          onEditorKeyDown={handleEditorKeyDown}
          onEditorPaste={handleEditorPaste}
          onEditorMouseDown={handleEditorMouseDown}
          onEditorClick={handleEditorClick}
          onEditorDoubleClick={handleEditorDoubleClick}
        />

        {quickAssistTrigger && (
          <EditorQuickAssistPopover
            open={Boolean(quickAssistTrigger)}
            triggerType={quickAssistTrigger.type}
            query={quickAssistTrigger.query}
            position={quickAssistTrigger.rect}
            options={quickAssistOptions}
            selectedIndex={quickAssistSelectedIdx}
            onSelectOption={handleSelectQuickAssistOption}
            onClose={() => setQuickAssistTrigger(null)}
            onDisableAssistType={(type) => {
              if (type === 'slash') {
                handleUpdateSmartAssistPrefs({
                  ...smartAssistPrefs,
                  enableSlashCommands: false,
                });
                showToast(
                  'تم إيقاف قائمة الأوامر السريعة (/) — يمكنك إعادة تفعيلها من زر «المساعد الذكي»'
                );
              } else {
                handleUpdateSmartAssistPrefs({
                  ...smartAssistPrefs,
                  enableInlineAutocomplete: false,
                });
                showToast(
                  'تم إيقاف الإكمال التلقائي للوسوم والبنود — يمكنك إعادة تفعيله من زر «المساعد الذكي»'
                );
              }
            }}
          />
        )}
      </div>

      <EditorBottomStatusBar
        autoSaveState={autoSaveState}
        unfilledCount={unfilledCount}
        totalVariablesCount={extractedPlaceholders.length}
        docMetrics={docMetrics}
        activeClauseTitle={outlineClauses.find((c) => c.enabled)?.title}
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
        onUpdateFieldValue={handleUpdateFieldValue}
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
        onOpenDiffComparison={(rev) => {
          setModalActiveBodyHtml(
            bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH
          );
          setDiffComparisonRevision(rev);
          setShowVersionDiffModal(true);
        }}
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
        onStartBlank={handleStartBlank}
        onSelectTemplateFromMultiSource={handleSelectTemplateFromMultiSource}
        onSelectSavedDocFromMultiSource={(doc) => {
          handleSelectDocument(doc);
          showToast(
            `تم فتح العقد "${doc.title}" بجميع خصائصه ومتغيراته المحفوظة`
          );
        }}
        onImportDocxFileFromMultiSource={handleParseAndPreviewDocx}
      />

      <OfficeSettingsModal
        isOpen={showOfficeSettingsModal}
        onClose={() => setShowOfficeSettingsModal(false)}
        localDirState={localDirState}
        onConnectLocalDirectory={() =>
          connectAndReloadLocalDirectory(async () => {
            await reloadLifecycleRecords();
          })
        }
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
        lastBackupAt={lastBackupAt}
        smartAssistPrefs={smartAssistPrefs}
        onChangeSmartAssistPrefs={setSmartAssistPrefs}
        onOpenOnboardingTour={() => setShowOnboardingTour(true)}
      />

      <WordTemplatesManagerModal
        isOpen={showWordTemplatesManager}
        onClose={() => setShowWordTemplatesManager(false)}
        onShowToast={showToast}
        onSelectExportTemplate={handleExportWithStandardTemplate}
      />
    </div>
  );
}
