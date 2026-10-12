'use client';

import React, { useCallback } from 'react';
import {
  ContractFolder,
  ContractOutlineClause,
  ContractPartyCard,
  DocumentRevision,
  NotaryClerk,
  OfficeTemplate,
  OfficeTemplateKind,
  SavedContractDerivedDoc,
  SavedDocument,
  VariableInputType,
} from '@/lib/types';
import {
  loadDocumentRevisions,
  loadSavedDocuments,
  saveDocumentRecord,
} from '@/lib/storage';
import { extractLiveContractClausesFromDOM } from '@/lib/clause-outline-engine';
import {
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
  syncSmartTagsFilledStateInDOM,
} from '@/lib/editor-utils';
import {
  DocxImportResult,
  importNotaryDocxFile,
  mergePlaceholdersIntoHtml,
} from '@/lib/docx-engine';
import { injectNotarialPaginationIntoHtml } from '@/lib/notarial-pagination';
import {
  DocumentSessionSnapshot,
  INITIAL_EMPTY_PARAGRAPH,
} from './useDocumentSession';

interface UseEditorDocumentActionsParams {
  docId: string;
  setDocId: React.Dispatch<React.SetStateAction<string>>;
  docTitle: string;
  setDocTitle: React.Dispatch<React.SetStateAction<string>>;
  showHeaderFooter: boolean;
  setShowHeaderFooter: React.Dispatch<React.SetStateAction<boolean>>;
  pageNumberingEnabled: boolean;
  setPageNumberingEnabled: React.Dispatch<React.SetStateAction<boolean>>;
  previewMergedMode: boolean;
  setMergedPreviewHtml: React.Dispatch<React.SetStateAction<string>>;
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  headerEditorRef: React.RefObject<HTMLDivElement | null>;
  footerEditorRef: React.RefObject<HTMLDivElement | null>;
  savedRangeRef: React.MutableRefObject<Range | null>;
  setUndoStack: React.Dispatch<React.SetStateAction<any[]>>;
  setRedoStack: React.Dispatch<React.SetStateAction<any[]>>;
  fieldValuesRef: React.MutableRefObject<Record<string, string>>;
  setFieldValues: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  fieldInputTypesRef: React.MutableRefObject<Record<string, VariableInputType>>;
  setFieldInputTypes: React.Dispatch<
    React.SetStateAction<Record<string, VariableInputType>>
  >;
  outlineClausesRef: React.MutableRefObject<ContractOutlineClause[]>;
  setOutlineClauses: React.Dispatch<
    React.SetStateAction<ContractOutlineClause[]>
  >;
  contractConditionsRef: React.MutableRefObject<
    Record<string, string | boolean>
  >;
  setContractConditions: React.Dispatch<
    React.SetStateAction<Record<string, string | boolean>>
  >;
  partyCardsRef: React.MutableRefObject<ContractPartyCard[]>;
  setPartyCards: React.Dispatch<React.SetStateAction<ContractPartyCard[]>>;
  setActiveClauseIdsInDoc: React.Dispatch<React.SetStateAction<string[]>>;
  setExtractedPlaceholders: React.Dispatch<React.SetStateAction<string[]>>;
  setSelectedEstateId: React.Dispatch<React.SetStateAction<string>>;
  setSelectedLotNumber: React.Dispatch<React.SetStateAction<string>>;
  setInlinePopover: React.Dispatch<React.SetStateAction<any>>;
  activeClerk: NotaryClerk;
  folders: ContractFolder[];
  setDocuments: (nextDocs: SavedDocument[]) => void;
  setRevisions: React.Dispatch<React.SetStateAction<DocumentRevision[]>>;
  setActiveDocument: React.Dispatch<React.SetStateAction<SavedDocument | null>>;
  activeDocumentRef: React.MutableRefObject<SavedDocument | null>;
  setActiveDerivedDoc: React.Dispatch<
    React.SetStateAction<SavedContractDerivedDoc | null>
  >;
  setAutoSaveState: React.Dispatch<
    React.SetStateAction<'saved' | 'saving' | 'error'>
  >;
  pendingDocxImport: DocxImportResult | null;
  setPendingDocxImport: React.Dispatch<
    React.SetStateAction<DocxImportResult | null>
  >;
  setShowDocxPreviewModal: React.Dispatch<React.SetStateAction<boolean>>;
  clearAutoSaveTimer: () => void;
  ensureBodyEditable: () => boolean;
  recordHistorySnapshot: (label?: string) => void;
  getSessionSnapshot: () => DocumentSessionSnapshot;
  syncPlaceholdersAndDraft: (
    customFieldValues?: Record<string, string>,
    customFieldInputTypes?: Record<string, VariableInputType>,
    customOutlineClauses?: ContractOutlineClause[],
    customPartyCards?: ContractPartyCard[]
  ) => void;
  handleSaveTemplateFromSnapshot: (args: {
    name: string;
    category: string;
    kind?: OfficeTemplateKind;
    code?: string;
    description?: string;
    clearFilledValues?: boolean;
    snapshot: DocumentSessionSnapshot;
  }) => Promise<void>;
  handleSaveDocumentAndRevision: (args: {
    snapshot: DocumentSessionSnapshot;
    liveOutline: ContractOutlineClause[];
    folders: ContractFolder[];
    summaryLabel?: string;
    explicitFolderId?: string | null;
  }) => Promise<SavedDocument>;
  showToast: (msg: string) => void;
}

export function useEditorDocumentActions({
  docId,
  setDocId,
  setDocTitle,
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
}: UseEditorDocumentActionsParams) {
  const handleSelectDocument = useCallback(
    (doc: SavedDocument) => {
      if (!bodyEditorRef.current) return;
      clearAutoSaveTimer();
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
    [
      activeDocumentRef,
      bodyEditorRef,
      clearAutoSaveTimer,
      contractConditionsRef,
      fieldInputTypesRef,
      fieldValuesRef,
      footerEditorRef,
      headerEditorRef,
      outlineClausesRef,
      partyCardsRef,
      setActiveDerivedDoc,
      setActiveDocument,
      setContractConditions,
      setDocId,
      setDocTitle,
      setFieldInputTypes,
      setFieldValues,
      setInlinePopover,
      setOutlineClauses,
      setPageNumberingEnabled,
      setPartyCards,
      setRedoStack,
      setSelectedEstateId,
      setSelectedLotNumber,
      setShowHeaderFooter,
      setUndoStack,
      syncPlaceholdersAndDraft,
    ]
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
      bodyEditorRef.current.innerHTML =
        derivedDoc.content || INITIAL_EMPTY_PARAGRAPH;
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
    [
      bodyEditorRef,
      fieldValuesRef,
      footerEditorRef,
      headerEditorRef,
      recordHistorySnapshot,
      setActiveDerivedDoc,
      setDocTitle,
      setFieldValues,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleParseAndPreviewDocx = useCallback(
    async (file: File) => {
      try {
        const imported = await importNotaryDocxFile(file);
        setPendingDocxImport(imported);
        setShowDocxPreviewModal(true);
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'تعذر قراءة ملف الوورد المرفوع'
        );
      }
    },
    [setPendingDocxImport, setShowDocxPreviewModal, showToast]
  );

  const handleApplyDocxImport = useCallback(
    (mode: 'replace' | 'insert', customSegmentedHtml?: string) => {
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
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      footerEditorRef,
      headerEditorRef,
      outlineClausesRef,
      pendingDocxImport,
      recordHistorySnapshot,
      savedRangeRef,
      setDocTitle,
      setExtractedPlaceholders,
      setOutlineClauses,
      setPendingDocxImport,
      setShowHeaderFooter,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleSaveCurrentAsTemplate = useCallback(
    async (
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
    },
    [getSessionSnapshot, handleSaveTemplateFromSnapshot]
  );

  const handleSaveCustomTemplateAdvanced = useCallback(
    async (data: {
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
    },
    [handleSaveCurrentAsTemplate]
  );

  const handleSaveCurrentDocumentAndRevision = useCallback(
    async (
      summaryLabel?: string,
      explicitFolderId?: string | null
    ): Promise<SavedDocument> => {
      clearAutoSaveTimer();
      const liveOutline = extractLiveContractClausesFromDOM(
        bodyEditorRef.current
      );
      const savedDoc = await handleSaveDocumentAndRevision({
        snapshot: getSessionSnapshot(),
        liveOutline,
        folders,
        summaryLabel,
        explicitFolderId,
      });
      setAutoSaveState('saved');
      return savedDoc;
    },
    [
      bodyEditorRef,
      clearAutoSaveTimer,
      folders,
      getSessionSnapshot,
      handleSaveDocumentAndRevision,
      setAutoSaveState,
    ]
  );

  const handleTakeManualSnapshot = useCallback(
    async (label: string) => {
      recordHistorySnapshot(label);
      await handleSaveCurrentDocumentAndRevision(label);
    },
    [handleSaveCurrentDocumentAndRevision, recordHistorySnapshot]
  );

  const handleRestoreSnapshot = useCallback(
    async (rev: DocumentRevision) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      if (!rev.bodyHtml || !rev.bodyHtml.trim()) {
        showToast(
          'هذه المراجعة سجل حدث ولا تحتوي على نسخة متن نصي للاستعادة'
        );
        return;
      }
      recordHistorySnapshot(
        'قبل استعادة النسخة: ' + (rev.summary || rev.documentTitle)
      );
      bodyEditorRef.current.innerHTML = rev.bodyHtml;
      normalizeNotaryContainerDOM(bodyEditorRef.current);
      const nextVals = rev.fieldValues || {};
      fieldValuesRef.current = nextVals;
      setFieldValues(nextVals);
      syncPlaceholdersAndDraft(nextVals);
      showToast(`تم استعادة النسخة: "${rev.summary || 'نسخة مرجعية'}"`);
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      fieldValuesRef,
      recordHistorySnapshot,
      setFieldValues,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleLoadTemplateFull = useCallback(
    (tpl: OfficeTemplate, clearPreviousClauses = true) => {
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
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      fieldValuesRef,
      footerEditorRef,
      headerEditorRef,
      recordHistorySnapshot,
      setActiveClauseIdsInDoc,
      setDocTitle,
      setFieldValues,
      setShowHeaderFooter,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleInsertTemplateAtCaret = useCallback(
    (tpl: OfficeTemplate) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      insertHtmlAtSelection(
        bodyEditorRef.current,
        tpl.bodyHtml,
        savedRangeRef.current
      );
      syncPlaceholdersAndDraft();
      showToast(`تم إدراج محتوى القالب "${tpl.name}" عند موضع المؤشر`);
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

  const handleStartBlank = useCallback(
    (clearPreviousClauses = true) => {
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
    },
    [
      activeDocumentRef,
      bodyEditorRef,
      fieldInputTypesRef,
      fieldValuesRef,
      footerEditorRef,
      headerEditorRef,
      setActiveClauseIdsInDoc,
      setActiveDocument,
      setDocId,
      setDocTitle,
      setFieldInputTypes,
      setFieldValues,
      setRedoStack,
      setSelectedEstateId,
      setSelectedLotNumber,
      setUndoStack,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleSelectTemplateFromMultiSource = useCallback(
    (
      tpl: OfficeTemplate,
      mode: 'replace' | 'insert',
      clearPreviousClauses: boolean = true
    ) => {
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
    },
    [
      activeDocumentRef,
      bodyEditorRef,
      ensureBodyEditable,
      fieldInputTypesRef,
      fieldValuesRef,
      recordHistorySnapshot,
      savedRangeRef,
      setActiveClauseIdsInDoc,
      setActiveDocument,
      setDocId,
      setDocTitle,
      setFieldInputTypes,
      setFieldValues,
      setRedoStack,
      setSelectedEstateId,
      setSelectedLotNumber,
      setUndoStack,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleContractIndexedOrSynced = useCallback(
    async (updatedDoc?: SavedDocument) => {
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
            injectNotarialPaginationIntoHtml(rawMerged, pageNumberingEnabled)
          );
        }
      }
    },
    [
      activeDocumentRef,
      bodyEditorRef,
      docId,
      fieldValuesRef,
      pageNumberingEnabled,
      previewMergedMode,
      setActiveDocument,
      setDocTitle,
      setDocuments,
      setFieldValues,
      setMergedPreviewHtml,
      setRevisions,
    ]
  );

  return {
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
  };
}
