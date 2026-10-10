'use client';

import React, { useCallback, useEffect, useRef, useState } from 'react';
import {
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  Clock,
  Download,
  Eye,
  FileEdit,
  FileOutput,
  FilePlus2,
  FileText,
  FileUp,
  FolderArchive,
  FormInput,
  GitCompare,
  HelpCircle,
  History,
  ListChecks,
  PanelRightClose,
  PanelRightOpen,
  Printer,
  Replace,
  RotateCcw,
  Save,
  Search,
  Sparkles,
  UserCheck,
  X,
} from 'lucide-react';
import EditorRibbon from './EditorRibbon';
import SidebarWorkspace, { SidebarTab } from './SidebarWorkspace';
import {
  DocxImportPreviewModal,
  MultiSourceStartModal,
  OnboardingTourModal,
  SaveAsTemplateModal,
  SmartVariablesModal,
  SnapshotsHistoryModal,
  VersionDiffModal,
  WordTemplatesModal,
} from './SmartModals';
import {
  ClauseVariableGroup,
  ContractFolder,
  ContractOutlineClause,
  CustomTemplate,
  DerivedDocTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  HistorySnapshot,
  ListNumberingStyle,
  NotaryClerk,
  NotaryClause,
  PartyField,
  SavedContractDerivedDoc,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  SubdivisionEstate,
  SubdivisionLot,
  ToolbarState,
  VariableInputType,
  WordTemplateDefinition,
} from '@/lib/types';
import {
  DEFAULT_DERIVED_DOC_TEMPLATES,
  deleteCustomTemplate,
  deleteDerivedDocTemplate,
  deleteDocumentRecord,
  deleteDocumentRevision,
  deleteDownloadArchiveItem,
  deleteNotaryClause,
  deletePartyRecord,
  deletePropertyRecord,
  deleteSubdivisionEstate,
  deleteWordTemplate,
  exportCustomTemplatesJson,
  exportFullBackupBundle,
  importCustomTemplatesJson,
  importFullBackupBundle,
  loadActiveDraftSession,
  loadContractFolders,
  loadCustomTemplates,
  loadDerivedDocTemplates,
  loadDocumentRevisions,
  loadDownloadArchive,
  loadNotaryClerks,
  loadNotaryClauses,
  loadPartyFields,
  loadSavedDocuments,
  loadSavedParties,
  loadSavedProperties,
  loadSubdivisionEstates,
  loadWordTemplates,
  saveActiveDraftSession,
  saveContractFolder,
  saveCustomTemplate,
  saveDerivedDocTemplate,
  saveDocumentRecord,
  saveDocumentRevision,
  saveDownloadArchiveItem,
  saveNotaryClerk,
  saveNotaryClause,
  savePartyFields,
  savePartyRecord,
  savePropertyRecord,
  saveSubdivisionEstate,
  saveWordTemplate,
  deleteContractFolder,
} from '@/lib/storage';
import {
  computeDocumentMetrics,
  convertSelectionToSmartTag,
  decorateSmartTagsInDOM,
  deleteContractClauseFromDOM,
  escapeHtml,
  extractLiveContractClausesFromDOM,
  extractPlaceholdersFromHtml,
  extractPlaceholdersGroupedByClause,
  findMatchesAcrossNodes,
  focusAndSelectMatchInDOM,
  getIntersectingBlockElements,
  insertHtmlAtSelection,
  insertOrWrapNewClauseAtSelection,
  moveContractClauseInDOM,
  MultiNodeTextMatch,
  normalizeNotaryContainerDOM,
  renameContractClauseInDOM,
  replaceMatchesAcrossNodes,
  restoreSerializedSelection,
  sanitizePastedWordHTML,
  scrollToContractClauseInDOM,
  serializeCurrentSelection,
  syncSmartTagsFilledStateInDOM,
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
import {
  DocxImportResult,
  downloadNotaryDocx,
  importNotaryDocxFile,
  lookupPlaceholderValue,
  mergePlaceholdersIntoHtml,
  normalizePlaceholderKey,
} from '@/lib/docx-engine';
import { parseContractIntoClauses } from '@/lib/docx-template-engine';

const INITIAL_EMPTY_PARAGRAPH = `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><br></p>`;

export default function NotaryEditorApp() {
  // Editor DOM refs
  const bodyEditorRef = useRef<HTMLDivElement | null>(null);
  const headerEditorRef = useRef<HTMLDivElement | null>(null);
  const footerEditorRef = useRef<HTMLDivElement | null>(null);
  const savedRangeRef = useRef<Range | null>(null);
  const activeTableCellRef = useRef<HTMLTableCellElement | null>(null);
  const selectedTableCellsRef = useRef<HTMLTableCellElement[]>([]);

  // Document state
  const [docId, setDocId] = useState<string>('doc_initial');
  const [docTitle, setDocTitle] = useState<string>('عقد توثيقي جديد');
  const [showHeaderFooter, setShowHeaderFooter] = useState<boolean>(false);
  const [pageNumberingEnabled, setPageNumberingEnabled] = useState<boolean>(true);
  const [previewMergedMode, setPreviewMergedMode] = useState<boolean>(false);
  const [mergedPreviewHtml, setMergedPreviewHtml] = useState<string>('');
  const [zoom, setZoom] = useState<number>(100);

  // Mode for editing a Derived Document Template or a Clause directly inside the A4 Editor
  const [editingDerivedTpl, setEditingDerivedTpl] = useState<DerivedDocTemplate | null>(null);
  const [editingClauseObj, setEditingClauseObj] = useState<NotaryClause | null>(null);
  const stashedMainDocRef = useRef<{
    bodyHtml: string;
    headerHtml: string;
    footerHtml: string;
    title: string;
  } | null>(null);

  // Sidebar & Persistent Data State (Default tab = 'clauses' per v2.4 usage priority)
  const [sidebarTab, setSidebarTab] = useState<SidebarTab>('clauses');
  const [isSidebarOpen, setIsSidebarOpen] = useState<boolean>(false);
  const [clauses, setClauses] = useState<NotaryClause[]>([]);
  const [activeClauseIdsInDoc, setActiveClauseIdsInDoc] = useState<string[]>([]);
  const [outlineClauses, setOutlineClauses] = useState<ContractOutlineClause[]>([]);
  const [partyFields, setPartyFields] = useState<PartyField[]>([]);
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const fieldValuesRef = useRef<Record<string, string>>({});
  useEffect(() => {
    fieldValuesRef.current = fieldValues;
  }, [fieldValues]);
  const [fieldInputTypes, setFieldInputTypes] = useState<Record<string, VariableInputType>>({});
  const fieldInputTypesRef = useRef<Record<string, VariableInputType>>({});
  useEffect(() => {
    fieldInputTypesRef.current = fieldInputTypes;
  }, [fieldInputTypes]);
  const [extractedPlaceholders, setExtractedPlaceholders] = useState<string[]>([]);
  const [clauseGroups, setClauseGroups] = useState<ClauseVariableGroup[]>([]);
  const [templates, setTemplates] = useState<CustomTemplate[]>([]);
  const [derivedTemplates, setDerivedTemplates] = useState<DerivedDocTemplate[]>(
    DEFAULT_DERIVED_DOC_TEMPLATES
  );
  const [wordTemplates, setWordTemplates] = useState<WordTemplateDefinition[]>([]);
  const [showWordTemplatesModal, setShowWordTemplatesModal] = useState<boolean>(false);
  const [documents, setDocuments] = useState<SavedDocument[]>([]);
  const documentsRef = useRef<SavedDocument[]>([]);
  useEffect(() => {
    documentsRef.current = documents;
  }, [documents]);
  const [revisions, setRevisions] = useState<DocumentRevision[]>([]);
  const [downloads, setDownloads] = useState<DownloadArchiveItem[]>([]);
  const [estates, setEstates] = useState<SubdivisionEstate[]>([]);
  const [selectedEstateId, setSelectedEstateId] = useState<string>('');
  const [selectedLotNumber, setSelectedLotNumber] = useState<string>('');
  const [savedParties, setSavedParties] = useState<SavedPartyRecord[]>([]);
  const [savedProperties, setSavedProperties] = useState<SavedPropertyRecord[]>([]);
  const [clerks, setClerks] = useState<NotaryClerk[]>([]);
  const [activeClerk, setActiveClerk] = useState<NotaryClerk>({
    id: 'notary-head',
    name: 'الموثق الرئيسي (الأستاذ)',
    role: 'notary',
    color: '#1e3a8a',
    createdAt: '2026-01-01T00:00:00.000Z',
  });
  const [folders, setFolders] = useState<ContractFolder[]>([]);
  const [activeDocument, setActiveDocument] = useState<SavedDocument | null>(null);
  const activeDocumentRef = useRef<SavedDocument | null>(null);
  useEffect(() => {
    activeDocumentRef.current = activeDocument;
  }, [activeDocument]);
  const [activeDerivedDoc, setActiveDerivedDoc] = useState<SavedContractDerivedDoc | null>(null);

  // Modals state
  const [showSmartVarsModal, setShowSmartVarsModal] = useState<boolean>(false);
  const [focusedVarKey, setFocusedVarKey] = useState<string | null>(null);
  const [showVersionDiffModal, setShowVersionDiffModal] = useState<boolean>(false);
  const [diffComparisonRevision, setDiffComparisonRevision] = useState<DocumentRevision | null>(null);
  const [showMultiSourceModal, setShowMultiSourceModal] = useState<boolean>(false);
  const [showDocxPreviewModal, setShowDocxPreviewModal] = useState<boolean>(false);
  const [pendingDocxImport, setPendingDocxImport] = useState<DocxImportResult | null>(null);
  const [showSaveAsTemplateModal, setShowSaveAsTemplateModal] = useState<boolean>(false);
  const [showSnapshotsHistoryModal, setShowSnapshotsHistoryModal] = useState<boolean>(false);
  const [modalActiveBodyHtml, setModalActiveBodyHtml] = useState<string>(INITIAL_EMPTY_PARAGRAPH);
  const [showOnboardingTour, setShowOnboardingTour] = useState<boolean>(false);

  // Auto-save & Status Bar metrics
  const [autoSaveState, setAutoSaveState] = useState<'saved' | 'saving'>('saved');
  const [docMetrics, setDocMetrics] = useState<{
    wordCount: number;
    charCount: number;
    estimatedPages: number;
  }>({ wordCount: 0, charCount: 0, estimatedPages: 1 });
  const autoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  // History Undo/Redo Stack
  const [undoStack, setUndoStack] = useState<HistorySnapshot[]>([]);
  const [redoStack, setRedoStack] = useState<HistorySnapshot[]>([]);

  // Find & Replace State
  const [showFindReplace, setShowFindReplace] = useState<boolean>(false);
  const [findQuery, setFindQuery] = useState<string>('');
  const [replaceQuery, setReplaceQuery] = useState<string>('');
  const [ignoreArabicHamzaAndDiacritics, setIgnoreArabicHamzaAndDiacritics] = useState<boolean>(true);
  const [activeMatchIdx, setActiveMatchIdx] = useState<number>(0);
  const [matches, setMatches] = useState<MultiNodeTextMatch[]>([]);
  const [matchCount, setMatchCount] = useState<number>(0);

  // Toolbar Active State
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

  // Status Toast Notification
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const showToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((prev) => (prev === msg ? null : prev));
    }, 3400);
  }, []);

  const getEditorZonesHtml = useCallback(() => {
    return {
      bodyHtml: bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH,
      headerHtml: headerEditorRef.current?.innerHTML || '',
      footerHtml: footerEditorRef.current?.innerHTML || '',
    };
  }, []);

  // Scan active clause containers and live # outline clauses inside the A4 editor
  const refreshActiveClausesInDOM = useCallback(() => {
    if (!bodyEditorRef.current) return;
    const containers = Array.from(
      bodyEditorRef.current.querySelectorAll('[data-clause-id]')
    ) as HTMLElement[];
    const ids = containers
      .map((el) => el.getAttribute('data-clause-id') || '')
      .filter(Boolean);
    setActiveClauseIdsInDoc(ids);
    const liveOutline = extractLiveContractClausesFromDOM(bodyEditorRef.current);
    setOutlineClauses(liveOutline);
  }, []);

  // Refresh extracted [...] placeholders, live clauses, metrics, and trigger atomic per-contract auto-save
  const syncPlaceholdersAndDraft = useCallback(
    (
      customFieldValues?: Record<string, string>,
      customFieldInputTypes?: Record<string, VariableInputType>
    ) => {
      const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
      const found = extractPlaceholdersFromHtml(headerHtml, bodyHtml, footerHtml);
      setExtractedPlaceholders(found);
      setClauseGroups(
        extractPlaceholdersGroupedByClause(bodyEditorRef.current, found)
      );
      setDocMetrics(computeDocumentMetrics(bodyHtml));
      const liveOutline = extractLiveContractClausesFromDOM(bodyEditorRef.current);
      setOutlineClauses(liveOutline);
      refreshActiveClausesInDOM();

      const activeValues = customFieldValues ?? fieldValuesRef.current ?? fieldValues;
      const activeInputTypes =
        customFieldInputTypes ?? fieldInputTypesRef.current ?? fieldInputTypes;
      syncSmartTagsFilledStateInDOM(bodyEditorRef.current, activeValues);
      syncSmartTagsFilledStateInDOM(headerEditorRef.current, activeValues);
      syncSmartTagsFilledStateInDOM(footerEditorRef.current, activeValues);

      if (previewMergedMode) {
        setMergedPreviewHtml(mergePlaceholdersIntoHtml(bodyHtml, activeValues));
      }

      if (editingDerivedTpl || editingClauseObj) {
        return;
      }

      const existingDoc =
        documentsRef.current.find((d) => d.id === docId) ||
        (activeDocumentRef.current?.id === docId ? activeDocumentRef.current : null);

      const inferredClient =
        existingDoc?.clientName ||
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
      saveActiveDraftSession(draft);

      setAutoSaveState('saving');
      if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
      autoSaveTimerRef.current = setTimeout(async () => {
        await saveDocumentRecord(draft);
        const updatedDocs = await loadSavedDocuments();
        setDocuments(updatedDocs);
        setActiveDocument((prev) => (prev?.id === draft.id ? draft : prev));
        setAutoSaveState('saved');
      }, 450);
    },
    [
      activeClerk.id,
      activeClerk.name,
      docId,
      docTitle,
      editingClauseObj,
      editingDerivedTpl,
      fieldInputTypes,
      fieldValues,
      getEditorZonesHtml,
      pageNumberingEnabled,
      previewMergedMode,
      refreshActiveClausesInDOM,
      selectedEstateId,
      selectedLotNumber,
      showHeaderFooter,
    ]
  );

  // Record a snapshot in the custom Undo Stack and optionally in persistent document revisions
  const recordHistorySnapshot = useCallback(
    (label?: string) => {
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
          author: 'النظام التوثيقي',
          bodyHtml,
          fieldValues: { ...fieldValuesRef.current },
          createdAt: new Date().toISOString(),
          summary: label,
        };
        saveDocumentRevision(rev).then(() => {
          loadDocumentRevisions().then((revs) => setRevisions(revs));
        });
      }
    },
    [docId, docTitle, getEditorZonesHtml]
  );

  const handleCreateFolder = useCallback(
    async (parentId: string | null, name: string) => {
      const newFolder: ContractFolder = {
        id: `folder_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        name,
        parentId,
        clerkId: activeClerk.id,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
      };
      await saveContractFolder(newFolder);
      setFolders(await loadContractFolders());
      showToast(`تم إنشاء المجلد "${name}"`);
    },
    [activeClerk.id, showToast]
  );

  const handleRenameFolder = useCallback(
    async (folderId: string, newName: string) => {
      const f = folders.find((x) => x.id === folderId);
      if (!f) return;
      const updated = { ...f, name: newName, updatedAt: new Date().toISOString() };
      await saveContractFolder(updated);
      setFolders(await loadContractFolders());
      showToast('تمت إعادة تسمية المجلد');
    },
    [folders, showToast]
  );

  const handleDeleteFolder = useCallback(
    async (folderId: string) => {
      const targetFolder = folders.find((f) => f.id === folderId);
      const fallbackParentId = targetFolder?.parentId ?? null;
      await deleteContractFolder(folderId);

      // Promote child folders to parent level
      const childFolders = folders.filter((f) => f.parentId === folderId);
      for (const cf of childFolders) {
        await saveContractFolder({
          ...cf,
          parentId: fallbackParentId,
          updatedAt: new Date().toISOString(),
        });
      }

      // Move documents in this folder to parent level
      const docsInFolder = documents.filter((d) => d.folderId === folderId);
      for (const d of docsInFolder) {
        const updatedDoc = {
          ...d,
          folderId: fallbackParentId,
          updatedAt: new Date().toISOString(),
        };
        await saveDocumentRecord(updatedDoc);
        if (activeDocument?.id === d.id) {
          setActiveDocument(updatedDoc);
        }
      }
      setFolders(await loadContractFolders());
      setDocuments(await loadSavedDocuments());
      showToast('تم حذف المجلد ونقل محتوياته بأمان');
    },
    [activeDocument?.id, documents, folders, showToast]
  );

  const handleMoveDocument = useCallback(
    async (targetDocId: string, targetFolderId: string | null) => {
      const d = documents.find((x) => x.id === targetDocId);
      if (!d) return;
      if ((d.folderId ?? null) === (targetFolderId ?? null)) return;
      const updated: SavedDocument = {
        ...d,
        folderId: targetFolderId,
        updatedAt: new Date().toISOString(),
      };
      await saveDocumentRecord(updated);
      setDocuments(await loadSavedDocuments());
      setActiveDocument((prev) => (prev?.id === targetDocId ? updated : prev));
      const folderName = targetFolderId
        ? folders.find((f) => f.id === targetFolderId)?.name || 'المجلد المختار'
        : 'الجذر الرئيسي';
      showToast(`تم نقل العقد "${d.title}" إلى "${folderName}"`);
    },
    [documents, folders, showToast]
  );

  const handleMoveFolder = useCallback(
    async (folderId: string, targetParentId: string | null) => {
      if (folderId === targetParentId) return;
      const f = folders.find((x) => x.id === folderId);
      if (!f) return;
      if ((f.parentId ?? null) === (targetParentId ?? null)) return;

      // Prevent moving a folder inside one of its own descendants
      let cursor = targetParentId;
      while (cursor) {
        if (cursor === folderId) {
          showToast('لا يمكن نقل المجلد داخل نفسه أو داخل أحد مجلداته الفرعية');
          return;
        }
        const parentNode = folders.find((x) => x.id === cursor);
        cursor = parentNode?.parentId ?? null;
      }

      const updated: ContractFolder = {
        ...f,
        parentId: targetParentId,
        updatedAt: new Date().toISOString(),
      };
      await saveContractFolder(updated);
      setFolders(await loadContractFolders());
      const targetName = targetParentId
        ? folders.find((x) => x.id === targetParentId)?.name || 'المجلد المختار'
        : 'الجذر الرئيسي';
      showToast(`تم نقل المجلد "${f.name}" إلى "${targetName}"`);
    },
    [folders, showToast]
  );

  const handleSelectDocument = useCallback(
    (doc: SavedDocument) => {
      if (!bodyEditorRef.current) return;
      // Flush any pending auto-save before switching contracts
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
      fieldValuesRef.current = isolatedValues;
      fieldInputTypesRef.current = isolatedInputTypes;
      setFieldValues(isolatedValues);
      setFieldInputTypes(isolatedInputTypes);
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
      syncPlaceholdersAndDraft(isolatedValues, isolatedInputTypes);
    },
    [syncPlaceholdersAndDraft]
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
      documentsRef.current = refreshed;
      handleSelectDocument(newDoc);
      const folderName = folderId
        ? folders.find((f) => f.id === folderId)?.name || 'المجلد المختار'
        : 'الجذر الرئيسي';
      showToast(`تم إنشاء عقد جديد معزول داخل "${folderName}"`);
    },
    [activeClerk, folders, handleSelectDocument, showToast]
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

  const handleUpdateDocumentMeta = useCallback(
    async (updated: SavedDocument) => {
      await saveDocumentRecord(updated);
      setDocuments(await loadSavedDocuments());
      setActiveDocument((prev) => (prev?.id === updated.id ? updated : prev));
      showToast('تم تحديث بيانات العقد');
    },
    [showToast]
  );

  // Initial Load from IndexedDB / LocalStorage
  useEffect(() => {
    let mounted = true;
    async function initWorkspace() {
      const [
        loadedTemplates,
        loadedDocs,
        loadedEstates,
        loadedClauses,
        loadedDerived,
        loadedRevs,
        loadedDownloads,
        loadedSavedParties,
        loadedSavedProps,
        loadedClerks,
        loadedFolders,
        loadedWordTpls,
      ] = await Promise.all([
        loadCustomTemplates(),
        loadSavedDocuments(),
        loadSubdivisionEstates(),
        loadNotaryClauses(),
        loadDerivedDocTemplates(),
        loadDocumentRevisions(),
        loadDownloadArchive(),
        loadSavedParties(),
        loadSavedProperties(),
        loadNotaryClerks(),
        loadContractFolders(),
        loadWordTemplates(),
      ]);
      if (!mounted) return;

      const loadedFields = loadPartyFields();
      setPartyFields(loadedFields);
      setTemplates(loadedTemplates);
      setDocuments(loadedDocs);
      setEstates(loadedEstates);
      setClauses(loadedClauses);
      setDerivedTemplates(loadedDerived);
      setWordTemplates(loadedWordTpls);
      setRevisions(loadedRevs);
      setDownloads(loadedDownloads);
      setSavedParties(loadedSavedParties);
      setSavedProperties(loadedSavedProps);
      setClerks(loadedClerks);
      if (loadedClerks.length > 0) setActiveClerk(loadedClerks[0]);
      setFolders(loadedFolders);

      if (loadedEstates.length > 0) {
        setSelectedEstateId(loadedEstates[0].id);
      }

      const activeDraft = loadActiveDraftSession();
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
        fieldValuesRef.current = initialVals;
        fieldInputTypesRef.current = initialTypes;
        setFieldValues(initialVals);
        setFieldInputTypes(initialTypes);
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
        setOutlineClauses(extractLiveContractClausesFromDOM(bodyEditorRef.current));
        setDocMetrics(computeDocumentMetrics(activeDraft.bodyHtml || ''));
        refreshActiveClausesInDOM();
        syncSmartTagsFilledStateInDOM(bodyEditorRef.current, initialVals);
        const matchingSaved = loadedDocs.find((d) => d.id === activeDraft.id);
        if (matchingSaved) {
          setActiveDocument(matchingSaved);
          activeDocumentRef.current = matchingSaved;
        }
      } else if (bodyEditorRef.current) {
        bodyEditorRef.current.innerHTML = INITIAL_EMPTY_PARAGRAPH;
      }

      if (typeof window !== 'undefined') {
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
  }, [refreshActiveClausesInDOM]);

  // Selection Bookmarking
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
  }, []);

  const restoreSelectionBookmarkIfNeeded = useCallback(() => {
    const sel = window.getSelection();
    if (!sel) return;
    if (
      sel.rangeCount > 0 &&
      (bodyEditorRef.current?.contains(sel.getRangeAt(0).commonAncestorContainer) ||
        headerEditorRef.current?.contains(sel.getRangeAt(0).commonAncestorContainer) ||
        footerEditorRef.current?.contains(sel.getRangeAt(0).commonAncestorContainer))
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
  }, []);

  // Synchronize Toolbar State with current Caret / Selection
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
      if (el.style?.backgroundColor && el.style.backgroundColor !== 'transparent') {
        highlightColor = el.style.backgroundColor;
      }
      if (['P', 'LI', 'DIV', 'H1', 'H2', 'H3', 'H4'].includes(tag)) {
        const rawAlign = (el.style?.textAlign || style.textAlign || 'justify').toLowerCase();
        if (rawAlign === 'center') align = 'center';
        else if (rawAlign === 'left') align = 'left';
        else if (rawAlign === 'right' || rawAlign === 'start') align = 'right';
        else align = 'justify';

        const rawDir = (el.getAttribute('dir') || style.direction || 'rtl').toLowerCase();
        dir = rawDir === 'ltr' ? 'ltr' : 'rtl';
      }
      if (tag === 'OL' || tag === 'UL') {
        const lType = el.getAttribute('data-list-type') as ListNumberingStyle | null;
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
        foundCell && ((foundCell.rowSpan || 1) > 1 || (foundCell.colSpan || 1) > 1)
      ),
    });
  }, []);

  useEffect(() => {
    const handler = () => syncToolbarWithSelection();
    document.addEventListener('selectionchange', handler);
    return () => document.removeEventListener('selectionchange', handler);
  }, [syncToolbarWithSelection]);

  // Undo & Redo
  const handleUndo = useCallback(() => {
    if (undoStack.length === 0) return;
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
    if (headerEditorRef.current) headerEditorRef.current.innerHTML = previous.headerHtml;
    if (footerEditorRef.current) footerEditorRef.current.innerHTML = previous.footerHtml;

    restoreSerializedSelection(
      {
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      },
      previous.selection
    );
    syncPlaceholdersAndDraft();
  }, [getEditorZonesHtml, syncPlaceholdersAndDraft, undoStack]);

  const handleRedo = useCallback(() => {
    if (redoStack.length === 0) return;
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
    if (headerEditorRef.current) headerEditorRef.current.innerHTML = next.headerHtml;
    if (footerEditorRef.current) footerEditorRef.current.innerHTML = next.footerHtml;

    restoreSerializedSelection(
      {
        body: bodyEditorRef.current,
        header: headerEditorRef.current,
        footer: footerEditorRef.current,
      },
      next.selection
    );
    syncPlaceholdersAndDraft();
  }, [getEditorZonesHtml, redoStack, syncPlaceholdersAndDraft]);

  // Formatting Commands
  const applyInlineCommand = (command: string, value?: string) => {
    restoreSelectionBookmarkIfNeeded();
    recordHistorySnapshot();
    document.execCommand(command, false, value);
    if (bodyEditorRef.current) normalizeNotaryContainerDOM(bodyEditorRef.current);
    syncToolbarWithSelection();
    syncPlaceholdersAndDraft();
  };

  const handleSetAlign = (align: 'right' | 'center' | 'left' | 'justify') => {
    restoreSelectionBookmarkIfNeeded();
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !bodyEditorRef.current) return;
    recordHistorySnapshot();

    const range = sel.getRangeAt(0);
    const activeZone = headerEditorRef.current?.contains(range.commonAncestorContainer)
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
  };

  const handleSetDirection = (dir: 'rtl' | 'ltr') => {
    restoreSelectionBookmarkIfNeeded();
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !bodyEditorRef.current) return;
    recordHistorySnapshot();

    const range = sel.getRangeAt(0);
    const activeZone = headerEditorRef.current?.contains(range.commonAncestorContainer)
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
  };

  const handleAdjustIndent = (delta: 'increase' | 'decrease') => {
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
      const nextPx = delta === 'increase' ? currentPx + 24 : Math.max(0, currentPx - 24);
      b.style[prop] = nextPx > 0 ? `${nextPx}px` : '';
    }
    syncPlaceholdersAndDraft();
  };

  const handleToggleList = (style: ListNumberingStyle) => {
    restoreSelectionBookmarkIfNeeded();
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();

    const isOrdered =
      style === 'decimal' || style === 'arabic-alpha' || style === 'arabic-abjad';
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

    normalizeNotaryContainerDOM(bodyEditorRef.current);
    syncToolbarWithSelection();
    syncPlaceholdersAndDraft();
  };

  // Table Creation & Virtual Grid Actions
  const handleInsertTable = (rows: number, cols: number) => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();
    const html = createNotaryTableHTML(rows, cols);
    insertHtmlAtSelection(bodyEditorRef.current, html, savedRangeRef.current);
    syncToolbarWithSelection();
    syncPlaceholdersAndDraft();
  };

  const handleTableAction = (
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
  };

  const handleInsertPageBreak = () => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();
    const breakHtml = `<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div><p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><br></p>`;
    insertHtmlAtSelection(bodyEditorRef.current, breakHtml, savedRangeRef.current);
    syncPlaceholdersAndDraft();
  };

  const handleInsertImageFile = (file: File) => {
    if (!bodyEditorRef.current) return;
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === 'string' && bodyEditorRef.current) {
        recordHistorySnapshot();
        const imgHtml = `<img src="${reader.result}" alt="${escapeHtml(
          file.name
        )}" style="max-width:100%;height:auto;" />`;
        insertHtmlAtSelection(bodyEditorRef.current, imgHtml, savedRangeRef.current);
        syncPlaceholdersAndDraft();
      }
    };
    reader.readAsDataURL(file);
  };

  // Double-click on a .smart-tag inside the A4 Editor -> Open SmartVariablesModal focused on that variable
  const handleEditorDoubleClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    const smartTag = target.closest('.smart-tag, .smart-placeholder') as HTMLElement | null;
    if (smartTag) {
      e.preventDefault();
      const varKey =
        smartTag.getAttribute('data-var') ||
        (smartTag.textContent || '').replace(/[\[\]{}]/g, '').trim();
      if (varKey) {
        setFocusedVarKey(varKey);
        setShowSmartVarsModal(true);
      }
    }
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
  // LIVE CONTRACT OUTLINE CLAUSES (# HEADINGS) & LIBRARY HANDLERS
  // =========================================================
  const handleScrollToLiveClause = (domIndex: number) => {
    if (!bodyEditorRef.current) return;
    scrollToContractClauseInDOM(bodyEditorRef.current, domIndex);
  };

  const handleMoveLiveClauseInDoc = (
    domIndex: number,
    direction: 'up' | 'down'
  ) => {
    if (!bodyEditorRef.current) return;
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

  const handleDeleteLiveClauseFromDoc = (domIndex: number) => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot('قبل حذف بند من العقد');
    const deleted = deleteContractClauseFromDOM(
      bodyEditorRef.current,
      domIndex
    );
    if (deleted) {
      syncPlaceholdersAndDraft();
      showToast('تم حذف البند من العقد الحالي');
    }
  };

  const handleRenameLiveClauseInDoc = (domIndex: number, newTitle: string) => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot('قبل تعديل عنوان بند في العقد');
    const renamed = renameContractClauseInDOM(
      bodyEditorRef.current,
      domIndex,
      newTitle
    );
    if (renamed) {
      syncPlaceholdersAndDraft();
      showToast(`تم تحديث عنوان البند إلى "# ${newTitle}"`);
    }
  };

  const handleToggleClauseInDoc = (clause: NotaryClause) => {
    if (!bodyEditorRef.current) return;
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
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();
    const containerHtml = `<div class="clause-container" data-clause-id="${
      clause.id
    }" data-clause-title="${escapeHtml(clause.title)}">${clause.contentHtml}</div>`;
    insertHtmlAtSelection(
      bodyEditorRef.current,
      containerHtml,
      savedRangeRef.current
    );
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج البند المقترح "${clause.title}" عند المؤشر`);
  };

  const handleMoveClauseOrder = async (
    clauseId: string,
    direction: 'up' | 'down'
  ) => {
    const idx = clauses.findIndex((c) => c.id === clauseId);
    if (idx < 0) return;
    const swapIdx = direction === 'up' ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= clauses.length) return;

    const updated = [...clauses];
    const temp = updated[idx];
    updated[idx] = updated[swapIdx];
    updated[swapIdx] = temp;

    for (let i = 0; i < updated.length; i++) {
      updated[i] = { ...updated[i], order: i };
      await saveNotaryClause(updated[i]);
    }
    setClauses(await loadNotaryClauses());
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
    const newClause: NotaryClause = {
      id: `clause_${Date.now()}`,
      title,
      category,
      contentHtml: bodyHtml,
      enabled: true,
      order: clauses.length,
      updatedAt: new Date().toISOString(),
    };
    await saveNotaryClause(newClause);
    setClauses(await loadNotaryClauses());
    showToast(`تم حفظ البند "${title}" في مكتبة البنود الجاهزة`);
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
  // DERIVED DOCUMENTS GENERATOR & IN-APP TEMPLATE EDITING
  // =========================================================
  const handleGenerateDerivedDoc = async (tpl: DerivedDocTemplate) => {
    try {
      const activeValues = fieldValuesRef.current || fieldValues;
      // Strictly rely on current Smart Variables Form + Parties & Property Designations
      const requiredVars = extractPlaceholdersFromHtml(
        tpl.headerHtml,
        tpl.bodyHtml,
        tpl.footerHtml
      );
      const missingVars = requiredVars.filter(
        (k) => !activeValues[k] || !activeValues[k].trim()
      );

      if (missingVars.length > 0) {
        showToast(
          `تنبيه: تم توليد "${tpl.name}" مع وجود (${missingVars.length}) متغير غير معبأ بعد في استمارة المتغيرات`
        );
      }

      const fileTitle = `${docTitle || 'عقد'} - ${tpl.name}`;
      await downloadNotaryDocx({
        title: fileTitle,
        bodyHtml: tpl.bodyHtml,
        headerHtml: tpl.headerHtml,
        footerHtml: tpl.footerHtml,
        pageNumberingEnabled: true,
        fieldValues: activeValues,
      });

      // Log into Download Archive
      const archiveItem: DownloadArchiveItem = {
        id: `dl_${Date.now()}`,
        documentId: docId,
        documentTitle: docTitle || 'عقد توثيقي',
        docTypeCode: tpl.code,
        docTypeLabel: tpl.name,
        fileName: `${fileTitle}.docx`,
        bodyHtml: mergePlaceholdersIntoHtml(tpl.bodyHtml, activeValues),
        headerHtml: mergePlaceholdersIntoHtml(tpl.headerHtml, activeValues),
        footerHtml: mergePlaceholdersIntoHtml(tpl.footerHtml, activeValues),
        fieldValues: { ...activeValues },
        createdAt: new Date().toISOString(),
      };
      await saveDownloadArchiveItem(archiveItem);
      setDownloads(await loadDownloadArchive());

      if (missingVars.length === 0) {
        showToast(`تم توليد وتصدير "${tpl.name}" (.docx) وحفظه في أرشيف التحميلات`);
      }
    } catch (err) {
      showToast(err instanceof Error ? err.message : 'تعذر توليد الوثيقة المشتقة');
    }
  };

  const handleEditDerivedTemplateInEditor = (tpl: DerivedDocTemplate) => {
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
    showToast(`أنت الآن تعدل قالب الوثيقة المشتقة "${tpl.name}" داخل ورقة الـ A4`);
  };

  const handleFinishEditingDerivedTemplate = async (saveChanges: boolean) => {
    if (!editingDerivedTpl || !bodyEditorRef.current) return;
    if (saveChanges) {
      const updatedTpl: DerivedDocTemplate = {
        ...editingDerivedTpl,
        bodyHtml: bodyEditorRef.current.innerHTML,
        headerHtml: headerEditorRef.current?.innerHTML || '',
        footerHtml: footerEditorRef.current?.innerHTML || '',
        updatedAt: new Date().toISOString(),
      };
      await saveDerivedDocTemplate(updatedTpl);
      setDerivedTemplates(await loadDerivedDocTemplates());
      showToast(`تم حفظ تعديلات قالب "${updatedTpl.name}" بنجاح`);
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

  const handleImportDocxForDerivedTemplate = async (
    tpl: DerivedDocTemplate,
    file: File
  ) => {
    try {
      const res = await importNotaryDocxFile(file);
      const updated: DerivedDocTemplate = {
        ...tpl,
        bodyHtml: res.bodyHtml,
        headerHtml: res.headerHtml,
        footerHtml: res.footerHtml,
        updatedAt: new Date().toISOString(),
      };
      await saveDerivedDocTemplate(updated);
      setDerivedTemplates(await loadDerivedDocTemplates());
      showToast(`تم تحديث قالب "${tpl.name}" من ملف الوورد المرفوع بنجاح`);
    } catch (err) {
      showToast(err instanceof Error ? err.message : 'تعذر قراءة ملف الوورد');
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
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();
    const clauseHtml = `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><span style="font-weight:bold;">التعيين العقاري (الحصة رقم ${escapeHtml(
      lot.lotNumber
    )}): </span>${escapeHtml(lot.fullDescription)}${
      estate.subdivisionDeedRef
        ? ` وذلك بموجب الجدول الوصفي للتقسيم المحرر بتاريخ ومراجع: ${escapeHtml(
            estate.subdivisionDeedRef
          )}.`
        : ''
    }</p>`;
    insertHtmlAtSelection(bodyEditorRef.current, clauseHtml, savedRangeRef.current);
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج فقرة تعيين الحصة رقم (${lot.lotNumber}) في العقد`);
  };

  const handleInsertFullSubdivisionTableAtCaret = (estate: SubdivisionEstate) => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();
    let rowsHtml = `<tr>
      <td style="width:14%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">رقم الحصة</p></td>
      <td style="width:20%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">الطبيعة</p></td>
      <td style="width:20%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">الطابق والعمارة</p></td>
      <td style="width:18%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">المساحة (م²)</p></td>
      <td style="width:28%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">الحصة في الأجزاء المشتركة</p></td>
    </tr>`;

    for (const lot of estate.lots) {
      rowsHtml += `<tr>
        <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
          lot.lotNumber
        )}</p></td>
        <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
          lot.nature
        )}</p></td>
        <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
          `${lot.floor} ${lot.building}`.trim()
        )}</p></td>
        <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
          lot.area
        )}</p></td>
        <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
          lot.commonShares
        )}</p></td>
      </tr>`;
    }

    const tableHtml = `<p dir="rtl" style="margin:0;line-height:1;font-weight:bold;">جدول الوصف التقسيمي للعقار: ${escapeHtml(
      estate.estateName
    )}</p><table dir="rtl" style="width:100%;max-width:120mm;border-collapse:collapse;table-layout:fixed;margin:0;"><tbody>${rowsHtml}</tbody></table><p dir="rtl" style="margin:0;line-height:1;"><br></p>`;

    insertHtmlAtSelection(bodyEditorRef.current, tableHtml, savedRangeRef.current);
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج جدول الوصف التقسيمي (${estate.estateName}) في العقد`);
  };

  // =========================================================
  // EXPORT TO WORD (.docx) + LOG TO DOWNLOAD ARCHIVE
  // =========================================================
  const handleExportCurrentToWord = async () => {
    try {
      const activeValues = fieldValuesRef.current || fieldValues;
      const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
      const titleToUse = docTitle || 'عقد_توثيقي';
      await downloadNotaryDocx({
        title: titleToUse,
        bodyHtml,
        headerHtml: showHeaderFooter ? headerHtml : '',
        footerHtml: showHeaderFooter ? footerHtml : '',
        pageNumberingEnabled,
        fieldValues: activeValues,
      });

      const archiveItem: DownloadArchiveItem = {
        id: `dl_${Date.now()}`,
        documentId: docId,
        documentTitle: titleToUse,
        docTypeCode: 'acte',
        docTypeLabel: 'الأصل WORD (العقد التوثيقي)',
        fileName: `${titleToUse}.docx`,
        bodyHtml,
        headerHtml: showHeaderFooter ? headerHtml : '',
        footerHtml: showHeaderFooter ? footerHtml : '',
        fieldValues: { ...activeValues },
        createdAt: new Date().toISOString(),
      };
      await saveDownloadArchiveItem(archiveItem);
      setDownloads(await loadDownloadArchive());

      showToast('تم تصدير ملف Word (.docx) وحفظه في أرشيف التحميلات بنجاح');
    } catch (err) {
      showToast(err instanceof Error ? err.message : 'تعذر تصدير ملف الوورد');
    }
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

  const handleApplyDocxImport = (mode: 'replace' | 'insert') => {
    if (!pendingDocxImport) return;

    if (mode === 'replace') {
      recordHistorySnapshot('قبل استبدال العقد بملف Word مستورد');
      setDocTitle(pendingDocxImport.title);
      if (bodyEditorRef.current) {
        bodyEditorRef.current.innerHTML = pendingDocxImport.bodyHtml;
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
      syncPlaceholdersAndDraft();
      showToast(
        `تم استبدال العقد بمحتوى ملف "${pendingDocxImport.title}" وتوحيد قياسات A4 بنجاح`
      );
    } else {
      // Insert at caret / savedRangeRef
      if (!bodyEditorRef.current) return;
      recordHistorySnapshot('قبل إدراج ملف Word عند المؤشر');
      insertHtmlAtSelection(
        bodyEditorRef.current,
        pendingDocxImport.bodyHtml,
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

  const handleImportDocxAsTemplate = async (files: FileList) => {
    let importedCount = 0;
    for (const file of Array.from(files)) {
      try {
        const res = await importNotaryDocxFile(file);
        const newTpl: CustomTemplate = {
          id: `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
          name: res.title,
          category: 'نماذج وورد مستوردة',
          bodyHtml: res.bodyHtml,
          headerHtml: res.headerHtml,
          footerHtml: res.footerHtml,
          pageNumberingEnabled: true,
          extractedPlaceholders: res.extractedPlaceholders,
          sourceFileName: file.name,
          updatedAt: new Date().toISOString(),
        };
        await saveCustomTemplate(newTpl);
        importedCount++;
      } catch {
        // Skip invalid file
      }
    }
    const updated = await loadCustomTemplates();
    setTemplates(updated);
    if (importedCount > 0) {
      showToast(`تم استيراد وحفظ (${importedCount}) قالب في مكتبة المكتب بنجاح`);
    }
  };

  const handleSaveCurrentAsTemplate = async (
    name: string,
    category: string,
    description?: string,
    clearFilledValues: boolean = true
  ) => {
    const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
    const placeholders = extractPlaceholdersFromHtml(
      headerHtml,
      bodyHtml,
      footerHtml
    );
    const newTpl: CustomTemplate = {
      id: `tpl_${Date.now()}`,
      name,
      category,
      description: description || '',
      bodyHtml,
      headerHtml,
      footerHtml,
      pageNumberingEnabled,
      extractedPlaceholders: placeholders,
      defaultFieldValues: clearFilledValues ? {} : { ...fieldValues },
      updatedAt: new Date().toISOString(),
    };
    await saveCustomTemplate(newTpl);
    setTemplates(await loadCustomTemplates());
    showToast(`تم حفظ القالب "${name}" في مكتبة المكتب بنجاح`);
  };

  const handleSaveCustomTemplateAdvanced = async (data: {
    name: string;
    category: string;
    description: string;
    clearFilledValues: boolean;
  }) => {
    await handleSaveCurrentAsTemplate(
      data.name,
      data.category,
      data.description,
      data.clearFilledValues
    );
  };

  const handleExportTemplatesJson = async () => {
    try {
      const jsonStr = await exportCustomTemplatesJson();
      const blob = new Blob([jsonStr], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `قوالب_مكتب_التوثيق_${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast('تم تصدير قوالب المكتب إلى ملف JSON بنجاح');
    } catch {
      showToast('تعذر تصدير ملف القوالب');
    }
  };

  const handleImportTemplatesJson = async (file: File) => {
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        const count = await importCustomTemplatesJson(parsed);
        setTemplates(await loadCustomTemplates());
        showToast(`تم استيراد وحفظ (${count}) قالب في مكتبة المكتب`);
      } else {
        showToast('ملف JSON غير صالح كقائمة قوالب');
      }
    } catch {
      showToast('تعذر قراءة ملف القوالب');
    }
  };

  const handleSaveCurrentDocumentAndRevision = async (
    summaryLabel?: string,
    explicitFolderId?: string | null
  ) => {
    const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
    const nowIso = new Date().toISOString();
    const existingDoc =
      documents.find((d) => d.id === docId) ||
      (activeDocument?.id === docId ? activeDocument : undefined);
    const resolvedFolderId =
      explicitFolderId !== undefined
        ? explicitFolderId
        : existingDoc?.folderId ?? activeDocument?.folderId ?? null;

    const inferredClient =
      existingDoc?.clientName ||
      activeDocument?.clientName ||
      fieldValues['الطرف_الأول_الاسم'] ||
      fieldValues['البائع'] ||
      fieldValues['المؤجر'] ||
      fieldValues['الموكل'] ||
      '';

    const liveOutline = extractLiveContractClausesFromDOM(bodyEditorRef.current);
    const docRecord: SavedDocument = {
      id: docId,
      title: docTitle || 'عقد بدون عنوان',
      bodyHtml,
      headerHtml,
      footerHtml,
      pageNumberingEnabled,
      showHeaderFooter,
      fieldValues,
      fieldInputTypes,
      outlineClauses: liveOutline,
      selectedEstateId,
      selectedLotNumber,
      folderId: resolvedFolderId,
      clerkId: existingDoc?.clerkId || activeDocument?.clerkId || activeClerk.id,
      clerkName: existingDoc?.clerkName || activeDocument?.clerkName || activeClerk.name,
      contractNumber:
        existingDoc?.contractNumber ||
        activeDocument?.contractNumber ||
        fieldValues['رقم_الفهرس'] ||
        `2026/${Math.floor(100 + Math.random() * 900)}`,
      year: existingDoc?.year || activeDocument?.year || new Date().getFullYear(),
      clientName: inferredClient,
      modelId: existingDoc?.modelId || activeDocument?.modelId,
      modelTitle: existingDoc?.modelTitle || activeDocument?.modelTitle,
      status: existingDoc?.status || activeDocument?.status || 'editing',
      derivedDocuments:
        existingDoc?.derivedDocuments || activeDocument?.derivedDocuments || [],
      updatedAt: nowIso,
      createdAt: existingDoc?.createdAt || activeDocument?.createdAt || nowIso,
    };
    await saveDocumentRecord(docRecord);
    const refreshedDocs = await loadSavedDocuments();
    setDocuments(refreshedDocs);
    documentsRef.current = refreshedDocs;
    setActiveDocument(docRecord);
    activeDocumentRef.current = docRecord;

    const rev: DocumentRevision = {
      id: `rev_${Date.now()}`,
      documentId: docId,
      documentTitle: docTitle || 'عقد توثيقي',
      author: activeClerk?.name || 'كاتب المكتب',
      bodyHtml,
      fieldValues: { ...fieldValues },
      createdAt: nowIso,
      summary: summaryLabel || 'حفظ نسخة مرجعية للعقد',
    };
    await saveDocumentRevision(rev);
    setRevisions(await loadDocumentRevisions());

    if (explicitFolderId !== undefined) {
      const folderName = explicitFolderId
        ? folders.find((f) => f.id === explicitFolderId)?.name || 'المجلد المختار'
        : 'الجذر الرئيسي';
      showToast(`تم حفظ العقد "${docRecord.title}" داخل "${folderName}" بنجاح`);
    } else {
      showToast(summaryLabel || 'تم حفظ العقد في شجرة المكتب وتسجيل نسخة مرجعية');
    }
  };

  const handleTakeManualSnapshot = async (label: string) => {
    recordHistorySnapshot(label);
    await handleSaveCurrentDocumentAndRevision(label);
  };

  const handleRestoreSnapshot = async (rev: DocumentRevision) => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot('قبل استعادة النسخة: ' + (rev.summary || rev.documentTitle));
    bodyEditorRef.current.innerHTML = rev.bodyHtml;
    normalizeNotaryContainerDOM(bodyEditorRef.current);
    const nextVals = rev.fieldValues || {};
    fieldValuesRef.current = nextVals;
    setFieldValues(nextVals);
    syncPlaceholdersAndDraft(nextVals);
    showToast(`تم استعادة النسخة: "${rev.summary || 'نسخة مرجعية'}"`);
  };

  const handleDeleteSnapshot = async (id: string) => {
    await deleteDocumentRevision(id);
    setRevisions(await loadDocumentRevisions());
    showToast('تم حذف النسخة من السجل');
  };

  const handleOpenSnapshotsHistoryModal = () => {
    setModalActiveBodyHtml(bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH);
    setShowSnapshotsHistoryModal(true);
  };

  const handleOpenVersionDiffModal = () => {
    setModalActiveBodyHtml(bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH);
    setDiffComparisonRevision(null);
    setShowVersionDiffModal(true);
  };

  const handleOpenDiffComparison = (rev: DocumentRevision) => {
    setModalActiveBodyHtml(bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH);
    setDiffComparisonRevision(rev);
    setShowVersionDiffModal(true);
  };

  // Saved Parties & Saved Properties Directory Handlers (v2.5)
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
    const prefix = role === 'party1' ? 'الطرف_الأول' : 'الطرف_الثاني';
    const vals = fieldValuesRef.current;
    const fullName = (vals[`${prefix}_الاسم`] || '').trim();
    if (!fullName) {
      showToast(
        `يرجى إدخال اسم ${
          role === 'party1' ? 'الطرف الأول' : 'الطرف الثاني'
        } أولاً قبل الحفظ في الدفتر`
      );
      return;
    }
    const existing = savedParties.find((p) => p.fullName === fullName);
    const record: SavedPartyRecord = {
      id: existing ? existing.id : `party_${Date.now()}`,
      fullName,
      birthDate: vals[`${prefix}_تاريخ_الميلاد`] || '',
      birthPlace: vals[`${prefix}_مكان_الميلاد`] || '',
      parentage: vals[`${prefix}_النسب`] || '',
      idCardRef: vals[`${prefix}_الهوية`] || '',
      address: vals[`${prefix}_الإقامة`] || '',
      updatedAt: new Date().toISOString(),
    };
    await savePartyRecord(record);
    setSavedParties(await loadSavedParties());
    showToast(`تم حفظ الطرف "${fullName}" في دفتر الأطراف المحفوظين`);
  };

  const handleDeleteSavedParty = async (id: string) => {
    await deletePartyRecord(id);
    setSavedParties(await loadSavedParties());
    showToast('تم حذف الطرف من الدفتر');
  };

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
    const vals = fieldValuesRef.current;
    const lotNumber = (vals['رقم_الحصة'] || '').trim();
    const nature = (vals['طبيعة_الحصة'] || '').trim();
    const fullDesc = (vals['تعيين_الحصة_الكامل'] || '').trim();
    if (!lotNumber && !fullDesc) {
      showToast('يرجى إدخال رقم الحصة أو التعيين الكامل أولاً قبل حفظ العقار في الدفتر');
      return;
    }
    const label = `${nature || 'عقار'} ${lotNumber ? `حصة رقم ${lotNumber}` : ''}`.trim();
    const record: SavedPropertyRecord = {
      id: `prop_${Date.now()}`,
      label,
      lotNumber,
      nature,
      floor: vals['الطابق'] || '',
      area: vals['المساحة'] || '',
      commonShares: vals['الأجزاء_المشتركة'] || '',
      fullDescription: fullDesc,
      deedRef: vals['مراجع_الوصف_التقسيمي'] || '',
      updatedAt: new Date().toISOString(),
    };
    await savePropertyRecord(record);
    setSavedProperties(await loadSavedProperties());
    showToast(`تم حفظ "${label}" في دفتر العقارات`);
  };

  const handleDeleteSavedProperty = async (id: string) => {
    await deletePropertyRecord(id);
    setSavedProperties(await loadSavedProperties());
    showToast('تم حذف العقار من الدفتر');
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

  // 1-Click Selection-to-Smart-Tag Converter (تحويل المحدد إلى وسم بضغطة واحدة دون نافذة)
  const handleConvertSelectionToSmartTag = () => {
    if (!bodyEditorRef.current) return;
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

  // # New Clause Toolbar Action (زر # بند جديد خاص بالعقد الحالي فقط)
  const handleInsertNewClauseHeadingAtCaret = async (customTitle?: string) => {
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();
    if (customTitle && customTitle.trim()) {
      const cleanTitle = customTitle.replace(/^#+\s*/, '').trim();
      const clauseId = `clause_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;
      const innerContentHtml = `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># ${escapeHtml(
        cleanTitle
      )}</span></p><p dir="rtl" style="${pStyle}">اكتب نص البند أو أدرج المتغيرات [اسم_المتغير] هنا.</p>`;
      const wrapperHtml = `<div class="clause-container" data-clause-id="${clauseId}" data-clause-title="${escapeHtml(
        cleanTitle
      )}">${innerContentHtml}</div><p dir="rtl" style="${pStyle}"><br></p>`;
      insertHtmlAtSelection(bodyEditorRef.current, wrapperHtml, savedRangeRef.current);
      syncPlaceholdersAndDraft();
      showToast(`تم إدراج عنوان البند "# ${cleanTitle}" في هيكل العقد الحالي`);
      return;
    }
    const { title } = insertOrWrapNewClauseAtSelection(
      bodyEditorRef.current,
      savedRangeRef.current,
      outlineClauses.length + 1
    );
    syncPlaceholdersAndDraft();
    showToast(`تم إدراج عنوان البند "${title}" في هيكل العقد الحالي`);
  };

  const handleInsertSmartTagAtCaret = (varKey: string) => {
    if (!bodyEditorRef.current || !varKey.trim()) return;
    const cleanKey = normalizePlaceholderKey(varKey).replace(/\s+/g, '_');
    if (!cleanKey) return;
    recordHistorySnapshot();
    if (!partyFields.some((f) => f.key === cleanKey)) {
      const updated: PartyField[] = [
        ...partyFields,
        {
          key: cleanKey,
          label: cleanKey.replace(/_/g, ' '),
          value: '',
          category: 'custom',
        },
      ];
      setPartyFields(updated);
      savePartyFields(updated);
    }
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
    if (!bodyEditorRef.current) return;
    recordHistorySnapshot();

    const activeValues: Record<string, string> = {
      ...fieldValues,
      ...fieldValuesRef.current,
      ...(explicitValues || {}),
    };
    fieldValuesRef.current = activeValues;
    setFieldValues(activeValues);

    // 1. Decorate any raw/manually-typed {{...}} into .smart-tag, then replace in live DOM + HTML
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

  // Find & Replace Handlers (Arabic-Aware with Hamza & Diacritics Normalization)
  useEffect(() => {
    if (!showFindReplace || !findQuery.trim() || !bodyEditorRef.current) {
      setMatches([]);
      setMatchCount(0);
      setActiveMatchIdx(0);
      return;
    }
    const foundMatches = findMatchesAcrossNodes(
      bodyEditorRef.current,
      findQuery,
      ignoreArabicHamzaAndDiacritics
    );
    setMatches(foundMatches);
    setMatchCount(foundMatches.length);
    if (foundMatches.length > 0) {
      setActiveMatchIdx(0);
      focusAndSelectMatchInDOM(foundMatches[0]);
    }
  }, [findQuery, showFindReplace, ignoreArabicHamzaAndDiacritics]);

  const handleNextMatch = () => {
    if (matches.length === 0) return;
    const next = (activeMatchIdx + 1) % matches.length;
    setActiveMatchIdx(next);
    focusAndSelectMatchInDOM(matches[next]);
  };

  const handlePrevMatch = () => {
    if (matches.length === 0) return;
    const prev = (activeMatchIdx - 1 + matches.length) % matches.length;
    setActiveMatchIdx(prev);
    focusAndSelectMatchInDOM(matches[prev]);
  };

  const handleRunReplace = (replaceAll: boolean) => {
    if (!bodyEditorRef.current || !findQuery) return;
    recordHistorySnapshot(
      replaceAll ? 'قبل استبدال الكل في البحث' : 'قبل استبدال تطابق نصي'
    );
    const count = replaceMatchesAcrossNodes(
      bodyEditorRef.current,
      findQuery,
      replaceQuery,
      replaceAll,
      activeMatchIdx,
      ignoreArabicHamzaAndDiacritics
    );
    const updated = findMatchesAcrossNodes(
      bodyEditorRef.current,
      findQuery,
      ignoreArabicHamzaAndDiacritics
    );
    setMatches(updated);
    setMatchCount(updated.length);
    if (updated.length > 0) {
      const nextIdx = Math.min(activeMatchIdx, updated.length - 1);
      setActiveMatchIdx(nextIdx);
      focusAndSelectMatchInDOM(updated[nextIdx]);
    }
    syncPlaceholdersAndDraft();
    showToast(`تم استبدال (${count}) تطابق بنجاح`);
  };

  // Compute unfilled variables count for the badge
  const unfilledCount = extractedPlaceholders.filter((k) => {
    const val = lookupPlaceholderValue(fieldValues, k);
    return !val || !val.trim();
  }).length;

  return (
    <div className="h-screen print:h-auto overflow-hidden print:overflow-visible flex flex-col bg-slate-100 text-slate-900">
      {/* TOP HEADER BAR: BRAND & CONTRACT TITLE (RIGHT) | FILE & EXPORT ACTIONS (LEFT) */}
      <header className="flex flex-wrap items-center justify-between gap-2 px-3 sm:px-4 py-2 bg-white border-b border-slate-200 shrink-0 no-print select-none">
        {/* Right: Brand & Editable Contract Title Input with Auto-Save Badge */}
        <div className="flex items-center gap-2.5 shrink-0">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-blue-900 text-white flex items-center justify-center font-bold text-sm shadow-2xs">
              م
            </div>
            <span className="text-sm font-bold tracking-tight text-slate-900 hidden sm:inline">
              الموثق الرقمي
            </span>
          </div>

          <div className="h-5 w-px bg-slate-200 hidden sm:block" />

          <div className="flex items-center gap-1.5 bg-slate-50 hover:bg-slate-100/90 border border-slate-200 focus-within:border-blue-900 focus-within:bg-white rounded-lg px-2.5 py-1 transition-all">
            <FileEdit className="w-3.5 h-3.5 text-slate-400 shrink-0" />
            <input
              type="text"
              value={docTitle}
              onChange={(e) => setDocTitle(e.target.value)}
              title="انقر لتعديل عنوان العقد مباشرة"
              className="bg-transparent text-xs font-bold text-slate-900 focus:outline-none w-36 sm:w-52 md:w-64 truncate"
              placeholder="عنوان العقد..."
            />
            <span className="text-[10px] shrink-0">
              {autoSaveState === 'saving' ? (
                <span className="text-amber-600 font-medium">جاري الحفظ...</span>
              ) : (
                <span className="text-emerald-600 font-medium">محفوظ ✓</span>
              )}
            </span>
          </div>
        </div>

        {/* Left: Document File & Export Commands */}
        <div className="flex items-center gap-1.5 sm:gap-2 shrink-0">
          {/* New Contract Modal */}
          <button
            type="button"
            onClick={() => setShowMultiSourceModal(true)}
            className="px-2.5 py-1.5 bg-white border border-slate-300 hover:bg-slate-50 text-slate-700 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
            title="بدء عقد جديد من قالب أو مسودة أو ملف وورد أو أرشيف التحميلات"
          >
            <FilePlus2 className="w-3.5 h-3.5 text-blue-900" />
            <span className="hidden md:inline">عقد جديد</span>
          </button>

          {/* Import Word (.docx) */}
          <label className="px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 cursor-pointer whitespace-nowrap shrink-0 transition-colors">
            <FileUp className="w-3.5 h-3.5 text-blue-900" />
            <span className="hidden md:inline">فتح Word</span>
            <input
              type="file"
              accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  handleImportDocxDirectlyToEditor(file);
                  e.target.value = '';
                }
              }}
            />
          </label>

          {/* Side-by-Side Merged Preview Toggle */}
          <button
            type="button"
            onClick={() => {
              const next = !previewMergedMode;
              if (next) {
                const currentHtml =
                  bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH;
                const activeVals = fieldValuesRef.current || fieldValues;
                setMergedPreviewHtml(mergePlaceholdersIntoHtml(currentHtml, activeVals));
                setPreviewMergedMode(true);
                setSidebarTab('parties');
                setIsSidebarOpen(true);
                showToast(
                  'وضع معاينة الدمج جنباً إلى جنب: عدّل أي قيمة في اليمين لترى العقد المدمج مباشرة'
                );
              } else {
                setPreviewMergedMode(false);
              }
            }}
            className={`px-2.5 py-1.5 border rounded-lg text-xs font-medium inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 transition-colors ${
              previewMergedMode
                ? 'bg-amber-100 border-amber-300 text-amber-950 font-bold'
                : 'bg-white border-slate-300 text-slate-700 hover:bg-slate-50'
            }`}
            title="معاينة دمج الحقول جنباً إلى جنب مع الاستمارة قبل التصدير"
          >
            <Eye className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">
              {previewMergedMode ? 'وضع التحرير' : 'معاينة'}
            </span>
          </button>

          {/* Snapshots & Diff Modal (filtered for current contract) */}
          <button
            type="button"
            onClick={handleOpenSnapshotsHistoryModal}
            className="px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 whitespace-nowrap shrink-0 transition-colors"
            title="عرض سجل اللقطات الزمنية والمقارنة الخاصة بهذا العقد"
          >
            <History className="w-3.5 h-3.5 text-blue-900" />
            <span className="hidden lg:inline">
              اللقطات ({revisions.filter((r) => r.documentId === docId).length})
            </span>
          </button>

          {/* Print */}
          <button
            type="button"
            onClick={() => window.print()}
            className="px-2.5 py-1.5 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-lg text-xs font-medium inline-flex items-center gap-1.5 shrink-0 transition-colors"
            title="طباعة العقد A4"
          >
            <Printer className="w-3.5 h-3.5" />
            <span className="hidden xl:inline">طباعة</span>
          </button>

          {/* Export Word (.docx) - Primary CTA */}
          <button
            type="button"
            onClick={handleExportCurrentToWord}
            className="px-3 py-1.5 text-xs font-semibold text-white bg-blue-900 rounded-lg hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 inline-flex items-center gap-1.5 shadow-2xs"
            title="تصدير العقد التوثيقي كاملاً إلى ملف Word (.docx)"
          >
            <Download className="w-4 h-4" />
            <span>تصدير Word</span>
          </button>

          {/* Editor Tour Guide */}
          <button
            type="button"
            onClick={() => setShowOnboardingTour(true)}
            className="p-1.5 text-slate-500 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-colors shrink-0"
            title="فتح الجولة الإرشادية التفاعلية للمحرر"
          >
            <HelpCircle className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Two-Row Word Formatting & Notary Productivity Ribbon */}
      <EditorRibbon
        toolbarState={toolbarState}
        canUndo={undoStack.length > 0}
        canRedo={redoStack.length > 0}
        zoom={zoom}
        showHeaderFooter={showHeaderFooter}
        pageNumberingEnabled={pageNumberingEnabled}
        showFindReplace={showFindReplace}
        unfilledCount={unfilledCount}
        totalVariablesCount={extractedPlaceholders.length}
        partyFields={partyFields}
        extractedPlaceholders={extractedPlaceholders}
        fieldValues={fieldValues}
        derivedTemplates={derivedTemplates}
        activeEditingDerivedId={editingDerivedTpl?.id || null}
        onOpenSmartVariablesModal={() => {
          decorateEditorZonesPreservingSelection();
          setFocusedVarKey(null);
          setShowSmartVarsModal(true);
        }}
        onConvertSelectionToSmartTag={handleConvertSelectionToSmartTag}
        onInsertSmartTagAtCaret={handleInsertSmartTagAtCaret}
        onInsertNewClauseHeadingAtCaret={handleInsertNewClauseHeadingAtCaret}
        onOpenVersionDiffModal={handleOpenVersionDiffModal}
        onGenerateDerivedDoc={handleGenerateDerivedDoc}
        onEditDerivedTemplateInEditor={handleEditDerivedTemplateInEditor}
        onSaveCurrentAsDerivedTemplate={async (name, code, description) => {
          const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
          const newDt: DerivedDocTemplate = {
            id: `custom_derived_${Date.now()}`,
            code,
            name,
            description,
            bodyHtml,
            headerHtml,
            footerHtml,
            updatedAt: new Date().toISOString(),
          };
          await saveDerivedDocTemplate(newDt);
          setDerivedTemplates(await loadDerivedDocTemplates());
          showToast(`تمت إضافة "${name}" إلى قوالب الوثائق المشتقة`);
        }}
        onImportDocxForDerivedTemplate={handleImportDocxForDerivedTemplate}
        onDeleteDerivedTemplate={async (id) => {
          await deleteDerivedDocTemplate(id);
          setDerivedTemplates(await loadDerivedDocTemplates());
          showToast('تم حذف القالب المشتق');
        }}
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
        onOpenWordTemplatesModal={() => {
          setModalActiveBodyHtml(
            bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH
          );
          setShowWordTemplatesModal(true);
        }}
        onSaveSelectionBookmark={saveSelectionBookmark}
      />

      {/* Contextual Banner when Editing a Derived Document Template or Clause inside the A4 Editor */}
      {(editingDerivedTpl || editingClauseObj) && (
        <div className="bg-amber-100 border-b border-amber-300 px-6 py-2 flex flex-wrap items-center justify-between gap-3 no-print">
          <div className="text-xs font-bold text-amber-950">
            {editingDerivedTpl
              ? `وضع تعديل قالب الوثيقة المشتقة: "${editingDerivedTpl.name}" — يمكنك تعديل النص والجداول ووسوم [...] بحرية تامة.`
              : `وضع تعديل صياغة البند الجاهز: "${editingClauseObj?.title}"`}
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() =>
                editingDerivedTpl
                  ? handleFinishEditingDerivedTemplate(true)
                  : handleFinishEditingClause(true)
              }
              className="px-3 py-1 bg-amber-900 text-white text-xs font-semibold rounded hover:bg-amber-950"
            >
              حفظ التعديلات والعودة للعقد
            </button>
            <button
              type="button"
              onClick={() =>
                editingDerivedTpl
                  ? handleFinishEditingDerivedTemplate(false)
                  : handleFinishEditingClause(false)
              }
              className="px-3 py-1 bg-white border border-amber-300 text-amber-950 text-xs font-medium rounded hover:bg-amber-50"
            >
              إلغاء والعودة للعقد
            </button>
          </div>
        </div>
      )}

      {/* Multi-Node Arabic-Aware Find & Replace Bar */}
      {showFindReplace && (
        <div className="bg-amber-50/95 border-b border-amber-200 px-4 py-2 flex flex-wrap items-center gap-2.5 text-xs no-print">
          <div className="flex items-center gap-1.5 font-semibold text-amber-950 shrink-0">
            <Search className="w-3.5 h-3.5 text-amber-900" />
            <span>بحث واستبدال عربي:</span>
          </div>
          <input
            type="text"
            value={findQuery}
            onChange={(e) => setFindQuery(e.target.value)}
            placeholder="ابحث عن نص (مثال: البائع)..."
            className="px-2.5 py-1 bg-white border border-slate-300 rounded w-44 focus:border-blue-900 focus:outline-none"
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                if (e.shiftKey) handlePrevMatch();
                else handleNextMatch();
              } else if (e.key === 'Escape') {
                setShowFindReplace(false);
              }
            }}
          />
          <input
            type="text"
            value={replaceQuery}
            onChange={(e) => setReplaceQuery(e.target.value)}
            placeholder="استبدال بـ..."
            className="px-2.5 py-1 bg-white border border-slate-300 rounded w-44 focus:border-blue-900 focus:outline-none"
            onKeyDown={(e) => {
              if (e.key === 'Enter') handleRunReplace(false);
              else if (e.key === 'Escape') setShowFindReplace(false);
            }}
          />

          <label className="flex items-center gap-1.5 text-[11px] text-amber-950 bg-white/70 px-2 py-1 rounded border border-amber-200 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={ignoreArabicHamzaAndDiacritics}
              onChange={(e) => setIgnoreArabicHamzaAndDiacritics(e.target.checked)}
              className="rounded border-slate-300 text-blue-900"
            />
            <span>تجاهل الهمزات والتشكيل (أ/إ/آ، ة/ه)</span>
          </label>

          <span className="text-slate-600 tabular-nums text-xs">
            {matchCount > 0
              ? `${activeMatchIdx + 1} من ${matchCount}`
              : '0 نتائج'}
          </span>

          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={handlePrevMatch}
              disabled={matchCount === 0}
              className="p-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40"
              title="التطابق السابق (Shift+Enter)"
            >
              <ChevronUp className="w-3.5 h-3.5" />
            </button>
            <button
              type="button"
              onClick={handleNextMatch}
              disabled={matchCount === 0}
              className="p-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40"
              title="التطابق التالي (Enter)"
            >
              <ChevronDown className="w-3.5 h-3.5" />
            </button>
          </div>

          <button
            type="button"
            onClick={() => handleRunReplace(false)}
            disabled={matchCount === 0}
            className="px-2.5 py-1 bg-white border border-slate-300 rounded hover:bg-slate-100 disabled:opacity-40 inline-flex items-center gap-1"
          >
            <Replace className="w-3.5 h-3.5" />
            <span>استبدال الحالي</span>
          </button>
          <button
            type="button"
            onClick={() => handleRunReplace(true)}
            disabled={matchCount === 0}
            className="px-2.5 py-1 bg-blue-900 text-white rounded hover:bg-blue-800 disabled:opacity-40 font-medium"
            title="استبدال كل المطابقات مع أخذ لقطة أمان تلقائية"
          >
            استبدال الكل ({matchCount})
          </button>
          <button
            type="button"
            onClick={() => setShowFindReplace(false)}
            className="mr-auto p-1 text-slate-500 hover:text-slate-800"
            title="إغلاق (Esc)"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* Toast Notification Banner */}
      {toastMessage && (
        <div className="fixed bottom-12 left-6 z-50 bg-slate-900 text-white px-4 py-2.5 rounded-lg shadow-lg text-xs font-medium flex items-center gap-2 no-print">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* MAIN WORKSPACE: RIGHT SIDEBAR + CENTER A4 CANVAS WITH HORIZONTAL & VERTICAL RULERS */}
      <div
        id="workspace"
        className="flex-1 flex flex-row min-h-0 overflow-hidden print:overflow-visible relative"
      >
        {/* RIGHT SIDEBAR (EXPANDED WORKSPACE VS 48PX SLIM ICON RAIL) */}
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
            onScrollToLiveClause={handleScrollToLiveClause}
            onMoveLiveClauseInDoc={handleMoveLiveClauseInDoc}
            onDeleteLiveClauseFromDoc={handleDeleteLiveClauseFromDoc}
            onRenameLiveClauseInDoc={handleRenameLiveClauseInDoc}
            onInsertNewClauseHeadingInDoc={handleInsertNewClauseHeadingAtCaret}
            clauses={clauses}
            activeClauseIdsInDoc={activeClauseIdsInDoc}
            onToggleClauseInDoc={handleToggleClauseInDoc}
            onInsertClauseAtCaret={handleInsertClauseAtCaret}
            onMoveClauseOrder={handleMoveClauseOrder}
            onSaveNewClause={handleSaveNewClause}
            onSaveSelectionAsClause={handleSaveSelectionAsClause}
            onEditClauseInEditor={handleEditClauseInEditor}
            onDeleteClause={async (id) => {
              await deleteNotaryClause(id);
              setClauses(await loadNotaryClauses());
            }}
            partyFields={partyFields}
            extractedPlaceholders={extractedPlaceholders}
            clauseGroups={clauseGroups}
            unfilledCount={unfilledCount}
            fieldValues={fieldValues}
            fieldInputTypes={fieldInputTypes}
            onOpenSmartVariablesModal={(focusKey) => {
              decorateEditorZonesPreservingSelection();
              setFocusedVarKey(focusKey || null);
              setShowSmartVarsModal(true);
            }}
            onUpdateFieldValue={(key, value) => {
              const next = { ...fieldValuesRef.current, [key]: value };
              fieldValuesRef.current = next;
              setFieldValues(next);
              syncPlaceholdersAndDraft(next);
            }}
            onChangeFieldInputType={handleChangeFieldInputType}
            onAddCustomField={(key, label) => {
              if (partyFields.some((f) => f.key === key)) return;
              const updated: PartyField[] = [
                ...partyFields,
                { key, label, value: '', category: 'custom' },
              ];
              setPartyFields(updated);
              savePartyFields(updated);
              showToast(`تمت إضافة الحقل [${key}]`);
            }}
            onInsertPlaceholderAtCaret={(key) => {
              if (!bodyEditorRef.current) return;
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
            onInsertValueAtCaret={(val) => {
              if (!bodyEditorRef.current || !val) return;
              recordHistorySnapshot();
              insertHtmlAtSelection(
                bodyEditorRef.current,
                escapeHtml(val),
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
              await saveSubdivisionEstate(estate);
              const list = await loadSubdivisionEstates();
              setEstates(list);
              setSelectedEstateId(estate.id);
              showToast(`تم حفظ جدول الوصف التقسيمي "${estate.estateName}"`);
            }}
            onDeleteEstate={async (id) => {
              await deleteSubdivisionEstate(id);
              const list = await loadSubdivisionEstates();
              setEstates(list);
              showToast('تم حذف جدول الوصف التقسيمي');
            }}
            onInsertLotClauseAtCaret={handleInsertLotClauseAtCaret}
            onInsertFullSubdivisionTableAtCaret={
              handleInsertFullSubdivisionTableAtCaret
            }
            savedParties={savedParties}
            savedProperties={savedProperties}
            onRecallPartyToRole={handleRecallPartyToRole}
            onSaveCurrentPartyToDirectory={handleSaveCurrentPartyToDirectory}
            onDeleteSavedParty={handleDeleteSavedParty}
            onRecallPropertyRecord={handleRecallPropertyRecord}
            onSaveCurrentPropertyToDirectory={handleSaveCurrentPropertyToDirectory}
            onDeleteSavedProperty={handleDeleteSavedProperty}
            templates={templates}
            onImportDocxAsTemplate={handleImportDocxAsTemplate}
            onSaveCurrentAsTemplate={handleSaveCurrentAsTemplate}
            onOpenSaveAsTemplateModal={() => setShowSaveAsTemplateModal(true)}
            onExportTemplatesJson={handleExportTemplatesJson}
            onImportTemplatesJson={handleImportTemplatesJson}
            onLoadTemplateFull={(tpl, clearPreviousClauses = true) => {
              if (!bodyEditorRef.current) return;
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
              if (!bodyEditorRef.current) return;
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
            onDeleteTemplate={async (id) => {
              await deleteCustomTemplate(id);
              setTemplates(await loadCustomTemplates());
              showToast('تم حذف القالب من المكتبة');
            }}
            derivedTemplates={derivedTemplates}
            activeEditingDerivedId={editingDerivedTpl?.id || null}
            onGenerateDerivedDoc={handleGenerateDerivedDoc}
            onEditDerivedTemplateInEditor={handleEditDerivedTemplateInEditor}
            onSaveCurrentAsDerivedTemplate={async (name, code, description) => {
              const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
              const newDt: DerivedDocTemplate = {
                id: `custom_derived_${Date.now()}`,
                code,
                name,
                description,
                bodyHtml,
                headerHtml,
                footerHtml,
                updatedAt: new Date().toISOString(),
              };
              await saveDerivedDocTemplate(newDt);
              setDerivedTemplates(await loadDerivedDocTemplates());
              showToast(`تمت إضافة "${name}" إلى قوالب الوثائق المشتقة`);
            }}
            onImportDocxForDerivedTemplate={handleImportDocxForDerivedTemplate}
            onDeleteDerivedTemplate={async (id) => {
              await deleteDerivedDocTemplate(id);
              setDerivedTemplates(await loadDerivedDocTemplates());
              showToast('تم حذف القالب المشتق');
            }}
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
            revisions={revisions}
            downloads={downloads}
            activeDocumentId={docId}
            currentEditorTitle={docTitle}
            onOpenDocument={handleSelectDocument}
            onOpenMultiSourceNewModal={() => setShowMultiSourceModal(true)}
            onSaveCurrentDocument={(targetFolderId?: string | null) =>
              handleSaveCurrentDocumentAndRevision(undefined, targetFolderId)
            }
            onDeleteDocument={async (id) => {
              await deleteDocumentRecord(id);
              setDocuments(await loadSavedDocuments());
              if (activeDocument?.id === id) {
                setActiveDocument(null);
              }
              showToast('تم حذف المستند المحفوظ');
            }}
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
            onOpenVersionDiffModal={handleOpenVersionDiffModal}
            onRestoreRevision={(rev) => {
              if (!bodyEditorRef.current) return;
              recordHistorySnapshot();
              bodyEditorRef.current.innerHTML = rev.bodyHtml;
              normalizeNotaryContainerDOM(bodyEditorRef.current);
              setFieldValues(rev.fieldValues || {});
              syncPlaceholdersAndDraft(rev.fieldValues || {});
              showToast('تم استعادة النسخة المختارة بنجاح');
            }}
            onOpenDownloadArchiveItemInEditor={(item) => {
              if (!bodyEditorRef.current) return;
              recordHistorySnapshot();
              setDocTitle(`${item.documentTitle} (${item.docTypeLabel})`);
              bodyEditorRef.current.innerHTML = item.bodyHtml;
              normalizeNotaryContainerDOM(bodyEditorRef.current);
              setFieldValues(item.fieldValues || {});
              syncPlaceholdersAndDraft(item.fieldValues || {});
              showToast(`تم فتح "${item.docTypeLabel}" في المحرر`);
            }}
            onRedownloadArchiveItemDocx={async (item) => {
              await downloadNotaryDocx({
                title: item.fileName.replace(/\.docx$/i, ''),
                bodyHtml: item.bodyHtml,
                headerHtml: item.headerHtml,
                footerHtml: item.footerHtml,
                pageNumberingEnabled: true,
                fieldValues: item.fieldValues,
              });
              showToast(`تم إعادة تحميل "${item.fileName}"`);
            }}
            onDeleteDownloadArchiveItem={async (id) => {
              await deleteDownloadArchiveItem(id);
              setDownloads(await loadDownloadArchive());
            }}
            onExportBackupJson={async () => {
              const bundle = await exportFullBackupBundle();
              const blob = new Blob([JSON.stringify(bundle, null, 2)], {
                type: 'application/json',
              });
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = `نسخة_احتياطية_الموثق_${new Date()
                .toISOString()
                .slice(0, 10)}.json`;
              document.body.appendChild(a);
              a.click();
              document.body.removeChild(a);
              URL.revokeObjectURL(url);
              showToast('تم تصدير النسخة الاحتياطية الشاملة للمكتب (JSON)');
            }}
            onImportBackupJson={async (file) => {
              try {
                const text = await file.text();
                const parsed = JSON.parse(text);
                const counts = await importFullBackupBundle(parsed);
                setTemplates(await loadCustomTemplates());
                setDocuments(await loadSavedDocuments());
                setEstates(await loadSubdivisionEstates());
                setClauses(await loadNotaryClauses());
                setDerivedTemplates(await loadDerivedDocTemplates());
                setRevisions(await loadDocumentRevisions());
                setDownloads(await loadDownloadArchive());
                setPartyFields(loadPartyFields());
                showToast(
                  `تم استعادة (${counts.templatesCount}) قالب و (${counts.clausesCount}) بند و (${counts.documentsCount}) عقد`
                );
              } catch {
                showToast('تعذر قراءة ملف النسخة الاحتياطية');
              }
            }}
            onSaveSelectionBookmark={saveSelectionBookmark}
          />
        </div>
      ) : (
        /* SLIM ACTIVITY RAIL (الشريط الجانبي يبدأ مطوياً كعمود أيقونات نحيف للأقسام الأربعة ويفتح عند اختيار قسم) */
        <aside className="flex w-14 bg-white border-l border-slate-200 flex-col items-center py-2.5 gap-2 h-full min-h-0 shrink-0 select-none no-print transition-all duration-150 shadow-2xs">
          {/* 1. Clauses Section */}
          <button
            type="button"
            onClick={() => {
              setSidebarTab('clauses');
              setIsSidebarOpen(true);
            }}
            className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-slate-600 hover:bg-blue-50 hover:text-blue-900"
            title="فتح قسم البنود الجاهزة"
          >
            <ListChecks className="w-4 h-4" />
            <span className="text-[10px] font-medium">البنود</span>
          </button>

          {/* 2. Parties Section */}
          <button
            type="button"
            onClick={() => {
              setSidebarTab('parties');
              setIsSidebarOpen(true);
            }}
            className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-slate-600 hover:bg-blue-50 hover:text-blue-900 relative"
            title="فتح قسم الأطراف والتعيينات"
          >
            <UserCheck className="w-4 h-4" />
            <span className="text-[10px] font-medium">الأطراف</span>
            {unfilledCount > 0 && (
              <span className="absolute top-1.5 left-1.5 w-2 h-2 rounded-full bg-red-600" />
            )}
          </button>

          {/* 3. Contract Templates Section */}
          <button
            type="button"
            onClick={() => {
              setSidebarTab('templates');
              setIsSidebarOpen(true);
            }}
            className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-slate-600 hover:bg-blue-50 hover:text-blue-900"
            title="فتح قسم قوالب عقود المكتب"
          >
            <FileText className="w-4 h-4" />
            <span className="text-[10px] font-medium">القوالب</span>
          </button>

          {/* 4. Archive & Saved Documents Section */}
          <button
            type="button"
            onClick={() => {
              setSidebarTab('documents');
              setIsSidebarOpen(true);
            }}
            className="w-11 py-2 rounded-lg transition-colors flex flex-col items-center justify-center gap-1 text-slate-600 hover:bg-blue-50 hover:text-blue-900"
            title="فتح قسم المستندات والأرشيف"
          >
            <FolderArchive className="w-4 h-4" />
            <span className="text-[10px] font-medium">الأرشيف</span>
          </button>
        </aside>
      )}

        {/* CENTER A4 CANVAS AREA WITH HORIZONTAL & VERTICAL RULERS */}
        <main className="flex-1 min-w-0 min-h-0 overflow-auto print:overflow-visible p-4 sm:p-8 flex flex-col items-center bg-slate-200/80">
          {/* Top Horizontal A4 Margin Ruler Indicator (21cm total: 7cm Right | 12cm Content | 2cm Left) */}
          <div
            className="mb-2 bg-white border border-slate-300 rounded-t shadow-2xs text-[10px] font-mono text-slate-500 flex items-center select-none overflow-hidden no-print"
            style={{
              width: '210mm',
              transform: `scale(${zoom / 100})`,
              transformOrigin: 'top center',
            }}
            dir="rtl"
          >
            <div
              style={{ width: '70mm' }}
              className="bg-slate-100 border-l border-dashed border-slate-400 py-1 text-center text-slate-500 tabular-nums"
              title="هامش أيمن توثيقي ثابت: 7 سنتيمتر"
            >
              هامش يمين: 7 سم
            </div>
            <div
              style={{ width: '120mm' }}
              className="py-1 text-center font-semibold text-blue-950 tabular-nums flex items-center justify-between px-2"
            >
              <span>0</span>
              <span>المساحة النصية الصافية (12 سم · Arial 13pt · تباعد 1.0)</span>
              <span>12</span>
            </div>
            <div
              style={{ width: '20mm' }}
              className="bg-slate-100 border-r border-dashed border-slate-400 py-1 text-center text-slate-500 tabular-nums"
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
                    contentEditable={!previewMergedMode}
                    suppressContentEditableWarning
                    dir="rtl"
                    onInput={() => syncPlaceholdersAndDraft()}
                    onKeyUp={(e) => {
                      if (e.key === ']' || e.key === '}') decorateEditorZonesPreservingSelection();
                    }}
                    onBlur={decorateEditorZonesPreservingSelection}
                    onKeyDown={handleEditorKeyDown}
                    onPaste={handleEditorPaste}
                    className="notary-editor-zone outline-none min-h-[18px]"
                  />
                </div>
              )}

              {/* MAIN CONTRACT BODY EDITABLE ZONE */}
              {previewMergedMode ? (
                <div
                  dir="rtl"
                  className="notary-editor-zone relative z-10 min-h-[210mm]"
                  dangerouslySetInnerHTML={{
                    __html: mergedPreviewHtml,
                  }}
                />
              ) : (
                <div
                  ref={bodyEditorRef}
                  contentEditable
                  suppressContentEditableWarning
                  dir="rtl"
                  onInput={() => syncPlaceholdersAndDraft()}
                  onKeyUp={(e) => {
                    if (e.key === ']' || e.key === '}') decorateEditorZonesPreservingSelection();
                  }}
                  onBlur={decorateEditorZonesPreservingSelection}
                  onKeyDown={handleEditorKeyDown}
                  onPaste={handleEditorPaste}
                  onMouseDown={handleEditorMouseDown}
                  onDoubleClick={handleEditorDoubleClick}
                  className="notary-editor-zone relative z-10 outline-none min-h-[210mm]"
                />
              )}

              {/* Optional Document Footer & Page Numbering Preview */}
              {(showHeaderFooter || pageNumberingEnabled) && (
                <div className="mt-4 pt-1 border-t border-dashed border-slate-300 relative z-10">
                  {showHeaderFooter && (
                    <>
                      <div className="text-[9px] text-slate-400 mb-0.5 select-none no-print">
                        تذييل العقد (Footer)
                      </div>
                      <div
                        ref={footerEditorRef}
                        contentEditable={!previewMergedMode}
                        suppressContentEditableWarning
                        dir="rtl"
                        onInput={() => syncPlaceholdersAndDraft()}
                        onKeyUp={(e) => {
                          if (e.key === ']' || e.key === '}') decorateEditorZonesPreservingSelection();
                        }}
                        onBlur={decorateEditorZonesPreservingSelection}
                        onKeyDown={handleEditorKeyDown}
                        onPaste={handleEditorPaste}
                        className="notary-editor-zone outline-none min-h-[18px]"
                      />
                    </>
                  )}
                  {pageNumberingEnabled && (
                    <div
                      dir="rtl"
                      className="text-center text-slate-600 select-none mt-1"
                      style={{
                        fontFamily: STRICT_FONT_FAMILY,
                        fontSize: `${STRICT_FONT_SIZE_PT}pt`,
                        lineHeight: 1,
                      }}
                    >
                      الصفحة 1 من {docMetrics.estimatedPages}
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </main>
      </div>

      {/* BOTTOM STATUS BAR (شريط الحالة التوثيقي) */}
      <footer className="bg-white border-t border-slate-200 px-4 py-1.5 text-[11px] text-slate-600 flex flex-wrap items-center justify-between gap-4 select-none no-print">
        <div className="flex items-center gap-3">
          <span className="inline-flex items-center gap-1.5 font-medium text-slate-800">
            <span
              className={`w-2 h-2 rounded-full ${
                autoSaveState === 'saved' ? 'bg-emerald-500' : 'bg-amber-500 animate-pulse'
              }`}
            />
            {autoSaveState === 'saved'
              ? 'تم الحفظ تلقائياً في قاعدة المكتب'
              : 'جارٍ الحفظ التلقائي...'}
          </span>

          <span className="text-slate-300">|</span>

          <button
            type="button"
            onClick={() => {
              setFocusedVarKey(null);
              setShowSmartVarsModal(true);
            }}
            className={`font-semibold hover:underline ${
              unfilledCount > 0 ? 'text-red-700' : 'text-emerald-700'
            }`}
          >
            {unfilledCount > 0
              ? `المتغيرات غير المعبأة: ${unfilledCount} (اضغط للتعبئة)`
              : `جميع متغيرات العقد مكتملة (${extractedPlaceholders.length})`}
          </button>
        </div>

        <div className="flex items-center gap-3 tabular-nums">
          <span>الكلمات: {docMetrics.wordCount}</span>
          <span className="text-slate-300">·</span>
          <span>الأحرف: {docMetrics.charCount}</span>
          <span className="text-slate-300">·</span>
          <span>الصفحات المقدرة: {docMetrics.estimatedPages}</span>
          <span className="text-slate-300">·</span>
          <span className="font-semibold text-slate-800">
            Arial 13pt · تباعد 1.0 · هوامش 7/2/1/6 سم
          </span>
        </div>
      </footer>

      {/* MODALS */}
      <SmartVariablesModal
        isOpen={showSmartVarsModal}
        onClose={() => setShowSmartVarsModal(false)}
        focusedVarKey={focusedVarKey}
        extractedPlaceholders={extractedPlaceholders}
        clauseGroups={clauseGroups}
        partyFields={partyFields}
        fieldValues={fieldValues}
        fieldInputTypes={fieldInputTypes}
        estates={estates}
        selectedEstateId={selectedEstateId}
        selectedLotNumber={selectedLotNumber}
        savedParties={savedParties}
        savedProperties={savedProperties}
        onSelectEstateAndLot={handleSelectEstateAndLot}
        onRecallPartyToRole={handleRecallPartyToRole}
        onSaveCurrentPartyToDirectory={handleSaveCurrentPartyToDirectory}
        onDeleteSavedParty={handleDeleteSavedParty}
        onRecallPropertyRecord={handleRecallPropertyRecord}
        onSaveCurrentPropertyToDirectory={handleSaveCurrentPropertyToDirectory}
        onDeleteSavedProperty={handleDeleteSavedProperty}
        onUpdateFieldValue={(key, value) => {
          const next = { ...fieldValuesRef.current, [key]: value };
          fieldValuesRef.current = next;
          setFieldValues(next);
          syncPlaceholdersAndDraft(next);
        }}
        onChangeFieldInputType={handleChangeFieldInputType}
        onBakeAllIntoDocument={handleBakeAllPlaceholdersIntoDocument}
      />

      <VersionDiffModal
        isOpen={showVersionDiffModal}
        onClose={() => {
          setShowVersionDiffModal(false);
          setDiffComparisonRevision(null);
        }}
        revisions={(() => {
          const contractRevs = revisions.filter((r) => r.documentId === docId);
          return diffComparisonRevision
            ? [
                diffComparisonRevision,
                ...contractRevs.filter((r) => r.id !== diffComparisonRevision.id),
              ]
            : contractRevs;
        })()}
        currentBodyHtml={modalActiveBodyHtml}
        currentTitle={docTitle}
        onRestoreRevision={handleRestoreSnapshot}
      />

      <DocxImportPreviewModal
        isOpen={showDocxPreviewModal}
        onClose={() => {
          setShowDocxPreviewModal(false);
          setPendingDocxImport(null);
        }}
        importResult={pendingDocxImport}
        onConfirmApply={handleApplyDocxImport}
      />

      <SaveAsTemplateModal
        isOpen={showSaveAsTemplateModal}
        onClose={() => setShowSaveAsTemplateModal(false)}
        currentDocTitle={docTitle}
        extractedPlaceholders={extractedPlaceholders}
        currentFieldValues={fieldValues}
        onConfirmSave={handleSaveCustomTemplateAdvanced}
      />

      <SnapshotsHistoryModal
        isOpen={showSnapshotsHistoryModal}
        onClose={() => setShowSnapshotsHistoryModal(false)}
        revisions={revisions.filter((r) => r.documentId === docId)}
        currentBodyHtml={modalActiveBodyHtml}
        currentTitle={docTitle}
        onTakeManualSnapshot={handleTakeManualSnapshot}
        onRestoreSnapshot={handleRestoreSnapshot}
        onDeleteSnapshot={handleDeleteSnapshot}
        onOpenDiffComparison={handleOpenDiffComparison}
      />

      <OnboardingTourModal
        isOpen={showOnboardingTour}
        onClose={() => setShowOnboardingTour(false)}
        onFinishTour={() => {
          if (typeof window !== 'undefined') {
            localStorage.setItem('notary_editor_tour_v27', 'true');
          }
          setShowOnboardingTour(false);
        }}
      />

      <MultiSourceStartModal
        isOpen={showMultiSourceModal}
        onClose={() => setShowMultiSourceModal(false)}
        templates={templates}
        documents={documents}
        downloads={downloads}
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
        onSelectTemplate={(tpl, mode, clearPreviousClauses) => {
          if (!bodyEditorRef.current) return;
          if (mode === 'insert') {
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
        onSelectSavedDoc={(doc) => {
          handleSelectDocument(doc);
          showToast(`تم فتح العقد "${doc.title}" بجميع خصائصه ومتغيراته المحفوظة`);
        }}
        onImportDocxFile={handleImportDocxDirectlyToEditor}
        onSelectDownloadArchiveItem={(item) => {
          if (!bodyEditorRef.current) return;
          setUndoStack([]);
          setRedoStack([]);
          setDocId(`doc_${Date.now()}`);
          setDocTitle(`${item.documentTitle} (${item.docTypeLabel})`);
          setActiveDocument(null);
          activeDocumentRef.current = null;
          bodyEditorRef.current.innerHTML = item.bodyHtml;
          normalizeNotaryContainerDOM(bodyEditorRef.current);
          const nextVals = { ...(item.fieldValues || {}) };
          fieldValuesRef.current = nextVals;
          fieldInputTypesRef.current = {};
          setFieldValues(nextVals);
          setFieldInputTypes({});
          syncPlaceholdersAndDraft(nextVals, {});
          showToast(`تم استيراد "${item.docTypeLabel}" من أرشيف التحميلات للمحرر`);
        }}
      />

      <WordTemplatesModal
        isOpen={showWordTemplatesModal}
        onClose={() => setShowWordTemplatesModal(false)}
        wordTemplates={wordTemplates}
        onSaveWordTemplate={async (tpl) => {
          await saveWordTemplate(tpl);
          setWordTemplates(await loadWordTemplates());
          showToast('تم حفظ قالب Word بنجاح');
        }}
        onDeleteWordTemplate={async (id) => {
          await deleteWordTemplate(id);
          setWordTemplates(await loadWordTemplates());
          showToast('تم حذف قالب Word');
        }}
        contractData={{
          officeName: activeClerk.name,
          officeAddr: 'مكتب التوثيق الرسمي',
          typeActe: docTitle,
          client1:
            fieldValues['الطرف_الأول_الاسم'] ||
            fieldValues['البائع'] ||
            fieldValues['الموكل'] ||
            '',
          client2:
            fieldValues['الطرف_الثاني_الاسم'] ||
            fieldValues['المشتري'] ||
            fieldValues['الوكيل'] ||
            '',
          dateActe: fieldValues['تاريخ_العقد'] || new Date().toLocaleDateString('ar-DZ'),
          clauses: (() => {
            const parsedFromSheet = parseContractIntoClauses(
              modalActiveBodyHtml,
              fieldValues
            );
            if (parsedFromSheet.length > 0) {
              return parsedFromSheet.map((c) => ({
                title: c.title,
                contentHtml: c.bodyHtml,
              }));
            }
            if (outlineClauses.length > 0) {
              return outlineClauses.map((c) => ({
                title: c.title,
                contentHtml: c.contentHtml,
              }));
            }
            return [];
          })(),
          rawContractBodyHtml: modalActiveBodyHtml,
          fieldValues,
        }}
      />
    </div>
  );
}
