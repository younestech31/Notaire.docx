'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  ContractFolder,
  ContractOutlineClause,
  DocumentRevision,
  NotaryClerk,
  OfficeTemplate,
  SavedContractDerivedDoc,
  SavedDocument,
} from '@/lib/types';
import {
  deleteContractFolder,
  deleteDocumentRecord,
  deleteDocumentRevision,
  loadActiveDraftSession,
  loadContractFolders,
  loadDocumentRevisions,
  loadSavedDocuments,
  saveActiveDraftSession,
  saveContractFolder,
  saveDocumentRecord,
  saveDocumentRevision,
} from '@/lib/storage';
import {
  exportEditorBodyToDocx,
  generateAndDownloadDerivedDoc,
} from '@/lib/document-io-workflows';
import { downloadNotaryDocx, mergePlaceholdersIntoHtml } from '@/lib/docx-engine';
import { isContractOfficiallyIndexed } from '@/lib/indexing-engine';
import type { DocumentSessionSnapshot } from './useDocumentSession';

export function useNotaryLifecycle(
  activeClerk: NotaryClerk,
  showToast: (msg: string) => void
) {
  const [documents, setDocumentsState] = useState<SavedDocument[]>([]);
  const documentsRef = useRef<SavedDocument[]>([]);
  const [revisions, setRevisions] = useState<DocumentRevision[]>([]);
  const [activeDocument, setActiveDocument] = useState<SavedDocument | null>(
    null
  );
  const activeDocumentRef = useRef<SavedDocument | null>(null);
  useEffect(() => {
    activeDocumentRef.current = activeDocument;
  }, [activeDocument]);

  const [activeDerivedDoc, setActiveDerivedDoc] =
    useState<SavedContractDerivedDoc | null>(null);
  const [unlockedDocId, setUnlockedDocId] = useState<string | null>(null);

  const setDocuments = useCallback((nextDocs: SavedDocument[]) => {
    documentsRef.current = nextDocs;
    setDocumentsState(nextDocs);
  }, []);

  const isContractLocked = useCallback(
    (
      docId: string,
      isEditingSubEntity: boolean
    ): { unlockIndexedBody: boolean; isBodyLocked: boolean } => {
      const unlockIndexedBody = unlockedDocId === docId;
      const isBodyLocked =
        isContractOfficiallyIndexed(activeDocument) &&
        !unlockIndexedBody &&
        !isEditingSubEntity;
      return { unlockIndexedBody, isBodyLocked };
    },
    [activeDocument, unlockedDocId]
  );

  const setUnlockIndexedBodyForDoc = useCallback(
    (docId: string, value: React.SetStateAction<boolean>) => {
      setUnlockedDocId((prev) => {
        const current = prev === docId;
        const next = typeof value === 'function' ? value(current) : value;
        return next ? docId : null;
      });
    },
    []
  );

  const reloadLifecycleRecords = useCallback(async () => {
    const [loadedDocs, loadedRevs, activeDraft] = await Promise.all([
      loadSavedDocuments(),
      loadDocumentRevisions(),
      loadActiveDraftSession(),
    ]);
    setDocuments(loadedDocs);
    setRevisions(loadedRevs);
    return {
      loadedDocs,
      loadedRevs,
      activeDraft,
    };
  }, [setDocuments]);

  const handleMoveDocument = useCallback(
    async (
      targetDocId: string,
      targetFolderId: string | null,
      folders: ContractFolder[]
    ) => {
      const d = documentsRef.current.find((x) => x.id === targetDocId);
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
    [setDocuments, showToast]
  );

  const handleDeleteFolderAndMigrateDocs = useCallback(
    async (
      folderId: string,
      folders: ContractFolder[],
      setFolders: (next: ContractFolder[]) => void
    ) => {
      const targetFolder = folders.find((f) => f.id === folderId);
      const fallbackParentId = targetFolder?.parentId ?? null;
      await deleteContractFolder(folderId);

      const childFolders = folders.filter((f) => f.parentId === folderId);
      for (const cf of childFolders) {
        await saveContractFolder({
          ...cf,
          parentId: fallbackParentId,
          updatedAt: new Date().toISOString(),
        });
      }

      const docsInFolder = documentsRef.current.filter(
        (d) => d.folderId === folderId
      );
      for (const d of docsInFolder) {
        const updatedDoc = {
          ...d,
          folderId: fallbackParentId,
          updatedAt: new Date().toISOString(),
        };
        await saveDocumentRecord(updatedDoc);
        if (activeDocumentRef.current?.id === d.id) {
          setActiveDocument(updatedDoc);
        }
      }
      setFolders(await loadContractFolders());
      setDocuments(await loadSavedDocuments());
      showToast('تم حذف المجلد ونقل محتوياته بأمان');
    },
    [setDocuments, showToast]
  );

  const handleUpdateDocumentMeta = useCallback(
    async (updated: SavedDocument) => {
      await saveDocumentRecord(updated);
      setDocuments(await loadSavedDocuments());
      setActiveDocument((prev) => (prev?.id === updated.id ? updated : prev));
      showToast('تم تحديث بيانات العقد');
    },
    [setDocuments, showToast]
  );

  const handleUpdateContractLifecycle = useCallback(
    async (targetDocId: string, patch: Partial<SavedDocument>) => {
      const target =
        documentsRef.current.find((d) => d.id === targetDocId) ||
        activeDocumentRef.current;
      if (!target) return;
      const updated: SavedDocument = {
        ...target,
        ...patch,
        updatedAt: new Date().toISOString(),
      };
      await saveDocumentRecord(updated);
      await saveActiveDraftSession(updated);
      const refreshed = await loadSavedDocuments();
      setDocuments(refreshed);
      if (activeDocumentRef.current?.id === updated.id) {
        setActiveDocument(updated);
      }
      if (patch.status || patch.indexNumber || patch.sealedOriginalSha256) {
        const isSealEvent = Boolean(
          patch.indexNumber || patch.sealedOriginalSha256
        );
        const rev: DocumentRevision = {
          id: `rev_lifecycle_${Date.now()}`,
          documentId: updated.id,
          documentTitle: updated.title,
          eventType: isSealEvent ? 'seal' : 'lifecycle',
          bodyHtml: isSealEvent ? updated.bodyHtml : undefined,
          author: activeClerk.name,
          fieldValues: updated.fieldValues || {},
          summary: isSealEvent
            ? `ترسيم في الفهرس الرسمي رقم ${updated.indexNumber || ''} وختم الأصل ببصمة SHA-256`
            : `تحديث حالة العقد التوثيقية (${updated.status})`,
          createdAt: new Date().toISOString(),
        };
        await saveDocumentRevision(rev);
        setRevisions(await loadDocumentRevisions());
      }
    },
    [activeClerk.name, setDocuments]
  );

  const handleSaveDocumentAndRevision = useCallback(
    async (params: {
      snapshot: DocumentSessionSnapshot;
      liveOutline: ContractOutlineClause[];
      folders: ContractFolder[];
      summaryLabel?: string;
      explicitFolderId?: string | null;
    }): Promise<SavedDocument> => {
      const {
        snapshot,
        liveOutline,
        folders,
        summaryLabel,
        explicitFolderId,
      } = params;
      const nowIso = new Date().toISOString();
      const existingDoc =
        documentsRef.current.find((d) => d.id === snapshot.docId) ||
        (activeDocumentRef.current?.id === snapshot.docId
          ? activeDocumentRef.current
          : undefined);
      const resolvedFolderId =
        explicitFolderId !== undefined
          ? explicitFolderId
          : existingDoc?.folderId ?? null;

      const activeVals = snapshot.fieldValues;
      const inferredClient =
        existingDoc?.clientName ||
        snapshot.partyCards[0]?.fullName ||
        activeVals['الطرف_الأول_الاسم'] ||
        activeVals['البائع'] ||
        activeVals['المؤجر'] ||
        activeVals['الموكل'] ||
        '';

      const docRecord: SavedDocument = {
        ...(existingDoc || {}),
        id: snapshot.docId,
        title: snapshot.docTitle || 'عقد بدون عنوان',
        bodyHtml: snapshot.bodyHtml,
        headerHtml: snapshot.headerHtml,
        footerHtml: snapshot.footerHtml,
        pageNumberingEnabled: snapshot.pageNumberingEnabled,
        showHeaderFooter: snapshot.showHeaderFooter,
        fieldValues: activeVals,
        fieldInputTypes: snapshot.fieldInputTypes,
        outlineClauses: liveOutline,
        contractConditions: snapshot.contractConditions,
        partyCards: snapshot.partyCards,
        selectedEstateId: snapshot.selectedEstateId,
        selectedLotNumber: snapshot.selectedLotNumber,
        folderId: resolvedFolderId,
        clerkId: existingDoc?.clerkId || activeClerk.id,
        clerkName: existingDoc?.clerkName || activeClerk.name,
        contractNumber:
          existingDoc?.contractNumber ||
          activeVals['رقم_الفهرس'] ||
          `2026/${Math.floor(100 + Math.random() * 900)}`,
        year: existingDoc?.year || new Date().getFullYear(),
        clientName: inferredClient,
        modelId: existingDoc?.modelId,
        modelTitle: existingDoc?.modelTitle,
        status: existingDoc?.status || 'editing',
        derivedDocuments: existingDoc?.derivedDocuments || [],
        updatedAt: nowIso,
        createdAt: existingDoc?.createdAt || nowIso,
      };
      await saveActiveDraftSession(docRecord);
      await saveDocumentRecord(docRecord);
      const refreshedDocs = await loadSavedDocuments();
      setDocuments(refreshedDocs);
      setActiveDocument(docRecord);
      activeDocumentRef.current = docRecord;

      const rev: DocumentRevision = {
        id: `rev_${Date.now()}`,
        documentId: snapshot.docId,
        documentTitle: snapshot.docTitle || 'عقد توثيقي',
        author: activeClerk?.name || 'كاتب المكتب',
        bodyHtml: snapshot.bodyHtml,
        fieldValues: { ...activeVals },
        createdAt: nowIso,
        summary: summaryLabel || 'حفظ نسخة مرجعية للعقد',
        eventType: 'save',
      };
      await saveDocumentRevision(rev);
      setRevisions(await loadDocumentRevisions());

      if (explicitFolderId !== undefined) {
        const folderName = explicitFolderId
          ? folders.find((f) => f.id === explicitFolderId)?.name ||
            'المجلد المختار'
          : 'الجذر الرئيسي';
        showToast(
          `تم حفظ العقد "${docRecord.title}" داخل "${folderName}" بنجاح`
        );
      } else {
        showToast(
          summaryLabel || 'تم حفظ العقد في شجرة المكتب وتسجيل نسخة مرجعية'
        );
      }
      return docRecord;
    },
    [activeClerk, setDocuments, showToast]
  );

  const handleDeleteSnapshot = useCallback(
    async (id: string) => {
      await deleteDocumentRevision(id);
      setRevisions(await loadDocumentRevisions());
      showToast('تم حذف النسخة من السجل');
    },
    [showToast]
  );

  const handleDeleteDocument = useCallback(
    async (id: string) => {
      await deleteDocumentRecord(id);
      setDocuments(await loadSavedDocuments());
      if (activeDocumentRef.current?.id === id) {
        setActiveDocument(null);
      }
      showToast('تم حذف المستند المحفوظ');
    },
    [setDocuments, showToast]
  );

  const handleExportSavedDocumentDocx = useCallback(
    async (doc: SavedDocument) => {
      const mergedBody = mergePlaceholdersIntoHtml(
        doc.bodyHtml,
        doc.fieldValues
      );
      const mergedHeader = mergePlaceholdersIntoHtml(
        doc.headerHtml,
        doc.fieldValues
      );
      const mergedFooter = mergePlaceholdersIntoHtml(
        doc.footerHtml,
        doc.fieldValues
      );
      await downloadNotaryDocx({
        title: doc.title,
        bodyHtml: mergedBody,
        headerHtml: doc.showHeaderFooter ? mergedHeader : '',
        footerHtml: doc.showHeaderFooter ? mergedFooter : '',
        pageNumberingEnabled: doc.pageNumberingEnabled,
        fieldValues: doc.fieldValues,
      });
      showToast(`تم تصدير «${doc.title}» بصيغة Word (.docx)`);
    },
    [showToast]
  );

  const handleExportContractWord = useCallback(
    async (snapshot: DocumentSessionSnapshot) => {
      try {
        const { fileName, revisions: nextRevs } = await exportEditorBodyToDocx({
          docId: snapshot.docId,
          docTitle: snapshot.docTitle,
          bodyHtml: snapshot.bodyHtml,
          headerHtml: snapshot.headerHtml,
          footerHtml: snapshot.footerHtml,
          showHeaderFooter: snapshot.showHeaderFooter,
          pageNumberingEnabled: snapshot.pageNumberingEnabled,
          fieldValues: snapshot.fieldValues,
          authorName: activeClerk.name,
        });
        setRevisions(nextRevs);
        showToast(
          `تم تصدير وتنزيل ملف Word (${fileName}) وتسجيل العملية في سجل مراجعات العقد`
        );
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'حدث خطأ أثناء تصدير ملف Word'
        );
      }
    },
    [activeClerk.name, showToast]
  );

  const handleGenerateDerivedDoc = useCallback(
    async (tpl: OfficeTemplate, snapshot: DocumentSessionSnapshot) => {
      try {
        const { missingVarsCount, revisions: nextRevs } =
          await generateAndDownloadDerivedDoc({
            docId: snapshot.docId,
            docTitle: snapshot.docTitle,
            tpl,
            fieldValues: snapshot.fieldValues,
            authorName: activeClerk.name,
          });
        setRevisions(nextRevs);
        if (missingVarsCount > 0) {
          showToast(
            `تنبيه: تم توليد "${tpl.name}" مع وجود (${missingVarsCount}) متغير غير معبأ بعد`
          );
        } else {
          showToast(
            `تم توليد وتنزيل "${tpl.name}" (.docx) وتسجيله في سجل مراجعات العقد`
          );
        }
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'تعذر توليد الوثيقة المشتقة'
        );
      }
    },
    [activeClerk.name, showToast]
  );

  return {
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
    unlockedDocId,
    setUnlockedDocId,
    isContractLocked,
    setUnlockIndexedBodyForDoc,
    reloadLifecycleRecords,
    handleMoveDocument,
    handleDeleteFolderAndMigrateDocs,
    handleUpdateDocumentMeta,
    handleUpdateContractLifecycle,
    handleSaveDocumentAndRevision,
    handleDeleteSnapshot,
    handleDeleteDocument,
    handleExportSavedDocumentDocx,
    handleExportContractWord,
    handleGenerateDerivedDoc,
  };
}
