'use client';

import { useCallback, useState } from 'react';
import {
  ContractFolder,
  ContractPartyCard,
  NotaryClerk,
  NotaryClause,
  OfficeTemplate,
  OfficeTemplateKind,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
} from '@/lib/types';
import {
  DEFAULT_OFFICE_TEMPLATES,
  deleteContractFolder,
  deleteNotaryClause,
  deleteOfficeTemplate,
  deletePartyRecord,
  deletePropertyRecord,
  deleteSubdivisionEstate,
  loadContractFolders,
  loadNotaryClauses,
  loadNotaryClerks,
  loadOfficeTemplates,
  loadSavedParties,
  loadSavedProperties,
  loadSubdivisionEstates,
  saveContractFolder,
  saveNotaryClause,
  saveOfficeTemplate,
  savePartyRecord,
  savePropertyRecord,
  saveSubdivisionEstate,
} from '@/lib/storage';
import { extractPlaceholdersFromHtml } from '@/lib/editor-utils';
import {
  downloadTemplatesJsonFile,
  importDocxFilesAsTemplates,
  importTemplatesJsonFile,
} from '@/lib/document-io-workflows';
import {
  type LocalOfficeDirectoryState,
  reconnectPersistedOfficeDirectory,
} from '@/lib/desktop-bridge';
import type { DocumentSessionSnapshot } from './useDocumentSession';

export function useOfficeCatalog(showToast: (msg: string) => void) {
  const [templates, setTemplates] = useState<OfficeTemplate[]>(
    DEFAULT_OFFICE_TEMPLATES
  );
  const [clauses, setClauses] = useState<NotaryClause[]>([]);
  const [estates, setEstates] = useState<SubdivisionEstate[]>([]);
  const [savedParties, setSavedParties] = useState<SavedPartyRecord[]>([]);
  const [savedProperties, setSavedProperties] = useState<SavedPropertyRecord[]>(
    []
  );
  const [clerks, setClerks] = useState<NotaryClerk[]>([]);
  const [activeClerk, setActiveClerk] = useState<NotaryClerk>({
    id: 'notary-head',
    name: 'الموثق الرئيسي (الأستاذ)',
    role: 'notary',
    color: '#1e3a8a',
    createdAt: '2026-01-01T00:00:00.000Z',
  });
  const [folders, setFolders] = useState<ContractFolder[]>([]);
  const [lastBackupAt, setLastBackupAt] = useState<string | null>(null);
  const [localDirState, setLocalDirState] = useState<LocalOfficeDirectoryState>({
    status: 'disconnected',
    folderName: null,
  });

  const reloadCatalogRecords = useCallback(async () => {
    const [
      loadedOfficeTemplates,
      loadedEstates,
      loadedClauses,
      loadedSavedParties,
      loadedSavedProps,
      loadedClerks,
      loadedFolders,
    ] = await Promise.all([
      loadOfficeTemplates(),
      loadSubdivisionEstates(),
      loadNotaryClauses(),
      loadSavedParties(),
      loadSavedProperties(),
      loadNotaryClerks(),
      loadContractFolders(),
    ]);

    setTemplates(loadedOfficeTemplates);
    setEstates(loadedEstates);
    setClauses(loadedClauses);
    setSavedParties(loadedSavedParties);
    setSavedProperties(loadedSavedProps);
    setClerks(loadedClerks);
    if (loadedClerks.length > 0) {
      setActiveClerk(loadedClerks[0]);
    }
    setFolders(loadedFolders);

    return {
      loadedOfficeTemplates,
      loadedEstates,
      loadedClauses,
      loadedSavedParties,
      loadedSavedProps,
      loadedClerks,
      loadedFolders,
    };
  }, []);

  const connectAndReloadLocalDirectory = useCallback(
    async (onAfterReload?: () => Promise<void>) => {
      try {
        const nextState = await reconnectPersistedOfficeDirectory();
        setLocalDirState(nextState);
        if (nextState.status === 'connected') {
          await reloadCatalogRecords();
          if (onAfterReload) {
            await onAfterReload();
          }
          showToast(
            `تم ربط مجلد المكتب المحلي (${nextState.folderName}) بنجاح — الحفظ والأصول المختومة على القرص مباشرة`
          );
        }
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'تعذر ربط مجلد المكتب المحلي'
        );
      }
    },
    [reloadCatalogRecords, showToast]
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
      showToast(`تم إنشاء المجلد «${name}»`);
    },
    [activeClerk.id, showToast]
  );

  const handleRenameFolder = useCallback(
    async (folderId: string, newName: string) => {
      const target = folders.find((f) => f.id === folderId);
      if (!target) return;
      await saveContractFolder({
        ...target,
        name: newName,
        updatedAt: new Date().toISOString(),
      });
      setFolders(await loadContractFolders());
      showToast(`تم تعديل اسم المجلد إلى «${newName}»`);
    },
    [folders, showToast]
  );

  const handleMoveFolder = useCallback(
    async (folderId: string, targetParentId: string | null) => {
      if (folderId === targetParentId) return;
      const f = folders.find((x) => x.id === folderId);
      if (!f) return;
      if ((f.parentId ?? null) === (targetParentId ?? null)) return;

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

  const handleDeleteFolderOnly = useCallback(
    async (folderId: string) => {
      await deleteContractFolder(folderId);
      setFolders(await loadContractFolders());
      showToast('تم حذف المجلد ونقل عقوده للجذر');
    },
    [showToast]
  );

  const handleSaveEstate = useCallback(
    async (estate: SubdivisionEstate) => {
      await saveSubdivisionEstate(estate);
      setEstates(await loadSubdivisionEstates());
      showToast(`تم حفظ جدول الوصف التقسيمي "${estate.estateName}"`);
    },
    [showToast]
  );

  const handleDeleteEstate = useCallback(
    async (id: string) => {
      await deleteSubdivisionEstate(id);
      setEstates(await loadSubdivisionEstates());
      showToast('تم حذف جدول الوصف التقسيمي');
    },
    [showToast]
  );

  const handleDeleteClause = useCallback(
    async (id: string) => {
      await deleteNotaryClause(id);
      setClauses(await loadNotaryClauses());
    },
    []
  );

  const handleImportStandardClausesPack = useCallback(async () => {
    try {
      const { OPTIONAL_STANDARD_CLAUSES_PACK } = await import(
        '@/lib/standard-clauses-pack'
      );
      const existingTitles = new Set(clauses.map((c) => c.title.trim()));
      let addedCount = 0;
      for (const item of OPTIONAL_STANDARD_CLAUSES_PACK) {
        if (!existingTitles.has(item.title.trim())) {
          await saveNotaryClause({
            ...item,
            id: `${item.id}_${Date.now()}_${addedCount}`,
            order: clauses.length + addedCount,
            updatedAt: new Date().toISOString(),
          });
          addedCount++;
        }
      }
      setClauses(await loadNotaryClauses());
      if (addedCount > 0) {
        showToast(`تم استيراد (${addedCount}) بند توثيقي قياسي إلى مكتبة المكتب`);
      } else {
        showToast('جميع البنود التوثيقية القياسية موجودة بالفعل في مكتبة المكتب');
      }
    } catch {
      showToast('تعذر تحميل حزمة البنود القياسية');
    }
  }, [clauses, showToast]);

  const handleMoveClauseOrder = useCallback(
    async (clauseId: string, direction: 'up' | 'down') => {
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
    },
    [clauses]
  );

  const handleSaveNewClause = useCallback(
    async (title: string, category: string, contentHtml: string) => {
      const newClause: NotaryClause = {
        id: `clause_${Date.now()}`,
        title,
        category,
        contentHtml,
        enabled: true,
        order: clauses.length,
        updatedAt: new Date().toISOString(),
      };
      await saveNotaryClause(newClause);
      setClauses(await loadNotaryClauses());
      showToast(`تم حفظ البند "${title}" في مكتبة البنود الجاهزة`);
    },
    [clauses.length, showToast]
  );

  const handleDeleteTemplate = useCallback(
    async (id: string) => {
      await deleteOfficeTemplate(id);
      setTemplates(await loadOfficeTemplates());
      showToast('تم حذف القالب من المكتبة الموحدة');
    },
    [showToast]
  );

  const handleSaveTemplateFromSnapshot = useCallback(
    async (params: {
      name: string;
      category: string;
      kind?: OfficeTemplateKind;
      code?: string;
      description?: string;
      clearFilledValues?: boolean;
      snapshot: DocumentSessionSnapshot;
    }) => {
      const {
        name,
        category,
        kind = 'contract',
        code,
        description = '',
        clearFilledValues = true,
        snapshot,
      } = params;
      const placeholders = extractPlaceholdersFromHtml(
        snapshot.headerHtml,
        snapshot.bodyHtml,
        snapshot.footerHtml
      );
      const newTpl: OfficeTemplate = {
        id: `tpl_${Date.now()}`,
        name,
        kind,
        code: code || (kind !== 'contract' ? name.replace(/\s+/g, '_') : undefined),
        category,
        description,
        bodyHtml: snapshot.bodyHtml,
        headerHtml: snapshot.headerHtml,
        footerHtml: snapshot.footerHtml,
        pageNumberingEnabled: snapshot.pageNumberingEnabled,
        extractedPlaceholders: placeholders,
        defaultFieldValues: clearFilledValues ? {} : { ...snapshot.fieldValues },
        updatedAt: new Date().toISOString(),
      };
      await saveOfficeTemplate(newTpl);
      setTemplates(await loadOfficeTemplates());
      showToast(`تم حفظ القالب "${name}" في مكتبة المكتب الموحدة بنجاح`);
    },
    [showToast]
  );

  const handleSaveCurrentPartyToDirectory = useCallback(
    async (role: 'party1' | 'party2', vals: Record<string, string>) => {
      const prefix = role === 'party1' ? 'الطرف_الأول' : 'الطرف_الثاني';
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
        nationalIdNin: vals[`${prefix}_الرقم_الوطني`] || '',
        idCardRef: vals[`${prefix}_الهوية`] || '',
        address: vals[`${prefix}_الإقامة`] || '',
        updatedAt: new Date().toISOString(),
      };
      await savePartyRecord(record);
      setSavedParties(await loadSavedParties());
      showToast(`تم حفظ الطرف "${fullName}" في دفتر الأطراف المحفوظين`);
    },
    [savedParties, showToast]
  );

  const handleSavePartyCardToDirectory = useCallback(
    async (card: ContractPartyCard) => {
      const fullName = (card.fullName || '').trim();
      if (!fullName) {
        showToast('يرجى إدخال الاسم الكامل للطرف أولاً قبل حفظه في دفتر الأطراف');
        return;
      }
      const existing = savedParties.find((p) => p.fullName === fullName);
      const record: SavedPartyRecord = {
        id: existing ? existing.id : `party_${Date.now()}`,
        fullName,
        birthDate: card.birthDate || '',
        birthPlace: card.birthPlace || '',
        parentage: card.filiation || '',
        nationalIdNin: card.nationalIdNin || '',
        idCardRef: card.idCardDetails || '',
        address: card.address || '',
        updatedAt: new Date().toISOString(),
      };
      await savePartyRecord(record);
      setSavedParties(await loadSavedParties());
      showToast(`تم حفظ الطرف "${fullName}" في دفتر الأطراف المحفوظين`);
    },
    [savedParties, showToast]
  );

  const handleDeleteSavedParty = useCallback(
    async (id: string) => {
      await deletePartyRecord(id);
      setSavedParties(await loadSavedParties());
      showToast('تم حذف الطرف من الدفتر');
    },
    [showToast]
  );

  const handleSaveCurrentPropertyToDirectory = useCallback(
    async (vals: Record<string, string>) => {
      const lotNumber = (vals['رقم_الحصة'] || '').trim();
      const nature = (vals['طبيعة_الحصة'] || '').trim();
      const fullDesc = (vals['تعيين_الحصة_الكامل'] || '').trim();
      if (!lotNumber && !fullDesc) {
        showToast(
          'يرجى إدخال رقم الحصة أو التعيين الكامل أولاً قبل حفظ العقار في الدفتر'
        );
        return;
      }
      const label = `${nature || 'عقار'} ${
        lotNumber ? `حصة رقم ${lotNumber}` : ''
      }`.trim();
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
    },
    [showToast]
  );

  const handleDeleteSavedProperty = useCallback(
    async (id: string) => {
      await deletePropertyRecord(id);
      setSavedProperties(await loadSavedProperties());
      showToast('تم حذف العقار من الدفتر');
    },
    [showToast]
  );

  const handleExportTemplatesJson = useCallback(async () => {
    try {
      await downloadTemplatesJsonFile();
      showToast('تم تصدير مكتبة قوالب المكتب الموحّدة بنجاح');
    } catch (err) {
      showToast(err instanceof Error ? err.message : 'تعذر تصدير القوالب');
    }
  }, [showToast]);

  const handleImportTemplatesJson = useCallback(
    async (file: File) => {
      try {
        const { count, templates: nextTemplates } =
          await importTemplatesJsonFile(file);
        setTemplates(nextTemplates);
        showToast(`تم استيراد ودمج ${count} قالب في مكتبة المكتب`);
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'ملف القوالب غير صالح للاستيراد'
        );
      }
    },
    [showToast]
  );

  const handleImportDocxAsTemplate = useCallback(
    async (files: FileList) => {
      try {
        const { importedCount, templates: updated } =
          await importDocxFilesAsTemplates(files);
        setTemplates(updated);
        if (importedCount > 0) {
          showToast(
            `تم استيراد ${importedCount} قالب Word وحفظها في قوالب المكتب`
          );
        }
      } catch (err) {
        showToast(
          err instanceof Error ? err.message : 'تعذر استيراد ملف القالب'
        );
      }
    },
    [showToast]
  );

  return {
    templates,
    setTemplates,
    clauses,
    setClauses,
    estates,
    setEstates,
    savedParties,
    setSavedParties,
    savedProperties,
    setSavedProperties,
    clerks,
    setClerks,
    activeClerk,
    setActiveClerk,
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
    handleDeleteFolderOnly,
    handleSaveEstate,
    handleDeleteEstate,
    handleDeleteClause,
    handleImportStandardClausesPack,
    handleMoveClauseOrder,
    handleSaveNewClause,
    handleDeleteTemplate,
    handleSaveTemplateFromSnapshot,
    handleSaveCurrentPartyToDirectory,
    handleSavePartyCardToDirectory,
    handleDeleteSavedParty,
    handleSaveCurrentPropertyToDirectory,
    handleDeleteSavedProperty,
    handleExportTemplatesJson,
    handleImportTemplatesJson,
    handleImportDocxAsTemplate,
  };
}
