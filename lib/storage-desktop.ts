import type {
  BackupBundle,
  ContractFolder,
  DocumentRevision,
  LegacyCustomTemplate,
  LegacyDerivedDocTemplate,
  LegacyDownloadArchiveItem,
  LegacyWordTemplateDefinition,
  NotaryClerk,
  NotaryClause,
  OfficeTemplate,
  PublicationTrack,
  RegistrationBatch,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SerializedSealedOriginal,
  SubdivisionEstate,
} from './types';
import { getDesktopBridge, type DesktopOfficeBridge } from './desktop-bridge';
import {
  base64ToUint8Array,
  buildArchiveRelativePath,
  computeSha256Hex,
  DEFAULT_NOTARY_CLERKS,
  mergeLegacyStoresIntoOfficeTemplates,
  migrateCustomTemplateToOfficeTemplate,
  migrateDerivedTemplateToOfficeTemplate,
  migrateDownloadArchiveItemToRevision,
  migrateWordTemplateToOfficeTemplate,
  normalizeAnyTemplateToOfficeTemplate,
  officeTemplateToCustomTemplate,
  officeTemplateToDerivedDocTemplate,
  type OfficeStore,
  type OfficeStoreRuntimeInfo,
  type ReadSealedOriginalResult,
  type SealContractOriginalInput,
  type SealContractOriginalResult,
  uint8ArrayToBase64,
} from './office-store';

function requireBridge(): DesktopOfficeBridge {
  const bridge = getDesktopBridge();
  if (!bridge) {
    throw new Error('جسر سطح المكتب غير متوفر');
  }
  return bridge;
}

async function desktopLoad<T>(collection: string): Promise<T[]> {
  const bridge = requireBridge();
  const res = await bridge.invoke<T[] | null>('storage_load', { collection });
  return Array.isArray(res) ? res : [];
}

async function desktopSave<T>(collection: string, item: T): Promise<void> {
  const bridge = requireBridge();
  await bridge.invoke<void>('storage_save', { collection, item });
}

async function desktopDelete(collection: string, id: string): Promise<void> {
  const bridge = requireBridge();
  await bridge.invoke<void>('storage_delete', { collection, id });
}

async function desktopClear(collection: string): Promise<void> {
  const bridge = requireBridge();
  await bridge.invoke<void>('storage_clear', { collection });
}

export async function getRuntimeInfo(): Promise<OfficeStoreRuntimeInfo> {
  const bridge = requireBridge();
  const dir = await bridge.getDataDirectory();
  if (!dir || typeof dir !== 'string' || !dir.trim()) {
    throw new Error('تعذر الحصول على مسار مجلد بيانات المكتب من جسر سطح المكتب');
  }
  return {
    mode: 'desktop',
    dataDirectory: dir.trim(),
  };
}

export async function sealContractOriginal(
  input: SealContractOriginalInput
): Promise<SealContractOriginalResult> {
  const bridge = requireBridge();
  const relativePath = buildArchiveRelativePath(input.year, input.indexNumber);
  const sha256 = await computeSha256Hex(input.docxBytes);
  const base64Data = uint8ArrayToBase64(input.docxBytes);

  const res = await bridge.invoke<SealContractOriginalResult | null>('seal_original', {
    documentId: input.documentId,
    year: input.year,
    indexNumber: input.indexNumber,
    relativePath,
    sha256,
    byteLength: input.docxBytes.byteLength,
    base64Data,
    docxBytes: input.docxBytes,
  });

  return {
    relativePath: res?.relativePath || relativePath,
    sha256: res?.sha256 || sha256,
  };
}

export async function readSealedOriginal(
  documentId: string
): Promise<ReadSealedOriginalResult | null> {
  const bridge = requireBridge();
  const res = await bridge.invoke<{
    relativePath: string;
    sha256: string;
    bytes?: Uint8Array | number[];
    base64Data?: string;
  } | null>('read_original', { documentId });

  if (!res) return null;
  let bytes: Uint8Array;
  if (res.bytes instanceof Uint8Array) {
    bytes = res.bytes;
  } else if (Array.isArray(res.bytes)) {
    bytes = new Uint8Array(res.bytes);
  } else if (res.base64Data) {
    bytes = base64ToUint8Array(res.base64Data);
  } else {
    return null;
  }

  const actualSha = await computeSha256Hex(bytes);
  if (res.sha256 && actualSha !== res.sha256) {
    throw new Error(
      `فشل التحقق من سلامة الأصل المختوم (${res.relativePath}): بصمة SHA-256 غير مطابقة`
    );
  }

  return {
    relativePath: res.relativePath,
    sha256: res.sha256 || actualSha,
    bytes,
  };
}

// 1. Clauses CRUD
export async function loadNotaryClauses(): Promise<NotaryClause[]> {
  const items = await desktopLoad<NotaryClause>('clauses');
  return items.sort((a, b) => a.order - b.order);
}

export async function saveNotaryClause(clause: NotaryClause): Promise<void> {
  await desktopSave('clauses', clause);
}

export async function deleteNotaryClause(id: string): Promise<void> {
  await desktopDelete('clauses', id);
}

// 2. Unified Office Templates CRUD (with idempotent migration flag)
export async function loadOfficeTemplates(): Promise<OfficeTemplate[]> {
  const [existingOffice, migrationFlags] = await Promise.all([
    desktopLoad<OfficeTemplate>('office_templates'),
    desktopLoad<{ id: string; migrated?: boolean }>('migration_flags'),
  ]);
  const isMigrated = migrationFlags.some(
    (f) => f.id === 'office_templates_migrated_v1' && f.migrated
  );

  if (isMigrated) {
    return existingOffice.sort((a, b) =>
      (b.updatedAt || '').localeCompare(a.updatedAt || '')
    );
  }

  const [legacyCustom, legacyDerived, legacyWord] = await Promise.all([
    desktopLoad<LegacyCustomTemplate>('templates'),
    desktopLoad<LegacyDerivedDocTemplate>('derived_templates'),
    desktopLoad<LegacyWordTemplateDefinition>('word_templates'),
  ]);

  const merged = mergeLegacyStoresIntoOfficeTemplates({
    existingOffice,
    legacyCustom,
    legacyDerived,
    legacyWord,
    seedDefaults: true,
  });

  for (const tpl of merged) {
    await desktopSave('office_templates', tpl);
  }
  await desktopSave('migration_flags', {
    id: 'office_templates_migrated_v1',
    migrated: true,
    updatedAt: new Date().toISOString(),
  });
  return merged;
}

export async function saveOfficeTemplate(
  template: OfficeTemplate
): Promise<void> {
  await desktopSave('office_templates', template);
}

export async function deleteOfficeTemplate(id: string): Promise<void> {
  await loadOfficeTemplates();
  await desktopDelete('office_templates', id);
  await desktopDelete('templates', id);
  await desktopDelete('derived_templates', id);
  await desktopDelete('word_templates', id);
}

export async function exportOfficeTemplatesJson(): Promise<string> {
  const templates = await loadOfficeTemplates();
  return JSON.stringify(templates, null, 2);
}

export async function importOfficeTemplatesJson(
  templates: unknown[]
): Promise<number> {
  let count = 0;
  for (const raw of templates) {
    const normalized = normalizeAnyTemplateToOfficeTemplate(raw);
    if (normalized) {
      await saveOfficeTemplate(normalized);
      count++;
    }
  }
  return count;
}

// 3. Document Revisions CRUD (Keeps latest 25 revisions per documentId + one-time migration of legacy downloads)
export async function loadDocumentRevisions(): Promise<DocumentRevision[]> {
  const migrationFlags = await desktopLoad<{ id: string; migrated?: boolean }>(
    'migration_flags'
  );
  const downloadsMigrated = migrationFlags.some(
    (f) => f.id === 'office_downloads_migrated_v1' && f.migrated
  );
  if (!downloadsMigrated) {
    const legacyDownloads = await desktopLoad<LegacyDownloadArchiveItem>('downloads');
    if (legacyDownloads.length > 0) {
      for (const dl of legacyDownloads) {
        if (dl && dl.id) {
          const rev = migrateDownloadArchiveItemToRevision(dl);
          await desktopSave('revisions', rev);
        }
      }
      await desktopClear('downloads').catch(() => {});
    }
    await desktopSave('migration_flags', {
      id: 'office_downloads_migrated_v1',
      migrated: true,
      updatedAt: new Date().toISOString(),
    });
  }

  const items = await desktopLoad<DocumentRevision>('revisions');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDocumentRevision(rev: DocumentRevision): Promise<void> {
  const isFullBodyEvent =
    !rev.eventType ||
    rev.eventType === 'snapshot' ||
    rev.eventType === 'seal' ||
    rev.eventType === 'save';
  const normalizedRev: DocumentRevision = {
    ...rev,
    eventType: rev.eventType || 'snapshot',
    ...(isFullBodyEvent ? { bodyHtml: rev.bodyHtml || '' } : { bodyHtml: undefined }),
  };
  await desktopSave('revisions', normalizedRev);
  const allRevs = await loadDocumentRevisions();
  const sameDocRevs = allRevs
    .filter((r) => r.documentId === rev.documentId)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  if (sameDocRevs.length > 25) {
    for (const oldRev of sameDocRevs.slice(25)) {
      await desktopDelete('revisions', oldRev.id);
    }
  }
}

export async function deleteDocumentRevision(id: string): Promise<void> {
  await desktopDelete('revisions', id);
}

// 6. Saved Documents CRUD
export async function loadSavedDocuments(): Promise<SavedDocument[]> {
  const items = await desktopLoad<SavedDocument>('documents');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveDocumentRecord(doc: SavedDocument): Promise<void> {
  await desktopSave('documents', doc);
}

export async function deleteDocumentRecord(id: string): Promise<void> {
  await desktopDelete('documents', id);
}

// 7. Subdivision Estates CRUD
export async function loadSubdivisionEstates(): Promise<SubdivisionEstate[]> {
  const items = await desktopLoad<SubdivisionEstate>('estates');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveSubdivisionEstate(estate: SubdivisionEstate): Promise<void> {
  await desktopSave('estates', estate);
}

export async function deleteSubdivisionEstate(id: string): Promise<void> {
  await desktopDelete('estates', id);
}

// 7b. Saved Parties & Properties Directory
export async function loadSavedParties(): Promise<SavedPartyRecord[]> {
  const items = await desktopLoad<SavedPartyRecord>('saved_parties');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePartyRecord(party: SavedPartyRecord): Promise<void> {
  await desktopSave('saved_parties', party);
}

export async function deletePartyRecord(id: string): Promise<void> {
  await desktopDelete('saved_parties', id);
}

export async function loadSavedProperties(): Promise<SavedPropertyRecord[]> {
  const items = await desktopLoad<SavedPropertyRecord>('saved_properties');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePropertyRecord(prop: SavedPropertyRecord): Promise<void> {
  await desktopSave('saved_properties', prop);
}

export async function deletePropertyRecord(id: string): Promise<void> {
  await desktopDelete('saved_properties', id);
}

// 7c. Clerks & Folders
export async function loadNotaryClerks(): Promise<NotaryClerk[]> {
  const items = await desktopLoad<NotaryClerk>('clerks');
  if (items.length === 0) {
    for (const c of DEFAULT_NOTARY_CLERKS) {
      await desktopSave('clerks', c);
    }
    return DEFAULT_NOTARY_CLERKS;
  }
  return items;
}

export async function saveNotaryClerk(clerk: NotaryClerk): Promise<void> {
  await desktopSave('clerks', clerk);
}

export async function deleteNotaryClerk(id: string): Promise<void> {
  await desktopDelete('clerks', id);
}

export async function loadContractFolders(): Promise<ContractFolder[]> {
  const items = await desktopLoad<ContractFolder>('folders');
  return items.sort((a, b) => a.name.localeCompare(b.name));
}

export async function saveContractFolder(folder: ContractFolder): Promise<void> {
  await desktopSave('folders', folder);
}

export async function deleteContractFolder(id: string): Promise<void> {
  await desktopDelete('folders', id);
}

// 7d. Registration Batches & Land Publication Tracks
export async function loadRegistrationBatches(): Promise<RegistrationBatch[]> {
  const items = await desktopLoad<RegistrationBatch>('registration_batches');
  return items.sort((a, b) => b.statusDate.localeCompare(a.statusDate));
}

export async function saveRegistrationBatch(batch: RegistrationBatch): Promise<void> {
  await desktopSave('registration_batches', batch);
}

export async function deleteRegistrationBatch(id: string): Promise<void> {
  await desktopDelete('registration_batches', id);
}

export async function loadPublicationTracks(): Promise<PublicationTrack[]> {
  const items = await desktopLoad<PublicationTrack>('publication_tracks');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePublicationTrack(track: PublicationTrack): Promise<void> {
  await desktopSave('publication_tracks', track);
}

export async function deletePublicationTrack(id: string): Promise<void> {
  await desktopDelete('publication_tracks', id);
}

// 9. Active Draft Session (Async from Disk / SQLite)
export async function loadActiveDraftSession(): Promise<SavedDocument | null> {
  const bridge = requireBridge();
  const res = await bridge.invoke<SavedDocument[] | SavedDocument | null>(
    'storage_load',
    { collection: 'active_draft' }
  );
  if (!res) return null;
  if (Array.isArray(res)) {
    return res.length > 0 ? res[0] : null;
  }
  if (typeof res === 'object' && 'id' in res) {
    return res as SavedDocument;
  }
  return null;
}

export async function saveActiveDraftSession(doc: SavedDocument): Promise<void> {
  const bridge = requireBridge();
  await bridge.invoke<void>('storage_save', {
    collection: 'active_draft',
    item: doc,
  });
}

// 10. Full Backup Export & Import (Triggers backup_database on desktop bridge)
export async function exportFullBackupBundle(): Promise<BackupBundle> {
  const bridge = requireBridge();
  await bridge.invoke('backup_database', {
    timestamp: new Date().toISOString(),
  });

  const [
    officeTemplates,
    documents,
    estates,
    clauses,
    revisions,
    savedParties,
    savedProperties,
    clerks,
    folders,
    registrationBatches,
    publicationTracks,
    sealedOriginals,
  ] = await Promise.all([
    loadOfficeTemplates(),
    loadSavedDocuments(),
    loadSubdivisionEstates(),
    loadNotaryClauses(),
    loadDocumentRevisions(),
    loadSavedParties(),
    loadSavedProperties(),
    loadNotaryClerks(),
    loadContractFolders(),
    loadRegistrationBatches(),
    loadPublicationTracks(),
    desktopLoad<SerializedSealedOriginal>('sealed_originals'),
  ]);

  const templates = officeTemplates
    .filter((t) => t.kind === 'contract')
    .map(officeTemplateToCustomTemplate);
  const derivedTemplates = officeTemplates
    .filter((t) => t.kind !== 'contract')
    .map(officeTemplateToDerivedDocTemplate);

  return {
    version: '2.8.0',
    exportedAt: new Date().toISOString(),
    clerks,
    folders,
    officeTemplates,
    templates,
    documents,
    estates,
    clauses,
    derivedTemplates,
    revisions,
    savedParties,
    savedProperties,
    registrationBatches,
    publicationTracks,
    ...(sealedOriginals.length > 0 ? { sealedOriginals } : {}),
  };
}

export async function importFullBackupBundle(bundle: BackupBundle): Promise<{
  templatesCount: number;
  documentsCount: number;
  estatesCount: number;
  clausesCount: number;
}> {
  requireBridge();
  let templatesCount = 0;
  let documentsCount = 0;
  let estatesCount = 0;
  let clausesCount = 0;

  if (Array.isArray(bundle.clerks)) {
    for (const cl of bundle.clerks) {
      if (cl && cl.id && cl.name) await saveNotaryClerk(cl);
    }
  }
  if (Array.isArray(bundle.folders)) {
    for (const f of bundle.folders) {
      if (f && f.id && f.name) await saveContractFolder(f);
    }
  }

  const importedTemplateIds = new Set<string>();
  if (Array.isArray(bundle.officeTemplates)) {
    for (const ot of bundle.officeTemplates) {
      const normalized = normalizeAnyTemplateToOfficeTemplate(ot);
      if (normalized) {
        await saveOfficeTemplate(normalized);
        importedTemplateIds.add(normalized.id);
        templatesCount++;
      }
    }
  }
  if (Array.isArray(bundle.templates)) {
    for (const tpl of bundle.templates) {
      if (tpl && tpl.id && tpl.name && !importedTemplateIds.has(tpl.id)) {
        await saveOfficeTemplate(migrateCustomTemplateToOfficeTemplate(tpl));
        importedTemplateIds.add(tpl.id);
        templatesCount++;
      }
    }
  }
  if (Array.isArray(bundle.derivedTemplates)) {
    for (const dt of bundle.derivedTemplates) {
      if (dt && dt.id && dt.name && !importedTemplateIds.has(dt.id)) {
        await saveOfficeTemplate(migrateDerivedTemplateToOfficeTemplate(dt));
        importedTemplateIds.add(dt.id);
      }
    }
  }
  if (Array.isArray(bundle.wordTemplates)) {
    for (const wt of bundle.wordTemplates) {
      if (wt && wt.id && wt.name && !importedTemplateIds.has(wt.id)) {
        await saveOfficeTemplate(migrateWordTemplateToOfficeTemplate(wt));
        importedTemplateIds.add(wt.id);
      }
    }
  }

  if (Array.isArray(bundle.documents)) {
    for (const doc of bundle.documents) {
      if (doc && doc.id && doc.title) {
        await saveDocumentRecord(doc);
        documentsCount++;
      }
    }
  }
  if (Array.isArray(bundle.estates)) {
    for (const est of bundle.estates) {
      if (est && est.id && est.estateName) {
        await saveSubdivisionEstate(est);
        estatesCount++;
      }
    }
  }
  if (Array.isArray(bundle.clauses)) {
    for (const cl of bundle.clauses) {
      if (cl && cl.id && cl.title) {
        await saveNotaryClause(cl);
        clausesCount++;
      }
    }
  }
  if (Array.isArray(bundle.revisions)) {
    for (const rev of bundle.revisions) {
      if (rev && rev.id) await saveDocumentRevision(rev);
    }
  }
  if (Array.isArray(bundle.downloads)) {
    for (const dl of bundle.downloads) {
      if (dl && dl.id) {
        await saveDocumentRevision(migrateDownloadArchiveItemToRevision(dl));
      }
    }
  }
  if (Array.isArray(bundle.savedParties)) {
    for (const p of bundle.savedParties) {
      if (p && p.id && p.fullName) await savePartyRecord(p);
    }
  }
  if (Array.isArray(bundle.savedProperties)) {
    for (const prop of bundle.savedProperties) {
      if (prop && prop.id && prop.label) await savePropertyRecord(prop);
    }
  }
  if (Array.isArray(bundle.registrationBatches)) {
    for (const batch of bundle.registrationBatches) {
      if (batch && batch.id) await saveRegistrationBatch(batch);
    }
  }
  if (Array.isArray(bundle.publicationTracks)) {
    for (const track of bundle.publicationTracks) {
      if (track && track.id && track.documentId) await savePublicationTrack(track);
    }
  }
  if (Array.isArray(bundle.sealedOriginals)) {
    for (const so of bundle.sealedOriginals) {
      if (so && so.documentId && so.base64Data) {
        const matchYear = so.relativePath.match(/archive\/(\d{4})\/([^/.]+)/);
        const year = matchYear ? parseInt(matchYear[1], 10) : new Date().getFullYear();
        const indexNumber = matchYear ? matchYear[2] : so.documentId;
        await sealContractOriginal({
          documentId: so.documentId,
          year,
          indexNumber,
          docxBytes: base64ToUint8Array(so.base64Data),
        });
      }
    }
  }

  return { templatesCount, documentsCount, estatesCount, clausesCount };
}

export const desktopOfficeStore: OfficeStore = {
  getRuntimeInfo,
  sealContractOriginal,
  readSealedOriginal,
  loadNotaryClauses,
  saveNotaryClause,
  deleteNotaryClause,
  loadOfficeTemplates,
  saveOfficeTemplate,
  deleteOfficeTemplate,
  exportOfficeTemplatesJson,
  importOfficeTemplatesJson,
  loadDocumentRevisions,
  saveDocumentRevision,
  deleteDocumentRevision,
  loadSavedDocuments,
  saveDocumentRecord,
  deleteDocumentRecord,
  loadSubdivisionEstates,
  saveSubdivisionEstate,
  deleteSubdivisionEstate,
  loadSavedParties,
  savePartyRecord,
  deletePartyRecord,
  loadSavedProperties,
  savePropertyRecord,
  deletePropertyRecord,
  loadNotaryClerks,
  saveNotaryClerk,
  deleteNotaryClerk,
  loadContractFolders,
  saveContractFolder,
  deleteContractFolder,
  loadRegistrationBatches,
  saveRegistrationBatch,
  deleteRegistrationBatch,
  loadPublicationTracks,
  savePublicationTrack,
  deletePublicationTrack,
  loadActiveDraftSession,
  saveActiveDraftSession,
  exportFullBackupBundle,
  importFullBackupBundle,
};
