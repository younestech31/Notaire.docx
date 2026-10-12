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
  SealedContractOriginal,
  SerializedSealedOriginal,
  SubdivisionEstate,
} from './types';
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

const DB_NAME = 'NotarySmartEditorDB';
const DB_VERSION = 9;
const STORE_OFFICE_TEMPLATES = 'office_templates';
const STORE_TEMPLATES = 'templates';
const STORE_DOCUMENTS = 'documents';
const STORE_ESTATES = 'estates';
const STORE_CLAUSES = 'clauses';
const STORE_DERIVED_TEMPLATES = 'derived_templates';
const STORE_REVISIONS = 'revisions';
const STORE_DOWNLOADS = 'downloads';
const STORE_SAVED_PARTIES = 'saved_parties';
const STORE_SAVED_PROPERTIES = 'saved_properties';
const STORE_CLERKS = 'clerks';
const STORE_FOLDERS = 'folders';
const STORE_WORD_TEMPLATES = 'word_templates';
const STORE_SEALED_ORIGINALS = 'sealed_originals';
const STORE_REGISTRATION_BATCHES = 'registration_batches';
const STORE_PUBLICATION_TRACKS = 'publication_tracks';

const LS_ACTIVE_DOC_KEY = 'notary_active_document_v2';
const LS_MIGRATED_TEMPLATES_KEY = 'office_templates_migrated_v1';
const LS_MIGRATED_DOWNLOADS_KEY = 'office_downloads_migrated_v1';
let memoryTemplatesMigrated = false;
let memoryDownloadsMigrated = false;

function isOfficeTemplatesMigrated(): boolean {
  if (memoryTemplatesMigrated) return true;
  if (typeof window !== 'undefined') {
    try {
      return localStorage.getItem(LS_MIGRATED_TEMPLATES_KEY) === 'true';
    } catch {
      return false;
    }
  }
  return false;
}

function setOfficeTemplatesMigrated(val: boolean): void {
  memoryTemplatesMigrated = val;
  if (typeof window !== 'undefined') {
    try {
      if (val) localStorage.setItem(LS_MIGRATED_TEMPLATES_KEY, 'true');
      else localStorage.removeItem(LS_MIGRATED_TEMPLATES_KEY);
    } catch {
      // ignore storage error
    }
  }
}

function isDownloadsMigrated(): boolean {
  if (memoryDownloadsMigrated) return true;
  if (typeof window !== 'undefined') {
    try {
      return localStorage.getItem(LS_MIGRATED_DOWNLOADS_KEY) === 'true';
    } catch {
      return false;
    }
  }
  return false;
}

function setDownloadsMigrated(val: boolean): void {
  memoryDownloadsMigrated = val;
  if (typeof window !== 'undefined') {
    try {
      if (val) localStorage.setItem(LS_MIGRATED_DOWNLOADS_KEY, 'true');
      else localStorage.removeItem(LS_MIGRATED_DOWNLOADS_KEY);
    } catch {
      // ignore storage error
    }
  }
}

interface StoredSealedOriginalRecord extends SealedContractOriginal {
  bytes?: Uint8Array;
  base64Data?: string;
}

function openDatabase(): Promise<IDBDatabase | null> {
  if (typeof window === 'undefined' || !('indexedDB' in window)) {
    return Promise.resolve(null);
  }
  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        const storesWithIdKey = [
          STORE_OFFICE_TEMPLATES,
          STORE_TEMPLATES,
          STORE_DOCUMENTS,
          STORE_ESTATES,
          STORE_CLAUSES,
          STORE_DERIVED_TEMPLATES,
          STORE_REVISIONS,
          STORE_DOWNLOADS,
          STORE_SAVED_PARTIES,
          STORE_SAVED_PROPERTIES,
          STORE_CLERKS,
          STORE_FOLDERS,
          STORE_WORD_TEMPLATES,
          STORE_REGISTRATION_BATCHES,
          STORE_PUBLICATION_TRACKS,
        ];
        for (const s of storesWithIdKey) {
          if (!db.objectStoreNames.contains(s)) {
            db.createObjectStore(s, { keyPath: 'id' });
          }
        }
        if (!db.objectStoreNames.contains(STORE_SEALED_ORIGINALS)) {
          db.createObjectStore(STORE_SEALED_ORIGINALS, { keyPath: 'documentId' });
        }
      };
      request.onsuccess = () => {
        const db = request.result;
        db.onversionchange = () => {
          db.close();
        };
        resolve(db);
      };
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function getAllFromStore<T>(storeName: string, lsFallbackKey: string): Promise<T[]> {
  try {
    const db = await openDatabase();
    if (db && db.objectStoreNames.contains(storeName)) {
      const dbResult = await new Promise<T[]>((resolve) => {
        try {
          const tx = db.transaction(storeName, 'readonly');
          const store = tx.objectStore(storeName);
          const req = store.getAll();
          req.onsuccess = () => resolve((req.result as T[]) || []);
          req.onerror = () => resolve([]);
        } catch {
          resolve([]);
        }
      });
      if (dbResult.length > 0) {
        return dbResult;
      }
    }
  } catch {
    // Fall back to localStorage below
  }
  if (typeof window !== 'undefined') {
    try {
      const raw = localStorage.getItem(lsFallbackKey);
      if (!raw) return [];
      return JSON.parse(raw) as T[];
    } catch {
      return [];
    }
  }
  return [];
}

async function putInStore<T extends { id: string }>(
  storeName: string,
  lsFallbackKey: string,
  item: T
): Promise<void> {
  let savedToIdb = false;
  try {
    const db = await openDatabase();
    if (db && db.objectStoreNames.contains(storeName)) {
      savedToIdb = await new Promise<boolean>((resolve) => {
        try {
          const tx = db.transaction(storeName, 'readwrite');
          tx.objectStore(storeName).put(item);
          tx.oncomplete = () => resolve(true);
          tx.onerror = () => resolve(false);
          tx.onabort = () => resolve(false);
        } catch {
          resolve(false);
        }
      });
    }
  } catch {
    savedToIdb = false;
  }

  // Only fall back to localStorage if IndexedDB was unavailable or failed
  if (!savedToIdb && typeof window !== 'undefined') {
    try {
      const raw = localStorage.getItem(lsFallbackKey);
      const existing: T[] = raw ? (JSON.parse(raw) as T[]) : [];
      const idx = existing.findIndex((x) => x.id === item.id);
      if (idx >= 0) existing[idx] = item;
      else existing.unshift(item);
      localStorage.setItem(lsFallbackKey, JSON.stringify(existing));
    } catch {
      // Ignore quota exceeded on emergency fallback
    }
  }
}

async function deleteFromStore(
  storeName: string,
  lsFallbackKey: string,
  id: string
): Promise<void> {
  let deletedFromIdb = false;
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(storeName)) {
    await new Promise<void>((resolve, reject) => {
      try {
        const tx = db.transaction(storeName, 'readwrite');
        tx.objectStore(storeName).delete(id);
        tx.oncomplete = () => {
          deletedFromIdb = true;
          resolve();
        };
        tx.onerror = () =>
          reject(
            new Error(
              `فشل حذف السجل من (${storeName}): ${
                tx.error?.message || 'خطأ في معاملة الحذف'
              }`
            )
          );
        tx.onabort = () =>
          reject(
            new Error(
              `تم إحباط معاملة الحذف من (${storeName}): ${
                tx.error?.message || ''
              }`
            )
          );
      } catch (err) {
        reject(
          err instanceof Error
            ? err
            : new Error(`فشل تنفيذ معاملة الحذف من (${storeName})`)
        );
      }
    });
  }
  if (!deletedFromIdb && typeof window !== 'undefined') {
    try {
      const raw = localStorage.getItem(lsFallbackKey);
      if (raw) {
        const existing = JSON.parse(raw) as { id: string }[];
        const filtered = existing.filter((x) => x.id !== id);
        localStorage.setItem(lsFallbackKey, JSON.stringify(filtered));
      }
    } catch {
      // Ignore localStorage fallback error
    }
  }
}

export async function getRuntimeInfo(): Promise<OfficeStoreRuntimeInfo> {
  return {
    mode: 'browser',
    dataDirectory: null,
  };
}

// =========================================================
// SEALED ORIGINALS (الأصول المعتمدة ببصمة SHA-256 — في IndexedDB حصراً دون localStorage)
// =========================================================
async function getAllSealedOriginals(): Promise<StoredSealedOriginalRecord[]> {
  const db = await openDatabase();
  if (!db || !db.objectStoreNames.contains(STORE_SEALED_ORIGINALS)) {
    return [];
  }
  return await new Promise<StoredSealedOriginalRecord[]>((resolve, reject) => {
    try {
      const tx = db.transaction(STORE_SEALED_ORIGINALS, 'readonly');
      const store = tx.objectStore(STORE_SEALED_ORIGINALS);
      const req = store.getAll();
      req.onsuccess = () =>
        resolve((req.result as StoredSealedOriginalRecord[]) || []);
      req.onerror = () =>
        reject(
          new Error(
            `فشل قراءة مخزن الأصول المعتمدة: ${req.error?.message || ''}`
          )
        );
    } catch (err) {
      reject(
        err instanceof Error ? err : new Error('فشل قراءة مخزن الأصول المعتمدة')
      );
    }
  });
}

async function putSealedOriginalRecord(
  record: StoredSealedOriginalRecord
): Promise<void> {
  const db = await openDatabase();
  if (!db || !db.objectStoreNames.contains(STORE_SEALED_ORIGINALS)) {
    throw new Error('مخزن الأصول المعتمدة غير متاح في قاعدة بيانات المتصفح');
  }
  await new Promise<void>((resolve, reject) => {
    try {
      const tx = db.transaction(STORE_SEALED_ORIGINALS, 'readwrite');
      tx.objectStore(STORE_SEALED_ORIGINALS).put(record);
      tx.oncomplete = () => resolve();
      tx.onerror = () =>
        reject(
          new Error(
            `فشل حفظ الأصل المعتمد: ${tx.error?.message || 'خطأ في المعاملة'}`
          )
        );
      tx.onabort = () =>
        reject(
          new Error(
            `تم إحباط حفظ الأصل المعتمد: ${tx.error?.message || ''}`
          )
        );
    } catch (err) {
      reject(err instanceof Error ? err : new Error('فشل حفظ الأصل المعتمد'));
    }
  });
}

export async function sealContractOriginal(
  input: SealContractOriginalInput
): Promise<SealContractOriginalResult> {
  const relativePath = buildArchiveRelativePath(input.year, input.indexNumber);
  const sha256 = await computeSha256Hex(input.docxBytes);

  const existingList = await getAllSealedOriginals();
  const conflicting = existingList.find(
    (item) =>
      (item.relativePath === relativePath ||
        item.documentId === input.documentId) &&
      item.sha256 !== sha256
  );

  if (conflicting) {
    throw new Error(
      `رفض ختم الأصل المعتمد: يوجد أصل معتمد مسبقاً في المسار (${conflicting.relativePath}) ببصمة SHA-256 مختلفة ولا يمكن استبداله`
    );
  }

  const record: StoredSealedOriginalRecord = {
    documentId: input.documentId,
    relativePath,
    sha256,
    sealedAt: new Date().toISOString(),
    byteLength: input.docxBytes.byteLength,
    bytes: new Uint8Array(input.docxBytes),
  };

  await putSealedOriginalRecord(record);
  return { relativePath, sha256 };
}

export async function readSealedOriginal(
  documentId: string
): Promise<ReadSealedOriginalResult | null> {
  const all = await getAllSealedOriginals();
  const found = all.find((x) => x.documentId === documentId);
  if (!found) return null;

  let bytes: Uint8Array | null = null;
  if (found.bytes instanceof Uint8Array) {
    bytes = found.bytes;
  } else if (found.base64Data) {
    bytes = base64ToUint8Array(found.base64Data);
  } else if (found.bytes && typeof found.bytes === 'object') {
    bytes = new Uint8Array(Object.values(found.bytes as Record<string, number>));
  }

  if (!bytes) return null;
  return {
    relativePath: found.relativePath,
    sha256: found.sha256,
    bytes,
  };
}

// 1. Clauses (البنود الجاهزة) CRUD
export async function loadNotaryClauses(): Promise<NotaryClause[]> {
  const items = await getAllFromStore<NotaryClause>(STORE_CLAUSES, 'notary_clauses_v2');
  return items.sort((a, b) => a.order - b.order);
}

export async function saveNotaryClause(clause: NotaryClause): Promise<void> {
  await putInStore(STORE_CLAUSES, 'notary_clauses_v2', clause);
}

export async function deleteNotaryClause(id: string): Promise<void> {
  await deleteFromStore(STORE_CLAUSES, 'notary_clauses_v2', id);
}

// 2. Unified Office Templates CRUD (with one-time idempotent migration from legacy stores)
export async function loadOfficeTemplates(): Promise<OfficeTemplate[]> {
  const existingOffice = await getAllFromStore<OfficeTemplate>(
    STORE_OFFICE_TEMPLATES,
    'notary_office_templates_v1'
  );

  if (isOfficeTemplatesMigrated() && existingOffice.length > 0) {
    return existingOffice.sort((a, b) =>
      (b.updatedAt || '').localeCompare(a.updatedAt || '')
    );
  }

  if (isOfficeTemplatesMigrated() && existingOffice.length === 0) {
    // Already migrated and user explicitly deleted all templates
    return [];
  }

  // First-time migration from legacy stores + default templates
  const [legacyCustom, legacyDerived, legacyWord] = await Promise.all([
    getAllFromStore<LegacyCustomTemplate>(STORE_TEMPLATES, 'notary_templates_v1'),
    getAllFromStore<LegacyDerivedDocTemplate>(
      STORE_DERIVED_TEMPLATES,
      'notary_derived_templates_v2'
    ),
    getAllFromStore<LegacyWordTemplateDefinition>(
      STORE_WORD_TEMPLATES,
      'notary_word_templates_v1'
    ),
  ]);

  const merged = mergeLegacyStoresIntoOfficeTemplates({
    existingOffice,
    legacyCustom,
    legacyDerived,
    legacyWord,
    seedDefaults: true,
  });

  for (const tpl of merged) {
    await putInStore(
      STORE_OFFICE_TEMPLATES,
      'notary_office_templates_v1',
      tpl
    );
  }
  setOfficeTemplatesMigrated(true);
  return merged;
}

export async function saveOfficeTemplate(
  template: OfficeTemplate
): Promise<void> {
  await putInStore(
    STORE_OFFICE_TEMPLATES,
    'notary_office_templates_v1',
    template
  );
}

export async function deleteOfficeTemplate(id: string): Promise<void> {
  // Ensure migration is materialized before deleting so defaults aren't lost or re-added
  if (!isOfficeTemplatesMigrated()) {
    await loadOfficeTemplates();
  }
  await deleteFromStore(
    STORE_OFFICE_TEMPLATES,
    'notary_office_templates_v1',
    id
  );
  await deleteFromStore(STORE_TEMPLATES, 'notary_templates_v1', id);
  await deleteFromStore(
    STORE_DERIVED_TEMPLATES,
    'notary_derived_templates_v2',
    id
  );
  await deleteFromStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', id);
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
  if (!isDownloadsMigrated()) {
    const legacyDownloads = await getAllFromStore<LegacyDownloadArchiveItem>(
      STORE_DOWNLOADS,
      'notary_downloads_v2'
    );
    let extraLsDownloads: LegacyDownloadArchiveItem[] = [];
    if (typeof window !== 'undefined') {
      try {
        const rawV1 = localStorage.getItem('notaire_download_archive_v1');
        if (rawV1) {
          const parsed = JSON.parse(rawV1);
          if (Array.isArray(parsed)) extraLsDownloads = parsed;
        }
      } catch {
        // ignore invalid legacy JSON
      }
    }
    const allLegacy = [...legacyDownloads, ...extraLsDownloads];
    if (allLegacy.length > 0) {
      for (const dl of allLegacy) {
        if (dl && dl.id) {
          const rev = migrateDownloadArchiveItemToRevision(dl);
          await putInStore(STORE_REVISIONS, 'notary_revisions_v2', rev);
          await deleteFromStore(STORE_DOWNLOADS, 'notary_downloads_v2', dl.id).catch(
            () => {}
          );
        }
      }
    }
    if (typeof window !== 'undefined') {
      try {
        localStorage.removeItem('notary_downloads_v2');
        localStorage.removeItem('notaire_download_archive_v1');
      } catch {
        // ignore storage error
      }
    }
    setDownloadsMigrated(true);
  }

  const items = await getAllFromStore<DocumentRevision>(
    STORE_REVISIONS,
    'notary_revisions_v2'
  );
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDocumentRevision(
  rev: DocumentRevision
): Promise<void> {
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
  await putInStore(STORE_REVISIONS, 'notary_revisions_v2', normalizedRev);

  // Prune older revisions for the same documentId, keeping only the latest 25
  const allRevs = await loadDocumentRevisions();
  const sameDocRevs = allRevs
    .filter((r) => r.documentId === rev.documentId)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  if (sameDocRevs.length > 25) {
    const toDelete = sameDocRevs.slice(25);
    for (const oldRev of toDelete) {
      await deleteFromStore(STORE_REVISIONS, 'notary_revisions_v2', oldRev.id);
    }
  }
}

export async function deleteDocumentRevision(id: string): Promise<void> {
  await deleteFromStore(STORE_REVISIONS, 'notary_revisions_v2', id);
}

// 6. Saved Documents CRUD
export async function loadSavedDocuments(): Promise<SavedDocument[]> {
  const items = await getAllFromStore<SavedDocument>(STORE_DOCUMENTS, 'notary_documents_v1');
  const normalized = items.map((doc) => {
    const hasIdx = Boolean(doc.indexNumber || doc.fieldValues?.['رقم_الفهرس']);
    const normalizedStatus =
      !doc.status
        ? hasIdx
          ? 'signed_indexed'
          : 'draft'
        : doc.status === 'signed' && hasIdx
        ? 'signed_indexed'
        : doc.status;
    return {
      ...doc,
      status: normalizedStatus,
    };
  });
  return normalized.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveDocumentRecord(doc: SavedDocument): Promise<void> {
  await putInStore(STORE_DOCUMENTS, 'notary_documents_v1', doc);
}

export async function deleteDocumentRecord(id: string): Promise<void> {
  await deleteFromStore(STORE_DOCUMENTS, 'notary_documents_v1', id);
}

// 7. Subdivision Estates CRUD
export async function loadSubdivisionEstates(): Promise<SubdivisionEstate[]> {
  const items = await getAllFromStore<SubdivisionEstate>(STORE_ESTATES, 'notary_estates_v1');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveSubdivisionEstate(estate: SubdivisionEstate): Promise<void> {
  await putInStore(STORE_ESTATES, 'notary_estates_v1', estate);
}

export async function deleteSubdivisionEstate(id: string): Promise<void> {
  await deleteFromStore(STORE_ESTATES, 'notary_estates_v1', id);
}

// 7b. Saved Parties & Properties Directory
export async function loadSavedParties(): Promise<SavedPartyRecord[]> {
  const items = await getAllFromStore<SavedPartyRecord>(
    STORE_SAVED_PARTIES,
    'notary_saved_parties_v25'
  );
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePartyRecord(party: SavedPartyRecord): Promise<void> {
  await putInStore(STORE_SAVED_PARTIES, 'notary_saved_parties_v25', party);
}

export async function deletePartyRecord(id: string): Promise<void> {
  await deleteFromStore(STORE_SAVED_PARTIES, 'notary_saved_parties_v25', id);
}

export async function loadSavedProperties(): Promise<SavedPropertyRecord[]> {
  const items = await getAllFromStore<SavedPropertyRecord>(
    STORE_SAVED_PROPERTIES,
    'notary_saved_properties_v25'
  );
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePropertyRecord(prop: SavedPropertyRecord): Promise<void> {
  await putInStore(STORE_SAVED_PROPERTIES, 'notary_saved_properties_v25', prop);
}

export async function deletePropertyRecord(id: string): Promise<void> {
  await deleteFromStore(STORE_SAVED_PROPERTIES, 'notary_saved_properties_v25', id);
}

// 7c. Clerks & Folders
export async function loadNotaryClerks(): Promise<NotaryClerk[]> {
  const items = await getAllFromStore<NotaryClerk>(STORE_CLERKS, 'notary_clerks_v1');
  if (items.length === 0) {
    for (const c of DEFAULT_NOTARY_CLERKS) {
      await putInStore(STORE_CLERKS, 'notary_clerks_v1', c);
    }
    return DEFAULT_NOTARY_CLERKS;
  }
  return items;
}

export async function saveNotaryClerk(clerk: NotaryClerk): Promise<void> {
  await putInStore(STORE_CLERKS, 'notary_clerks_v1', clerk);
}

export async function deleteNotaryClerk(id: string): Promise<void> {
  await deleteFromStore(STORE_CLERKS, 'notary_clerks_v1', id);
}

export async function loadContractFolders(): Promise<ContractFolder[]> {
  const items = await getAllFromStore<ContractFolder>(STORE_FOLDERS, 'notary_folders_v1');
  return items.sort((a, b) => a.name.localeCompare(b.name));
}

export async function saveContractFolder(folder: ContractFolder): Promise<void> {
  await putInStore(STORE_FOLDERS, 'notary_folders_v1', folder);
}

export async function deleteContractFolder(id: string): Promise<void> {
  await deleteFromStore(STORE_FOLDERS, 'notary_folders_v1', id);
}

// 7d. Registration Batches & Land Publication Tracks
export async function loadRegistrationBatches(): Promise<RegistrationBatch[]> {
  const items = await getAllFromStore<RegistrationBatch>(
    STORE_REGISTRATION_BATCHES,
    'notary_registration_batches_v1'
  );
  return items.sort((a, b) => b.statusDate.localeCompare(a.statusDate));
}

export async function saveRegistrationBatch(batch: RegistrationBatch): Promise<void> {
  await putInStore(STORE_REGISTRATION_BATCHES, 'notary_registration_batches_v1', batch);
}

export async function deleteRegistrationBatch(id: string): Promise<void> {
  await deleteFromStore(STORE_REGISTRATION_BATCHES, 'notary_registration_batches_v1', id);
}

export async function loadPublicationTracks(): Promise<PublicationTrack[]> {
  const items = await getAllFromStore<PublicationTrack>(
    STORE_PUBLICATION_TRACKS,
    'notary_publication_tracks_v1'
  );
  return items.sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
}

export async function savePublicationTrack(track: PublicationTrack): Promise<void> {
  await putInStore(STORE_PUBLICATION_TRACKS, 'notary_publication_tracks_v1', track);
}

export async function deletePublicationTrack(id: string): Promise<void> {
  await deleteFromStore(STORE_PUBLICATION_TRACKS, 'notary_publication_tracks_v1', id);
}

// 9. Active Draft Session
export async function loadActiveDraftSession(): Promise<SavedDocument | null> {
  if (typeof window === 'undefined') return null;
  try {
    const raw = localStorage.getItem(LS_ACTIVE_DOC_KEY);
    return raw ? (JSON.parse(raw) as SavedDocument) : null;
  } catch {
    return null;
  }
}

export async function saveActiveDraftSession(doc: SavedDocument): Promise<void> {
  if (typeof window === 'undefined') return;
  localStorage.setItem(LS_ACTIVE_DOC_KEY, JSON.stringify(doc));
}

// 10. Full Backup Export & Import (including optional Base64 sealedOriginals)
export async function exportFullBackupBundle(): Promise<BackupBundle> {
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
    rawSealedOriginals,
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
    getAllSealedOriginals(),
  ]);

  const templates = officeTemplates
    .filter((t) => t.kind === 'contract')
    .map(officeTemplateToCustomTemplate);
  const derivedTemplates = officeTemplates
    .filter((t) => t.kind !== 'contract')
    .map(officeTemplateToDerivedDocTemplate);

  const sealedOriginals: SerializedSealedOriginal[] = rawSealedOriginals.map((item) => ({
    documentId: item.documentId,
    relativePath: item.relativePath,
    sha256: item.sha256,
    sealedAt: item.sealedAt,
    byteLength: item.byteLength,
    base64Data:
      item.base64Data || (item.bytes ? uint8ArrayToBase64(item.bytes) : undefined),
  }));

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
  let templatesCount = 0;
  let documentsCount = 0;
  let estatesCount = 0;
  let clausesCount = 0;

  if (Array.isArray(bundle.clerks)) {
    for (const cl of bundle.clerks) {
      if (cl && cl.id && cl.name) {
        await saveNotaryClerk(cl);
      }
    }
  }

  if (Array.isArray(bundle.folders)) {
    for (const f of bundle.folders) {
      if (f && f.id && f.name) {
        await saveContractFolder(f);
      }
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

  if (importedTemplateIds.size > 0) {
    setOfficeTemplatesMigrated(true);
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
      if (rev && rev.id) {
        await saveDocumentRevision(rev);
      }
    }
  }

  if (Array.isArray(bundle.downloads)) {
    for (const dl of bundle.downloads) {
      if (dl && dl.id) {
        await saveDocumentRevision(migrateDownloadArchiveItemToRevision(dl));
      }
    }
    setDownloadsMigrated(true);
  }

  if (Array.isArray(bundle.savedParties)) {
    for (const p of bundle.savedParties) {
      if (p && p.id && p.fullName) {
        await savePartyRecord(p);
      }
    }
  }

  if (Array.isArray(bundle.savedProperties)) {
    for (const prop of bundle.savedProperties) {
      if (prop && prop.id && prop.label) {
        await savePropertyRecord(prop);
      }
    }
  }

  if (Array.isArray(bundle.registrationBatches)) {
    for (const batch of bundle.registrationBatches) {
      if (batch && batch.id) {
        await saveRegistrationBatch(batch);
      }
    }
  }

  if (Array.isArray(bundle.publicationTracks)) {
    for (const track of bundle.publicationTracks) {
      if (track && track.id && track.documentId) {
        await savePublicationTrack(track);
      }
    }
  }

  if (Array.isArray(bundle.sealedOriginals)) {
    const existingSealed = await getAllSealedOriginals();
    for (const so of bundle.sealedOriginals) {
      if (so && so.documentId && so.relativePath && so.sha256) {
        const conflict = existingSealed.find(
          (item) =>
            (item.documentId === so.documentId || item.relativePath === so.relativePath) &&
            item.sha256 !== so.sha256
        );
        if (conflict) {
          // Never overwrite an already sealed original if the SHA-256 fingerprint differs
          continue;
        }
        const bytes = so.base64Data ? base64ToUint8Array(so.base64Data) : undefined;
        await putSealedOriginalRecord({
          documentId: so.documentId,
          relativePath: so.relativePath,
          sha256: so.sha256,
          sealedAt: so.sealedAt || new Date().toISOString(),
          byteLength: so.byteLength || (bytes ? bytes.byteLength : 0),
          bytes,
        });
      }
    }
  }

  return { templatesCount, documentsCount, estatesCount, clausesCount };
}

export const browserOfficeStore: OfficeStore = {
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
