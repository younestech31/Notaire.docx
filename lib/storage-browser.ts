import {
  BackupBundle,
  ContractFolder,
  CustomTemplate,
  DerivedDocTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  NotaryClerk,
  NotaryClause,
  PartyField,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SealedContractOriginal,
  SerializedSealedOriginal,
  SubdivisionEstate,
  WordTemplateDefinition,
} from './types';
import { DEFAULT_WORD_TEMPLATES } from './docx-template-engine';
import {
  base64ToUint8Array,
  buildArchiveRelativePath,
  computeSha256Hex,
  DEFAULT_DERIVED_DOC_TEMPLATES,
  DEFAULT_NOTARY_CLERKS,
  DEFAULT_PARTY_FIELDS,
  OfficeStore,
  OfficeStoreRuntimeInfo,
  ReadSealedOriginalResult,
  SealContractOriginalInput,
  SealContractOriginalResult,
  uint8ArrayToBase64,
} from './office-store';

const DB_NAME = 'NotarySmartEditorDB';
const DB_VERSION = 6;
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

const LS_FIELDS_KEY = 'notary_default_party_fields_v2';
const LS_ACTIVE_DOC_KEY = 'notary_active_document_v2';
const LS_SEALED_ORIGINALS_KEY = 'notary_sealed_originals_v1';

interface StoredSealedOriginalRecord extends SealedContractOriginal {
  bytes?: Uint8Array;
  base64Data?: string;
}

function openDatabase(): Promise<IDBDatabase | null> {
  if (typeof window === 'undefined' || !('indexedDB' in window)) {
    return Promise.resolve(null);
  }
  return new Promise((resolve, reject) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        const storesWithIdKey = [
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
      request.onerror = () =>
        reject(
          new Error(
            `تعذر فتح قاعدة بيانات المتصفح المحلية: ${
              request.error?.message || 'خطأ غير معروف'
            }`
          )
        );
      request.onblocked = () =>
        reject(new Error('قاعدة بيانات المتصفح محجوبة بواسطة علامة تبويب أخرى'));
    } catch (err) {
      reject(
        err instanceof Error
          ? err
          : new Error('تعذر تهيئة قاعدة بيانات المتصفح المحلية')
      );
    }
  });
}

async function getAllFromStore<T>(storeName: string, lsFallbackKey: string): Promise<T[]> {
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(storeName)) {
    const dbResult = await new Promise<T[]>((resolve, reject) => {
      try {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.getAll();
        req.onsuccess = () => resolve((req.result as T[]) || []);
        req.onerror = () =>
          reject(
            new Error(
              `فشل قراءة المخزن (${storeName}): ${req.error?.message || ''}`
            )
          );
      } catch (err) {
        reject(err instanceof Error ? err : new Error(`فشل قراءة المخزن (${storeName})`));
      }
    });
    if (dbResult.length > 0) {
      return dbResult;
    }
  }
  if (typeof window !== 'undefined') {
    const raw = localStorage.getItem(lsFallbackKey);
    if (!raw) return [];
    return JSON.parse(raw) as T[];
  }
  return [];
}

async function putInStore<T extends { id: string }>(
  storeName: string,
  lsFallbackKey: string,
  item: T
): Promise<void> {
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(storeName)) {
    await new Promise<void>((resolve, reject) => {
      try {
        const tx = db.transaction(storeName, 'readwrite');
        tx.objectStore(storeName).put(item);
        tx.oncomplete = () => resolve();
        tx.onerror = () =>
          reject(
            new Error(
              `فشل حفظ السجل في (${storeName}): ${
                tx.error?.message || 'خطأ في معاملة التخزين'
              }`
            )
          );
        tx.onabort = () =>
          reject(
            new Error(
              `تم إحباط معاملة الحفظ في (${storeName}): ${
                tx.error?.message || 'تجاوز السعة أو خطأ داخلي'
              }`
            )
          );
      } catch (err) {
        reject(
          err instanceof Error
            ? err
            : new Error(`فشل تنفيذ معاملة الحفظ في (${storeName})`)
        );
      }
    });
  }
  if (typeof window !== 'undefined') {
    try {
      const existing = await getAllFromStore<T>(storeName, lsFallbackKey);
      const idx = existing.findIndex((x) => x.id === item.id);
      if (idx >= 0) existing[idx] = item;
      else existing.unshift(item);
      localStorage.setItem(lsFallbackKey, JSON.stringify(existing));
    } catch {
      // Secondary localStorage mirror may hit quota on large docs; primary IDB transaction already committed
    }
  }
}

async function deleteFromStore(
  storeName: string,
  lsFallbackKey: string,
  id: string
): Promise<void> {
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(storeName)) {
    await new Promise<void>((resolve, reject) => {
      try {
        const tx = db.transaction(storeName, 'readwrite');
        tx.objectStore(storeName).delete(id);
        tx.oncomplete = () => resolve();
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
  if (typeof window !== 'undefined') {
    try {
      const existing = await getAllFromStore<{ id: string }>(storeName, lsFallbackKey);
      const filtered = existing.filter((x) => x.id !== id);
      localStorage.setItem(lsFallbackKey, JSON.stringify(filtered));
    } catch {
      // Secondary localStorage mirror ignore
    }
  }
}

export function getRuntimeInfo(): OfficeStoreRuntimeInfo {
  return {
    mode: 'browser',
    dataDirectory: null,
  };
}

// =========================================================
// SEALED ORIGINALS (الأصول المعتمدة ببصمة SHA-256)
// =========================================================
async function getAllSealedOriginals(): Promise<StoredSealedOriginalRecord[]> {
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(STORE_SEALED_ORIGINALS)) {
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
  if (typeof window !== 'undefined') {
    const raw = localStorage.getItem(LS_SEALED_ORIGINALS_KEY);
    return raw ? (JSON.parse(raw) as StoredSealedOriginalRecord[]) : [];
  }
  return [];
}

async function putSealedOriginalRecord(
  record: StoredSealedOriginalRecord
): Promise<void> {
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(STORE_SEALED_ORIGINALS)) {
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
  if (typeof window !== 'undefined') {
    try {
      const serialized: SerializedSealedOriginal = {
        documentId: record.documentId,
        relativePath: record.relativePath,
        sha256: record.sha256,
        sealedAt: record.sealedAt,
        byteLength: record.byteLength,
        base64Data:
          record.base64Data ||
          (record.bytes ? uint8ArrayToBase64(record.bytes) : undefined),
      };
      const raw = localStorage.getItem(LS_SEALED_ORIGINALS_KEY);
      const list: SerializedSealedOriginal[] = raw ? JSON.parse(raw) : [];
      const idx = list.findIndex((x) => x.documentId === record.documentId);
      if (idx >= 0) list[idx] = serialized;
      else list.unshift(serialized);
      localStorage.setItem(LS_SEALED_ORIGINALS_KEY, JSON.stringify(list));
    } catch {
      // Ignore secondary localStorage quota if bytes are large
    }
  }
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
    base64Data: uint8ArrayToBase64(input.docxBytes),
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

// 2. Derived Document Templates CRUD
export async function loadDerivedDocTemplates(): Promise<DerivedDocTemplate[]> {
  const saved = await getAllFromStore<DerivedDocTemplate>(
    STORE_DERIVED_TEMPLATES,
    'notary_derived_templates_v2'
  );
  if (saved.length === 0) {
    for (const defTpl of DEFAULT_DERIVED_DOC_TEMPLATES) {
      await putInStore(STORE_DERIVED_TEMPLATES, 'notary_derived_templates_v2', defTpl);
    }
    return DEFAULT_DERIVED_DOC_TEMPLATES;
  }
  return saved;
}

export async function saveDerivedDocTemplate(tpl: DerivedDocTemplate): Promise<void> {
  await putInStore(STORE_DERIVED_TEMPLATES, 'notary_derived_templates_v2', tpl);
}

export async function deleteDerivedDocTemplate(id: string): Promise<void> {
  await deleteFromStore(STORE_DERIVED_TEMPLATES, 'notary_derived_templates_v2', id);
}

// 2b. Word Templates CRUD
export async function loadWordTemplates(): Promise<WordTemplateDefinition[]> {
  const saved = await getAllFromStore<WordTemplateDefinition>(
    STORE_WORD_TEMPLATES,
    'notary_word_templates_v1'
  );
  if (saved.length === 0) {
    for (const defTpl of DEFAULT_WORD_TEMPLATES) {
      await putInStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', defTpl);
    }
    return DEFAULT_WORD_TEMPLATES;
  }
  return saved;
}

export async function saveWordTemplate(tpl: WordTemplateDefinition): Promise<void> {
  await putInStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', tpl);
}

export async function deleteWordTemplate(id: string): Promise<void> {
  await deleteFromStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', id);
}

// 3. Document Revisions CRUD (Keeps latest 20 revisions per documentId)
export async function loadDocumentRevisions(): Promise<DocumentRevision[]> {
  const items = await getAllFromStore<DocumentRevision>(STORE_REVISIONS, 'notary_revisions_v2');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDocumentRevision(rev: DocumentRevision): Promise<void> {
  await putInStore(STORE_REVISIONS, 'notary_revisions_v2', rev);

  // Prune older revisions for the same documentId, keeping only the latest 20
  const allRevs = await loadDocumentRevisions();
  const sameDocRevs = allRevs
    .filter((r) => r.documentId === rev.documentId)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));

  if (sameDocRevs.length > 20) {
    const toDelete = sameDocRevs.slice(20);
    for (const oldRev of toDelete) {
      await deleteFromStore(STORE_REVISIONS, 'notary_revisions_v2', oldRev.id);
    }
  }
}

export async function deleteDocumentRevision(id: string): Promise<void> {
  await deleteFromStore(STORE_REVISIONS, 'notary_revisions_v2', id);
}

// 4. Download Archive CRUD
export async function loadDownloadArchive(): Promise<DownloadArchiveItem[]> {
  const items = await getAllFromStore<DownloadArchiveItem>(STORE_DOWNLOADS, 'notary_downloads_v2');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDownloadArchiveItem(item: DownloadArchiveItem): Promise<void> {
  await putInStore(STORE_DOWNLOADS, 'notary_downloads_v2', item);
}

export async function deleteDownloadArchiveItem(id: string): Promise<void> {
  await deleteFromStore(STORE_DOWNLOADS, 'notary_downloads_v2', id);
}

export async function clearDownloadArchive(): Promise<void> {
  const items = await loadDownloadArchive();
  for (const item of items) {
    await deleteDownloadArchiveItem(item.id);
  }
}

// 5. Office Custom Templates CRUD
export async function loadCustomTemplates(): Promise<CustomTemplate[]> {
  const items = await getAllFromStore<CustomTemplate>(STORE_TEMPLATES, 'notary_templates_v1');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveCustomTemplate(template: CustomTemplate): Promise<void> {
  await putInStore(STORE_TEMPLATES, 'notary_templates_v1', template);
}

export async function deleteCustomTemplate(id: string): Promise<void> {
  await deleteFromStore(STORE_TEMPLATES, 'notary_templates_v1', id);
}

export async function exportCustomTemplatesJson(): Promise<string> {
  const templates = await loadCustomTemplates();
  return JSON.stringify(templates, null, 2);
}

export async function importCustomTemplatesJson(templates: CustomTemplate[]): Promise<number> {
  let count = 0;
  for (const tpl of templates) {
    if (tpl && tpl.name && (tpl.bodyHtml !== undefined || tpl.id)) {
      const sanitized: CustomTemplate = {
        id: tpl.id || `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        name: tpl.name,
        category: tpl.category || 'عام',
        description: tpl.description || '',
        bodyHtml: tpl.bodyHtml || '',
        headerHtml: tpl.headerHtml || '',
        footerHtml: tpl.footerHtml || '',
        pageNumberingEnabled: !!tpl.pageNumberingEnabled,
        extractedPlaceholders: Array.isArray(tpl.extractedPlaceholders) ? tpl.extractedPlaceholders : [],
        defaultFieldValues: tpl.defaultFieldValues || {},
        updatedAt: tpl.updatedAt || new Date().toISOString(),
      };
      await saveCustomTemplate(sanitized);
      count++;
    }
  }
  return count;
}

// 6. Saved Documents CRUD
export async function loadSavedDocuments(): Promise<SavedDocument[]> {
  const items = await getAllFromStore<SavedDocument>(STORE_DOCUMENTS, 'notary_documents_v1');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
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

// 8. Default / Custom Party Fields
export function loadPartyFields(): PartyField[] {
  if (typeof window === 'undefined') return DEFAULT_PARTY_FIELDS;
  try {
    const raw = localStorage.getItem(LS_FIELDS_KEY);
    if (!raw) return DEFAULT_PARTY_FIELDS;
    const parsed = JSON.parse(raw) as PartyField[];
    if (!Array.isArray(parsed) || parsed.length === 0) return DEFAULT_PARTY_FIELDS;
    const existingKeys = new Set(parsed.map((f) => f.key));
    const merged = parsed.map((f) => {
      if ((f.key === 'تاريخ_العقد' || f.key === 'رقم_الفهرس') && f.category === 'custom') {
        return { ...f, category: 'index' as const };
      }
      return f;
    });
    for (const defField of DEFAULT_PARTY_FIELDS) {
      if (!existingKeys.has(defField.key)) {
        merged.push(defField);
      }
    }
    return merged;
  } catch {
    return DEFAULT_PARTY_FIELDS;
  }
}

export function savePartyFields(fields: PartyField[]): void {
  if (typeof window === 'undefined') return;
  localStorage.setItem(LS_FIELDS_KEY, JSON.stringify(fields));
}

// 9. Active Draft Session
export function loadActiveDraftSession(): SavedDocument | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = localStorage.getItem(LS_ACTIVE_DOC_KEY);
    return raw ? (JSON.parse(raw) as SavedDocument) : null;
  } catch {
    return null;
  }
}

export function saveActiveDraftSession(doc: SavedDocument): void {
  if (typeof window === 'undefined') return;
  localStorage.setItem(LS_ACTIVE_DOC_KEY, JSON.stringify(doc));
}

// 10. Full Backup Export & Import (including optional Base64 sealedOriginals)
export async function exportFullBackupBundle(): Promise<BackupBundle> {
  const [
    templates,
    documents,
    estates,
    clauses,
    derivedTemplates,
    revisions,
    downloads,
    savedParties,
    savedProperties,
    clerks,
    folders,
    rawSealedOriginals,
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
    getAllSealedOriginals(),
  ]);
  const defaultFields = loadPartyFields();
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
    version: '2.6.0',
    exportedAt: new Date().toISOString(),
    clerks,
    folders,
    templates,
    documents,
    estates,
    clauses,
    derivedTemplates,
    revisions,
    downloads,
    savedParties,
    savedProperties,
    defaultFields,
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

  if (Array.isArray(bundle.templates)) {
    for (const tpl of bundle.templates) {
      if (tpl && tpl.id && tpl.name) {
        await saveCustomTemplate(tpl);
        templatesCount++;
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

  if (Array.isArray(bundle.derivedTemplates)) {
    for (const dt of bundle.derivedTemplates) {
      if (dt && dt.id && dt.name) {
        await saveDerivedDocTemplate(dt);
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
        await saveDownloadArchiveItem(dl);
      }
    }
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

  if (Array.isArray(bundle.sealedOriginals)) {
    for (const so of bundle.sealedOriginals) {
      if (so && so.documentId && so.relativePath && so.sha256) {
        const bytes = so.base64Data ? base64ToUint8Array(so.base64Data) : undefined;
        await putSealedOriginalRecord({
          documentId: so.documentId,
          relativePath: so.relativePath,
          sha256: so.sha256,
          sealedAt: so.sealedAt || new Date().toISOString(),
          byteLength: so.byteLength || (bytes ? bytes.byteLength : 0),
          bytes,
          base64Data: so.base64Data,
        });
      }
    }
  }

  if (Array.isArray(bundle.defaultFields) && bundle.defaultFields.length > 0) {
    savePartyFields(bundle.defaultFields);
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
  loadDerivedDocTemplates,
  saveDerivedDocTemplate,
  deleteDerivedDocTemplate,
  loadWordTemplates,
  saveWordTemplate,
  deleteWordTemplate,
  loadDocumentRevisions,
  saveDocumentRevision,
  deleteDocumentRevision,
  loadDownloadArchive,
  saveDownloadArchiveItem,
  deleteDownloadArchiveItem,
  clearDownloadArchive,
  loadCustomTemplates,
  saveCustomTemplate,
  deleteCustomTemplate,
  exportCustomTemplatesJson,
  importCustomTemplatesJson,
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
  loadPartyFields,
  savePartyFields,
  loadActiveDraftSession,
  saveActiveDraftSession,
  exportFullBackupBundle,
  importFullBackupBundle,
};
