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
  SerializedSealedOriginal,
  SubdivisionEstate,
  WordTemplateDefinition,
} from './types';
import { DEFAULT_WORD_TEMPLATES } from './docx-template-engine';
import { getDesktopBridge, DesktopOfficeBridge } from './desktop-bridge';
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

  return {
    relativePath: res.relativePath,
    sha256: res.sha256,
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

// 2. Derived Document Templates CRUD
export async function loadDerivedDocTemplates(): Promise<DerivedDocTemplate[]> {
  const saved = await desktopLoad<DerivedDocTemplate>('derived_templates');
  if (saved.length === 0) {
    for (const defTpl of DEFAULT_DERIVED_DOC_TEMPLATES) {
      await desktopSave('derived_templates', defTpl);
    }
    return DEFAULT_DERIVED_DOC_TEMPLATES;
  }
  return saved;
}

export async function saveDerivedDocTemplate(tpl: DerivedDocTemplate): Promise<void> {
  await desktopSave('derived_templates', tpl);
}

export async function deleteDerivedDocTemplate(id: string): Promise<void> {
  await desktopDelete('derived_templates', id);
}

// 2b. Word Templates CRUD
export async function loadWordTemplates(): Promise<WordTemplateDefinition[]> {
  const saved = await desktopLoad<WordTemplateDefinition>('word_templates');
  if (saved.length === 0) {
    for (const defTpl of DEFAULT_WORD_TEMPLATES) {
      await desktopSave('word_templates', defTpl);
    }
    return DEFAULT_WORD_TEMPLATES;
  }
  return saved;
}

export async function saveWordTemplate(tpl: WordTemplateDefinition): Promise<void> {
  await desktopSave('word_templates', tpl);
}

export async function deleteWordTemplate(id: string): Promise<void> {
  await desktopDelete('word_templates', id);
}

// 3. Document Revisions CRUD (Keeps latest 20 revisions per documentId)
export async function loadDocumentRevisions(): Promise<DocumentRevision[]> {
  const items = await desktopLoad<DocumentRevision>('revisions');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDocumentRevision(rev: DocumentRevision): Promise<void> {
  await desktopSave('revisions', rev);
  const allRevs = await loadDocumentRevisions();
  const sameDocRevs = allRevs
    .filter((r) => r.documentId === rev.documentId)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  if (sameDocRevs.length > 20) {
    for (const oldRev of sameDocRevs.slice(20)) {
      await desktopDelete('revisions', oldRev.id);
    }
  }
}

export async function deleteDocumentRevision(id: string): Promise<void> {
  await desktopDelete('revisions', id);
}

// 4. Download Archive CRUD
export async function loadDownloadArchive(): Promise<DownloadArchiveItem[]> {
  const items = await desktopLoad<DownloadArchiveItem>('downloads');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDownloadArchiveItem(item: DownloadArchiveItem): Promise<void> {
  await desktopSave('downloads', item);
}

export async function deleteDownloadArchiveItem(id: string): Promise<void> {
  await desktopDelete('downloads', id);
}

export async function clearDownloadArchive(): Promise<void> {
  await desktopClear('downloads');
}

// 5. Custom Templates CRUD
export async function loadCustomTemplates(): Promise<CustomTemplate[]> {
  const items = await desktopLoad<CustomTemplate>('templates');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveCustomTemplate(template: CustomTemplate): Promise<void> {
  await desktopSave('templates', template);
}

export async function deleteCustomTemplate(id: string): Promise<void> {
  await desktopDelete('templates', id);
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

// 8. Party Fields (Async from Disk / SQLite)
export async function loadPartyFields(): Promise<PartyField[]> {
  const bridge = requireBridge();
  const res = await bridge.invoke<
    PartyField[] | Array<{ id: string; fields?: PartyField[] }> | null
  >('storage_load', { collection: 'party_fields' });

  let parsed: PartyField[] = [];
  if (Array.isArray(res) && res.length > 0) {
    const first = res[0] as { fields?: PartyField[]; key?: string };
    if (Array.isArray(first.fields)) {
      parsed = first.fields;
    } else if (typeof first.key === 'string') {
      parsed = res as PartyField[];
    }
  }

  if (parsed.length === 0) return DEFAULT_PARTY_FIELDS;
  const existingKeys = new Set(parsed.map((f) => f.key));
  const merged = [...parsed];
  for (const defField of DEFAULT_PARTY_FIELDS) {
    if (!existingKeys.has(defField.key)) {
      merged.push(defField);
    }
  }
  return merged;
}

export async function savePartyFields(fields: PartyField[]): Promise<void> {
  const bridge = requireBridge();
  await bridge.invoke<void>('storage_save', {
    collection: 'party_fields',
    item: { id: 'default_party_fields', fields },
  });
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
    sealedOriginals,
    defaultFields,
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
    desktopLoad<SerializedSealedOriginal>('sealed_originals'),
    loadPartyFields(),
  ]);

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
      if (dt && dt.id && dt.name) await saveDerivedDocTemplate(dt);
    }
  }
  if (Array.isArray(bundle.revisions)) {
    for (const rev of bundle.revisions) {
      if (rev && rev.id) await saveDocumentRevision(rev);
    }
  }
  if (Array.isArray(bundle.downloads)) {
    for (const dl of bundle.downloads) {
      if (dl && dl.id) await saveDownloadArchiveItem(dl);
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
  if (Array.isArray(bundle.defaultFields) && bundle.defaultFields.length > 0) {
    await savePartyFields(bundle.defaultFields);
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
