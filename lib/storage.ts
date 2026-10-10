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
  SubdivisionEstate,
  WordTemplateDefinition,
} from './types';
import { getDesktopBridge } from './desktop-bridge';
import {
  DEFAULT_DERIVED_DOC_TEMPLATES,
  DEFAULT_NOTARY_CLERKS,
  DEFAULT_PARTY_FIELDS,
  OfficeStore,
  OfficeStoreRuntimeInfo,
  ReadSealedOriginalResult,
  SealContractOriginalInput,
  SealContractOriginalResult,
} from './office-store';
import { browserOfficeStore } from './storage-browser';
import { desktopOfficeStore } from './storage-desktop';

export {
  DEFAULT_DERIVED_DOC_TEMPLATES,
  DEFAULT_NOTARY_CLERKS,
  DEFAULT_PARTY_FIELDS,
};

export function getActiveOfficeStore(): OfficeStore {
  return getDesktopBridge() ? desktopOfficeStore : browserOfficeStore;
}

export function getRuntimeInfo(): OfficeStoreRuntimeInfo {
  return getActiveOfficeStore().getRuntimeInfo();
}

export function sealContractOriginal(
  input: SealContractOriginalInput
): Promise<SealContractOriginalResult> {
  return getActiveOfficeStore().sealContractOriginal(input);
}

export function readSealedOriginal(
  documentId: string
): Promise<ReadSealedOriginalResult | null> {
  return getActiveOfficeStore().readSealedOriginal(documentId);
}

// 1. Clauses
export function loadNotaryClauses(): Promise<NotaryClause[]> {
  return getActiveOfficeStore().loadNotaryClauses();
}

export function saveNotaryClause(clause: NotaryClause): Promise<void> {
  return getActiveOfficeStore().saveNotaryClause(clause);
}

export function deleteNotaryClause(id: string): Promise<void> {
  return getActiveOfficeStore().deleteNotaryClause(id);
}

// 2. Derived Document Templates
export function loadDerivedDocTemplates(): Promise<DerivedDocTemplate[]> {
  return getActiveOfficeStore().loadDerivedDocTemplates();
}

export function saveDerivedDocTemplate(tpl: DerivedDocTemplate): Promise<void> {
  return getActiveOfficeStore().saveDerivedDocTemplate(tpl);
}

export function deleteDerivedDocTemplate(id: string): Promise<void> {
  return getActiveOfficeStore().deleteDerivedDocTemplate(id);
}

// 2b. Word Templates
export function loadWordTemplates(): Promise<WordTemplateDefinition[]> {
  return getActiveOfficeStore().loadWordTemplates();
}

export function saveWordTemplate(tpl: WordTemplateDefinition): Promise<void> {
  return getActiveOfficeStore().saveWordTemplate(tpl);
}

export function deleteWordTemplate(id: string): Promise<void> {
  return getActiveOfficeStore().deleteWordTemplate(id);
}

// 3. Document Revisions
export function loadDocumentRevisions(): Promise<DocumentRevision[]> {
  return getActiveOfficeStore().loadDocumentRevisions();
}

export function saveDocumentRevision(rev: DocumentRevision): Promise<void> {
  return getActiveOfficeStore().saveDocumentRevision(rev);
}

export function deleteDocumentRevision(id: string): Promise<void> {
  return getActiveOfficeStore().deleteDocumentRevision(id);
}

// 4. Download Archive
export function loadDownloadArchive(): Promise<DownloadArchiveItem[]> {
  return getActiveOfficeStore().loadDownloadArchive();
}

export function saveDownloadArchiveItem(item: DownloadArchiveItem): Promise<void> {
  return getActiveOfficeStore().saveDownloadArchiveItem(item);
}

export function deleteDownloadArchiveItem(id: string): Promise<void> {
  return getActiveOfficeStore().deleteDownloadArchiveItem(id);
}

export function clearDownloadArchive(): Promise<void> {
  return getActiveOfficeStore().clearDownloadArchive();
}

// 5. Custom Templates
export function loadCustomTemplates(): Promise<CustomTemplate[]> {
  return getActiveOfficeStore().loadCustomTemplates();
}

export function saveCustomTemplate(template: CustomTemplate): Promise<void> {
  return getActiveOfficeStore().saveCustomTemplate(template);
}

export function deleteCustomTemplate(id: string): Promise<void> {
  return getActiveOfficeStore().deleteCustomTemplate(id);
}

export function exportCustomTemplatesJson(): Promise<string> {
  return getActiveOfficeStore().exportCustomTemplatesJson();
}

export function importCustomTemplatesJson(templates: CustomTemplate[]): Promise<number> {
  return getActiveOfficeStore().importCustomTemplatesJson(templates);
}

// 6. Saved Documents
export function loadSavedDocuments(): Promise<SavedDocument[]> {
  return getActiveOfficeStore().loadSavedDocuments();
}

export function saveDocumentRecord(doc: SavedDocument): Promise<void> {
  return getActiveOfficeStore().saveDocumentRecord(doc);
}

export function deleteDocumentRecord(id: string): Promise<void> {
  return getActiveOfficeStore().deleteDocumentRecord(id);
}

// 7. Subdivision Estates
export function loadSubdivisionEstates(): Promise<SubdivisionEstate[]> {
  return getActiveOfficeStore().loadSubdivisionEstates();
}

export function saveSubdivisionEstate(estate: SubdivisionEstate): Promise<void> {
  return getActiveOfficeStore().saveSubdivisionEstate(estate);
}

export function deleteSubdivisionEstate(id: string): Promise<void> {
  return getActiveOfficeStore().deleteSubdivisionEstate(id);
}

// 7b. Saved Parties & Properties Directory
export function loadSavedParties(): Promise<SavedPartyRecord[]> {
  return getActiveOfficeStore().loadSavedParties();
}

export function savePartyRecord(party: SavedPartyRecord): Promise<void> {
  return getActiveOfficeStore().savePartyRecord(party);
}

export function deletePartyRecord(id: string): Promise<void> {
  return getActiveOfficeStore().deletePartyRecord(id);
}

export function loadSavedProperties(): Promise<SavedPropertyRecord[]> {
  return getActiveOfficeStore().loadSavedProperties();
}

export function savePropertyRecord(prop: SavedPropertyRecord): Promise<void> {
  return getActiveOfficeStore().savePropertyRecord(prop);
}

export function deletePropertyRecord(id: string): Promise<void> {
  return getActiveOfficeStore().deletePropertyRecord(id);
}

// 7c. Clerks & Folders
export function loadNotaryClerks(): Promise<NotaryClerk[]> {
  return getActiveOfficeStore().loadNotaryClerks();
}

export function saveNotaryClerk(clerk: NotaryClerk): Promise<void> {
  return getActiveOfficeStore().saveNotaryClerk(clerk);
}

export function deleteNotaryClerk(id: string): Promise<void> {
  return getActiveOfficeStore().deleteNotaryClerk(id);
}

export function loadContractFolders(): Promise<ContractFolder[]> {
  return getActiveOfficeStore().loadContractFolders();
}

export function saveContractFolder(folder: ContractFolder): Promise<void> {
  return getActiveOfficeStore().saveContractFolder(folder);
}

export function deleteContractFolder(id: string): Promise<void> {
  return getActiveOfficeStore().deleteContractFolder(id);
}

// 8. Party Fields
export function loadPartyFields(): PartyField[] {
  return getActiveOfficeStore().loadPartyFields();
}

export function savePartyFields(fields: PartyField[]): void {
  getActiveOfficeStore().savePartyFields(fields);
}

// 9. Active Draft Session
export function loadActiveDraftSession(): SavedDocument | null {
  return getActiveOfficeStore().loadActiveDraftSession();
}

export function saveActiveDraftSession(doc: SavedDocument): void {
  getActiveOfficeStore().saveActiveDraftSession(doc);
}

// 10. Full Backup Export & Import
export function exportFullBackupBundle(): Promise<BackupBundle> {
  return getActiveOfficeStore().exportFullBackupBundle();
}

export function importFullBackupBundle(bundle: BackupBundle): Promise<{
  templatesCount: number;
  documentsCount: number;
  estatesCount: number;
  clausesCount: number;
}> {
  return getActiveOfficeStore().importFullBackupBundle(bundle);
}
