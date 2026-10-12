import type {
  BackupBundle,
  ContractFolder,
  DocumentRevision,
  NotaryClerk,
  NotaryClause,
  OfficeTemplate,
  PublicationTrack,
  RegistrationBatch,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
} from './types';
import { getDesktopBridge } from './desktop-bridge';
import {
  DEFAULT_NOTARY_CLERKS,
  DEFAULT_OFFICE_TEMPLATES,
  type OfficeStore,
  type OfficeStoreRuntimeInfo,
  type ReadSealedOriginalResult,
  type SealContractOriginalInput,
  type SealContractOriginalResult,
} from './office-store';
import { browserOfficeStore } from './storage-browser';
import { desktopOfficeStore } from './storage-desktop';

export {
  DEFAULT_NOTARY_CLERKS,
  DEFAULT_OFFICE_TEMPLATES,
};

export function getActiveOfficeStore(): OfficeStore {
  return getDesktopBridge() ? desktopOfficeStore : browserOfficeStore;
}

export function getRuntimeInfo(): Promise<OfficeStoreRuntimeInfo> {
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

// 2. Unified Office Templates
export function loadOfficeTemplates(): Promise<OfficeTemplate[]> {
  return getActiveOfficeStore().loadOfficeTemplates();
}

export function saveOfficeTemplate(template: OfficeTemplate): Promise<void> {
  return getActiveOfficeStore().saveOfficeTemplate(template);
}

export function deleteOfficeTemplate(id: string): Promise<void> {
  return getActiveOfficeStore().deleteOfficeTemplate(id);
}

export function exportOfficeTemplatesJson(): Promise<string> {
  return getActiveOfficeStore().exportOfficeTemplatesJson();
}

export function importOfficeTemplatesJson(templates: unknown[]): Promise<number> {
  return getActiveOfficeStore().importOfficeTemplatesJson(templates);
}

// 3. Unified Document Revisions & Export History
export function loadDocumentRevisions(): Promise<DocumentRevision[]> {
  return getActiveOfficeStore().loadDocumentRevisions();
}

export function saveDocumentRevision(rev: DocumentRevision): Promise<void> {
  return getActiveOfficeStore().saveDocumentRevision(rev);
}

export function deleteDocumentRevision(id: string): Promise<void> {
  return getActiveOfficeStore().deleteDocumentRevision(id);
}

// 4. Saved Documents
export function loadSavedDocuments(): Promise<SavedDocument[]> {
  return getActiveOfficeStore().loadSavedDocuments();
}

export function saveDocumentRecord(doc: SavedDocument): Promise<void> {
  return getActiveOfficeStore().saveDocumentRecord(doc);
}

export function deleteDocumentRecord(id: string): Promise<void> {
  return getActiveOfficeStore().deleteDocumentRecord(id);
}

// 5. Subdivision Estates
export function loadSubdivisionEstates(): Promise<SubdivisionEstate[]> {
  return getActiveOfficeStore().loadSubdivisionEstates();
}

export function saveSubdivisionEstate(estate: SubdivisionEstate): Promise<void> {
  return getActiveOfficeStore().saveSubdivisionEstate(estate);
}

export function deleteSubdivisionEstate(id: string): Promise<void> {
  return getActiveOfficeStore().deleteSubdivisionEstate(id);
}

// 6. Saved Parties & Properties Directory
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

// 7. Clerks & Folders
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

// 8. Registration Batches & Land Publication Tracks
export function loadRegistrationBatches(): Promise<RegistrationBatch[]> {
  return getActiveOfficeStore().loadRegistrationBatches();
}

export function saveRegistrationBatch(batch: RegistrationBatch): Promise<void> {
  return getActiveOfficeStore().saveRegistrationBatch(batch);
}

export function deleteRegistrationBatch(id: string): Promise<void> {
  return getActiveOfficeStore().deleteRegistrationBatch(id);
}

export function loadPublicationTracks(): Promise<PublicationTrack[]> {
  return getActiveOfficeStore().loadPublicationTracks();
}

export function savePublicationTrack(track: PublicationTrack): Promise<void> {
  return getActiveOfficeStore().savePublicationTrack(track);
}

export function deletePublicationTrack(id: string): Promise<void> {
  return getActiveOfficeStore().deletePublicationTrack(id);
}

// 9. Active Draft Session
export function loadActiveDraftSession(): Promise<SavedDocument | null> {
  return getActiveOfficeStore().loadActiveDraftSession();
}

export function saveActiveDraftSession(doc: SavedDocument): Promise<void> {
  return getActiveOfficeStore().saveActiveDraftSession(doc);
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
