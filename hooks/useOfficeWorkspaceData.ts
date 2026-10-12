'use client';

import { useCallback, useRef, useState } from 'react';
import {
  ContractFolder,
  DocumentRevision,
  NotaryClerk,
  NotaryClause,
  OfficeTemplate,
  PartyField,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
} from '@/lib/types';
import {
  DEFAULT_OFFICE_TEMPLATES,
  loadActiveDraftSession,
  loadContractFolders,
  loadDocumentRevisions,
  loadNotaryClauses,
  loadNotaryClerks,
  loadOfficeTemplates,
  loadPartyFields,
  loadSavedDocuments,
  loadSavedParties,
  loadSavedProperties,
  loadSubdivisionEstates,
} from '@/lib/storage';
import {
  type LocalOfficeDirectoryState,
  reconnectPersistedOfficeDirectory,
} from '@/lib/desktop-bridge';

export function useOfficeWorkspaceData() {
  const [templates, setTemplates] = useState<OfficeTemplate[]>(
    DEFAULT_OFFICE_TEMPLATES
  );
  const [documents, setDocuments] = useState<SavedDocument[]>([]);
  const documentsRef = useRef<SavedDocument[]>([]);
  const [revisions, setRevisions] = useState<DocumentRevision[]>([]);
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
  const [partyFields, setPartyFields] = useState<PartyField[]>([]);
  const [lastBackupAt, setLastBackupAt] = useState<string | null>(null);
  const [localDirState, setLocalDirState] = useState<LocalOfficeDirectoryState>({
    status: 'disconnected',
    folderName: null,
  });

  const updateDocumentsState = useCallback((nextDocs: SavedDocument[]) => {
    documentsRef.current = nextDocs;
    setDocuments(nextDocs);
  }, []);

  const reloadAllOfficeRecords = useCallback(async () => {
    const [
      loadedOfficeTemplates,
      loadedDocs,
      loadedEstates,
      loadedClauses,
      loadedRevs,
      loadedSavedParties,
      loadedSavedProps,
      loadedClerks,
      loadedFolders,
      loadedFields,
      activeDraft,
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
      loadPartyFields(),
      loadActiveDraftSession(),
    ]);

    setTemplates(loadedOfficeTemplates);
    updateDocumentsState(loadedDocs);
    setEstates(loadedEstates);
    setClauses(loadedClauses);
    setRevisions(loadedRevs);
    setSavedParties(loadedSavedParties);
    setSavedProperties(loadedSavedProps);
    setClerks(loadedClerks);
    if (loadedClerks.length > 0) {
      setActiveClerk(loadedClerks[0]);
    }
    setFolders(loadedFolders);
    setPartyFields(loadedFields);

    return {
      loadedOfficeTemplates,
      loadedDocs,
      loadedEstates,
      loadedClauses,
      loadedRevs,
      loadedSavedParties,
      loadedSavedProps,
      loadedClerks,
      loadedFolders,
      loadedFields,
      activeDraft,
    };
  }, [updateDocumentsState]);

  const connectAndReloadLocalDirectory = useCallback(
    async (showToast: (msg: string) => void) => {
      try {
        const nextState = await reconnectPersistedOfficeDirectory();
        setLocalDirState(nextState);
        if (nextState.status === 'connected') {
          await reloadAllOfficeRecords();
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
    [reloadAllOfficeRecords]
  );

  return {
    templates,
    setTemplates,
    documents,
    setDocuments: updateDocumentsState,
    documentsRef,
    revisions,
    setRevisions,
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
    partyFields,
    setPartyFields,
    lastBackupAt,
    setLastBackupAt,
    localDirState,
    setLocalDirState,
    reloadAllOfficeRecords,
    connectAndReloadLocalDirectory,
  };
}
