'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import {
  ClauseVariableGroup,
  ContractOutlineClause,
  ContractPartyCard,
  NotaryClause,
  OfficeTemplate,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  VariableInputType,
} from '@/lib/types';

export const INITIAL_EMPTY_PARAGRAPH = `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><br></p>`;

export interface DocumentSessionSnapshot {
  docId: string;
  docTitle: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  showHeaderFooter: boolean;
  pageNumberingEnabled: boolean;
  fieldValues: Record<string, string>;
  fieldInputTypes: Record<string, VariableInputType>;
  partyCards: ContractPartyCard[];
  outlineClauses: ContractOutlineClause[];
  contractConditions: Record<string, string | boolean>;
  selectedEstateId: string;
  selectedLotNumber: string;
}

export function useDocumentSession() {
  const bodyEditorRef = useRef<HTMLDivElement | null>(null);
  const headerEditorRef = useRef<HTMLDivElement | null>(null);
  const footerEditorRef = useRef<HTMLDivElement | null>(null);

  const [docId, setDocId] = useState<string>('doc_initial');
  const [docTitle, setDocTitle] = useState<string>('عقد توثيقي جديد');
  const [showHeaderFooter, setShowHeaderFooter] = useState<boolean>(false);
  const [pageNumberingEnabled, setPageNumberingEnabled] = useState<boolean>(true);
  const [previewMergedMode, setPreviewMergedMode] = useState<boolean>(false);
  const [mergedPreviewHtml, setMergedPreviewHtml] = useState<string>('');
  const [printPaginatedHtml, setPrintPaginatedHtml] = useState<string>('');
  const [zoom, setZoom] = useState<number>(100);

  // Temporary sub-editing mode for an OfficeTemplate or a NotaryClause inside the A4 Editor
  const [editingDerivedTpl, setEditingDerivedTpl] =
    useState<OfficeTemplate | null>(null);
  const [editingClauseObj, setEditingClauseObj] =
    useState<NotaryClause | null>(null);
  const stashedMainDocRef = useRef<{
    bodyHtml: string;
    headerHtml: string;
    footerHtml: string;
    title: string;
  } | null>(null);

  // Live clauses & conditions in current session
  const [activeClauseIdsInDoc, setActiveClauseIdsInDoc] = useState<string[]>([]);
  const [outlineClauses, setOutlineClauses] = useState<ContractOutlineClause[]>(
    []
  );
  const outlineClausesRef = useRef<ContractOutlineClause[]>([]);
  useEffect(() => {
    outlineClausesRef.current = outlineClauses;
  }, [outlineClauses]);

  const [contractConditions, setContractConditions] = useState<
    Record<string, string | boolean>
  >({});
  const contractConditionsRef = useRef<Record<string, string | boolean>>({});
  useEffect(() => {
    contractConditionsRef.current = contractConditions;
  }, [contractConditions]);

  // Structured party cards & free placeholders
  const [partyCards, setPartyCards] = useState<ContractPartyCard[]>([]);
  const partyCardsRef = useRef<ContractPartyCard[]>([]);
  useEffect(() => {
    partyCardsRef.current = partyCards;
  }, [partyCards]);

  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const fieldValuesRef = useRef<Record<string, string>>({});
  useEffect(() => {
    fieldValuesRef.current = fieldValues;
  }, [fieldValues]);

  const [fieldInputTypes, setFieldInputTypes] = useState<
    Record<string, VariableInputType>
  >({});
  const fieldInputTypesRef = useRef<Record<string, VariableInputType>>({});
  useEffect(() => {
    fieldInputTypesRef.current = fieldInputTypes;
  }, [fieldInputTypes]);

  const [extractedPlaceholders, setExtractedPlaceholders] = useState<string[]>(
    []
  );
  const [clauseGroups, setClauseGroups] = useState<ClauseVariableGroup[]>([]);
  const [selectedEstateId, setSelectedEstateId] = useState<string>('');
  const [selectedLotNumber, setSelectedLotNumber] = useState<string>('');

  // Auto-save & Status Bar metrics
  const [autoSaveState, setAutoSaveState] = useState<
    'saved' | 'saving' | 'error'
  >('saved');
  const [docMetrics, setDocMetrics] = useState<{
    wordCount: number;
    charCount: number;
    estimatedPages: number;
  }>({ wordCount: 0, charCount: 0, estimatedPages: 1 });

  const getEditorZonesHtml = useCallback(() => {
    return {
      bodyHtml: bodyEditorRef.current?.innerHTML || INITIAL_EMPTY_PARAGRAPH,
      headerHtml: headerEditorRef.current?.innerHTML || '',
      footerHtml: footerEditorRef.current?.innerHTML || '',
    };
  }, []);

  const getSessionSnapshot = useCallback((): DocumentSessionSnapshot => {
    const { bodyHtml, headerHtml, footerHtml } = getEditorZonesHtml();
    return {
      docId,
      docTitle,
      bodyHtml,
      headerHtml,
      footerHtml,
      showHeaderFooter,
      pageNumberingEnabled,
      fieldValues: fieldValuesRef.current,
      fieldInputTypes: fieldInputTypesRef.current,
      partyCards: partyCardsRef.current,
      outlineClauses: outlineClausesRef.current,
      contractConditions: contractConditionsRef.current,
      selectedEstateId,
      selectedLotNumber,
    };
  }, [
    docId,
    docTitle,
    getEditorZonesHtml,
    pageNumberingEnabled,
    selectedEstateId,
    selectedLotNumber,
    showHeaderFooter,
  ]);

  return {
    bodyEditorRef,
    headerEditorRef,
    footerEditorRef,
    docId,
    setDocId,
    docTitle,
    setDocTitle,
    showHeaderFooter,
    setShowHeaderFooter,
    pageNumberingEnabled,
    setPageNumberingEnabled,
    previewMergedMode,
    setPreviewMergedMode,
    mergedPreviewHtml,
    setMergedPreviewHtml,
    printPaginatedHtml,
    setPrintPaginatedHtml,
    zoom,
    setZoom,
    editingDerivedTpl,
    setEditingDerivedTpl,
    editingClauseObj,
    setEditingClauseObj,
    stashedMainDocRef,
    activeClauseIdsInDoc,
    setActiveClauseIdsInDoc,
    outlineClauses,
    setOutlineClauses,
    outlineClausesRef,
    contractConditions,
    setContractConditions,
    contractConditionsRef,
    partyCards,
    setPartyCards,
    partyCardsRef,
    fieldValues,
    setFieldValues,
    fieldValuesRef,
    fieldInputTypes,
    setFieldInputTypes,
    fieldInputTypesRef,
    extractedPlaceholders,
    setExtractedPlaceholders,
    clauseGroups,
    setClauseGroups,
    selectedEstateId,
    setSelectedEstateId,
    selectedLotNumber,
    setSelectedLotNumber,
    autoSaveState,
    setAutoSaveState,
    docMetrics,
    setDocMetrics,
    getEditorZonesHtml,
    getSessionSnapshot,
  };
}
