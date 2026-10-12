'use client';

import React, { useCallback, useMemo } from 'react';
import {
  ContractOutlineClause,
  ContractPartyCard,
  SavedPartyRecord,
  SavedPropertyRecord,
  SubdivisionEstate,
  SubdivisionLot,
  VariableInputType,
} from '@/lib/types';
import { applyPartyCardsToFieldValues } from '@/lib/clause-outline-engine';
import {
  convertSelectionToSmartTag,
  decorateSmartTagsInDOM,
  escapeHtml,
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
} from '@/lib/editor-utils';
import {
  lookupPlaceholderValue,
  mergePlaceholdersIntoHtml,
  normalizePlaceholderKey,
} from '@/lib/docx-engine';
import {
  buildFullSubdivisionTableHtml,
  buildSubdivisionLotClauseHtml,
} from '@/lib/document-io-workflows';

interface UseEditorPartiesParams {
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  headerEditorRef: React.RefObject<HTMLDivElement | null>;
  footerEditorRef: React.RefObject<HTMLDivElement | null>;
  savedRangeRef: React.MutableRefObject<Range | null>;
  estates: SubdivisionEstate[];
  extractedPlaceholders: string[];
  partyCardsRef: React.MutableRefObject<ContractPartyCard[]>;
  setPartyCards: React.Dispatch<React.SetStateAction<ContractPartyCard[]>>;
  fieldValues: Record<string, string>;
  fieldValuesRef: React.MutableRefObject<Record<string, string>>;
  setFieldValues: React.Dispatch<React.SetStateAction<Record<string, string>>>;
  fieldInputTypesRef: React.MutableRefObject<Record<string, VariableInputType>>;
  setFieldInputTypes: React.Dispatch<
    React.SetStateAction<Record<string, VariableInputType>>
  >;
  setSelectedEstateId: React.Dispatch<React.SetStateAction<string>>;
  setSelectedLotNumber: React.Dispatch<React.SetStateAction<string>>;
  ensureBodyEditable: () => boolean;
  recordHistorySnapshot: (label?: string) => void;
  syncPlaceholdersAndDraft: (
    customFieldValues?: Record<string, string>,
    customFieldInputTypes?: Record<string, VariableInputType>,
    customOutlineClauses?: ContractOutlineClause[],
    customPartyCards?: ContractPartyCard[]
  ) => void;
  saveRolePartyToCatalog: (
    role: 'party1' | 'party2',
    fieldValues: Record<string, string>
  ) => Promise<void>;
  savePropertyToCatalog: (fieldValues: Record<string, string>) => Promise<void>;
  showToast: (msg: string) => void;
}

export function useEditorParties({
  bodyEditorRef,
  headerEditorRef,
  footerEditorRef,
  savedRangeRef,
  estates,
  extractedPlaceholders,
  partyCardsRef,
  setPartyCards,
  fieldValues,
  fieldValuesRef,
  setFieldValues,
  fieldInputTypesRef,
  setFieldInputTypes,
  setSelectedEstateId,
  setSelectedLotNumber,
  ensureBodyEditable,
  recordHistorySnapshot,
  syncPlaceholdersAndDraft,
  saveRolePartyToCatalog,
  savePropertyToCatalog,
  showToast,
}: UseEditorPartiesParams) {
  const handleUpdateFieldValue = useCallback(
    (key: string, value: string) => {
      const next = { ...fieldValuesRef.current, [key]: value };
      fieldValuesRef.current = next;
      setFieldValues(next);
      syncPlaceholdersAndDraft(next);
    },
    [fieldValuesRef, setFieldValues, syncPlaceholdersAndDraft]
  );

  const handleChangeFieldInputType = useCallback(
    (key: string, inputType: 'text' | 'number' | 'date') => {
      const nextTypes: Record<string, VariableInputType> = {
        ...fieldInputTypesRef.current,
        [key]: inputType,
      };
      fieldInputTypesRef.current = nextTypes;
      setFieldInputTypes(nextTypes);
      syncPlaceholdersAndDraft(fieldValuesRef.current, nextTypes);
    },
    [
      fieldInputTypesRef,
      fieldValuesRef,
      setFieldInputTypes,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleUpdatePartyCards = useCallback(
    (nextCards: ContractPartyCard[]) => {
      partyCardsRef.current = nextCards;
      setPartyCards(nextCards);
      const nextValues = applyPartyCardsToFieldValues(
        nextCards,
        fieldValuesRef.current,
        extractedPlaceholders
      );
      fieldValuesRef.current = nextValues;
      setFieldValues(nextValues);
      syncPlaceholdersAndDraft(nextValues, undefined, undefined, nextCards);
    },
    [
      extractedPlaceholders,
      fieldValuesRef,
      partyCardsRef,
      setFieldValues,
      setPartyCards,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleRecallPartyToRole = useCallback(
    (party: SavedPartyRecord, role: 'party1' | 'party2') => {
      const prefix = role === 'party1' ? 'الطرف_الأول' : 'الطرف_الثاني';
      const next: Record<string, string> = {
        ...fieldValuesRef.current,
        [`${prefix}_الاسم`]: party.fullName || '',
        [`${prefix}_تاريخ_الميلاد`]: party.birthDate || '',
        [`${prefix}_مكان_الميلاد`]: party.birthPlace || '',
        [`${prefix}_النسب`]: party.parentage || '',
        [`${prefix}_الهوية`]: party.idCardRef || '',
        [`${prefix}_الإقامة`]: party.address || '',
      };
      fieldValuesRef.current = next;
      setFieldValues(next);
      syncPlaceholdersAndDraft(next);
      showToast(
        `تم استدعاء بيانات "${party.fullName}" في حقول ${
          role === 'party1' ? 'الطرف الأول' : 'الطرف الثاني'
        }`
      );
    },
    [fieldValuesRef, setFieldValues, showToast, syncPlaceholdersAndDraft]
  );

  const handleSaveCurrentPartyToDirectory = useCallback(
    async (role: 'party1' | 'party2') => {
      await saveRolePartyToCatalog(role, fieldValuesRef.current);
    },
    [fieldValuesRef, saveRolePartyToCatalog]
  );

  const handleRecallPropertyRecord = useCallback(
    (prop: SavedPropertyRecord) => {
      const next: Record<string, string> = {
        ...fieldValuesRef.current,
        رقم_الحصة: prop.lotNumber || '',
        طبيعة_الحصة: prop.nature || '',
        الطابق: prop.floor || '',
        المساحة: prop.area || '',
        الأجزاء_المشتركة: prop.commonShares || '',
        تعيين_الحصة_الكامل: prop.fullDescription || '',
        مراجع_الوصف_التقسيمي: prop.deedRef || '',
      };
      fieldValuesRef.current = next;
      setFieldValues(next);
      syncPlaceholdersAndDraft(next);
      showToast(`تم استدعاء بيانات العقار "${prop.label}" في حقول التعيين`);
    },
    [fieldValuesRef, setFieldValues, showToast, syncPlaceholdersAndDraft]
  );

  const handleSaveCurrentPropertyToDirectory = useCallback(async () => {
    await savePropertyToCatalog(fieldValuesRef.current);
  }, [fieldValuesRef, savePropertyToCatalog]);

  const handleSelectEstateAndLot = useCallback(
    (estateId: string, lotNumber: string) => {
      setSelectedEstateId(estateId);
      setSelectedLotNumber(lotNumber);

      const estate = estates.find((e) => e.id === estateId);
      const lot = estate?.lots.find((l) => l.lotNumber === lotNumber);

      if (estate && lot) {
        const nextValues: Record<string, string> = {
          ...fieldValuesRef.current,
          رقم_الحصة: lot.lotNumber,
          طبيعة_الحصة: lot.nature,
          الطابق: `${lot.floor}${lot.building ? ` (${lot.building})` : ''}`,
          المساحة: lot.area,
          الأجزاء_المشتركة: lot.commonShares,
          تعيين_الحصة_الكامل: lot.fullDescription,
          مراجع_الوصف_التقسيمي: estate.subdivisionDeedRef,
        };
        fieldValuesRef.current = nextValues;
        setFieldValues(nextValues);
        syncPlaceholdersAndDraft(nextValues);
        showToast(
          `تم ربط بيانات الحصة رقم (${lot.lotNumber}) بالتعيين العقاري تلقائياً`
        );
      }
    },
    [
      estates,
      fieldValuesRef,
      setFieldValues,
      setSelectedEstateId,
      setSelectedLotNumber,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleInsertLotClauseAtCaret = useCallback(
    (estate: SubdivisionEstate, lot: SubdivisionLot) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      const clauseHtml = buildSubdivisionLotClauseHtml(estate, lot);
      insertHtmlAtSelection(
        bodyEditorRef.current,
        clauseHtml,
        savedRangeRef.current
      );
      syncPlaceholdersAndDraft();
      showToast(`تم إدراج فقرة تعيين الحصة رقم (${lot.lotNumber}) في العقد`);
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      savedRangeRef,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleInsertFullSubdivisionTableAtCaret = useCallback(
    (estate: SubdivisionEstate) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();
      const tableHtml = buildFullSubdivisionTableHtml(estate);
      insertHtmlAtSelection(
        bodyEditorRef.current,
        tableHtml,
        savedRangeRef.current
      );
      syncPlaceholdersAndDraft();
      showToast(`تم إدراج جدول الوصف التقسيمي (${estate.estateName}) في العقد`);
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      savedRangeRef,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleConvertSelectionToSmartTag = useCallback(() => {
    if (!bodyEditorRef.current || !ensureBodyEditable()) return;
    recordHistorySnapshot();
    const createdVar = convertSelectionToSmartTag(
      bodyEditorRef.current,
      savedRangeRef.current
    );
    if (!createdVar) {
      showToast(
        'حدد أي كلمة أو جملة داخل ورقة العقد أولاً ثم اضغط «[ ] تحويل المحدد لوسم»'
      );
      return;
    }
    syncPlaceholdersAndDraft();
    showToast(
      `تم تحويل النص المحدد إلى وسم ذكي [${createdVar}] مرتبط بهذا العقد`
    );
  }, [
    bodyEditorRef,
    ensureBodyEditable,
    recordHistorySnapshot,
    savedRangeRef,
    showToast,
    syncPlaceholdersAndDraft,
  ]);

  const handleInsertSmartTagAtCaret = useCallback(
    (varKey: string) => {
      if (!bodyEditorRef.current || !varKey.trim() || !ensureBodyEditable())
        return;
      const cleanKey = normalizePlaceholderKey(varKey).replace(/\s+/g, '_');
      if (!cleanKey) return;
      recordHistorySnapshot();
      const tagHtml = `<span class="smart-tag" data-var="${escapeHtml(
        cleanKey
      )}">[${escapeHtml(cleanKey)}]</span>&nbsp;`;
      insertHtmlAtSelection(
        bodyEditorRef.current,
        tagHtml,
        savedRangeRef.current
      );
      syncPlaceholdersAndDraft();
      showToast(`تم إدراج الوسم [${cleanKey}] عند موضع المؤشر`);
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      recordHistorySnapshot,
      savedRangeRef,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const handleBakeAllPlaceholdersIntoDocument = useCallback(
    (explicitValues?: Record<string, string>) => {
      if (!bodyEditorRef.current || !ensureBodyEditable()) return;
      recordHistorySnapshot();

      const activeValues: Record<string, string> = {
        ...fieldValues,
        ...fieldValuesRef.current,
        ...(explicitValues || {}),
      };
      fieldValuesRef.current = activeValues;
      setFieldValues(activeValues);

      const replaceLiveZoneTags = (zoneEl: HTMLElement | null) => {
        if (!zoneEl) return;
        decorateSmartTagsInDOM(zoneEl);
        const liveSpans = Array.from(
          zoneEl.querySelectorAll('.smart-tag, .smart-placeholder, [data-var]')
        ) as HTMLElement[];
        for (const span of liveSpans) {
          const rawKey =
            span.getAttribute('data-var') || (span.textContent || '');
          const val = lookupPlaceholderValue(activeValues, rawKey);
          if (val !== undefined && val.trim() !== '') {
            const textNode = document.createTextNode(val.trim());
            span.parentNode?.replaceChild(textNode, span);
          }
        }
        zoneEl.innerHTML = mergePlaceholdersIntoHtml(
          zoneEl.innerHTML,
          activeValues
        );
      };

      replaceLiveZoneTags(bodyEditorRef.current);
      normalizeNotaryContainerDOM(bodyEditorRef.current);

      if (headerEditorRef.current && headerEditorRef.current.innerHTML) {
        replaceLiveZoneTags(headerEditorRef.current);
      }
      if (footerEditorRef.current && footerEditorRef.current.innerHTML) {
        replaceLiveZoneTags(footerEditorRef.current);
      }

      syncPlaceholdersAndDraft(activeValues);
      showToast('تم دمج واستبدال قيم الحقول داخل نص العقد بنجاح');
    },
    [
      bodyEditorRef,
      ensureBodyEditable,
      fieldValues,
      fieldValuesRef,
      footerEditorRef,
      headerEditorRef,
      recordHistorySnapshot,
      setFieldValues,
      showToast,
      syncPlaceholdersAndDraft,
    ]
  );

  const unfilledCount = useMemo(
    () =>
      extractedPlaceholders.filter((k) => {
        const val = lookupPlaceholderValue(fieldValues, k);
        return !val || !val.trim();
      }).length,
    [extractedPlaceholders, fieldValues]
  );

  return {
    unfilledCount,
    handleUpdateFieldValue,
    handleChangeFieldInputType,
    handleUpdatePartyCards,
    handleRecallPartyToRole,
    handleSaveCurrentPartyToDirectory,
    handleRecallPropertyRecord,
    handleSaveCurrentPropertyToDirectory,
    handleSelectEstateAndLot,
    handleInsertLotClauseAtCaret,
    handleInsertFullSubdivisionTableAtCaret,
    handleConvertSelectionToSmartTag,
    handleInsertSmartTagAtCaret,
    handleBakeAllPlaceholdersIntoDocument,
  };
}
