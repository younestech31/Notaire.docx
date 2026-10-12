'use client';

import { useEffect, useState } from 'react';
import {
  findMatchesAcrossNodes,
  focusAndSelectMatchInDOM,
  MultiNodeTextMatch,
  replaceMatchesAcrossNodes,
} from '@/lib/editor-utils';

interface UseFindAndReplaceParams {
  bodyEditorRef: React.RefObject<HTMLDivElement | null>;
  ensureBodyEditable: () => boolean;
  recordHistorySnapshot: (label?: string) => void;
  syncPlaceholdersAndDraft: () => void;
  showToast: (msg: string) => void;
}

export function useFindAndReplace({
  bodyEditorRef,
  ensureBodyEditable,
  recordHistorySnapshot,
  syncPlaceholdersAndDraft,
  showToast,
}: UseFindAndReplaceParams) {
  const [showFindReplace, setShowFindReplace] = useState<boolean>(false);
  const [findQuery, setFindQuery] = useState<string>('');
  const [replaceQuery, setReplaceQuery] = useState<string>('');
  const [ignoreArabicHamzaAndDiacritics, setIgnoreArabicHamzaAndDiacritics] =
    useState<boolean>(true);
  const [activeMatchIdx, setActiveMatchIdx] = useState<number>(0);
  const [matches, setMatches] = useState<MultiNodeTextMatch[]>([]);
  const [matchCount, setMatchCount] = useState<number>(0);

  useEffect(() => {
    if (!showFindReplace || !findQuery.trim() || !bodyEditorRef.current) {
      setMatches([]);
      setMatchCount(0);
      setActiveMatchIdx(0);
      return;
    }
    const foundMatches = findMatchesAcrossNodes(
      bodyEditorRef.current,
      findQuery,
      ignoreArabicHamzaAndDiacritics
    );
    setMatches(foundMatches);
    setMatchCount(foundMatches.length);
    if (foundMatches.length > 0) {
      setActiveMatchIdx(0);
      focusAndSelectMatchInDOM(foundMatches[0]);
    }
  }, [
    bodyEditorRef,
    findQuery,
    ignoreArabicHamzaAndDiacritics,
    showFindReplace,
  ]);

  const handleNextMatch = () => {
    if (matches.length === 0) return;
    const next = (activeMatchIdx + 1) % matches.length;
    setActiveMatchIdx(next);
    focusAndSelectMatchInDOM(matches[next]);
  };

  const handlePrevMatch = () => {
    if (matches.length === 0) return;
    const prev = (activeMatchIdx - 1 + matches.length) % matches.length;
    setActiveMatchIdx(prev);
    focusAndSelectMatchInDOM(matches[prev]);
  };

  const handleRunReplace = (replaceAll: boolean) => {
    if (!bodyEditorRef.current || !findQuery || !ensureBodyEditable()) return;
    recordHistorySnapshot(
      replaceAll ? 'قبل استبدال الكل في البحث' : 'قبل استبدال تطابق نصي'
    );
    const count = replaceMatchesAcrossNodes(
      bodyEditorRef.current,
      findQuery,
      replaceQuery,
      replaceAll,
      activeMatchIdx,
      ignoreArabicHamzaAndDiacritics
    );
    const updated = findMatchesAcrossNodes(
      bodyEditorRef.current,
      findQuery,
      ignoreArabicHamzaAndDiacritics
    );
    setMatches(updated);
    setMatchCount(updated.length);
    if (updated.length > 0) {
      const nextIdx = Math.min(activeMatchIdx, updated.length - 1);
      setActiveMatchIdx(nextIdx);
      focusAndSelectMatchInDOM(updated[nextIdx]);
    }
    syncPlaceholdersAndDraft();
    showToast(`تم استبدال (${count}) تطابق بنجاح`);
  };

  return {
    showFindReplace,
    setShowFindReplace,
    findQuery,
    setFindQuery,
    replaceQuery,
    setReplaceQuery,
    ignoreArabicHamzaAndDiacritics,
    setIgnoreArabicHamzaAndDiacritics,
    activeMatchIdx,
    matches,
    matchCount,
    handleNextMatch,
    handlePrevMatch,
    handleRunReplace,
  };
}
