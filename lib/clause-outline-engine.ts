import type {
  ClauseRecognitionKind,
  ContractOutlineClause,
  ContractPartyCard,
} from './types';
import {
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
} from './types';
import { lookupPlaceholderValue } from './docx-engine';
import {
  decorateSmartTagsInDOM,
  escapeHtml,
  extractPlaceholdersFromHtml,
  insertHtmlAtSelection,
  normalizeNotaryContainerDOM,
  syncSmartTagsFilledStateInDOM,
} from './editor-utils';

/**
 * # New Clause Toolbar Action (Flexible Heading Model):
 * Converts the current line/selection (or inserts at caret) into a bold `# عنوان البند`
 * heading paragraph followed by an empty paragraph for free writing — without forcing
 * a `.clause-container` wrapper or mandatory placeholder text.
 */
export function insertOrWrapNewClauseAtSelection(
  editorEl: HTMLElement,
  savedRange: Range | null,
  clauseIndexHint: number = 1,
  customTitle?: string
): { clauseId: string; title: string; contentHtml: string } {
  editorEl.focus();
  const sel = window.getSelection();
  if (
    sel &&
    sel.rangeCount === 0 &&
    savedRange &&
    editorEl.contains(savedRange.commonAncestorContainer)
  ) {
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }

  let rawTitle = customTitle ? customTitle.replace(/^#+\s*/, '').trim() : '';
  if (!rawTitle && sel && sel.rangeCount > 0) {
    const r = sel.getRangeAt(0);
    if (!r.collapsed && editorEl.contains(r.commonAncestorContainer)) {
      rawTitle = r.toString().replace(/^#+\s*/, '').trim();
      r.deleteContents();
    } else if (r.collapsed && editorEl.contains(r.commonAncestorContainer)) {
      let block: Node | null = r.startContainer;
      while (block && block !== editorEl && (block as HTMLElement).tagName !== 'P') {
        block = block.parentNode;
      }
      if (block && (block as HTMLElement).tagName === 'P') {
        const pText = (block.textContent || '').trim();
        if (pText.startsWith('#')) {
          rawTitle = pText.replace(/^#+\s*/, '').trim();
          (block as HTMLElement).innerHTML = '<br>';
        }
      }
    }
  }

  const title = rawTitle || `بند رقم ${clauseIndexHint}`;
  const clauseId = `hash_clause_${Date.now()}`;
  const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;

  const headingAndBlankLineHtml = `<p dir="rtl" style="${pStyle}" data-clause-id="${clauseId}"><span style="font-weight:bold;"># ${escapeHtml(
    title
  )}</span></p><p dir="rtl" style="${pStyle}"><br></p>`;

  insertHtmlAtSelection(editorEl, headingAndBlankLineHtml, savedRange);
  return { clauseId, title, contentHtml: headingAndBlankLineHtml };
}

/**
 * Notarial standard section keywords for Algerian/Arab contract drafting:
 */
export const NOTARIAL_HEADING_KEYWORDS = [
  'التمهيد',
  'تمهيد',
  'مقدمة العقد',
  'أصل الملكية',
  'سند الملكية',
  'بيان أصل الملكية',
  'تعيين العقار',
  'بيان العقار',
  'وصف المبيع',
  'موضوع العقد',
  'موضوع البيع',
  'موضوع الاتفاق',
  'موضوع الوكالة',
  'شروط البيع',
  'شروط الاتفاق',
  'شروط الإيجار',
  'الشروط والأحكام',
  'شروط العقد',
  'الثمن وطريقة السداد',
  'الثمن والوفاء',
  'الأجرة وطريقة الدفع',
  'بدل الإيجار',
  'الثمن',
  'التزامات البائع',
  'التزامات المشتري',
  'التزامات الطرفين',
  'التزامات المستأجر',
  'التزامات المؤجر',
  'الشرط الجزائي',
  'شرط الفسخ',
  'الفسخ والشرط الجزائي',
  'فسخ العقد',
  'النزاعات والاختصاص القضائي',
  'الاختصاص القضائي',
  'حل النزاعات',
  'النزاعات القضائية',
  'المصادقة والتوقيع',
  'إقرار وتعهد',
  'إقرارات وتعهدات',
  'شهادة الشهود',
  'المصاريف والأتعاب',
  'خاتمة العقد',
  'حالة الأماكن',
  'التصريحات الجبائية',
  'التأمينات والضمانات',
  'الوفاة وانتقال الحقوق',
] as const;

export interface ClauseHeadingMatch {
  isHeading: boolean;
  kind: ClauseRecognitionKind;
  cleanTitle: string;
}

/**
 * Intelligent hybrid recognizer for notarial clauses in live contract DOM:
 * Recognizes:
 * 1. Explicit containers (.clause-container or [data-clause-id])
 * 2. Rapid hash syntax: `# عنوان البند` or `## البند`
 * 3. Law articles: `المادة الأولى: ...`, `المادة 1: ...`, `المادة العاشرة`
 * 4. Contract clauses: `البند الأول: ...`, `البند 2: ...`
 * 5. Verbal ordinal enumeration: `أولاً: ...`, `ثانياً: ...`, `عاشراً: ...`
 * 6. Standard notarial headings: `التمهيد:`, `أصل الملكية:`, `شروط البيع:`
 * 7. Heading tags (H1-H4) or standalone bold title paragraphs
 */
export function detectNotaryClauseHeading(el: HTMLElement): ClauseHeadingMatch {
  if (!el || el.nodeType !== 1) {
    return { isHeading: false, kind: 'clause', cleanTitle: '' };
  }

  const isContainer = el.classList.contains('clause-container');
  const plainText = (el.textContent || '')
    .replace(/[\u200B-\u200D\uFEFF\u200E\u200F\u061C]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  // Rule 1: div.clause-container is a closed container
  if (isContainer) {
    const rawTitle =
      el.getAttribute('data-clause-title') ||
      el.querySelector('strong, b, h1, h2, h3, h4')?.textContent?.trim() ||
      'بند';
    const cleanTitle = rawTitle.replace(/^#+\s*/, '').replace(/:$/, '').trim();
    const kindAttr = el.getAttribute('data-clause-kind') as ClauseRecognitionKind | null;
    return {
      isHeading: true,
      kind: kindAttr || 'container',
      cleanTitle,
    };
  }

  // Rule 2: Paragraph starting with # (with or without data-clause-id)
  if (plainText.startsWith('#')) {
    const rawTitle = el.getAttribute('data-clause-title') || plainText;
    const cleanTitle = rawTitle.replace(/^#+\s*/, '').replace(/:$/, '').trim();
    const kindAttr = (el.getAttribute('data-clause-kind') as ClauseRecognitionKind) || 'hash';
    return {
      isHeading: true,
      kind: kindAttr,
      cleanTitle: cleanTitle || 'بند غير معنون',
    };
  }

  // Rule 3: Any other element with data-clause-id that does NOT start with # is NOT a clause/heading.
  // Ignore stamp and treat as normal text.
  return { isHeading: false, kind: 'clause', cleanTitle: '' };
}

interface DomClauseSegment {
  id: string;
  title: string;
  domIndex: number;
  isContainer: boolean;
  locked: boolean;
  explicitOrder?: number;
  clauseKind?: ClauseRecognitionKind;
  elements: HTMLElement[];
}

function collectDomClauseSegments(rootEl: HTMLElement): DomClauseSegment[] {
  const children = Array.from(rootEl.children) as HTMLElement[];
  const segments: DomClauseSegment[] = [];

  let currentSegment: DomClauseSegment | null = null;

  const flushSegment = () => {
    if (currentSegment) {
      segments.push(currentSegment);
      currentSegment = null;
    }
  };

  children.forEach((child, idx) => {
    if (child.getAttribute('data-notary-pagination') === 'true') {
      return;
    }
    if (child.classList.contains('page-break')) {
      if (currentSegment) {
        currentSegment.elements.push(child);
      }
      return;
    }

    if (child.classList.contains('clause-container')) {
      flushSegment();
      const rawTitle =
        child.getAttribute('data-clause-title') ||
        child.querySelector('strong, b, h1, h2, h3, h4')?.textContent?.trim() ||
        `بند رقم ${segments.length + 1}`;
      const cleanTitle = rawTitle.replace(/^#+\s*/, '').replace(/:$/, '').trim();
      const cId = child.getAttribute('data-clause-id') || `dom_container_${idx}`;
      const locked =
        child.getAttribute('data-clause-locked') === 'true' ||
        child.getAttribute('contenteditable') === 'false';
      const orderAttr = child.getAttribute('data-clause-order');
      const explicitOrder =
        orderAttr !== null && orderAttr !== '' && !Number.isNaN(Number(orderAttr))
          ? Number(orderAttr)
          : undefined;
      const detectedKind =
        (child.getAttribute('data-clause-kind') as ClauseRecognitionKind) || 'container';

      segments.push({
        id: cId,
        title: cleanTitle || `بند رقم ${segments.length + 1}`,
        domIndex: idx,
        isContainer: true,
        locked,
        explicitOrder,
        clauseKind: detectedKind,
        elements: [child],
      });
      currentSegment = null;
      return;
    }

    const headingInfo = detectNotaryClauseHeading(child);
    if (headingInfo.isHeading) {
      flushSegment();
      const cId = child.getAttribute('data-clause-id') || `hash_clause_${idx}`;

      currentSegment = {
        id: cId,
        title: headingInfo.cleanTitle,
        domIndex: idx,
        isContainer: false,
        locked: false,
        clauseKind: headingInfo.kind,
        elements: [child],
      };
    } else if (currentSegment) {
      currentSegment.elements.push(child);
    }
  });

  flushSegment();
  return segments;
}

/**
 * Lazily upgrades top-level clause segments into `<div class="clause-container">`
 * ONLY when required by structural operations (lock, disable/hide, conditions, reorder).
 * Preserves the segment ID (`seg.id` or `matchedPrev.id`) so the calling operation can
 * immediately target the newly created container.
 */
export function upgradeTopLevelHashSegmentsToContainers(
  rootEl: HTMLElement | null,
  existingOutline: ContractOutlineClause[] = []
): void {
  if (!rootEl) return;
  const segments = collectDomClauseSegments(rootEl);
  if (segments.length === 0) return;

  segments.forEach((seg, idx) => {
    const matchedPrev = existingOutline.find(
      (c) => c.id === seg.id || c.domIndex === seg.domIndex || c.title === seg.title
    );
    const clauseOrder =
      matchedPrev?.order !== undefined
        ? matchedPrev.order
        : seg.explicitOrder !== undefined
        ? seg.explicitOrder
        : idx;
    const isLocked = matchedPrev?.locked ?? seg.locked ?? false;

    if (seg.isContainer) {
      const containerEl = seg.elements[0];
      if (!containerEl.getAttribute('data-clause-id')) {
        containerEl.setAttribute('data-clause-id', matchedPrev?.id || seg.id);
      }
      containerEl.setAttribute('data-clause-order', String(clauseOrder));
      containerEl.setAttribute('data-clause-title', seg.title);
      if (seg.clauseKind) {
        containerEl.setAttribute('data-clause-kind', seg.clauseKind);
      }
      containerEl.setAttribute('data-clause-locked', isLocked ? 'true' : 'false');
      if (isLocked) {
        containerEl.setAttribute('contenteditable', 'false');
        containerEl.style.backgroundColor = 'rgba(248, 250, 252, 0.75)';
        containerEl.style.borderRight = '3px solid #f59e0b';
        containerEl.style.paddingRight = '6px';
      }
    } else {
      const firstEl = seg.elements[0];
      if (!firstEl || !firstEl.parentNode) return;
      const wrapper = document.createElement('div');
      wrapper.className = 'clause-container';
      const newId = matchedPrev?.id || seg.id || `clause_${Date.now()}_${idx}`;
      wrapper.setAttribute('data-clause-id', newId);
      wrapper.setAttribute('data-clause-order', String(clauseOrder));
      wrapper.setAttribute('data-clause-title', seg.title);
      if (seg.clauseKind) {
        wrapper.setAttribute('data-clause-kind', seg.clauseKind);
      }
      wrapper.setAttribute('data-clause-locked', isLocked ? 'true' : 'false');
      if (isLocked) {
        wrapper.setAttribute('contenteditable', 'false');
        wrapper.style.backgroundColor = 'rgba(248, 250, 252, 0.75)';
        wrapper.style.borderRight = '3px solid #f59e0b';
        wrapper.style.paddingRight = '6px';
      }
      firstEl.parentNode.insertBefore(wrapper, firstEl);
      for (const el of seg.elements) {
        wrapper.appendChild(el);
      }
    }
  });
}

function buildClausePreviewAndHtml(seg: DomClauseSegment): {
  innerContentHtml: string;
  previewText: string;
  variables: string[];
} {
  let innerContentHtml = '';
  let previewSource = '';

  if (seg.isContainer) {
    const containerEl = seg.elements[0];
    innerContentHtml = containerEl ? containerEl.innerHTML : '';
    if (containerEl) {
      const children = Array.from(containerEl.children);
      if (children.length > 1) {
        previewSource = children
          .slice(1)
          .map((c) => c.textContent || '')
          .join(' ');
      } else {
        // Container has only the heading line (or single element): check if there is text after the first line
        const rawText = (containerEl.textContent || '').trim();
        const withoutHeading = rawText
          .replace(/^#+\s*[^\n\r]*/, '')
          .trim();
        previewSource =
          withoutHeading && withoutHeading !== seg.title ? withoutHeading : '';
      }
    }
  } else {
    const tempDiv = document.createElement('div');
    seg.elements.forEach((el) => tempDiv.appendChild(el.cloneNode(true)));
    innerContentHtml = tempDiv.innerHTML;
    if (seg.elements.length > 1) {
      previewSource = seg.elements
        .slice(1)
        .map((e) => e.textContent || '')
        .join(' ');
    } else {
      previewSource = '';
    }
  }

  const variables = extractPlaceholdersFromHtml(innerContentHtml);
  const previewText = previewSource.replace(/\s+/g, ' ').trim().slice(0, 110);
  return { innerContentHtml, previewText, variables };
}

/**
 * Extracts live contract clauses from the open A4 editor DOM and merges them
 * with any disabled (`enabled: false`) clauses in `prevOutline` (Single Source of Truth).
 */
export function extractLiveContractClausesFromDOM(
  rootEl: HTMLElement | null,
  prevOutline: ContractOutlineClause[] = []
): ContractOutlineClause[] {
  if (!rootEl) return prevOutline;
  const segments = collectDomClauseSegments(rootEl);

  const disabledClauses = prevOutline.filter((c) => c.enabled === false);

  const activeClauses: ContractOutlineClause[] = segments.map((seg, i) => {
    const prevMatch = prevOutline.find(
      (c) => c.id === seg.id || (c.enabled !== false && c.title === seg.title && c.domIndex === seg.domIndex)
    );
    const { innerContentHtml, previewText, variables } = buildClausePreviewAndHtml(seg);
    const order =
      seg.explicitOrder !== undefined
        ? seg.explicitOrder
        : prevMatch?.order !== undefined
        ? prevMatch.order
        : i;

    return {
      id: seg.id,
      index: i + 1,
      order,
      title: seg.title,
      previewText,
      variables,
      contentHtml: innerContentHtml,
      enabled: true,
      locked: prevMatch?.locked ?? seg.locked ?? false,
      domIndex: seg.domIndex,
      isContainer: seg.isContainer,
      condition: prevMatch?.condition,
      clauseKind: seg.clauseKind ?? prevMatch?.clauseKind ?? 'clause',
    };
  });

  if (disabledClauses.length === 0) {
    return activeClauses
      .sort((a, b) => a.order - b.order)
      .map((c, idx) => ({ ...c, index: idx + 1, order: idx }));
  }

  const merged = [...activeClauses, ...disabledClauses].sort((a, b) => {
    if (a.order !== b.order) return a.order - b.order;
    return a.index - b.index;
  });

  return merged.map((c, idx) => ({
    ...c,
    index: idx + 1,
    order: idx,
  }));
}

export function toggleContractClauseEnabledInDOM(
  rootEl: HTMLElement | null,
  currentOutline: ContractOutlineClause[],
  clauseId: string,
  fieldValues: Record<string, string> = {}
): ContractOutlineClause[] {
  if (!rootEl) return currentOutline;

  upgradeTopLevelHashSegmentsToContainers(rootEl, currentOutline);
  const syncedOutline = extractLiveContractClausesFromDOM(rootEl, currentOutline);
  const targetIdx = syncedOutline.findIndex((c) => c.id === clauseId);
  if (targetIdx < 0) return syncedOutline;

  const targetClause = syncedOutline[targetIdx];
  const nextEnabled = !targetClause.enabled;

  if (!nextEnabled) {
    const domEl = rootEl.querySelector(
      `[data-clause-id="${CSS.escape(clauseId)}"]`
    ) as HTMLElement | null;
    let latestHtml = targetClause.contentHtml;
    if (domEl) {
      latestHtml = domEl.innerHTML;
      domEl.remove();
    }
    const vars = extractPlaceholdersFromHtml(latestHtml);
    const updated = syncedOutline.map((c, i) =>
      i === targetIdx
        ? {
            ...c,
            enabled: false,
            contentHtml: latestHtml,
            variables: vars,
            domIndex: -1,
          }
        : c
    );
    normalizeNotaryContainerDOM(rootEl, false);
    return updated;
  } else {
    const newDiv = document.createElement('div');
    newDiv.className = 'clause-container';
    newDiv.setAttribute('data-clause-id', targetClause.id);
    newDiv.setAttribute('data-clause-order', String(targetClause.order));
    newDiv.setAttribute('data-clause-title', targetClause.title);
    newDiv.setAttribute('data-clause-locked', targetClause.locked ? 'true' : 'false');
    if (targetClause.locked) {
      newDiv.setAttribute('contenteditable', 'false');
      newDiv.style.backgroundColor = 'rgba(248, 250, 252, 0.75)';
      newDiv.style.borderRight = '3px solid #f59e0b';
      newDiv.style.paddingRight = '6px';
    }
    newDiv.innerHTML = targetClause.contentHtml;

    const activeContainers = Array.from(
      rootEl.querySelectorAll('.clause-container[data-clause-id]')
    ) as HTMLElement[];

    let nextActiveEl: HTMLElement | null = null;
    for (const el of activeContainers) {
      const elId = el.getAttribute('data-clause-id') || '';
      const outlineItem = syncedOutline.find((c) => c.id === elId);
      const attrOrder = el.getAttribute('data-clause-order');
      const elOrder =
        outlineItem !== undefined
          ? outlineItem.order
          : attrOrder !== null && !Number.isNaN(Number(attrOrder))
          ? Number(attrOrder)
          : Infinity;
      if (elOrder > targetClause.order) {
        nextActiveEl = el;
        break;
      }
    }

    if (nextActiveEl && nextActiveEl.parentNode === rootEl) {
      rootEl.insertBefore(newDiv, nextActiveEl);
    } else {
      const lastChild = rootEl.lastElementChild as HTMLElement | null;
      if (
        lastChild &&
        lastChild.tagName === 'P' &&
        (lastChild.textContent || '').trim() === '' &&
        !lastChild.classList.contains('clause-container')
      ) {
        rootEl.insertBefore(newDiv, lastChild);
      } else {
        rootEl.appendChild(newDiv);
      }
    }

    decorateSmartTagsInDOM(newDiv);
    syncSmartTagsFilledStateInDOM(newDiv, fieldValues);
    normalizeNotaryContainerDOM(rootEl, false);

    const updatedOutline = syncedOutline.map((c, i) =>
      i === targetIdx ? { ...c, enabled: true, isContainer: true } : c
    );
    return extractLiveContractClausesFromDOM(rootEl, updatedOutline);
  }
}

export function toggleContractClauseLockedInDOM(
  rootEl: HTMLElement | null,
  currentOutline: ContractOutlineClause[],
  clauseId: string
): ContractOutlineClause[] {
  if (!rootEl) return currentOutline;
  upgradeTopLevelHashSegmentsToContainers(rootEl, currentOutline);
  const synced = extractLiveContractClausesFromDOM(rootEl, currentOutline);
  const idx = synced.findIndex((c) => c.id === clauseId);
  if (idx < 0) return synced;

  const nextLocked = !synced[idx].locked;
  const domEl = rootEl.querySelector(
    `[data-clause-id="${CSS.escape(clauseId)}"]`
  ) as HTMLElement | null;

  if (domEl) {
    domEl.setAttribute('data-clause-locked', nextLocked ? 'true' : 'false');
    if (nextLocked) {
      domEl.setAttribute('contenteditable', 'false');
      domEl.style.backgroundColor = 'rgba(248, 250, 252, 0.75)';
      domEl.style.borderRight = '3px solid #f59e0b';
      domEl.style.paddingRight = '6px';
    } else {
      domEl.removeAttribute('contenteditable');
      domEl.style.backgroundColor = '';
      domEl.style.borderRight = '';
      domEl.style.paddingRight = '';
    }
  }

  return synced.map((c, i) => (i === idx ? { ...c, locked: nextLocked } : c));
}

export function reorderContractClausesInDOM(
  rootEl: HTMLElement | null,
  currentOutline: ContractOutlineClause[],
  sourceClauseId: string,
  targetClauseId: string
): ContractOutlineClause[] {
  if (!rootEl || sourceClauseId === targetClauseId) return currentOutline;
  upgradeTopLevelHashSegmentsToContainers(rootEl, currentOutline);
  const synced = extractLiveContractClausesFromDOM(rootEl, currentOutline);

  const fromIdx = synced.findIndex((c) => c.id === sourceClauseId);
  const toIdx = synced.findIndex((c) => c.id === targetClauseId);
  if (fromIdx < 0 || toIdx < 0) return synced;

  const reordered = [...synced];
  const [moved] = reordered.splice(fromIdx, 1);
  reordered.splice(toIdx, 0, moved);

  const normalized = reordered.map((c, idx) => ({
    ...c,
    index: idx + 1,
    order: idx,
  }));

  const activeDomElements: HTMLElement[] = [];
  for (const item of normalized) {
    if (!item.enabled) continue;
    const el = rootEl.querySelector(
      `[data-clause-id="${CSS.escape(item.id)}"]`
    ) as HTMLElement | null;
    if (el) {
      el.setAttribute('data-clause-order', String(item.order));
      activeDomElements.push(el);
    }
  }

  if (activeDomElements.length > 0) {
    const firstContainerInDom = rootEl.querySelector(
      '.clause-container[data-clause-id]'
    ) as HTMLElement | null;
    const anchorParent = firstContainerInDom?.parentNode || rootEl;
    const anchorSibling = firstContainerInDom;

    const frag = document.createDocumentFragment();
    for (const el of activeDomElements) {
      frag.appendChild(el);
    }
    if (anchorSibling && anchorSibling.parentNode === anchorParent) {
      anchorParent.insertBefore(frag, anchorSibling);
    } else {
      rootEl.appendChild(frag);
    }
  }

  normalizeNotaryContainerDOM(rootEl, false);
  return extractLiveContractClausesFromDOM(rootEl, normalized);
}

export function scrollToContractClauseInDOM(
  rootEl: HTMLElement | null,
  domIndexOrClauseId: number | string
): void {
  if (!rootEl) return;
  let target: HTMLElement | null = null;
  if (typeof domIndexOrClauseId === 'string') {
    target = rootEl.querySelector(
      `[data-clause-id="${CSS.escape(domIndexOrClauseId)}"]`
    ) as HTMLElement | null;
    if (!target) {
      const segments = collectDomClauseSegments(rootEl);
      const matchedSeg = segments.find((s) => s.id === domIndexOrClauseId);
      if (matchedSeg && matchedSeg.elements.length > 0) {
        target = matchedSeg.elements[0];
      }
    }
  }
  if (!target && typeof domIndexOrClauseId === 'number' && domIndexOrClauseId >= 0) {
    target = (rootEl.children[domIndexOrClauseId] as HTMLElement) || null;
  }
  if (!target) return;

  target.scrollIntoView({ behavior: 'smooth', block: 'center' });
  const prevTransition = target.style.transition;
  const prevBg = target.style.backgroundColor;
  target.style.transition = 'background-color 0.25s ease';
  target.style.backgroundColor = '#dbeafe';
  setTimeout(() => {
    target!.style.backgroundColor = prevBg;
    setTimeout(() => {
      target!.style.transition = prevTransition;
    }, 250);
  }, 900);
}

export function moveContractClauseInDOM(
  rootEl: HTMLElement | null,
  domIndex: number,
  direction: 'up' | 'down',
  currentOutline: ContractOutlineClause[] = []
): boolean {
  if (!rootEl) return false;
  upgradeTopLevelHashSegmentsToContainers(rootEl, currentOutline);
  const synced = extractLiveContractClausesFromDOM(rootEl, currentOutline);
  const segIdx = synced.findIndex((s) => s.domIndex === domIndex);
  if (segIdx < 0) return false;

  const swapIdx = direction === 'up' ? segIdx - 1 : segIdx + 1;
  if (swapIdx < 0 || swapIdx >= synced.length) return false;

  reorderContractClausesInDOM(rootEl, synced, synced[segIdx].id, synced[swapIdx].id);
  return true;
}

export function deleteContractClauseFromDOM(
  rootEl: HTMLElement | null,
  domIndexOrId: number | string
): boolean {
  if (!rootEl) return false;
  if (typeof domIndexOrId === 'string') {
    const el = rootEl.querySelector(
      `[data-clause-id="${CSS.escape(domIndexOrId)}"]`
    ) as HTMLElement | null;
    if (el) {
      el.remove();
      normalizeNotaryContainerDOM(rootEl, false);
      return true;
    }
  }
  const segments = collectDomClauseSegments(rootEl);
  const seg = segments.find(
    (s) => s.domIndex === domIndexOrId || s.id === domIndexOrId
  );
  if (!seg) return false;

  for (const el of seg.elements) {
    el.remove();
  }
  normalizeNotaryContainerDOM(rootEl, false);
  return true;
}

export function renameContractClauseInDOM(
  rootEl: HTMLElement | null,
  domIndexOrId: number | string,
  newTitle: string
): boolean {
  if (!rootEl || !newTitle.trim()) return false;
  const cleanTitle = newTitle.replace(/^#+\s*/, '').trim();
  if (!cleanTitle) return false;

  const segments = collectDomClauseSegments(rootEl);
  const seg = segments.find(
    (s) => s.domIndex === domIndexOrId || s.id === domIndexOrId
  );
  if (!seg) return false;

  const firstEl = seg.elements[0];
  if (!firstEl) return false;

  if (seg.isContainer) {
    firstEl.setAttribute('data-clause-title', cleanTitle);
    const strongEl = firstEl.querySelector('strong, b, span[style*="bold"]');
    if (strongEl) {
      strongEl.textContent = `# ${cleanTitle}`;
    }
  } else {
    firstEl.innerHTML = `<span style="font-weight:bold;"># ${escapeHtml(cleanTitle)}</span>`;
  }

  normalizeNotaryContainerDOM(rootEl, false);
  return true;
}

export interface CandidateImportedClause {
  id: string;
  title: string;
  paragraphsHtml: string[];
  previewText: string;
}

export function segmentHtmlIntoCandidateClauses(html: string): CandidateImportedClause[] {
  if (!html || typeof window === 'undefined' || typeof DOMParser === 'undefined') {
    return [];
  }
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${html}</div>`, 'text/html');
  const root = doc.body.firstElementChild as HTMLElement | null;
  if (!root) return [];

  const children = Array.from(root.children) as HTMLElement[];
  if (children.length === 0) return [];

  const notaryHeadingPatterns =
    /^(?:#+\s*|أولاً|أولا|ثانياً|ثانيا|ثالثاً|ثالثا|رابعاً|رابعا|خامساً|خامسا|سادساً|سادسا|سابعاً|سابعا|البند\s+|المادة\s+|تعيين\s+|أصل\s+الملكية|الملكية\s+والانتفاع|الثمن\s*|الشروط\s+|التصاريح\s+|الحالة\s+المدنية|تعيين\s+الأطراف|موضوع\s+العقد|موضوع\s+الوكالة|الإيجاب\s+والقبول)/i;

  const candidates: CandidateImportedClause[] = [];
  let current: CandidateImportedClause = {
    id: `imp_clause_0`,
    title: 'ديباجة العقد ومقدمة الأطراف',
    paragraphsHtml: [],
    previewText: '',
  };

  const flushCurrent = () => {
    if (current.paragraphsHtml.length > 0) {
      const tmp = document.createElement('div');
      tmp.innerHTML = current.paragraphsHtml.join(' ');
      current.previewText = (tmp.textContent || '')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 120);
      candidates.push(current);
    }
  };

  children.forEach((child, idx) => {
    if (child.classList.contains('clause-container')) {
      flushCurrent();
      const cTitle =
        child.getAttribute('data-clause-title') ||
        child.querySelector('strong, b')?.textContent?.trim() ||
        `بند رقم ${candidates.length + 1}`;
      const innerEls = Array.from(child.children).map((e) => e.outerHTML);
      current = {
        id: `imp_clause_${idx + 1}`,
        title: cTitle.replace(/^#+\s*/, '').replace(/:$/, '').trim(),
        paragraphsHtml: innerEls.length > 0 ? innerEls : [child.innerHTML],
        previewText: '',
      };
      flushCurrent();
      current = {
        id: `imp_clause_after_${idx + 1}`,
        title: `بند رقم ${candidates.length + 1}`,
        paragraphsHtml: [],
        previewText: '',
      };
      return;
    }

    const plain = (child.textContent || '').replace(/\s+/g, ' ').trim();
    const isHashHeading = plain.startsWith('#');

    if (isHashHeading) {
      flushCurrent();
      const cleanTitle = plain.replace(/^#+\s*/, '').replace(/:$/, '').trim();
      current = {
        id: `imp_clause_${idx + 1}`,
        title: cleanTitle || `بند رقم ${candidates.length + 1}`,
        paragraphsHtml: [child.outerHTML],
        previewText: '',
      };
    } else {
      current.paragraphsHtml.push(child.outerHTML);
    }
  });

  flushCurrent();
  return candidates;
}

export function compileCandidateClausesToHtml(
  clauses: CandidateImportedClause[]
): string {
  const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;

  return clauses
    .filter((c) => c.paragraphsHtml.length > 0)
    .map((c, idx) => {
      const cleanTitle = c.title.replace(/^#+\s*/, '').trim() || `بند رقم ${idx + 1}`;
      const firstParaPlain = c.paragraphsHtml[0]
        ? c.paragraphsHtml[0].replace(/<[^>]+>/g, '').trim()
        : '';
      const hasHashPrefix = firstParaPlain.startsWith('#');
      const matchesTitleWithoutHash =
        !hasHashPrefix &&
        firstParaPlain.replace(/:$/, '').trim() === cleanTitle;

      const headingHtml = `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># ${escapeHtml(
        cleanTitle
      )}</span></p>`;

      if (hasHashPrefix) {
        return c.paragraphsHtml.join('');
      }
      if (matchesTitleWithoutHash) {
        return `${headingHtml}${c.paragraphsHtml.slice(1).join('')}`;
      }
      return `${headingHtml}${c.paragraphsHtml.join('')}`;
    })
    .join('');
}

export function buildPartyCardsVariableMap(
  partyCards: ContractPartyCard[]
): Record<string, string> {
  const out: Record<string, string> = {};
  if (!Array.isArray(partyCards) || partyCards.length === 0) return out;

  const withDefiniteArticle = (r: string): string => {
    const clean = r.trim().replace(/\s+/g, '_');
    if (!clean) return 'الطرف';
    return clean.startsWith('ال') ? clean : `ال${clean}`;
  };

  const withoutDefiniteArticle = (r: string): string => {
    const clean = r.trim().replace(/\s+/g, '_');
    return clean.startsWith('ال') ? clean.slice(2) : clean;
  };

  const isFirstPartyRole = (role: string) =>
    ['بائع', 'موكل', 'واهب'].includes(role);
  const isSecondPartyRole = (role: string) =>
    ['مشتري', 'وكيل', 'موهوب'].includes(role);

  for (const card of partyCards) {
    const rawRole =
      card.role === 'أخرى' ? card.customRoleLabel?.trim() || 'طرف' : card.role;
    const defRole = withDefiniteArticle(rawRole);
    const bareRole = withoutDefiniteArticle(rawRole);
    const idx = Math.max(1, card.index || 1);

    const fieldPairs: [string[], string][] = [
      [['الاسم', 'الاسم_واللقب', 'اللقب_والاسم'], card.fullName],
      [['تاريخ_الميلاد', 'الميلاد'], card.birthDate],
      [['مكان_الميلاد'], card.birthPlace],
      [['النسب', 'ابن'], card.filiation],
      [['الرقم_الوطني', 'NIN', 'رقم_التعريف_الوطني'], card.nationalIdNin],
      [['الهوية', 'بطاقة_الهوية', 'وثيقة_الهوية'], card.idCardDetails],
      [['العنوان', 'الإقامة', 'السكن'], card.address],
      [['الممثل_القانوني', 'نائب'], card.legalRepresentative || ''],
    ];

    for (const [suffixes, val] of fieldPairs) {
      if (!val || !val.trim()) continue;
      const trimmedVal = val.trim();
      for (const suffix of suffixes) {
        out[`${defRole}_${idx}_${suffix}`] = trimmedVal;
        out[`${bareRole}_${idx}_${suffix}`] = trimmedVal;
        out[`${defRole}.${idx}.${suffix}`] = trimmedVal;

        if (idx === 1) {
          out[`${defRole}_${suffix}`] = trimmedVal;
          out[`${bareRole}_${suffix}`] = trimmedVal;
          out[`${defRole}.${suffix}`] = trimmedVal;
          out[`${bareRole}.${suffix}`] = trimmedVal;
        }
      }
    }

    if (idx === 1 && isFirstPartyRole(card.role)) {
      if (card.fullName?.trim()) out['الطرف_الأول_الاسم'] = card.fullName.trim();
      if (card.birthDate?.trim()) out['الطرف_الأول_تاريخ_الميلاد'] = card.birthDate.trim();
      if (card.birthPlace?.trim()) out['الطرف_الأول_مكان_الميلاد'] = card.birthPlace.trim();
      if (card.filiation?.trim()) out['الطرف_الأول_النسب'] = card.filiation.trim();
      if (card.idCardDetails?.trim()) out['الطرف_الأول_الهوية'] = card.idCardDetails.trim();
      if (card.nationalIdNin?.trim()) out['الطرف_الأول_الرقم_الوطني'] = card.nationalIdNin.trim();
      if (card.address?.trim()) out['الطرف_الأول_الإقامة'] = card.address.trim();
    } else if (idx === 1 && isSecondPartyRole(card.role)) {
      if (card.fullName?.trim()) out['الطرف_الثاني_الاسم'] = card.fullName.trim();
      if (card.birthDate?.trim()) out['الطرف_الثاني_تاريخ_الميلاد'] = card.birthDate.trim();
      if (card.birthPlace?.trim()) out['الطرف_الثاني_مكان_الميلاد'] = card.birthPlace.trim();
      if (card.filiation?.trim()) out['الطرف_الثاني_النسب'] = card.filiation.trim();
      if (card.idCardDetails?.trim()) out['الطرف_الثاني_الهوية'] = card.idCardDetails.trim();
      if (card.nationalIdNin?.trim()) out['الطرف_الثاني_الرقم_الوطني'] = card.nationalIdNin.trim();
      if (card.address?.trim()) out['الطرف_الثاني_الإقامة'] = card.address.trim();
    }
  }

  return out;
}

export function applyPartyCardsToFieldValues(
  partyCards: ContractPartyCard[],
  currentFieldValues: Record<string, string>,
  extractedPlaceholders: string[] = []
): Record<string, string> {
  const next: Record<string, string> = { ...currentFieldValues };
  const mapped = buildPartyCardsVariableMap(partyCards);
  for (const [k, v] of Object.entries(mapped)) {
    next[k] = v;
  }
  for (const ph of extractedPlaceholders) {
    const val = lookupPlaceholderValue(mapped, ph);
    if (val !== undefined && val.trim() !== '') {
      next[ph] = val;
    }
  }
  return next;
}

export function setContractClauseLockedInDOM(
  rootEl: HTMLElement | null,
  currentOutline: ContractOutlineClause[],
  clauseId: string,
  locked: boolean
): ContractOutlineClause[] {
  if (!rootEl) return currentOutline;
  upgradeTopLevelHashSegmentsToContainers(rootEl, currentOutline);
  const synced = extractLiveContractClausesFromDOM(rootEl, currentOutline);
  const idx = synced.findIndex((c) => c.id === clauseId);
  if (idx < 0) return synced;

  const domEl = rootEl.querySelector(
    `[data-clause-id="${CSS.escape(clauseId)}"]`
  ) as HTMLElement | null;

  if (domEl) {
    domEl.setAttribute('data-clause-locked', locked ? 'true' : 'false');
    domEl.setAttribute('data-locked', locked ? 'true' : 'false');
    if (locked) {
      domEl.setAttribute('contenteditable', 'false');
      domEl.style.backgroundColor = 'rgba(248, 250, 252, 0.75)';
      domEl.style.borderRight = '3px solid #f59e0b';
      domEl.style.paddingRight = '6px';
    } else {
      domEl.removeAttribute('contenteditable');
      domEl.style.backgroundColor = '';
      domEl.style.borderRight = '';
      domEl.style.paddingRight = '';
    }
  }

  return synced.map((c, i) => (i === idx ? { ...c, locked } : c));
}

export const reorderLiveContractClausesInDOM = reorderContractClausesInDOM;

export function scrollToContractClauseByIdInDOM(
  rootEl: HTMLElement | null,
  clauseId: string,
  fallbackDomIndex?: number
): void {
  if (!rootEl) return;
  const byId = rootEl.querySelector(
    `[data-clause-id="${CSS.escape(clauseId)}"]`
  ) as HTMLElement | null;
  if (byId) {
    scrollToContractClauseInDOM(rootEl, clauseId);
  } else if (fallbackDomIndex !== undefined) {
    scrollToContractClauseInDOM(rootEl, fallbackDomIndex);
  }
}
