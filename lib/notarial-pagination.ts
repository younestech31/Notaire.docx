import { STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';

/**
 * Converts a page number (1, 2, 3...) into its official Arabic feminine ordinal word
 * for the closing line of the contract ("الصفحة التاسعة والأخيرة").
 */
export function getArabicOrdinalPageWord(pageNumber: number): string {
  const n = Math.max(1, Math.floor(pageNumber));
  const firstTwenty: Record<number, string> = {
    1: 'الأولى',
    2: 'الثانية',
    3: 'الثالثة',
    4: 'الرابعة',
    5: 'الخامسة',
    6: 'السادسة',
    7: 'السابعة',
    8: 'الثامنة',
    9: 'التاسعة',
    10: 'العاشرة',
    11: 'الحادية عشرة',
    12: 'الثانية عشرة',
    13: 'الثالثة عشرة',
    14: 'الرابعة عشرة',
    15: 'الخامسة عشرة',
    16: 'السادسة عشرة',
    17: 'السابعة عشرة',
    18: 'الثامنة عشرة',
    19: 'التاسعة عشرة',
    20: 'العشرون',
  };
  if (firstTwenty[n]) return firstTwenty[n];

  const compoundUnits: Record<number, string> = {
    1: 'الحادية',
    2: 'الثانية',
    3: 'الثالثة',
    4: 'الرابعة',
    5: 'الخامسة',
    6: 'السادسة',
    7: 'السابعة',
    8: 'الثامنة',
    9: 'التاسعة',
  };
  const tensWords: Record<number, string> = {
    20: 'العشرون',
    30: 'الثلاثون',
    40: 'الأربعون',
    50: 'الخمسون',
    60: 'الستون',
    70: 'السبعون',
    80: 'الثمانون',
    90: 'التسعون',
  };

  if (n >= 21 && n <= 99) {
    const unit = n % 10;
    const ten = Math.floor(n / 10) * 10;
    if (unit === 0) return tensWords[ten];
    return `${compoundUnits[unit]} و${tensWords[ten]}`;
  }

  if (n === 100) return 'المائة';
  if (n >= 101 && n <= 199) {
    return `المائة و${getArabicOrdinalPageWord(n - 100)}`;
  }
  if (n === 200) return 'المائتان';

  return `رقم ${n}`;
}

export function buildNotarialPaginationContinueText(pageNumber: number): string {
  return `الصفحة ${getArabicOrdinalPageWord(pageNumber)} ../..`;
}

export function buildNotarialPaginationClosingText(totalPages: number): string {
  return `الصفحة ${getArabicOrdinalPageWord(totalPages)} والأخيرة`;
}

/**
 * Strips any previously injected notarial pagination markers (`data-notary-pagination="true"`)
 * from an HTML string so editing and re-pagination remain idempotent.
 */
export function stripNotarialPaginationFromHtml(html: string): string {
  if (!html) return '';
  return html
    .replace(
      /<(?:p|div)\b[^>]*\bdata-notary-pagination=["']true["'][^>]*>[\s\S]*?<\/(?:p|div)>/gi,
      ''
    )
    .trim();
}

const VOID_HTML_TAGS = new Set([
  'br',
  'hr',
  'img',
  'input',
  'meta',
  'link',
  'col',
  'area',
  'base',
  'embed',
  'param',
  'source',
  'track',
  'wbr',
]);

function isAtomicTagDeclaration(fullOpenTag: string): boolean {
  return /\b(?:smart-tag|smart-placeholder|data-var|data-placeholder)\b/i.test(
    fullOpenTag
  );
}

/**
 * Counts the number of logical words inside an HTML snippet, treating `[متغير]` or `.smart-tag`
 * chips as a single indivisible word.
 */
export function countLogicalWordsInHtml(innerHtml: string): number {
  if (!innerHtml) return 0;
  const tagOrTextRegex = /(<[^>]+>)|([^<]+)/g;
  const openTags: { tagName: string; fullOpenTag: string; isAtomic: boolean }[] = [];
  let wordCount = 0;
  let atomicHasCountedWord = false;
  let match: RegExpExecArray | null;

  while ((match = tagOrTextRegex.exec(innerHtml)) !== null) {
    const tagToken = match[1];
    const textToken = match[2];

    if (tagToken) {
      const closeMatch = tagToken.match(/^<\s*\/\s*([a-zA-Z0-9:-]+)/);
      if (closeMatch) {
        const closingName = closeMatch[1].toLowerCase();
        for (let k = openTags.length - 1; k >= 0; k--) {
          if (openTags[k].tagName === closingName) {
            openTags.splice(k, 1);
            break;
          }
        }
        if (!openTags.some((t) => t.isAtomic)) {
          atomicHasCountedWord = false;
        }
      } else {
        const openMatch = tagToken.match(/^<\s*([a-zA-Z0-9:-]+)/);
        const selfClosing = /\/\s*>$/.test(tagToken);
        if (openMatch && !selfClosing) {
          const tagName = openMatch[1].toLowerCase();
          if (!VOID_HTML_TAGS.has(tagName)) {
            openTags.push({
              tagName,
              fullOpenTag: tagToken,
              isAtomic: isAtomicTagDeclaration(tagToken),
            });
          }
        }
      }
    } else if (textToken) {
      const insideAtomic = openTags.some((t) => t.isAtomic);
      if (insideAtomic) {
        if (textToken.trim() && !atomicHasCountedWord) {
          wordCount++;
          atomicHasCountedWord = true;
        }
      } else {
        const words =
          textToken.match(/\[\s*[^\[\]<>]{1,60}?\s*\]|\{\{\s*[^}]+\s*\}\}|\S+/g) || [];
        wordCount += words.length;
      }
    }
  }

  return wordCount;
}

/**
 * Splits an innerHTML string at `targetWordCount` while preserving all inline HTML formatting
 * (`<strong>`, `<b>`, `<u>`, `<em>`, `<span ...>`, `.smart-tag`, `[متغير]`) across both halves.
 * Never splits inside an atomic `.smart-tag` / `[data-var]` or raw `[متغير]` placeholder.
 */
export function splitHtmlParagraphPreservingTags(
  innerHtml: string,
  targetWordCount: number
): { firstHtml: string; secondHtml: string } | null {
  if (!innerHtml || targetWordCount < 1) return null;

  const tagOrTextRegex = /(<[^>]+>)|([^<]+)/g;
  const openTags: { tagName: string; fullOpenTag: string; isAtomic: boolean }[] = [];
  let currentWords = 0;
  let atomicHasCountedWord = false;
  let splitCharIndex: number | null = null;
  let openTagsAtSplit: { tagName: string; fullOpenTag: string; isAtomic: boolean }[] = [];

  let match: RegExpExecArray | null;
  while ((match = tagOrTextRegex.exec(innerHtml)) !== null) {
    const tokenStart = match.index;
    const tagToken = match[1];
    const textToken = match[2];

    if (tagToken) {
      const closeMatch = tagToken.match(/^<\s*\/\s*([a-zA-Z0-9:-]+)/);
      if (closeMatch) {
        const closingName = closeMatch[1].toLowerCase();
        for (let k = openTags.length - 1; k >= 0; k--) {
          if (openTags[k].tagName === closingName) {
            openTags.splice(k, 1);
            break;
          }
        }
        const stillInsideAtomic = openTags.some((t) => t.isAtomic);
        if (!stillInsideAtomic) {
          atomicHasCountedWord = false;
          if (currentWords >= targetWordCount && splitCharIndex === null) {
            splitCharIndex = tokenStart + tagToken.length;
            openTagsAtSplit = openTags.map((t) => ({ ...t }));
            break;
          }
        }
      } else {
        const openMatch = tagToken.match(/^<\s*([a-zA-Z0-9:-]+)/);
        const selfClosing = /\/\s*>$/.test(tagToken);
        if (openMatch && !selfClosing) {
          const tagName = openMatch[1].toLowerCase();
          if (!VOID_HTML_TAGS.has(tagName)) {
            openTags.push({
              tagName,
              fullOpenTag: tagToken,
              isAtomic: isAtomicTagDeclaration(tagToken),
            });
          }
        }
      }
    } else if (textToken) {
      const insideAtomic = openTags.some((t) => t.isAtomic);
      if (insideAtomic) {
        if (textToken.trim() && !atomicHasCountedWord) {
          currentWords++;
          atomicHasCountedWord = true;
        }
      } else {
        const wordRegex = /\[\s*[^\[\]<>]{1,60}?\s*\]|\{\{\s*[^}]+\s*\}\}|\S+/g;
        let wMatch: RegExpExecArray | null;
        while ((wMatch = wordRegex.exec(textToken)) !== null) {
          currentWords++;
          if (currentWords >= targetWordCount) {
            splitCharIndex = tokenStart + wMatch.index + wMatch[0].length;
            openTagsAtSplit = openTags.map((t) => ({ ...t }));
            break;
          }
        }
        if (splitCharIndex !== null) {
          break;
        }
      }
    }
  }

  if (splitCharIndex === null) return null;

  const rawFirst = innerHtml.slice(0, splitCharIndex);
  const rawSecond = innerHtml.slice(splitCharIndex);

  // Ensure the remainder actually has visible text content
  const remainderPlain = rawSecond
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .trim();
  if (!remainderPlain) return null;

  const closingTagsForFirst = openTagsAtSplit
    .slice()
    .reverse()
    .map((t) => `</${t.tagName}>`)
    .join('');

  const reopeningTagsForSecond = openTagsAtSplit.map((t) => t.fullOpenTag).join('');

  // Trim leading whitespace on the first text node of rawSecond so the new page starts flush right
  const cleanedSecond = rawSecond.replace(/^(\s|&nbsp;)+/, '');

  return {
    firstHtml: `${rawFirst}${closingTagsForFirst}`,
    secondHtml: `${reopeningTagsForSecond}${cleanedSecond}`,
  };
}

/**
 * Creates a hidden offscreen DOM measurement container matching the exact Algerian Notarial A4
 * printable column dimensions:
 * - Net text width: 120mm (210mm A4 width - 70mm right margin - 20mm left margin)
 * - Net printable height: 227mm (297mm A4 height - 10mm top margin - 60mm bottom margin)
 * - Reserved bottom space for standalone `../..` line: ~9mm -> max content height per page = 218mm
 * - Font: Arial 13pt, line-height 1.0, RTL, justified.
 */
function create12cmNotarialProbe(): {
  probe: HTMLDivElement;
  maxPageHeightPx: number;
  lineHeightPx: number;
  measureElementHeight: (el: HTMLElement) => number;
  findFittingFormattedSplit: (
    el: HTMLElement,
    availableHeightPx: number
  ) => { firstHtml: string; secondHtml: string } | null;
  destroy: () => void;
} | null {
  if (typeof window === 'undefined' || typeof document === 'undefined') return null;

  const probe = document.createElement('div');
  probe.setAttribute('dir', 'rtl');
  probe.setAttribute('aria-hidden', 'true');
  probe.style.position = 'fixed';
  probe.style.top = '-10000px';
  probe.style.left = '-10000px';
  probe.style.width = '120mm';
  probe.style.minWidth = '120mm';
  probe.style.maxWidth = '120mm';
  probe.style.padding = '0';
  probe.style.margin = '0';
  probe.style.visibility = 'hidden';
  probe.style.pointerEvents = 'none';
  probe.style.zIndex = '-9999';
  probe.style.fontFamily = STRICT_FONT_FAMILY;
  probe.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
  probe.style.lineHeight = '1';
  probe.style.textAlign = 'justify';
  probe.style.direction = 'rtl';
  probe.style.boxSizing = 'border-box';

  // Height ruler: 218mm (227mm printable box minus 9mm reserved for the standalone `../..` line + margin)
  const heightRuler = document.createElement('div');
  heightRuler.style.height = '218mm';
  heightRuler.style.width = '1px';
  probe.appendChild(heightRuler);

  // Single line ruler
  const lineRuler = document.createElement('p');
  lineRuler.style.margin = '0';
  lineRuler.style.padding = '0';
  lineRuler.style.lineHeight = '1';
  lineRuler.style.fontFamily = STRICT_FONT_FAMILY;
  lineRuler.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
  lineRuler.textContent = 'سطر قياس المعايرة التوثيقية';
  probe.appendChild(lineRuler);

  const testSlot = document.createElement('div');
  testSlot.style.width = '100%';
  testSlot.style.margin = '0';
  testSlot.style.padding = '0';
  probe.appendChild(testSlot);

  document.body.appendChild(probe);

  const measuredMaxHeight = heightRuler.getBoundingClientRect().height;
  const measuredLineHeight = lineRuler.getBoundingClientRect().height;
  // Fallback if running in a headless/jsdom environment where getBoundingClientRect returns 0
  const lineHeightPx = measuredLineHeight > 4 ? measuredLineHeight : 17.5;
  const maxPageHeightPx =
    measuredMaxHeight > 100 ? measuredMaxHeight : 40 * lineHeightPx;

  const applyStrictCloneStyles = (clone: HTMLElement) => {
    clone.style.fontFamily = STRICT_FONT_FAMILY;
    clone.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    clone.style.lineHeight = '1';
    if (clone.tagName.toUpperCase() === 'TABLE') {
      clone.style.width = '100%';
      clone.style.borderCollapse = 'collapse';
    } else {
      clone.style.margin = '0';
    }
  };

  const measureElementHeight = (el: HTMLElement): number => {
    testSlot.innerHTML = '';
    const clone = el.cloneNode(true) as HTMLElement;
    applyStrictCloneStyles(clone);
    testSlot.appendChild(clone);
    const h = clone.getBoundingClientRect().height;
    return h > 0 ? h : lineHeightPx;
  };

  /**
   * Binary-searches the exact word split boundary for an overflowing `<p>` or `<li>` inside
   * the 120mm column while preserving all inner HTML tags (`<strong>`, `<u>`, `.smart-tag`, etc.).
   */
  const findFittingFormattedSplit = (
    el: HTMLElement,
    availableHeightPx: number
  ): { firstHtml: string; secondHtml: string } | null => {
    if (availableHeightPx < lineHeightPx * 2) return null;
    const tag = el.tagName.toUpperCase();
    if (tag !== 'P' && tag !== 'LI') return null;
    if (el.querySelector('table, img')) return null;

    const rawInnerHtml = el.innerHTML;
    const totalWords = countLogicalWordsInHtml(rawInnerHtml);
    if (totalWords < 16) return null;

    testSlot.innerHTML = '';
    const probeBlock = el.cloneNode(false) as HTMLElement;
    applyStrictCloneStyles(probeBlock);
    testSlot.appendChild(probeBlock);

    let low = 4;
    let high = totalWords - 4;
    let bestSplit: { firstHtml: string; secondHtml: string } | null = null;

    while (low <= high) {
      const mid = Math.floor((low + high) / 2);
      const candidate = splitHtmlParagraphPreservingTags(rawInnerHtml, mid);
      if (!candidate) {
        high = mid - 1;
        continue;
      }
      probeBlock.innerHTML = candidate.firstHtml;
      const measuredH = probeBlock.getBoundingClientRect().height;
      // If running in headless DOM where height is 0, estimate 9 words per 12cm line
      const effectiveH =
        measuredH > 0 ? measuredH : Math.ceil(mid / 9) * lineHeightPx;

      if (effectiveH <= availableHeightPx) {
        bestSplit = candidate;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    return bestSplit;
  };

  const destroy = () => {
    if (probe.parentNode) {
      probe.parentNode.removeChild(probe);
    }
  };

  return {
    probe,
    maxPageHeightPx,
    lineHeightPx,
    measureElementHeight,
    findFittingFormattedSplit,
    destroy,
  };
}

/**
 * Pure-JS fallback pagination injector used when `DOMParser` / `window` is unavailable
 * (e.g. Node.js automated test suites or server-side verification).
 */
function injectNotarialPaginationFallback(
  cleanHtml: string,
  highPrecisionSplit: boolean
): string {
  const WORDS_PER_LINE = 9;
  const MAX_LINES_PER_PAGE = 40;

  const blockRegex =
    /<(p|li|table|div)\b([^>]*)>([\s\S]*?)<\/\1>|<div\b[^>]*class=["'][^"']*page-break[^"']*["'][^>]*>\s*<\/div>/gi;

  interface ParsedBlock {
    fullMatch: string;
    tag: string;
    attrs: string;
    inner: string;
    isPageBreak: boolean;
    isDisabledClause: boolean;
  }

  const blocks: ParsedBlock[] = [];
  let match: RegExpExecArray | null;
  while ((match = blockRegex.exec(cleanHtml)) !== null) {
    const fullMatch = match[0];
    const tag = (match[1] || 'div').toLowerCase();
    const attrs = match[2] || '';
    const inner = match[3] || '';
    const isPageBreak =
      /class=["'][^"']*page-break[^"']*["']/i.test(fullMatch) ||
      /data-page-break=["']true["']/i.test(fullMatch);
    const isDisabledClause = /data-clause-disabled=["']true["']/i.test(attrs);
    blocks.push({ fullMatch, tag, attrs, inner, isPageBreak, isDisabledClause });
  }

  if (blocks.length === 0) {
    const closing = buildNotarialPaginationClosingText(1);
    return `${cleanHtml}<p dir="rtl" align="left" data-notary-pagination="true" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:left;direction:rtl;font-weight:bold;">${closing}</p>`;
  }

  const makeMarkerP = (markerText: string) =>
    `<p dir="rtl" align="left" data-notary-pagination="true" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:left;direction:rtl;font-weight:bold;">${markerText}</p>`;

  const makePageBreakDiv = (pageNumber: number) =>
    `<div class="page-break" contenteditable="false" data-page-break="true" data-notary-pagination="true">نهاية الصفحة ${pageNumber} — بداية الصفحة ${pageNumber + 1}</div>`;

  const outputBlocks: string[] = [];
  let linesOnCurrentPage = 0;
  let totalPages = 1;

  let i = 0;
  while (i < blocks.length) {
    const b = blocks[i];
    if (b.isDisabledClause) {
      outputBlocks.push(b.fullMatch);
      i++;
      continue;
    }

    if (b.isPageBreak) {
      outputBlocks.push(makeMarkerP(buildNotarialPaginationContinueText(totalPages)));
      outputBlocks.push(b.fullMatch);
      totalPages++;
      linesOnCurrentPage = 0;
      i++;
      continue;
    }

    const words = countLogicalWordsInHtml(b.inner);
    const estLines =
      b.tag === 'table'
        ? Math.max(2, (b.inner.match(/<tr\b/gi) || []).length * 2)
        : Math.max(1, Math.ceil(words / WORDS_PER_LINE));

    if (linesOnCurrentPage + estLines <= MAX_LINES_PER_PAGE) {
      outputBlocks.push(b.fullMatch);
      linesOnCurrentPage += estLines;
      i++;
      continue;
    }

    const remainingLines = MAX_LINES_PER_PAGE - linesOnCurrentPage;
    if (
      highPrecisionSplit &&
      remainingLines >= 3 &&
      (b.tag === 'p' || b.tag === 'li') &&
      !/<(table|img)\b/i.test(b.inner) &&
      words >= 16
    ) {
      const targetWords = Math.min(words - 4, remainingLines * WORDS_PER_LINE);
      if (targetWords >= 6) {
        const splitRes = splitHtmlParagraphPreservingTags(b.inner, targetWords);
        if (splitRes) {
          outputBlocks.push(`<${b.tag}${b.attrs}>${splitRes.firstHtml}</${b.tag}>`);
          outputBlocks.push(makeMarkerP(buildNotarialPaginationContinueText(totalPages)));
          outputBlocks.push(makePageBreakDiv(totalPages));
          totalPages++;
          linesOnCurrentPage = 0;
          blocks.splice(i + 1, 0, {
            fullMatch: `<${b.tag}${b.attrs}>${splitRes.secondHtml}</${b.tag}>`,
            tag: b.tag,
            attrs: b.attrs,
            inner: splitRes.secondHtml,
            isPageBreak: false,
            isDisabledClause: false,
          });
          i++;
          continue;
        }
      }
    }

    if (linesOnCurrentPage > 0) {
      outputBlocks.push(makeMarkerP(buildNotarialPaginationContinueText(totalPages)));
      outputBlocks.push(makePageBreakDiv(totalPages));
      totalPages++;
      linesOnCurrentPage = 0;
      continue;
    }

    outputBlocks.push(b.fullMatch);
    if (i < blocks.length - 1) {
      outputBlocks.push(makeMarkerP(buildNotarialPaginationContinueText(totalPages)));
      outputBlocks.push(makePageBreakDiv(totalPages));
      totalPages++;
      linesOnCurrentPage = 0;
    } else {
      linesOnCurrentPage = estLines;
    }
    i++;
  }

  const closingText = buildNotarialPaginationClosingText(totalPages);
  outputBlocks.push(makeMarkerP(closingText));
  return outputBlocks.join('\n');
}

/**
 * Injects official Algerian Notarial pagination markers into the contract HTML:
 * - As a standalone line (`<p dir="rtl" style="...text-align:left;">`) immediately BELOW the last line of text on each page
 * - Left-aligned within the 12cm net text column (`text-align: left; direction: rtl`)
 * - Non-final pages: `الصفحة {العدد بالحروف} ../..`
 * - Final page: `الصفحة {العدد بالحروف} والأخيرة`
 * - Never in headers/footers and never numeric `1/9`
 *
 * Uses real 120mm x 227mm DOM height measurement, places `../..` BEFORE any block that would overflow
 * the 227mm printable box, and splits extra-long paragraphs across page boundaries while preserving
 * all inner HTML formatting (`<strong>`, `<u>`, `.smart-tag`, `[متغير]`).
 */
export function injectNotarialPaginationIntoHtml(
  html: string,
  pageNumberingEnabled: boolean = true,
  highPrecisionDomSplit: boolean = true
): string {
  const cleanHtml = stripNotarialPaginationFromHtml(html);
  if (!pageNumberingEnabled || !cleanHtml.trim()) return cleanHtml;

  if (typeof window !== 'undefined' && typeof DOMParser !== 'undefined') {
    const probeCtx = create12cmNotarialProbe();
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(`<div>${cleanHtml}</div>`, 'text/html');
      const root = doc.body.firstElementChild as HTMLElement | null;
      if (!root) return cleanHtml;

      const collectBlocks = (): HTMLElement[] =>
        Array.from(
          root.querySelectorAll('p, li, table, div.page-break, [data-page-break="true"]')
        ).filter((el) => {
          if (el.closest('[data-clause-disabled="true"]')) return false;
          const inTable = el.tagName.toUpperCase() !== 'TABLE' && el.closest('table');
          return !inTable;
        }) as HTMLElement[];

      const blocks = collectBlocks();
      if (blocks.length === 0) return cleanHtml;

      const maxPageHeightPx = probeCtx ? probeCtx.maxPageHeightPx : 40 * 17.5;
      const lineHeightPx = probeCtx ? probeCtx.lineHeightPx : 17.5;

      let heightOnCurrentPagePx = 0;
      let totalPagesCount = 1;
      const pageEndElements: {
        element: HTMLElement;
        insertPageBreakAfter: boolean;
        pageNumber: number;
      }[] = [];
      let lastContentElement: HTMLElement | null = null;

      let i = 0;
      while (i < blocks.length) {
        const el = blocks[i];
        const isPageBreak =
          el.classList.contains('page-break') ||
          el.getAttribute('data-page-break') === 'true';

        if (isPageBreak) {
          if (
            lastContentElement &&
            !pageEndElements.some((item) => item.element === lastContentElement)
          ) {
            pageEndElements.push({
              element: lastContentElement,
              insertPageBreakAfter: false,
              pageNumber: totalPagesCount,
            });
          }
          totalPagesCount++;
          heightOnCurrentPagePx = 0;
          i++;
          continue;
        }

        const text = (el.textContent || '').replace(/\s+/g, ' ').trim();
        if (!text && el.tagName.toUpperCase() !== 'TABLE') {
          i++;
          continue;
        }

        let blockHeightPx = probeCtx ? probeCtx.measureElementHeight(el) : 0;
        if (!blockHeightPx || blockHeightPx <= 0) {
          if (el.tagName.toUpperCase() === 'TABLE') {
            const rowCount = el.querySelectorAll('tr').length || 1;
            blockHeightPx = rowCount * 2 * lineHeightPx;
          } else {
            const words = countLogicalWordsInHtml(el.innerHTML);
            blockHeightPx = Math.max(1, Math.ceil(words / 9)) * lineHeightPx;
          }
        }

        const remainingHeightPx = maxPageHeightPx - heightOnCurrentPagePx;

        // Case A: Element fits on the current page
        if (heightOnCurrentPagePx + blockHeightPx <= maxPageHeightPx) {
          heightOnCurrentPagePx += blockHeightPx;
          lastContentElement = el;
          i++;
          continue;
        }

        // Case B: Element overflows the current page!
        // Check if `el` is a long paragraph (`<p>` or `<li>`) that can be split while preserving inner HTML tags
        if (
          highPrecisionDomSplit &&
          probeCtx &&
          remainingHeightPx >= lineHeightPx * 2.5 &&
          (el.tagName.toUpperCase() === 'P' || el.tagName.toUpperCase() === 'LI') &&
          !el.querySelector('table, img')
        ) {
          const formattedSplit = probeCtx.findFittingFormattedSplit(
            el,
            remainingHeightPx
          );
          if (formattedSplit) {
            const continuationEl = el.cloneNode(false) as HTMLElement;
            el.innerHTML = formattedSplit.firstHtml;
            continuationEl.innerHTML = formattedSplit.secondHtml;
            el.insertAdjacentElement('afterend', continuationEl);

            pageEndElements.push({
              element: el,
              insertPageBreakAfter: true,
              pageNumber: totalPagesCount,
            });
            totalPagesCount++;
            heightOnCurrentPagePx = 0;
            lastContentElement = el;
            blocks.splice(i + 1, 0, continuationEl);
            i++;
            continue;
          }
        }

        // Case C: Close the current page right after `lastContentElement` (BEFORE `el`), and start `el` on the new page
        if (
          heightOnCurrentPagePx > 0 &&
          lastContentElement &&
          !pageEndElements.some((item) => item.element === lastContentElement)
        ) {
          pageEndElements.push({
            element: lastContentElement,
            insertPageBreakAfter: true,
            pageNumber: totalPagesCount,
          });
          totalPagesCount++;
          heightOnCurrentPagePx = blockHeightPx;
          lastContentElement = el;

          const hasMoreAfterEl = blocks.slice(i + 1).some((nextEl) => {
            if (
              nextEl.classList.contains('page-break') ||
              nextEl.getAttribute('data-page-break') === 'true'
            ) {
              return false;
            }
            return (
              nextEl.tagName.toUpperCase() === 'TABLE' ||
              (nextEl.textContent || '').trim().length > 0
            );
          });
          if (heightOnCurrentPagePx >= maxPageHeightPx && hasMoreAfterEl) {
            pageEndElements.push({
              element: el,
              insertPageBreakAfter: true,
              pageNumber: totalPagesCount,
            });
            totalPagesCount++;
            heightOnCurrentPagePx = 0;
          }
          i++;
          continue;
        }

        // Case D: `el` is the first element on the page and exceeds `maxPageHeightPx`
        lastContentElement = el;
        const hasMoreContentAhead = blocks.slice(i + 1).some((nextEl) => {
          if (
            nextEl.classList.contains('page-break') ||
            nextEl.getAttribute('data-page-break') === 'true'
          ) {
            return false;
          }
          return (
            nextEl.tagName.toUpperCase() === 'TABLE' ||
            (nextEl.textContent || '').trim().length > 0
          );
        });

        if (hasMoreContentAhead) {
          pageEndElements.push({
            element: el,
            insertPageBreakAfter: true,
            pageNumber: totalPagesCount,
          });
          totalPagesCount++;
          heightOnCurrentPagePx = 0;
        } else {
          heightOnCurrentPagePx = blockHeightPx;
        }
        i++;
      }

      // Ensure lastContentElement is never marked as a non-final page end
      if (
        lastContentElement &&
        pageEndElements.length > 0 &&
        pageEndElements[pageEndElements.length - 1].element === lastContentElement
      ) {
        pageEndElements.pop();
        totalPagesCount = Math.max(1, totalPagesCount - 1);
      }

      const createStandalonePaginationParagraph = (markerText: string): HTMLElement => {
        const p = doc.createElement('p');
        p.setAttribute('dir', 'rtl');
        p.setAttribute('align', 'left');
        p.setAttribute('data-notary-pagination', 'true');
        p.style.margin = '0';
        p.style.lineHeight = '1';
        p.style.fontFamily = STRICT_FONT_FAMILY;
        p.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
        p.style.textAlign = 'left';
        p.style.direction = 'rtl';
        p.style.fontWeight = 'bold';
        p.textContent = markerText;
        return p;
      };

      // 1. Insert standalone marker paragraph directly BELOW the last content element of each non-final page
      for (const { element, insertPageBreakAfter, pageNumber } of pageEndElements) {
        const markerP = createStandalonePaginationParagraph(
          buildNotarialPaginationContinueText(pageNumber)
        );
        element.insertAdjacentElement('afterend', markerP);
        if (insertPageBreakAfter) {
          const pbDiv = doc.createElement('div');
          pbDiv.className = 'page-break';
          pbDiv.setAttribute('contenteditable', 'false');
          pbDiv.setAttribute('data-page-break', 'true');
          pbDiv.setAttribute('data-notary-pagination', 'true');
          pbDiv.textContent = `نهاية الصفحة ${pageNumber} — بداية الصفحة ${pageNumber + 1}`;
          markerP.insertAdjacentElement('afterend', pbDiv);
        }
      }

      // 2. Insert standalone `الصفحة {العدد بالحروف} والأخيرة` paragraph directly BELOW the final content element
      if (lastContentElement) {
        const closingText = buildNotarialPaginationClosingText(totalPagesCount);
        const closingP = createStandalonePaginationParagraph(closingText);
        lastContentElement.insertAdjacentElement('afterend', closingP);
      }

      return root.innerHTML;
    } catch {
      return injectNotarialPaginationFallback(cleanHtml, highPrecisionDomSplit);
    } finally {
      probeCtx?.destroy();
    }
  }

  return injectNotarialPaginationFallback(cleanHtml, highPrecisionDomSplit);
}

/**
 * Counts words and measured A4 pages from HTML string for the Bottom Status Bar.
 * Supports `useDomMeasurement` flag so keystroke handlers can stay fast while idle/print/seal use exact DOM measurement.
 */
export function computeDocumentMetrics(
  html: string,
  useDomMeasurement: boolean = true
): {
  wordCount: number;
  charCount: number;
  estimatedPages: number;
} {
  const cleanHtml = stripNotarialPaginationFromHtml(html || '');
  const plain = cleanHtml
    .replace(/<div[^>]*data-clause-disabled="true"[^>]*>[\s\S]*?<\/div>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const words = plain ? plain.split(' ').filter(Boolean) : [];
  const wordCount = words.length;
  const charCount = plain.length;

  if (
    useDomMeasurement &&
    typeof window !== 'undefined' &&
    typeof DOMParser !== 'undefined' &&
    cleanHtml.trim()
  ) {
    const probeCtx = create12cmNotarialProbe();
    if (probeCtx) {
      try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(`<div>${cleanHtml}</div>`, 'text/html');
        const root = doc.body.firstElementChild as HTMLElement | null;
        if (root) {
          const blocks = Array.from(
            root.querySelectorAll('p, li, table, div.page-break, [data-page-break="true"]')
          ).filter((el) => {
            if (el.closest('[data-clause-disabled="true"]')) return false;
            const inTable = el.tagName.toUpperCase() !== 'TABLE' && el.closest('table');
            return !inTable;
          }) as HTMLElement[];

          let pages = 1;
          let accHeight = 0;
          for (let i = 0; i < blocks.length; i++) {
            const el = blocks[i];
            const isPb =
              el.classList.contains('page-break') ||
              el.getAttribute('data-page-break') === 'true';
            if (isPb) {
              pages++;
              accHeight = 0;
              continue;
            }
            const text = (el.textContent || '').trim();
            if (!text && el.tagName.toUpperCase() !== 'TABLE') continue;
            const h = probeCtx.measureElementHeight(el);
            if (accHeight > 0 && accHeight + h > probeCtx.maxPageHeightPx) {
              pages++;
              accHeight = h;
            } else {
              accHeight += h;
            }
          }
          return { wordCount, charCount, estimatedPages: Math.max(1, pages) };
        }
      } catch {
        // Fallback below
      } finally {
        probeCtx.destroy();
      }
    }
  }

  const explicitBreaks = (
    cleanHtml.match(/class="[^"]*page-break[^"]*"|data-page-break="true"/gi) || []
  ).length;
  const estimatedPages = Math.max(1, Math.ceil(wordCount / 320) + explicitBreaks);

  return { wordCount, charCount, estimatedPages };
}
