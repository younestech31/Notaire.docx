import { STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';

/**
 * Converts a page number (1, 2, 3...) into its official Arabic feminine ordinal word
 * for the closing line of the contract ("الصفحة التاسعة والأخيرة").
 */
export function getArabicOrdinalPageWord(pageNumber: number): string {
  const n = Math.max(1, Math.floor(pageNumber));
  const map: Record<number, string> = {
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
    21: 'الحادية والعشرون',
    22: 'الثانية والعشرون',
    23: 'الثالثة والعشرون',
    24: 'الرابعة والعشرون',
    25: 'الخامسة والعشرون',
    26: 'السادسة والعشرون',
    27: 'السابعة والعشرون',
    28: 'الثامنة والعشرون',
    29: 'التاسعة والعشرون',
    30: 'الثلاثون',
    31: 'الحادية والثلاثون',
    32: 'الثانية والثلاثون',
    33: 'الثالثة والثلاثون',
    34: 'الرابعة والثلاثون',
    35: 'الخامسة والثلاثون',
    36: 'السادسة والثلاثون',
    37: 'السابعة والثلاثون',
    38: 'الثامنة والثلاثون',
    39: 'التاسعة والثلاثون',
    40: 'الأربعون',
    41: 'الحادية والأربعون',
    42: 'الثانية والأربعون',
    43: 'الثالثة والأربعون',
    44: 'الرابعة والأربعون',
    45: 'الخامسة والأربعون',
    46: 'السادسة والأربعون',
    47: 'السابعة والأربعون',
    48: 'الثامنة والأربعون',
    49: 'التاسعة والأربعون',
    50: 'الخمسون',
  };
  return map[n] || `${n}`;
}

export function buildNotarialPaginationClosingText(totalPages: number): string {
  return `الصفحة ${getArabicOrdinalPageWord(totalPages)} والأخيرة`;
}

/**
 * Strips any previously injected transient pagination markers (`data-notary-pagination="true"`)
 * whether they are standalone `<p>` lines, page-break `<div>`s, or legacy `<span>` tags.
 */
export function stripNotarialPaginationFromHtml(html: string): string {
  if (!html) return '';
  return html
    .replace(
      /<p[^>]*data-notary-pagination="true"[^>]*>[\s\S]*?<\/p>/gi,
      ''
    )
    .replace(
      /<div[^>]*data-notary-pagination="true"[^>]*>[\s\S]*?<\/div>/gi,
      ''
    )
    .replace(
      /<span[^>]*data-notary-pagination="true"[^>]*>[\s\S]*?<\/span>/gi,
      ''
    );
}

/**
 * Creates a hidden offscreen 120mm net text column probe matching the official Algerian Notarial A4 specification:
 * - Page width: 210mm, Right margin: 70mm, Left margin: 20mm -> Net column width = 120mm
 * - Page height: 297mm, Top margin: 10mm, Bottom margin: 60mm -> Net printable height = 227mm
 * - Font: Arial 13pt, line-height: 1.0
 */
function create12cmNotarialProbe(): {
  probe: HTMLDivElement;
  maxPageHeightPx: number;
  lineHeightPx: number;
  measureElementHeight: (el: HTMLElement) => number;
  findFittingTextSplitIndex: (el: HTMLElement, availableHeightPx: number) => number | null;
  destroy: () => void;
} | null {
  if (typeof document === 'undefined' || !document.body) return null;

  const probe = document.createElement('div');
  probe.setAttribute('dir', 'rtl');
  probe.setAttribute('aria-hidden', 'true');
  probe.style.position = 'fixed';
  probe.style.top = '-9999px';
  probe.style.left = '-9999px';
  probe.style.width = '120mm';
  probe.style.visibility = 'hidden';
  probe.style.pointerEvents = 'none';
  probe.style.fontFamily = STRICT_FONT_FAMILY;
  probe.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
  probe.style.lineHeight = '1';
  probe.style.textAlign = 'justify';
  probe.style.direction = 'rtl';
  probe.style.boxSizing = 'border-box';
  probe.style.padding = '0';
  probe.style.margin = '0';

  const heightRuler = document.createElement('div');
  heightRuler.style.height = '227mm';
  heightRuler.style.width = '1px';
  heightRuler.style.padding = '0';
  heightRuler.style.margin = '0';

  const sampleLine = document.createElement('p');
  sampleLine.style.margin = '0';
  sampleLine.style.lineHeight = '1';
  sampleLine.style.fontFamily = STRICT_FONT_FAMILY;
  sampleLine.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
  sampleLine.textContent = '../..';

  const testSlot = document.createElement('div');
  testSlot.style.width = '100%';
  testSlot.style.margin = '0';
  testSlot.style.padding = '0';

  probe.appendChild(heightRuler);
  probe.appendChild(sampleLine);
  probe.appendChild(testSlot);
  document.body.appendChild(probe);

  const rawPageHeightPx = heightRuler.getBoundingClientRect().height || 858;
  const lineHeightPx = sampleLine.getBoundingClientRect().height || 17.5;
  // Reserve 1.5 lines at the bottom of the 227mm printable box for the standalone `../..` line
  const maxPageHeightPx = Math.max(220, rawPageHeightPx - lineHeightPx * 1.5);

  const applyStrictCloneStyles = (clone: HTMLElement) => {
    clone.style.margin = '0';
    if (!clone.style.lineHeight) clone.style.lineHeight = '1';
    if (!clone.style.fontFamily) clone.style.fontFamily = STRICT_FONT_FAMILY;
    if (!clone.style.fontSize) clone.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    if (clone.tagName.toUpperCase() === 'TABLE') {
      clone.style.width = '100%';
      clone.style.borderCollapse = 'collapse';
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
   * For a long plain/formatted paragraph `<p>` or `<li>` that overflows the remaining height on the current page,
   * binary-searches the word boundary that fits inside `availableHeightPx` in the 120mm column.
   * Returns the number of words that fit on the current page, or null if it shouldn't be split.
   */
  const findFittingTextSplitIndex = (
    el: HTMLElement,
    availableHeightPx: number
  ): number | null => {
    if (availableHeightPx < lineHeightPx * 2) return null;
    const tag = el.tagName.toUpperCase();
    if (tag !== 'P' && tag !== 'LI') return null;
    // Avoid splitting paragraphs that contain complex nested tables or images
    if (el.querySelector('table, img')) return null;

    const fullText = (el.textContent || '').replace(/\s+/g, ' ').trim();
    const words = fullText.split(' ').filter(Boolean);
    if (words.length < 16) return null;

    testSlot.innerHTML = '';
    const probeP = document.createElement('p');
    probeP.setAttribute('dir', el.getAttribute('dir') || 'rtl');
    probeP.style.margin = '0';
    probeP.style.lineHeight = '1';
    probeP.style.fontFamily = STRICT_FONT_FAMILY;
    probeP.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    probeP.style.textAlign = el.style.textAlign || 'justify';
    if (el.style.fontWeight) probeP.style.fontWeight = el.style.fontWeight;
    testSlot.appendChild(probeP);

    let low = 4;
    let high = words.length - 4;
    let bestWordCount: number | null = null;

    while (low <= high) {
      const mid = Math.floor((low + high) / 2);
      probeP.textContent = words.slice(0, mid).join(' ');
      const h = probeP.getBoundingClientRect().height;
      if (h <= availableHeightPx) {
        bestWordCount = mid;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    if (!bestWordCount || bestWordCount < 6 || bestWordCount > words.length - 4) {
      return null;
    }
    return bestWordCount;
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
    findFittingTextSplitIndex,
    destroy,
  };
}

/**
 * Injects official Algerian Notarial pagination markers into the contract HTML:
 * - As a standalone line (`<p dir="rtl" style="...text-align:left;">`) immediately BELOW the last line of text on each page
 * - Left-aligned within the 12cm net text column (`text-align: left; direction: rtl`)
 * - Non-final pages: `../..`
 * - Final page: `الصفحة {العدد بالحروف} والأخيرة`
 * - Never in headers/footers and never numeric `1/9`
 *
 * Uses real 120mm x 227mm DOM height measurement, places `../..` BEFORE any block that would overflow
 * the 227mm printable box, and splits extra-long paragraphs across page boundaries when needed.
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

      const maxPageHeightPx = probeCtx ? probeCtx.maxPageHeightPx : 41 * 17.5;
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
              insertPageBreakAfter: false, // explicit page-break div already follows
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
            const words = text.split(' ').filter(Boolean).length;
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
        // First, check if `el` is a long paragraph that can be cleanly split across the page boundary
        if (
          highPrecisionDomSplit &&
          probeCtx &&
          remainingHeightPx >= lineHeightPx * 2.5 &&
          (el.tagName.toUpperCase() === 'P' || el.tagName.toUpperCase() === 'LI') &&
          !el.querySelector('table, img, .smart-tag, [data-var]')
        ) {
          const splitWordCount = probeCtx.findFittingTextSplitIndex(
            el,
            remainingHeightPx
          );
          if (splitWordCount !== null) {
            const allWords = text.split(' ').filter(Boolean);
            const firstPartText = allWords.slice(0, splitWordCount).join(' ');
            const secondPartText = allWords.slice(splitWordCount).join(' ');

            if (firstPartText && secondPartText) {
              const continuationEl = el.cloneNode(false) as HTMLElement;
              continuationEl.textContent = secondPartText;
              el.textContent = firstPartText;
              el.insertAdjacentElement('afterend', continuationEl);

              pageEndElements.push({
                element: el,
                insertPageBreakAfter: true,
                pageNumber: totalPagesCount,
              });
              totalPagesCount++;
              heightOnCurrentPagePx = 0;
              lastContentElement = el;
              // Insert continuationEl into blocks right after `i` so it is processed at the top of the next page
              blocks.splice(i + 1, 0, continuationEl);
              i++;
              continue;
            }
          }
        }

        // Case C: If there is already content on the current page (`lastContentElement`),
        // close the current page right after `lastContentElement` (BEFORE `el`), and start `el` on the new page!
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

          // If `el` by itself is taller than a full page and there is still content after `el`, close this page too
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

        // Case D: `el` is the very first element on the page and exceeds `maxPageHeightPx`
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

      // 1. Insert standalone `../..` paragraph directly BELOW the last content element of each non-final page
      for (const { element, insertPageBreakAfter, pageNumber } of pageEndElements) {
        const markerP = createStandalonePaginationParagraph('../..');
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
      return cleanHtml;
    } finally {
      probeCtx?.destroy();
    }
  }

  return cleanHtml;
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
