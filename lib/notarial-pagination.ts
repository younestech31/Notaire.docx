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
 * whether they are standalone `<p>` lines or legacy `<span>` tags.
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
 * Measures real rendered block heights inside a hidden offscreen 120mm net text column
 * matching the official Algerian Notarial A4 specification:
 * - Page width: 210mm, Right margin: 70mm, Left margin: 20mm -> Net column width = 120mm
 * - Page height: 297mm, Top margin: 10mm, Bottom margin: 60mm -> Net printable height = 227mm
 * - Font: Arial 13pt, line-height: 1.0
 */
function measureBlocksHeightsIn12cmProbe(blocks: HTMLElement[]): {
  heightsPx: number[];
  maxPageHeightPx: number;
  lineHeightPx: number;
} | null {
  if (typeof document === 'undefined' || !document.body) return null;

  const probe = document.createElement('div');
  probe.setAttribute('dir', 'rtl');
  probe.setAttribute('aria-hidden', 'true');
  probe.style.position = 'fixed';
  probe.style.top = '-9999px';
  probe.style.left = '-9999px';
  probe.style.width = '120mm';
  probe.style.minHeight = '227mm';
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

  // Reference ruler element for exact 227mm printable height and 1 line height in current DPI
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

  probe.appendChild(heightRuler);
  probe.appendChild(sampleLine);

  const clones: HTMLElement[] = blocks.map((el) => {
    const clone = el.cloneNode(true) as HTMLElement;
    clone.style.margin = '0';
    if (!clone.style.lineHeight) clone.style.lineHeight = '1';
    if (!clone.style.fontFamily) clone.style.fontFamily = STRICT_FONT_FAMILY;
    if (!clone.style.fontSize) clone.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    probe.appendChild(clone);
    return clone;
  });

  document.body.appendChild(probe);

  try {
    const rawPageHeightPx = heightRuler.getBoundingClientRect().height || 858;
    const lineHeightPx = sampleLine.getBoundingClientRect().height || 17.5;
    // Reserve one standalone line height at the bottom of the 227mm printable area for `../..`
    const maxPageHeightPx = Math.max(200, rawPageHeightPx - lineHeightPx * 1.25);
    const heightsPx = clones.map((c) => {
      const h = c.getBoundingClientRect().height;
      return h > 0 ? h : lineHeightPx;
    });
    return { heightsPx, maxPageHeightPx, lineHeightPx };
  } finally {
    if (probe.parentNode) {
      probe.parentNode.removeChild(probe);
    }
  }
}

/**
 * Injects official Algerian Notarial pagination markers into the contract HTML:
 * - As a standalone line (`<p dir="rtl" style="...text-align:left;">`) immediately BELOW the last line of text on each page
 * - Left-aligned within the 12cm net text column (`text-align: left; direction: rtl`)
 * - Non-final pages: `../..`
 * - Final page: `الصفحة {العدد بالحروف} والأخيرة`
 * - Never in headers/footers and never numeric `1/9`
 * Uses real DOM height measurement inside an offscreen 120mm x 227mm probe (with fallback when SSR).
 */
export function injectNotarialPaginationIntoHtml(
  html: string,
  pageNumberingEnabled: boolean = true
): string {
  const cleanHtml = stripNotarialPaginationFromHtml(html);
  if (!pageNumberingEnabled || !cleanHtml.trim()) return cleanHtml;

  if (typeof window !== 'undefined' && typeof DOMParser !== 'undefined') {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(`<div>${cleanHtml}</div>`, 'text/html');
      const root = doc.body.firstElementChild as HTMLElement | null;
      if (!root) return cleanHtml;

      // Collect top-level or clause-level paragraphs, tables, and explicit page breaks in reading order
      const blocks = Array.from(
        root.querySelectorAll('p, li, table, div.page-break, [data-page-break="true"]')
      ).filter((el) => {
        if (el.closest('[data-clause-disabled="true"]')) return false;
        const inTable = el.tagName.toUpperCase() !== 'TABLE' && el.closest('table');
        return !inTable;
      }) as HTMLElement[];

      if (blocks.length === 0) return cleanHtml;

      const measured = measureBlocksHeightsIn12cmProbe(blocks);
      const maxPageHeightPx = measured ? measured.maxPageHeightPx : 42 * 17.5;
      const lineHeightPx = measured ? measured.lineHeightPx : 17.5;

      let heightOnCurrentPagePx = 0;
      let totalPagesCount = 1;
      const pageEndElements: { element: HTMLElement; insertPageBreakAfter: boolean }[] = [];
      let lastContentElement: HTMLElement | null = null;

      for (let i = 0; i < blocks.length; i++) {
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
            });
          }
          totalPagesCount++;
          heightOnCurrentPagePx = 0;
          continue;
        }

        const text = (el.textContent || '').replace(/\s+/g, ' ').trim();
        if (!text && el.tagName.toUpperCase() !== 'TABLE') continue;

        lastContentElement = el;
        let blockHeightPx = measured ? measured.heightsPx[i] : 0;
        if (!blockHeightPx || blockHeightPx <= 0) {
          if (el.tagName.toUpperCase() === 'TABLE') {
            const rowCount = el.querySelectorAll('tr').length || 1;
            blockHeightPx = rowCount * 2 * lineHeightPx;
          } else {
            const words = text.split(' ').filter(Boolean).length;
            blockHeightPx = Math.max(1, Math.ceil(words / 9)) * lineHeightPx;
          }
        }

        heightOnCurrentPagePx += blockHeightPx;

        // Check if there is still non-empty content after this block
        const hasMoreContentAhead = blocks.slice(i + 1).some((nextEl) => {
          if (
            nextEl.classList.contains('page-break') ||
            nextEl.getAttribute('data-page-break') === 'true'
          ) {
            return false;
          }
          return (nextEl.textContent || '').trim().length > 0;
        });

        if (heightOnCurrentPagePx >= maxPageHeightPx && hasMoreContentAhead) {
          pageEndElements.push({
            element: el,
            insertPageBreakAfter: true,
          });
          totalPagesCount++;
          heightOnCurrentPagePx = 0;
        }
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
      for (const { element, insertPageBreakAfter } of pageEndElements) {
        const markerP = createStandalonePaginationParagraph('../..');
        element.insertAdjacentElement('afterend', markerP);
        if (insertPageBreakAfter) {
          const pbDiv = doc.createElement('div');
          pbDiv.className = 'page-break';
          pbDiv.setAttribute('contenteditable', 'false');
          pbDiv.setAttribute('data-page-break', 'true');
          pbDiv.setAttribute('data-notary-pagination', 'true');
          pbDiv.textContent = 'فاصل صفحات';
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
    }
  }

  return cleanHtml;
}

/**
 * Counts words and measured A4 pages from HTML string for the Bottom Status Bar.
 */
export function computeDocumentMetrics(html: string): {
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

  if (typeof window !== 'undefined' && typeof DOMParser !== 'undefined' && cleanHtml.trim()) {
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

        const measured = measureBlocksHeightsIn12cmProbe(blocks);
        if (measured) {
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
            accHeight += measured.heightsPx[i] || measured.lineHeightPx;
            if (accHeight >= measured.maxPageHeightPx && i < blocks.length - 1) {
              pages++;
              accHeight = 0;
            }
          }
          return { wordCount, charCount, estimatedPages: Math.max(1, pages) };
        }
      }
    } catch {
      // Fallback below
    }
  }

  const explicitBreaks = (
    cleanHtml.match(/class="[^"]*page-break[^"]*"|data-page-break="true"/gi) || []
  ).length;
  const estimatedPages = Math.max(1, Math.ceil(wordCount / 320) + explicitBreaks);

  return { wordCount, charCount, estimatedPages };
}
