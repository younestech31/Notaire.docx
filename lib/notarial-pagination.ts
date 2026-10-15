import {
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  NOTARIAL_PAGE_CONTENT_LINES,
  STRICT_LINE_SPACING_TWIPS,
} from './types';

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
 * Injects official Algerian Notarial fixed-line grid pagination markers into contract HTML:
 * - Exactly 46 lines per page (`NOTARIAL_PAGE_CONTENT_LINES`)
 * - Exact line spacing 260 twips (13pt)
 * - Continuation marker (`../..`) on non-final pages, closing marker («... والأخيرة») on the final page
 * - Page break div after each page marker.
 */
export function injectNotarialPaginationIntoHtml(
  html: string,
  pageNumberingEnabled: boolean = true
): string {
  const cleanHtml = stripNotarialPaginationFromHtml(html);
  if (!pageNumberingEnabled || !cleanHtml.trim()) return cleanHtml;

  const LINES = NOTARIAL_PAGE_CONTENT_LINES; // 46 lines

  if (typeof window !== 'undefined' && typeof DOMParser !== 'undefined') {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(`<div>${cleanHtml}</div>`, 'text/html');
      const root = doc.body.firstElementChild as HTMLElement | null;
      if (!root) return cleanHtml;

      const rawBlocks = Array.from(root.children) as HTMLElement[];
      if (rawBlocks.length === 0) return cleanHtml;

      interface PageGroup {
        elements: HTMLElement[];
        linesUsed: number;
      }

      const pages: PageGroup[] = [];
      let currentPage: PageGroup = { elements: [], linesUsed: 0 };

      const getBlockEstimatedLines = (el: HTMLElement): number => {
        const tag = el.tagName.toUpperCase();
        if (tag === 'TABLE') {
          const rows = el.querySelectorAll('tr').length || 1;
          return Math.max(2, rows * 2);
        }
        if (el.classList.contains('page-break') || el.getAttribute('data-page-break') === 'true') {
          return 0;
        }
        const words = countLogicalWordsInHtml(el.innerHTML);
        return Math.max(1, Math.ceil(words / 9));
      };

      for (let i = 0; i < rawBlocks.length; i++) {
        const el = rawBlocks[i];
        const isPb = el.classList.contains('page-break') || el.getAttribute('data-page-break') === 'true';

        if (isPb) {
          if (currentPage.elements.length > 0) {
            pages.push(currentPage);
            currentPage = { elements: [], linesUsed: 0 };
          }
          continue;
        }

        const estLines = getBlockEstimatedLines(el);

        if (currentPage.linesUsed + estLines <= LINES - 1) {
          currentPage.elements.push(el);
          currentPage.linesUsed += estLines;
        } else {
          const tag = el.tagName.toUpperCase();
          const availLines = (LINES - 1) - currentPage.linesUsed;
          if ((tag === 'P' || tag === 'LI') && availLines >= 2) {
            const splitResult = splitHtmlParagraphPreservingTags(el.innerHTML, availLines * 9);
            if (splitResult) {
              const firstEl = el.cloneNode(false) as HTMLElement;
              firstEl.innerHTML = splitResult.firstHtml;
              currentPage.elements.push(firstEl);
              currentPage.linesUsed = LINES - 1;
              pages.push(currentPage);

              currentPage = { elements: [], linesUsed: 0 };
              const secondEl = el.cloneNode(false) as HTMLElement;
              secondEl.innerHTML = splitResult.secondHtml;
              const remLines = getBlockEstimatedLines(secondEl);
              currentPage.elements.push(secondEl);
              currentPage.linesUsed += remLines;
              continue;
            }
          }

          if (currentPage.elements.length > 0) {
            pages.push(currentPage);
            currentPage = { elements: [], linesUsed: 0 };
          }
          currentPage.elements.push(el);
          currentPage.linesUsed += estLines;
        }
      }

      if (currentPage.elements.length > 0) {
        pages.push(currentPage);
      }

      if (pages.length === 0) return cleanHtml;

      const totalPages = pages.length;
      const newRoot = doc.createElement('div');

      const pBaseStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;line:${STRICT_LINE_SPACING_TWIPS};lineRule:exact;`;
      const markerStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:left;direction:rtl;font-weight:bold;line:${STRICT_LINE_SPACING_TWIPS};lineRule:exact;`;

      pages.forEach((page, idx) => {
        const pageNum = idx + 1;
        const isFinal = pageNum === totalPages;

        while (page.linesUsed < LINES - 1) {
          const emptyP = doc.createElement('p');
          emptyP.setAttribute('dir', 'rtl');
          emptyP.setAttribute('style', pBaseStyle);
          emptyP.innerHTML = '<br>';
          page.elements.push(emptyP);
          page.linesUsed++;
        }

        const markerP = doc.createElement('p');
        markerP.setAttribute('dir', 'rtl');
        markerP.setAttribute('align', 'left');
        markerP.setAttribute('data-notary-pagination', 'true');
        markerP.setAttribute('style', markerStyle);
        markerP.textContent = isFinal
          ? buildNotarialPaginationClosingText(totalPages)
          : buildNotarialPaginationContinueText(pageNum);
        page.elements.push(markerP);
        page.linesUsed++;

        for (const el of page.elements) {
          newRoot.appendChild(el.cloneNode(true));
        }

        if (!isFinal) {
          const pbDiv = doc.createElement('div');
          pbDiv.className = 'page-break';
          pbDiv.setAttribute('contenteditable', 'false');
          pbDiv.setAttribute('data-page-break', 'true');
          pbDiv.setAttribute('data-notary-pagination', 'true');
          newRoot.appendChild(pbDiv);
        }
      });

      return newRoot.innerHTML;
    } catch {
      return cleanHtml;
    }
  }

  return cleanHtml;
}

/**
 * Counts words and measured pages from HTML string for document metrics.
 */
export function computeDocumentMetrics(
  html: string,
  _useDomMeasurement: boolean = true
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

  const markerCount = (html.match(/data-notary-pagination=["']true["']/gi) || []).length;
  const estimatedPages = Math.max(1, markerCount || Math.ceil(wordCount / 380));

  return { wordCount, charCount, estimatedPages };
}
