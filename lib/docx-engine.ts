import JSZip from 'jszip';
import { buildVirtualTableGrid } from './table-grid';
import {
  A4_SIZE_TWIPS,
  NET_CONTENT_WIDTH_TWIPS,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_HALF_PT,
  STRICT_FONT_SIZE_PT,
  STRICT_LINE_SPACING_TWIPS,
  STRICT_MARGINS_TWIPS,
} from './types';

export interface DocxExportOptions {
  title: string;
  bodyHtml: string;
  headerHtml?: string;
  footerHtml?: string;
  pageNumberingEnabled?: boolean;
  fieldValues?: Record<string, string>;
}

export interface DocxImportResult {
  title: string;
  fileName: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  extractedPlaceholders: string[];
  stats: {
    paragraphCount: number;
    wordCount: number;
    tableCount: number;
  };
}

interface MediaRelationship {
  rId: string;
  fileName: string;
  data: Uint8Array;
  contentType: string;
}

interface InlineRunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  highlight?: string;
}

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Converts CSS RGB or Hex color string into 6-digit uppercase hex (without #).
 */
function parseColorToHex6(colorStr?: string | null): string | null {
  if (!colorStr) return null;
  const trimmed = colorStr.trim().toLowerCase();
  if (!trimmed || trimmed === 'transparent' || trimmed === 'inherit' || trimmed === 'initial') {
    return null;
  }
  if (trimmed.startsWith('#')) {
    const hex = trimmed.slice(1);
    if (hex.length === 3) {
      return hex
        .split('')
        .map((c) => c + c)
        .join('')
        .toUpperCase();
    }
    if (hex.length >= 6) {
      return hex.slice(0, 6).toUpperCase();
    }
  }
  const rgbMatch = trimmed.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (rgbMatch) {
    const r = Math.min(255, parseInt(rgbMatch[1], 10)).toString(16).padStart(2, '0');
    const g = Math.min(255, parseInt(rgbMatch[2], 10)).toString(16).padStart(2, '0');
    const b = Math.min(255, parseInt(rgbMatch[3], 10)).toString(16).padStart(2, '0');
    return `${r}${g}${b}`.toUpperCase();
  }
  return null;
}

/**
 * Normalizes a variable key by stripping braces, invisible BiDi/zero-width marks,
 * non-breaking spaces, and unifying whitespace/underscores.
 */
export function normalizePlaceholderKey(raw: string): string {
  return String(raw || '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/[{}\[\]\u200B-\u200F\u061C\uFEFF]/g, '')
    .replace(/[\u00A0\s]+/g, ' ')
    .trim();
}

export function lookupPlaceholderValue(
  fieldValues: Record<string, string>,
  rawKey: string
): string | undefined {
  if (!fieldValues || !rawKey) return undefined;
  const cleaned = normalizePlaceholderKey(rawKey);
  if (!cleaned) return undefined;

  const withUnderscores = cleaned.replace(/[\s.]+/g, '_');
  const withDots = cleaned.replace(/[\s_]+/g, '.');
  const withSpaces = cleaned.replace(/[_.]+/g, ' ');

  if (fieldValues[cleaned] !== undefined && fieldValues[cleaned].trim() !== '') {
    return fieldValues[cleaned];
  }
  if (
    fieldValues[withUnderscores] !== undefined &&
    fieldValues[withUnderscores].trim() !== ''
  ) {
    return fieldValues[withUnderscores];
  }
  if (fieldValues[withDots] !== undefined && fieldValues[withDots].trim() !== '') {
    return fieldValues[withDots];
  }
  if (fieldValues[withSpaces] !== undefined && fieldValues[withSpaces].trim() !== '') {
    return fieldValues[withSpaces];
  }

  const canonicalTarget = withUnderscores.toLowerCase();
  for (const [k, v] of Object.entries(fieldValues)) {
    if (v === undefined || v.trim() === '') continue;
    const canonicalK = normalizePlaceholderKey(k).replace(/[\s.]+/g, '_').toLowerCase();
    if (canonicalK === canonicalTarget) {
      return v;
    }
  }
  return undefined;
}

/**
 * Replaces {{placeholder}} tokens and <span class="smart-tag" data-var="..."> elements
 * in HTML with their actual values from fieldValues.
 */
export function mergePlaceholdersIntoHtml(
  html: string,
  fieldValues: Record<string, string> = {}
): string {
  if (!html) return '';

  let workingHtml = html;

  // 1. If DOMParser is available in browser, replace all .smart-tag / [data-var] spans,
  // single-node {{...}}, and cross-node split {{...}} cleanly in the DOM tree
  if (typeof window !== 'undefined' && typeof DOMParser !== 'undefined') {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(`<div>${workingHtml}</div>`, 'text/html');
      const container = doc.body.firstElementChild as HTMLElement | null;
      if (container) {
        // 1a. Replace all .smart-tag / .smart-placeholder / [data-var] elements
        const tagSpans = Array.from(
          container.querySelectorAll('.smart-tag, .smart-placeholder, [data-var]')
        ) as HTMLElement[];

        for (const span of tagSpans) {
          const rawVar =
            span.getAttribute('data-var') || (span.textContent || '');
          const val = lookupPlaceholderValue(fieldValues, rawVar);
          if (val !== undefined && val.trim() !== '') {
            span.replaceWith(doc.createTextNode(val.trim()));
          }
        }

        // 1b. Replace [...] and legacy {{...}} inside individual Text nodes
        const walker = doc.createTreeWalker(container, NodeFilter.SHOW_TEXT);
        const textNodes: Text[] = [];
        let curr: Node | null;
        while ((curr = walker.nextNode())) {
          textNodes.push(curr as Text);
        }
        for (const tNode of textNodes) {
          const txt = tNode.nodeValue || '';
          if (
            (txt.includes('[') && txt.includes(']')) ||
            (txt.includes('{{') && txt.includes('}}'))
          ) {
            const replaced = txt.replace(
              /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g,
              (full, bKey, cKey) => {
                const rk = bKey || cKey || '';
                const val = lookupPlaceholderValue(fieldValues, rk);
                return val !== undefined && val.trim() !== '' ? val.trim() : full;
              }
            );
            if (replaced !== txt) {
              tNode.nodeValue = replaced;
            }
          }
        }

        // 1c. Replace cross-node split [...] or {{...}} inside block elements
        const blocks = Array.from(
          container.querySelectorAll('p, li, td, th, div, h1, h2, h3, h4')
        ) as HTMLElement[];
        const leafBlocks = blocks.filter(
          (b) => !blocks.some((other) => other !== b && b.contains(other))
        );
        for (const block of leafBlocks) {
          const bText = block.textContent || '';
          const hasBrackets = bText.includes('[') && bText.includes(']');
          const hasBraces = bText.includes('{{') && bText.includes('}}');
          if (!hasBrackets && !hasBraces) continue;
          const bWalker = doc.createTreeWalker(block, NodeFilter.SHOW_TEXT);
          const segs: { node: Text; start: number; end: number }[] = [];
          let full = '';
          let n: Node | null;
          while ((n = bWalker.nextNode())) {
            const tn = n as Text;
            const len = (tn.nodeValue || '').length;
            if (len > 0) {
              segs.push({ node: tn, start: full.length, end: full.length + len });
              full += tn.nodeValue;
            }
          }
          const matches = Array.from(
            full.matchAll(/\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g)
          ).reverse();
          for (const m of matches) {
            if (m.index === undefined) continue;
            const rawK = m[1] || m[2] || '';
            const val = lookupPlaceholderValue(fieldValues, rawK);
            if (val === undefined || val.trim() === '') continue;
            const mStart = m.index;
            const mEnd = mStart + m[0].length;
            const sEntry = segs.find((s) => mStart >= s.start && mStart < s.end);
            const eEntry = segs.find((s) => mEnd > s.start && mEnd <= s.end);
            if (sEntry && eEntry) {
              if (sEntry.node === eEntry.node) {
                const v = sEntry.node.nodeValue || '';
                sEntry.node.nodeValue =
                  v.slice(0, mStart - sEntry.start) +
                  val.trim() +
                  v.slice(mEnd - sEntry.start);
              } else {
                const r = doc.createRange();
                r.setStart(sEntry.node, mStart - sEntry.start);
                r.setEnd(eEntry.node, mEnd - eEntry.start);
                r.deleteContents();
                r.insertNode(doc.createTextNode(val.trim()));
              }
            }
          }
        }

        workingHtml = container.innerHTML;
      }
    } catch {
      // Fallback to regex below
    }
  }

  // 2. Regex replacement for any <span ... data-var="KEY" ...>...</span> regardless of attribute order
  workingHtml = workingHtml.replace(
    /<span[^>]*data-var="([^"]+)"[^>]*>[\s\S]*?<\/span>/gi,
    (fullMatch, rawKey) => {
      const val = lookupPlaceholderValue(fieldValues, rawKey);
      return val !== undefined && val.trim() !== '' ? escapeXml(val.trim()) : fullMatch;
    }
  );

  // 3. Regex replacement for <span class="...smart-tag...">...</span>
  workingHtml = workingHtml.replace(
    /<span[^>]*class="[^"]*(?:smart-tag|smart-placeholder)[^"]*"[^>]*>(?:\[\s*([^\[\]<]+?)\s*\]|\{\{\s*([^}<]+?)\s*\}\})<\/span>/gi,
    (fullMatch, bKey, cKey) => {
      const val = lookupPlaceholderValue(fieldValues, bKey || cKey || '');
      return val !== undefined && val.trim() !== '' ? escapeXml(val.trim()) : fullMatch;
    }
  );

  // 4. Replace any remaining raw [key] or {{key}} occurrences in text
  return workingHtml.replace(
    /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{(?:&nbsp;|\s|<[^>]+>)*([^}<]+?)(?:&nbsp;|\s|<[^>]+>)*\}\}/gi,
    (fullMatch, bKey, cKey) => {
      const val = lookupPlaceholderValue(fieldValues, bKey || cKey || '');
      return val !== undefined && val.trim() !== '' ? escapeXml(val.trim()) : fullMatch;
    }
  );
}

/**
 * Splits mixed Arabic/Latin text into homogeneous BiDi segments.
 * Pure Latin word segments ([A-Za-z]+) have isRtl = false;
 * Arabic, numbers, and punctuation in Arabic context have isRtl = true.
 */
function splitTextIntoBidiRuns(text: string, defaultDir: 'rtl' | 'ltr'): { text: string; isRtl: boolean }[] {
  if (!text) return [];
  const normalized = text.replace(/\u00A0/g, '\u00A0');
  const tokens: { text: string; isRtl: boolean }[] = [];

  // Match contiguous Latin letters (plus numbers directly attached to Latin words) vs everything else
  const regex = /([A-Za-z]+(?:[0-9A-Za-z\-_./]*[A-Za-z0-9])?)|([^A-Za-z]+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(normalized)) !== null) {
    if (match[1]) {
      tokens.push({ text: match[1], isRtl: false });
    } else if (match[2]) {
      // Check if default paragraph direction is LTR and segment has no Arabic characters
      const hasArabic = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/.test(match[2]);
      tokens.push({
        text: match[2],
        isRtl: hasArabic ? true : defaultDir === 'rtl',
      });
    }
  }
  return tokens;
}

/**
 * Builds strict OOXML <w:rPr> in exact ECMA-376 element order,
 * enforcing Arial 13pt (w:sz=26, w:szCs=26) always and forever.
 */
function buildRunPropertiesXml(style: InlineRunStyle, isRtl: boolean): string {
  const parts: string[] = [];
  // 1. w:rFonts (Always Arial)
  parts.push(
    `<w:rFonts w:ascii="${STRICT_FONT_FAMILY}" w:hAnsi="${STRICT_FONT_FAMILY}" w:eastAsia="${STRICT_FONT_FAMILY}" w:cs="${STRICT_FONT_FAMILY}"/>`
  );
  // 2. w:b & w:bCs
  if (style.bold) {
    parts.push('<w:b/><w:bCs/>');
  }
  // 3. w:i & w:iCs
  if (style.italic) {
    parts.push('<w:i/><w:iCs/>');
  }
  // 4. w:color
  const hexColor = parseColorToHex6(style.color);
  if (hexColor && hexColor !== '000000') {
    parts.push(`<w:color w:val="${hexColor}"/>`);
  }
  // 5. w:sz & w:szCs (Always 26 = 13pt)
  parts.push(
    `<w:sz w:val="${STRICT_FONT_SIZE_HALF_PT}"/><w:szCs w:val="${STRICT_FONT_SIZE_HALF_PT}"/>`
  );
  // 6. w:u
  if (style.underline) {
    parts.push('<w:u w:val="single"/>');
  }
  // 7. w:shd (for background highlight color)
  const hexBg = parseColorToHex6(style.highlight);
  if (hexBg && hexBg !== 'FFFFFF') {
    parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${hexBg}"/>`);
  }
  // 8. w:rtl
  if (isRtl) {
    parts.push('<w:rtl/>');
  }
  // 9. w:lang
  parts.push('<w:lang w:val="ar-DZ" w:bidi="ar-DZ"/>');

  return `<w:rPr>${parts.join('')}</w:rPr>`;
}

/**
 * Builds strict OOXML <w:pPr> in exact ECMA-376 element order:
 * w:keepNext -> w:pageBreakBefore -> w:numPr -> w:pBdr -> w:shd -> w:bidi -> w:spacing -> w:ind -> w:jc -> w:rPr
 * Always enforces single line spacing (w:line="240") and 0 paragraph spacing (w:before="0" w:after="0").
 */
function buildParagraphPropertiesXml(options: {
  isRtl: boolean;
  align?: string;
  pageBreakBefore?: boolean;
  numId?: number;
  ilvl?: number;
  indentStartTwips?: number;
  indentEndTwips?: number;
}): string {
  const parts: string[] = [];

  if (options.pageBreakBefore) {
    parts.push('<w:pageBreakBefore/>');
  }

  if (options.numId !== undefined && options.ilvl !== undefined) {
    parts.push(
      `<w:numPr><w:ilvl w:val="${Math.min(8, Math.max(0, options.ilvl))}"/><w:numId w:val="${options.numId}"/></w:numPr>`
    );
  }

  if (options.isRtl) {
    parts.push('<w:bidi/>');
  }

  // Strict single line spacing (1.0 = 240) and zero before/after spacing
  parts.push(
    `<w:spacing w:before="0" w:after="0" w:line="${STRICT_LINE_SPACING_TWIPS}" w:lineRule="auto"/>`
  );

  if (options.indentStartTwips || options.indentEndTwips) {
    const rightVal = options.isRtl ? options.indentStartTwips || 0 : options.indentEndTwips || 0;
    const leftVal = options.isRtl ? options.indentEndTwips || 0 : options.indentStartTwips || 0;
    parts.push(`<w:ind w:right="${rightVal}" w:left="${leftVal}"/>`);
  }

  // Alignment mapping: in OpenXML with w:bidi, "both" = justify, "center" = center
  let jcVal = 'both'; // Default notarial contract alignment is justified (ضبط كلي) or right
  const rawAlign = (options.align || '').toLowerCase();
  if (rawAlign === 'center') jcVal = 'center';
  else if (rawAlign === 'justify' || rawAlign === 'both') jcVal = 'both';
  else if (rawAlign === 'left') jcVal = options.isRtl ? 'right' : 'left';
  else if (rawAlign === 'right') jcVal = options.isRtl ? 'left' : 'right';
  else jcVal = 'both';

  parts.push(`<w:jc w:val="${jcVal}"/>`);

  // Paragraph mark run properties (Arial 13pt)
  parts.push(buildRunPropertiesXml({}, options.isRtl));

  return `<w:pPr>${parts.join('')}</w:pPr>`;
}

/**
 * Converts a Base64 Data URL image into a MediaRelationship and returns the <w:drawing> XML.
 */
function buildImageDrawingXml(
  imgEl: HTMLImageElement,
  mediaList: MediaRelationship[],
  maxWidthTwips: number
): string {
  const src = imgEl.getAttribute('src') || '';
  const dataMatch = src.match(/^data:(image\/([a-zA-Z0-9+.-]+));base64,(.+)$/);
  if (!dataMatch) return '';

  const contentType = dataMatch[1];
  const ext = dataMatch[2] === 'jpeg' ? 'jpg' : dataMatch[2].replace(/[^a-z0-9]/gi, '') || 'png';
  const base64Data = dataMatch[3];

  try {
    const binaryStr = atob(base64Data);
    const bytes = new Uint8Array(binaryStr.length);
    for (let i = 0; i < binaryStr.length; i++) {
      bytes[i] = binaryStr.charCodeAt(i);
    }

    const idx = mediaList.length + 1;
    const rId = `rIdImg${idx}`;
    const fileName = `image${idx}.${ext}`;
    mediaList.push({ rId, fileName, data: bytes, contentType });

    // Calculate dimensions in EMUs (1 twip = 635 EMUs, 1 px = 9525 EMUs)
    const maxWidthEmu = maxWidthTwips * 635;
    const attrWidthPx = parseInt(imgEl.getAttribute('width') || '', 10) || imgEl.width || 220;
    const attrHeightPx = parseInt(imgEl.getAttribute('height') || '', 10) || imgEl.height || 140;

    let cx = attrWidthPx * 9525;
    let cy = attrHeightPx * 9525;
    if (cx > maxWidthEmu) {
      const ratio = maxWidthEmu / cx;
      cx = Math.round(maxWidthEmu);
      cy = Math.round(cy * ratio);
    }

    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${idx}" name="Picture ${idx}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${idx}" name="${fileName}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  } catch {
    return '';
  }
}

/**
 * Recursively serializes inline child nodes of a paragraph/list-item into <w:r> elements.
 */
function serializeInlineNodesToRunsXml(
  node: Node,
  inheritedStyle: InlineRunStyle,
  paragraphDir: 'rtl' | 'ltr',
  mediaList: MediaRelationship[],
  maxWidthTwips: number
): { runsXml: string; hasPageBreak: boolean } {
  let runsXml = '';
  let hasPageBreak = false;

  if (node.nodeType === Node.TEXT_NODE) {
    const rawText = node.nodeValue || '';
    if (!rawText) return { runsXml: '', hasPageBreak: false };

    // Handle tabs inside text
    const tabParts = rawText.split('\t');
    tabParts.forEach((part, idx) => {
      if (part.length > 0) {
        const bidiRuns = splitTextIntoBidiRuns(part, paragraphDir);
        for (const bRun of bidiRuns) {
          const rPr = buildRunPropertiesXml(inheritedStyle, bRun.isRtl);
          runsXml += `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(bRun.text)}</w:t></w:r>`;
        }
      }
      if (idx < tabParts.length - 1) {
        const rPr = buildRunPropertiesXml(inheritedStyle, paragraphDir === 'rtl');
        runsXml += `<w:r>${rPr}<w:tab/></w:r>`;
      }
    });

    return { runsXml, hasPageBreak: false };
  }

  if (node.nodeType !== Node.ELEMENT_NODE) {
    return { runsXml: '', hasPageBreak: false };
  }

  const el = node as HTMLElement;
  const tag = el.tagName.toUpperCase();

  if (el.classList.contains('page-break') || el.getAttribute('data-page-break') === 'true') {
    return { runsXml: '', hasPageBreak: true };
  }

  if (tag === 'BR') {
    const rPr = buildRunPropertiesXml(inheritedStyle, paragraphDir === 'rtl');
    // If <br> is the sole empty placeholder at the end of a paragraph, skip emitting an extra line break
    if (!el.previousSibling && !el.nextSibling) {
      return { runsXml: '', hasPageBreak: false };
    }
    return { runsXml: `<w:r>${rPr}<w:br/></w:r>`, hasPageBreak: false };
  }

  if (tag === 'IMG') {
    const drawingXml = buildImageDrawingXml(el as HTMLImageElement, mediaList, maxWidthTwips);
    return { runsXml: drawingXml, hasPageBreak: false };
  }

  // Compute merged inline style
  const nextStyle: InlineRunStyle = { ...inheritedStyle };
  if (tag === 'B' || tag === 'STRONG' || el.style?.fontWeight === 'bold' || Number(el.style?.fontWeight) >= 600) {
    nextStyle.bold = true;
  }
  if (tag === 'I' || tag === 'EM' || el.style?.fontStyle === 'italic') {
    nextStyle.italic = true;
  }
  if (tag === 'U' || el.style?.textDecoration?.includes('underline')) {
    nextStyle.underline = true;
  }
  const isSmartTag =
    el.classList.contains('smart-tag') ||
    el.classList.contains('smart-placeholder') ||
    el.classList.contains('smart-placeholder-filled');

  if (el.style?.color && !isSmartTag) {
    nextStyle.color = el.style.color;
  }
  if (el.style?.backgroundColor && !isSmartTag) {
    nextStyle.highlight = el.style.backgroundColor;
  }

  for (const child of Array.from(el.childNodes)) {
    const res = serializeInlineNodesToRunsXml(child, nextStyle, paragraphDir, mediaList, maxWidthTwips);
    runsXml += res.runsXml;
    if (res.hasPageBreak) hasPageBreak = true;
  }

  return { runsXml, hasPageBreak };
}

/**
 * Serializes an HTMLTableElement into a complete OpenXML <w:tbl> using the 2D Virtual Grid Matrix,
 * handling colSpan (w:gridSpan), rowSpan (w:vMerge), borders, shading, and nested tables.
 */
function serializeTableToXml(
  tableEl: HTMLTableElement,
  mediaList: MediaRelationship[],
  availableWidthTwips: number
): string {
  const grid = buildVirtualTableGrid(tableEl);
  if (grid.numRows === 0 || grid.numCols === 0) return '';

  const colWidthTwips = Math.floor(availableWidthTwips / grid.numCols);
  const gridColsXml = Array.from({ length: grid.numCols })
    .map(() => `<w:gridCol w:w="${colWidthTwips}"/>`)
    .join('');

  const tblPrXml = `<w:tblPr><w:bidiVisual/><w:tblW w:w="${availableWidthTwips}" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>`;

  let rowsXml = '';

  for (let r = 0; r < grid.numRows; r++) {
    const trEl = grid.rows[r];
    let trPrXml = '';
    const heightPx = parseInt(trEl?.style?.height || '', 10);
    if (heightPx > 0) {
      trPrXml = `<w:trPr><w:trHeight w:val="${Math.round(heightPx * 15)}" w:hRule="atLeast"/></w:trPr>`;
    }

    let cellsXml = '';
    let c = 0;
    while (c < grid.numCols) {
      const slot = grid.matrix[r][c];
      if (!slot) {
        // Fallback empty cell
        const emptyP = `<w:p>${buildParagraphPropertiesXml({ isRtl: true })}</w:p>`;
        cellsXml += `<w:tc><w:tcPr><w:tcW w:w="${colWidthTwips}" w:type="dxa"/></w:tcPr>${emptyP}</w:tc>`;
        c++;
        continue;
      }

      const cellWidthTwips = colWidthTwips * slot.colSpan;
      const tcPrParts: string[] = [`<w:tcW w:w="${cellWidthTwips}" w:type="dxa"/>`];

      if (slot.colSpan > 1) {
        tcPrParts.push(`<w:gridSpan w:val="${slot.colSpan}"/>`);
      }

      if (slot.rowSpan > 1) {
        if (slot.originRow === r) {
          tcPrParts.push('<w:vMerge w:val="restart"/>');
        } else {
          tcPrParts.push('<w:vMerge/>');
        }
      }

      // Check cell custom border or shading
      const cellEl = slot.cell;
      if (cellEl.style?.borderStyle === 'none' || cellEl.getAttribute('data-border') === 'none') {
        tcPrParts.push(
          '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'
        );
      }
      const bgHex = parseColorToHex6(cellEl.style?.backgroundColor);
      if (bgHex && bgHex !== 'FFFFFF') {
        tcPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bgHex}"/>`);
      }

      if (slot.originRow === r) {
        // Origin cell: serialize its block children
        let innerXml = serializeBlockContainerToXml(cellEl, mediaList, cellWidthTwips - 120);
        if (!innerXml.trim()) {
          innerXml = `<w:p>${buildParagraphPropertiesXml({ isRtl: true })}</w:p>`;
        } else if (innerXml.trim().endsWith('</w:tbl>')) {
          // OpenXML strict requirement: <w:tc> must always end with a <w:p>, never a <w:tbl>
          innerXml += `<w:p>${buildParagraphPropertiesXml({ isRtl: true })}</w:p>`;
        }
        cellsXml += `<w:tc><w:tcPr>${tcPrParts.join('')}</w:tcPr>${innerXml}</w:tc>`;
      } else {
        // Merged continuation cell: emit required empty paragraph
        const emptyP = `<w:p>${buildParagraphPropertiesXml({ isRtl: true })}</w:p>`;
        cellsXml += `<w:tc><w:tcPr>${tcPrParts.join('')}</w:tcPr>${emptyP}</w:tc>`;
      }

      c += slot.colSpan;
    }

    rowsXml += `<w:tr>${trPrXml}${cellsXml}</w:tr>`;
  }

  return `<w:tbl>${tblPrXml}<w:tblGrid>${gridColsXml}</w:tblGrid>${rowsXml}</w:tbl>`;
}

/**
 * Serializes block children (paragraphs, headings, lists, tables, page-breaks) into OpenXML body elements.
 */
function serializeBlockContainerToXml(
  container: HTMLElement,
  mediaList: MediaRelationship[],
  availableWidthTwips: number,
  listState: { nextPageBreak: boolean } = { nextPageBreak: false }
): string {
  let xml = '';
  const childNodes = Array.from(container.childNodes);

  // If container has only text/inline nodes and no block elements, wrap as a single paragraph
  const hasBlockChild = childNodes.some(
    (n) =>
      n.nodeType === Node.ELEMENT_NODE &&
      ['P', 'DIV', 'UL', 'OL', 'TABLE', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE'].includes(
        (n as HTMLElement).tagName.toUpperCase()
      )
  );

  if (!hasBlockChild) {
    const dir = (container.getAttribute('dir') || 'rtl').toLowerCase() === 'ltr' ? 'ltr' : 'rtl';
    const align = container.style?.textAlign || container.getAttribute('align') || 'justify';
    const { runsXml } = serializeInlineNodesToRunsXml(container, {}, dir, mediaList, availableWidthTwips);
    const pPr = buildParagraphPropertiesXml({
      isRtl: dir === 'rtl',
      align,
      pageBreakBefore: listState.nextPageBreak,
    });
    listState.nextPageBreak = false;
    return `<w:p>${pPr}${runsXml}</w:p>`;
  }

  for (const node of childNodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      const txt = (node.nodeValue || '').trim();
      if (txt) {
        const pPr = buildParagraphPropertiesXml({
          isRtl: true,
          align: 'justify',
          pageBreakBefore: listState.nextPageBreak,
        });
        listState.nextPageBreak = false;
        const { runsXml } = serializeInlineNodesToRunsXml(node, {}, 'rtl', mediaList, availableWidthTwips);
        xml += `<w:p>${pPr}${runsXml}</w:p>`;
      }
      continue;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const el = node as HTMLElement;
    const tag = el.tagName.toUpperCase();

    if (el.classList.contains('page-break') || el.getAttribute('data-page-break') === 'true') {
      listState.nextPageBreak = true;
      continue;
    }

    if (tag === 'TABLE') {
      xml += serializeTableToXml(el as HTMLTableElement, mediaList, availableWidthTwips);
      continue;
    }

    if (tag === 'UL' || tag === 'OL') {
      xml += serializeListToXml(el, 0, mediaList, availableWidthTwips, listState);
      continue;
    }

    // Unwrap visual-only clause-container during Word export so clauses merge 100% natively
    if (tag === 'DIV' && (el.classList.contains('clause-container') || el.hasAttribute('data-clause-id'))) {
      xml += serializeBlockContainerToXml(el, mediaList, availableWidthTwips, listState);
      continue;
    }

    if (['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE'].includes(tag)) {
      const dir = (el.getAttribute('dir') || el.style?.direction || 'rtl').toLowerCase() === 'ltr' ? 'ltr' : 'rtl';
      const align = el.style?.textAlign || el.getAttribute('align') || 'justify';
      const padRightPx = parseInt(el.style?.paddingRight || el.style?.marginRight || '0', 10) || 0;
      const padLeftPx = parseInt(el.style?.paddingLeft || el.style?.marginLeft || '0', 10) || 0;

      const baseStyle: InlineRunStyle = {};
      if (['H1', 'H2', 'H3', 'H4'].includes(tag) || el.style?.fontWeight === 'bold') {
        baseStyle.bold = true;
      }

      const { runsXml, hasPageBreak } = serializeInlineNodesToRunsXml(
        el,
        baseStyle,
        dir,
        mediaList,
        availableWidthTwips
      );

      const pPr = buildParagraphPropertiesXml({
        isRtl: dir === 'rtl',
        align,
        pageBreakBefore: listState.nextPageBreak,
        indentStartTwips: dir === 'rtl' ? padRightPx * 15 : padLeftPx * 15,
        indentEndTwips: dir === 'rtl' ? padLeftPx * 15 : padRightPx * 15,
      });
      listState.nextPageBreak = hasPageBreak;

      xml += `<w:p>${pPr}${runsXml}</w:p>`;
    }
  }

  return xml;
}

/**
 * Converts HTML <ol>/<ul> lists (including nested lists) into OpenXML <w:p> elements with <w:numPr>.
 * numId mapping in numbering.xml:
 * 1 = decimal (1., 2., 3.)
 * 2 = arabic-alpha (أ، ب، ج، د)
 * 3 = arabic-abjad (أبجد، هوز)
 * 4 = bullet (•)
 * 5 = dash (-)
 */
function serializeListToXml(
  listEl: HTMLElement,
  level: number,
  mediaList: MediaRelationship[],
  availableWidthTwips: number,
  listState: { nextPageBreak: boolean }
): string {
  let xml = '';
  const tag = listEl.tagName.toUpperCase();
  const listTypeAttr = listEl.getAttribute('data-list-type') || '';

  let numId = tag === 'UL' ? 4 : 1;
  if (listTypeAttr === 'arabic-alpha') numId = 2;
  else if (listTypeAttr === 'arabic-abjad') numId = 3;
  else if (listTypeAttr === 'dash') numId = 5;
  else if (listTypeAttr === 'bullet') numId = 4;

  for (const child of Array.from(listEl.children)) {
    const cTag = child.tagName.toUpperCase();
    if (cTag === 'UL' || cTag === 'OL') {
      xml += serializeListToXml(child as HTMLElement, level + 1, mediaList, availableWidthTwips, listState);
      continue;
    }
    if (cTag === 'LI') {
      const liEl = child as HTMLElement;
      const dir = (liEl.getAttribute('dir') || 'rtl').toLowerCase() === 'ltr' ? 'ltr' : 'rtl';
      const align = liEl.style?.textAlign || 'justify';

      // Separate nested lists inside <li> from inline text
      const cloneLi = liEl.cloneNode(true) as HTMLElement;
      const nestedLists = Array.from(cloneLi.querySelectorAll(':scope > ul, :scope > ol')) as HTMLElement[];
      nestedLists.forEach((nl) => nl.remove());

      const { runsXml, hasPageBreak } = serializeInlineNodesToRunsXml(
        cloneLi,
        {},
        dir,
        mediaList,
        availableWidthTwips
      );

      const pPr = buildParagraphPropertiesXml({
        isRtl: dir === 'rtl',
        align,
        pageBreakBefore: listState.nextPageBreak,
        numId,
        ilvl: level,
      });
      listState.nextPageBreak = hasPageBreak;
      xml += `<w:p>${pPr}${runsXml}</w:p>`;

      // Now serialize any nested lists that were inside this <li>
      const realNested = Array.from(liEl.querySelectorAll(':scope > ul, :scope > ol')) as HTMLElement[];
      for (const nl of realNested) {
        xml += serializeListToXml(nl, level + 1, mediaList, availableWidthTwips, listState);
      }
    }
  }

  return xml;
}

function buildNumberingXml(): string {
  const buildAbstractNum = (abstractNumId: number, numFmt: string, lvlTextPattern: string) => {
    let levelsXml = '';
    for (let ilvl = 0; ilvl <= 8; ilvl++) {
      const indentRight = (ilvl + 1) * 360;
      const lvlText = lvlTextPattern.replace('%L', `%${ilvl + 1}`);
      levelsXml += `<w:lvl w:ilvl="${ilvl}">
        <w:start w:val="1"/>
        <w:numFmt w:val="${numFmt}"/>
        <w:lvlText w:val="${lvlText}"/>
        <w:lvlJc w:val="right"/>
        <w:pPr>
          <w:bidi/>
          <w:spacing w:before="0" w:after="0" w:line="${STRICT_LINE_SPACING_TWIPS}" w:lineRule="auto"/>
          <w:ind w:right="${indentRight}" w:hanging="180"/>
        </w:pPr>
        <w:rPr>
          <w:rFonts w:ascii="${STRICT_FONT_FAMILY}" w:hAnsi="${STRICT_FONT_FAMILY}" w:eastAsia="${STRICT_FONT_FAMILY}" w:cs="${STRICT_FONT_FAMILY}"/>
          <w:sz w:val="${STRICT_FONT_SIZE_HALF_PT}"/>
          <w:szCs w:val="${STRICT_FONT_SIZE_HALF_PT}"/>
          <w:rtl/>
        </w:rPr>
      </w:lvl>`;
    }
    return `<w:abstractNum w:abstractNumId="${abstractNumId}"><w:multiLevelType w:val="hybridMultilevel"/>${levelsXml}</w:abstractNum>`;
  };

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${buildAbstractNum(1, 'decimal', '%L.')}
  ${buildAbstractNum(2, 'arabicAlpha', '%L-')}
  ${buildAbstractNum(3, 'arabicAbjad', '%L-')}
  ${buildAbstractNum(4, 'bullet', '•')}
  ${buildAbstractNum(5, 'bullet', '-')}
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>
  <w:num w:numId="4"><w:abstractNumId w:val="4"/></w:num>
  <w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num>
</w:numbering>`;
}

function buildStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="${STRICT_FONT_FAMILY}" w:hAnsi="${STRICT_FONT_FAMILY}" w:eastAsia="${STRICT_FONT_FAMILY}" w:cs="${STRICT_FONT_FAMILY}"/>
        <w:sz w:val="${STRICT_FONT_SIZE_HALF_PT}"/>
        <w:szCs w:val="${STRICT_FONT_SIZE_HALF_PT}"/>
        <w:rtl/>
        <w:lang w:val="ar-DZ" w:bidi="ar-DZ"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:bidi/>
        <w:spacing w:before="0" w:after="0" w:line="${STRICT_LINE_SPACING_TWIPS}" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:bidi/>
      <w:spacing w:before="0" w:after="0" w:line="${STRICT_LINE_SPACING_TWIPS}" w:lineRule="auto"/>
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${STRICT_FONT_FAMILY}" w:hAnsi="${STRICT_FONT_FAMILY}" w:eastAsia="${STRICT_FONT_FAMILY}" w:cs="${STRICT_FONT_FAMILY}"/>
      <w:sz w:val="${STRICT_FONT_SIZE_HALF_PT}"/>
      <w:szCs w:val="${STRICT_FONT_SIZE_HALF_PT}"/>
      <w:rtl/>
      <w:lang w:val="ar-DZ" w:bidi="ar-DZ"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}

function buildSettingsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:view w:val="print"/>
  <w:zoom w:percent="100"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>`;
}

/**
 * Generates a complete, standard-compliant Microsoft Word (.docx) Blob
 * with strict Arial 13pt, 1.0 single line spacing, 0 paragraph spacing,
 * and fixed Notarial margins (Top: 1cm, Right: 7cm, Bottom: 6cm, Left: 2cm).
 */
export async function generateNotaryDocxBlob(options: DocxExportOptions): Promise<Blob> {
  const zip = new JSZip();
  const parser = new DOMParser();

  const mergedBodyHtml = mergePlaceholdersIntoHtml(options.bodyHtml, options.fieldValues);
  const mergedHeaderHtml = mergePlaceholdersIntoHtml(options.headerHtml || '', options.fieldValues);
  const mergedFooterHtml = mergePlaceholdersIntoHtml(options.footerHtml || '', options.fieldValues);

  const bodyDoc = parser.parseFromString(`<div>${mergedBodyHtml}</div>`, 'text/html');
  const bodyRoot = bodyDoc.body.firstElementChild as HTMLElement;

  const mediaList: MediaRelationship[] = [];
  let bodyElementsXml = serializeBlockContainerToXml(bodyRoot, mediaList, NET_CONTENT_WIDTH_TWIPS);

  // OpenXML requires <w:body> to have at least one <w:p> and not end directly with <w:tbl>
  if (!bodyElementsXml.trim() || bodyElementsXml.trim().endsWith('</w:tbl>')) {
    bodyElementsXml += `<w:p>${buildParagraphPropertiesXml({ isRtl: true })}</w:p>`;
  }

  const hasHeader = mergedHeaderHtml.replace(/<[^>]+>/g, '').trim().length > 0;
  const hasFooter =
    mergedFooterHtml.replace(/<[^>]+>/g, '').trim().length > 0 || Boolean(options.pageNumberingEnabled);

  // Section Properties with strict Notarial Margins (Right: 7cm=3969, Left: 2cm=1134, Top: 1cm=567, Bottom: 6cm=3402)
  const sectRefs: string[] = [];
  if (hasHeader) {
    sectRefs.push(
      '<w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdHeader1"/>'
    );
  }
  if (hasFooter) {
    sectRefs.push(
      '<w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdFooter1"/>'
    );
  }

  const sectPrXml = `<w:sectPr>
    ${sectRefs.join('')}
    <w:pgSz w:w="${A4_SIZE_TWIPS.width}" w:h="${A4_SIZE_TWIPS.height}"/>
    <w:pgMar w:top="${STRICT_MARGINS_TWIPS.top}" w:right="${STRICT_MARGINS_TWIPS.right}" w:bottom="${STRICT_MARGINS_TWIPS.bottom}" w:left="${STRICT_MARGINS_TWIPS.left}" w:header="${STRICT_MARGINS_TWIPS.header}" w:footer="${STRICT_MARGINS_TWIPS.footer}" w:gutter="0"/>
    <w:bidi/>
  </w:sectPr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
  <w:body>
    ${bodyElementsXml}
    ${sectPrXml}
  </w:body>
</w:document>`;

  // Header & Footer XML
  if (hasHeader) {
    const hDoc = parser.parseFromString(`<div>${mergedHeaderHtml}</div>`, 'text/html');
    const hRoot = hDoc.body.firstElementChild as HTMLElement;
    const hInnerXml =
      serializeBlockContainerToXml(hRoot, [], NET_CONTENT_WIDTH_TWIPS) ||
      `<w:p>${buildParagraphPropertiesXml({ isRtl: true, align: 'center' })}</w:p>`;
    zip.file(
      'word/header1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${hInnerXml}
</w:hdr>`
    );
  }

  if (hasFooter) {
    const fDoc = parser.parseFromString(`<div>${mergedFooterHtml}</div>`, 'text/html');
    const fRoot = fDoc.body.firstElementChild as HTMLElement;
    let fInnerXml = '';
    if (mergedFooterHtml.replace(/<[^>]+>/g, '').trim().length > 0) {
      fInnerXml += serializeBlockContainerToXml(fRoot, [], NET_CONTENT_WIDTH_TWIPS);
    }
    if (options.pageNumberingEnabled) {
      const rPr = buildRunPropertiesXml({}, true);
      fInnerXml += `<w:p>${buildParagraphPropertiesXml({ isRtl: true, align: 'center' })}
        <w:r>${rPr}<w:t xml:space="preserve">صفحة </w:t></w:r>
        <w:fldSimple w:instr="PAGE"><w:r>${rPr}<w:t>1</w:t></w:r></w:fldSimple>
        <w:r>${rPr}<w:t xml:space="preserve"> من </w:t></w:r>
        <w:fldSimple w:instr="NUMPAGES"><w:r>${rPr}<w:t>1</w:t></w:r></w:fldSimple>
      </w:p>`;
    }
    zip.file(
      'word/footer1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${fInnerXml}
</w:ftr>`
    );
  }

  // Write Media files
  for (const media of mediaList) {
    zip.file(`word/media/${media.fileName}`, media.data);
  }

  // Build word/_rels/document.xml.rels
  const relEntries: string[] = [
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
  ];
  if (hasHeader) {
    relEntries.push(
      '<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
    );
  }
  if (hasFooter) {
    relEntries.push(
      '<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
    );
  }
  for (const media of mediaList) {
    relEntries.push(
      `<Relationship Id="${media.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${media.fileName}"/>`
    );
  }

  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relEntries.join('\n  ')}
</Relationships>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  ${hasHeader ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}
  ${hasFooter ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}
</Types>`
  );

  zip.file('word/document.xml', documentXml);
  zip.file('word/styles.xml', buildStylesXml());
  zip.file('word/numbering.xml', buildNumberingXml());
  zip.file('word/settings.xml', buildSettingsXml());

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

/**
 * Triggers instant download (or Web Share if requested) of the generated .docx file.
 */
export async function downloadNotaryDocx(options: DocxExportOptions): Promise<Uint8Array> {
  const blob = await generateNotaryDocxBlob(options);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const safeName = (options.title || 'عقد_توثيقي')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_');
  const fileName = safeName.endsWith('.docx') ? safeName : `${safeName}.docx`;

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  return bytes;
}

/**
 * Parses an imported .docx file (word/document.xml, word/header*.xml, word/footer*.xml, word/media/*)
 * into clean HTML preserving bold, italic, underline, colors, alignments, lists, and complex tables
 * (rowSpan/colSpan/borders/shading) while strictly normalizing font to Arial 13pt and 1.0 line spacing.
 */
export async function importNotaryDocxFile(file: File): Promise<DocxImportResult> {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const parser = new DOMParser();

  // 1. Load relationships & media images as Base64 data URLs
  const mediaDataUrlMap: Record<string, string> = {};
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    const relsXml = await relsFile.async('string');
    const relsDoc = parser.parseFromString(relsXml, 'application/xml');
    const relNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));
    for (const rel of relNodes) {
      const rId = rel.getAttribute('Id') || '';
      const target = rel.getAttribute('Target') || '';
      if (rId && target.includes('media/')) {
        const cleanPath = target.startsWith('/') ? target.slice(1) : `word/${target.replace(/^\.\.\//, '')}`;
        const imgFile = zip.file(cleanPath);
        if (imgFile) {
          const base64 = await imgFile.async('base64');
          const ext = cleanPath.split('.').pop()?.toLowerCase() || 'png';
          const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
          mediaDataUrlMap[rId] = `data:${mime};base64,${base64}`;
        }
      }
    }
  }

  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new Error('ملف الوورد غير صالح أو لا يحتوي على word/document.xml');
  }

  const docXmlStr = await docFile.async('string');
  const xmlDoc = parser.parseFromString(docXmlStr, 'application/xml');
  const bodyNode = xmlDoc.getElementsByTagName('w:body')[0] || xmlDoc.getElementsByTagName('body')[0];

  const bodyHtml = bodyNode ? convertXmlContainerToHtml(bodyNode, mediaDataUrlMap) : '';

  // Optional header1.xml & footer1.xml
  let headerHtml = '';
  const headerFile = zip.file('word/header1.xml');
  if (headerFile) {
    const hXml = await headerFile.async('string');
    const hDoc = parser.parseFromString(hXml, 'application/xml');
    if (hDoc.documentElement) {
      headerHtml = convertXmlContainerToHtml(hDoc.documentElement, mediaDataUrlMap);
    }
  }

  let footerHtml = '';
  const footerFile = zip.file('word/footer1.xml');
  if (footerFile) {
    const fXml = await footerFile.async('string');
    const fDoc = parser.parseFromString(fXml, 'application/xml');
    if (fDoc.documentElement) {
      footerHtml = convertXmlContainerToHtml(fDoc.documentElement, mediaDataUrlMap);
    }
  }

  // Normalize bracket placeholders like [البائع] (and any legacy {{البائع}}) into [البائع] / .smart-tag
  const normalizeBracketPlaceholdersInHtml = (htmlStr: string): string => {
    if (!htmlStr) return '';
    // Only replace [text] or {{text}} outside HTML tags when text is 1..45 chars and doesn't contain HTML/newlines
    return htmlStr.replace(/(>[^<]*)|(<[^>]+>)/g, (segment) => {
      if (segment.startsWith('<')) return segment;
      return segment.replace(
        /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}<>]{1,45}?)\s*\}\}/g,
        (fullMatch, bInner, cInner) => {
          const inner = bInner || cInner || '';
          const cleanKey = normalizePlaceholderKey(inner).replace(/\s+/g, '_');
          if (!cleanKey || /^\d+$/.test(cleanKey)) return fullMatch;
          return `<span class="smart-tag" data-var="${escapeXml(cleanKey)}">[${escapeXml(cleanKey)}]</span>`;
        }
      );
    });
  };

  const normalizedBodyHtml = normalizeBracketPlaceholdersInHtml(bodyHtml);
  const normalizedHeaderHtml = normalizeBracketPlaceholdersInHtml(headerHtml);
  const normalizedFooterHtml = normalizeBracketPlaceholdersInHtml(footerHtml);

  // Extract [...] (and legacy {{...}}) and data-var placeholders
  const combinedHtml = `${normalizedHeaderHtml} ${normalizedBodyHtml} ${normalizedFooterHtml}`;
  const combinedText = combinedHtml.replace(/<[^>]+>/g, ' ');
  const placeholderRegex = /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g;
  const placeholders = new Set<string>();
  let m: RegExpExecArray | null;
  while ((m = placeholderRegex.exec(combinedText)) !== null) {
    const raw = m[1] || m[2] || '';
    const clean = normalizePlaceholderKey(raw).replace(/\s+/g, '_');
    if (clean && !/^\d+$/.test(clean)) placeholders.add(clean);
  }

  const paragraphCount = (normalizedBodyHtml.match(/<p\b/gi) || []).length || 1;
  const tableCount = (normalizedBodyHtml.match(/<table\b/gi) || []).length;
  const plainWords = combinedText
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(Boolean);

  const title = file.name.replace(/\.docx$/i, '');
  return {
    title,
    fileName: file.name,
    bodyHtml:
      normalizedBodyHtml ||
      `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;"><br></p>`,
    headerHtml: normalizedHeaderHtml,
    footerHtml: normalizedFooterHtml,
    extractedPlaceholders: Array.from(placeholders),
    stats: {
      paragraphCount,
      wordCount: plainWords.length,
      tableCount,
    },
  };
}

function getLocalTagName(node: Element): string {
  return (node.localName || node.tagName.replace(/^.*:/, '')).toLowerCase();
}

function findDirectChild(parent: Element, localName: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (getLocalTagName(child) === localName) return child;
  }
  return null;
}

function convertXmlContainerToHtml(
  container: Element,
  mediaDataUrlMap: Record<string, string>
): string {
  let html = '';

  for (const child of Array.from(container.children)) {
    const tag = getLocalTagName(child);
    if (tag === 'p') {
      html += convertXmlParagraphToHtml(child, mediaDataUrlMap);
    } else if (tag === 'tbl') {
      html += convertXmlTableToHtml(child, mediaDataUrlMap);
    } else if (tag === 'sdt') {
      // Structured Document Tag: unwrap w:sdtContent
      const sdtContent = findDirectChild(child, 'sdtcontent');
      if (sdtContent) {
        html += convertXmlContainerToHtml(sdtContent, mediaDataUrlMap);
      }
    }
  }

  return html;
}

function convertXmlParagraphToHtml(
  pNode: Element,
  mediaDataUrlMap: Record<string, string>
): string {
  const pPr = findDirectChild(pNode, 'ppr');
  let align = 'justify';
  let isRtl = true;
  let hasPageBreakBefore = false;

  if (pPr) {
    const bidiEl = findDirectChild(pPr, 'bidi');
    if (bidiEl && (bidiEl.getAttribute('w:val') === '0' || bidiEl.getAttribute('val') === '0')) {
      isRtl = false;
    }
    const jcEl = findDirectChild(pPr, 'jc');
    if (jcEl) {
      const val = (jcEl.getAttribute('w:val') || jcEl.getAttribute('val') || '').toLowerCase();
      if (val === 'center') align = 'center';
      else if (val === 'both' || val === 'distribute') align = 'justify';
      else if (val === 'left') align = isRtl ? 'right' : 'left';
      else if (val === 'right') align = isRtl ? 'left' : 'right';
    }
    if (findDirectChild(pPr, 'pagebreakbefore')) {
      hasPageBreakBefore = true;
    }
  }

  let innerHtml = '';
  let hasInlinePageBreak = false;

  const processParagraphChild = (el: Element) => {
    const tag = getLocalTagName(el);
    if (tag === 'r') {
      const res = convertXmlRunToHtml(el, mediaDataUrlMap);
      innerHtml += res.html;
      if (res.hasPageBreak) hasInlinePageBreak = true;
    } else if (tag === 'hyperlink' || tag === 'fldsimple' || tag === 'smarttag') {
      for (const sub of Array.from(el.children)) {
        processParagraphChild(sub);
      }
    } else if (tag === 'sdt') {
      const sdtContent = findDirectChild(el, 'sdtcontent');
      if (sdtContent) {
        for (const sub of Array.from(sdtContent.children)) {
          processParagraphChild(sub);
        }
      }
    }
  };

  for (const child of Array.from(pNode.children)) {
    processParagraphChild(child);
  }

  let prefixBreak = '';
  if (hasPageBreakBefore) {
    prefixBreak =
      '<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div>';
  }

  // If paragraph only had a page break and no text, emit only the page break without an extra empty line
  if (hasInlinePageBreak && innerHtml.trim() === '') {
    return `${prefixBreak}<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div>`;
  }

  const suffixBreak = hasInlinePageBreak
    ? '<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div>'
    : '';

  const content = innerHtml.trim() ? innerHtml : '<br>';
  return `${prefixBreak}<p dir="${isRtl ? 'rtl' : 'ltr'}" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:${align};">${content}</p>${suffixBreak}`;
}

function convertXmlRunToHtml(
  rNode: Element,
  mediaDataUrlMap: Record<string, string>
): { html: string; hasPageBreak: boolean } {
  const rPr = findDirectChild(rNode, 'rpr');
  let bold = false;
  let italic = false;
  let underline = false;
  let color = '';
  let bg = '';

  if (rPr) {
    const bEl = findDirectChild(rPr, 'b') || findDirectChild(rPr, 'bcs');
    if (bEl && bEl.getAttribute('w:val') !== '0' && bEl.getAttribute('val') !== '0') {
      bold = true;
    }
    const iEl = findDirectChild(rPr, 'i') || findDirectChild(rPr, 'ics');
    if (iEl && iEl.getAttribute('w:val') !== '0' && iEl.getAttribute('val') !== '0') {
      italic = true;
    }
    const uEl = findDirectChild(rPr, 'u');
    if (uEl && uEl.getAttribute('w:val') !== 'none' && uEl.getAttribute('val') !== 'none') {
      underline = true;
    }
    const cEl = findDirectChild(rPr, 'color');
    const cVal = cEl?.getAttribute('w:val') || cEl?.getAttribute('val') || '';
    if (cVal && cVal !== 'auto' && cVal !== '000000') {
      color = `#${cVal}`;
    }
    const shdEl = findDirectChild(rPr, 'shd');
    const fillVal = shdEl?.getAttribute('w:fill') || shdEl?.getAttribute('fill') || '';
    if (fillVal && fillVal !== 'auto' && fillVal !== 'FFFFFF') {
      bg = `#${fillVal}`;
    }
  }

  let runContent = '';
  let hasPageBreak = false;

  for (const child of Array.from(rNode.children)) {
    const tag = getLocalTagName(child);
    if (tag === 't') {
      runContent += escapeXml(child.textContent || '');
    } else if (tag === 'tab') {
      runContent += '&nbsp;&nbsp;&nbsp;&nbsp;';
    } else if (tag === 'br') {
      const brType = child.getAttribute('w:type') || child.getAttribute('type') || '';
      if (brType === 'page') {
        hasPageBreak = true;
      } else {
        runContent += '<br>';
      }
    } else if (tag === 'drawing' || tag === 'pict') {
      // Check for embedded blip image or textbox content (w:txbxContent)
      const blips = Array.from(child.getElementsByTagName('*')).filter(
        (el) => getLocalTagName(el) === 'blip'
      );
      for (const blip of blips) {
        const embedId = blip.getAttribute('r:embed') || blip.getAttribute('embed') || '';
        if (embedId && mediaDataUrlMap[embedId]) {
          runContent += `<img src="${mediaDataUrlMap[embedId]}" alt="صورة مضمنة" style="max-width:100%;height:auto;" />`;
        }
      }
      const txbxList = Array.from(child.getElementsByTagName('*')).filter(
        (el) => getLocalTagName(el) === 'txbxcontent'
      );
      for (const txbx of txbxList) {
        runContent += convertXmlContainerToHtml(txbx, mediaDataUrlMap);
      }
    }
  }

  if (!runContent) return { html: '', hasPageBreak };

  const styles: string[] = [
    `font-family:${STRICT_FONT_FAMILY}`,
    `font-size:${STRICT_FONT_SIZE_PT}pt`,
    'line-height:1',
  ];
  if (bold) styles.push('font-weight:bold');
  if (italic) styles.push('font-style:italic');
  if (underline) styles.push('text-decoration:underline');
  if (color) styles.push(`color:${color}`);
  if (bg) styles.push(`background-color:${bg}`);

  if (bold || italic || underline || color || bg) {
    return {
      html: `<span style="${styles.join(';')}">${runContent}</span>`,
      hasPageBreak,
    };
  }

  return { html: runContent, hasPageBreak };
}

/**
 * Reconstructs an HTML <table> from OpenXML <w:tbl>, accurately computing
 * colSpan from <w:gridSpan> and rowSpan from <w:vMerge> across rows.
 */
function convertXmlTableToHtml(
  tblNode: Element,
  mediaDataUrlMap: Record<string, string>
): string {
  const trNodes = Array.from(tblNode.children).filter((c) => getLocalTagName(c) === 'tr');
  if (trNodes.length === 0) return '';

  interface ParsedCellInfo {
    contentHtml: string;
    colSpan: number;
    vMergeState: 'none' | 'restart' | 'continue';
    rowSpan: number;
    bgHex: string | null;
    borderHidden: boolean;
  }

  const parsedRows: ParsedCellInfo[][] = [];

  for (const tr of trNodes) {
    const tcNodes = Array.from(tr.children).filter((c) => getLocalTagName(c) === 'tc');
    const rowCells: ParsedCellInfo[] = [];

    for (const tc of tcNodes) {
      const tcPr = findDirectChild(tc, 'tcpr');
      let colSpan = 1;
      let vMergeState: 'none' | 'restart' | 'continue' = 'none';
      let bgHex: string | null = null;
      let borderHidden = false;

      if (tcPr) {
        const gridSpanEl = findDirectChild(tcPr, 'gridspan');
        if (gridSpanEl) {
          colSpan = Math.max(
            1,
            parseInt(gridSpanEl.getAttribute('w:val') || gridSpanEl.getAttribute('val') || '1', 10)
          );
        }
        const vMergeEl = findDirectChild(tcPr, 'vmerge');
        if (vMergeEl) {
          const val = vMergeEl.getAttribute('w:val') || vMergeEl.getAttribute('val');
          vMergeState = val === 'restart' ? 'restart' : 'continue';
        }
        const shdEl = findDirectChild(tcPr, 'shd');
        const fill = shdEl?.getAttribute('w:fill') || shdEl?.getAttribute('fill') || '';
        if (fill && fill !== 'auto' && fill !== 'FFFFFF') {
          bgHex = `#${fill}`;
        }
        const tcBorders = findDirectChild(tcPr, 'tcborders');
        if (tcBorders) {
          const topB = findDirectChild(tcBorders, 'top');
          const topVal = topB?.getAttribute('w:val') || topB?.getAttribute('val') || '';
          if (topVal === 'nil' || topVal === 'none') {
            borderHidden = true;
          }
        }
      }

      const contentHtml = convertXmlContainerToHtml(tc, mediaDataUrlMap);
      rowCells.push({
        contentHtml:
          contentHtml ||
          `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;"><br></p>`,
        colSpan,
        vMergeState,
        rowSpan: 1,
        bgHex,
        borderHidden,
      });
    }

    parsedRows.push(rowCells);
  }

  // Resolve vertical merges (vMerge restart -> continue) by tracking column grid positions
  const activeVSpans: { originCell: ParsedCellInfo; colIndex: number }[] = [];

  for (let r = 0; r < parsedRows.length; r++) {
    let colCursor = 0;
    const row = parsedRows[r];
    for (const cell of row) {
      if (cell.vMergeState === 'restart') {
        const existingIdx = activeVSpans.findIndex((x) => x.colIndex === colCursor);
        if (existingIdx >= 0) activeVSpans.splice(existingIdx, 1);
        activeVSpans.push({ originCell: cell, colIndex: colCursor });
      } else if (cell.vMergeState === 'continue') {
        const active = activeVSpans.find((x) => x.colIndex === colCursor);
        if (active) {
          active.originCell.rowSpan += 1;
        }
      } else {
        const existingIdx = activeVSpans.findIndex((x) => x.colIndex === colCursor);
        if (existingIdx >= 0) activeVSpans.splice(existingIdx, 1);
      }
      colCursor += cell.colSpan;
    }
  }

  let tableHtml = `<table dir="rtl" style="width:100%;max-width:120mm;border-collapse:collapse;table-layout:fixed;margin:0;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;line-height:1;"><tbody>`;

  for (const row of parsedRows) {
    tableHtml += '<tr>';
    for (const cell of row) {
      if (cell.vMergeState === 'continue') continue;
      const attrs: string[] = [];
      if (cell.colSpan > 1) attrs.push(`colspan="${cell.colSpan}"`);
      if (cell.rowSpan > 1) attrs.push(`rowspan="${cell.rowSpan}"`);

      const borderCss = cell.borderHidden ? 'border:none;' : 'border:1px solid #000000;';
      const bgCss = cell.bgHex ? `background-color:${cell.bgHex};` : '';

      tableHtml += `<td ${attrs.join(' ')} style="${borderCss}${bgCss}padding:2px 4px;vertical-align:top;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;line-height:1;">${cell.contentHtml}</td>`;
    }
    tableHtml += '</tr>';
  }

  tableHtml += '</tbody></table>';
  return tableHtml;
}
