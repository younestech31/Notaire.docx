import type {
  ClauseVariableGroup,
  ListNumberingStyle,
  SerializedSelectionPath,
} from './types';
import {
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
} from './types';
import { lookupPlaceholderValue, normalizePlaceholderKey } from './docx-engine';

/**
 * Computes the DOM index path of a node relative to rootEl.
 */
export function getNodePath(rootEl: HTMLElement, targetNode: Node): number[] {
  const path: number[] = [];
  let current: Node | null = targetNode;
  while (current && current !== rootEl) {
    const parent: Node | null = current.parentNode;
    if (!parent) break;
    const index = Array.prototype.indexOf.call(parent.childNodes, current);
    path.unshift(index);
    current = parent;
  }
  return path;
}

export function resolveNodeFromPath(rootEl: HTMLElement, path: number[]): Node | null {
  let current: Node = rootEl;
  for (const idx of path) {
    if (!current.childNodes || idx < 0 || idx >= current.childNodes.length) {
      return current;
    }
    current = current.childNodes[idx];
  }
  return current;
}

export function serializeCurrentSelection(
  zones: {
    body: HTMLElement | null;
    header: HTMLElement | null;
    footer: HTMLElement | null;
  }
): SerializedSelectionPath | null {
  if (typeof window === 'undefined') return null;
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);

  for (const [zoneName, el] of Object.entries(zones) as [
    'body' | 'header' | 'footer',
    HTMLElement | null
  ][]) {
    if (el && el.contains(range.commonAncestorContainer)) {
      return {
        zone: zoneName,
        startPath: getNodePath(el, range.startContainer),
        startOffset: range.startOffset,
        endPath: getNodePath(el, range.endContainer),
        endOffset: range.endOffset,
      };
    }
  }
  return null;
}

export function restoreSerializedSelection(
  zones: {
    body: HTMLElement | null;
    header: HTMLElement | null;
    footer: HTMLElement | null;
  },
  saved: SerializedSelectionPath | null
): void {
  if (!saved || typeof window === 'undefined') return;
  const rootEl = zones[saved.zone];
  if (!rootEl) return;

  try {
    const startNode = resolveNodeFromPath(rootEl, saved.startPath);
    const endNode = resolveNodeFromPath(rootEl, saved.endPath);
    if (!startNode || !endNode) return;

    const range = document.createRange();
    const maxStart =
      startNode.nodeType === Node.TEXT_NODE
        ? (startNode.textContent || '').length
        : startNode.childNodes.length;
    const maxEnd =
      endNode.nodeType === Node.TEXT_NODE
        ? (endNode.textContent || '').length
        : endNode.childNodes.length;

    range.setStart(startNode, Math.min(saved.startOffset, maxStart));
    range.setEnd(endNode, Math.min(saved.endOffset, maxEnd));

    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(range);
    }
  } catch {
    // Ignore invalid DOM offsets
  }
}

/**
 * Returns all block elements (p, li, div, h1..h6) inside rootEl that intersect with the given Range.
 */
export function getIntersectingBlockElements(
  rootEl: HTMLElement,
  range: Range
): HTMLElement[] {
  const candidates = Array.from(
    rootEl.querySelectorAll('p, li, h1, h2, h3, h4, h5, h6')
  ) as HTMLElement[];

  const matched = candidates.filter((block) => {
    try {
      return range.intersectsNode(block);
    } catch {
      return false;
    }
  });

  if (matched.length > 0) return matched;

  let curr: Node | null = range.startContainer;
  while (curr && curr !== rootEl) {
    if (
      curr.nodeType === Node.ELEMENT_NODE &&
      ['P', 'LI', 'DIV', 'TD', 'TH'].includes((curr as HTMLElement).tagName)
    ) {
      return [curr as HTMLElement];
    }
    curr = curr.parentNode;
  }
  return [];
}

/**
 * Scans text nodes (and cross-node split runs) inside rootEl and wraps any raw [variable] (or legacy {{variable}}) tokens into
 * <span class="smart-tag" data-var="variable">[variable]</span>
 */
export function decorateSmartTagsInDOM(rootEl: HTMLElement): void {
  // 1. Strip any legacy contenteditable="false" on existing smart tags and normalize {{...}} display to [...]
  const existingTags = Array.from(
    rootEl.querySelectorAll('.smart-tag, .smart-placeholder')
  ) as HTMLElement[];
  for (const tagEl of existingTags) {
    if (tagEl.hasAttribute('contenteditable')) {
      tagEl.removeAttribute('contenteditable');
    }
    const rawKey =
      tagEl.getAttribute('data-var') ||
      (tagEl.textContent || '').replace(/[\[\]{}]/g, '');
    const cleanKey = normalizePlaceholderKey(rawKey).replace(/\s+/g, '_');
    if (cleanKey) {
      tagEl.setAttribute('data-var', cleanKey);
      if (tagEl.textContent !== `[${cleanKey}]`) {
        tagEl.textContent = `[${cleanKey}]`;
      }
    }
  }

  // 2. Single-text-node [...] (and legacy {{...}}) wrapping
  const walker = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT);
  const nodesToProcess: Text[] = [];
  let current: Node | null;

  while ((current = walker.nextNode())) {
    const tNode = current as Text;
    const parentEl = tNode.parentElement;
    if (
      parentEl &&
      (parentEl.classList.contains('smart-tag') ||
        parentEl.classList.contains('smart-placeholder') ||
        parentEl.closest('.smart-tag, .smart-placeholder'))
    ) {
      continue;
    }
    if (
      tNode.nodeValue &&
      (/\[\s*[^\[\]<>]{1,45}?\s*\]/.test(tNode.nodeValue) ||
        /\{\{\s*[^}]+?\s*\}\}/.test(tNode.nodeValue))
    ) {
      nodesToProcess.push(tNode);
    }
  }

  for (const tNode of nodesToProcess) {
    const text = tNode.nodeValue || '';
    const regex = /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g;
    let lastIndex = 0;
    let match: RegExpExecArray | null;
    const frag = document.createDocumentFragment();

    while ((match = regex.exec(text)) !== null) {
      if (match.index > lastIndex) {
        frag.appendChild(document.createTextNode(text.slice(lastIndex, match.index)));
      }
      const rawCaptured = match[1] || match[2] || '';
      const varName = normalizePlaceholderKey(rawCaptured).replace(/\s+/g, '_');
      if (varName && !/^\d+$/.test(varName)) {
        const span = document.createElement('span');
        span.className = 'smart-tag';
        span.setAttribute('data-var', varName);
        span.textContent = `[${varName}]`;
        frag.appendChild(span);
      } else {
        frag.appendChild(document.createTextNode(match[0]));
      }
      lastIndex = regex.lastIndex;
    }

    if (lastIndex < text.length) {
      frag.appendChild(document.createTextNode(text.slice(lastIndex)));
    }

    tNode.parentNode?.replaceChild(frag, tNode);
  }

  // 3. Cross-node split [...] or {{...}} wrapping inside leaf block elements
  const blocks = Array.from(
    rootEl.querySelectorAll('p, li, td, th, h1, h2, h3, h4')
  ) as HTMLElement[];
  const leafBlocks = blocks.filter(
    (b) => !blocks.some((other) => other !== b && b.contains(other))
  );

  for (const block of leafBlocks) {
    const bWalker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    const segs: { node: Text; start: number; end: number }[] = [];
    let full = '';
    let n: Node | null;
    while ((n = bWalker.nextNode())) {
      const tn = n as Text;
      if (tn.parentElement?.closest('.smart-tag, .smart-placeholder')) {
        continue;
      }
      const len = (tn.nodeValue || '').length;
      if (len > 0) {
        segs.push({ node: tn, start: full.length, end: full.length + len });
        full += tn.nodeValue;
      }
    }
    const hasSplitBrackets = full.includes('[') && full.includes(']');
    const hasSplitBraces = full.includes('{{') && full.includes('}}');
    if (!hasSplitBrackets && !hasSplitBraces) continue;
    const matches = Array.from(
      full.matchAll(/\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g)
    ).reverse();
    for (const m of matches) {
      if (m.index === undefined) continue;
      const rawCap = m[1] || m[2] || '';
      const varName = normalizePlaceholderKey(rawCap).replace(/\s+/g, '_');
      if (!varName || /^\d+$/.test(varName)) continue;
      const mStart = m.index;
      const mEnd = mStart + m[0].length;
      const sEntry = segs.find((s) => mStart >= s.start && mStart < s.end);
      const eEntry = segs.find((s) => mEnd > s.start && mEnd <= s.end);
      if (sEntry && eEntry && sEntry.node !== eEntry.node) {
        try {
          const r = document.createRange();
          r.setStart(sEntry.node, mStart - sEntry.start);
          r.setEnd(eEntry.node, mEnd - eEntry.start);
          r.deleteContents();
          const span = document.createElement('span');
          span.className = 'smart-tag';
          span.setAttribute('data-var', varName);
          span.textContent = `[${varName}]`;
          r.insertNode(span);
        } catch {
          // Ignore range boundary errors
        }
      }
    }
  }
}

/**
 * Updates the visual state of all .smart-tag elements inside rootEl based on whether
 * their variable has a non-empty value in fieldValues:
 * - Unfilled: pink (.smart-tag)
 * - Filled (before final replacement): soft green (.smart-tag.smart-tag-filled) with tooltip showing the value
 */
export function syncSmartTagsFilledStateInDOM(
  rootEl: HTMLElement | null,
  fieldValues: Record<string, string>
): void {
  if (!rootEl) return;
  const spans = Array.from(
    rootEl.querySelectorAll('.smart-tag, .smart-placeholder, [data-var]')
  ) as HTMLElement[];

  for (const span of spans) {
    if (span.hasAttribute('contenteditable')) {
      span.removeAttribute('contenteditable');
    }
    const rawKey =
      span.getAttribute('data-var') ||
      (span.textContent || '').replace(/[\[\]{}]/g, '');
    const key = normalizePlaceholderKey(rawKey).replace(/\s+/g, '_');
    if (!key) continue;
    const val = lookupPlaceholderValue(fieldValues, key);
    if (val !== undefined && val.trim() !== '') {
      span.classList.add('smart-tag-filled');
      span.setAttribute(
        'title',
        `معبأ: "${val.trim()}" — (نقر مزدوج للتعديل)`
      );
    } else {
      span.classList.remove('smart-tag-filled');
      span.setAttribute(
        'title',
        `غير معبأ ([${key}]) — (نقر مزدوج للتعبئة)`
      );
    }
  }
}

/**
 * Enforces the strict Notary Office rule on a container:
 * - Font is always Arial 13pt
 * - Line height is always 1.0
 * - Zero top/bottom margins on paragraphs
 * - Decorates {{...}} smart tags
 */
export function normalizeNotaryContainerDOM(
  rootEl: HTMLElement,
  decorateTags: boolean = true
): void {
  if (rootEl.innerHTML.trim() === '') {
    rootEl.innerHTML = `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><br></p>`;
    return;
  }

  if (decorateTags) {
    decorateSmartTagsInDOM(rootEl);
  }

  const allElements = Array.from(rootEl.querySelectorAll('*')) as HTMLElement[];
  for (const el of allElements) {
    if (el.classList.contains('page-break')) continue;

    if (
      (el.classList.contains('smart-tag') ||
        el.classList.contains('smart-placeholder')) &&
      el.hasAttribute('contenteditable')
    ) {
      el.removeAttribute('contenteditable');
    }

    if (el.style) {
      el.style.fontFamily = STRICT_FONT_FAMILY;
      el.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
      el.style.lineHeight = '1';
      if (
        ['P', 'DIV', 'LI', 'UL', 'OL', 'TABLE', 'H1', 'H2', 'H3', 'H4'].includes(el.tagName)
      ) {
        el.style.marginTop = '0';
        el.style.marginBottom = '0';
      }
    }
    if (el.hasAttribute('face')) el.removeAttribute('face');
    if (el.hasAttribute('size')) el.removeAttribute('size');
  }

  const lastChild = rootEl.lastElementChild;
  if (
    lastChild &&
    (lastChild.classList.contains('page-break') ||
      lastChild.tagName === 'TABLE' ||
      lastChild.classList.contains('clause-container'))
  ) {
    const trailingP = document.createElement('p');
    trailingP.setAttribute('dir', 'rtl');
    trailingP.style.margin = '0';
    trailingP.style.lineHeight = '1';
    trailingP.style.fontFamily = STRICT_FONT_FAMILY;
    trailingP.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    trailingP.style.textAlign = 'justify';
    trailingP.appendChild(document.createElement('br'));
    rootEl.appendChild(trailingP);
  }
}

/**
 * Sanitizes HTML pasted from Microsoft Word or external sources while preserving
 * bold, italic, underline, colors, tables (rowSpan/colSpan/borders), and numbered/bulleted lists,
 * converting [tags] into smart-tags and # lines into clause headings,
 * and strictly enforcing Arial 13pt and 1.0 line spacing with 0 paragraph margins.
 */
export function sanitizePastedWordHTML(rawHtml: string, fallbackText: string): string {
  if (!rawHtml || rawHtml.trim() === '') {
    const lines = fallbackText.split(/\r?\n/);
    const htmlChunks: string[] = [];
    let activeListType: 'ol' | 'ul' | null = null;
    let activeDataListType: ListNumberingStyle = 'decimal';

    const closeActiveList = () => {
      if (activeListType) {
        htmlChunks.push(`</${activeListType}>`);
        activeListType = null;
      }
    };

    for (const rawLine of lines) {
      const trimmed = rawLine.trim();
      const orderedMatch = /^(\d+|[أبجدهوزحطيكلمنسعفصقرشتثخذضظغ])[\.\-\)]\s+(.+)$/.exec(trimmed);
      const bulletMatch = /^([•\*\-–—])\s+(.+)$/.exec(trimmed);

      if (orderedMatch) {
        const marker = orderedMatch[1];
        const itemText = orderedMatch[2];
        const desiredStyle: ListNumberingStyle = /^\d+$/.test(marker)
          ? 'decimal'
          : 'arabic-alpha';
        if (activeListType !== 'ol' || activeDataListType !== desiredStyle) {
          closeActiveList();
          activeListType = 'ol';
          activeDataListType = desiredStyle;
          htmlChunks.push(
            `<ol dir="rtl" data-list-type="${desiredStyle}" style="margin:0;padding-right:22px;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;">`
          );
        }
        htmlChunks.push(
          `<li dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;">${escapeHtml(
            itemText
          )}</li>`
        );
      } else if (bulletMatch) {
        const marker = bulletMatch[1];
        const itemText = bulletMatch[2];
        const desiredStyle: ListNumberingStyle = marker === '-' || marker === '–' ? 'dash' : 'bullet';
        if (activeListType !== 'ul' || activeDataListType !== desiredStyle) {
          closeActiveList();
          activeListType = 'ul';
          activeDataListType = desiredStyle;
          htmlChunks.push(
            `<ul dir="rtl" data-list-type="${desiredStyle}" style="margin:0;padding-right:22px;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;">`
          );
        }
        htmlChunks.push(
          `<li dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;">${escapeHtml(
            itemText
          )}</li>`
        );
      } else {
        closeActiveList();
        const isHashHeading = trimmed.startsWith('#');
        htmlChunks.push(
          `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;${
            isHashHeading ? 'font-weight:bold;' : ''
          }">${trimmed ? escapeHtml(rawLine) : '<br>'}</p>`
        );
      }
    }
    closeActiveList();

    if (typeof document !== 'undefined') {
      const tempContainer = document.createElement('div');
      tempContainer.innerHTML = htmlChunks.join('');
      decorateSmartTagsInDOM(tempContainer);
      return tempContainer.innerHTML;
    }
    // SSR / Node test fallback: string-based smart-tag replacement
    return htmlChunks
      .join('')
      .replace(/\[\s*([^\[\]<>]{1,45}?)\s*\]/g, (_m, key) => {
        const clean = key.trim().replace(/\s+/g, '_');
        return `<span class="smart-tag" data-var="${clean}">[${clean}]</span>`;
      });
  }

  const cleaned = rawHtml
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<meta[^>]*>/gi, '')
    .replace(/<link[^>]*>/gi, '')
    .replace(/<\/?o:[^>]*>/gi, '')
    .replace(/<\/?w:[^>]*>/gi, '')
    .replace(/<\/?v:[^>]*>/gi, '')
    .replace(/<\/?m:[^>]*>/gi, '');

  if (typeof DOMParser === 'undefined') {
    return sanitizePastedWordHTML('', fallbackText || rawHtml);
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(cleaned, 'text/html');

  // Step 1: Convert Word MsoListParagraph ormso-list paragraphs into proper <ol> / <ul> lists
  const topBlocks = Array.from(doc.body.children) as HTMLElement[];
  let currentListEl: HTMLOListElement | HTMLUListElement | null = null;
  let currentListKind: 'ol' | 'ul' | null = null;

  for (const block of topBlocks) {
    const tag = block.tagName.toUpperCase();
    const classAttr = block.getAttribute('class') || '';
    const styleAttr = block.getAttribute('style') || '';
    const rawText = (block.textContent || '').replace(/\u00A0/g, ' ').trim();

    const isWordListPara =
      /MsoListParagraph/i.test(classAttr) || /mso-list/i.test(styleAttr);
    const orderedPrefixMatch = /^(\d+|[أبجدهوزحطيكلمنسعفصقرشتثخذضظغ])[\.\-\)]\s+/.exec(
      rawText
    );
    const bulletPrefixMatch = /^([•·▪◦\*\-–—])\s+/.exec(rawText);

    if (
      (tag === 'P' || tag === 'DIV') &&
      (isWordListPara || orderedPrefixMatch || bulletPrefixMatch)
    ) {
      const isBullet = Boolean(bulletPrefixMatch && !orderedPrefixMatch);
      const targetKind: 'ol' | 'ul' = isBullet ? 'ul' : 'ol';

      // Strip the hardcoded Word list marker span or leading text marker
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      const firstText = walker.nextNode() as Text | null;
      if (firstText && firstText.nodeValue) {
        firstText.nodeValue = firstText.nodeValue
          .replace(/^[\s\u00A0]*(\d+|[أبجدهوزحطيكلمنسعفصقرشتثخذضظغ])[\.\-\)][\s\u00A0]+/, '')
          .replace(/^[\s\u00A0]*[•·▪◦\*\-–—][\s\u00A0]+/, '');
      }

      if (!currentListEl || currentListKind !== targetKind) {
        currentListEl = doc.createElement(targetKind);
        currentListKind = targetKind;
        currentListEl.setAttribute('dir', 'rtl');
        currentListEl.setAttribute(
          'data-list-type',
          targetKind === 'ul' ? 'bullet' : 'decimal'
        );
        block.parentNode?.insertBefore(currentListEl, block);
      }

      const li = doc.createElement('li');
      li.setAttribute('dir', 'rtl');
      li.innerHTML = block.innerHTML;
      currentListEl.appendChild(li);
      block.remove();
    } else {
      currentListEl = null;
      currentListKind = null;
    }
  }

  // Step 2: Sanitize all elements while preserving bold, underline, italic, lists, and tables
  const allNodes = Array.from(doc.body.querySelectorAll('*')) as HTMLElement[];
  for (const el of allNodes) {
    const tag = el.tagName.toUpperCase();
    if (['SCRIPT', 'STYLE', 'META', 'LINK', 'OBJECT', 'IFRAME'].includes(tag)) {
      el.remove();
      continue;
    }

    const computedAlign = el.style?.textAlign || el.getAttribute('align') || '';
    const fontWeight = el.style?.fontWeight || '';
    const fontStyle = el.style?.fontStyle || '';
    const textDecoration =
      el.style?.textDecoration || el.style?.textDecorationLine || '';
    const color = el.style?.color || '';
    const bg = el.style?.backgroundColor || el.style?.background || '';
    const border = el.style?.border || '';
    const dir = el.getAttribute('dir') || el.style?.direction || 'rtl';
    const dataListType = el.getAttribute('data-list-type');

    el.removeAttribute('class');
    el.removeAttribute('id');
    el.removeAttribute('lang');
    el.removeAttribute('face');
    el.removeAttribute('size');
    el.removeAttribute('style');

    if (dataListType) {
      el.setAttribute('data-list-type', dataListType);
    }

    el.style.fontFamily = STRICT_FONT_FAMILY;
    el.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    el.style.lineHeight = '1';

    if (['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(tag)) {
      el.style.marginTop = '0';
      el.style.marginBottom = '0';
      el.setAttribute('dir', dir === 'ltr' ? 'ltr' : 'rtl');
      if (computedAlign) {
        el.style.textAlign =
          computedAlign === 'justify' ? 'justify' : computedAlign;
      } else {
        el.style.textAlign = 'justify';
      }
      const plain = (el.textContent || '').trim();
      if (plain.startsWith('#')) {
        el.style.fontWeight = 'bold';
      }
    }

    if (tag === 'OL' || tag === 'UL') {
      el.setAttribute('dir', 'rtl');
      if (!el.getAttribute('data-list-type')) {
        el.setAttribute('data-list-type', tag === 'UL' ? 'bullet' : 'decimal');
      }
      el.style.marginTop = '0';
      el.style.marginBottom = '0';
      el.style.paddingRight = '22px';
    }

    if (tag === 'B' || tag === 'STRONG' || fontWeight === 'bold' || Number(fontWeight) >= 600) {
      el.style.fontWeight = 'bold';
    }
    if (tag === 'I' || tag === 'EM' || fontStyle === 'italic') {
      el.style.fontStyle = 'italic';
    }
    if (tag === 'U' || textDecoration.includes('underline')) {
      el.style.textDecoration = 'underline';
    }
    if (color && color !== 'windowtext' && color !== '#000000' && color !== 'rgb(0, 0, 0)') {
      el.style.color = color;
    }
    if (bg && bg !== 'transparent' && !bg.includes('none')) {
      el.style.backgroundColor = bg;
    }

    if (tag === 'TABLE') {
      el.setAttribute('dir', 'rtl');
      el.style.width = '100%';
      el.style.maxWidth = '120mm';
      el.style.borderCollapse = 'collapse';
      el.style.tableLayout = 'fixed';
      el.style.marginTop = '0';
      el.style.marginBottom = '0';
    }
    if (tag === 'TD' || tag === 'TH') {
      el.style.border = border || '1px solid #000000';
      el.style.padding = '2px 4px';
      el.style.verticalAlign = 'top';
    }
  }

  decorateSmartTagsInDOM(doc.body);
  return doc.body.innerHTML;
}

export function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extracts all unique [placeholder] (and legacy {{placeholder}}) names from HTML strings.
 */
export function extractPlaceholdersFromHtml(...htmlParts: string[]): string[] {
  const combined = htmlParts.join(' ');
  const found = new Set<string>();

  // Extract from data-var attributes as well as [...] / {{...}} text
  const dataVarRegex = /data-var="([^"]+)"/g;
  let dm: RegExpExecArray | null;
  while ((dm = dataVarRegex.exec(combined)) !== null) {
    const clean = normalizePlaceholderKey(dm[1]).replace(/\s+/g, '_');
    if (clean && !/^\d+$/.test(clean)) found.add(clean);
  }

  const textOnly = combined.replace(/<[^>]+>/g, '').replace(/&nbsp;/gi, ' ');
  const regex = /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(textOnly)) !== null) {
    const raw = match[1] || match[2] || '';
    const key = normalizePlaceholderKey(raw).replace(/\s+/g, '_');
    if (key && !/^\d+$/.test(key)) found.add(key);
  }
  return Array.from(found);
}

/**
 * Extracts placeholders from the live A4 Editor DOM grouped by the Clause (.clause-container)
 * or section heading they appear under, powering the "Dynamic Variables Form by Clause".
 */
export function extractPlaceholdersGroupedByClause(
  rootEl: HTMLElement | null,
  fallbackPlaceholders: string[] = []
): ClauseVariableGroup[] {
  if (!rootEl) {
    return fallbackPlaceholders.length > 0
      ? [
          {
            clauseId: 'general_doc',
            clauseTitle: '1. فقرات العقد العامة والديباجة',
            variables: fallbackPlaceholders,
          },
        ]
      : [];
  }

  const groups: ClauseVariableGroup[] = [];
  const assignedVars = new Set<string>();

  const extractVarsFromElement = (el: HTMLElement): string[] => {
    const vars: string[] = [];
    const seen = new Set<string>();
    const tagEls = Array.from(
      el.querySelectorAll('.smart-tag, .smart-placeholder, [data-var]')
    ) as HTMLElement[];
    for (const t of tagEls) {
      const rawK =
        t.getAttribute('data-var') ||
        (t.textContent || '').replace(/[\[\]{}]/g, '');
      const k = normalizePlaceholderKey(rawK).replace(/\s+/g, '_');
      if (k && !/^\d+$/.test(k) && !seen.has(k)) {
        seen.add(k);
        vars.push(k);
      }
    }
    const textMatches = (el.textContent || '').matchAll(
      /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}]+?)\s*\}\}/g
    );
    for (const m of textMatches) {
      const raw = m[1] || m[2] || '';
      const k = normalizePlaceholderKey(raw).replace(/\s+/g, '_');
      if (k && !/^\d+$/.test(k) && !seen.has(k)) {
        seen.add(k);
        vars.push(k);
      }
    }
    return vars;
  };

  // Walk top-level children of rootEl in document order
  const children = Array.from(rootEl.children) as HTMLElement[];
  let currentGeneralTitle = 'ديباجة العقد والفقرات التمهيدية';
  let currentGeneralId = 'clause_preamble';
  let currentGeneralVars: string[] = [];
  let currentGeneralTexts: string[] = [];

  const flushGeneralGroup = () => {
    if (currentGeneralVars.length > 0) {
      groups.push({
        clauseId: `${currentGeneralId}_${groups.length + 1}`,
        clauseTitle: `${groups.length + 1}. ${currentGeneralTitle}`,
        variables: [...currentGeneralVars],
        clauseReferenceText: currentGeneralTexts.join(' ').replace(/\s+/g, ' ').trim(),
      });
      currentGeneralVars = [];
      currentGeneralTexts = [];
    }
  };

  for (const child of children) {
    if (
      child.classList.contains('clause-container') ||
      child.hasAttribute('data-clause-id')
    ) {
      flushGeneralGroup();
      const cId = child.getAttribute('data-clause-id') || `clause_${groups.length + 1}`;
      const cTitle =
        child.getAttribute('data-clause-title') ||
        child.querySelector('strong, b')?.textContent?.trim() ||
        `بند رقم ${groups.length + 1}`;
      const cVars = extractVarsFromElement(child);
      if (cVars.length > 0) {
        cVars.forEach((v) => assignedVars.add(v));
        const refText = (child.textContent || '')
          .replace(/^#+\s*/, '')
          .replace(/\s+/g, ' ')
          .trim();
        groups.push({
          clauseId: cId,
          clauseTitle: `${groups.length + 1}. ${cTitle}`,
          variables: cVars,
          clauseReferenceText: refText,
        });
      }
    } else {
      // Check if this paragraph is a section/clause heading (e.g. starts with # or is short bold line)
      const plainText = (child.textContent || '').trim();
      const firstStrong = child.querySelector('strong, b');
      const isHeadingParagraph =
        plainText.startsWith('#') ||
        (plainText.length > 2 &&
          plainText.length <= 65 &&
          !plainText.includes('[') &&
          !plainText.includes('{{') &&
          (child.style.fontWeight === 'bold' ||
            (firstStrong && firstStrong.textContent?.trim() === plainText)));

      if (isHeadingParagraph) {
        flushGeneralGroup();
        currentGeneralTitle = plainText.replace(/^#+\s*/, '').replace(/:$/, '').trim();
        currentGeneralId = `heading_${groups.length + 1}`;
      } else if (plainText) {
        currentGeneralTexts.push(plainText);
      }

      const pVars = extractVarsFromElement(child);
      for (const v of pVars) {
        if (!currentGeneralVars.includes(v)) {
          currentGeneralVars.push(v);
          assignedVars.add(v);
        }
      }
    }
  }

  flushGeneralGroup();

  // Catch any fallback placeholders (e.g. in header/footer) not yet grouped
  const remaining = fallbackPlaceholders.filter((k) => !assignedVars.has(k));
  if (remaining.length > 0) {
    groups.push({
      clauseId: 'clause_other',
      clauseTitle: `${groups.length + 1}. متغيرات الترويسة أو متغيرات إضافية`,
      variables: remaining,
    });
  }

  return groups;
}

/**
 * 1-Click Selection-to-Smart-Tag Converter (like localnotaire's [ ] button):
 * Takes the currently highlighted word/phrase inside the A4 editor (or word at caret)
 * and immediately replaces it in-place with <span class="smart-tag" data-var="...">{{...}}</span>
 * without opening any dialog or modal.
 */
export function convertSelectionToSmartTag(
  editorEl: HTMLElement,
  savedRange: Range | null
): string | null {
  editorEl.focus();
  const sel = window.getSelection();
  if (!sel) return null;

  if (
    (sel.rangeCount === 0 || sel.getRangeAt(0).collapsed) &&
    savedRange &&
    editorEl.contains(savedRange.commonAncestorContainer)
  ) {
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }

  if (sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!editorEl.contains(range.commonAncestorContainer)) return null;

  let selectedText = range.toString().trim();

  // If no text is highlighted, try to grab the word around the caret in the text node
  if (!selectedText && range.collapsed && range.startContainer.nodeType === Node.TEXT_NODE) {
    const tNode = range.startContainer as Text;
    const text = tNode.nodeValue || '';
    const offset = range.startOffset;
    let start = offset;
    let end = offset;
    while (start > 0 && !/\s|[.,،؛:()[\]{}]/.test(text[start - 1])) {
      start--;
    }
    while (end < text.length && !/\s|[.,،؛:()[\]{}]/.test(text[end])) {
      end++;
    }
    if (end > start) {
      range.setStart(tNode, start);
      range.setEnd(tNode, end);
      selectedText = range.toString().trim();
    }
  }

  if (!selectedText) return null;

  // Clean variable name: strip braces/brackets and convert spaces to underscores
  const varName = normalizePlaceholderKey(selectedText).replace(/\s+/g, '_');

  if (!varName) return null;

  const span = document.createElement('span');
  span.className = 'smart-tag';
  span.setAttribute('data-var', varName);
  span.textContent = `[${varName}]`;

  range.deleteContents();
  range.insertNode(span);

  // Place caret right after the newly created smart tag
  const afterSpace = document.createTextNode('\u00A0');
  span.parentNode?.insertBefore(afterSpace, span.nextSibling);

  const newRange = document.createRange();
  newRange.setStartAfter(afterSpace);
  newRange.collapse(true);
  sel.removeAllRanges();
  sel.addRange(newRange);

  normalizeNotaryContainerDOM(editorEl, false);
  return varName;
}


/**
 * Inserts HTML snippet or template at the active caret position inside the editor.
 */
export function insertHtmlAtSelection(
  editorEl: HTMLElement,
  htmlToInsert: string,
  savedRange: Range | null
): void {
  editorEl.focus();
  const sel = window.getSelection();
  if (!sel) return;

  if (savedRange && editorEl.contains(savedRange.commonAncestorContainer)) {
    sel.removeAllRanges();
    sel.addRange(savedRange);
  }

  if (sel.rangeCount === 0) {
    const fallbackRange = document.createRange();
    fallbackRange.selectNodeContents(editorEl);
    fallbackRange.collapse(false);
    sel.addRange(fallbackRange);
  }

  const range = sel.getRangeAt(0);
  range.deleteContents();

  const temp = document.createElement('div');
  temp.innerHTML = htmlToInsert;
  normalizeNotaryContainerDOM(temp);

  const frag = document.createDocumentFragment();
  let lastInsertedNode: Node | null = null;

  const children = Array.from(temp.childNodes);
  const isSingleParagraph =
    children.length === 1 &&
    children[0].nodeType === Node.ELEMENT_NODE &&
    (children[0] as HTMLElement).tagName === 'P';

  if (isSingleParagraph) {
    const pEl = children[0] as HTMLElement;
    while (pEl.firstChild) {
      lastInsertedNode = pEl.firstChild;
      frag.appendChild(lastInsertedNode);
    }
  } else {
    while (temp.firstChild) {
      lastInsertedNode = temp.firstChild;
      frag.appendChild(lastInsertedNode);
    }
  }

  range.insertNode(frag);

  if (lastInsertedNode) {
    const newRange = document.createRange();
    newRange.setStartAfter(lastInsertedNode);
    newRange.collapse(true);
    sel.removeAllRanges();
    sel.addRange(newRange);
  }

  normalizeNotaryContainerDOM(editorEl);
}

/**
 * Multi-Node Find & Replace across formatted text runs using TreeWalker(NodeFilter.SHOW_TEXT).
 */
export interface MultiNodeTextMatch {
  startNode: Text;
  startOffset: number;
  endNode: Text;
  endOffset: number;
  matchedText: string;
}

/**
 * Normalizes Arabic text for smart search (strips Tashkeel/Tatweel and unifies Hamza/Alef/Ta-Marbuta)
 * while maintaining an exact index map back to original string positions.
 */
function buildNormalizedSearchString(
  input: string,
  ignoreArabicHamzaAndDiacritics: boolean
): { normText: string; normToOrigStart: number[]; normToOrigEnd: number[] } {
  if (!ignoreArabicHamzaAndDiacritics) {
    const starts = Array.from({ length: input.length }, (_, i) => i);
    const ends = Array.from({ length: input.length }, (_, i) => i + 1);
    return {
      normText: input.toLowerCase(),
      normToOrigStart: starts,
      normToOrigEnd: ends,
    };
  }

  let normText = '';
  const normToOrigStart: number[] = [];
  const normToOrigEnd: number[] = [];

  const isDiacritic = (ch: string) => /[\u064B-\u065F\u0670\u0640]/.test(ch);
  const normalizeArabicChar = (ch: string): string => {
    if (/[أإآٱ]/.test(ch)) return 'ا';
    if (ch === 'ؤ') return 'و';
    if (ch === 'ئ' || ch === 'ى') return 'ي';
    if (ch === 'ة') return 'ه';
    return ch.toLowerCase();
  };

  for (let i = 0; i < input.length; i++) {
    const ch = input[i];
    if (isDiacritic(ch)) {
      if (normToOrigEnd.length > 0) {
        normToOrigEnd[normToOrigEnd.length - 1] = i + 1;
      }
      continue;
    }
    normText += normalizeArabicChar(ch);
    normToOrigStart.push(i);
    normToOrigEnd.push(i + 1);
  }

  return { normText, normToOrigStart, normToOrigEnd };
}

export function findMatchesAcrossNodes(
  rootEl: HTMLElement,
  query: string,
  ignoreArabicHamzaAndDiacritics: boolean = true
): MultiNodeTextMatch[] {
  if (!query || !query.trim()) return [];
  const { normText: normQuery } = buildNormalizedSearchString(
    query,
    ignoreArabicHamzaAndDiacritics
  );
  if (!normQuery) return [];

  const results: MultiNodeTextMatch[] = [];

  const blocks = Array.from(
    rootEl.querySelectorAll('p, li, td, th, h1, h2, h3, h4')
  ) as HTMLElement[];
  const containers = blocks.length > 0 ? blocks : [rootEl];

  const leafBlocks = containers.filter(
    (b) => !containers.some((other) => other !== b && b.contains(other))
  );

  for (const block of leafBlocks) {
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    const textNodes: { node: Text; start: number; end: number }[] = [];
    let fullText = '';
    let current: Node | null;

    while ((current = walker.nextNode())) {
      const tNode = current as Text;
      const len = (tNode.nodeValue || '').length;
      if (len > 0) {
        textNodes.push({
          node: tNode,
          start: fullText.length,
          end: fullText.length + len,
        });
        fullText += tNode.nodeValue;
      }
    }

    if (!fullText) continue;

    const { normText, normToOrigStart, normToOrigEnd } = buildNormalizedSearchString(
      fullText,
      ignoreArabicHamzaAndDiacritics
    );
    if (!normText || normText.length < normQuery.length) continue;

    let searchIdx = 0;
    while (searchIdx <= normText.length - normQuery.length) {
      const normMatchStart = normText.indexOf(normQuery, searchIdx);
      if (normMatchStart === -1) break;
      const normMatchEnd = normMatchStart + normQuery.length - 1;

      const matchStart = normToOrigStart[normMatchStart];
      const matchEnd = normToOrigEnd[normMatchEnd];

      const startEntry = textNodes.find((t) => matchStart >= t.start && matchStart < t.end);
      const endEntry = textNodes.find((t) => matchEnd > t.start && matchEnd <= t.end);

      if (startEntry && endEntry) {
        results.push({
          startNode: startEntry.node,
          startOffset: matchStart - startEntry.start,
          endNode: endEntry.node,
          endOffset: matchEnd - endEntry.start,
          matchedText: fullText.slice(matchStart, matchEnd),
        });
      }

      searchIdx = normMatchStart + normQuery.length;
    }
  }

  return results;
}

export function focusAndSelectMatchInDOM(match: MultiNodeTextMatch | null): void {
  if (!match || typeof window === 'undefined') return;
  try {
    const range = document.createRange();
    range.setStart(match.startNode, match.startOffset);
    range.setEnd(match.endNode, match.endOffset);
    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(range);
    }
    const parentEl = match.startNode.parentElement;
    if (parentEl) {
      parentEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  } catch {
    // Ignore detached DOM node errors
  }
}

export function replaceMatchesAcrossNodes(
  rootEl: HTMLElement,
  query: string,
  replacement: string,
  replaceAll: boolean,
  matchIndex: number = 0,
  ignoreArabicHamzaAndDiacritics: boolean = true
): number {
  const matches = findMatchesAcrossNodes(
    rootEl,
    query,
    ignoreArabicHamzaAndDiacritics
  );
  if (matches.length === 0) return 0;

  const targetMatches = replaceAll
    ? [...matches].reverse()
    : [matches[Math.min(Math.max(0, matchIndex), matches.length - 1)]];

  let replacedCount = 0;

  for (const m of targetMatches) {
    try {
      if (m.startNode === m.endNode) {
        const val = m.startNode.nodeValue || '';
        m.startNode.nodeValue =
          val.slice(0, m.startOffset) + replacement + val.slice(m.endOffset);
      } else {
        const range = document.createRange();
        range.setStart(m.startNode, m.startOffset);
        range.setEnd(m.endNode, m.endOffset);
        range.deleteContents();
        m.startNode.nodeValue =
          (m.startNode.nodeValue || '').slice(0, m.startOffset) +
          replacement +
          (m.startNode.nodeValue || '').slice(m.startOffset);
      }
      replacedCount++;
    } catch {
      // Ignore DOM mutation boundary error
    }
  }

  normalizeNotaryContainerDOM(rootEl);
  return replacedCount;
}

/**
 * Computes word-by-word Diff tokens between two HTML versions for the Version Diff Viewer.
 */
export interface DiffSegment {
  type: 'equal' | 'added' | 'removed';
  text: string;
}

export function computeTextDiff(oldHtml: string, newHtml: string): DiffSegment[] {
  const toWords = (h: string) =>
    h
      .replace(/<\/p>|<\/div>|<\/tr>|<\/li>|<br\s*\/?>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .split(/(\s+)/)
      .filter((t) => t.length > 0);

  const a = toWords(oldHtml);
  const b = toWords(newHtml);

  // Limit LCS matrix size for performance on very large contracts
  const maxTokens = 1600;
  const aSlice = a.slice(0, maxTokens);
  const bSlice = b.slice(0, maxTokens);
  const n = aSlice.length;
  const m = bSlice.length;

  const dp: Uint16Array[] = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));

  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      if (aSlice[i] === bSlice[j]) {
        dp[i][j] = dp[i + 1][j + 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
  }

  const raw: DiffSegment[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (aSlice[i] === bSlice[j]) {
      raw.push({ type: 'equal', text: aSlice[i] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      raw.push({ type: 'removed', text: aSlice[i] });
      i++;
    } else {
      raw.push({ type: 'added', text: bSlice[j] });
      j++;
    }
  }
  while (i < n) {
    raw.push({ type: 'removed', text: aSlice[i++] });
  }
  while (j < m) {
    raw.push({ type: 'added', text: bSlice[j++] });
  }

  // Merge adjacent segments of the same type
  const merged: DiffSegment[] = [];
  for (const seg of raw) {
    if (merged.length > 0 && merged[merged.length - 1].type === seg.type) {
      merged[merged.length - 1].text += seg.text;
    } else {
      merged.push({ ...seg });
    }
  }
  return merged;
}

/**
 * Modern Range-based Inline Formatting Engine (replaces legacy document.execCommand where possible).
 * Wraps or unwraps selected text in <strong>, <em>, <u>, or styled <span> while preserving selection.
 */
export function applyModernInlineFormattingToRange(
  rootEl: HTMLElement,
  command: 'bold' | 'italic' | 'underline' | 'foreColor' | 'hiliteColor',
  value?: string
): boolean {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return false;
  const range = sel.getRangeAt(0);
  if (!rootEl.contains(range.commonAncestorContainer)) return false;

  // If selection is collapsed, or spans across multiple block paragraphs,
  // check if single block can be cleanly wrapped via Range; otherwise fallback safely.
  if (range.collapsed) {
    // Expand to current word at caret if user clicks Bold/Underline on a word without dragging
    const node = range.startContainer;
    if (node.nodeType === Node.TEXT_NODE && node.nodeValue) {
      const text = node.nodeValue;
      const offset = range.startOffset;
      let start = offset;
      let end = offset;
      while (start > 0 && !/\s/.test(text[start - 1])) start--;
      while (end < text.length && !/\s/.test(text[end])) end++;
      if (end > start) {
        range.setStart(node, start);
        range.setEnd(node, end);
      } else {
        return false;
      }
    } else {
      return false;
    }
  }

  const tagMap: Record<string, string> = {
    bold: 'strong',
    italic: 'em',
    underline: 'u',
  };

  // Check if selection is already inside the target inline formatting element
  let ancestor: HTMLElement | null =
    range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
      ? (range.commonAncestorContainer as HTMLElement)
      : range.commonAncestorContainer.parentElement;

  if (command === 'bold' || command === 'italic' || command === 'underline') {
    const targetTag = tagMap[command].toUpperCase();
    const altTag = command === 'bold' ? 'B' : command === 'italic' ? 'I' : 'U';

    while (ancestor && ancestor !== rootEl) {
      const t = ancestor.tagName.toUpperCase();
      const hasInlineStyle =
        (command === 'bold' &&
          (ancestor.style.fontWeight === 'bold' ||
            Number(ancestor.style.fontWeight) >= 600)) ||
        (command === 'italic' && ancestor.style.fontStyle === 'italic') ||
        (command === 'underline' &&
          ancestor.style.textDecoration.includes('underline'));

      if (t === targetTag || t === altTag || hasInlineStyle) {
        // Unwrap formatting if the entire element is selected or contains the range
        if (t === targetTag || t === altTag) {
          const parent = ancestor.parentNode;
          if (parent) {
            const firstChild = ancestor.firstChild;
            const lastChild = ancestor.lastChild;
            while (ancestor.firstChild) {
              parent.insertBefore(ancestor.firstChild, ancestor);
            }
            parent.removeChild(ancestor);
            if (firstChild && lastChild) {
              const newRange = document.createRange();
              newRange.setStartBefore(firstChild);
              newRange.setEndAfter(lastChild);
              sel.removeAllRanges();
              sel.addRange(newRange);
            }
            normalizeNotaryContainerDOM(rootEl);
            return true;
          }
        } else if (hasInlineStyle) {
          if (command === 'bold') ancestor.style.fontWeight = 'normal';
          if (command === 'italic') ancestor.style.fontStyle = 'normal';
          if (command === 'underline') ancestor.style.textDecoration = 'none';
          return true;
        }
      }
      ancestor = ancestor.parentElement;
    }
  }

  // Ensure range doesn't split across block elements; if it stays within a block, wrap via extractContents
  const blocks = getIntersectingBlockElements(rootEl, range);
  if (blocks.length <= 1) {
    try {
      const wrapper = document.createElement(
        command === 'bold'
          ? 'strong'
          : command === 'italic'
          ? 'em'
          : command === 'underline'
          ? 'u'
          : 'span'
      );
      if (command === 'foreColor' && value) {
        wrapper.style.color = value;
      } else if (command === 'hiliteColor' && value) {
        wrapper.style.backgroundColor = value;
      }
      const contents = range.extractContents();
      wrapper.appendChild(contents);
      range.insertNode(wrapper);

      const nextRange = document.createRange();
      nextRange.selectNodeContents(wrapper);
      sel.removeAllRanges();
      sel.addRange(nextRange);
      normalizeNotaryContainerDOM(rootEl);
      return true;
    } catch {
      return false;
    }
  }

  // Multi-block selection: apply formatting to each block's inline content cleanly
  for (const block of blocks) {
    if (command === 'bold') {
      const isBold =
        block.style.fontWeight === 'bold' || Number(block.style.fontWeight) >= 600;
      block.style.fontWeight = isBold ? '' : 'bold';
    } else if (command === 'italic') {
      block.style.fontStyle = block.style.fontStyle === 'italic' ? '' : 'italic';
    } else if (command === 'underline') {
      block.style.textDecoration = block.style.textDecoration.includes('underline')
        ? ''
        : 'underline';
    } else if (command === 'foreColor' && value) {
      block.style.color = value;
    } else if (command === 'hiliteColor' && value) {
      block.style.backgroundColor = value === 'transparent' ? '' : value;
    }
  }
  normalizeNotaryContainerDOM(rootEl);
  return true;
}

/**
 * Modern Range-based List Toggle (converts intersecting paragraphs to <ol>/<ul> or unwraps them back to <p>).
 */
export function applyModernListToggleToRange(
  rootEl: HTMLElement,
  style: ListNumberingStyle
): boolean {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return false;
  const range = sel.getRangeAt(0);
  if (!rootEl.contains(range.commonAncestorContainer)) return false;

  const isOrdered =
    style === 'decimal' ||
    style === 'arabic-alpha' ||
    style === 'arabic-abjad';
  const targetListTag = isOrdered ? 'ol' : 'ul';

  // Check if selection is already inside an <ol> or <ul>
  let currentEl: HTMLElement | null =
    range.commonAncestorContainer.nodeType === Node.ELEMENT_NODE
      ? (range.commonAncestorContainer as HTMLElement)
      : range.commonAncestorContainer.parentElement;

  let existingList: HTMLElement | null = null;
  while (currentEl && currentEl !== rootEl) {
    if (currentEl.tagName === 'OL' || currentEl.tagName === 'UL') {
      existingList = currentEl;
      break;
    }
    currentEl = currentEl.parentElement;
  }

  if (existingList) {
    const currentStyle = existingList.getAttribute('data-list-type');
    const currentTag = existingList.tagName.toLowerCase();

    // If same list style is clicked again, unwrap the list back into standard notarial <p> paragraphs
    if (currentStyle === style && currentTag === targetListTag) {
      const parent = existingList.parentNode;
      if (!parent) return false;
      const items = Array.from(existingList.querySelectorAll(':scope > li'));
      let lastP: HTMLElement | null = null;
      for (const li of items) {
        const p = document.createElement('p');
        p.setAttribute('dir', li.getAttribute('dir') || 'rtl');
        p.style.margin = '0';
        p.style.lineHeight = '1';
        p.style.fontFamily = STRICT_FONT_FAMILY;
        p.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
        p.style.textAlign = 'justify';
        p.innerHTML = li.innerHTML || '<br>';
        parent.insertBefore(p, existingList);
        lastP = p;
      }
      parent.removeChild(existingList);
      if (lastP) {
        const newRange = document.createRange();
        newRange.selectNodeContents(lastP);
        newRange.collapse(false);
        sel.removeAllRanges();
        sel.addRange(newRange);
      }
      normalizeNotaryContainerDOM(rootEl);
      return true;
    }

    // Otherwise switch list type/tag in-place
    if (currentTag !== targetListTag) {
      const replacementList = document.createElement(targetListTag);
      replacementList.setAttribute('dir', 'rtl');
      replacementList.setAttribute('data-list-type', style);
      replacementList.style.margin = '0';
      replacementList.style.paddingRight = '22px';
      replacementList.style.lineHeight = '1';
      replacementList.style.fontFamily = STRICT_FONT_FAMILY;
      replacementList.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
      while (existingList.firstChild) {
        replacementList.appendChild(existingList.firstChild);
      }
      existingList.parentNode?.replaceChild(replacementList, existingList);
    } else {
      existingList.setAttribute('data-list-type', style);
      existingList.setAttribute('dir', 'rtl');
    }
    normalizeNotaryContainerDOM(rootEl);
    return true;
  }

  // Convert intersecting block paragraphs into a new <ol> or <ul>
  const blocks = getIntersectingBlockElements(rootEl, range).filter(
    (b) => b.tagName === 'P' || b.tagName === 'DIV'
  );
  if (blocks.length === 0) return false;

  const firstBlock = blocks[0];
  const parent = firstBlock.parentNode;
  if (!parent) return false;

  const listEl = document.createElement(targetListTag);
  listEl.setAttribute('dir', 'rtl');
  listEl.setAttribute('data-list-type', style);
  listEl.style.margin = '0';
  listEl.style.paddingRight = '22px';
  listEl.style.lineHeight = '1';
  listEl.style.fontFamily = STRICT_FONT_FAMILY;
  listEl.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;

  parent.insertBefore(listEl, firstBlock);

  let lastLi: HTMLElement | null = null;
  for (const b of blocks) {
    const li = document.createElement('li');
    li.setAttribute('dir', b.getAttribute('dir') || 'rtl');
    li.style.margin = '0';
    li.style.lineHeight = '1';
    li.style.fontFamily = STRICT_FONT_FAMILY;
    li.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    li.style.textAlign = 'justify';
    li.innerHTML = b.innerHTML || '<br>';
    listEl.appendChild(li);
    b.remove();
    lastLi = li;
  }

  if (lastLi) {
    const newRange = document.createRange();
    newRange.selectNodeContents(lastLi);
    newRange.collapse(false);
    sel.removeAllRanges();
    sel.addRange(newRange);
  }

  normalizeNotaryContainerDOM(rootEl);
  return true;
}

/**
 * Atomic Smart Tag Navigation & Deletion Helper:
 * Detects if the collapsed caret is immediately adjacent to or inside a `.smart-tag` element.
 */
export function findAdjacentSmartTagElement(
  direction: 'backward' | 'forward'
): HTMLElement | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || !sel.isCollapsed) return null;
  const range = sel.getRangeAt(0);
  const node = range.startContainer;
  const offset = range.startOffset;

  // Case 1: Caret is inside a .smart-tag
  const parentEl =
    node.nodeType === Node.ELEMENT_NODE
      ? (node as HTMLElement)
      : node.parentElement;
  const insideTag = parentEl?.closest?.(
    '.smart-tag, .smart-placeholder'
  ) as HTMLElement | null;
  if (insideTag) {
    return insideTag;
  }

  const isSmartTagNode = (n: Node | null): HTMLElement | null => {
    if (!n || n.nodeType !== Node.ELEMENT_NODE) return null;
    const el = n as HTMLElement;
    if (
      el.classList.contains('smart-tag') ||
      el.classList.contains('smart-placeholder')
    ) {
      return el;
    }
    return null;
  };

  // Case 2: Caret is in a Text node at the very start (backward) or very end (forward)
  if (node.nodeType === Node.TEXT_NODE) {
    const textLen = (node.nodeValue || '').length;
    if (direction === 'backward' && offset === 0) {
      let prev = node.previousSibling;
      while (
        prev &&
        prev.nodeType === Node.TEXT_NODE &&
        (prev.nodeValue || '') === ''
      ) {
        prev = prev.previousSibling;
      }
      return isSmartTagNode(prev);
    }
    if (direction === 'forward' && offset === textLen) {
      let next = node.nextSibling;
      while (
        next &&
        next.nodeType === Node.TEXT_NODE &&
        (next.nodeValue || '') === ''
      ) {
        next = next.nextSibling;
      }
      return isSmartTagNode(next);
    }
  }

  // Case 3: Caret is in an Element container between child nodes
  if (node.nodeType === Node.ELEMENT_NODE) {
    const children = node.childNodes;
    if (direction === 'backward' && offset > 0) {
      return isSmartTagNode(children[offset - 1]);
    }
    if (direction === 'forward' && offset < children.length) {
      return isSmartTagNode(children[offset]);
    }
  }

  return null;
}

/**
 * Moves the caret cleanly before or after an atomic `.smart-tag` element so the user skips over it as a single unit.
 */
export function moveCaretOutsideSmartTag(
  tagEl: HTMLElement,
  position: 'before' | 'after'
): void {
  const sel = window.getSelection();
  if (!sel) return;
  const range = document.createRange();
  if (position === 'before') {
    range.setStartBefore(tagEl);
  } else {
    range.setStartAfter(tagEl);
  }
  range.collapse(true);
  sel.removeAllRanges();
  sel.addRange(range);
}

/**
 * Detects inline autocomplete or quick-command triggers at the current caret position:
 * - `[` followed by 0–35 chars (without closing `]`) -> tag autocomplete
 * - `#` at the start of a line -> clause heading autocomplete
 * - `/` at the start of a line -> quick command menu
 */
export interface CaretInlineTrigger {
  type: 'tag' | 'clause' | 'slash';
  query: string;
  rect: { top: number; left: number; bottom: number };
  textNode: Text;
  triggerOffset: number;
  caretOffset: number;
}

export function detectCaretInlineTrigger(
  rootEl: HTMLElement
): CaretInlineTrigger | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || !sel.isCollapsed) return null;
  const range = sel.getRangeAt(0);
  if (!rootEl.contains(range.commonAncestorContainer)) return null;

  const node = range.startContainer;
  if (node.nodeType !== Node.TEXT_NODE) return null;
  const textNode = node as Text;
  const parentEl = textNode.parentElement;
  if (!parentEl || parentEl.closest('.smart-tag, .smart-placeholder')) {
    return null;
  }

  const textBeforeCaret = (textNode.nodeValue || '').slice(0, range.startOffset);
  if (!textBeforeCaret) return null;

  // Check if this text node is at the beginning of its block paragraph
  const blockEl = parentEl.closest('p, div, li, h1, h2, h3, h4') as HTMLElement | null;
  let fullBlockPrefix = textBeforeCaret;
  if (blockEl) {
    try {
      const preRange = document.createRange();
      preRange.selectNodeContents(blockEl);
      preRange.setEnd(textNode, range.startOffset);
      fullBlockPrefix = preRange.toString();
    } catch {
      fullBlockPrefix = textBeforeCaret;
    }
  }

  const measureRect = (triggerIndexInNode: number) => {
    try {
      const measureRange = document.createRange();
      measureRange.setStart(textNode, triggerIndexInNode);
      measureRange.setEnd(textNode, range.startOffset);
      const r = measureRange.getBoundingClientRect();
      if (r.width > 0 || r.height > 0) {
        return { top: r.top, left: r.left, bottom: r.bottom };
      }
    } catch {
      // Fallback
    }
    const fallbackRect = parentEl.getBoundingClientRect();
    return {
      top: fallbackRect.top,
      left: fallbackRect.left,
      bottom: fallbackRect.bottom,
    };
  };

  // 1. Check `[` tag trigger anywhere in the line (not closed by `]`)
  const bracketMatch = /\[([^\[\]\s]{0,35})$/.exec(textBeforeCaret);
  if (bracketMatch) {
    const triggerOffset = range.startOffset - bracketMatch[0].length;
    return {
      type: 'tag',
      query: bracketMatch[1] || '',
      rect: measureRect(triggerOffset),
      textNode,
      triggerOffset,
      caretOffset: range.startOffset,
    };
  }

  // 2. Check `#` clause heading trigger at start of line
  const hashMatch = /^\s*#([^\r\n#]{0,30})$/.exec(fullBlockPrefix);
  if (hashMatch && textBeforeCaret.includes('#')) {
    const localHashIdx = textBeforeCaret.lastIndexOf('#');
    if (localHashIdx !== -1) {
      return {
        type: 'clause',
        query: hashMatch[1].trim(),
        rect: measureRect(localHashIdx),
        textNode,
        triggerOffset: localHashIdx,
        caretOffset: range.startOffset,
      };
    }
  }

  // 3. Check `/` slash command trigger at start of line
  const slashMatch = /^\s*\/([^\r\n/\s]{0,20})$/.exec(fullBlockPrefix);
  if (slashMatch && textBeforeCaret.includes('/')) {
    const localSlashIdx = textBeforeCaret.lastIndexOf('/');
    if (localSlashIdx !== -1) {
      return {
        type: 'slash',
        query: slashMatch[1].trim(),
        rect: measureRect(localSlashIdx),
        textNode,
        triggerOffset: localSlashIdx,
        caretOffset: range.startOffset,
      };
    }
  }

  return null;
}

