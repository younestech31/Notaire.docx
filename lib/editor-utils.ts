import {
  ClauseVariableGroup,
  ContractOutlineClause,
  ContractPartyCard,
  SerializedSelectionPath,
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
 * bold, italic, underline, colors, tables (rowSpan/colSpan/borders), and lists,
 * and strictly enforcing Arial 13pt and 1.0 line spacing with 0 paragraph margins.
 */
export function sanitizePastedWordHTML(rawHtml: string, fallbackText: string): string {
  if (!rawHtml || rawHtml.trim() === '') {
    return fallbackText
      .split(/\r?\n/)
      .map(
        (line) =>
          `<p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;text-align:justify;">${
            line.trim() ? escapeHtml(line) : '<br>'
          }</p>`
      )
      .join('');
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

  const parser = new DOMParser();
  const doc = parser.parseFromString(cleaned, 'text/html');

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
    const textDecoration = el.style?.textDecoration || '';
    const color = el.style?.color || '';
    const bg = el.style?.backgroundColor || el.style?.background || '';
    const border = el.style?.border || '';
    const dir = el.getAttribute('dir') || el.style?.direction || 'rtl';

    el.removeAttribute('class');
    el.removeAttribute('id');
    el.removeAttribute('lang');
    el.removeAttribute('face');
    el.removeAttribute('size');
    el.removeAttribute('style');

    el.style.fontFamily = STRICT_FONT_FAMILY;
    el.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
    el.style.lineHeight = '1';

    if (['P', 'DIV', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(tag)) {
      el.style.marginTop = '0';
      el.style.marginBottom = '0';
      el.setAttribute('dir', dir === 'ltr' ? 'ltr' : 'rtl');
      if (computedAlign) {
        el.style.textAlign = computedAlign === 'justify' ? 'justify' : computedAlign;
      }
    }

    if (fontWeight === 'bold' || Number(fontWeight) >= 600) {
      el.style.fontWeight = 'bold';
    }
    if (fontStyle === 'italic') {
      el.style.fontStyle = 'italic';
    }
    if (textDecoration.includes('underline')) {
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

  const flushGeneralGroup = () => {
    if (currentGeneralVars.length > 0) {
      groups.push({
        clauseId: `${currentGeneralId}_${groups.length + 1}`,
        clauseTitle: `${groups.length + 1}. ${currentGeneralTitle}`,
        variables: [...currentGeneralVars],
      });
      currentGeneralVars = [];
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
        groups.push({
          clauseId: cId,
          clauseTitle: `${groups.length + 1}. ${cTitle}`,
          variables: cVars,
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
 * # New Clause Toolbar Action:
 * Converts the current line/selection (or inserts at caret) into a structured
 * .clause-container with a bold heading so it immediately links with the Clauses Sidebar.
 */
export function insertOrWrapNewClauseAtSelection(
  editorEl: HTMLElement,
  savedRange: Range | null,
  clauseIndexHint: number = 1
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

  let rawTitle = '';
  if (sel && sel.rangeCount > 0) {
    const r = sel.getRangeAt(0);
    if (!r.collapsed && editorEl.contains(r.commonAncestorContainer)) {
      rawTitle = r.toString().trim();
      r.deleteContents();
    } else if (r.collapsed && editorEl.contains(r.commonAncestorContainer)) {
      // Check if current paragraph starts with # or has a short title
      let block: Node | null = r.startContainer;
      while (block && block !== editorEl && (block as HTMLElement).tagName !== 'P') {
        block = block.parentNode;
      }
      if (block && (block as HTMLElement).tagName === 'P') {
        const pText = (block.textContent || '').trim();
        if (pText.startsWith('#') || (pText.length > 0 && pText.length <= 60)) {
          rawTitle = pText.replace(/^#+\s*/, '').trim();
          (block as HTMLElement).innerHTML = '<br>';
        }
      }
    }
  }

  const title = rawTitle || `بند رقم ${clauseIndexHint}`;
  const clauseId = `clause_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;

  const innerContentHtml = `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># ${escapeHtml(
    title
  )}</span></p><p dir="rtl" style="${pStyle}">اكتب نص البند أو أدرج المتغيرات [اسم_المتغير] هنا.</p>`;

  const wrapperHtml = `<div class="clause-container" data-clause-id="${clauseId}" data-clause-order="${clauseIndexHint - 1}" data-clause-title="${escapeHtml(
    title
  )}" data-clause-locked="false">${innerContentHtml}</div><p dir="rtl" style="${pStyle}"><br></p>`;

  insertHtmlAtSelection(editorEl, wrapperHtml, savedRange);
  return { clauseId, title, contentHtml: innerContentHtml };
}

/**
 * Helper to segment rootEl.children into logical contract clauses based on:
 * 1. Explicit .clause-container[data-clause-id] elements
 * 2. Paragraphs starting with `#` (Notaire Local Layer 2 syntax: `# عنوان البند`)
 */
interface DomClauseSegment {
  id: string;
  title: string;
  domIndex: number;
  isContainer: boolean;
  locked: boolean;
  explicitOrder?: number;
  elements: HTMLElement[];
}

function collectDomClauseSegments(rootEl: HTMLElement): DomClauseSegment[] {
  const children = Array.from(rootEl.children) as HTMLElement[];
  const segments: DomClauseSegment[] = [];

  let currentHashSegment: DomClauseSegment | null = null;

  const flushHashSegment = () => {
    if (currentHashSegment) {
      segments.push(currentHashSegment);
      currentHashSegment = null;
    }
  };

  children.forEach((child, idx) => {
    if (child.classList.contains('page-break')) {
      if (currentHashSegment) {
        currentHashSegment.elements.push(child);
      }
      return;
    }

    if (
      child.classList.contains('clause-container') ||
      child.hasAttribute('data-clause-id')
    ) {
      flushHashSegment();
      const rawTitle =
        child.getAttribute('data-clause-title') ||
        child.querySelector('strong, b, h1, h2, h3, h4')?.textContent?.trim() ||
        `بند رقم ${segments.length + 1}`;
      const cleanTitle = rawTitle.replace(/^#+\s*/, '').replace(/:$/, '').trim();
      const cId =
        child.getAttribute('data-clause-id') || `dom_clause_${idx}`;
      const locked =
        child.getAttribute('data-clause-locked') === 'true' ||
        child.getAttribute('contenteditable') === 'false';
      const orderAttr = child.getAttribute('data-clause-order');
      const explicitOrder =
        orderAttr !== null && orderAttr !== '' && !Number.isNaN(Number(orderAttr))
          ? Number(orderAttr)
          : undefined;
      segments.push({
        id: cId,
        title: cleanTitle || `بند رقم ${segments.length + 1}`,
        domIndex: idx,
        isContainer: true,
        locked,
        explicitOrder,
        elements: [child],
      });
      return;
    }

    const plainText = (child.textContent || '').trim();
    if (plainText.startsWith('#')) {
      flushHashSegment();
      const cleanTitle =
        plainText.replace(/^#+\s*/, '').replace(/:$/, '').trim() ||
        `بند رقم ${segments.length + 1}`;
      currentHashSegment = {
        id: `hash_clause_${idx}`,
        title: cleanTitle,
        domIndex: idx,
        isContainer: false,
        locked: false,
        elements: [child],
      };
    } else if (currentHashSegment) {
      currentHashSegment.elements.push(child);
    }
  });

  flushHashSegment();
  return segments;
}

/**
 * Safely upgrades any top-level `# عنوان البند` segments into `<div class="clause-container">`
 * on initial document load or explicit structure normalization, and stamps `data-clause-order`.
 * Never called during `onInput` typing!
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
      (c) => c.id === seg.id || c.title === seg.title
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
      containerEl.setAttribute('data-clause-locked', isLocked ? 'true' : 'false');
      if (isLocked) {
        containerEl.setAttribute('contenteditable', 'false');
        containerEl.style.backgroundColor = 'rgba(248, 250, 252, 0.75)';
        containerEl.style.borderRight = '3px solid #f59e0b';
        containerEl.style.paddingRight = '6px';
      }
    } else {
      // Wrap raw `#` heading and its body paragraphs in a .clause-container in-place
      const firstEl = seg.elements[0];
      if (!firstEl || !firstEl.parentNode) return;
      const wrapper = document.createElement('div');
      wrapper.className = 'clause-container';
      const newId =
        matchedPrev?.id && !matchedPrev.id.startsWith('hash_clause_')
          ? matchedPrev.id
          : `clause_${Date.now()}_${idx}`;
      wrapper.setAttribute('data-clause-id', newId);
      wrapper.setAttribute('data-clause-order', String(clauseOrder));
      wrapper.setAttribute('data-clause-title', seg.title);
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
        previewSource = (containerEl.textContent || '').replace(/^#+\s*/, '');
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
      previewSource = (seg.elements[0]?.textContent || '').replace(/^#+\s*/, '');
    }
  }

  const variables = extractPlaceholdersFromHtml(innerContentHtml);
  const previewText = previewSource.replace(/\s+/g, ' ').trim().slice(0, 110);
  return { innerContentHtml, previewText, variables };
}

/**
 * Extracts live contract clauses from the open A4 editor DOM and merges them
 * with any disabled (`enabled: false`) clauses in `prevOutline` (Single Source of Truth),
 * without mutating the DOM during user typing!
 */
export function extractLiveContractClausesFromDOM(
  rootEl: HTMLElement | null,
  prevOutline: ContractOutlineClause[] = []
): ContractOutlineClause[] {
  if (!rootEl) return prevOutline;
  const segments = collectDomClauseSegments(rootEl);

  const disabledClauses = prevOutline.filter((c) => c.enabled === false);

  // Build updated active clauses from DOM
  const activeClauses: ContractOutlineClause[] = segments.map((seg, i) => {
    const prevMatch = prevOutline.find((c) => c.id === seg.id);
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
    };
  });

  if (disabledClauses.length === 0) {
    return activeClauses
      .sort((a, b) => a.order - b.order)
      .map((c, idx) => ({ ...c, index: idx + 1, order: idx }));
  }

  // Merge active and disabled clauses by their `order`
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

/**
 * Toggles a clause's `enabled` state:
 * - When disabled (`enabled = false`): captures latest `innerHTML`, removes `div.clause-container` from DOM immediately, preserves `order` and `contentHtml` in `outlineClauses`.
 * - When re-enabled (`enabled = true`): creates `div.clause-container` with `data-clause-order` and inserts it via `insertBefore` right before the first active `.clause-container` whose `order > clause.order`.
 */
export function toggleContractClauseEnabledInDOM(
  rootEl: HTMLElement | null,
  currentOutline: ContractOutlineClause[],
  clauseId: string,
  fieldValues: Record<string, string> = {}
): ContractOutlineClause[] {
  if (!rootEl) return currentOutline;

  // Ensure any raw `#` segments are upgraded before toggling
  upgradeTopLevelHashSegmentsToContainers(rootEl, currentOutline);
  const syncedOutline = extractLiveContractClausesFromDOM(rootEl, currentOutline);
  const targetIdx = syncedOutline.findIndex((c) => c.id === clauseId);
  if (targetIdx < 0) return syncedOutline;

  const targetClause = syncedOutline[targetIdx];
  const nextEnabled = !targetClause.enabled;

  if (!nextEnabled) {
    // Disabling: read latest innerHTML from DOM element and remove it from A4 sheet
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
    // Re-enabling: create .clause-container and insert deterministically by `order`
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

    // Deterministic insertion by `order`: find the first active .clause-container with order > targetClause.order
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
      // Check if last child is an empty trailing spacer paragraph
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

    // Recalculate domIndex while preserving `order`
    const updatedOutline = syncedOutline.map((c, i) =>
      i === targetIdx ? { ...c, enabled: true, isContainer: true } : c
    );
    return extractLiveContractClausesFromDOM(rootEl, updatedOutline);
  }
}

/**
 * Toggles a clause's `locked` state (`contentEditable="false"`) inside the A4 editor DOM.
 */
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

/**
 * Reorders clauses via Drag & Drop (`sourceClauseId` -> `targetClauseId`),
 * updates `order` (0, 1, 2...) in `outlineClauses`, and reorders active `.clause-container`
 * elements in-place inside `rootEl`.
 */
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

  // Reorder active .clause-container elements in-place in rootEl according to `normalized`
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
    // Find the insertion anchor (position of the first clause container in rootEl)
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

/**
 * Scrolls smoothly to a specific clause inside the A4 editor by `clauseId` (or `domIndex` fallback) and highlights it.
 */
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

/**
 * Moves a live clause up or down using the Single Source of Truth reorder engine.
 */
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

/**
 * Deletes a live clause (container or `#` heading + its body paragraphs) from the open contract DOM.
 */
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

/**
 * Renames a clause heading directly inside the open contract DOM.
 */
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

/**
 * Phase 2: Candidate Clause Segmentation for Imported Word (.docx) Documents
 * Automatically detects candidate clauses from `#` lines or short bold heading paragraphs
 * so the notary can manually review, rename, merge, or split them in DocxImportPreviewModal.
 */
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
    const hasBoldTag =
      child.style.fontWeight === 'bold' ||
      !!child.querySelector('strong, b, span[style*="bold"]');
    const isShortHeading =
      plain.startsWith('#') ||
      (plain.length >= 3 &&
        plain.length <= 75 &&
        !plain.includes('[') &&
        (notaryHeadingPatterns.test(plain) ||
          (hasBoldTag && plain.endsWith(':'))));

    if (isShortHeading) {
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

/**
 * Compiles reviewed `CandidateImportedClause[]` into `<div class="clause-container">` blocks
 * ready for the Single Source of Truth editor.
 */
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
      const hasHeadingLine =
        firstParaPlain.startsWith('#') ||
        firstParaPlain.replace(/:$/, '').trim() === cleanTitle;

      const headingHtml = hasHeadingLine
        ? ''
        : `<p dir="rtl" style="${pStyle}"><span style="font-weight:bold;"># ${escapeHtml(
            cleanTitle
          )}</span></p>`;

      return `<div class="clause-container" data-clause-id="clause_imp_${Date.now()}_${idx}" data-clause-order="${idx}" data-clause-title="${escapeHtml(
        cleanTitle
      )}" data-clause-locked="false">${headingHtml}${c.paragraphsHtml.join('')}</div>`;
    })
    .join('');
}

/**
 * Phase 4: Automatic Variable Binding from Structured Party Cards (`role` + `index`)
 * Maps each party card's fields into `fieldValues` under all standard Algerian notary variable conventions:
 * - Indexed with definite article: `البائع_1_الاسم`, `المشتري_2_الاسم`
 * - Indexed without definite article: `بائع_1_الاسم`, `مشتري_2_الاسم`
 * - For index === 1: `البائع_الاسم`, `البائع.الاسم`, and `الطرف_الأول_*` / `الطرف_الثاني_*`
 */
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
    const defRole = withDefiniteArticle(rawRole); // e.g., البائع
    const bareRole = withoutDefiniteArticle(rawRole); // e.g., بائع
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
        // 1. Indexed role keys: البائع_1_الاسم, بائع_1_الاسم, البائع.1.الاسم
        out[`${defRole}_${idx}_${suffix}`] = trimmedVal;
        out[`${bareRole}_${idx}_${suffix}`] = trimmedVal;
        out[`${defRole}.${idx}.${suffix}`] = trimmedVal;

        // 2. If index === 1, also populate non-indexed role keys: البائع_الاسم, البائع.الاسم
        if (idx === 1) {
          out[`${defRole}_${suffix}`] = trimmedVal;
          out[`${bareRole}_${suffix}`] = trimmedVal;
          out[`${defRole}.${suffix}`] = trimmedVal;
          out[`${bareRole}.${suffix}`] = trimmedVal;
        }
      }
    }

    // 3. Also map index === 1 first/second party roles to standard الطرف_الأول_* / الطرف_الثاني_*
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

/**
 * Merges structured party cards into `fieldValues` while also matching any extracted placeholders
 * in the active contract (supporting both underscore and dot notation).
 */
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
 * Counts words and estimated A4 pages from HTML string for the Bottom Status Bar.
 */
export function computeDocumentMetrics(html: string): {
  wordCount: number;
  charCount: number;
  estimatedPages: number;
} {
  const plain = html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!plain) {
    return { wordCount: 0, charCount: 0, estimatedPages: 1 };
  }
  const words = plain.split(' ').filter(Boolean);
  const pageBreakMatches = (html.match(/data-page-break="true"/g) || []).length;
  const estimatedPages = Math.max(1, Math.ceil(words.length / 320) + pageBreakMatches);
  return {
    wordCount: words.length,
    charCount: plain.length,
    estimatedPages,
  };
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
