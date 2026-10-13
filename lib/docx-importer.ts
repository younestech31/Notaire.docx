import JSZip from 'jszip';
import { STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';
import { normalizePlaceholderKey } from './docx-engine';
import { purifyNotarialBodyHtml } from './notarial-body-purifier';

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

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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
        const cleanPath = target.startsWith('/')
          ? target.slice(1)
          : `word/${target.replace(/^\.\.\//, '')}`;
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
  const bodyNode =
    xmlDoc.getElementsByTagName('w:body')[0] || xmlDoc.getElementsByTagName('body')[0];

  const bodyHtml = bodyNode ? convertXmlContainerToHtml(bodyNode, mediaDataUrlMap) : '';

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

  const normalizeBracketPlaceholdersInHtml = (htmlStr: string): string => {
    if (!htmlStr) return '';
    return htmlStr.replace(/(>[^<]*)|(<[^>]+>)/g, (segment) => {
      if (segment.startsWith('<')) return segment;
      return segment.replace(
        /\[\s*([^\[\]<>]{1,45}?)\s*\]|\{\{\s*([^}<>]{1,45}?)\s*\}\}/g,
        (fullMatch, bInner, cInner) => {
          const inner = bInner || cInner || '';
          const cleanKey = normalizePlaceholderKey(inner).replace(/\s+/g, '_');
          if (!cleanKey || /^\d+$/.test(cleanKey)) return fullMatch;
          return `<span class="smart-tag" data-var="${escapeXml(cleanKey)}">[${escapeXml(
            cleanKey
          )}]</span>`;
        }
      );
    });
  };

  const normalizedBodyHtml = normalizeBracketPlaceholdersInHtml(bodyHtml);
  const normalizedHeaderHtml = normalizeBracketPlaceholdersInHtml(headerHtml);
  const normalizedFooterHtml = normalizeBracketPlaceholdersInHtml(footerHtml);

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
    bodyHtml: purifyNotarialBodyHtml(normalizedBodyHtml, { keepSingleEmptyShell: true }),
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

  const trimmedInner = innerHtml.trim();
  if (!trimmedInner && !hasInlinePageBreak && !hasPageBreakBefore && !pNode.querySelector('img, table')) {
    return '';
  }

  let prefixBreak = '';
  if (hasPageBreakBefore) {
    prefixBreak =
      '<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div>';
  }

  if (hasInlinePageBreak && trimmedInner === '') {
    return `${prefixBreak}<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div>`;
  }

  const suffixBreak = hasInlinePageBreak
    ? '<div class="page-break" contenteditable="false" data-page-break="true">فاصل صفحات</div>'
    : '';

  const content = trimmedInner || '<br>';
  return `${prefixBreak}<p dir="${
    isRtl ? 'rtl' : 'ltr'
  }" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:${align};">${content}</p>${suffixBreak}`;
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

      tableHtml += `<td ${attrs.join(
        ' '
      )} style="${borderCss}${bgCss}padding:2px 4px;vertical-align:top;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;line-height:1;">${
        cell.contentHtml
      }</td>`;
    }
    tableHtml += '</tr>';
  }

  tableHtml += '</tbody></table>';
  return tableHtml;
}
