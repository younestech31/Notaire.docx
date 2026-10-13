/**
 * Notarial Body Purifier (محرك تطهير وتنظيم متن العقد التوثيقي).
 *
 * Enforces dense, professional notarial formatting:
 * - Eliminates empty paragraphs and <div> shells without visible text.
 * - Strips old non-notary page breaks and pagination markers.
 * - Preserves tables, images, smart tags ([var]), and clause headings.
 * - Enforces strict inline typography (Arial 13pt, line-height 1, margin 0).
 */

import { stripNotarialPaginationFromHtml } from './notarial-pagination';

export interface PurifyOptions {
  keepSingleEmptyShell?: boolean;
}

export function purifyNotarialBodyHtml(html: string, options?: PurifyOptions): string {
  if (!html || typeof html !== 'string') {
    return options?.keepSingleEmptyShell
      ? '<p dir="rtl" style="margin:0; line-height:1; font-family:Arial; font-size:13pt;"><br></p>'
      : '';
  }

  // 1. Strip existing notarial pagination markers first
  let cleaned = stripNotarialPaginationFromHtml(html);

  if (typeof window === 'undefined') {
    // Server-side fallback regex cleanup
    cleaned = cleaned
      .replace(/<p[^>]*>\s*(?:<br\s*\/?>|&nbsp;|\s)*<\/p>/gi, '')
      .replace(/<div[^>]*>\s*(?:<br\s*\/?>|&nbsp;|\s)*<\/div>/gi, '');
    if (!cleaned.trim() && options?.keepSingleEmptyShell) {
      return '<p dir="rtl" style="margin:0; line-height:1; font-family:Arial; font-size:13pt;"><br></p>';
    }
    return cleaned;
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${cleaned}</div>`, 'text/html');
  const container = doc.body.firstElementChild as HTMLElement;
  if (!container) {
    return options?.keepSingleEmptyShell
      ? '<p dir="rtl" style="margin:0; line-height:1; font-family:Arial; font-size:13pt;"><br></p>'
      : '';
  }

  // Remove non-notary page breaks
  const pageBreaks = container.querySelectorAll('.page-break, [data-page-break="true"]');
  pageBreaks.forEach((pb) => {
    if (!pb.hasAttribute('data-notary-pagination')) {
      pb.remove();
    }
  });

  // Process all children blocks
  const children = Array.from(container.children);
  for (const child of children) {
    const tag = child.tagName.toUpperCase();
    if (tag === 'TABLE' || tag === 'FIGURE' || tag === 'IMG') {
      continue;
    }

    const text = (child.textContent || '').replace(/\u00a0/g, ' ').trim();
    const hasSpecial = child.querySelector('img, table, .smart-tag');

    if (!text && !hasSpecial) {
      child.remove();
      continue;
    }

    if (tag === 'P' || tag === 'DIV' || tag.startsWith('H')) {
      child.setAttribute('dir', 'rtl');
      const existingStyle = child.getAttribute('style') || '';
      const cleanedStyle = existingStyle
        .replace(/margin:[^;]+(;|$)/gi, '')
        .replace(/line-height:[^;]+(;|$)/gi, '')
        .replace(/font-family:[^;]+(;|$)/gi, '')
        .replace(/font-size:[^;]+(;|$)/gi, '');
      child.setAttribute(
        'style',
        `margin:0; line-height:1; ${cleanedStyle}`.trim()
      );
    }
  }

  const finalHtml = container.innerHTML.trim();

  if (!finalHtml && options?.keepSingleEmptyShell) {
    return '<p dir="rtl" style="margin:0; line-height:1; font-family:Arial; font-size:13pt;"><br></p>';
  }

  return finalHtml;
}
