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
    .replace(
      /<(?:p|div)\b[^>]*\bdata-page-break=["']true["'][^>]*>[\s\S]*?<\/(?:p|div)>/gi,
      ''
    )
    .replace(/<div\b[^>]*\bclass=["'][^"']*page-break[^"']*["'][^>]*>[\s\S]*?<\/div>/gi, '')
    .trim();
}
