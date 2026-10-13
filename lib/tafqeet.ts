/**
 * Algerian Notarial Legal Arabic Tafqeet Engine (محرك التفقيط القانوني الجزائري).
 * 
 * Provides spelling-out functions in standard Algerian notarial Arabic for:
 * 1. Monetary amounts in Algerian Dinars (DZD)
 * 2. Full legal deed dates (اليوم والشهر والسنة بالأحرف)
 * 3. Deed indexing years (سنة الفهرسة)
 * 4. Contract page numbers count (عدد الصفحات)
 */

const ONES = [
  '',
  'واحد',
  'اثنان',
  'ثلاثة',
  'أربعة',
  'خمسة',
  'ستة',
  'سبعة',
  'ثمانية',
  'تسعة',
  'عشرة',
  'أحد عشر',
  'اثنا عشر',
  'ثلاثة عشر',
  'أربعة عشر',
  'خمسة عشر',
  'ستة عشر',
  'سبعة عشر',
  'ثمانية عشر',
  'تسعة عشر',
];

const TENS = [
  '',
  'عشرة',
  'عشرون',
  'ثلاثون',
  'أربعون',
  'خمسون',
  'ستون',
  'سبعون',
  'ثمانون',
  'تسعون',
];

const HUNDREDS = [
  '',
  'مائة',
  'مائتان',
  'ثلاثمائة',
  'أربعمائة',
  'خمسمائة',
  'ستمائة',
  'سبعمائة',
  'ثمانمائة',
  'تسعمائة',
];

const ARABIC_MONTHS = [
  'جانفي',
  'فيفري',
  'مارس',
  'أفريل',
  'ماي',
  'جوان',
  'جويلية',
  'أوت',
  'سبتمبر',
  'أكتوبر',
  'نوفمبر',
  'ديسمبر',
];

const DAY_ORDINALS: Record<number, string> = {
  1: 'الأول',
  2: 'الثاني',
  3: 'الثالث',
  4: 'الرابع',
  5: 'الخامس',
  6: 'السادس',
  7: 'السابع',
  8: 'الثامن',
  9: 'التاسع',
  10: 'العاشر',
  11: 'الحادي عشر',
  12: 'الثاني عشر',
  13: 'الثالث عشر',
  14: 'الرابع عشر',
  15: 'الخامس عشر',
  16: 'السادس عشر',
  17: 'السابع عشر',
  18: 'الثامن عشر',
  19: 'التاسع عشر',
  20: 'العشرين',
  21: 'الحادي والعشرين',
  22: 'الثاني والعشرين',
  23: 'الثالث والعشرين',
  24: 'الرابع والعشرين',
  25: 'الخامس والعشرين',
  26: 'السادس والعشرين',
  27: 'السابع والعشرين',
  28: 'الثامن والعشرين',
  29: 'التاسع والعشرين',
  30: 'الثلاثين',
  31: 'الحادي والثلاثين',
};

/**
 * Converts a positive number (0 to 999) to Arabic words.
 */
function convertHundredsGroup(n: number): string {
  if (n === 0) return '';
  const parts: string[] = [];

  const h = Math.floor(n / 100);
  const rem = n % 100;

  if (h > 0) {
    parts.push(HUNDREDS[h]);
  }

  if (rem > 0) {
    if (rem < 20) {
      parts.push(ONES[rem]);
    } else {
      const o = rem % 10;
      const t = Math.floor(rem / 10);
      if (o > 0) {
        parts.push(`${ONES[o]} و${TENS[t]}`);
      } else {
        parts.push(TENS[t]);
      }
    }
  }

  return parts.join(' و');
}

/**
 * Converts any positive integer into standard legal Arabic words.
 */
export function integerToAlgerianWords(num: number): string {
  if (num === 0) return 'صفر';
  if (num < 0) return `سالب ${integerToAlgerianWords(Math.abs(num))}`;

  const billions = Math.floor(num / 1_000_000_000);
  const millions = Math.floor((num % 1_000_000_000) / 1_000_000);
  const thousands = Math.floor((num % 1_000_000) / 1_000);
  const remainder = num % 1_000;

  const chunks: string[] = [];

  if (billions > 0) {
    if (billions === 1) chunks.push('مليار');
    else if (billions === 2) chunks.push('ملياران');
    else if (billions >= 3 && billions <= 10) chunks.push(`${convertHundredsGroup(billions)} ملايير`);
    else chunks.push(`${convertHundredsGroup(billions)} مليار`);
  }

  if (millions > 0) {
    if (millions === 1) chunks.push('مليون');
    else if (millions === 2) chunks.push('مليونان');
    else if (millions >= 3 && millions <= 10) chunks.push(`${convertHundredsGroup(millions)} ملايين`);
    else chunks.push(`${convertHundredsGroup(millions)} مليون`);
  }

  if (thousands > 0) {
    if (thousands === 1) chunks.push('ألف');
    else if (thousands === 2) chunks.push('ألفان');
    else if (thousands >= 3 && thousands <= 10) chunks.push(`${convertHundredsGroup(thousands)} آلاف`);
    else chunks.push(`${convertHundredsGroup(thousands)} ألف`);
  }

  if (remainder > 0) {
    chunks.push(convertHundredsGroup(remainder));
  }

  return chunks.join(' و');
}

/**
 * Converts a monetary string or number into Algerian Dinar legal words (دينار جزائري).
 */
export function currencyToAlgerianDinarWords(raw: string | number): string {
  const clean = String(raw ?? '').replace(/[^\d.]/g, '');
  if (!clean) return '';
  const num = parseFloat(clean);
  if (Number.isNaN(num) || num <= 0) return 'صفر دينار جزائري';

  const dinars = Math.floor(num);
  const cents = Math.round((num - dinars) * 100);

  let result = `${integerToAlgerianWords(dinars)} دينار جزائري`;
  if (cents > 0) {
    result += ` و${integerToAlgerianWords(cents)} سنتيم`;
  }
  return `${result} فقط لا غير`;
}

/**
 * Converts a date into full Algerian notarial wording:
 * e.g. "2026/02/15" -> "الخامس عشر من شهر فيفري سنة ألفين وستة وعشرين (15/02/2026)"
 */
export function dateToAlgerianLegalWords(rawDate: string): {
  fullWithNumbers: string;
  lettersOnly: string;
  yearLettersOnly: string;
  dayMonthOnly: string;
} {
  const parts = String(rawDate || '').split(/[-/.]/);
  let year = new Date().getFullYear();
  let month = 1;
  let day = 1;

  if (parts.length === 3) {
    if (parts[0].length === 4) {
      year = parseInt(parts[0], 10) || year;
      month = parseInt(parts[1], 10) || 1;
      day = parseInt(parts[2], 10) || 1;
    } else {
      day = parseInt(parts[0], 10) || 1;
      month = parseInt(parts[1], 10) || 1;
      year = parseInt(parts[2], 10) || year;
    }
  }

  const dayStr = DAY_ORDINALS[day] || integerToAlgerianWords(day);
  const monthStr = ARABIC_MONTHS[Math.min(Math.max(month - 1, 0), 11)];
  const yearStr = integerToAlgerianWords(year);

  const lettersOnly = `في يوم ${dayStr} من شهر ${monthStr} سنة ${yearStr}`;
  const dayFormatted = String(day).padStart(2, '0');
  const monthFormatted = String(month).padStart(2, '0');
  const fullWithNumbers = `${lettersOnly} الموافق لـ (${dayFormatted}/${monthFormatted}/${year})`;
  const yearLettersOnly = `سنة ${yearStr} (${year})`;
  const dayMonthOnly = `يوم ${dayStr} من شهر ${monthStr} (${dayFormatted}/${monthFormatted})`;

  return {
    fullWithNumbers,
    lettersOnly,
    yearLettersOnly,
    dayMonthOnly,
  };
}

/**
 * Converts page count to notarial text:
 * e.g. 3 -> "ثلاث صفحات (03)"
 */
export function pagesCountToAlgerianWords(pages: number): string {
  const n = Math.max(1, pages);
  const padded = String(n).padStart(2, '0');
  if (n === 1) return `صفحة واحدة (${padded})`;
  if (n === 2) return `صفحتان (${padded})`;
  if (n >= 3 && n <= 10) return `${integerToAlgerianWords(n)} صفحات (${padded})`;
  return `${integerToAlgerianWords(n)} صفحة (${padded})`;
}
