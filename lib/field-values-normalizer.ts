/**
 * Central Notarial Field Values Normalizer & Resolver.
 *
 * Provides a single, unified entry point to normalize, synchronize, and pre-evaluate
 * all contract field values across the three supported tiers:
 * 1. Canonical Arabic Keys (e.g., [رقم_الفهرس], [الطرف_الأول_الاسم], [الثمن_بالأرقام])
 * 2. Standard Notarial Symbols (e.g., ${num_acte}, ${c_1}, ${c_2}, ${act_value}, ${pages2lettre_acte})
 * 3. Standard Latin Tags (e.g., [num_repertoire], [partie1_nom], [prix_chiffres], [prix_lettres])
 *
 * Also pre-computes legal Tafqeet for currency, dates, and page counts so that
 * storage, indexing, and Word export always have complete, coherent metadata.
 */

import {
  evaluateDynamicNotarySymbols,
  SYMBOL_KEY_TO_ARABIC_KEYS_MAP,
} from './symbol-map';
import {
  expandFieldValuesWithStandardLatinTags,
  LATIN_TO_ARABIC_KEYS_MAP,
  STANDARD_NOTARY_TAG_DEFINITIONS,
} from './notary-word-variables';
import {
  currencyToAlgerianDinarWords,
  dateToAlgerianLegalWords,
  pagesCountToAlgerianWords,
} from './tafqeet';

export interface FieldValuesNormalizerOptions {
  contractBodyHtml?: string;
  pageCount?: number;
}

/**
 * Normalizes all contract field values across Arabic keys, standard symbols, and Latin tags.
 * Ensures bidirectional consistency and evaluates legal wording (Tafqeet).
 */
export function normalizeAllNotaryFieldValues(
  fieldValues: Record<string, string> = {},
  options?: FieldValuesNormalizerOptions
): Record<string, string> {
  const result: Record<string, string> = {};

  // 1. Initial copy and cleanup of existing non-empty values
  for (const [k, v] of Object.entries(fieldValues || {})) {
    if (v !== undefined && v !== null && v.trim() !== '') {
      result[k] = v.trim();
      const normUnder = k.replace(/[\s.]+/g, '_');
      if (!result[normUnder]) {
        result[normUnder] = v.trim();
      }
    }
  }

  // 2. Reverse resolution: Latin tag -> Arabic keys
  for (const [latinKey, arKeys] of Object.entries(LATIN_TO_ARABIC_KEYS_MAP)) {
    const latinVal = result[latinKey] || result[`[${latinKey}]`];
    if (latinVal) {
      for (const arKey of arKeys) {
        if (!result[arKey]) {
          result[arKey] = latinVal;
        }
        const arUnder = arKey.replace(/[\s.]+/g, '_');
        if (!result[arUnder]) {
          result[arUnder] = latinVal;
        }
      }
    }
  }

  // 3. Reverse resolution: Symbol key -> Arabic keys (e.g., num_acte -> رقم_الفهرس)
  for (const [symKey, arKeys] of Object.entries(SYMBOL_KEY_TO_ARABIC_KEYS_MAP)) {
    const symVal = result[symKey] || result[`\${${symKey}}`];
    if (symVal) {
      for (const arKey of arKeys) {
        if (!result[arKey]) {
          result[arKey] = symVal;
        }
        const arUnder = arKey.replace(/[\s.]+/g, '_');
        if (!result[arUnder]) {
          result[arUnder] = symVal;
        }
      }
    }
  }

  // 4. Resolve Dynamic Notary Symbols and Tafqeet
  const pageCount = options?.pageCount !== undefined && options.pageCount > 0
    ? options.pageCount
    : (result['pages_acte'] ? parseInt(result['pages_acte'], 10) || 1 : 1);

  const evaluatedSymbols = evaluateDynamicNotarySymbols(result, {
    contractBodyHtml: options?.contractBodyHtml,
    pageCount,
  });

  for (const [symKey, symVal] of Object.entries(evaluatedSymbols)) {
    if (symVal !== undefined && symVal.trim() !== '') {
      result[symKey] = symVal.trim();
      result[`\${${symKey}}`] = symVal.trim();
    }
  }

  // 5. Expand and synchronize Standard Latin Tags
  const withLatin = expandFieldValuesWithStandardLatinTags(
    result,
    options?.contractBodyHtml
  );

  for (const [latinKey, latinVal] of Object.entries(withLatin)) {
    if (latinVal !== undefined && latinVal.trim() !== '') {
      result[latinKey] = latinVal.trim();
    }
  }

  // 6. Explicit Financial and Tafqeet Synchronization
  const priceVal =
    result['الثمن_بالأرقام'] ||
    result['الثمن'] ||
    result['prix_chiffres'] ||
    result['act_value'] ||
    result['total'] ||
    '';

  if (priceVal) {
    const priceLetters = currencyToAlgerianDinarWords(priceVal);
    if (!result['الثمن_بالأحرف']) result['الثمن_بالأحرف'] = priceLetters;
    if (!result['الثمن_بالحروف']) result['الثمن_بالحروف'] = priceLetters;
    if (!result['prix_lettres']) result['prix_lettres'] = priceLetters;
    if (!result['totalLettre']) result['totalLettre'] = priceLetters;
    if (!result['totalLettre_ar']) result['totalLettre_ar'] = priceLetters;
    if (!result['act_value_lettre']) result['act_value_lettre'] = priceLetters;
  }

  // 7. Explicit Date Tafqeet Synchronization
  const dateVal =
    result['تاريخ_العقد'] ||
    result['date_acte'] ||
    '';

  if (dateVal) {
    const dateWords = dateToAlgerianLegalWords(dateVal);
    if (dateWords) {
      if (!result['تاريخ_العقد_بالحروف']) result['تاريخ_العقد_بالحروف'] = dateWords.fullWithNumbers;
      if (!result['date_acte_lettres']) result['date_acte_lettres'] = dateWords.fullWithNumbers;
      if (!result['date2lettre_acte']) result['date2lettre_acte'] = dateWords.fullWithNumbers;
      if (!result['date2lettre_acte_only']) result['date2lettre_acte_only'] = dateWords.lettersOnly;
      if (!result['anne2lettre_acte']) result['anne2lettre_acte'] = dateWords.yearLettersOnly;
      if (!result['jourmois2lettre_acte']) result['jourmois2lettre_acte'] = dateWords.dayMonthOnly;
    }
  }

  // 8. Explicit Page Count Tafqeet Synchronization
  result['pages_acte'] = String(pageCount);
  result['pages2lettre_acte'] = pagesCountToAlgerianWords(pageCount);

  // 9. Synchronize Parties shortcuts (c_1, c_2, client_name)
  const p1 =
    result['الطرف_الأول_الاسم'] ||
    result['الطرف_الأول'] ||
    result['partie1_nom_complet'] ||
    result['partie1_nom'] ||
    '';
  if (p1) {
    if (!result['c_1']) result['c_1'] = p1;
    if (!result['partie1_nom']) result['partie1_nom'] = p1;
  }

  const p2 =
    result['الطرف_الثاني_الاسم'] ||
    result['الطرف_الثاني'] ||
    result['partie2_nom_complet'] ||
    result['partie2_nom'] ||
    '';
  if (p2) {
    if (!result['c_2']) result['c_2'] = p2;
    if (!result['partie2_nom']) result['partie2_nom'] = p2;
    if (!result['client_name']) result['client_name'] = p2;
  }

  // 10. Contract body text injection into clause tags (REMOVED to prevent escaping as a single block)
  // if (options?.contractBodyHtml) {
  //   result['clause'] = options.contractBodyHtml;
  //   result['contenu_acte'] = options.contractBodyHtml;
  //   result['corps_acte'] = options.contractBodyHtml;
  //   result['texte_acte'] = options.contractBodyHtml;
  // }

  return result;
}
