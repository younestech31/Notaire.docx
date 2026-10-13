/**
 * Canonical Notarial Symbol System (دليل الرموز التوثيقية الرسمية المعتمد).
 * 
 * Provides full bidirectional mapping and automatic value evaluation for
 * standard official symbols used in Algerian notarial Word templates (${key}).
 */

import {
  currencyToAlgerianDinarWords,
  dateToAlgerianLegalWords,
  integerToAlgerianWords,
  pagesCountToAlgerianWords,
} from './tafqeet';

export type NotarySymbolCategory =
  | 'office'
  | 'acte'
  | 'tafqeet'
  | 'registration_publication'
  | 'payment_deposit'
  | 'fiscal'
  | 'block_loop';

export interface NotarySymbolItem {
  symbol: string; // e.g. "${num_acte}"
  key: string; // e.g. "num_acte"
  label: string; // Arabic explanation
  category: NotarySymbolCategory;
  categoryLabel: string;
  sampleValue: string;
  mappedFieldKeys: string[];
}

export const CANONICAL_NOTARY_SYMBOLS: NotarySymbolItem[] = [
  // 1. حلقة تكرار البنود (Block Loop)
  {
    symbol: '${block}...${/block}',
    key: 'block',
    label: 'حلقة تكرار بنود العقد (بداية ونهاية)',
    category: 'block_loop',
    categoryLabel: 'حلقة البنود التكرارية',
    sampleValue: 'تكرار لكل بند نشط بالعقد',
    mappedFieldKeys: ['متن_العقد', 'بنود_العقد'],
  },
  {
    symbol: '${titre}',
    key: 'titre',
    label: 'عنوان البند داخل حلقة التكرار',
    category: 'block_loop',
    categoryLabel: 'حلقة البنود التكرارية',
    sampleValue: 'تعيين الحصة المفرزة والمساحة',
    mappedFieldKeys: ['عنوان_البند'],
  },
  {
    symbol: '${clause}',
    key: 'clause',
    label: 'محتوى البند بالكامل داخل حلقة التكرار',
    category: 'block_loop',
    categoryLabel: 'حلقة البنود التكرارية',
    sampleValue: 'نص البند التوثيقي كاملاً مع التنسيق',
    mappedFieldKeys: ['محتوى_البند'],
  },

  // 2. بيانات مكتب التوثيق (Office)
  {
    symbol: '${office_name}',
    key: 'office_name',
    label: 'اسم مكتب التوثيق بالعربية',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'مكتب الأستاذ عبد الرحمن بوزيد',
    mappedFieldKeys: ['المكتب', 'اسم_المكتب', 'الموثق'],
  },
  {
    symbol: '${office_name_fr}',
    key: 'office_name_fr',
    label: 'اسم مكتب التوثيق بالفرنسية',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'Étude Notariale Me Abdelrahmane Bouzid',
    mappedFieldKeys: ['المكتب_فر', 'اسم_المكتب_فر'],
  },
  {
    symbol: '${office_addr}',
    key: 'office_addr',
    label: 'عنوان المكتب بالعربية',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'حي ديدوش مراد، الجزائر الوسطى',
    mappedFieldKeys: ['عنوان_المكتب', 'مقر_المكتب'],
  },
  {
    symbol: '${office_addr_fr}',
    key: 'office_addr_fr',
    label: 'عنوان المكتب بالفرنسية',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'Rue Didouche Mourad, Alger Centre',
    mappedFieldKeys: ['عنوان_المكتب_فر'],
  },
  {
    symbol: '${office_wilaya}',
    key: 'office_wilaya',
    label: 'ولاية مكتب التوثيق',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'الجزائر',
    mappedFieldKeys: ['ولاية_المكتب', 'الولاية'],
  },
  {
    symbol: '${office_tasjil}',
    key: 'office_tasjil',
    label: 'مصلحة / مفتشية التسجيل المختصة للمكتب',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'مفتشية التسجيل الجزائر الوسطى',
    mappedFieldKeys: ['مفتشية_التسجيل'],
  },
  {
    symbol: '${office_tel}',
    key: 'office_tel',
    label: 'رقم هاتف المكتب',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: '021 63 45 78',
    mappedFieldKeys: ['هاتف_المكتب'],
  },
  {
    symbol: '${office_fax}',
    key: 'office_fax',
    label: 'فاكس المكتب',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: '021 63 45 79',
    mappedFieldKeys: ['فاكس_المكتب'],
  },
  {
    symbol: '${office_mail}',
    key: 'office_mail',
    label: 'البريد الإلكتروني للمكتب',
    category: 'office',
    categoryLabel: 'بيانات المكتب التوثيقي',
    sampleValue: 'etude.notaire@dz.legal',
    mappedFieldKeys: ['بريد_المكتب'],
  },

  // 3. بيانات العقد والأطراف (Acte & Parties)
  {
    symbol: '${type_acte}',
    key: 'type_acte',
    label: 'نوع العقد (عقد بيع، هبة، قسمة...)',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: 'عقد بيع عقاري رسمي',
    mappedFieldKeys: ['نوع_العقد', 'طبيعة_العقد'],
  },
  {
    symbol: '${date_acte}',
    key: 'date_acte',
    label: 'تاريخ العقد (أرقام)',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: '2026/02/15',
    mappedFieldKeys: ['تاريخ_العقد', 'تاريخ_اليوم'],
  },
  {
    symbol: '${num_acte}',
    key: 'num_acte',
    label: 'رقم فهرس العقد الرسمي',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: '2026/0142',
    mappedFieldKeys: ['رقم_الفهرس', 'الفهرس', 'رقم_العقد'],
  },
  {
    symbol: '${c_1}',
    key: 'c_1',
    label: 'الطرف الأول (البائع / المتصرف)',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: 'السيد بن عيسى أحمد بن مصطفى',
    mappedFieldKeys: ['الطرف_الأول_الاسم', 'الطرف_الأول'],
  },
  {
    symbol: '${c_2}',
    key: 'c_2',
    label: 'الطرف الثاني (المشتري / المتصرف إليه)',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: 'السيد براهيمي كريم بن عبد الله',
    mappedFieldKeys: ['الطرف_الثاني_الاسم', 'الطرف_الثاني'],
  },
  {
    symbol: '${act_value}',
    key: 'act_value',
    label: 'قيمة العقد المصرحة (الثمن بالأرقام)',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: '15,000,000.00',
    mappedFieldKeys: ['الثمن_بالأرقام', 'الثمن', 'قيمة_العقد'],
  },
  {
    symbol: '${today_date}',
    key: 'today_date',
    label: 'تاريخ اليوم الحالي لتوليد الوثيقة',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: '2026/02/15',
    mappedFieldKeys: ['تاريخ_التوليد', 'تاريخ_اليوم'],
  },
  {
    symbol: '${today_year}',
    key: 'today_year',
    label: 'سنة التوليد الحالية',
    category: 'acte',
    categoryLabel: 'بيانات العقد والأطراف',
    sampleValue: '2026',
    mappedFieldKeys: ['سنة_التوليد', 'السنة'],
  },

  // 4. التفقيط القانوني (Tafqeet)
  {
    symbol: '${pages_acte}',
    key: 'pages_acte',
    label: 'عدد صفحات العقد (أرقام)',
    category: 'tafqeet',
    categoryLabel: 'التفقيط والصياغة بالحروف',
    sampleValue: '3',
    mappedFieldKeys: ['عدد_الصفحات'],
  },
  {
    symbol: '${pages2lettre_acte}',
    key: 'pages2lettre_acte',
    label: 'عدد الصفحات (حروف + رقم)',
    category: 'tafqeet',
    categoryLabel: 'التفقيط والصياغة بالحروف',
    sampleValue: 'ثلاث صفحات (03)',
    mappedFieldKeys: ['عدد_الصفحات_بالحروف'],
  },
  {
    symbol: '${date2lettre_acte}',
    key: 'date2lettre_acte',
    label: 'تاريخ الفهرسة (حروف + أرقام)',
    category: 'tafqeet',
    categoryLabel: 'التفقيط والصياغة بالحروف',
    sampleValue: 'في يوم الخامس عشر من شهر فيفري سنة ألفين وستة وعشرين الموافق لـ (15/02/2026)',
    mappedFieldKeys: ['تاريخ_العقد_بالأحرف', 'تاريخ_الفهرسة_حروف_وارقام'],
  },
  {
    symbol: '${date2lettre_acte_only}',
    key: 'date2lettre_acte_only',
    label: 'تاريخ الفهرسة (حروف فقط)',
    category: 'tafqeet',
    categoryLabel: 'التفقيط والصياغة بالحروف',
    sampleValue: 'في يوم الخامس عشر من شهر فيفري سنة ألفين وستة وعشرين',
    mappedFieldKeys: ['تاريخ_العقد_حروف_فقط'],
  },
  {
    symbol: '${anne2lettre_acte}',
    key: 'anne2lettre_acte',
    label: 'سنة الفهرسة (حروف + رقم)',
    category: 'tafqeet',
    categoryLabel: 'التفقيط والصياغة بالحروف',
    sampleValue: 'سنة ألفين وستة وعشرين (2026)',
    mappedFieldKeys: ['سنة_العقد_بالحروف'],
  },
  {
    symbol: '${jourmois2lettre_acte}',
    key: 'jourmois2lettre_acte',
    label: 'يوم وشهر الفهرسة + التاريخ الكامل',
    category: 'tafqeet',
    categoryLabel: 'التفقيط والصياغة بالحروف',
    sampleValue: 'يوم الخامس عشر من شهر فيفري (15/02)',
    mappedFieldKeys: ['يوم_وشهر_الفهرسة'],
  },

  // 5. التسجيل والشهر العقاري (Registration & Publication)
  {
    symbol: '${tasjil_acte}',
    key: 'tasjil_acte',
    label: 'مراجع وتأشيرة التسجيل',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: 'سجل 14 ص 88 قسيمة 210',
    mappedFieldKeys: ['مراجع_التسجيل', 'تأشيرة_التسجيل'],
  },
  {
    symbol: '${tasjil_date}',
    key: 'tasjil_date',
    label: 'تاريخ تأشيرة التسجيل',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: '2026/02/20',
    mappedFieldKeys: ['تاريخ_التسجيل'],
  },
  {
    symbol: '${tasjil_tax}',
    key: 'tasjil_tax',
    label: 'مبلغ رسوم التسجيل (دج)',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: '750,000.00',
    mappedFieldKeys: ['حقوق_التسجيل', 'رسوم_التسجيل'],
  },
  {
    symbol: '${timbre_tax}',
    key: 'timbre_tax',
    label: 'رسوم الطابع الجبائي (دج)',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: '2,500.00',
    mappedFieldKeys: ['رسوم_الطابع', 'حقوق_الطابع'],
  },
  {
    symbol: '${pub_conserv}',
    key: 'pub_conserv',
    label: 'المحافظة العقارية المختصة',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: 'المحافظة العقارية لسيدي امحمد',
    mappedFieldKeys: ['المحافظة_العقارية'],
  },
  {
    symbol: '${pub_conserv_wilaya}',
    key: 'pub_conserv_wilaya',
    label: 'ولاية المحافظة العقارية',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: 'الجزائر',
    mappedFieldKeys: ['ولاية_المحافظة'],
  },
  {
    symbol: '${pub_tax}',
    key: 'pub_tax',
    label: 'مبلغ رسوم الشهر العقاري (دج)',
    category: 'registration_publication',
    categoryLabel: 'التسجيل والشهر العقاري',
    sampleValue: '150,000.00',
    mappedFieldKeys: ['رسم_الإشهار_العقاري', 'رسوم_الشهر'],
  },

  // 6. بطاقة الدفع وقسائم الإيداع (Payment & Deposit)
  {
    symbol: '${date}',
    key: 'date',
    label: 'تاريخ عملية الدفع / الإيداع',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '2026/02/15',
    mappedFieldKeys: ['تاريخ_الدفع', 'تاريخ_العملية', 'تاريخ_العقد'],
  },
  {
    symbol: '${client_n}',
    key: 'client_n',
    label: 'رقم حساب المودع لدى الخزينة',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '007.01234567.89',
    mappedFieldKeys: ['رقم_حساب_الخزينة', 'رقم_المودع'],
  },
  {
    symbol: '${total}',
    key: 'total',
    label: 'المبلغ الإجمالي المدفوع (أرقام)',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '15,000,000.00',
    mappedFieldKeys: ['المبلغ_الإجمالي', 'الثمن_بالأرقام', 'الثمن'],
  },
  {
    symbol: '${totalLettre}',
    key: 'totalLettre',
    label: 'المبلغ الإجمالي بالحروف',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: 'خمسة عشر مليون دينار جزائري فقط لا غير',
    mappedFieldKeys: ['الثمن_بالأحرف', 'المبلغ_بالحروف'],
  },
  {
    symbol: '${totalLettre_ar}',
    key: 'totalLettre_ar',
    label: 'المبلغ بالحروف بالعربية',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: 'خمسة عشر مليون دينار جزائري',
    mappedFieldKeys: ['الثمن_بالأحرف'],
  },
  {
    symbol: '${totalLettre_fr}',
    key: 'totalLettre_fr',
    label: 'المبلغ بالحروف بالفرنسية',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: 'Quinze millions de dinars algériens',
    mappedFieldKeys: ['المبلغ_بالفرنسية'],
  },
  {
    symbol: '${client_name}',
    key: 'client_name',
    label: 'اسم المودع / العميل',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: 'براهيمي كريم',
    mappedFieldKeys: ['اسم_المودع', 'الطرف_الثاني_الاسم'],
  },
  {
    symbol: '${Motif}',
    key: 'Motif',
    label: 'سبب الدفع أو موضوع الإيداع',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: 'تسوية ثمن شراء شقة ومصاريف التوثيق',
    mappedFieldKeys: ['سبب_الدفع', 'موضوع_الإيداع'],
  },
  {
    symbol: '${espece}',
    key: 'espece',
    label: 'المبلغ المدفوع نقداً (دج)',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '0.00',
    mappedFieldKeys: ['المبلغ_نقدا'],
  },
  {
    symbol: '${cheque}',
    key: 'cheque',
    label: 'المبلغ المدفوع بالصك (دج)',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '15,000,000.00',
    mappedFieldKeys: ['المبلغ_بالصك'],
  },
  {
    symbol: '${cheq_num}',
    key: 'cheq_num',
    label: 'رقم الصك البنكي',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '4589210',
    mappedFieldKeys: ['رقم_الصك'],
  },
  {
    symbol: '${cheq_banque}',
    key: 'cheq_banque',
    label: 'البنك المسحوب عليه الصك',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: 'بنك الجزائر الخارجي (BEA)',
    mappedFieldKeys: ['بنك_الصك'],
  },
  {
    symbol: '${cheq_date}',
    key: 'cheq_date',
    label: 'تاريخ تحرير الصك',
    category: 'payment_deposit',
    categoryLabel: 'الدفع وقسائم الإيداع',
    sampleValue: '2026/02/10',
    mappedFieldKeys: ['تاريخ_الصك'],
  },

  // 7. رموز الوضعية الجبائية (Série H n° 59)
  {
    symbol: '${fiscal_par1}',
    key: 'fiscal_par1',
    label: 'الفقرة 1: تعيين العقار والوصف التقسيمي',
    category: 'fiscal',
    categoryLabel: 'الوضعية الجبائية (Série H n° 59)',
    sampleValue: 'الشقة ذات الاستعمال السكني الكائنة بالطابق الثالث، الحصة رقم 15 بمساحة 110 م²',
    mappedFieldKeys: ['تعيين_الحصة_الكامل', 'تعيين_العقار'],
  },
  {
    symbol: '${fiscal_par2}',
    key: 'fiscal_par2',
    label: 'الفقرة 2: أصل الملكية والمشتقات العقارية',
    category: 'fiscal',
    categoryLabel: 'الوضعية الجبائية (Série H n° 59)',
    sampleValue: 'عقد بيع مشهر بالمحافظة العقارية مجلد 2100 رقم 45 بتاريخ 2018/04/12',
    mappedFieldKeys: ['أصل_الملكية', 'مراجع_الوصف_التقسيمي'],
  },
  {
    symbol: '${fiscal_par3}',
    key: 'fiscal_par3',
    label: 'الفقرة 3: هوية البائع (المتصرف)',
    category: 'fiscal',
    categoryLabel: 'الوضعية الجبائية (Série H n° 59)',
    sampleValue: 'بن عيسى أحمد، المولود بـ الجزائر، الحامل للرقم الوطني 119801601002345678',
    mappedFieldKeys: ['الطرف_الأول_الاسم'],
  },
  {
    symbol: '${fiscal_par4}',
    key: 'fiscal_par4',
    label: 'الفقرة 4: هوية المشتري (المتصرف إليه)',
    category: 'fiscal',
    categoryLabel: 'الوضعية الجبائية (Série H n° 59)',
    sampleValue: 'براهيمي كريم، المولود بـ بئر مراد رايس، الحامل للرقم الوطني 119881604009876543',
    mappedFieldKeys: ['الطرف_الثاني_الاسم'],
  },
  {
    symbol: '${fiscal_par5}',
    key: 'fiscal_par5',
    label: 'الفقرة 5: الثمن المصرح به + إيداع الخمس القانوني',
    category: 'fiscal',
    categoryLabel: 'الوضعية الجبائية (Série H n° 59)',
    sampleValue: 'الثمن 15,000,000.00 دج، إيداع الخمس القانوني: 3,000,000.00 دج',
    mappedFieldKeys: ['الثمن_بالأرقام', 'إيداع_الخمس'],
  },
];

/**
 * Fast lookup maps
 */
export const SYMBOL_KEY_TO_ARABIC_KEYS_MAP: Record<string, string[]> = {};
export const ARABIC_KEY_TO_SYMBOL_KEYS_MAP: Record<string, string[]> = {};

for (const item of CANONICAL_NOTARY_SYMBOLS) {
  SYMBOL_KEY_TO_ARABIC_KEYS_MAP[item.key.toLowerCase()] = item.mappedFieldKeys;
  for (const arKey of item.mappedFieldKeys) {
    const norm = arKey.replace(/[\s.]+/g, '_');
    if (!ARABIC_KEY_TO_SYMBOL_KEYS_MAP[norm]) {
      ARABIC_KEY_TO_SYMBOL_KEYS_MAP[norm] = [];
    }
    ARABIC_KEY_TO_SYMBOL_KEYS_MAP[norm].push(item.key);
  }
}

/**
 * Automatically computes and enriches a fieldValues map with all
 * official ${symbol} variables, including calculated Tafqeet dates,
 * page counts, and financial values.
 */
export function expandFieldValuesWithCanonicalSymbols(
  fieldValues: Record<string, string> = {},
  options?: {
    pageCount?: number;
    contractBodyHtml?: string;
  }
): Record<string, string> {
  const result: Record<string, string> = { ...fieldValues };

  // 1. Resolve date & year values
  const dateValue =
    fieldValues['تاريخ_العقد'] ||
    fieldValues['date_acte'] ||
    new Date().toISOString().slice(0, 10).replace(/-/g, '/');
  const dateWords = dateToAlgerianLegalWords(dateValue);

  result['date_acte'] = dateValue;
  result['today_date'] = dateValue;
  result['today_year'] = String(new Date().getFullYear());
  result['date2lettre_acte'] =
    fieldValues['date2lettre_acte'] || dateWords.fullWithNumbers;
  result['date2lettre_acte_only'] =
    fieldValues['date2lettre_acte_only'] || dateWords.lettersOnly;
  result['anne2lettre_acte'] =
    fieldValues['anne2lettre_acte'] || dateWords.yearLettersOnly;
  result['jourmois2lettre_acte'] =
    fieldValues['jourmois2lettre_acte'] || dateWords.dayMonthOnly;

  // 2. Resolve page counts
  const pages = options?.pageCount || 1;
  result['pages_acte'] = String(pages);
  result['pages2lettre_acte'] = pagesCountToAlgerianWords(pages);

  // 3. Resolve price & Tafqeet
  const rawPrice =
    fieldValues['الثمن_بالأرقام'] ||
    fieldValues['الثمن'] ||
    fieldValues['act_value'] ||
    fieldValues['total'] ||
    '';
  if (rawPrice) {
    result['act_value'] = rawPrice;
    result['total'] = rawPrice;
    const priceWords = currencyToAlgerianDinarWords(rawPrice);
    if (!result['totalLettre']) result['totalLettre'] = priceWords;
    if (!result['totalLettre_ar']) result['totalLettre_ar'] = priceWords;
    if (!result['الثمن_بالأحرف']) result['الثمن_بالأحرف'] = priceWords;
  }

  // 4. Resolve parties shortcuts
  const c1 =
    fieldValues['الطرف_الأول_الاسم'] ||
    fieldValues['الطرف_الأول'] ||
    fieldValues['partie1_nom_complet'] ||
    fieldValues['partie1_nom'] ||
    '';
  const c2 =
    fieldValues['الطرف_الثاني_الاسم'] ||
    fieldValues['الطرف_الثاني'] ||
    fieldValues['partie2_nom_complet'] ||
    fieldValues['partie2_nom'] ||
    '';
  if (c1) result['c_1'] = c1;
  if (c2) result['c_2'] = c2;
  if (!result['client_name'] && c2) result['client_name'] = c2;

  // 5. Resolve office metadata defaults
  if (!result['office_name']) {
    result['office_name'] =
      fieldValues['المكتب'] ||
      fieldValues['الموثق'] ||
      'مكتب التوثيق الرسمي';
  }

  // 6. Map all remaining canonical symbol keys from Arabic values
  for (const item of CANONICAL_NOTARY_SYMBOLS) {
    if (result[item.key] && result[item.key].trim() !== '') continue;
    for (const arKey of item.mappedFieldKeys) {
      const val = fieldValues[arKey] || fieldValues[arKey.replace(/[\s.]+/g, '_')];
      if (val !== undefined && val.trim() !== '') {
        result[item.key] = val.trim();
        break;
      }
    }
  }

  // 7. Inject full contract text if available
  if (options?.contractBodyHtml) {
    result['clause'] = options.contractBodyHtml;
    result['contenu_acte'] = options.contractBodyHtml;
    result['corps_acte'] = options.contractBodyHtml;
  }

  return result;
}
