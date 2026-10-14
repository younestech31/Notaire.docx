/**
 * Standard Latin Variables & Placeholders Dictionary for Algerian Notary Offices.
 * 
 * Maps between Algerian notarial contract Arabic fields (الطرف الأول، الفهرس، العقار...)
 * and standard Latin tags ([partie1_nom], [num_repertoire], [date_acte]...).
 */

export interface StandardNotaryTagDefinition {
  tag: string;
  label: string;
  category: 'parties' | 'acte' | 'immeuble' | 'finance' | 'texte';
  sampleValue: string;
  mappedFieldKeys: string[];
}

export const STANDARD_NOTARY_TAG_DEFINITIONS: StandardNotaryTagDefinition[] = [
  // --- PARTIE 1 (الطرف الأول / البائع / المتصرف) ---
  {
    tag: 'partie1_nom',
    label: 'الطرف الأول: الاسم واللقب',
    category: 'parties',
    sampleValue: 'بن عيسى أحمد',
    mappedFieldKeys: ['الطرف_الأول_الاسم', 'الطرف_الأول_اللقب', 'الطرف_الأول_الاسم_واللقب', 'الطرف_الأول_الاسم_الكامل'],
  },
  {
    tag: 'partie1_prenom',
    label: 'الطرف الأول: الاسم الشخصي',
    category: 'parties',
    sampleValue: 'أحمد',
    mappedFieldKeys: ['الطرف_الأول_الاسم_الشخصي', 'الطرف_الأول_الاسم'],
  },
  {
    tag: 'partie1_nom_complet',
    label: 'الطرف الأول: الاسم الكامل',
    category: 'parties',
    sampleValue: 'بن عيسى أحمد بن مصطفى',
    mappedFieldKeys: ['الطرف_الأول_الاسم', 'الطرف_الأول_الاسم_الكامل'],
  },
  {
    tag: 'partie1_date_naissance',
    label: 'الطرف الأول: تاريخ الميلاد',
    category: 'parties',
    sampleValue: '1980/05/14',
    mappedFieldKeys: ['الطرف_الأول_تاريخ_الميلاد'],
  },
  {
    tag: 'partie1_lieu_naissance',
    label: 'الطرف الأول: مكان الميلاد',
    category: 'parties',
    sampleValue: 'الجزائر الوسطى',
    mappedFieldKeys: ['الطرف_الأول_مكان_الميلاد'],
  },
  {
    tag: 'partie1_filiation',
    label: 'الطرف الأول: النسب (اسم الأب والأم)',
    category: 'parties',
    sampleValue: 'بن مصطفى وفاطمة الزهراء بن علي',
    mappedFieldKeys: ['الطرف_الأول_النسب'],
  },
  {
    tag: 'partie1_adresse',
    label: 'الطرف الأول: العنوان الكامل',
    category: 'parties',
    sampleValue: 'حي 500 مسكن، عمارة 12، بئر مراد رايس، الجزائر',
    mappedFieldKeys: ['الطرف_الأول_العنوان'],
  },
  {
    tag: 'partie1_cni',
    label: 'الطرف الأول: بطاقة التعريف / جواز السفر',
    category: 'parties',
    sampleValue: 'بطاقة تعريف بيومترية رقم 1122334455 صادرة بتاريخ 2020/02/10 عن دائرة بئر مراد رايس',
    mappedFieldKeys: ['الطرف_الأول_بطاقة_التعريف', 'الطرف_الأول_رقم_الهوية', 'الطرف_الأول_الهوية'],
  },
  {
    tag: 'partie1_nin',
    label: 'الطرف الأول: رقم التعريف الوطني (NIN)',
    category: 'parties',
    sampleValue: '180160100234567890',
    mappedFieldKeys: ['الطرف_الأول_الرقم_الوطني', 'الطرف_الأول_NIN'],
  },
  {
    tag: 'partie1_profession',
    label: 'الطرف الأول: المهنة',
    category: 'parties',
    sampleValue: 'تاجر',
    mappedFieldKeys: ['الطرف_الأول_المهنة'],
  },
  {
    tag: 'partie1_nationalite',
    label: 'الطرف الأول: الجنسية',
    category: 'parties',
    sampleValue: 'جزائرية',
    mappedFieldKeys: ['الطرف_الأول_الجنسية'],
  },
  {
    tag: 'partie1_situation_familiale',
    label: 'الطرف الأول: الحالة العائلية',
    category: 'parties',
    sampleValue: 'متزوج بنظام الاشتراك في الأموال',
    mappedFieldKeys: ['الطرف_الأول_الحالة_العائلية'],
  },

  // --- PARTIE 2 (الطرف الثاني / المشتري / المتصرف إليه) ---
  {
    tag: 'partie2_nom',
    label: 'الطرف الثاني: الاسم واللقب',
    category: 'parties',
    sampleValue: 'منصوري كريم',
    mappedFieldKeys: ['الطرف_الثاني_الاسم', 'الطرف_الثاني_اللقب', 'الطرف_الثاني_الاسم_واللقب', 'الطرف_الثاني_الاسم_الكامل'],
  },
  {
    tag: 'partie2_prenom',
    label: 'الطرف الثاني: الاسم الشخصي',
    category: 'parties',
    sampleValue: 'كريم',
    mappedFieldKeys: ['الطرف_الثاني_الاسم_الشخصي', 'الطرف_الثاني_الاسم'],
  },
  {
    tag: 'partie2_nom_complet',
    label: 'الطرف الثاني: الاسم الكامل',
    category: 'parties',
    sampleValue: 'منصوري كريم بن عمر',
    mappedFieldKeys: ['الطرف_الثاني_الاسم', 'الطرف_الثاني_الاسم_الكامل'],
  },
  {
    tag: 'partie2_date_naissance',
    label: 'الطرف الثاني: تاريخ الميلاد',
    category: 'parties',
    sampleValue: '1985/11/20',
    mappedFieldKeys: ['الطرف_الثاني_تاريخ_الميلاد'],
  },
  {
    tag: 'partie2_lieu_naissance',
    label: 'الطرف الثاني: مكان الميلاد',
    category: 'parties',
    sampleValue: 'البليدة',
    mappedFieldKeys: ['الطرف_الثاني_مكان_الميلاد'],
  },
  {
    tag: 'partie2_filiation',
    label: 'الطرف الثاني: النسب',
    category: 'parties',
    sampleValue: 'بن عمر وخديجة بلقاسم',
    mappedFieldKeys: ['الطرف_الثاني_النسب'],
  },
  {
    tag: 'partie2_adresse',
    label: 'الطرف الثاني: العنوان الكامل',
    category: 'parties',
    sampleValue: 'شارع الشهداء رقم 45، الأبيار، الجزائر',
    mappedFieldKeys: ['الطرف_الثاني_العنوان'],
  },
  {
    tag: 'partie2_cni',
    label: 'الطرف الثاني: بطاقة التعريف / جواز السفر',
    category: 'parties',
    sampleValue: 'بطاقة بيومترية رقم 9988776655 صادرة بتاريخ 2021/04/15 عن دائرة بوزريعة',
    mappedFieldKeys: ['الطرف_الثاني_بطاقة_التعريف', 'الطرف_الثاني_رقم_الهوية', 'الطرف_الثاني_الهوية'],
  },
  {
    tag: 'partie2_nin',
    label: 'الطرف الثاني: رقم التعريف الوطني (NIN)',
    category: 'parties',
    sampleValue: '185160200987654321',
    mappedFieldKeys: ['الطرف_الثاني_الرقم_الوطني', 'الطرف_الثاني_NIN'],
  },
  {
    tag: 'partie2_profession',
    label: 'الطرف الثاني: المهنة',
    category: 'parties',
    sampleValue: 'مهندس معماري',
    mappedFieldKeys: ['الطرف_الثاني_المهنة'],
  },
  {
    tag: 'partie2_nationalite',
    label: 'الطرف الثاني: الجنسية',
    category: 'parties',
    sampleValue: 'جزائرية',
    mappedFieldKeys: ['الطرف_الثاني_الجنسية'],
  },
  {
    tag: 'partie2_situation_familiale',
    label: 'الطرف الثاني: الحالة العائلية',
    category: 'parties',
    sampleValue: 'متزوج',
    mappedFieldKeys: ['الطرف_الثاني_الحالة_العائلية'],
  },

  // --- NOTAIRE & ETUDE (الموثق والمكتب التوثيقي) ---
  {
    tag: 'notaire_nom',
    label: 'اسم ولقب الموثق',
    category: 'acte',
    sampleValue: 'الأستاذ عبد الرحمن بوزيد',
    mappedFieldKeys: ['اسم_الموثق', 'الموثق', 'الأستاذ', 'الموثق_المحرر'],
  },
  {
    tag: 'notaire_qualite',
    label: 'صفة الموثق',
    category: 'acte',
    sampleValue: 'موثق لدى محكمة سيدي امحمد',
    mappedFieldKeys: ['صفة_الموثق', 'المحكمة_المختصة'],
  },
  {
    tag: 'etude_commune',
    label: 'بلدية ومقر المكتب',
    category: 'acte',
    sampleValue: 'الجزائر الوسطى',
    mappedFieldKeys: ['مقر_المكتب', 'بلدية_المكتب', 'البلدية'],
  },
  {
    tag: 'etude_wilaya',
    label: 'ولاية المكتب التوثيقي',
    category: 'acte',
    sampleValue: 'الجزائر',
    mappedFieldKeys: ['ولاية_المكتب', 'الولاية'],
  },
  {
    tag: 'etude_adresse',
    label: 'عنوان المكتب التوثيقي',
    category: 'acte',
    sampleValue: '04 شارع ديدوش مراد، الجزائر الوسطى',
    mappedFieldKeys: ['عنوان_المكتب'],
  },
  {
    tag: 'etude_tel',
    label: 'هاتف المكتب',
    category: 'acte',
    sampleValue: '021.73.45.67',
    mappedFieldKeys: ['هاتف_المكتب', 'الهاتف'],
  },
  {
    tag: 'chambre_regionale',
    label: 'الغرفة الجهوية للموثقين',
    category: 'acte',
    sampleValue: 'الغرفة الجهوية لموثقي ناحية الوسط',
    mappedFieldKeys: ['الغرفة_الجهوية', 'الغرفة_الوطنية'],
  },

  // --- ACTE & REPERTOIRE (بيانات العقد والفهرس) ---
  {
    tag: 'num_repertoire',
    label: 'رقم الفهرس الرسمي للعقد',
    category: 'acte',
    sampleValue: '1428',
    mappedFieldKeys: ['رقم_الفهرس', 'الفهرس', 'رقم_العقد'],
  },
  {
    tag: 'annee_acte',
    label: 'سنة إبرام العقد',
    category: 'acte',
    sampleValue: '2025',
    mappedFieldKeys: ['سنة_العقد', 'السنة'],
  },
  {
    tag: 'date_acte',
    label: 'تاريخ إبرام العقد (أرقام)',
    category: 'acte',
    sampleValue: '2025/03/12',
    mappedFieldKeys: ['تاريخ_العقد', 'التاريخ', 'تاريخ_التحرير'],
  },
  {
    tag: 'date_acte_lettres',
    label: 'تاريخ العقد بالحروف القانونية',
    category: 'acte',
    sampleValue: 'الثاني عشر من شهر مارس سنة ألفين وخمسة وعشرين',
    mappedFieldKeys: ['تاريخ_العقد_بالحروف', 'تاريخ_التحرير_بالحروف'],
  },
  {
    tag: 'type_acte',
    label: 'نوع العقد / طبيعة التصرف',
    category: 'acte',
    sampleValue: 'عقد بيع شقة سكنية',
    mappedFieldKeys: ['نوع_العقد', 'طبيعة_العقد', 'موضوع_العقد'],
  },
  {
    tag: 'objet_acte',
    label: 'موضوع العقد المختصر',
    category: 'acte',
    sampleValue: 'بيع عقار سكني مفرز',
    mappedFieldKeys: ['موضوع_العقد', 'نوع_العقد'],
  },
  {
    tag: 'inspection_enregistrement',
    label: 'مفتشية التسجيل والطابع المختصة',
    category: 'acte',
    sampleValue: 'مفتشية التسجيل بسيدي امحمد',
    mappedFieldKeys: ['مفتشية_التسجيل', 'مصلحة_التسجيل'],
  },
  {
    tag: 'conservation_fonciere',
    label: 'المحافظة العقارية المختصة',
    category: 'acte',
    sampleValue: 'المحافظة العقارية للجزائر - سيدي امحمد',
    mappedFieldKeys: ['المحافظة_العقارية'],
  },

  // --- IMMOBILIER & BIEN (التعيين العقاري والمسح) ---
  {
    tag: 'designation_bien',
    label: 'التعيين العقاري الكامل',
    category: 'immeuble',
    sampleValue: 'شقة سكنية كائنة بالطابق الثالث من العمارة (أ) متكونة من أربع غرف ومطبخ ومرفقات',
    mappedFieldKeys: ['التعيين_العقاري', 'تعيين_العقار', 'العقار', 'تعيين_الحصة_الكامل', 'تعيين_الحصة'],
  },
  {
    tag: 'commune_bien',
    label: 'بلدية موقع العقار',
    category: 'immeuble',
    sampleValue: 'الجزائر الوسطى',
    mappedFieldKeys: ['بلدية_العقار'],
  },
  {
    tag: 'section_cadastrale',
    label: 'القسم المسحي',
    category: 'immeuble',
    sampleValue: '04',
    mappedFieldKeys: ['القسم_المسحي', 'القسم'],
  },
  {
    tag: 'ilot_cadastral',
    label: 'مجموعة الملكيات / الجزيرة',
    category: 'immeuble',
    sampleValue: '12',
    mappedFieldKeys: ['مجموعة_الملكيات', 'الجزيرة'],
  },
  {
    tag: 'numero_plan',
    label: 'رقم المخطط المسحي',
    category: 'immeuble',
    sampleValue: '148',
    mappedFieldKeys: ['رقم_المخطط', 'المخطط'],
  },
  {
    tag: 'lot_numero',
    label: 'رقم الحصة العقارية المفرزة',
    category: 'immeuble',
    sampleValue: '07',
    mappedFieldKeys: ['رقم_الحصة', 'الحصة'],
  },
  {
    tag: 'lot_nature',
    label: 'طبيعة الحصة العقارية',
    category: 'immeuble',
    sampleValue: 'شقة سكنية',
    mappedFieldKeys: ['طبيعة_الحصة', 'طبيعة_العقار'],
  },
  {
    tag: 'lot_superficie',
    label: 'المساحة الصافية بالمتر المربع',
    category: 'immeuble',
    sampleValue: '115.50',
    mappedFieldKeys: ['مساحة_العقار', 'المساحة'],
  },
  {
    tag: 'lot_quote_part',
    label: 'الأجزاء المشتركة في الملكية المشتركة',
    category: 'immeuble',
    sampleValue: '125 / 1000 جزء من الأجزاء المشتركة العامة',
    mappedFieldKeys: ['الأجزاء_المشتركة'],
  },
  {
    tag: 'origine_propriete',
    label: 'أصل الملكية والسندات السابقة',
    category: 'immeuble',
    sampleValue: 'عقد بيع مشهر بالمحافظة العقارية تحت حجم 120 رقم 45 بتاريخ 2015/06/18',
    mappedFieldKeys: ['أصل_الملكية', 'سند_الملكية'],
  },
  {
    tag: 'ref_division',
    label: 'مراجع الوصف التقسيمي للملكية المشتركة (EDD)',
    category: 'immeuble',
    sampleValue: 'عقد وصف تقسيمي محرر بتاريخ 2010/01/20 مشهر بحجم 88 رقم 12',
    mappedFieldKeys: ['الوصف_التقسيمي', 'مراجع_الوصف_التقسيمي'],
  },

  // --- FINANCES & PRIX (المبالغ المالية والرسوم) ---
  {
    tag: 'prix_chiffres',
    label: 'الثمن بالأرقام (دج)',
    category: 'finance',
    sampleValue: '18500000.00',
    mappedFieldKeys: ['الثمن_بالأرقام', 'ثمن_البيع_بالأرقام', 'الثمن', 'المبلغ_بالأرقام'],
  },
  {
    tag: 'prix_lettres',
    label: 'الثمن بالحروف القانونية (دج)',
    category: 'finance',
    sampleValue: 'ثمانية عشر مليون وخمسمائة ألف دينار جزائري',
    mappedFieldKeys: ['الثمن_بالحروف', 'الثمن_بالأحرف', 'ثمن_البيع_بالحروف', 'ثمن_البيع_بالأحرف', 'المبلغ_بالحروف', 'المبلغ_بالأحرف'],
  },
  {
    tag: 'modalites_paiement',
    label: 'طريقة وكيفية الوفاء بالثمن',
    category: 'finance',
    sampleValue: 'تم دفع الثمن بالكامل نقداً ومحاسبة بمجلس العقد بين أيدي الموثق بواسطة صك بنكي',
    mappedFieldKeys: ['طريقة_الدفع', 'كيفية_الوفاء', 'شروط_الدفع'],
  },
  {
    tag: 'banque_cheque',
    label: 'بيانات الصك البنكي أو التحويل',
    category: 'finance',
    sampleValue: 'صك معتمد رقم 0045892 مسحوب على بنك التنمية المحلية (BDL) وكالة الجزائر',
    mappedFieldKeys: ['بيانات_الصك', 'رقم_الصك'],
  },
  {
    tag: 'droits_enregistrement',
    label: 'حقوق التسجيل المقدرة (دج)',
    category: 'finance',
    sampleValue: '925000.00 دج (5%)',
    mappedFieldKeys: ['حقوق_التسجيل', 'رسوم_التسجيل'],
  },
  {
    tag: 'taxe_publicite',
    label: 'رسم الشهر العقاري (دج)',
    category: 'finance',
    sampleValue: '185000.00 دج (1%)',
    mappedFieldKeys: ['رسوم_الشهر', 'رسم_الشهر_العقاري'],
  },
  {
    tag: 'honoraires_notaire',
    label: 'أتعاب الموثق المقررة قانوناً (دج)',
    category: 'finance',
    sampleValue: '145000.00 دج طبقاً للجدول الرسمي للتعريفة التوثيقية',
    mappedFieldKeys: ['أتعاب_الموثق', 'الأتعاب'],
  },

  // --- TEXTE & CORPS (متن العقد والبنود الكاملة) ---
  {
    tag: 'contenu_acte',
    label: 'المتن الكامل لعقد التوثيق (جميع البنود المنسقة)',
    category: 'texte',
    sampleValue: '[المتن الكامل للعقد كما حرر بالمحرر التوثيقي]',
    mappedFieldKeys: ['متن_العقد', 'نص_العقد', 'محتوى_العقد'],
  },
  {
    tag: 'texte_acte',
    label: 'نص العقد (مرادف لـ contenu_acte)',
    category: 'texte',
    sampleValue: '[نص العقد المحرر]',
    mappedFieldKeys: ['نص_العقد', 'متن_العقد'],
  },
  {
    tag: 'corps_acte',
    label: 'متن العقد (مرادف لـ contenu_acte)',
    category: 'texte',
    sampleValue: '[متن العقد المحرر داخل صفحة A4]',
    mappedFieldKeys: ['متن_العقد', 'نص_العقد'],
  },
];

/**
 * Builds a fast lookup map: Latin tag -> list of mapped Arabic keys
 */
export const LATIN_TO_ARABIC_KEYS_MAP: Record<string, string[]> = {};
for (const def of STANDARD_NOTARY_TAG_DEFINITIONS) {
  LATIN_TO_ARABIC_KEYS_MAP[def.tag.toLowerCase()] = def.mappedFieldKeys;
}

/**
 * Builds a fast lookup map: Arabic key -> list of mapped Latin tags
 */
export const ARABIC_TO_LATIN_TAGS_MAP: Record<string, string[]> = {};
for (const def of STANDARD_NOTARY_TAG_DEFINITIONS) {
  for (const arKey of def.mappedFieldKeys) {
    const norm = arKey.replace(/[\s.]+/g, '_');
    if (!ARABIC_TO_LATIN_TAGS_MAP[norm]) {
      ARABIC_TO_LATIN_TAGS_MAP[norm] = [];
    }
    ARABIC_TO_LATIN_TAGS_MAP[norm].push(def.tag);
  }
}

/**
 * Expands a fieldValues map to include all corresponding standard Latin tags
 * so that Word templates with Latin tags ([partie1_nom], [date_acte], etc.)
 * resolve immediately.
 */
export function expandFieldValuesWithStandardLatinTags(
  fieldValues: Record<string, string> = {},
  contractBodyHtml?: string
): Record<string, string> {
  const expanded: Record<string, string> = { ...fieldValues };

  // 1. Populate Latin tags from Arabic fieldValues
  for (const def of STANDARD_NOTARY_TAG_DEFINITIONS) {
    if (expanded[def.tag]) continue;
    for (const arKey of def.mappedFieldKeys) {
      const val = fieldValues[arKey] || fieldValues[arKey.replace(/[\s.]+/g, '_')];
      if (val !== undefined && val.trim() !== '') {
        expanded[def.tag] = val.trim();
        break;
      }
    }
  }

  // 2. Contract body paragraphs are spliced natively into templates at block-level,
  // never injected as escaped raw string values into text fields.

  return expanded;
}


