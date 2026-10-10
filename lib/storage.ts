import {
  BackupBundle,
  ContractFolder,
  CustomTemplate,
  DerivedDocTemplate,
  DocumentRevision,
  DownloadArchiveItem,
  NotaryClerk,
  NotaryClause,
  PartyField,
  SavedDocument,
  SavedPartyRecord,
  SavedPropertyRecord,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  SubdivisionEstate,
  WordTemplateDefinition,
} from './types';
import { DEFAULT_WORD_TEMPLATES } from './docx-template-engine';

const DB_NAME = 'NotarySmartEditorDB';
const DB_VERSION = 5;
const STORE_TEMPLATES = 'templates';
const STORE_DOCUMENTS = 'documents';
const STORE_ESTATES = 'estates';
const STORE_CLAUSES = 'clauses';
const STORE_DERIVED_TEMPLATES = 'derived_templates';
const STORE_REVISIONS = 'revisions';
const STORE_DOWNLOADS = 'downloads';
const STORE_SAVED_PARTIES = 'saved_parties';
const STORE_SAVED_PROPERTIES = 'saved_properties';
const STORE_CLERKS = 'clerks';
const STORE_FOLDERS = 'folders';
const STORE_WORD_TEMPLATES = 'word_templates';

const LS_FIELDS_KEY = 'notary_default_party_fields_v2';
const LS_ACTIVE_DOC_KEY = 'notary_active_document_v2';

export const DEFAULT_PARTY_FIELDS: PartyField[] = [
  { key: 'الطرف_الأول_الاسم', label: 'الطرف الأول (الاسم واللقب)', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_تاريخ_الميلاد', label: 'تاريخ ميلاد الطرف الأول', value: '', category: 'party1', inputType: 'date' },
  { key: 'الطرف_الأول_مكان_الميلاد', label: 'مكان ميلاد الطرف الأول', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_النسب', label: 'اسم الأب والأم (بن/بنت)', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_الهوية', label: 'رقم ونوع وثيقة الهوية (ط.ت.و / ر.س)', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الأول_الإقامة', label: 'عنوان إقامة الطرف الأول', value: '', category: 'party1', inputType: 'text' },
  { key: 'الطرف_الثاني_الاسم', label: 'الطرف الثاني (الاسم واللقب)', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_تاريخ_الميلاد', label: 'تاريخ ميلاد الطرف الثاني', value: '', category: 'party2', inputType: 'date' },
  { key: 'الطرف_الثاني_مكان_الميلاد', label: 'مكان ميلاد الطرف الثاني', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_النسب', label: 'اسم الأب والأم (بن/بنت)', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_الهوية', label: 'رقم ونوع وثيقة الهوية (ط.ت.و / ر.س)', value: '', category: 'party2', inputType: 'text' },
  { key: 'الطرف_الثاني_الإقامة', label: 'عنوان إقامة الطرف الثاني', value: '', category: 'party2', inputType: 'text' },
  { key: 'رقم_الحصة', label: 'رقم الحصة العقارية', value: '', category: 'property', inputType: 'number' },
  { key: 'طبيعة_الحصة', label: 'طبيعة الحصة (شقة/محل/مرآب)', value: '', category: 'property', inputType: 'select', options: ['شقة سكنية', 'محل تجاري', 'مرآب', 'قطعة أرضية', 'مسكن فردي'] },
  { key: 'الطابق', label: 'الطابق والعمارة', value: '', category: 'property', inputType: 'text' },
  { key: 'المساحة', label: 'المساحة (م²)', value: '', category: 'property', inputType: 'number' },
  { key: 'الأجزاء_المشتركة', label: 'الحصة في الأجزاء المشتركة', value: '', category: 'property', inputType: 'text' },
  { key: 'تعيين_الحصة_الكامل', label: 'التعيين الكامل للحصة (من جدول الوصف)', value: '', category: 'property', inputType: 'text' },
  { key: 'مراجع_الوصف_التقسيمي', label: 'مراجع عقد الوصف التقسيمي', value: '', category: 'property', inputType: 'text' },
  { key: 'الثمن_بالأحرف', label: 'الثمن الإجمالي بالأحرف', value: '', category: 'financial', inputType: 'text' },
  { key: 'الثمن_بالأرقام', label: 'الثمن الإجمالي بالأرقام', value: '', category: 'financial', inputType: 'number' },
  { key: 'تاريخ_العقد', label: 'تاريخ تحرير العقد', value: '', category: 'custom', inputType: 'date' },
  { key: 'رقم_الفهرس', label: 'رقم الفهرس السنوي', value: '', category: 'custom', inputType: 'text' },
];

const pStyle = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;
const pCenterBold = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:center;font-weight:bold;`;

/**
 * Editable Derived Document Templates (Fully editable in-app inside the A4 Editor!).
 * Rely exclusively on the current Smart Variables Form + Parties & Property Designations.
 */
export const DEFAULT_DERIVED_DOC_TEMPLATES: DerivedDocTemplate[] = [
  {
    id: 'derived_extrait',
    code: 'extrait',
    name: 'المستخرج (Extrait)',
    description: 'مستخرج العقد الموجه لمفتشية التسجيل والطابع — قابل للتعديل الكامل داخل المحرر',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">مستخرج عقد توثيقي لأجل التسجيل</p>
      <p dir="rtl" style="${pStyle}">فهرس رقم: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> — بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الطرف الأول:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، المولود بتاريخ <span class="smart-tag" data-var="الطرف_الأول_تاريخ_الميلاد">[الطرف_الأول_تاريخ_الميلاد]</span> بـ <span class="smart-tag" data-var="الطرف_الأول_مكان_الميلاد">[الطرف_الأول_مكان_الميلاد]</span>، بن <span class="smart-tag" data-var="الطرف_الأول_النسب">[الطرف_الأول_النسب]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الأول_الإقامة">[الطرف_الأول_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الطرف الثاني:</span> <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، المولود بتاريخ <span class="smart-tag" data-var="الطرف_الثاني_تاريخ_الميلاد">[الطرف_الثاني_تاريخ_الميلاد]</span> بـ <span class="smart-tag" data-var="الطرف_الثاني_مكان_الميلاد">[الطرف_الثاني_مكان_الميلاد]</span>، بن <span class="smart-tag" data-var="الطرف_الثاني_النسب">[الطرف_الثاني_النسب]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الثاني_الإقامة">[الطرف_الثاني_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">تعيين العقار (الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span>):</span> <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الثمن الإجمالي:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
    `.trim(),
  },
  {
    id: 'derived_publication',
    code: 'publication',
    name: 'إجراء الشهر العقاري (Publication)',
    description: 'وثيقة إجراء الشهر العقاري الموجهة للمحافظة العقارية — قابلة للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">قائمة إجراء الشهر العقاري</p>
      <p dir="rtl" style="${pStyle}">موجب العقد المحرر بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> تحت رقم الفهرس: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">أولاً — تعيين الأطراف:</span></p>
      <p dir="rtl" style="${pStyle}">1. المتصرف (البائع): <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، تاريخ ومكان الميلاد: <span class="smart-tag" data-var="الطرف_الأول_تاريخ_الميلاد">[الطرف_الأول_تاريخ_الميلاد]</span> <span class="smart-tag" data-var="الطرف_الأول_مكان_الميلاد">[الطرف_الأول_مكان_الميلاد]</span>، النسب: <span class="smart-tag" data-var="الطرف_الأول_النسب">[الطرف_الأول_النسب]</span>.</p>
      <p dir="rtl" style="${pStyle}">2. المتصرف إليه (المشتري): <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، تاريخ ومكان الميلاد: <span class="smart-tag" data-var="الطرف_الثاني_تاريخ_الميلاد">[الطرف_الثاني_تاريخ_الميلاد]</span> <span class="smart-tag" data-var="الطرف_الثاني_مكان_الميلاد">[الطرف_الثاني_مكان_الميلاد]</span>، النسب: <span class="smart-tag" data-var="الطرف_الثاني_النسب">[الطرف_الثاني_النسب]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">ثانياً — التعيين الدقيق للعقار والوصف التقسيمي:</span></p>
      <p dir="rtl" style="${pStyle}">الحصة رقم: <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> — الطبيعة: <span class="smart-tag" data-var="طبيعة_الحصة">[طبيعة_الحصة]</span> — الطابق: <span class="smart-tag" data-var="الطابق">[الطابق]</span> — المساحة: <span class="smart-tag" data-var="المساحة">[المساحة]</span> م² — الأجزاء المشتركة: <span class="smart-tag" data-var="الأجزاء_المشتركة">[الأجزاء_المشتركة]</span>.</p>
      <p dir="rtl" style="${pStyle}">الوصف الكامل: <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}">مراجع الوصف التقسيمي: <span class="smart-tag" data-var="مراجع_الوصف_التقسيمي">[مراجع_الوصف_التقسيمي]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">ثالثاً — الثمن والشروط المالية:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
    `.trim(),
  },
  {
    id: 'derived_attestation',
    code: 'attestation',
    name: 'شهادة توثيقية / شهادة بيع (Attestation)',
    description: 'شهادة تسلم للأطراف لإثبات إبرام العقد بمكتب الموثق — قابلة للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">شهادة توثيقية</p>
      <p dir="rtl" style="${pStyle}">يشهد الموثق الموقع أسفله أنه بموجب عقد محرر بمكتبنا بتاريخ <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> تحت رقم الفهرس <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span>:</p>
      <p dir="rtl" style="${pStyle}">قام السيد(ة): <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، الحامل لوثيقة الهوية: <span class="smart-tag" data-var="الطرف_الأول_الهوية">[الطرف_الأول_الهوية]</span>.</p>
      <p dir="rtl" style="${pStyle}">بالتصرف لفائدة السيد(ة): <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، الحامل لوثيقة الهوية: <span class="smart-tag" data-var="الطرف_الثاني_الهوية">[الطرف_الثاني_الهوية]</span>.</p>
      <p dir="rtl" style="${pStyle}">في العقار المعين كما يلي: <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}">سلمت هذه الشهادة للمعني بالأمر للإدلاء بها في حدود ما يسمح به القانون.</p>
    `.trim(),
  },
  {
    id: 'derived_proce_verbal',
    code: 'proce_verbal',
    name: 'الصيغة التنفيذية / النسخة (Grosse)',
    description: 'النسخة الممهورة بالصيغة التنفيذية الرسمية — قابلة للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenterBold}">باسم الشعب الجزائري — نسخة تنفيذية</p>
      <p dir="rtl" style="${pStyle}">موجب العقد المحرر بتاريخ <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> تحت رقم <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span> بين <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span> و <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span> بخصوص الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> (<span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span>) بمبلغ <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span>.</p>
      <p dir="rtl" style="${pStyle}">وبناءً على ذلك، فإن الجمهورية الجزائرية الديمقراطية الشعبية تدعو وتأمر جميع المحضرين وكذا كل الأعوان الذين طُلب منهم ذلك أن ينفذوا هذا العقد، وعلى النواب العامين ووكلاء الجمهورية لدى المحاكم أن يمدوا يد المساعدة اللازمة لتنفيذه، وعلى جميع قادة وضباط القوة العمومية أن يمدوا يد المعونة بقوة القانون متى طُلب منهم ذلك بصفة قانونية.</p>
    `.trim(),
  },
  {
    id: 'derived_extrait_donation',
    code: 'extrait_donation',
    name: 'مستخرج الهبات (Extrait Donation)',
    description: 'مستخرج خاص بعقود الهبة والتبرعات — قابل للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">مستخرج عقد هبة</p>
      <p dir="rtl" style="${pStyle}">بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span> — فهرس رقم: <span class="smart-tag" data-var="رقم_الفهرس">[رقم_الفهرس]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الواهب:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الأول_الإقامة">[الطرف_الأول_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">الموهوب له:</span> <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الثاني_الإقامة">[الطرف_الثاني_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">العقار الموهوب:</span> <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">التقويم المالي للهبة:</span> <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
    `.trim(),
  },
  {
    id: 'derived_br4bis',
    code: 'br4bis',
    name: 'طلب PR4BIS (طلب معلومات عقارية)',
    description: 'استمارة طلب معلومات على الإجراءات المشهرة من المحافظة العقارية — قابلة للتعديل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">طلب معلومات على الإجراءات المشهرة (PR4BIS)</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">المالك الحالي / الطرف الأول:</span> <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>، تاريخ ومكان الميلاد: <span class="smart-tag" data-var="الطرف_الأول_تاريخ_الميلاد">[الطرف_الأول_تاريخ_الميلاد]</span> بـ <span class="smart-tag" data-var="الطرف_الأول_مكان_الميلاد">[الطرف_الأول_مكان_الميلاد]</span>، النسب: <span class="smart-tag" data-var="الطرف_الأول_النسب">[الطرف_الأول_النسب]</span>.</p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">تعيين العقار المطلوب بشأنه المعلومات:</span> الحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> — <span class="smart-tag" data-var="تعيين_الحصة_الكامل">[تعيين_الحصة_الكامل]</span></p>
      <p dir="rtl" style="${pStyle}"><span style="font-weight:bold;">مراجع الوصف التقسيمي / السند:</span> <span class="smart-tag" data-var="مراجع_الوصف_التقسيمي">[مراجع_الوصف_التقسيمي]</span></p>
    `.trim(),
  },
  {
    id: 'derived_recu_depot',
    code: 'recu_depot',
    name: 'وصل استلام الوديعة (Reçu de Dépôt)',
    description: 'وصل استلام المبالغ أو الوثائق المودعة بالمكتب — قابل للتعديل الكامل',
    headerHtml: '',
    footerHtml: '',
    updatedAt: '2026-01-01T00:00:00.000Z',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">وصل استلام وديعة</p>
      <p dir="rtl" style="${pStyle}">بتاريخ: <span class="smart-tag" data-var="تاريخ_العقد">[تاريخ_العقد]</span></p>
      <p dir="rtl" style="${pStyle}">استلمنا من السيد(ة): <span class="smart-tag" data-var="الطرف_الثاني_الاسم">[الطرف_الثاني_الاسم]</span>، المقيم بـ <span class="smart-tag" data-var="الطرف_الثاني_الإقامة">[الطرف_الثاني_الإقامة]</span>.</p>
      <p dir="rtl" style="${pStyle}">مبلغاً قدره: <span class="smart-tag" data-var="الثمن_بالأحرف">[الثمن_بالأحرف]</span> (<span class="smart-tag" data-var="الثمن_بالأرقام">[الثمن_بالأرقام]</span>).</p>
      <p dir="rtl" style="${pStyle}">وذلك بخصوص المعاملة المتعلقة بالحصة رقم <span class="smart-tag" data-var="رقم_الحصة">[رقم_الحصة]</span> مع السيد(ة) <span class="smart-tag" data-var="الطرف_الأول_الاسم">[الطرف_الأول_الاسم]</span>.</p>
    `.trim(),
  },
];

function openDatabase(): Promise<IDBDatabase | null> {
  if (typeof window === 'undefined' || !('indexedDB' in window)) {
    return Promise.resolve(null);
  }
  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        const stores = [
          STORE_TEMPLATES,
          STORE_DOCUMENTS,
          STORE_ESTATES,
          STORE_CLAUSES,
          STORE_DERIVED_TEMPLATES,
          STORE_REVISIONS,
          STORE_DOWNLOADS,
          STORE_SAVED_PARTIES,
          STORE_SAVED_PROPERTIES,
          STORE_CLERKS,
          STORE_FOLDERS,
          STORE_WORD_TEMPLATES,
        ];
        for (const s of stores) {
          if (!db.objectStoreNames.contains(s)) {
            db.createObjectStore(s, { keyPath: 'id' });
          }
        }
      };
      request.onsuccess = () => {
        const db = request.result;
        db.onversionchange = () => {
          db.close();
        };
        resolve(db);
      };
      request.onerror = () => resolve(null);
      request.onblocked = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
}

async function getAllFromStore<T>(storeName: string, lsFallbackKey: string): Promise<T[]> {
  const db = await openDatabase();
  if (db && db.objectStoreNames.contains(storeName)) {
    const dbResult = await new Promise<T[] | null>((resolve) => {
      try {
        const tx = db.transaction(storeName, 'readonly');
        const store = tx.objectStore(storeName);
        const req = store.getAll();
        req.onsuccess = () => resolve((req.result as T[]) || []);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
    if (dbResult && dbResult.length > 0) {
      return dbResult;
    }
  }
  if (typeof window !== 'undefined') {
    try {
      const raw = localStorage.getItem(lsFallbackKey);
      return raw ? (JSON.parse(raw) as T[]) : [];
    } catch {
      return [];
    }
  }
  return [];
}

async function putInStore<T extends { id: string }>(
  storeName: string,
  lsFallbackKey: string,
  item: T
): Promise<void> {
  const db = await openDatabase();
  if (db) {
    await new Promise<void>((resolve) => {
      try {
        const tx = db.transaction(storeName, 'readwrite');
        tx.objectStore(storeName).put(item);
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch {
        resolve();
      }
    });
  }
  if (typeof window !== 'undefined') {
    try {
      const existing = await getAllFromStore<T>(storeName, lsFallbackKey);
      const idx = existing.findIndex((x) => x.id === item.id);
      if (idx >= 0) existing[idx] = item;
      else existing.unshift(item);
      localStorage.setItem(lsFallbackKey, JSON.stringify(existing));
    } catch {
      // Ignore quota errors
    }
  }
}

async function deleteFromStore(
  storeName: string,
  lsFallbackKey: string,
  id: string
): Promise<void> {
  const db = await openDatabase();
  if (db) {
    await new Promise<void>((resolve) => {
      try {
        const tx = db.transaction(storeName, 'readwrite');
        tx.objectStore(storeName).delete(id);
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch {
        resolve();
      }
    });
  }
  if (typeof window !== 'undefined') {
    try {
      const existing = await getAllFromStore<{ id: string }>(storeName, lsFallbackKey);
      const filtered = existing.filter((x) => x.id !== id);
      localStorage.setItem(lsFallbackKey, JSON.stringify(filtered));
    } catch {
      // Ignore
    }
  }
}

// 1. Clauses (البنود الجاهزة) CRUD
export async function loadNotaryClauses(): Promise<NotaryClause[]> {
  const items = await getAllFromStore<NotaryClause>(STORE_CLAUSES, 'notary_clauses_v2');
  return items.sort((a, b) => a.order - b.order);
}

export async function saveNotaryClause(clause: NotaryClause): Promise<void> {
  await putInStore(STORE_CLAUSES, 'notary_clauses_v2', clause);
}

export async function deleteNotaryClause(id: string): Promise<void> {
  await deleteFromStore(STORE_CLAUSES, 'notary_clauses_v2', id);
}

// 2. Derived Document Templates (قوالب الوثائق المشتقة القابلة للتعديل) CRUD
export async function loadDerivedDocTemplates(): Promise<DerivedDocTemplate[]> {
  const saved = await getAllFromStore<DerivedDocTemplate>(
    STORE_DERIVED_TEMPLATES,
    'notary_derived_templates_v2'
  );
  if (saved.length === 0) {
    for (const defTpl of DEFAULT_DERIVED_DOC_TEMPLATES) {
      await putInStore(STORE_DERIVED_TEMPLATES, 'notary_derived_templates_v2', defTpl);
    }
    return DEFAULT_DERIVED_DOC_TEMPLATES;
  }
  return saved;
}

export async function saveDerivedDocTemplate(tpl: DerivedDocTemplate): Promise<void> {
  await putInStore(STORE_DERIVED_TEMPLATES, 'notary_derived_templates_v2', tpl);
}

export async function deleteDerivedDocTemplate(id: string): Promise<void> {
  await deleteFromStore(STORE_DERIVED_TEMPLATES, 'notary_derived_templates_v2', id);
}

// 2b. Word Templates (قوالب Word الرسمية والمخصصة) CRUD
export async function loadWordTemplates(): Promise<WordTemplateDefinition[]> {
  const saved = await getAllFromStore<WordTemplateDefinition>(
    STORE_WORD_TEMPLATES,
    'notary_word_templates_v1'
  );
  if (saved.length === 0) {
    for (const defTpl of DEFAULT_WORD_TEMPLATES) {
      await putInStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', defTpl);
    }
    return DEFAULT_WORD_TEMPLATES;
  }
  return saved;
}

export async function saveWordTemplate(tpl: WordTemplateDefinition): Promise<void> {
  await putInStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', tpl);
}

export async function deleteWordTemplate(id: string): Promise<void> {
  await deleteFromStore(STORE_WORD_TEMPLATES, 'notary_word_templates_v1', id);
}

// 3. Document Revisions (سجل النسخ ومقارنة التعديلات Diff) CRUD
export async function loadDocumentRevisions(): Promise<DocumentRevision[]> {
  const items = await getAllFromStore<DocumentRevision>(STORE_REVISIONS, 'notary_revisions_v2');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDocumentRevision(rev: DocumentRevision): Promise<void> {
  await putInStore(STORE_REVISIONS, 'notary_revisions_v2', rev);
}

export async function deleteDocumentRevision(id: string): Promise<void> {
  await deleteFromStore(STORE_REVISIONS, 'notary_revisions_v2', id);
}

// 4. Download Archive (أرشيف التحميلات) CRUD
export async function loadDownloadArchive(): Promise<DownloadArchiveItem[]> {
  const items = await getAllFromStore<DownloadArchiveItem>(STORE_DOWNLOADS, 'notary_downloads_v2');
  return items.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}

export async function saveDownloadArchiveItem(item: DownloadArchiveItem): Promise<void> {
  await putInStore(STORE_DOWNLOADS, 'notary_downloads_v2', item);
}

export async function deleteDownloadArchiveItem(id: string): Promise<void> {
  await deleteFromStore(STORE_DOWNLOADS, 'notary_downloads_v2', id);
}

export async function clearDownloadArchive(): Promise<void> {
  const items = await loadDownloadArchive();
  for (const item of items) {
    await deleteDownloadArchiveItem(item.id);
  }
}

// 5. Office Custom Templates CRUD (Starts empty — user adds their own)
export async function loadCustomTemplates(): Promise<CustomTemplate[]> {
  const items = await getAllFromStore<CustomTemplate>(STORE_TEMPLATES, 'notary_templates_v1');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveCustomTemplate(template: CustomTemplate): Promise<void> {
  await putInStore(STORE_TEMPLATES, 'notary_templates_v1', template);
}

export async function deleteCustomTemplate(id: string): Promise<void> {
  await deleteFromStore(STORE_TEMPLATES, 'notary_templates_v1', id);
}

export async function exportCustomTemplatesJson(): Promise<string> {
  const templates = await loadCustomTemplates();
  return JSON.stringify(templates, null, 2);
}

export async function importCustomTemplatesJson(templates: CustomTemplate[]): Promise<number> {
  let count = 0;
  for (const tpl of templates) {
    if (tpl && tpl.name && (tpl.bodyHtml !== undefined || tpl.id)) {
      const sanitized: CustomTemplate = {
        id: tpl.id || `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        name: tpl.name,
        category: tpl.category || 'عام',
        description: tpl.description || '',
        bodyHtml: tpl.bodyHtml || '',
        headerHtml: tpl.headerHtml || '',
        footerHtml: tpl.footerHtml || '',
        pageNumberingEnabled: !!tpl.pageNumberingEnabled,
        extractedPlaceholders: Array.isArray(tpl.extractedPlaceholders) ? tpl.extractedPlaceholders : [],
        defaultFieldValues: tpl.defaultFieldValues || {},
        updatedAt: tpl.updatedAt || new Date().toISOString(),
      };
      await saveCustomTemplate(sanitized);
      count++;
    }
  }
  return count;
}

// 6. Saved Documents CRUD
export async function loadSavedDocuments(): Promise<SavedDocument[]> {
  const items = await getAllFromStore<SavedDocument>(STORE_DOCUMENTS, 'notary_documents_v1');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveDocumentRecord(doc: SavedDocument): Promise<void> {
  await putInStore(STORE_DOCUMENTS, 'notary_documents_v1', doc);
}

export async function deleteDocumentRecord(id: string): Promise<void> {
  await deleteFromStore(STORE_DOCUMENTS, 'notary_documents_v1', id);
}

// 7. Subdivision Estates (جداول الوصف التقسيمي) CRUD
export async function loadSubdivisionEstates(): Promise<SubdivisionEstate[]> {
  const items = await getAllFromStore<SubdivisionEstate>(STORE_ESTATES, 'notary_estates_v1');
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function saveSubdivisionEstate(estate: SubdivisionEstate): Promise<void> {
  await putInStore(STORE_ESTATES, 'notary_estates_v1', estate);
}

export async function deleteSubdivisionEstate(id: string): Promise<void> {
  await deleteFromStore(STORE_ESTATES, 'notary_estates_v1', id);
}

// 7b. Saved Parties & Saved Properties Directory (دفتر الأطراف والعقارات القابل للاستدعاء)
export async function loadSavedParties(): Promise<SavedPartyRecord[]> {
  const items = await getAllFromStore<SavedPartyRecord>(
    STORE_SAVED_PARTIES,
    'notary_saved_parties_v25'
  );
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePartyRecord(party: SavedPartyRecord): Promise<void> {
  await putInStore(STORE_SAVED_PARTIES, 'notary_saved_parties_v25', party);
}

export async function deletePartyRecord(id: string): Promise<void> {
  await deleteFromStore(STORE_SAVED_PARTIES, 'notary_saved_parties_v25', id);
}

export async function loadSavedProperties(): Promise<SavedPropertyRecord[]> {
  const items = await getAllFromStore<SavedPropertyRecord>(
    STORE_SAVED_PROPERTIES,
    'notary_saved_properties_v25'
  );
  return items.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
}

export async function savePropertyRecord(prop: SavedPropertyRecord): Promise<void> {
  await putInStore(STORE_SAVED_PROPERTIES, 'notary_saved_properties_v25', prop);
}

export async function deletePropertyRecord(id: string): Promise<void> {
  await deleteFromStore(STORE_SAVED_PROPERTIES, 'notary_saved_properties_v25', id);
}

export const DEFAULT_NOTARY_CLERKS: NotaryClerk[] = [
  { id: 'notary-head', name: 'الموثق الرئيسي (الأستاذ)', role: 'notary', color: '#1e3a8a', createdAt: '2026-01-01T00:00:00.000Z' },
  { id: 'clerk-1', name: 'الكاتب الأول (أمين)', role: 'clerk', color: '#047857', createdAt: '2026-01-01T00:00:00.000Z' },
  { id: 'clerk-2', name: 'الكاتب الثاني (فاطمة)', role: 'clerk', color: '#b91c1c', createdAt: '2026-01-01T00:00:00.000Z' },
  { id: 'clerk-3', name: 'الكاتب الثالث (سفيان)', role: 'clerk', color: '#d97706', createdAt: '2026-01-01T00:00:00.000Z' },
];

export async function loadNotaryClerks(): Promise<NotaryClerk[]> {
  const items = await getAllFromStore<NotaryClerk>(STORE_CLERKS, 'notary_clerks_v1');
  if (items.length === 0) {
    for (const c of DEFAULT_NOTARY_CLERKS) {
      await putInStore(STORE_CLERKS, 'notary_clerks_v1', c);
    }
    return DEFAULT_NOTARY_CLERKS;
  }
  return items;
}

export async function saveNotaryClerk(clerk: NotaryClerk): Promise<void> {
  await putInStore(STORE_CLERKS, 'notary_clerks_v1', clerk);
}

export async function deleteNotaryClerk(id: string): Promise<void> {
  await deleteFromStore(STORE_CLERKS, 'notary_clerks_v1', id);
}

export async function loadContractFolders(): Promise<ContractFolder[]> {
  const items = await getAllFromStore<ContractFolder>(STORE_FOLDERS, 'notary_folders_v1');
  return items.sort((a, b) => a.name.localeCompare(b.name));
}

export async function saveContractFolder(folder: ContractFolder): Promise<void> {
  await putInStore(STORE_FOLDERS, 'notary_folders_v1', folder);
}

export async function deleteContractFolder(id: string): Promise<void> {
  await deleteFromStore(STORE_FOLDERS, 'notary_folders_v1', id);
}

// 8. Default / Custom Party Fields
export function loadPartyFields(): PartyField[] {
  if (typeof window === 'undefined') return DEFAULT_PARTY_FIELDS;
  try {
    const raw = localStorage.getItem(LS_FIELDS_KEY);
    if (!raw) return DEFAULT_PARTY_FIELDS;
    const parsed = JSON.parse(raw) as PartyField[];
    return Array.isArray(parsed) && parsed.length > 0 ? parsed : DEFAULT_PARTY_FIELDS;
  } catch {
    return DEFAULT_PARTY_FIELDS;
  }
}

export function savePartyFields(fields: PartyField[]): void {
  if (typeof window === 'undefined') return;
  try {
    localStorage.setItem(LS_FIELDS_KEY, JSON.stringify(fields));
  } catch {
    // Ignore
  }
}

// 9. Active Draft Session
export function loadActiveDraftSession(): SavedDocument | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = localStorage.getItem(LS_ACTIVE_DOC_KEY);
    return raw ? (JSON.parse(raw) as SavedDocument) : null;
  } catch {
    return null;
  }
}

export function saveActiveDraftSession(doc: SavedDocument): void {
  if (typeof window === 'undefined') return;
  try {
    localStorage.setItem(LS_ACTIVE_DOC_KEY, JSON.stringify(doc));
  } catch {
    // Ignore
  }
}

// 10. Full Backup Export & Import
export async function exportFullBackupBundle(): Promise<BackupBundle> {
  const [
    templates,
    documents,
    estates,
    clauses,
    derivedTemplates,
    revisions,
    downloads,
    savedParties,
    savedProperties,
    clerks,
    folders,
  ] = await Promise.all([
    loadCustomTemplates(),
    loadSavedDocuments(),
    loadSubdivisionEstates(),
    loadNotaryClauses(),
    loadDerivedDocTemplates(),
    loadDocumentRevisions(),
    loadDownloadArchive(),
    loadSavedParties(),
    loadSavedProperties(),
    loadNotaryClerks(),
    loadContractFolders(),
  ]);
  const defaultFields = loadPartyFields();
  return {
    version: '2.5.0',
    exportedAt: new Date().toISOString(),
    clerks,
    folders,
    templates,
    documents,
    estates,
    clauses,
    derivedTemplates,
    revisions,
    downloads,
    savedParties,
    savedProperties,
    defaultFields,
  };
}

export async function importFullBackupBundle(bundle: BackupBundle): Promise<{
  templatesCount: number;
  documentsCount: number;
  estatesCount: number;
  clausesCount: number;
}> {
  let templatesCount = 0;
  let documentsCount = 0;
  let estatesCount = 0;
  let clausesCount = 0;

  if (Array.isArray(bundle.clerks)) {
    for (const cl of bundle.clerks) {
      if (cl && cl.id && cl.name) {
        await saveNotaryClerk(cl);
      }
    }
  }

  if (Array.isArray(bundle.folders)) {
    for (const f of bundle.folders) {
      if (f && f.id && f.name) {
        await saveContractFolder(f);
      }
    }
  }

  if (Array.isArray(bundle.templates)) {
    for (const tpl of bundle.templates) {
      if (tpl && tpl.id && tpl.name) {
        await saveCustomTemplate(tpl);
        templatesCount++;
      }
    }
  }

  if (Array.isArray(bundle.documents)) {
    for (const doc of bundle.documents) {
      if (doc && doc.id && doc.title) {
        await saveDocumentRecord(doc);
        documentsCount++;
      }
    }
  }

  if (Array.isArray(bundle.estates)) {
    for (const est of bundle.estates) {
      if (est && est.id && est.estateName) {
        await saveSubdivisionEstate(est);
        estatesCount++;
      }
    }
  }

  if (Array.isArray(bundle.clauses)) {
    for (const cl of bundle.clauses) {
      if (cl && cl.id && cl.title) {
        await saveNotaryClause(cl);
        clausesCount++;
      }
    }
  }

  if (Array.isArray(bundle.derivedTemplates)) {
    for (const dt of bundle.derivedTemplates) {
      if (dt && dt.id && dt.name) {
        await saveDerivedDocTemplate(dt);
      }
    }
  }

  if (Array.isArray(bundle.revisions)) {
    for (const rev of bundle.revisions) {
      if (rev && rev.id) {
        await saveDocumentRevision(rev);
      }
    }
  }

  if (Array.isArray(bundle.downloads)) {
    for (const dl of bundle.downloads) {
      if (dl && dl.id) {
        await saveDownloadArchiveItem(dl);
      }
    }
  }

  if (Array.isArray(bundle.defaultFields) && bundle.defaultFields.length > 0) {
    savePartyFields(bundle.defaultFields);
  }

  return { templatesCount, documentsCount, estatesCount, clausesCount };
}
