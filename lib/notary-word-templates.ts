/**
 * Built-in Word (.docx) Templates Library for Algerian Notary Offices.
 *
 * Implements the 6 standard Algerian notary document types:
 * 1. أصل العقد (Acte Original / Minute)
 * 2. ملف العقد المرجعي (Contrat de Référence)
 * 3. نسخة العقد (Expédition / Copie Authentique)
 * 4. المستخرج (Extrait d'Acte)
 * 5. شهادة البيع (Certificat de Vente)
 * 6. جدول الإرسال (Bordereau d'Envoi)
 *
 * All templates strictly utilize modern canonical `${}` and `${block}` loop architecture.
 */

import {
  downloadNotaryDocx,
  expandBlockLoop,
  mergeEditorVariables,
  mergePlaceholdersIntoHtml,
  mergeTemplateSymbols,
} from './docx-engine';
import { normalizeAllNotaryFieldValues } from './field-values-normalizer';
import { purifyNotarialBodyHtml } from './notarial-body-purifier';
import { expandFieldValuesWithStandardLatinTags } from './notary-word-variables';
import { STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';

export type StandardNotaryTemplateId =
  | 'acte_original'
  | 'contrat_reference'
  | 'expedition'
  | 'extrait'
  | 'certificat_vente'
  | 'bordereau_envoi';

export interface StandardNotaryTemplateConfig {
  id: StandardNotaryTemplateId;
  name: string;
  frenchTitle: string;
  category: string;
  description: string;
  defaultFileName: string;
  headerHtml: string;
  bodyHtml: string;
  footerHtml: string;
  pageNumberingEnabled: boolean;
  supportedTags: string[];
}

const pBase = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;`;
const pCenterBold = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:14pt;text-align:center;font-weight:bold;`;
const pCenter = `margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:center;`;
const tableStyle = `width:100%;border-collapse:collapse;margin:4px 0;`;
const thStyle = `border:1px solid #000000;background-color:#F1F5F9;padding:3px 6px;font-weight:bold;font-size:12pt;text-align:center;`;
const tdStyle = `border:1px solid #000000;padding:3px 6px;font-size:12pt;`;

const registrationBoxHtml = `
<table dir="rtl" style="width:100%;border-collapse:collapse;margin:0 0 10px 0;">
  <tr>
    <td style="border:1px solid #000000;padding:6px 10px;height:85px;vertical-align:top;text-align:right;">
      <p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:11pt;font-weight:bold;text-align:right;">إطار خاص بالتسجيل</p>
      <p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:10pt;color:#555555;text-align:right;"><br><br></p>
    </td>
  </tr>
</table>
`.trim();

export const BUILTIN_STANDARD_WORD_TEMPLATES: Record<
  StandardNotaryTemplateId,
  StandardNotaryTemplateConfig
> = {
  // 1. أصل العقد (Acte Original / Minute)
  acte_original: {
    id: 'acte_original',
    name: 'أصل العقد',
    frenchTitle: 'Acte Original / Minute',
    category: 'أصل العقد',
    description:
      'العقد التوثيقي الأصلي الكامل المحفوظ بأرشيف مكتب التوثيق مع إطار التسجيل بالصفحة الأولى وهوامش التوثيق الرسمية وعلامات الترقيم التوثيقية الجزائرية.',
    defaultFileName: 'قالب_أصل_العقد_النموذجي',
    headerHtml: '',
    bodyHtml: `
      ${registrationBoxHtml}
      <p dir="rtl" style="${pCenterBold}">أصل عقد توثيقي رسمي</p>
      <p dir="rtl" style="${pCenter}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">وزارة العدل — الغرفة الوطنية للموثقين</p>
      <p dir="rtl" style="${pCenter}"><span style="font-weight:bold;">مكتب الأستاذ:</span> \${office_name} — ولاية \${office_wilaya}</p>
      <p dir="rtl" style="${pCenter}"><span style="font-weight:bold;">فهرس رقم:</span> \${num_acte} — <span style="font-weight:bold;">بتاريخ:</span> \${date_acte}</p>
      <p dir="rtl" style="${pCenter}">----------------------------------------------------------</p>
      <p dir="rtl" style="${pBase}">في سنة \${date2lettre_acte}، الموافق لـ \${date_acte}.</p>
      <p dir="rtl" style="${pBase}">أمامنا نحن الأستاذ \${office_name}، الموثق بمكتب التوثيق الكائن مقره بـ \${office_addr}، ولاية \${office_wilaya}.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-template-part="parties" class="parties-wrapper">
        <p dir="rtl" style="${pCenterBold}">حـــــــــــــضـــــــــــــر</p>
        <p dir="rtl" style="${pBase}"><br></p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الأول (المتصرف):</span> \${c_1}.</p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الثاني (المتصرف إليه):</span> \${c_2}.</p>
        <p dir="rtl" style="${pBase}"><br></p>
      </div>
      <div data-clause-container="true">
        \${block}
        <p dir="rtl" style="font-weight:bold; margin-bottom:4pt; text-align:justify;">\${titre}</p>
        <div dir="rtl" style="margin:0; line-height:1; text-align:justify;">\${clause}</div>
        \${/block}
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">إثبات التلاوة والإشهاد:</span> تمت تلاوة هذا العقد على الأطراف الحاضرين، فصرحوا بفهمهم لجميع بنوده ووافقوا عليها بحضور الموثق الموقع أسفله، وحرر على أصل واحد محفوظ بأرشيف المكتب تحت رقم الفهرس \${num_acte}.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;"><p dir="rtl" style="${pBase}">توقيع الطرف الأول</p><br><br></td>
          <td style="${tdStyle};width:50%;text-align:center;"><p dir="rtl" style="${pBase}">توقيع وخاتم الموثق</p><br><br></td>
        </tr>
      </table>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'office_name',
      'office_addr',
      'office_wilaya',
      'office_tel',
      'office_mail',
      'num_acte',
      'c_1',
      'c_2',
      'date_acte',
      'date2lettre_acte',
      'act_value',
      'titre',
      'clause',
      'block',
    ],
  },

  // 2. ملف العقد المرجعي (Contrat de Référence)
  contrat_reference: {
    id: 'contrat_reference',
    name: 'ملف العقد المرجعي',
    frenchTitle: 'Contrat de Référence',
    category: 'ملف العقد المرجعي',
    description:
      'ملف العقد المرجعي النموذجي الشامل الذي يُنسخ تنسيقه وهوية المكتب تلقائياً لباقي العقود، مع إطار التسجيل وجداول الأطراف والتعيين المرجعي.',
    defaultFileName: 'قالب_ملف_العقد_المرجعي_النموذجي',
    headerHtml: '',
    bodyHtml: `
      ${registrationBoxHtml}
      <p dir="rtl" style="${pCenterBold}">بطاقة ملف العقد المرجعي الشامل</p>
      <p dir="rtl" style="${pCenter}">الجمهورية الجزائرية الديمقراطية الشعبية — مكتب التوثيق</p>
      <p dir="rtl" style="${pCenter}">مكتب الأستاذ: \${office_name} — ولاية \${office_wilaya}</p>
      <p dir="rtl" style="${pCenter}">----------------------------------------------------------</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:25%;">رقم الفهرس المرجعي</td>
          <td style="${tdStyle};width:25%;">\${num_acte}</td>
          <td style="${thStyle};width:25%;">تاريخ التحرير</td>
          <td style="${tdStyle};width:25%;">\${date_acte}</td>
        </tr>
        <tr>
          <td style="${thStyle};width:25%;">الموثق المحرر</td>
          <td style="${tdStyle};width:25%;">\${office_name}</td>
          <td style="${thStyle};width:25%;">مقر المكتب والولاية</td>
          <td style="${tdStyle};width:25%;">\${office_addr} / \${office_wilaya}</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-template-part="parties" class="parties-wrapper">
        <p dir="rtl" style="${pCenterBold}">جدول الأطراف المرجعي</p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الأول:</span> \${c_1}</p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الثاني:</span> \${c_2}</p>
        <p dir="rtl" style="${pBase}"><br></p>
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        \${block}
        <p dir="rtl" style="font-weight:bold; margin-bottom:4pt; text-align:justify;">\${titre}</p>
        <div dir="rtl" style="margin:0; line-height:1; text-align:justify;">\${clause}</div>
        \${/block}
      </div>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'num_acte',
      'date_acte',
      'office_name',
      'office_addr',
      'office_wilaya',
      'c_1',
      'c_2',
      'block',
      'titre',
      'clause',
    ],
  },

  // 3. نسخة العقد (Expédition / Copie Authentique)
  expedition: {
    id: 'expedition',
    name: 'نسخة العقد',
    frenchTitle: 'Expédition / Copie Authentique',
    category: 'نسخة العقد',
    description:
      'النسخة الرسمية المسلمة للأطراف مع إطار التسجيل وديباجة النسخة والمطابقة للأصل، والصيغة التنفيذية أو التأشيرة الرسمية.',
    defaultFileName: 'قالب_نسخة_العقد_النموذجي',
    headerHtml: '',
    bodyHtml: `
      ${registrationBoxHtml}
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">باسم الشعب الجزائري — نسخة رسمية مطابقة للأصل</p>
      <p dir="rtl" style="${pCenter}">مكتب الأستاذ: \${office_name} — فهرس رقم: \${num_acte}</p>
      <p dir="rtl" style="${pCenter}">----------------------------------------------------------</p>
      <p dir="rtl" style="${pCenterBold}">نسخة رسمية من عقد توثيقي</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-template-part="parties" class="parties-wrapper">
        <p dir="rtl" style="${pBase}">نحن الأستاذ \${office_name}، الموثق بمكتب التوثيق بـ \${office_addr}، ولاية \${office_wilaya}.</p>
        <p dir="rtl" style="${pBase}">نشهد ونسلم هذه النسخة الرسمية المطابقة تماماً لأصل العقد المحفوظ بأرشيف مكتبنا، المحرر بتاريخ \${date_acte}، والمقيد بالفهرس الرسمي تحت رقم \${num_acte}، والذي نصّه الآتي حرفياً:</p>
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        \${block}
        <p dir="rtl" style="font-weight:bold; margin-bottom:4pt; text-align:justify;">\${titre}</p>
        <div dir="rtl" style="margin:0; line-height:1; text-align:justify;">\${clause}</div>
        \${/block}
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">شهادة المطابقة للأصل والصيغة التنفيذية:</span></p>
      <p dir="rtl" style="${pBase}">أشهد أنا الموثق الموقع أسفله أن هذه النسخة مطابقة حرفياً وكلياً لأصل العقد المحفوظ بمكتبنا دون زيادة أو نقصان، وسلّمت للمعني بالأمر لاستعمالها فيما يفرضه القانون وحفظ حقوقه.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">حرر بـ \${office_addr} في \${date_acte}</p>
            <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">توقيع وخاتم الموثق</span></p>
            <br><br>
          </td>
        </tr>
      </table>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'office_name',
      'office_addr',
      'office_wilaya',
      'num_acte',
      'date_acte',
      'block',
      'titre',
      'clause',
    ],
  },

  // 4. المستخرج (Extrait d'Acte)
  extrait: {
    id: 'extrait',
    name: 'المستخرج',
    frenchTitle: "Extrait d'Acte",
    category: 'المستخرج',
    description:
      'مستخرج العقد الموجه لمفتشية التسجيل والضرائب أو المحافظة العقارية، متضمناً إطار التسجيل والأطراف والثمن والتأشيرات.',
    defaultFileName: 'قالب_المستخرج_النموذجي',
    headerHtml: '',
    bodyHtml: `
      ${registrationBoxHtml}
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">وزارة المالية — المديرية العامة للضرائب</p>
      <p dir="rtl" style="${pCenter}">مستخرج عقد توثيقي لأجل إجراءات التسجيل والطابع</p>
      <p dir="rtl" style="${pCenter}">----------------------------------------------------------</p>
      <p dir="rtl" style="${pCenterBold}">مستخرج عقد توثيقي</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:25%;">رقم الفهرس الرسمي</td>
          <td style="${tdStyle};width:25%;">\${num_acte}</td>
          <td style="${thStyle};width:25%;">تاريخ إبرام العقد</td>
          <td style="${tdStyle};width:25%;">\${date_acte}</td>
        </tr>
        <tr>
          <td style="${thStyle};width:25%;">الموثق المحرر</td>
          <td style="${tdStyle};width:25%;">\${office_name}</td>
          <td style="${thStyle};width:25%;">مقر المكتب والولاية</td>
          <td style="${tdStyle};width:25%;">\${office_addr} / \${office_wilaya}</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-template-part="parties" class="parties-wrapper">
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">1. الطرف الأول (المتصرف):</span> \${c_1}.</p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">2. الطرف الثاني (المتصرف إليه):</span> \${c_2}.</p>
        <p dir="rtl" style="${pBase}"><br></p>
      </div>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">3. الثمن والشروط المالية:</span> تم التصرف بمبلغ إجمالي قدره \${act_value} دج.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        \${block}
        <p dir="rtl" style="font-weight:bold; margin-bottom:4pt; text-align:justify;">\${titre}</p>
        <div dir="rtl" style="margin:0; line-height:1; text-align:justify;">\${clause}</div>
        \${/block}
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}">سلم هذا المستخرج لمفتشية التسجيل والطابع لاستيفاء الحقوق والرسوم المقررة قانوناً.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">تأشيرة مفتشية التسجيل</p><br><br>
          </td>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">خاتم وتوقيع الموثق</p><br><br>
          </td>
        </tr>
      </table>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'num_acte',
      'date_acte',
      'office_name',
      'office_addr',
      'office_wilaya',
      'c_1',
      'c_2',
      'act_value',
      'block',
      'titre',
      'clause',
    ],
  },

  // 5. شهادة البيع (Certificat / Attestation de Vente)
  certificat_vente: {
    id: 'certificat_vente',
    name: 'شهادة البيع',
    frenchTitle: 'Certificat / Attestation de Vente',
    category: 'شهادة البيع',
    description:
      'شهادة إدارية رسمية تثبت واقعة البيع أو نقل الملكية للأطراف المعنية للإدلاء بها لدى الهيئات الإدارية والشركات مع إطار التسجيل.',
    defaultFileName: 'قالب_شهادة_البيع_النموذجي',
    headerHtml: '',
    bodyHtml: `
      ${registrationBoxHtml}
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">وزارة العدل — مكتب التوثيق</p>
      <p dir="rtl" style="${pCenter}">مكتب الأستاذ: \${office_name} — الموثق بـ \${office_addr}</p>
      <p dir="rtl" style="${pCenter}">----------------------------------------------------------</p>
      <p dir="rtl" style="${pCenterBold}">شهادة توثيقية بالبيع والتصرف</p>
      <p dir="rtl" style="${pCenterBold}">Attestation Notariée de Vente</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-template-part="parties" class="parties-wrapper">
        <p dir="rtl" style="${pBase}">يشهد الأستاذ \${office_name}، الموثق بمكتب التوثيق الكائن مقره بـ \${office_addr}، ولاية \${office_wilaya}.</p>
        <p dir="rtl" style="${pBase}">أنه بموجب عقد توثيقي رسمي تم تحريره وإبرامه بمكتبنا بتاريخ \${date_acte}، والمقيد بسجلات الفهرس الرسمي تحت رقم \${num_acte}:</p>
        <p dir="rtl" style="${pBase}"><br></p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الأول:</span> \${c_1}.</p>
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الثاني:</span> \${c_2}.</p>
        <p dir="rtl" style="${pBase}"><br></p>
      </div>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الثمن الإجمالي للبيع:</span> \${act_value} دج.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        \${block}
        <p dir="rtl" style="font-weight:bold; margin-bottom:4pt; text-align:justify;">\${titre}</p>
        <div dir="rtl" style="margin:0; line-height:1; text-align:justify;">\${clause}</div>
        \${/block}
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}">سلمت هذه الشهادة للمعني بالأمر بناءً على طلبه لاستعمالها والإدلاء بها لدى الهيئات والإدارات المختصة في حدود ما يقتضيه القانون والتنظيم المعمول به.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">حرر بـ \${office_addr} في: \${date_acte}</p>
          </td>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">خاتم وتوقيع الموثق</span></p><br><br>
          </td>
        </tr>
      </table>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'office_name',
      'office_addr',
      'office_wilaya',
      'date_acte',
      'num_acte',
      'c_1',
      'c_2',
      'act_value',
      'block',
      'titre',
      'clause',
    ],
  },

  // 6. جدول الإرسال (Bordereau d'Envoi / Dépôt)
  bordereau_envoi: {
    id: 'bordereau_envoi',
    name: 'جدول الإرسال',
    frenchTitle: "Bordereau d'Envoi / Dépôt",
    category: 'جدول الإرسال',
    description:
      'جدول إرسال وإيداع المعاملات والعقود الرسمية لدى مفتشية التسجيل أو المحافظة العقارية مع بيان الوثائق المرفقة.',
    defaultFileName: 'قالب_جدول_الإرسال_النموذجي',
    headerHtml: '',
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">مكتب التوثيق — الأستاذ: \${office_name}</p>
      <p dir="rtl" style="${pCenter}">العنوان: \${office_addr}، ولاية \${office_wilaya}</p>
      <p dir="rtl" style="${pCenter}">----------------------------------------------------------</p>
      <p dir="rtl" style="${pCenterBold}">جدول إرسال وإيداع الوثائق التوثيقية</p>
      <p dir="rtl" style="${pCenterBold}">Bordereau d'Envoi Notarié</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-template-part="parties" class="parties-wrapper">
        <table dir="rtl" style="${tableStyle}">
          <tr>
            <td style="${thStyle};width:30%;">الجهة المرسل إليها:</td>
            <td style="${tdStyle};width:70%;">السيد مفتش التسجيل والطابع / السيد المحافظ العقاري</td>
          </tr>
          <tr>
            <td style="${thStyle};width:30%;">تاريخ الإرسال:</td>
            <td style="${tdStyle};width:70%;">\${date_acte}</td>
          </tr>
          <tr>
            <td style="${thStyle};width:30%;">رقم وتاريخ العقد:</td>
            <td style="${tdStyle};width:70%;">فهرس رقم \${num_acte} بتاريخ \${date_acte}</td>
          </tr>
          <tr>
            <td style="${thStyle};width:30%;">الأطراف:</td>
            <td style="${tdStyle};width:70%;">\${c_1} و \${c_2}</td>
          </tr>
        </table>
        <p dir="rtl" style="${pBase}"><br></p>
      </div>
      <p dir="rtl" style="${pCenterBold}">بيان المعاملة والوثائق المرفقة</p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:10%;">الرقم</td>
          <td style="${thStyle};width:45%;">بيان الوثائق المودعة</td>
          <td style="${thStyle};width:15%;">عدد النسخ</td>
          <td style="${thStyle};width:30%;">ملاحظات وتأشيرات</td>
        </tr>
        <tr>
          <td style="${tdStyle};text-align:center;">01</td>
          <td style="${tdStyle};"><span style="font-weight:bold;">أصل العقد التوثيقي:</span> بيع عقاري وموثق</td>
          <td style="${tdStyle};text-align:center;">01</td>
          <td style="${tdStyle};">لأجل التسجيل والشهر</td>
        </tr>
        <tr>
          <td style="${tdStyle};text-align:center;">02</td>
          <td style="${tdStyle};"><span style="font-weight:bold;">مستخرج العقد:</span> مخصص لمصالح الضرائب</td>
          <td style="${tdStyle};text-align:center;">02</td>
          <td style="${tdStyle};">استيفاء الحقوق والرسوم</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        \${block}
        <p dir="rtl" style="font-weight:bold; margin-bottom:4pt; text-align:justify;">\${titre}</p>
        <div dir="rtl" style="margin:0; line-height:1; text-align:justify;">\${clause}</div>
        \${/block}
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">وصل استلام الجهة المودع لديها:</p>
            <br><br>
          </td>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">توقيع وخاتم الموثق المرسل:</p>
            <br><br>
          </td>
        </tr>
      </table>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'office_name',
      'office_addr',
      'office_wilaya',
      'date_acte',
      'num_acte',
      'c_1',
      'c_2',
      'block',
      'titre',
      'clause',
    ],
  },
};

const STORAGE_CUSTOM_TEMPLATE_PREFIX = 'notaire_custom_docx_tpl_';

/**
 * Checks if a customized template exists in browser storage.
 */
export function hasCustomizedNotaryTemplate(
  id: StandardNotaryTemplateId
): boolean {
  if (typeof window === 'undefined') return false;
  return Boolean(
    window.localStorage.getItem(`${STORAGE_CUSTOM_TEMPLATE_PREFIX}${id}`)
  );
}

/**
 * Saves a user-customized .docx template file as Base64 in browser storage.
 */
export async function saveCustomizedNotaryTemplateFile(
  id: StandardNotaryTemplateId,
  file: File
): Promise<void> {
  if (typeof window === 'undefined') return;

  const arrayBuffer = await file.arrayBuffer();
  let binary = '';
  const bytes = new Uint8Array(arrayBuffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  const base64 = window.btoa(binary);

  const payload = JSON.stringify({
    fileName: file.name,
    fileSize: file.size,
    updatedAt: new Date().toISOString(),
    base64Data: base64,
  });

  window.localStorage.setItem(`${STORAGE_CUSTOM_TEMPLATE_PREFIX}${id}`, payload);
}

/**
 * Gets customized template metadata if uploaded.
 */
export function getCustomizedNotaryTemplateMeta(
  id: StandardNotaryTemplateId
): {
  fileName: string;
  fileSize: number;
  updatedAt: string;
} | null {
  if (typeof window === 'undefined') return null;
  const raw = window.localStorage.getItem(
    `${STORAGE_CUSTOM_TEMPLATE_PREFIX}${id}`
  );
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return {
      fileName: parsed.fileName || 'قالب_مخصص.docx',
      fileSize: parsed.fileSize || 0,
      updatedAt: parsed.updatedAt || '',
    };
  } catch {
    return null;
  }
}

/**
 * Resets a customized template back to standard built-in.
 */
export function resetNotaryTemplateToDefault(
  id: StandardNotaryTemplateId
): void {
  if (typeof window === 'undefined') return;
  window.localStorage.removeItem(`${STORAGE_CUSTOM_TEMPLATE_PREFIX}${id}`);
}

/**
 * Downloads the clean, unmerged official standard .docx template
 * for any of the 6 templates, allowing the notary to inspect or customize it.
 */
export async function downloadDefaultNotaryDocxTemplate(
  id: StandardNotaryTemplateId
): Promise<void> {
  const config = BUILTIN_STANDARD_WORD_TEMPLATES[id];
  if (!config) return;

  await downloadNotaryDocx({
    title: config.defaultFileName,
    bodyHtml: config.bodyHtml,
    headerHtml: '',
    footerHtml: config.footerHtml,
    pageNumberingEnabled: config.pageNumberingEnabled,
    fieldValues: {}, // keep [tags] unmerged as tags in the template
    omitRepeatingHeader: true,
  });
}

/**
 * Exports a contract using the specified standard Word template.
 */
export async function exportContractWithStandardTemplate(params: {
  templateId: StandardNotaryTemplateId;
  contractTitle: string;
  bodyHtml: string;
  fieldValues: Record<string, string>;
}): Promise<{ fileName: string; templateName: string }> {
  const { templateId, contractTitle, bodyHtml, fieldValues } = params;
  const config =
    BUILTIN_STANDARD_WORD_TEMPLATES[templateId] ||
    BUILTIN_STANDARD_WORD_TEMPLATES.acte_original;

  const normalizedVals = normalizeAllNotaryFieldValues(fieldValues);
  const activeVals = expandFieldValuesWithStandardLatinTags(normalizedVals);

  const mergedBody = expandBlockLoop(
    config.bodyHtml,
    [],
    purifyNotarialBodyHtml(bodyHtml)
  );

  const finalBody = mergeEditorVariables(mergedBody, activeVals);
  const finalHeader = mergeEditorVariables(config.headerHtml, activeVals);
  const finalFooter = mergeEditorVariables(config.footerHtml, activeVals);

  const fileName = `${contractTitle || config.defaultFileName}.docx`;
  await downloadNotaryDocx({
    title: contractTitle || config.defaultFileName,
    bodyHtml: finalBody,
    headerHtml: finalHeader,
    footerHtml: finalFooter,
    pageNumberingEnabled: config.pageNumberingEnabled,
    fieldValues: activeVals,
    omitRepeatingHeader: true,
  });

  return {
    fileName,
    templateName: config.name,
  };
}
