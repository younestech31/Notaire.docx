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
 * All templates strictly utilize standard Latin placeholder tags
 * ([partie1_nom], [date_acte], [num_repertoire], [contenu_acte], etc.)
 * in Word documents and support instant download, customization via .docx upload,
 * and direct export with dynamic data injection.
 */

import {
  downloadNotaryDocx,
  mergePlaceholdersIntoHtml,
} from './docx-engine';
import { normalizeAllNotaryFieldValues } from './field-values-normalizer';
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
      'العقد التوثيقي الأصلي الكامل المحفوظ بأرشيف مكتب التوثيق مع هوامش التوثيق الرسمية وعلامات الترقيم التوثيقية الجزائرية.',
    defaultFileName: 'قالب_أصل_العقد_النموذجي',
    headerHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">وزارة العدل — الغرفة الوطنية للموثقين</p>
      <p dir="rtl" style="${pCenter}"><span style="font-weight:bold;">مكتب الأستاذ:</span> [notaire_nom] — الموثق لدى محكمة [etude_commune] — ولاية [etude_wilaya]</p>
      <p dir="rtl" style="${pCenter}"><span style="font-weight:bold;">فهرس رقم:</span> [num_repertoire] / [annee_acte] — <span style="font-weight:bold;">بتاريخ:</span> [date_acte]</p>
      <hr style="border-top:1px solid #000000;margin:4px 0;"/>
    `.trim(),
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">أصل عقد توثيقي رسمي</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}">في سنة [date_acte_lettres]، الموافق لـ [date_acte].</p>
      <p dir="rtl" style="${pBase}">أمامنا نحن الأستاذ [notaire_nom]، الموثق بمكتب التوثيق الكائن مقره بـ [etude_commune]، ولاية [etude_wilaya].</p>
      <p dir="rtl" style="${pBase}">حضر الأطراف الآتي بيانهم:</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الأول (المتصرف):</span> السيد [partie1_nom_complet]، المولود بتاريخ [partie1_date_naissance] بـ [partie1_lieu_naissance]، [partie1_filiation]، الحامل للرقم الوطني [partie1_nin]، المقيم بـ [partie1_adresse]، الحامل لوثيقة الهوية: [partie1_cni]، المهنة: [partie1_profession].</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الطرف الثاني (المتصرف إليه):</span> السيد [partie2_nom_complet]، المولود بتاريخ [partie2_date_naissance] بـ [partie2_lieu_naissance]، [partie2_filiation]، الحامل للرقم الوطني [partie2_nin]، المقيم بـ [partie2_adresse]، الحامل لوثيقة الهوية: [partie2_cni]، المهنة: [partie2_profession].</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;"># متن وشروط العقد</span></p>
        <p dir="rtl" style="${pBase}">[contenu_acte]</p>
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">إثبات التلاوة والإشهاد:</span> تمت تلاوة هذا العقد على الأطراف الحاضرين، فصرحوا بفهمهم لجميع بنوده ووافقوا عليها بحضور الموثق الموقع أسفله، وحرر على أصل واحد محفوظ بأرشيف المكتب تحت رقم الفهرس [num_repertoire].</p>
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
      'notaire_nom',
      'etude_commune',
      'etude_wilaya',
      'num_repertoire',
      'annee_acte',
      'date_acte',
      'date_acte_lettres',
      'partie1_nom_complet',
      'partie1_date_naissance',
      'partie1_lieu_naissance',
      'partie1_filiation',
      'partie1_nin',
      'partie1_adresse',
      'partie1_cni',
      'partie1_profession',
      'partie2_nom_complet',
      'partie2_date_naissance',
      'partie2_lieu_naissance',
      'partie2_filiation',
      'partie2_nin',
      'partie2_adresse',
      'partie2_cni',
      'partie2_profession',
      'contenu_acte',
    ],
  },

  // 2. ملف العقد المرجعي (Contrat de Référence)
  contrat_reference: {
    id: 'contrat_reference',
    name: 'ملف العقد المرجعي',
    frenchTitle: 'Contrat de Référence',
    category: 'ملف العقد المرجعي',
    description:
      'ملف العقد المرجعي النموذجي الشامل الذي يُنسخ تنسيقه وهوية المكتب تلقائياً لباقي العقود، مع جداول الأطراف والتعيين المرجعي.',
    defaultFileName: 'قالب_ملف_العقد_المرجعي_النموذجي',
    headerHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية — مكتب التوثيق</p>
      <p dir="rtl" style="${pCenter}">الملف المرجعي التوثيقي الموحد للعقد | Référence Notariale</p>
      <hr style="border-top:1px solid #000000;margin:4px 0;"/>
    `.trim(),
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">بطاقة ملف العقد المرجعي الشامل</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:25%;">رقم الفهرس المرجعي</td>
          <td style="${tdStyle};width:25%;">[num_repertoire]</td>
          <td style="${thStyle};width:25%;">تاريخ التحرير</td>
          <td style="${tdStyle};width:25%;">[date_acte]</td>
        </tr>
        <tr>
          <td style="${thStyle};width:25%;">الموثق المحرر</td>
          <td style="${tdStyle};width:25%;">[notaire_nom]</td>
          <td style="${thStyle};width:25%;">مقر المكتب والولاية</td>
          <td style="${tdStyle};width:25%;">[etude_commune] / [etude_wilaya]</td>
        </tr>
        <tr>
          <td style="${thStyle};width:25%;">مفتشية التسجيل</td>
          <td style="${tdStyle};width:25%;">[inspection_enregistrement]</td>
          <td style="${thStyle};width:25%;">المحافظة العقارية</td>
          <td style="${tdStyle};width:25%;">[conservation_fonciere]</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pCenterBold}">جدول الأطراف المرجعي</p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:20%;">الصفة</td>
          <td style="${thStyle};width:30%;">الاسم واللقب والنسب</td>
          <td style="${thStyle};width:25%;">الميلاد ورقم التعريف (NIN)</td>
          <td style="${thStyle};width:25%;">العنوان ووثيقة الهوية</td>
        </tr>
        <tr>
          <td style="${tdStyle};font-weight:bold;">الطرف الأول</td>
          <td style="${tdStyle};">[partie1_nom_complet]<br>[partie1_filiation]</td>
          <td style="${tdStyle};">[partie1_date_naissance] بـ [partie1_lieu_naissance]<br>NIN: [partie1_nin]</td>
          <td style="${tdStyle};">[partie1_adresse]<br>[partie1_cni]</td>
        </tr>
        <tr>
          <td style="${tdStyle};font-weight:bold;">الطرف الثاني</td>
          <td style="${tdStyle};">[partie2_nom_complet]<br>[partie2_filiation]</td>
          <td style="${tdStyle};">[partie2_date_naissance] بـ [partie2_lieu_naissance]<br>NIN: [partie2_nin]</td>
          <td style="${tdStyle};">[partie2_adresse]<br>[partie2_cni]</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pCenterBold}">جدول التعيين العقاري والبيانات المالية</p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:30%;">التعيين العقاري الكامل</td>
          <td style="${tdStyle};width:70%;">[designation_bien] (الحصة رقم [lot_numero]، بمساحة [lot_superficie] م²، الأجزاء المشتركة: [lot_quote_part])</td>
        </tr>
        <tr>
          <td style="${thStyle};width:30%;">الثمن المالي المتفق عليه</td>
          <td style="${tdStyle};width:70%;"><span style="font-weight:bold;">[prix_chiffres] دج</span> ([prix_lettres])</td>
        </tr>
        <tr>
          <td style="${thStyle};width:30%;">مراجع الوصف التقسيمي</td>
          <td style="${tdStyle};width:70%;">[ref_division]</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;"># متن العقد والبنود المنسوخة تلقائياً</span></p>
        <p dir="rtl" style="${pBase}">[contenu_acte]</p>
      </div>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'num_repertoire',
      'date_acte',
      'notaire_nom',
      'etude_commune',
      'etude_wilaya',
      'inspection_enregistrement',
      'conservation_fonciere',
      'partie1_nom_complet',
      'partie1_filiation',
      'partie1_date_naissance',
      'partie1_lieu_naissance',
      'partie1_nin',
      'partie1_adresse',
      'partie1_cni',
      'partie2_nom_complet',
      'partie2_filiation',
      'partie2_date_naissance',
      'partie2_lieu_naissance',
      'partie2_nin',
      'partie2_adresse',
      'partie2_cni',
      'designation_bien',
      'lot_numero',
      'lot_superficie',
      'lot_quote_part',
      'prix_chiffres',
      'prix_lettres',
      'ref_division',
      'contenu_acte',
    ],
  },

  // 3. نسخة العقد (Expédition / Copie Authentique)
  expedition: {
    id: 'expedition',
    name: 'نسخة العقد',
    frenchTitle: 'Expédition / Copie Authentique',
    category: 'نسخة العقد',
    description:
      'النسخة الرسمية المسلمة للأطراف مع ديباجة النسخة والمطابقة للأصل، والصيغة التنفيذية أو التأشيرة الرسمية.',
    defaultFileName: 'قالب_نسخة_العقد_النموذجي',
    headerHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">باسم الشعب الجزائري — نسخة رسمية مطابقة للأصل</p>
      <p dir="rtl" style="${pCenter}">مكتب الأستاذ: [notaire_nom] — فهرس رقم: [num_repertoire]</p>
      <hr style="border-top:1px solid #000000;margin:4px 0;"/>
    `.trim(),
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">نسخة رسمية من عقد توثيقي</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}">نحن الأستاذ [notaire_nom]، الموثق بمكتب التوثيق بـ [etude_commune]، ولاية [etude_wilaya].</p>
      <p dir="rtl" style="${pBase}">نشهد ونسلم هذه النسخة الرسمية المطابقة تماماً لأصل العقد المحفوظ بأرشيف مكتبنا، المحرر بتاريخ [date_acte]، والمقيد بالفهرس الرسمي تحت رقم [num_repertoire] لسنة [annee_acte]، والذي نصّه الآتي حرفياً:</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <div data-clause-container="true">
        <p dir="rtl" style="${pBase}"><span style="font-weight:bold;"># متن العقد المسلّم نسخته</span></p>
        <p dir="rtl" style="${pBase}">[contenu_acte]</p>
      </div>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">شهادة المطابقة للأصل والصيغة التنفيذية:</span></p>
      <p dir="rtl" style="${pBase}">أشهد أنا الموثق الموقع أسفله أن هذه النسخة مطابقة حرفياً وكلياً لأصل العقد المحفوظ بمكتبنا دون زيادة أو نقصان، وسلّمت للمعني بالأمر لاستعمالها فيما يفرضه القانون وحفظ حقوقه.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">مراجع التسجيل والشهر:</p>
            <p dir="rtl" style="${pBase}">[ref_enregistrement] — [ref_publicite]</p>
          </td>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">حرر بـ [etude_commune] في [date_acte]</p>
            <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">توقيع وخاتم الموثق</span></p>
            <br><br>
          </td>
        </tr>
      </table>
    `.trim(),
    footerHtml: ``,
    pageNumberingEnabled: true,
    supportedTags: [
      'notaire_nom',
      'num_repertoire',
      'etude_commune',
      'etude_wilaya',
      'date_acte',
      'annee_acte',
      'contenu_acte',
      'ref_enregistrement',
      'ref_publicite',
    ],
  },

  // 4. المستخرج (Extrait d'Acte)
  extrait: {
    id: 'extrait',
    name: 'المستخرج',
    frenchTitle: "Extrait d'Acte",
    category: 'المستخرج',
    description:
      'مستخرج العقد الموجه لمفتشية التسجيل والضرائب أو المحافظة العقارية، متضمناً الأطراف والعقار والثمن والتأشيرات.',
    defaultFileName: 'قالب_المستخرج_النموذجي',
    headerHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">وزارة المالية — المديرية العامة للضرائب</p>
      <p dir="rtl" style="${pCenter}">مستخرج عقد توثيقي لأجل إجراءات التسجيل والطابع</p>
      <hr style="border-top:1px solid #000000;margin:4px 0;"/>
    `.trim(),
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">مستخرج عقد توثيقي</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:25%;">رقم الفهرس الرسمي</td>
          <td style="${tdStyle};width:25%;">[num_repertoire]</td>
          <td style="${thStyle};width:25%;">تاريخ إبرام العقد</td>
          <td style="${tdStyle};width:25%;">[date_acte]</td>
        </tr>
        <tr>
          <td style="${thStyle};width:25%;">الموثق المحرر</td>
          <td style="${tdStyle};width:25%;">[notaire_nom]</td>
          <td style="${thStyle};width:25%;">مفتشية التسجيل</td>
          <td style="${tdStyle};width:25%;">[inspection_enregistrement]</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">1. الطرف الأول (المتصرف):</span> [partie1_nom_complet]، المولود بـ [partie1_lieu_naissance] في [partie1_date_naissance]، [partie1_filiation]، رقم التعريف الوطني (NIN): [partie1_nin]، المقيم بـ [partie1_adresse].</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">2. الطرف الثاني (المتصرف إليه):</span> [partie2_nom_complet]، المولود بـ [partie2_lieu_naissance] في [partie2_date_naissance]، [partie2_filiation]، رقم التعريف الوطني (NIN): [partie2_nin]، المقيم بـ [partie2_adresse].</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">3. التعيين العقاري:</span> [designation_bien] (الحصة رقم [lot_numero]، طبيعتها: [lot_nature]، المساحة: [lot_superficie] م²، الأجزاء المشتركة: [lot_quote_part]).</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">4. الثمن والشروط المالية:</span> تم التصرف بمبلغ إجمالي قدره [prix_chiffres] دج ([prix_lettres]).</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">5. مراجع التسجيل السابقة أو الشهر:</span> [ref_enregistrement] — [ref_publicite].</p>
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
      'num_repertoire',
      'date_acte',
      'notaire_nom',
      'inspection_enregistrement',
      'partie1_nom_complet',
      'partie1_lieu_naissance',
      'partie1_date_naissance',
      'partie1_filiation',
      'partie1_nin',
      'partie1_adresse',
      'partie2_nom_complet',
      'partie2_lieu_naissance',
      'partie2_date_naissance',
      'partie2_filiation',
      'partie2_nin',
      'partie2_adresse',
      'designation_bien',
      'lot_numero',
      'lot_nature',
      'lot_superficie',
      'lot_quote_part',
      'prix_chiffres',
      'prix_lettres',
      'ref_enregistrement',
      'ref_publicite',
    ],
  },

  // 5. شهادة البيع (Certificat / Attestation de Vente)
  certificat_vente: {
    id: 'certificat_vente',
    name: 'شهادة البيع',
    frenchTitle: 'Certificat / Attestation de Vente',
    category: 'شهادة البيع',
    description:
      'شهادة إدارية رسمية تثبت واقعة البيع أو نقل الملكية للأطراف المعنية للإدلاء بها لدى الهيئات الإدارية والشركات.',
    defaultFileName: 'قالب_شهادة_البيع_النموذجي',
    headerHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">وزارة العدل — مكتب التوثيق</p>
      <p dir="rtl" style="${pCenter}">مكتب الأستاذ: [notaire_nom] — الموثق بـ [etude_commune]</p>
      <hr style="border-top:1px solid #000000;margin:4px 0;"/>
    `.trim(),
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">شهادة توثيقية بالبيع والتصرف</p>
      <p dir="rtl" style="${pCenterBold}">Attestation Notariée de Vente</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}">يشهد الأستاذ [notaire_nom]، الموثق بمكتب التوثيق الكائن مقره بـ [etude_commune]، ولاية [etude_wilaya].</p>
      <p dir="rtl" style="${pBase}">أنه بموجب عقد توثيقي رسمي تم تحريره وإبرامه بمكتبنا بتاريخ [date_acte]، والمقيد بسجلات الفهرس الرسمي تحت رقم [num_repertoire] لسنة [annee_acte]:</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">قام السيد(ة):</span> [partie1_nom_complet]، المولود بتاريخ [partie1_date_naissance] بـ [partie1_lieu_naissance]، الحامل لوثيقة الهوية: [partie1_cni] والرقم الوطني [partie1_nin]، المقيم بـ [partie1_adresse].</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">بالبيع والتنازل النهائي لفائدة السيد(ة):</span> [partie2_nom_complet]، المولود بتاريخ [partie2_date_naissance] بـ [partie2_lieu_naissance]، الحامل لوثيقة الهوية: [partie2_cni] والرقم الوطني [partie2_nin]، المقيم بـ [partie2_adresse].</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">بخصوص العقار المعين كما يلي:</span></p>
      <p dir="rtl" style="${pBase}">[designation_bien] (الحصة رقم [lot_numero]، بمساحة قدرها [lot_superficie] م²، والأجزاء المشتركة [lot_quote_part]).</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">الثمن الإجمالي للبيع:</span> [prix_chiffres] دج ([prix_lettres]).</p>
      <p dir="rtl" style="${pBase}"><span style="font-weight:bold;">مراجع التسجيل:</span> [ref_enregistrement] لدى [inspection_enregistrement].</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <p dir="rtl" style="${pBase}">سلمت هذه الشهادة للمعني بالأمر بناءً على طلبه لاستعمالها والإدلاء بها لدى الهيئات والإدارات المختصة في حدود ما يقتضيه القانون والتنظيم المعمول به.</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${tdStyle};width:50%;text-align:center;">
            <p dir="rtl" style="${pBase}">حرر بـ [etude_commune] في: [date_acte]</p>
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
      'notaire_nom',
      'etude_commune',
      'etude_wilaya',
      'date_acte',
      'num_repertoire',
      'annee_acte',
      'partie1_nom_complet',
      'partie1_date_naissance',
      'partie1_lieu_naissance',
      'partie1_cni',
      'partie1_nin',
      'partie1_adresse',
      'partie2_nom_complet',
      'partie2_date_naissance',
      'partie2_lieu_naissance',
      'partie2_cni',
      'partie2_nin',
      'partie2_adresse',
      'designation_bien',
      'lot_numero',
      'lot_superficie',
      'lot_quote_part',
      'prix_chiffres',
      'prix_lettres',
      'ref_enregistrement',
      'inspection_enregistrement',
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
    headerHtml: `
      <p dir="rtl" style="${pCenterBold}">الجمهورية الجزائرية الديمقراطية الشعبية</p>
      <p dir="rtl" style="${pCenter}">مكتب التوثيق — الأستاذ: [notaire_nom]</p>
      <p dir="rtl" style="${pCenter}">العنوان: [etude_commune]، ولاية [etude_wilaya]</p>
      <hr style="border-top:1px solid #000000;margin:4px 0;"/>
    `.trim(),
    bodyHtml: `
      <p dir="rtl" style="${pCenterBold}">جدول إرسال وإيداع الوثائق التوثيقية</p>
      <p dir="rtl" style="${pCenterBold}">Bordereau d'Envoi Notarié</p>
      <p dir="rtl" style="${pBase}"><br></p>
      <table dir="rtl" style="${tableStyle}">
        <tr>
          <td style="${thStyle};width:30%;">الجهة المرسل إليها:</td>
          <td style="${tdStyle};width:70%;">السيد مفتش التسجيل والطابع لدى [inspection_enregistrement] / السيد المحافظ العقاري لدى [conservation_fonciere]</td>
        </tr>
        <tr>
          <td style="${thStyle};width:30%;">تاريخ الإرسال:</td>
          <td style="${tdStyle};width:70%;">[date_acte]</td>
        </tr>
        <tr>
          <td style="${thStyle};width:30%;">رقم وتاريخ العقد:</td>
          <td style="${tdStyle};width:70%;">فهرس رقم [num_repertoire] بتاريخ [date_acte]</td>
        </tr>
      </table>
      <p dir="rtl" style="${pBase}"><br></p>
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
          <td style="${tdStyle};"><span style="font-weight:bold;">أصل العقد التوثيقي:</span> بيع عقاري بين [partie1_nom] و [partie2_nom]</td>
          <td style="${tdStyle};text-align:center;">01</td>
          <td style="${tdStyle};">لأجل التسجيل والشهر</td>
        </tr>
        <tr>
          <td style="${tdStyle};text-align:center;">02</td>
          <td style="${tdStyle};"><span style="font-weight:bold;">مستخرج العقد:</span> مخصص لمصالح الضرائب</td>
          <td style="${tdStyle};text-align:center;">02</td>
          <td style="${tdStyle};">استيفاء الحقوق والرسوم</td>
        </tr>
        <tr>
          <td style="${tdStyle};text-align:center;">03</td>
          <td style="${tdStyle};"><span style="font-weight:bold;">جدول الوصف التقسيمي والسند:</span> مراجع [ref_division]</td>
          <td style="${tdStyle};text-align:center;">01</td>
          <td style="${tdStyle};">الحصة رقم [lot_numero]</td>
        </tr>
        <tr>
          <td style="${tdStyle};text-align:center;">04</td>
          <td style="${tdStyle};"><span style="font-weight:bold;">استمارة الشهر العقاري (PR4BIS):</span> بخصوص [conservation_fonciere]</td>
          <td style="${tdStyle};text-align:center;">02</td>
          <td style="${tdStyle};">لإشهار الحق العيني</td>
        </tr>
      </table>
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
      'notaire_nom',
      'etude_commune',
      'etude_wilaya',
      'inspection_enregistrement',
      'conservation_fonciere',
      'date_acte',
      'num_repertoire',
      'partie1_nom',
      'partie2_nom',
      'ref_division',
      'lot_numero',
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
    headerHtml: config.headerHtml,
    footerHtml: config.footerHtml,
    pageNumberingEnabled: config.pageNumberingEnabled,
    fieldValues: {}, // keep [tags] unmerged as tags in the template
  });
}

/**
 * Compiles a contract session into a chosen standard template,
 * replacing all Latin tags with contract data, and downloads the resulting .docx.
 */
export async function exportContractWithStandardTemplate(params: {
  templateId: StandardNotaryTemplateId;
  contractTitle: string;
  bodyHtml: string;
  fieldValues: Record<string, string>;
}): Promise<{
  fileName: string;
  templateName: string;
}> {
  const { templateId, contractTitle, bodyHtml, fieldValues } = params;
  const config = BUILTIN_STANDARD_WORD_TEMPLATES[templateId];
  if (!config) {
    throw new Error(`القالب "${templateId}" غير معروف في النظام`);
  }

  // 1. Fully normalize all field values across Arabic, symbols, and Latin tags
  const expandedValues = normalizeAllNotaryFieldValues(fieldValues, {
    contractBodyHtml: bodyHtml,
  });

  // 2. Perform placeholder injection into template body & header
  const mergedBodyHtml = mergePlaceholdersIntoHtml(
    config.bodyHtml,
    expandedValues
  );
  const mergedHeaderHtml = mergePlaceholdersIntoHtml(
    config.headerHtml,
    expandedValues
  );
  const mergedFooterHtml = mergePlaceholdersIntoHtml(
    config.footerHtml,
    expandedValues
  );

  // 3. Construct a clear notarial file name
  const cleanTitle = (contractTitle || 'عقد_توثيقي')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_');
  const fileName = `${cleanTitle}_${config.name.replace(/\s+/g, '_')}.docx`;

  // 4. Generate and trigger download
  await downloadNotaryDocx({
    title: fileName.replace(/\.docx$/, ''),
    bodyHtml: mergedBodyHtml,
    headerHtml: mergedHeaderHtml,
    footerHtml: mergedFooterHtml,
    pageNumberingEnabled: config.pageNumberingEnabled,
    fieldValues: expandedValues,
  });

  return {
    fileName,
    templateName: config.name,
  };
}


