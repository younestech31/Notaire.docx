import JSZip from 'jszip';
import { WordTemplateDefinition, STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';
import {
  generateNotaryDocxBlob,
  DocxExportOptions,
  downloadNotaryDocx,
  mergePlaceholdersIntoHtml,
} from './docx-engine';

/**
 * Parses the contract HTML into clauses based on:
 * 1. Explicit .clause-container[data-clause-id] elements
 * 2. Paragraphs starting with `#` (Layer 2 Notaire Local syntax: `# عنوان البند`)
 */
export function parseContractIntoClauses(
  contractHtml: string,
  fieldValues: Record<string, string> = {}
): { title: string; bodyHtml: string }[] {
  if (!contractHtml || typeof document === 'undefined') return [];
  const mergedHtml = mergePlaceholdersIntoHtml(contractHtml, fieldValues);
  const container = document.createElement('div');
  container.innerHTML = mergedHtml;

  const result: { title: string; bodyHtml: string }[] = [];

  // Check if there are .clause-container blocks or # headings
  const children = Array.from(container.children) as HTMLElement[];
  let currentTitle: string | null = null;
  let currentNodesHtml: string[] = [];

  const flushCurrent = () => {
    if (currentTitle && currentNodesHtml.length > 0) {
      result.push({
        title: currentTitle,
        bodyHtml: currentNodesHtml.join('\n'),
      });
    }
    currentTitle = null;
    currentNodesHtml = [];
  };

  for (const child of children) {
    if (
      child.classList.contains('clause-container') ||
      child.hasAttribute('data-clause-id')
    ) {
      flushCurrent();
      const rawTitle =
        child.getAttribute('data-clause-title') ||
        child.querySelector('h1,h2,h3,h4,strong,b')?.textContent ||
        'بند تعاقدي';
      const cleanTitle = rawTitle.replace(/^#\s*/, '').trim();
      const clone = child.cloneNode(true) as HTMLElement;
      const firstBlock = clone.firstElementChild as HTMLElement | null;
      if (firstBlock && (firstBlock.textContent || '').trim().startsWith('#')) {
        firstBlock.remove();
      }
      result.push({
        title: cleanTitle || 'بند تعاقدي',
        bodyHtml: clone.innerHTML.trim() || `<p dir="rtl">${cleanTitle}</p>`,
      });
      continue;
    }

    const text = (child.textContent || '').trim();
    if (text.startsWith('#')) {
      flushCurrent();
      currentTitle = text.replace(/^#+\s*/, '').trim() || 'بند تعاقدي';
    } else if (currentTitle !== null) {
      currentNodesHtml.push(child.outerHTML);
    }
  }
  flushCurrent();

  return result;
}


export const DEFAULT_WORD_TEMPLATES: WordTemplateDefinition[] = [
  {
    id: 'tpl_original',
    name: 'قالب العقد الأصلي الرسمي (Original)',
    type: 'original',
    description: 'قالب العقد التوثيقي الأصلي مع الهوامش الرسمية وبنود العقد وحلقة البنود `${block}`.',
    isDefault: true,
    updatedAt: '2026-01-01T00:00:00.000Z',
    headerHtml: `<div style="text-align:center; font-weight:bold; font-size:12pt;">${'${office_name}'} — ${'${office_addr}'}</div>`,
    footerHtml: `<div style="text-align:center; font-size:10pt;">الجمهورية الجزائرية الديمقراطية الشعبية</div>`,
    bodyHtml: `
      <p style="text-align:center; font-weight:bold; font-size:14pt;">${'${type_acte}'}</p>
      <p style="text-align:justify;">حرر بمكتب التوثيق لـ: ${'${office_name}'}، بتاريخ: ${'${date_acte}'} الموافق لـ: ${'${date_lettre}'}.</p>
      <p style="text-align:justify;"><b>الطرف الأول:</b> ${'${client_1}'}</p>
      <p style="text-align:justify;"><b>الطرف الثاني:</b> ${'${client_2}'}</p>
      <hr/>
      ${'${block}'}
      <p style="font-weight:bold; margin-top:10px;">${'${titre}'}</p>
      <p style="text-align:justify;">${'${clause}'}</p>
      ${'${/block}'}
      <hr/>
      <p style="text-align:justify; font-weight:bold;">الوضعية الجبائية والتسجيل:</p>
      <p style="text-align:justify;">1. أصل الملكية: ${'${fiscal_1}'}</p>
      <p style="text-align:justify;">2. تعيين الحصة: ${'${fiscal_2}'}</p>
      <p style="text-align:justify;">3. الأطراف: ${'${fiscal_3}'}</p>
      <p style="text-align:justify;">4. الثمن والتصريح: ${'${fiscal_4}'}</p>
      <p style="text-align:justify;">5. مراجع الإشهر: ${'${fiscal_5}'}</p>
    `.trim(),
  },
  {
    id: 'tpl_extract',
    name: 'قالب المستخرج (Extract)',
    type: 'extract',
    description: 'قالب مستخرج العقد الموجه لمفتشية التسجيل والطابع.',
    isDefault: true,
    updatedAt: '2026-01-01T00:00:00.000Z',
    headerHtml: `<div style="text-align:center; font-weight:bold;">مستخرج عقد توثيقي — ${'${office_name}'}</div>`,
    footerHtml: `<div style="text-align:center; font-size:10pt;">مستخرج رسمي للتسجيل</div>`,
    bodyHtml: `
      <p style="text-align:center; font-weight:bold; font-size:13pt;">مستخرج ${'${type_acte}'}</p>
      <p style="text-align:justify;">التاريخ: ${'${date_acte}'}</p>
      <p style="text-align:justify;"><b>المتصرف (الطرف الأول):</b> ${'${client_1}'}</p>
      <p style="text-align:justify;"><b>المتصرف إليه (الطرف الثاني):</b> ${'${client_2}'}</p>
      <p style="text-align:justify;"><b>تعيين العقار والحصة:</b> ${'${fiscal_2}'}</p>
      <p style="text-align:justify;"><b>الثمن والتقويم:</b> ${'${fiscal_4}'}</p>
    `.trim(),
  },
  {
    id: 'tpl_registration',
    name: 'قالب الشهر العقاري (Registration / Publication)',
    type: 'registration',
    description: 'قالب وثيقة إجراء الشهر العقاري الموجهة للمحافظة العقارية.',
    isDefault: true,
    updatedAt: '2026-01-01T00:00:00.000Z',
    headerHtml: `<div style="text-align:center; font-weight:bold;">المحافظة العقارية — وثيقة إجراء الشهر</div>`,
    footerHtml: `<div style="text-align:center; font-size:10pt;">نسخة موثقة للشهر</div>`,
    bodyHtml: `
      <p style="text-align:center; font-weight:bold; font-size:13pt;">إجراء شهر ${'${type_acte}'}</p>
      <p style="text-align:justify;">بموجب العقد المحرر في: ${'${date_acte}'} بمكتب ${'${office_name}'}.</p>
      <p style="text-align:justify;"><b>الأطراف:</b><br/>- البائع: ${'${client_1}'}<br/>- المشتري: ${'${client_2}'}</p>
      <p style="text-align:justify;"><b>أصل الملكية والسند:</b> ${'${fiscal_1}'}</p>
      <p style="text-align:justify;"><b>الوصف التقسيمي الدقيق:</b> ${'${fiscal_2}'}</p>
      <p style="text-align:justify;"><b>المبلغ المالي:</b> ${'${fiscal_4}'}</p>
    `.trim(),
  },
  {
    id: 'tpl_fiscal',
    name: 'قالب الوضعية الجبائية (Fiscal Position)',
    type: 'fiscal',
    description: 'قالب مخصص لعرض وتصدير الجوانب الجبائية الخمسة للعقد.',
    isDefault: true,
    updatedAt: '2026-01-01T00:00:00.000Z',
    headerHtml: `<div style="text-align:center; font-weight:bold;">الوضعية الجبائية للعقد</div>`,
    footerHtml: `<div style="text-align:center;">مكتب الموثق ${'${office_name}'}</div>`,
    bodyHtml: `
      <p style="text-align:center; font-weight:bold; font-size:13pt;">بطاقة الوضعية الجبائية — ${'${type_acte}'}</p>
      <p style="text-align:justify;"><b>1. أصل الملكية:</b> ${'${fiscal_1}'}</p>
      <p style="text-align:justify;"><b>2. تعيين العقار والحصة:</b> ${'${fiscal_2}'}</p>
      <p style="text-align:justify;"><b>3. الأطراف والقدرة القانونية:</b> ${'${fiscal_3}'}</p>
      <p style="text-align:justify;"><b>4. الثمن وشروط الدفع:</b> ${'${fiscal_4}'}</p>
      <p style="text-align:justify;"><b>5. مراجع الشهر والوضع العقاري:</b> ${'${fiscal_5}'}</p>
    `.trim(),
  },
];

/**
 * Pours contract content, clauses, and variables into a WordTemplateDefinition
 * and produces final HTML ready for Word export (.docx).
 */
export function renderTemplateWithContractData(
  template: WordTemplateDefinition,
  contractData: {
    officeName?: string;
    officeAddr?: string;
    typeActe?: string;
    client1?: string;
    client2?: string;
    dateActe?: string;
    dateLettre?: string;
    fiscal1?: string;
    fiscal2?: string;
    fiscal3?: string;
    fiscal4?: string;
    fiscal5?: string;
    clauses?: { title: string; contentHtml: string }[];
    fieldValues?: Record<string, string>;
  }
): { bodyHtml: string; headerHtml: string; footerHtml: string } {
  const officeName = contractData.officeName || contractData.fieldValues?.['اسم_المكتب'] || 'مكتب التوثيق';
  const officeAddr = contractData.officeAddr || contractData.fieldValues?.['عنوان_المكتب'] || 'الجزائر العاصمة';
  const typeActe = contractData.typeActe || contractData.fieldValues?.['نوع_العقد'] || 'عقد توثيقي';
  const client1 = contractData.client1 || contractData.fieldValues?.['الطرف_الأول_الاسم'] || 'الطرف الأول';
  const client2 = contractData.client2 || contractData.fieldValues?.['الطرف_الثاني_الاسم'] || 'الطرف الثاني';
  const dateActe = contractData.dateActe || contractData.fieldValues?.['تاريخ_العقد'] || new Date().toLocaleDateString('ar-DZ');
  const dateLettre = contractData.dateLettre || contractData.fieldValues?.['تاريخ_العقد_بالحروف'] || 'اليوم';

  const fiscal1 = contractData.fiscal1 || contractData.fieldValues?.['أصل_الملكية'] || 'مشهار بالمحافظة العقارية وفق السند القانوني.';
  const fiscal2 = contractData.fiscal2 || contractData.fieldValues?.['تعيين_الحصة_الكامل'] || contractData.fieldValues?.['طبيعة_الحصة'] || 'الحصة العقارية المعينة بدقة.';
  const fiscal3 = contractData.fiscal3 || contractData.fieldValues?.['الأطراف'] || 'الأطراف المذكور أعلاه كاملو الأهلية والقانون.';
  const fiscal4 = contractData.fiscal4 || contractData.fieldValues?.['الثمن_بالأحرف'] || contractData.fieldValues?.['الثمن_بالأرقام'] || 'القيمة المتفق عليها رضاءً وقانوناً.';
  const fiscal5 = contractData.fiscal5 || contractData.fieldValues?.['مراجع_الوصف_التقسيمي'] || 'الشهر العقاري الجاري.';

  const replacements: Record<string, string> = {
    '${office_name}': officeName,
    '${office_addr}': officeAddr,
    '${type_acte}': typeActe,
    '${client_1}': client1,
    '${client_2}': client2,
    '${date_acte}': dateActe,
    '${date_lettre}': dateLettre,
    '${fiscal_1}': fiscal1,
    '${fiscal_2}': fiscal2,
    '${fiscal_3}': fiscal3,
    '${fiscal_4}': fiscal4,
    '${fiscal_5}': fiscal5,
  };

  let headerHtml = template.headerHtml || '';
  let footerHtml = template.footerHtml || '';
  let bodyHtml = template.bodyHtml || '';

  // Apply simple variable replacements
  for (const [key, val] of Object.entries(replacements)) {
    const regex = new RegExp(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    headerHtml = headerHtml.replace(regex, val);
    footerHtml = footerHtml.replace(regex, val);
    bodyHtml = bodyHtml.replace(regex, val);
  }

  // Handle ${block} ... ${/block} clause loops
  const blockRegex = /\$\{\s*block\s*\}([\s\S]*?)\$\{\s*\/block\s*\}/g;
  bodyHtml = bodyHtml.replace(blockRegex, (_match, blockTemplate) => {
    const clausesList = contractData.clauses || [];
    if (clausesList.length === 0) {
      // Fallback if no structured clauses, use default sample clauses
      return `
        <p style="font-weight:bold;">البند الأول: التعيين والتعرف</p>
        <p style="text-align:justify;">${fiscal2}</p>
        <p style="font-weight:bold; margin-top:8px;">البند الثاني: الثمن والشروط المالية</p>
        <p style="text-align:justify;">${fiscal4}</p>
      `;
    }

    return clausesList
      .map((cl) => {
        let itemHtml = blockTemplate;
        const mergedClauseHtml = mergePlaceholdersIntoHtml(
          cl.contentHtml || '',
          contractData.fieldValues || {}
        );
        itemHtml = itemHtml.replace(/\$\{\s*titre\s*\}/g, cl.title || 'بند قانوني');
        itemHtml = itemHtml.replace(/\$\{\s*clause\s*\}/g, mergedClauseHtml);
        return itemHtml;
      })
      .join('\n');
  });

  return { bodyHtml, headerHtml, footerHtml };
}

export async function exportTemplateAsDocx(
  template: WordTemplateDefinition,
  contractData: Parameters<typeof renderTemplateWithContractData>[1]
): Promise<void> {
  const rendered = renderTemplateWithContractData(template, contractData);
  const options: DocxExportOptions = {
    title: template.name || 'وثيقة_موثقة',
    bodyHtml: rendered.bodyHtml,
    headerHtml: rendered.headerHtml,
    footerHtml: rendered.footerHtml,
    pageNumberingEnabled: true,
    fieldValues: contractData.fieldValues,
  };
  await downloadNotaryDocx(options);
}
