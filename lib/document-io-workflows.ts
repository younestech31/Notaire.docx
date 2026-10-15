import {
  DocumentRevision,
  OfficeTemplate,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  SubdivisionEstate,
  SubdivisionLot,
} from './types';
import {
  exportFullBackupBundle,
  exportOfficeTemplatesJson,
  importOfficeTemplatesJson,
  loadDocumentRevisions,
  loadOfficeTemplates,
  saveDocumentRevision,
  saveOfficeTemplate,
} from './storage';
import {
  downloadNotaryDocx,
  importNotaryDocxFile,
} from './docx-engine';
import { escapeHtml, extractPlaceholdersFromHtml } from './editor-utils';
import { purifyNotarialBodyHtml } from './notarial-body-purifier';
import {
  exportContractWithStandardTemplate,
  StandardNotaryTemplateId,
} from './notary-word-templates';

export function buildSubdivisionLotClauseHtml(
  estate: SubdivisionEstate,
  lot: SubdivisionLot
): string {
  return `<p dir="rtl" style="margin:0;line-height:1;font-family:${STRICT_FONT_FAMILY};font-size:${STRICT_FONT_SIZE_PT}pt;text-align:justify;"><span style="font-weight:bold;">التعيين العقاري (الحصة رقم ${escapeHtml(
    lot.lotNumber
  )}): </span>${escapeHtml(lot.fullDescription)}${
    estate.subdivisionDeedRef
      ? ` وذلك بموجب الجدول الوصفي للتقسيم المحرر بتاريخ ومراجع: ${escapeHtml(
          estate.subdivisionDeedRef
        )}.`
      : ''
  }</p>`;
}

export function buildFullSubdivisionTableHtml(estate: SubdivisionEstate): string {
  let rowsHtml = `<tr>
    <td style="width:14%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">رقم الحصة</p></td>
    <td style="width:20%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">الطبيعة</p></td>
    <td style="width:20%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">الطابق والعمارة</p></td>
    <td style="width:18%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">المساحة (م²)</p></td>
    <td style="width:28%;border:1px solid #000000;background-color:#F1F5F9;padding:2px 4px;font-weight:bold;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;font-weight:bold;">الحصة في الأجزاء المشتركة</p></td>
  </tr>`;

  for (const lot of estate.lots) {
    rowsHtml += `<tr>
      <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
        lot.lotNumber
      )}</p></td>
      <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
        lot.nature
      )}</p></td>
      <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
        `${lot.floor} ${lot.building}`.trim()
      )}</p></td>
      <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
        lot.area
      )}</p></td>
      <td style="border:1px solid #000000;padding:2px 4px;"><p dir="rtl" style="margin:0;line-height:1;text-align:center;">${escapeHtml(
        lot.commonShares
      )}</p></td>
    </tr>`;
  }

  return `<p dir="rtl" style="margin:0;line-height:1;font-weight:bold;">جدول الوصف التقسيمي للعقار: ${escapeHtml(
    estate.estateName
  )}</p><table dir="rtl" style="width:100%;max-width:120mm;border-collapse:collapse;table-layout:fixed;margin:0;"><tbody>${rowsHtml}</tbody></table><p dir="rtl" style="margin:0;line-height:1;"><br></p>`;
}

/**
 * Exports the contract as a .docx file and records a lightweight 'export' event
 * in DocumentRevision without sealing the official original (sealing is strictly
 * reserved for officiallyIndexAndSealContract).
 */
export async function exportEditorBodyToDocx(params: {
  docId: string;
  docTitle: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  showHeaderFooter: boolean;
  pageNumberingEnabled: boolean;
  fieldValues: Record<string, string>;
  authorName?: string;
}): Promise<{
  fileName: string;
  revisions: DocumentRevision[];
}> {
  const titleToUse = params.docTitle || 'عقد_توثيقي';
  const purifiedBody = purifyNotarialBodyHtml(params.bodyHtml);
  
  await downloadNotaryDocx({
    title: titleToUse,
    bodyHtml: purifiedBody,
    headerHtml: params.showHeaderFooter ? params.headerHtml : '',
    footerHtml: params.showHeaderFooter ? params.footerHtml : '',
    pageNumberingEnabled: params.pageNumberingEnabled,
    fieldValues: params.fieldValues,
    omitRepeatingHeader: !params.showHeaderFooter,
  });

  const fileName = `${titleToUse}.docx`;
  const exportRevision: DocumentRevision = {
    id: `rev_exp_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    documentId: params.docId,
    documentTitle: titleToUse,
    author: params.authorName || 'الموثق الرئيسي',
    fieldValues: { ...params.fieldValues },
    createdAt: new Date().toISOString(),
    eventType: 'export',
    fileName,
    docTypeCode: 'acte',
    summary: `تصدير متن العقد (تطهير): ${fileName}`,
  };
  await saveDocumentRevision(exportRevision);
  const revisions = await loadDocumentRevisions();

  return {
    fileName,
    revisions,
  };
}

export async function exportWithStandardTemplate(params: {
  templateId: string;
  contractTitle: string;
  bodyHtml: string;
  fieldValues: Record<string, string>;
}): Promise<{ fileName: string; templateName: string }> {
  return await exportContractWithStandardTemplate({
    templateId: params.templateId as StandardNotaryTemplateId,
    contractTitle: params.contractTitle,
    bodyHtml: params.bodyHtml,
    fieldValues: params.fieldValues,
  });
}

export async function generateAndDownloadDerivedDoc(params: {
  docId: string;
  docTitle: string;
  tpl: OfficeTemplate;
  fieldValues: Record<string, string>;
  authorName?: string;
}): Promise<{
  missingVarsCount: number;
  revisions: DocumentRevision[];
}> {
  const { docId, docTitle, tpl, fieldValues } = params;
  const requiredVars = extractPlaceholdersFromHtml(
    tpl.headerHtml,
    tpl.bodyHtml,
    tpl.footerHtml
  );
  const missingVars = requiredVars.filter(
    (k) => !fieldValues[k] || !fieldValues[k].trim()
  );

  const fileTitle = `${docTitle || 'عقد'} - ${tpl.name}`;
  await downloadNotaryDocx({
    title: fileTitle,
    bodyHtml: tpl.bodyHtml,
    headerHtml: tpl.headerHtml,
    footerHtml: tpl.footerHtml,
    pageNumberingEnabled: tpl.pageNumberingEnabled !== false,
    fieldValues,
  });

  const docTypeCode = tpl.docTypeCode || tpl.code || 'extrait';
  const fileName = `${fileTitle}.docx`;

  const derivedRevision: DocumentRevision = {
    id: `rev_drv_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    documentId: docId,
    documentTitle: docTitle || 'عقد توثيقي',
    author: params.authorName || 'الموثق الرئيسي',
    fieldValues: { ...fieldValues },
    createdAt: new Date().toISOString(),
    eventType: 'derived',
    fileName,
    docTypeCode,
    summary: `توليد وتصدير وثيقة مشتقة: ${tpl.name} (${fileName})`,
  };
  await saveDocumentRevision(derivedRevision);
  const revisions = await loadDocumentRevisions();

  return {
    missingVarsCount: missingVars.length,
    revisions,
  };
}

export async function importDocxFilesAsTemplates(files: FileList): Promise<{
  importedCount: number;
  officeTemplates: OfficeTemplate[];
  templates: OfficeTemplate[];
}> {
  let importedCount = 0;
  for (const file of Array.from(files)) {
    try {
      const res = await importNotaryDocxFile(file);
      const newTpl: OfficeTemplate = {
        id: `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        name: res.title,
        kind: 'contract',
        category: 'نماذج وورد مستوردة',
        docTypeCode: 'acte',
        bodyHtml: res.bodyHtml,
        headerHtml: res.headerHtml,
        footerHtml: res.footerHtml,
        pageNumberingEnabled: true,
        extractedPlaceholders: res.extractedPlaceholders,
        sourceFileName: file.name,
        updatedAt: new Date().toISOString(),
      };
      await saveOfficeTemplate(newTpl);
      importedCount++;
    } catch {
      // Skip invalid file
    }
  }
  const officeTemplates = await loadOfficeTemplates();
  return {
    importedCount,
    officeTemplates,
    templates: officeTemplates,
  };
}

export async function downloadTemplatesJsonFile(): Promise<void> {
  const jsonStr = await exportOfficeTemplatesJson();
  const blob = new Blob([jsonStr], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `قوالب_مكتب_التوثيق_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function importTemplatesJsonFile(file: File): Promise<{
  count: number;
  officeTemplates: OfficeTemplate[];
  templates: OfficeTemplate[];
}> {
  const text = await file.text();
  const parsed = JSON.parse(text);
  if (!Array.isArray(parsed)) {
    throw new Error('ملف JSON غير صالح كقائمة قوالب');
  }
  const count = await importOfficeTemplatesJson(parsed);
  const officeTemplates = await loadOfficeTemplates();
  return {
    count,
    officeTemplates,
    templates: officeTemplates,
  };
}

export async function downloadOfficeBackupBundleFile(): Promise<string> {
  const bundle = await exportFullBackupBundle();
  const blob = new Blob([JSON.stringify(bundle, null, 2)], {
    type: 'application/json',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `نسخة_احتياطية_الموثق_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  return bundle.exportedAt;
}
