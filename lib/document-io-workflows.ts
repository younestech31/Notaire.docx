import {
  CustomTemplate,
  DerivedDocTemplate,
  DownloadArchiveItem,
  STRICT_FONT_FAMILY,
  STRICT_FONT_SIZE_PT,
  SubdivisionEstate,
  SubdivisionLot,
} from './types';
import {
  exportCustomTemplatesJson,
  exportFullBackupBundle,
  importCustomTemplatesJson,
  loadCustomTemplates,
  loadDownloadArchive,
  saveCustomTemplate,
  saveDownloadArchiveItem,
  sealContractOriginal,
} from './storage';
import {
  downloadNotaryDocx,
  importNotaryDocxFile,
  mergePlaceholdersIntoHtml,
} from './docx-engine';
import { escapeHtml, extractPlaceholdersFromHtml } from './editor-utils';

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

export async function exportAndSealContractToWord(params: {
  docId: string;
  docTitle: string;
  bodyHtml: string;
  headerHtml: string;
  footerHtml: string;
  showHeaderFooter: boolean;
  pageNumberingEnabled: boolean;
  fieldValues: Record<string, string>;
  year: number;
  indexNumber: string;
}): Promise<{
  relativePath: string;
  sha256: string;
  downloads: DownloadArchiveItem[];
}> {
  const titleToUse = params.docTitle || 'عقد_توثيقي';
  const docxBytes = await downloadNotaryDocx({
    title: titleToUse,
    bodyHtml: params.bodyHtml,
    headerHtml: params.showHeaderFooter ? params.headerHtml : '',
    footerHtml: params.showHeaderFooter ? params.footerHtml : '',
    pageNumberingEnabled: params.pageNumberingEnabled,
    fieldValues: params.fieldValues,
  });

  const sealed = await sealContractOriginal({
    documentId: params.docId,
    year: params.year,
    indexNumber: params.indexNumber,
    docxBytes,
  });

  const archiveItem: DownloadArchiveItem = {
    id: `dl_${Date.now()}`,
    documentId: params.docId,
    documentTitle: titleToUse,
    docTypeCode: 'acte',
    docTypeLabel: `الأصل المعتمد (${sealed.relativePath})`,
    fileName: `${titleToUse}.docx`,
    bodyHtml: params.bodyHtml,
    headerHtml: params.showHeaderFooter ? params.headerHtml : '',
    footerHtml: params.showHeaderFooter ? params.footerHtml : '',
    fieldValues: { ...params.fieldValues },
    createdAt: new Date().toISOString(),
  };
  await saveDownloadArchiveItem(archiveItem);
  const downloads = await loadDownloadArchive();

  return {
    relativePath: sealed.relativePath,
    sha256: sealed.sha256,
    downloads,
  };
}

export async function generateAndDownloadDerivedDoc(params: {
  docId: string;
  docTitle: string;
  tpl: DerivedDocTemplate;
  fieldValues: Record<string, string>;
}): Promise<{
  missingVarsCount: number;
  downloads: DownloadArchiveItem[];
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
    pageNumberingEnabled: true,
    fieldValues,
  });

  const archiveItem: DownloadArchiveItem = {
    id: `dl_${Date.now()}`,
    documentId: docId,
    documentTitle: docTitle || 'عقد توثيقي',
    docTypeCode: tpl.code,
    docTypeLabel: tpl.name,
    fileName: `${fileTitle}.docx`,
    bodyHtml: mergePlaceholdersIntoHtml(tpl.bodyHtml, fieldValues),
    headerHtml: mergePlaceholdersIntoHtml(tpl.headerHtml, fieldValues),
    footerHtml: mergePlaceholdersIntoHtml(tpl.footerHtml, fieldValues),
    fieldValues: { ...fieldValues },
    createdAt: new Date().toISOString(),
  };
  await saveDownloadArchiveItem(archiveItem);
  const downloads = await loadDownloadArchive();

  return {
    missingVarsCount: missingVars.length,
    downloads,
  };
}

export async function importDocxFilesAsTemplates(files: FileList): Promise<{
  importedCount: number;
  templates: CustomTemplate[];
}> {
  let importedCount = 0;
  for (const file of Array.from(files)) {
    try {
      const res = await importNotaryDocxFile(file);
      const newTpl: CustomTemplate = {
        id: `tpl_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        name: res.title,
        category: 'نماذج وورد مستوردة',
        bodyHtml: res.bodyHtml,
        headerHtml: res.headerHtml,
        footerHtml: res.footerHtml,
        pageNumberingEnabled: true,
        extractedPlaceholders: res.extractedPlaceholders,
        sourceFileName: file.name,
        updatedAt: new Date().toISOString(),
      };
      await saveCustomTemplate(newTpl);
      importedCount++;
    } catch {
      // Skip invalid file
    }
  }
  const templates = await loadCustomTemplates();
  return { importedCount, templates };
}

export async function downloadTemplatesJsonFile(): Promise<void> {
  const jsonStr = await exportCustomTemplatesJson();
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
  templates: CustomTemplate[];
}> {
  const text = await file.text();
  const parsed = JSON.parse(text);
  if (!Array.isArray(parsed)) {
    throw new Error('ملف JSON غير صالح كقائمة قوالب');
  }
  const count = await importCustomTemplatesJson(parsed);
  const templates = await loadCustomTemplates();
  return { count, templates };
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
