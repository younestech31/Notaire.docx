import {
  ContractLifecycleStatus,
  DocumentRevision,
  PublicationTrack,
  RegistrationBatch,
  SavedDocument,
} from './types';
import { generateNotaryDocxBlob } from './docx-engine';
import {
  loadSavedDocuments,
  saveDocumentRecord,
  saveDocumentRevision,
  sealContractOriginal,
} from './storage';

/**
 * Normalizes legacy or mixed contract lifecycle status values into canonical states.
 */
export function normalizeContractStatus(
  status?: ContractLifecycleStatus | string,
  hasIndexNumber?: boolean
): ContractLifecycleStatus {
  if (!status) {
    return hasIndexNumber ? 'signed_indexed' : 'draft';
  }
  if (status === 'signed' && hasIndexNumber) {
    return 'signed_indexed';
  }
  return status as ContractLifecycleStatus;
}

/**
 * Checks whether a contract is officially indexed and sealed (so its main body should be locked
 * against accidental raw text edits while allowing post-signature metadata updates).
 */
export function isContractOfficiallyIndexed(doc?: SavedDocument | null): boolean {
  if (!doc) return false;
  const norm = normalizeContractStatus(doc.status, Boolean(doc.indexNumber));
  return (
    norm === 'signed_indexed' ||
    norm === 'deposited_for_registration' ||
    norm === 'registration_rejected' ||
    norm === 'registered' ||
    norm === 'deposited_for_publication' ||
    norm === 'publication_rejected' ||
    norm === 'published' ||
    Boolean(doc.sealedOriginalSha256)
  );
}

/**
 * Parses numeric index sequence and year from strings like "142/2026", "2026/142", or "142".
 */
export function parseIndexNumberAndYear(
  rawIndex: string,
  fallbackYear: number = new Date().getFullYear()
): { sequence: number; year: number; formatted: string } {
  const cleaned = String(rawIndex || '').trim();
  const slashParts = cleaned.split(/[\/\-]/).map((s) => s.trim());

  if (slashParts.length === 2) {
    const p1 = parseInt(slashParts[0], 10);
    const p2 = parseInt(slashParts[1], 10);
    if (!isNaN(p1) && !isNaN(p2)) {
      if (p1 >= 1990 && p1 <= 2100) {
        return { sequence: p2, year: p1, formatted: `${p2}/${p1}` };
      }
      if (p2 >= 1990 && p2 <= 2100) {
        return { sequence: p1, year: p2, formatted: `${p1}/${p2}` };
      }
    }
  }

  const numMatch = cleaned.match(/\d+/);
  const seq = numMatch ? parseInt(numMatch[0], 10) : 1;
  return {
    sequence: Math.max(1, seq),
    year: fallbackYear,
    formatted: `${Math.max(1, seq)}/${fallbackYear}`,
  };
}

/**
 * Computes the next sequential annual index number (`آخر فهرس + 1`) for the given year
 * across all saved documents in the office (Decree 08-244).
 */
export function suggestNextAnnualIndexNumber(
  documents: SavedDocument[],
  year: number = new Date().getFullYear()
): string {
  let maxSeq = 0;
  for (const doc of documents) {
    const rawIdx = doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'];
    if (!rawIdx) continue;
    const parsed = parseIndexNumberAndYear(rawIdx, doc.year || year);
    if (parsed.year === year && parsed.sequence > maxSeq) {
      maxSeq = parsed.sequence;
    }
  }
  return `${maxSeq + 1}/${year}`;
}

/**
 * Suggests the next deposit number (`رقم الإيداع`) for the given year.
 */
export function suggestNextDepositNumber(
  documents: SavedDocument[],
  year: number = new Date().getFullYear()
): string {
  let maxSeq = 0;
  for (const doc of documents) {
    const rawDep = doc.depositNumber || doc.fieldValues?.['رقم_الإيداع'];
    if (!rawDep) continue;
    const parsed = parseIndexNumberAndYear(rawDep, doc.year || year);
    if (parsed.year === year && parsed.sequence > maxSeq) {
      maxSeq = parsed.sequence;
    }
  }
  return `${maxSeq + 1}/${year}`;
}

/**
 * Validates that `indexNumber` is not already assigned to another contract in the same year.
 */
export function findConflictingIndexedDocument(
  documents: SavedDocument[],
  targetDocId: string,
  indexNumber: string,
  year: number = new Date().getFullYear()
): SavedDocument | undefined {
  const targetParsed = parseIndexNumberAndYear(indexNumber, year);
  return documents.find((doc) => {
    if (doc.id === targetDocId) return false;
    const docRawIdx = doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'];
    if (!docRawIdx) return false;
    const docParsed = parseIndexNumberAndYear(docRawIdx, doc.year || year);
    return (
      docParsed.year === targetParsed.year &&
      docParsed.sequence === targetParsed.sequence
    );
  });
}

/**
 * Executes the complete official indexing and sealing workflow (Decree 08-244):
 * 1. Checks for duplicate index in the same year
 * 2. Injects `[رقم_الفهرس]`, `[رقم_الإيداع]`, and `[تاريخ_العقد]` into `fieldValues`
 * 3. Generates final `.docx` and seals original via `sealContractOriginal` (SHA-256)
 * 4. Updates document with `indexNumber`, `depositNumber`, `indexedAt`, `status: 'signed_indexed'`, `sealedOriginalPath`, `sealedOriginalSha256`
 * 5. Records an audit revision in `DocumentRevision`
 */
export async function officiallyIndexAndSealContract(params: {
  document: SavedDocument;
  allDocuments: SavedDocument[];
  indexNumberInput: string;
  depositNumberInput?: string;
  contractDateInput?: string;
  authorName?: string;
}): Promise<{
  updatedDocument: SavedDocument;
  relativePath: string;
  sha256: string;
}> {
  const plainBody = (params.document.bodyHtml || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!plainBody) {
    throw new Error(
      'لا يمكن اعتماد أو فهرسة عقد فارغ بلا متن نصي محفوظ. يرجى كتابة متن العقد وحفظه في المكتب قبل ختم الأصل.'
    );
  }

  const currentYear = params.document.year || new Date().getFullYear();
  const parsedIndex = parseIndexNumberAndYear(
    params.indexNumberInput || suggestNextAnnualIndexNumber(params.allDocuments, currentYear),
    currentYear
  );

  const conflict = findConflictingIndexedDocument(
    params.allDocuments,
    params.document.id,
    parsedIndex.formatted,
    parsedIndex.year
  );
  if (conflict) {
    throw new Error(
      `رقم الفهرس (${parsedIndex.formatted}) مستخدم مسبقاً في نفس السنة للعقد "${conflict.title}". الفهرس الرسمي متسلسل ولا يقبل التكرار.`
    );
  }

  const nowIso = new Date().toISOString();
  const formattedDate =
    params.contractDateInput ||
    params.document.fieldValues?.['تاريخ_العقد'] ||
    nowIso.slice(0, 10);
  const depositNo =
    (params.depositNumberInput || '').trim() ||
    params.document.depositNumber ||
    suggestNextDepositNumber(params.allDocuments, parsedIndex.year);

  const updatedFieldValues: Record<string, string> = {
    ...(params.document.fieldValues || {}),
    رقم_الفهرس: parsedIndex.formatted,
    رقم_الإيداع: depositNo,
    تاريخ_العقد: formattedDate,
  };

  // Generate final .docx bytes and seal with SHA-256
  const docxBlob = await generateNotaryDocxBlob({
    title: params.document.title,
    bodyHtml: params.document.bodyHtml,
    headerHtml: params.document.showHeaderFooter ? params.document.headerHtml : '',
    footerHtml: params.document.showHeaderFooter ? params.document.footerHtml : '',
    pageNumberingEnabled: params.document.pageNumberingEnabled !== false,
    fieldValues: updatedFieldValues,
  });
  const docxBytes = new Uint8Array(await docxBlob.arrayBuffer());

  const sealResult = await sealContractOriginal({
    documentId: params.document.id,
    year: parsedIndex.year,
    indexNumber: String(parsedIndex.sequence),
    docxBytes,
  });

  // Clean any old [فهرس: ...] suffix from title before setting clean title
  const baseTitle = params.document.title.replace(/\s*\[فهرس:[^\]]+\]\s*$/g, '').trim();

  const updatedDocument: SavedDocument = {
    ...params.document,
    title: baseTitle,
    depositNumber: depositNo,
    indexNumber: parsedIndex.formatted,
    contractNumber: parsedIndex.formatted,
    year: parsedIndex.year,
    indexedAt: nowIso,
    status: 'signed_indexed',
    sealedOriginalPath: sealResult.relativePath,
    sealedOriginalSha256: sealResult.sha256,
    fieldValues: updatedFieldValues,
    updatedAt: nowIso,
  };

  await saveDocumentRecord(updatedDocument);

  const revision: DocumentRevision = {
    id: `rev_index_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    documentId: updatedDocument.id,
    documentTitle: updatedDocument.title,
    author: params.authorName || updatedDocument.clerkName || 'الموثق الرئيسي',
    bodyHtml: updatedDocument.bodyHtml,
    fieldValues: updatedFieldValues,
    createdAt: nowIso,
    summary: `فهرسة رسمية برقم (${parsedIndex.formatted}) وإيداع (${depositNo}) وختم الأصل ببصمة SHA-256 (${sealResult.sha256.slice(0, 12)})`,
  };
  await saveDocumentRevision(revision);

  return {
    updatedDocument,
    relativePath: sealResult.relativePath,
    sha256: sealResult.sha256,
  };
}

/**
 * Formats the official registration reference string `[مراجع_التسجيل]` from a RegistrationBatch.
 */
export function formatRegistrationReferenceText(batch: RegistrationBatch): string {
  const datePart = batch.registrationDate || batch.statusDate;
  if (batch.status === 'rejected') {
    return [
      `رُفض للتسوية/التصحيح بـ${batch.inspectorateName || 'مفتشية التسجيل والطابع'}`,
      datePart ? `بتاريخ ${datePart}` : '',
      batch.notes ? `(${batch.notes})` : '',
    ]
      .filter(Boolean)
      .join(' ');
  }
  const receiptPart = batch.receiptNumber
    ? `حسب الوصل رقم ${batch.receiptNumber}`
    : 'مودع للتسجيل';
  const feesPart = batch.totalFees
    ? `بحقوق قدرها ${batch.totalFees.toLocaleString('ar-DZ')} دج`
    : '';
  return [
    `سُجّل بـ${batch.inspectorateName || 'مفتشية التسجيل والطابع'}`,
    datePart ? `بتاريخ ${datePart}` : '',
    receiptPart,
    feesPart,
  ]
    .filter(Boolean)
    .join(' ');
}

/**
 * Synchronizes a saved or updated `RegistrationBatch` back to all linked contracts:
 * - Matches contracts by `contractIds` and by `indexNumbers`
 * - Sets `registrationBatchId`, `inspectorateName`, and updates `status` (`deposited_for_registration`, `registered`, or `registration_rejected`)
 * - Automatically populates `fieldValues['مراجع_التسجيل']` (and updates derived documents if present)
 */
export async function syncRegistrationBatchToContracts(
  batch: RegistrationBatch,
  authorName: string = 'الموثق الرئيسي'
): Promise<SavedDocument[]> {
  const allDocs = await loadSavedDocuments();
  const normalizedBatchIndices = new Set(
    (batch.indexNumbers || []).map((idx) =>
      parseIndexNumberAndYear(idx, batch.year).formatted
    )
  );
  const idSet = new Set(batch.contractIds || []);

  const targetStatus: ContractLifecycleStatus =
    batch.status === 'registered'
      ? 'registered'
      : batch.status === 'rejected'
      ? 'registration_rejected'
      : 'deposited_for_registration';
  const regRefText = formatRegistrationReferenceText(batch);
  const nowIso = new Date().toISOString();

  const updatedList: SavedDocument[] = [];

  for (const doc of allDocs) {
    const docIdxRaw = doc.indexNumber || doc.fieldValues?.['رقم_الفهرس'] || '';
    const docIdxFormatted = docIdxRaw
      ? parseIndexNumberAndYear(docIdxRaw, doc.year || batch.year).formatted
      : '';

    const isMatched =
      idSet.has(doc.id) ||
      (docIdxFormatted && normalizedBatchIndices.has(docIdxFormatted));

    if (!isMatched) continue;

    const nextFieldValues: Record<string, string> = {
      ...(doc.fieldValues || {}),
      مراجع_التسجيل: regRefText,
      مفتشية_التسجيل: batch.inspectorateName || 'مفتشية التسجيل والطابع',
    };
    if (batch.receiptNumber) {
      nextFieldValues['وصل_التسجيل'] = batch.receiptNumber;
    }
    if (batch.registrationDate || batch.statusDate) {
      nextFieldValues['تاريخ_التسجيل'] = batch.registrationDate || batch.statusDate;
    }

    // Keep existing published status if already published
    const nextStatus: ContractLifecycleStatus =
      doc.status === 'published' || doc.status === 'deposited_for_publication'
        ? doc.status
        : targetStatus;

    const updatedDerived = (doc.derivedDocuments || []).map((dd) => ({
      ...dd,
      fieldValues: {
        ...(dd.fieldValues || {}),
        مراجع_التسجيل: regRefText,
      },
    }));

    const updatedDoc: SavedDocument = {
      ...doc,
      registrationBatchId: batch.id,
      inspectorateName: batch.inspectorateName,
      status: nextStatus,
      fieldValues: nextFieldValues,
      derivedDocuments: updatedDerived,
      updatedAt: nowIso,
    };

    await saveDocumentRecord(updatedDoc);
    updatedList.push(updatedDoc);

    await saveDocumentRevision({
      id: `rev_reg_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      documentId: updatedDoc.id,
      documentTitle: updatedDoc.title,
      author: authorName,
      bodyHtml: updatedDoc.bodyHtml,
      fieldValues: nextFieldValues,
      createdAt: nowIso,
      summary: `تحديث دفعة التسجيل (${
        batch.status === 'registered'
          ? 'مسجل'
          : batch.status === 'rejected'
          ? 'مرفوض للتسوية'
          : 'مودع للتسجيل'
      }): ${regRefText}`,
    });
  }

  return updatedList;
}

/**
 * Formats the official land publication reference string `[مراجع_الشهر]` from a PublicationTrack.
 */
export function formatPublicationReferenceText(track: PublicationTrack): string {
  const officePart = `${track.conservationOffice || 'المحافظة العقارية'}${
    track.wilaya ? ` (${track.wilaya})` : ''
  }`;
  if (track.status === 'rejected') {
    return [
      `رفض/تحفظ بالشهر العقاري لدى ${officePart}`,
      track.depositDate ? `بتاريخ ${track.depositDate}` : '',
      track.notes ? `(${track.notes})` : '',
    ]
      .filter(Boolean)
      .join(' ');
  }
  if (track.status === 'published' || (track.volume && track.number)) {
    const pubDate = track.publicationDate || track.depositDate || '';
    return [
      `شُهر بـ${officePart}`,
      pubDate ? `بتاريخ ${pubDate}` : '',
      track.volume ? `مجلد ${track.volume}` : '',
      track.number ? `رقم ${track.number}` : '',
    ]
      .filter(Boolean)
      .join(' ');
  }
  return [
    `أودع للشهر بـ${officePart}`,
    track.depositDate ? `بتاريخ ${track.depositDate}` : '',
  ]
    .filter(Boolean)
    .join(' ');
}

/**
 * Synchronizes a saved or updated `PublicationTrack` back to its linked contract:
 * - Updates contract status (`deposited_for_publication`, `published`, or `publication_rejected`)
 * - Populates `fieldValues['مراجع_الشهر']`, `المحافظة_العقارية`, `مجلد_الشهر`, `رقم_الشهر`
 */
export async function syncPublicationTrackToContract(
  track: PublicationTrack,
  authorName: string = 'الموثق الرئيسي'
): Promise<SavedDocument | null> {
  const allDocs = await loadSavedDocuments();
  const trackIdxFormatted = track.indexNumber
    ? parseIndexNumberAndYear(track.indexNumber).formatted
    : '';

  const targetDoc = allDocs.find((d) => {
    if (d.id === track.documentId) return true;
    const dIdx = d.indexNumber || d.fieldValues?.['رقم_الفهرس'];
    if (dIdx && trackIdxFormatted) {
      return parseIndexNumberAndYear(dIdx, d.year).formatted === trackIdxFormatted;
    }
    return false;
  });

  if (!targetDoc) return null;

  const pubRefText = formatPublicationReferenceText(track);
  const nowIso = new Date().toISOString();

  const nextFieldValues: Record<string, string> = {
    ...(targetDoc.fieldValues || {}),
    مراجع_الشهر: pubRefText,
    المحافظة_العقارية: track.conservationOffice || '',
  };
  if (track.volume) nextFieldValues['مجلد_الشهر'] = track.volume;
  if (track.number) nextFieldValues['رقم_الشهر'] = track.number;
  if (track.publicationDate) nextFieldValues['تاريخ_الشهر'] = track.publicationDate;

  const nextStatus: ContractLifecycleStatus =
    track.status === 'published'
      ? 'published'
      : track.status === 'rejected'
      ? 'publication_rejected'
      : 'deposited_for_publication';

  const updatedDerived = (targetDoc.derivedDocuments || []).map((dd) => ({
    ...dd,
    fieldValues: {
      ...(dd.fieldValues || {}),
      مراجع_الشهر: pubRefText,
    },
  }));

  const updatedDoc: SavedDocument = {
    ...targetDoc,
    conservationOffice: track.conservationOffice,
    wilaya: track.wilaya || targetDoc.wilaya,
    status: nextStatus,
    fieldValues: nextFieldValues,
    derivedDocuments: updatedDerived,
    updatedAt: nowIso,
  };

  await saveDocumentRecord(updatedDoc);

  await saveDocumentRevision({
    id: `rev_pub_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    documentId: updatedDoc.id,
    documentTitle: updatedDoc.title,
    author: authorName,
    bodyHtml: updatedDoc.bodyHtml,
    fieldValues: nextFieldValues,
    createdAt: nowIso,
    summary: `تحديث مسار الشهر العقاري (${
      track.status === 'published'
        ? 'مشهر'
        : track.status === 'rejected'
        ? 'رفض/تحفظ عقاري'
        : 'قيد الإيداع'
    }): ${pubRefText}`,
  });

  return updatedDoc;
}
