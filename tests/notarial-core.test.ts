import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {
  buildNotarialPaginationClosingText,
  countLogicalWordsInHtml,
  getArabicOrdinalPageWord,
  injectNotarialPaginationIntoHtml,
  splitHtmlParagraphPreservingTags,
  stripNotarialPaginationFromHtml,
} from '../lib/notarial-pagination.ts';
import {
  findConflictingIndexedDocument,
  formatPublicationReferenceText,
  formatRegistrationReferenceText,
  isContractOfficiallyIndexed,
  officiallyIndexAndSealContract,
  parseIndexNumberAndYear,
  suggestNextAnnualIndexNumber,
  syncRegistrationBatchToContracts,
} from '../lib/indexing-engine.ts';
import {
  createLocalFileSystemBridge,
  type FileSystemDirectoryHandleLike,
} from '../lib/desktop-bridge.ts';
import { desktopOfficeStore } from '../lib/storage-desktop.ts';
import { generateNotaryDocxBlob } from '../lib/docx-engine.ts';
import type {
  PublicationTrack,
  RegistrationBatch,
  SavedDocument,
} from '../lib/types.ts';

function makeDummyDoc(overrides: Partial<SavedDocument> = {}): SavedDocument {
  return {
    id: 'doc_1',
    title: 'عقد بيع عقاري',
    bodyHtml: '<p dir="rtl">نص العقد الرسمي المحرر بمكتب التوثيق</p>',
    headerHtml: '',
    footerHtml: '',
    pageNumberingEnabled: true,
    showHeaderFooter: false,
    fieldValues: {},
    updatedAt: '2026-03-01T10:00:00.000Z',
    ...overrides,
  };
}

function createInMemoryDirHandle(name = 'notary_office'): FileSystemDirectoryHandleLike {
  const dirs = new Map<string, FileSystemDirectoryHandleLike>();
  const files = new Map<string, Uint8Array>();

  const toUint8 = async (data: string | BufferSource | Blob): Promise<Uint8Array> => {
    if (typeof data === 'string') return new TextEncoder().encode(data);
    if (data instanceof Uint8Array) return new Uint8Array(data);
    if (ArrayBuffer.isView(data)) {
      return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
    }
    if (data instanceof ArrayBuffer) return new Uint8Array(data);
    if (typeof Blob !== 'undefined' && data instanceof Blob) {
      return new Uint8Array(await data.arrayBuffer());
    }
    return new Uint8Array(0);
  };

  return {
    name,
    async getDirectoryHandle(subName: string, options?: { create?: boolean }) {
      const existing = dirs.get(subName);
      if (existing) return existing;
      if (!options?.create) throw new Error(`Directory not found: ${subName}`);
      const created = createInMemoryDirHandle(subName);
      dirs.set(subName, created);
      return created;
    },
    async getFileHandle(fileName: string, options?: { create?: boolean }) {
      if (!files.has(fileName) && !options?.create) {
        throw new Error(`File not found: ${fileName}`);
      }
      return {
        async getFile() {
          const bytes = files.get(fileName) || new Uint8Array(0);
          return {
            async text() {
              return new TextDecoder().decode(bytes);
            },
            async arrayBuffer() {
              return bytes.buffer.slice(
                bytes.byteOffset,
                bytes.byteOffset + bytes.byteLength
              );
            },
          };
        },
        async createWritable() {
          let pending = new Uint8Array(0);
          return {
            async write(data: string | BufferSource | Blob) {
              pending = await toUint8(data);
            },
            async close() {
              files.set(fileName, pending);
            },
          };
        },
      };
    },
  };
}

test('Notarial Pagination: Arabic feminine ordinal closing line (1 to 200)', () => {
  assert.equal(getArabicOrdinalPageWord(1), 'الأولى');
  assert.equal(getArabicOrdinalPageWord(9), 'التاسعة');
  assert.equal(getArabicOrdinalPageWord(30), 'الثلاثون');
  assert.equal(getArabicOrdinalPageWord(31), 'الحادية والثلاثون');
  assert.equal(getArabicOrdinalPageWord(45), 'الخامسة والأربعون');
  assert.equal(getArabicOrdinalPageWord(99), 'التاسعة والتسعون');
  assert.equal(getArabicOrdinalPageWord(100), 'المائة');
  assert.equal(getArabicOrdinalPageWord(105), 'المائة والخامسة');
  assert.equal(
    buildNotarialPaginationClosingText(35),
    'الصفحة الخامسة والثلاثون والأخيرة'
  );
});

test('Notarial Pagination: splitHtmlParagraphPreservingTags preserves bold, underline, and smart-tag chips', () => {
  const html =
    'حضر لدى الموثق <strong>السيد <span class="smart-tag" data-var="اسم_البائع">[اسم_البائع]</span> المولود بتاريخ <u>12 جانفي 1980</u> ببلدية الجزائر الوسطى</strong> وصرحوا بما يلي في هذا العقد الرسمي';

  const totalWords = countLogicalWordsInHtml(html);
  assert.ok(totalWords >= 16);

  const split = splitHtmlParagraphPreservingTags(html, 10);
  assert.ok(split !== null);

  assert.ok(
    split.firstHtml.includes(
      '<span class="smart-tag" data-var="اسم_البائع">[اسم_البائع]</span>'
    )
  );
  assert.ok(split.firstHtml.endsWith('</u></strong>'));
  assert.ok(split.secondHtml.startsWith('<strong><u>'));
  assert.ok(split.secondHtml.includes('</strong>'));
});

test('Notarial Pagination: injectNotarialPaginationIntoHtml splits multi-page contract while preserving inner formatting', () => {
  const repeatedFormattedClause = Array.from({ length: 55 }, (_, idx) =>
    idx === 20
      ? '<strong>بند جوهري مسطر <u>تحتوي فقرته على التزام صريح</u> من البائع <span class="smart-tag" data-var="البائع">[البائع]</span></strong>'
      : 'كلمة توثيقية رسمية في متن العقد العقاري المبرم بين الطرفين'
  ).join(' ');

  const inputHtml = `<p dir="rtl">${repeatedFormattedClause}</p>`;
  const paginated = injectNotarialPaginationIntoHtml(inputHtml, true, true);

  assert.ok(paginated.includes('../..'));
  assert.ok(paginated.includes('الصفحة الثانية والأخيرة'));
  assert.ok(
    paginated.includes('<span class="smart-tag" data-var="البائع">[البائع]</span>')
  );
  assert.ok(paginated.includes('<u>تحتوي فقرته على التزام صريح</u>'));

  const stripped = stripNotarialPaginationFromHtml(paginated);
  assert.ok(!stripped.includes('data-notary-pagination="true"'));
});

test('Indexing Engine: annual sequence suggestion and duplicate conflict detection', () => {
  const docs: SavedDocument[] = [
    makeDummyDoc({ id: 'doc_1', indexNumber: '14/2026', year: 2026 }),
    makeDummyDoc({ id: 'doc_2', indexNumber: '2026/35', year: 2026 }),
    makeDummyDoc({ id: 'doc_3', indexNumber: '99/2025', year: 2025 }),
  ];

  assert.deepEqual(parseIndexNumberAndYear('35/2026'), {
    sequence: 35,
    year: 2026,
    formatted: '35/2026',
  });
  assert.deepEqual(parseIndexNumberAndYear('2026/35'), {
    sequence: 35,
    year: 2026,
    formatted: '35/2026',
  });

  assert.equal(suggestNextAnnualIndexNumber(docs, 2026), '36/2026');

  const conflict = findConflictingIndexedDocument(
    docs,
    'doc_new',
    '35/2026',
    2026
  );
  assert.equal(conflict?.id, 'doc_2');

  const noConflictOtherYear = findConflictingIndexedDocument(
    docs,
    'doc_new',
    '35/2027',
    2027
  );
  assert.equal(noConflictOtherYear, undefined);
});

test('Indexing Engine: rejects empty contract body when indexing', async () => {
  const emptyDoc = makeDummyDoc({
    id: 'doc_empty',
    bodyHtml: '<p dir="rtl"><br></p>',
  });

  await assert.rejects(
    () =>
      officiallyIndexAndSealContract({
        document: emptyDoc,
        allDocuments: [emptyDoc],
        indexNumberInput: '1/2026',
      }),
    /لا يمكن اعتماد أو فهرسة عقد فارغ/
  );
});

test('Indexing Engine: lifecycle lock and rejection formatting for Registration and Publication', () => {
  const rejectedRegDoc = makeDummyDoc({
    status: 'registration_rejected',
    indexNumber: '40/2026',
  });
  const rejectedPubDoc = makeDummyDoc({
    status: 'publication_rejected',
    indexNumber: '41/2026',
  });
  const draftDoc = makeDummyDoc({ status: 'draft' });

  assert.equal(isContractOfficiallyIndexed(rejectedRegDoc), true);
  assert.equal(isContractOfficiallyIndexed(rejectedPubDoc), true);
  assert.equal(isContractOfficiallyIndexed(draftDoc), false);

  const rejectedBatch: RegistrationBatch = {
    id: 'reg_1',
    batchNumber: 1,
    batchReference: 'دفعة مارس',
    year: 2026,
    inspectorateName: 'مفتشية التسجيل والطابع - الجزائر',
    contractIds: ['doc_1'],
    indexNumbers: ['40/2026'],
    totalFees: 15000,
    status: 'rejected',
    statusDate: '2026-03-15',
    notes: 'نقص في أصل شهادة الميلاد',
    updatedAt: '2026-03-15T10:00:00.000Z',
  };
  const regRef = formatRegistrationReferenceText(rejectedBatch);
  assert.ok(regRef.includes('رُفض للتسوية/التصحيح'));
  assert.ok(regRef.includes('نقص في أصل شهادة الميلاد'));

  const rejectedTrack: PublicationTrack = {
    id: 'pub_1',
    documentId: 'doc_1',
    contractTitle: 'عقد بيع عقاري',
    indexNumber: '41/2026',
    conservationOffice: 'المحافظة العقارية ببئر مراد رايس',
    wilaya: 'الجزائر',
    depositDate: '2026-03-20',
    status: 'rejected',
    notes: 'خطأ في رقم القسم المساحي',
    updatedAt: '2026-03-20T10:00:00.000Z',
  };
  const pubRef = formatPublicationReferenceText(rejectedTrack);
  assert.ok(pubRef.includes('رفض/تحفظ بالشهر العقاري'));
  assert.ok(pubRef.includes('خطأ في رقم القسم المساحي'));
});

test('Desktop Bridge & Multi-Contract Batch Sync: storage_save, seal_original, read_original, and syncRegistrationBatchToContracts', async () => {
  const memDir = createInMemoryDirHandle('notaire_disk_test');
  const bridge = createLocalFileSystemBridge(memDir);
  (globalThis as unknown as { window: Record<string, unknown> }).window = {
    __NOTAIRE_DESKTOP__: bridge,
  };

  try {
    const docA = makeDummyDoc({
      id: 'doc_a',
      title: 'عقد بيع شقة رقم 1',
      indexNumber: '10/2026',
      year: 2026,
      status: 'signed_indexed',
    });
    const docB = makeDummyDoc({
      id: 'doc_b',
      title: 'عقد هبة عقارية رقم 2',
      indexNumber: '11/2026',
      year: 2026,
      status: 'signed_indexed',
    });

    await desktopOfficeStore.saveDocumentRecord(docA);
    await desktopOfficeStore.saveDocumentRecord(docB);

    const loaded = await desktopOfficeStore.loadSavedDocuments();
    assert.equal(loaded.length, 2);

    // Test sealing and reading back original with SHA-256 verification
    const sampleBytes = new TextEncoder().encode('NOTAIRE-DOCX-SEAL-TEST');
    const sealed = await desktopOfficeStore.sealContractOriginal({
      documentId: 'doc_a',
      year: 2026,
      indexNumber: '10/2026',
      docxBytes: sampleBytes,
    });
    assert.ok(sealed.sha256.length === 64);

    const readBack = await desktopOfficeStore.readSealedOriginal('doc_a');
    assert.ok(readBack !== null);
    assert.equal(readBack.sha256, sealed.sha256);
    assert.deepEqual(readBack.bytes, sampleBytes);

    // Test multi-contract RegistrationBatch synchronization (registered -> rejected)
    const batchRegistered: RegistrationBatch = {
      id: 'batch_10_11',
      batchNumber: 5,
      year: 2026,
      inspectorateName: 'مفتشية التسجيل والطابع بئر مراد رايس',
      contractIds: ['doc_a'],
      indexNumbers: ['10/2026', '11/2026'],
      totalFees: 45000,
      receiptNumber: 'W-9988',
      registrationDate: '2026-04-02',
      status: 'registered',
      statusDate: '2026-04-02',
    };

    const syncedRegistered = await syncRegistrationBatchToContracts(
      batchRegistered,
      'الأستاذ الموثق'
    );
    assert.equal(syncedRegistered.length, 2);
    for (const d of syncedRegistered) {
      assert.equal(d.status, 'registered');
      assert.equal(d.registrationBatchId, 'batch_10_11');
      assert.ok(d.fieldValues['مراجع_التسجيل']?.includes('W-9988'));
    }

    const batchRejected: RegistrationBatch = {
      ...batchRegistered,
      status: 'rejected',
      receiptNumber: undefined,
      notes: 'مطلوب استكمال شهادة السلبية',
    };
    const syncedRejected = await syncRegistrationBatchToContracts(
      batchRejected,
      'الأستاذ الموثق'
    );
    assert.equal(syncedRejected.length, 2);
    for (const d of syncedRejected) {
      assert.equal(d.status, 'registration_rejected');
      assert.ok(
        d.fieldValues['مراجع_التسجيل']?.includes('مطلوب استكمال شهادة السلبية')
      );
    }
  } finally {
    delete (globalThis as unknown as { window?: unknown }).window;
  }
});

test('DOCX Engine: generates valid .docx with multi-page ../.. pagination, table, and merged variables', async () => {
  const longParagraph = Array.from({ length: 50 }, () =>
    'فقرة توثيقية رسمية محررة بمكتب التوثيق لبيان التزامات المتعاقدين'
  ).join(' ');

  const bodyHtml = `
    <p dir="rtl">حضر السيد <strong>[البائع]</strong> وباع العقار المبين في الجدول أدناه:</p>
    <table dir="rtl">
      <tbody>
        <tr><td>رقم الحصة</td><td>المساحة</td></tr>
        <tr><td>[رقم_الحصة]</td><td>120 م²</td></tr>
      </tbody>
    </table>
    <p dir="rtl">${longParagraph}</p>
  `;

  const blob = await generateNotaryDocxBlob({
    title: 'عقد بيع تجريبي متعدد الصفحات',
    bodyHtml,
    pageNumberingEnabled: true,
    fieldValues: {
      البائع: 'أحمد بن محمد',
      رقم_الحصة: '14',
    },
  });

  const arrayBuffer = await blob.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docXmlFile = zip.file('word/document.xml');
  assert.ok(docXmlFile !== null);

  const docXml = await docXmlFile.async('string');
  assert.ok(docXml.includes('أحمد بن محمد'));
  assert.ok(docXml.includes('<w:tbl>'));
  assert.ok(docXml.includes('../..'));
  assert.ok(docXml.includes('الصفحة الثانية والأخيرة'));
});

test('Unified OfficeTemplate Migration & Lightweight Export Revisions: migrates legacy templates and stores lightweight export logs', async () => {
  const rootDir = createInMemoryDirHandle('office_unified_migration');
  const bridge = createLocalFileSystemBridge(rootDir, 'مجلد_المكتب_الموحد');

  (globalThis as unknown as { window: Record<string, unknown> }).window = {
    __NOTAIRE_DESKTOP__: bridge,
  };

  try {
    // Seed legacy stores (CustomTemplate, DerivedDocTemplate, WordTemplateDefinition)
    await desktopOfficeStore.saveCustomTemplate({
      id: 'legacy_custom_1',
      name: 'قالب عقد بيع قديم',
      category: 'عقود البيع',
      description: 'قالب من المنظومة الأولى',
      headerHtml: '',
      bodyHtml: '<p>نص عقد البيع القديم [البائع]</p>',
      footerHtml: '',
      pageNumberingEnabled: true,
      extractedPlaceholders: ['البائع'],
      updatedAt: '2026-03-01T10:00:00.000Z',
    });

    await desktopOfficeStore.saveDerivedDocTemplate({
      id: 'legacy_derived_1',
      code: 'G1',
      name: 'استمارة مستخرج ضريبي قديمة',
      description: 'قالب من منظومة المشتقات القديمة',
      headerHtml: '',
      bodyHtml: '<p>مستخرج للضرائب [البائع]</p>',
      footerHtml: '',
      updatedAt: '2026-03-02T10:00:00.000Z',
    });

    await desktopOfficeStore.saveWordTemplate({
      id: 'legacy_word_1',
      name: 'قالب وورد قديم',
      type: 'extract',
      description: 'قالب من منظومة وورد القديمة',
      headerHtml: '',
      bodyHtml: '<p>مراسلة إدارية [الموثق]</p>',
      footerHtml: '',
      updatedAt: '2026-03-03T10:00:00.000Z',
    });

    // Load unified OfficeTemplates — should automatically migrate all three legacy templates + built-in derived templates
    const unified = await desktopOfficeStore.loadOfficeTemplates();
    const migratedContract = unified.find((t) => t.id === 'legacy_custom_1');
    const migratedExtract = unified.find((t) => t.id === 'legacy_derived_1');
    const migratedWord = unified.find((t) => t.id === 'legacy_word_1');

    assert.ok(migratedContract, 'Legacy CustomTemplate must be migrated');
    assert.equal(migratedContract.kind, 'contract');

    assert.ok(migratedExtract, 'Legacy DerivedDocTemplate must be migrated');
    assert.equal(migratedExtract.kind, 'extract');
    assert.equal(migratedExtract.code, 'G1');
    assert.equal(migratedExtract.docTypeCode, 'G1');

    assert.ok(migratedWord, 'Legacy WordTemplateDefinition must be migrated');

    // Save a lightweight export revision (eventType: 'export', empty bodyHtml)
    await desktopOfficeStore.saveDocumentRevision({
      id: 'rev_export_1',
      documentId: 'doc_1',
      documentTitle: 'عقد بيع عقاري',
      author: 'الأستاذ الموثق',
      bodyHtml: '',
      fieldValues: {},
      summary: 'تصدير نسخة Word يومية: عقد_بيع.docx',
      createdAt: '2026-03-10T12:00:00.000Z',
      eventType: 'export',
      fileName: 'عقد_بيع.docx',
      docTypeCode: 'acte',
    });

    const revisions = await desktopOfficeStore.loadDocumentRevisions();
    const exportRev = revisions.find((r) => r.id === 'rev_export_1');
    assert.ok(exportRev);
    assert.equal(exportRev.eventType, 'export');
    assert.equal(exportRev.bodyHtml, '');
    assert.equal(exportRev.fileName, 'عقد_بيع.docx');
  } finally {
    delete (globalThis as unknown as { window?: unknown }).window;
  }
});

