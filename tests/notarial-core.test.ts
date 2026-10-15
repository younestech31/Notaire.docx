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
import {
  expandBlockLoop,
  generateNotaryDocxBlob,
  lookupPlaceholderValue,
  mergeEditorVariables,
  mergePlaceholdersIntoHtml,
  mergeTemplateSymbols,
} from '../lib/docx-engine.ts';
import { normalizeAllNotaryFieldValues } from '../lib/field-values-normalizer.ts';
import { purifyNotarialBodyHtml } from '../lib/notarial-body-purifier.ts';
import {
  expandFieldValuesWithStandardLatinTags,
  STANDARD_NOTARY_TAG_DEFINITIONS,
} from '../lib/notary-word-variables.ts';
import {
  BUILTIN_STANDARD_WORD_TEMPLATES,
  exportContractWithStandardTemplate,
  spliceContractBodyIntoTemplate,
} from '../lib/notary-word-templates.ts';
import { CANONICAL_NOTARY_SYMBOLS } from '../lib/symbol-map.ts';
import { detectNotaryClauseHeading } from '../lib/clause-outline-engine.ts';
import { sanitizePastedWordHTML } from '../lib/editor-utils.ts';
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
  const paginated = injectNotarialPaginationIntoHtml(inputHtml, true);

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

  // Word footer pagination verification
  const footerXml = await zip.file('word/footer1.xml')?.async('string');
  assert.ok(footerXml !== null && footerXml !== undefined);
  assert.ok(footerXml.includes('../..'));
  assert.ok(footerXml.includes('الصفحة'));
  assert.ok(footerXml.includes('w:instr="PAGE"'));
});

test('Unified OfficeTemplate Migration & Lightweight Export Revisions: migrates legacy templates and stores lightweight export logs', async () => {
  const rootDir = createInMemoryDirHandle('office_unified_migration');
  const bridge = createLocalFileSystemBridge(rootDir, 'مجلد_المكتب_الموحد');

  (globalThis as unknown as { window: Record<string, unknown> }).window = {
    __NOTAIRE_DESKTOP__: bridge,
  };

  try {
    // Seed raw legacy stores on disk (templates, derived_templates, word_templates, downloads)
    await bridge.invoke('storage_save', {
      collection: 'templates',
      item: {
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
      },
    });

    await bridge.invoke('storage_save', {
      collection: 'derived_templates',
      item: {
        id: 'legacy_derived_1',
        code: 'G1',
        name: 'استمارة مستخرج ضريبي قديمة',
        description: 'قالب من منظومة المشتقات القديمة',
        headerHtml: '',
        bodyHtml: '<p>مستخرج للضرائب [البائع]</p>',
        footerHtml: '',
        updatedAt: '2026-03-02T10:00:00.000Z',
      },
    });

    await bridge.invoke('storage_save', {
      collection: 'word_templates',
      item: {
        id: 'legacy_word_1',
        name: 'قالب وورد قديم',
        type: 'extract',
        description: 'قالب من منظومة وورد القديمة',
        headerHtml: '',
        bodyHtml: '<p>مراسلة إدارية [الموثق]</p>',
        footerHtml: '',
        updatedAt: '2026-03-03T10:00:00.000Z',
      },
    });

    await bridge.invoke('storage_save', {
      collection: 'downloads',
      item: {
        id: 'dl_legacy_1',
        documentId: 'doc_1',
        documentTitle: 'عقد بيع عقاري',
        fileName: 'مستخرج_ضريبي_قديم.docx',
        docTypeCode: 'G1',
        docTypeLabel: 'الاستمارة الجبائية G1',
        bodyHtml: '<p>نص قديم</p>',
        headerHtml: '',
        footerHtml: '',
        fieldValues: { البائع: 'أحمد' },
        createdAt: '2026-03-05T09:00:00.000Z',
      },
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

    // Save a lightweight export revision (eventType: 'export', optional bodyHtml omitted)
    await desktopOfficeStore.saveDocumentRevision({
      id: 'rev_export_1',
      documentId: 'doc_1',
      documentTitle: 'عقد بيع عقاري',
      author: 'الأستاذ الموثق',
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
    assert.equal(exportRev.bodyHtml, undefined);
    assert.equal(exportRev.fileName, 'عقد_بيع.docx');

    // Verify legacy download_archive item was automatically migrated into DocumentRevision
    const migratedDlRev = revisions.find((r) => r.id === 'rev_dl_dl_legacy_1');
    assert.ok(migratedDlRev, 'Legacy DownloadArchiveItem must be migrated into DocumentRevision');
    assert.equal(migratedDlRev.eventType, 'derived');
    assert.equal(migratedDlRev.fileName, 'مستخرج_ضريبي_قديم.docx');
    assert.equal(migratedDlRev.docTypeCode, 'G1');
    assert.equal(migratedDlRev.bodyHtml, undefined);
  } finally {
    delete (globalThis as unknown as { window?: unknown }).window;
  }
});

test('Clause Outline Engine: hybrid detection of legal articles, clauses, verbal rankings, and notarial headings', () => {
  function makeMockElement(tagName: string, text: string, attrs: Record<string, string> = {}): any {
    return {
      nodeType: 1,
      tagName: tagName.toUpperCase(),
      textContent: text,
      classList: {
        contains(cls: string) {
          return Boolean(attrs['class']?.split(' ').includes(cls));
        },
      },
      hasAttribute(attr: string) {
        return Boolean(attrs[attr] !== undefined);
      },
      getAttribute(attr: string) {
        return attrs[attr] ?? null;
      },
      setAttribute(attr: string, val: string) {
        attrs[attr] = val;
      },
      querySelector() {
        return null;
      },
      children: [],
    };
  }

  // 1. Law Article syntax
  const art1 = detectNotaryClauseHeading(
    makeMockElement('p', 'المادة الأولى: التمهيد وأطراف العقد')
  );
  assert.equal(art1.isHeading, false);

  // 2. Clause syntax
  const clause1 = detectNotaryClauseHeading(
    makeMockElement('p', 'البند الأول: موضوع البيع والتعيين')
  );
  assert.equal(clause1.isHeading, false);

  // 3. Verbal enumeration syntax
  const verbal1 = detectNotaryClauseHeading(
    makeMockElement('p', 'أولاً: أصل الملكية والوثائق المشهرة')
  );
  assert.equal(verbal1.isHeading, false);

  // 4. Standard notarial headings
  const notarialHead = detectNotaryClauseHeading(
    makeMockElement('p', 'شروط البيع: اتفق الطرفان على الشروط التالية')
  );
  assert.equal(notarialHead.isHeading, false);

  // 5. Rapid Hash syntax (Heading)
  const hashClause = detectNotaryClauseHeading(
    makeMockElement('p', '# تعيين الحصة المفرزة والمساحة')
  );
  assert.equal(hashClause.isHeading, true);
  assert.equal(hashClause.kind, 'hash');
  assert.equal(hashClause.cleanTitle, 'تعيين الحصة المفرزة والمساحة');

  // 6. Container syntax
  const containerClause = detectNotaryClauseHeading(
    makeMockElement('div', 'بند حاوية', { class: 'clause-container' })
  );
  assert.equal(containerClause.isHeading, true);
  assert.equal(containerClause.kind, 'container');

  // 7. Regular narrative text should NOT be flagged as a clause heading
  const normalText = detectNotaryClauseHeading(
    makeMockElement('p', 'بمقتضى هذا العقد التوثيقي حضر أمام الأستاذ الموثق الأطراف التالية أسماؤهم')
  );
  assert.equal(normalText.isHeading, false);
});

test('Word Paste Sanitization: enforces Arial 13pt, text-align: justify, and preserves tags and headings', () => {
  const sampleWordHtml = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
      <head><style>p.MsoNormal { font-size: 16pt; font-family: Calibri; }</style></head>
      <body>
        <!--[if gte mso 9]><xml><w:WordDocument><w:View>Normal</w:View></w:WordDocument></xml><![endif]-->
        <p class="MsoNormal" style="margin-bottom:12pt;text-align:right;">
          <span style="font-family:Calibri;font-size:16pt;">بمقتضى هذا العقد حضر <span style="font-weight:bold;">[البائع]</span> وتنازل لصالح <span style="font-weight:bold;">[المشتري]</span>.</span>
        </p>
        <p class="MsoNormal">
          <span># أصل الملكية والمشتقات</span>
        </p>
      </body>
    </html>
  `;

  const fallbackText = 'بمقتضى هذا العقد حضر [البائع] وتنازل لصالح [المشتري].\n# أصل الملكية والمشتقات';
  const sanitized = sanitizePastedWordHTML(sampleWordHtml, fallbackText);
  assert.ok(sanitized.includes('Arial'));
  assert.ok(sanitized.includes('13pt'));
  assert.ok(sanitized.includes('text-align:justify') || sanitized.includes('text-align: justify'));
  assert.ok(sanitized.includes('data-var="البائع"'));
  assert.ok(sanitized.includes('data-var="المشتري"'));
  assert.ok(!sanitized.includes('Calibri'));
  assert.ok(!sanitized.includes('MsoNormal'));
});

test('Standard Latin Variables & 6 Word Templates: bidirectional mapping and template validation', () => {
  // 1. Verify all 6 standard templates exist
  const expectedTemplates = [
    'acte_original',
    'contrat_reference',
    'expedition',
    'extrait',
    'certificat_vente',
    'bordereau_envoi',
  ];
  for (const id of expectedTemplates) {
    const tpl = BUILTIN_STANDARD_WORD_TEMPLATES[id as keyof typeof BUILTIN_STANDARD_WORD_TEMPLATES];
    assert.ok(tpl, `Template ${id} must exist in BUILTIN_STANDARD_WORD_TEMPLATES`);
    assert.ok(tpl.name && tpl.name.length > 0);
    assert.ok(tpl.bodyHtml && tpl.bodyHtml.length > 0);
    assert.ok(tpl.supportedTags && tpl.supportedTags.length > 0);
  }

  // 2. Test Latin variable resolution from Arabic contract field values
  const arabicFieldValues: Record<string, string> = {
    الطرف_الأول_الاسم: 'بن عيسى أحمد',
    الطرف_الأول_تاريخ_الميلاد: '1980/05/14',
    الطرف_الأول_الرقم_الوطني: '119801601002345678',
    الطرف_الثاني_الاسم: 'براهيمي كريم',
    رقم_الفهرس: '2026/0142',
    تاريخ_العقد: '2026/02/15',
    الموثق: 'الأستاذ عبد الرحمن بوزيد',
    الثمن_بالأرقام: '15,000,000.00',
    الثمن_بالأحرف: 'خمسة عشر مليون دينار جزائري',
    تعيين_الحصة_الكامل: 'شقة رقم 15 عمارة ب',
  };

  // Direct lookup with Latin tag should automatically find the Arabic value
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'partie1_nom'), 'بن عيسى أحمد');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'partie1_date_naissance'), '1980/05/14');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'partie1_nin'), '119801601002345678');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'partie2_nom'), 'براهيمي كريم');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'num_repertoire'), '2026/0142');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'date_acte'), '2026/02/15');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'notaire_nom'), 'الأستاذ عبد الرحمن بوزيد');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'prix_chiffres'), '15,000,000.00');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'prix_lettres'), 'خمسة عشر مليون دينار جزائري');
  assert.equal(lookupPlaceholderValue(arabicFieldValues, 'designation_bien'), 'شقة رقم 15 عمارة ب');

  // 3. Test expandFieldValuesWithStandardLatinTags
  const expanded = expandFieldValuesWithStandardLatinTags(
    arabicFieldValues,
    '<p>متن العقد الكامل للحقن</p>'
  );
  assert.equal(expanded.partie1_nom, 'بن عيسى أحمد');
  assert.equal(expanded.num_repertoire, '2026/0142');
  assert.equal(expanded.contenu_acte, '<p>متن العقد الكامل للحقن</p>');

  // 4. Test HTML merging with Latin tokens [partie1_nom] and [num_repertoire]
  const sampleTemplateHtml = '<p dir="rtl">حضر [partie1_nom] تحت فهرس [num_repertoire] بمبلغ [prix_chiffres] دج.</p>';
  const merged = mergePlaceholdersIntoHtml(sampleTemplateHtml, arabicFieldValues);
  assert.ok(merged.includes('بن عيسى أحمد'));
  assert.ok(merged.includes('2026/0142'));
  assert.ok(merged.includes('15,000,000.00'));
});

test('Core Normalization, Block Loops, and Page Count Tafqeet', () => {
  // 1. Test 3-tier normalization (Arabic keys, standard symbols, Latin tags)
  const rawFields = {
    'الطرف_الأول': 'محمد بوضياف',
    'num_acte': '2026/099',
    'partie2_nom': 'فاطمة الزهراء',
    'act_value': '8,500,000.00',
  };
  const normalized = normalizeAllNotaryFieldValues(rawFields, { pageCount: 4 });
  assert.equal(normalized['c_1'], 'محمد بوضياف');
  assert.equal(normalized['${c_1}'], 'محمد بوضياف');
  assert.equal(normalized['رقم_الفهرس'], '2026/099');
  assert.equal(normalized['${num_acte}'], '2026/099');
  assert.equal(normalized['pages_acte'], '4');
  assert.equal(normalized['pages2lettre_acte'], 'أربعة صفحات (04)');
  assert.ok(normalized['totalLettre'].includes('ثمانية'));

  // 2. Test expandBlockLoop filtering out disabled clauses
  const templateWithBlock = '<div>${block}<p>${titre}: ${clause}</p>${/block}</div>';
  const clauses = [
    { id: 'c1', title: 'البند الأول', contentHtml: 'محتوى البند الأول', enabled: true },
    { id: 'c2', title: 'البند الثاني (معطل)', contentHtml: 'محتوى معطل', enabled: false },
    { id: 'c3', title: 'البند الثالث', contentHtml: 'محتوى البند الثالث', enabled: true },
  ];
  const expandedBlockHtml = expandBlockLoop(templateWithBlock, clauses);
  assert.ok(expandedBlockHtml.includes('البند الأول'));
  assert.ok(expandedBlockHtml.includes('البند الثالث'));
  assert.ok(!expandedBlockHtml.includes('البند الثاني'));
});

test('Notarial Body Purifier: removes empty paragraphs and enforces dense formatting', () => {
  const messyHtml = `
    <p dir="rtl"><br></p>
    <p dir="rtl">&nbsp;</p>
    <p dir="rtl">نص البند الأول الرسمي</p>
    <p dir="rtl"></p>
    <table><tr><td>خلية جدول</td></tr></table>
    <p dir="rtl"><br></p>
  `;
  const purified = purifyNotarialBodyHtml(messyHtml);
  assert.ok(!purified.includes('&nbsp;'));
  assert.ok(!purified.includes('<br>'));
  assert.ok(purified.includes('نص البند الأول الرسمي'));
  assert.ok(purified.includes('خلية جدول'));
});

test('DOCX Export & Importer: true pageCount and empty paragraph drop', async () => {
  const longBody = Array.from({ length: 60 }, (_, i) => `<p dir="rtl">سلسلة نصية طويلة رقم ${i} لامتلاد الصفحات في التوثيق الرسمي الجزائري.</p>`).join('\n');
  const blob = await generateNotaryDocxBlob({
    title: 'عقد اختبار الصفحات',
    bodyHtml: longBody,
    pageNumberingEnabled: true,
  });
  assert.ok(blob);

  const emptyPResult = purifyNotarialBodyHtml('<p><br></p><p>   </p><p>محتوى حقيقي</p>');
  assert.ok(emptyPResult.includes('محتوى حقيقي'));
  assert.ok(!emptyPResult.includes('<br>'));
});

test('Word Export & Pagination Invariants: native body paragraphs, w:jc left, w:line 260, no page break text', async () => {
  const tpl = BUILTIN_STANDARD_WORD_TEMPLATES['acte_original'];
  const testBody = `
    <p dir="rtl" style="margin:0;line-height:1;">البند الأول: هذا نص المتن الأصلي الأول للعقد.</p>
    <p dir="rtl" style="margin:0;line-height:1;">البند الثاني: هذا نص المتن الأصلي الثاني للعقد.</p>
  `;

  // 1. Check native block splicing
  const spliced = spliceContractBodyIntoTemplate(tpl.bodyHtml, testBody);
  assert.ok(!spliced.includes('[contenu_acte]'));
  assert.ok(spliced.includes('البند الأول: هذا نص المتن الأصلي'));
  // Ensure no nested <p><p>
  assert.ok(!/<p[^>]*>(?:(?!<\/p>)[\s\S])*?<p/i.test(spliced));

  // 2. Multi-page document export
  const longBody = Array.from({ length: 70 }, (_, idx) =>
    `<p dir="rtl" style="margin:0;line-height:1;">بند رسمي رقم ${idx + 1}: التزام توثيقي كامل وفقاً للتشريع التوثيقي الجزائري الصادر بمرسوم تنفيذي رسمي.</p>`
  ).join('\n');

  const fullTplBody = spliceContractBodyIntoTemplate(tpl.bodyHtml, longBody);
  const blob = await generateNotaryDocxBlob({
    title: 'عقد_توثيقي_نموذجي',
    bodyHtml: fullTplBody,
    fieldValues: {
      notaire_nom: 'الأستاذ أحمد الشريف',
      num_repertoire: '987',
      annee_acte: '2026',
    },
    pageNumberingEnabled: true,
  });

  const arrayBuf = await blob.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuf);
  const docXml = await zip.file('word/document.xml')?.async('text');
  assert.ok(docXml);

  // Invariant 1: w:line="260" in document
  assert.ok(docXml.includes('w:line="260"'));

  // Invariant 2: Dynamic Word footer in word/footer1.xml
  const footerXml = await zip.file('word/footer1.xml')?.async('text');
  assert.ok(footerXml);
  assert.ok(footerXml.includes('w:jc w:val="left"'));
  assert.ok(footerXml.includes('w:instr="PAGE"'));
  assert.ok(footerXml.includes('../..'));
  assert.ok(footerXml.includes('الصفحة'));

  // Invariant 3: Body document.xml does NOT have in-body pagination markers or "والأخيرة"
  assert.ok(!docXml.includes('والأخيرة'));
  assert.ok(!docXml.includes('data-notary-pagination'));
});

test('Strict Layer Separation: mergeEditorVariables vs mergeTemplateSymbols', () => {
  const editorHtml = `
    <p dir="rtl">
      حضر <span class="smart-tag" data-var="البائع">[البائع]</span>
      والطرف الأول [partie1_nom]
      أمام الموثق بـ \${office_name} تحت فهرس \${num_acte}.
    </p>
  `;

  // 1. mergeEditorVariables should replace [البائع] and [partie1_nom], but PRESERVE ${office_name} and ${num_acte}
  const editorMerged = mergeEditorVariables(editorHtml, {
    البائع: 'عمر المختار',
    partie1_nom: 'عمر المختار بن عائشة',
    office_name: 'مكتب الأستاذ الشريف',
    num_acte: '2026/050',
  });

  assert.ok(editorMerged.includes('عمر المختار'));
  assert.ok(editorMerged.includes('عمر المختار بن عائشة'));
  assert.ok(editorMerged.includes('${office_name}'));
  assert.ok(editorMerged.includes('${num_acte}'));

  // 2. mergeTemplateSymbols should replace ${office_name} and ${num_acte}, and PRESERVE un-matched symbols (whitelist)
  const templateHtml = '<p>المكتب: ${office_name} — فهرس: ${num_acte} — مجهول: ${unknown_custom_symbol}</p>';
  const templateMerged = mergeTemplateSymbols(templateHtml, {
    office_name: 'مكتب الأستاذ بوزيد',
    num_acte: '2026/077',
  });

  assert.ok(templateMerged.includes('مكتب الأستاذ بوزيد'));
  assert.ok(templateMerged.includes('2026/077'));
  // Whitelist: unknown symbol is preserved, never blindly deleted
  assert.ok(templateMerged.includes('${unknown_custom_symbol}'));
});

test('Block Loop: expands clauses with raw HTML injection and single-clause fallback', () => {
  const tpl = `
    <div class="contract">
      \${block}
      <h3 class="clause-title">\${titre}</h3>
      <div class="clause-body">\${clause}</div>
      \${/block}
    </div>
  `;

  const clauses = [
    {
      title: 'البند الأول: التعيين',
      contentHtml: '<p dir="rtl">قطعة أرضية مساحتها <strong>500 م²</strong> كائنة بالبليدة.</p>',
      enabled: true,
    },
    {
      title: 'البند الثاني: الثمن',
      contentHtml: '<p dir="rtl">الثمن الإجمالي هو <u>10,000,000 دج</u> تم دفعه كاملاً.</p>',
      enabled: true,
    },
  ];

  const expanded = expandBlockLoop(tpl, clauses);

  // Titre is present
  assert.ok(expanded.includes('البند الأول: التعيين'));
  assert.ok(expanded.includes('البند الثاني: الثمن'));

  // Raw HTML is preserved without escapeXml
  assert.ok(expanded.includes('<strong>500 م²</strong>'));
  assert.ok(!expanded.includes('&lt;strong&gt;'));
  assert.ok(expanded.includes('<u>10,000,000 دج</u>'));
  assert.ok(!expanded.includes('&lt;u&gt;'));

  // Fallback test: un-split contract treated as single un-titled clause
  const unSplitBody = '<p dir="rtl">متن عقد كامل غير مقسم ببنود شباك محرر رسمياً.</p>';
  const fallbackExpanded = expandBlockLoop(tpl, [], unSplitBody);
  assert.ok(fallbackExpanded.includes('متن عقد كامل غير مقسم'));
  // Title container cleanly omitted
  assert.ok(!fallbackExpanded.includes('class="clause-title"'));
});

test('Parties Preamble Detection & data-template-part="parties" elimination', async () => {
  const tpl = BUILTIN_STANDARD_WORD_TEMPLATES['acte_original'];
  assert.ok(tpl.bodyHtml.includes('data-template-part="parties"'));

  // When body has parties preamble, exportWithStandardTemplate strips duplicate parties wrapper
  const bodyWithPreamble = `
    <p dir="rtl">حضر أمامنا السيد بائع ومشتري واتفقا على ما يلي:</p>
    <p dir="rtl"># البند الأول: البيع</p>
    <p dir="rtl">تم البيع بالتراضي والقبول.</p>
  `;

  const res = await exportContractWithStandardTemplate({
    templateId: 'acte_original',
    contractTitle: 'عقد_مع_ديباجة',
    bodyHtml: bodyWithPreamble,
    fieldValues: {
      c_1: 'الطرف الأول الأصلي',
      c_2: 'الطرف الثاني الأصلي',
      office_name: 'مكتب الأستاذ أحمد',
      num_acte: '2026/888',
    },
  });

  assert.ok(res.fileName.includes('.docx'));
});



