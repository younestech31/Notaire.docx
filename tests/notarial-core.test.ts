import test from 'node:test';
import assert from 'node:assert/strict';
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
} from '../lib/indexing-engine.ts';
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

test('Notarial Pagination: Arabic feminine ordinal closing line', () => {
  assert.equal(getArabicOrdinalPageWord(1), 'الأولى');
  assert.equal(getArabicOrdinalPageWord(9), 'التاسعة');
  assert.equal(
    buildNotarialPaginationClosingText(5),
    'الصفحة الخامسة والأخيرة'
  );
});

test('Notarial Pagination: splitHtmlParagraphPreservingTags preserves bold, underline, and smart-tag chips', () => {
  const html =
    'حضر لدى الموثق <strong>السيد <span class="smart-tag" data-var="اسم_البائع">[اسم_البائع]</span> المولود بتاريخ <u>12 جانفي 1980</u> ببلدية الجزائر الوسطى</strong> وصرحوا بما يلي في هذا العقد الرسمي';

  const totalWords = countLogicalWordsInHtml(html);
  assert.ok(totalWords >= 16);

  // Split right inside the <strong> and <u> block (around word 10)
  const split = splitHtmlParagraphPreservingTags(html, 10);
  assert.ok(split !== null);

  // First part must close <u> and <strong> properly and keep smart-tag intact
  assert.ok(
    split.firstHtml.includes(
      '<span class="smart-tag" data-var="اسم_البائع">[اسم_البائع]</span>'
    )
  );
  assert.ok(split.firstHtml.endsWith('</u></strong>'));

  // Second part must reopen <strong><u> and close them properly
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

  // Should inject ../.. for page 1 and الصفحة الثانية والأخيرة at the end
  assert.ok(paginated.includes('../..'));
  assert.ok(paginated.includes('الصفحة الثانية والأخيرة'));
  // Must preserve the smart-tag and underline/strong tags across the split
  assert.ok(
    paginated.includes('<span class="smart-tag" data-var="البائع">[البائع]</span>')
  );
  assert.ok(paginated.includes('<u>تحتوي فقرته على التزام صريح</u>'));

  // Stripping pagination restores clean content without pagination markers
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

  // Next index for 2026 must be 36/2026 (ignoring 99/2025)
  assert.equal(suggestNextAnnualIndexNumber(docs, 2026), '36/2026');

  // Conflict detection finds doc_2 when attempting to use 35/2026 on doc_new
  const conflict = findConflictingIndexedDocument(
    docs,
    'doc_new',
    '35/2026',
    2026
  );
  assert.equal(conflict?.id, 'doc_2');

  // Same index in a different year (2027) is allowed
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
