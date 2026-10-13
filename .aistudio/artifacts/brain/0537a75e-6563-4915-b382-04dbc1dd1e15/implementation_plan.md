# Implementation Plan - Final Precision Fixes: Live Density & Accurate Pagination

## Objective
Implement the 5 targeted precision fixes:
1. Remove trailing paragraph generator (`trailingP`) in `normalizeNotaryContainerDOM` (`lib/editor-utils.ts`).
2. Reflect purification in editor state after successful save/autosave (`NotaryEditorApp.tsx` / `storage-browser.ts`).
3. Compute exact `pageCount` from pagination injection output in `generateNotaryDocxBlob` (`lib/docx-engine.ts`) before final `normalizeAllNotaryFieldValues`.
4. Drop empty paragraphs at source in `convertXmlParagraphToHtml` (`lib/docx-importer.ts`).
5. Add corresponding unit tests in `tests/notarial-core.test.ts`.

## Proposed Changes

### 1. `lib/editor-utils.ts` (`normalizeNotaryContainerDOM`)
- Remove the code block adding `trailingP` after tables, page breaks, or clauses.
- Retain only the single empty shell check (`<p dir="rtl"><br></p>`) when container has zero text/blocks.

### 2. Editor State Sync After Save (`NotaryEditorApp.tsx`)
- In save/autosave handlers, after `saveActiveDraftSession` or `saveDocumentRecord`, update the editor's active document `bodyHtml` state with the purified version (`purifyNotarialBodyHtml(doc.bodyHtml)`).

### 3. DOCX Export Accurate Page Count (`lib/docx-engine.ts`)
- In `generateNotaryDocxBlob`:
  1. Purify `bodyHtml`.
  2. Merge initial placeholders.
  3. Inject pagination (`injectNotarialPaginationIntoHtml`) -> `mergedBodyHtml`.
  4. Count pages accurately from `mergedBodyHtml` (e.g. counting `data-notary-pagination` elements + 1, or parsing closing pagination text).
  5. Re-run `normalizeAllNotaryFieldValues` with the exact computed `pageCount`.
  6. Regenerate final merged body & headers/footers with updated field values.

### 4. Docx Importer (`lib/docx-importer.ts`)
- In `convertXmlParagraphToHtml`, ensure empty paragraphs without text, page breaks, or tables return `''`.

### 5. Unit Tests (`tests/notarial-core.test.ts`)
- Add unit tests for importer empty paragraph dropping and DOCX export page count accuracy.

## Verification
- Run `npm test` and `compile_applet`.
