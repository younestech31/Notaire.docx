# Implementation Plan - Dense Notarial Body Purification & Clean Pagination

## Objective
Implement strict notarial body purification (zero empty lines, paragraph-after-paragraph density, fixed 7/2/1/6 margins, clean export/preview-only pagination) as requested.

## Proposed Changes

### 1. New Utility: `lib/notarial-body-purifier.ts`
- Implement `purifyNotarialBodyHtml(html: string, options?: { keepSingleEmptyShell?: boolean }): string`.
- Removes empty `<p>`/`<div>` without visible text, folds consecutive `<br>`, strips old pagination markers, preserves tables, images, `.smart-tag`, clause headers, and enforces strict base inline styles (`margin:0; line-height:1; font-family:Arial; font-size:13pt`).

### 2. Hook up Purification & Prevent Source Empty Paragraphs
- **`lib/docx-importer.ts`**: Return `''` for empty paragraphs in `convertXmlParagraphToHtml`, and purify final imported HTML.
- **`lib/editor-utils.ts`**: Update `sanitizePastedWordHTML` to ignore empty lines and run `purifyNotarialBodyHtml` on output.
- **`lib/docx-engine.ts`**: Purify `bodyHtml` (and header/footer) before merging placeholders and injecting pagination in `generateNotaryDocxBlob`. Pass real `pageCount`.
- **`lib/indexing-engine.ts`**: Purify contract body before sealing original.

### 3. Clean Pagination Layer (`lib/notarial-pagination.ts`)
- Ensure stored document drafts strip pagination markers before saving to IndexedDB/localStorage.
- Apply `purifyNotarialBodyHtml` -> `injectNotarialPaginationIntoHtml` exclusively during preview and `.docx` export.

### 4. Unit Tests (`tests/notarial-core.test.ts`)
- Add tests for body purification, empty paragraph dropping, pagination on purified text, and `pages2lettre_acte` accuracy.

## Verification
- Run `npm test` and `compile_applet`.
