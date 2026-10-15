# Implementation Plan - Fixed-Line Grid Notarial Pagination

## Goal
Replace dynamic Word footer `PAGE` numbering and variable DOM-height pagination with a strict fixed-line grid pagination system in `lib/notarial-pagination.ts` and `lib/docx-engine.ts`, adhering strictly to Algerian notarial formatting rules (46 lines per page, exact 13pt / 260 twip line height, word pagination markers like "الصفحة الأولى ../.." on continuation pages and "… والأخيرة" on the final page, left-aligned on the last line of the page grid, followed by page breaks).

## User Requirements & Technical Specification
1. **Remove Word Footer Pagination**: Remove `footer1.xml` generation and `PAGE` field insertions in `lib/docx-engine.ts`. All pagination markers must be injected directly into the document body text before OOXML serialization.
2. **Fixed-Line Grid in `lib/notarial-pagination.ts`**:
   - Use `NOTARIAL_PAGE_CONTENT_LINES` and `STRICT_LINE_SPACING_TWIPS` (260 twips = 13pt).
   - Strip existing pagination (`stripNotarialPaginationFromHtml`).
   - Collect atomic block elements (`<p>`, `<li>`, `<table>`).
   - Fill each page up to `LINES` (46 lines). Pad remaining space with blank lines/paragraphs if needed to ensure exact vertical grid alignment.
   - Insert page marker paragraph:
     - Non-final pages: `buildNotarialPaginationContinueText(n)` -> «الصفحة الأولى ../..»
     - Final page: `buildNotarialPaginationClosingText(total)` -> «الصفحة … والأخيرة»
   - Follow continuation marker with `<div class="page-break" data-page-break="true"></div>` (no text).
3. **OOXML Serialization in `lib/docx-engine.ts`**:
   - For `data-notary-pagination="true"` paragraphs:
     ```xml
     <w:pPr>
       <w:spacing w:before="0" w:after="0" w:line="260" w:lineRule="exact"/>
       <w:jc w:val="left"/>
     </w:pPr>
     ```
     (No `<w:bidi/>` on marker paragraphs).
   - `div.page-break` -> `<w:br w:type="page"/>` only.
   - Body paragraphs: `line=260`, `lineRule=exact`, `before/after=0`.
4. **Page Count Calculation**:
   - `pageCount = number of p[data-notary-pagination]`.
   - Pass `pageCount` to `normalizeAllNotaryFieldValues` for `pages2lettre_acte`.
5. **Export Pipeline**:
   - Purify -> Merge (plain text) -> insertActeBodyHtml -> injectNotarialPaginationIntoHtml (fixed grid) -> serialize to document.xml.
