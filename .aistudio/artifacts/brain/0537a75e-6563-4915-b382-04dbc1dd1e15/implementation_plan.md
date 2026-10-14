# Arabic Textual Notarial Pagination Directly Below Page Paragraphs

Implementation plan to position Algerian notarial page numbering in Arabic words (`الصفحة الأولى ../..` to `الصفحة ... والأخيرة`) directly below the text and clauses of each page, while enforcing **paragraph-level atomic paging** so that paragraphs and clauses are never split across pages.

### User Decisions Confirmed

> [!IMPORTANT]
> The clarifying choices confirmed by the user govern this implementation:
> - **Decision 1: Textual Numbering in Arabic Words**: The page number must remain written in Arabic words:
>   - Non-final pages: `الصفحة الأولى ../..`, `الصفحة الثانية ../..`, `الصفحة الثالثة ../..`, etc.
>   - Final page: `الصفحة ... والأخيرة` (e.g. `الصفحة الرابعة والأخيرة`).
> - **Decision 2: Direct Placement Below Page Writing**: The marker must sit directly below the writing of the clauses and body on that page (above the bottom footer margin).
> - **Decision 3: Atomic Paragraph Paging (No Mid-Paragraph Splitting)**: Paragraphs and clauses will never be sliced in the middle. If a paragraph exceeds the available lines/height of the current page, the entire paragraph moves atomically to the top of the next page.

---

### 1. Document Architecture & Pagination Rules

#### Page Budget & Calculation
- Page height: 297mm (A4). Margins: Top 10mm, Bottom 60mm (notary stamp margin), Right 70mm, Left 20mm.
- Net writable height: 227mm (~40-42 lines at 260 twips / 18px line-height).
- Page 1 accounts for the top registration box (*إطار خاص بالتسجيل*) and deed preamble (~10-12 lines), leaving ~30 lines for clauses.
- Subsequent pages have full ~40 lines available for clauses and tables.

#### Layout Flow per Page:
```
┌────────────────────────────────────────────────────────┐
│ [Clauses, Terms, or Tables on this Page]              │
│ (Fully preserved intact paragraphs - no cut words)     │
│                                                        │
│ [Last complete paragraph fitting on this page]         │
│                                                        │
│ الصفحة الأولى ../..                                   │
│ (Left-aligned, Arial 13pt bold, line spacing 1.0)     │
└────────────────────────────────────────────────────────┘
==================== [PAGE BREAK] ========================
┌────────────────────────────────────────────────────────┐
│ [Next paragraph starts cleanly at the top of Page 2]   │
│ ...                                                    │
│ [Last complete paragraph fitting on Page 2]            │
│                                                        │
│ الصفحة الثانية ../..                                  │
└────────────────────────────────────────────────────────┘
```

---

### 2. Implementation Steps

#### Step 1: `lib/notarial-pagination.ts`
- **Enforce Paragraph-Atomic Paging**:
  - In `paginateHtmlContentByNotarialStandards` and `injectNotarialPaginationIntoHtml`:
    - Completely disable splitting within paragraphs (`splitParagraphBetweenSentences` / word slicing).
    - Measure each block element (paragraphs `<p>`, tables `<table>`, list items `<li>`).
    - When adding the next block would exceed the maximum allowable page height:
      1. Insert the continuation marker paragraph `الصفحة {الترتيب بالحروف} ../..` directly below the previous block.
      2. Insert a clean page break (`<div class="page-break" data-page-break="true"></div>`).
      3. Start the new page with the overflowing block at the top.
      4. Reset the page accumulator and increment page count.
    - Below the final document block (signatures/reading confirmation), append `الصفحة {الترتيب بالحروف} والأخيرة`.

#### Step 2: `lib/docx-engine.ts`
- **Serialize In-Body Textual Markers with Clean OpenXML**:
  - The continuation marker `<p data-notary-pagination="true">` is serialized as:
    ```xml
    <w:p>
      <w:pPr>
        <w:bidi/>
        <w:jc w:val="left"/>
        <w:spacing w:before="120" w:after="0" w:line="260" w:lineRule="exact"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
          <w:sz w:val="26"/>
          <w:szCs w:val="26"/>
          <w:b/>
          <w:rtl/>
          <w:lang w:val="ar-DZ" w:bidi="ar-DZ"/>
        </w:rPr>
        <w:t xml:space="preserve">الصفحة الأولى ../..</w:t>
      </w:r>
    </w:p>
    ```
  - The `<div class="page-break">` emits `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`.
  - Disable `footerReference` so the page number is not duplicated in the Word footer.

#### Step 3: `lib/notary-word-templates.ts` & `lib/document-io-workflows.ts`
- Verify standard templates and exports cleanly invoke this atomic pagination pipeline.

---

### 3. Verification Plan

1. **Compilation**: Run `compile_applet`.
2. **Linting**: Run `lint_applet`.
3. **Execution Verification**: Run automated tests generating a sample 3-page deed:
   - Check that the page marker is in Arabic letters (`الصفحة الأولى ../..`, `الصفحة الثانية ../..`, `الصفحة الثالثة والأخيرة`).
   - Check that it sits directly under the last paragraph of each page.
   - Verify that no paragraphs are split across page boundaries.
