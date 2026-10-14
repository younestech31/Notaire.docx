# Footer-Based Notarial Pagination & Natural Body Flow Architecture

Implementation plan to transition Algerian notarial document generation from brittle mid-paragraph body DOM splitting to native Word OpenXML Footer (`w:footerReference` / `word/footer1.xml`) with dynamic field codes (`w:fldSimple w:instr="PAGE"`), leaving the entire contract body to flow naturally as unbroken paragraphs.

### User Decisions Confirmed

> [!IMPORTANT]
> The clarifying choices confirmed by the user govern this implementation:
> - **Decision 1**: Page numbering will be generated natively in the Word footer using dynamic OpenXML page fields (`PAGE` / `NUMPAGES`) formatted with the Algerian notarial continuation structure (`الصفحة {PAGE} ../..`).
> - **Decision 2**: Contract body paragraphs and clauses will flow completely naturally without DOM element splitting or artificial mid-sentence/mid-paragraph breaks.

---

### 1. The Core Problem & Architectural Shift

| Dimension | Previous In-Body Slicing Model | New Dynamic Footer Model |
|---|---|---|
| **Mechanism** | Measured DOM heights in pixels and inserted marker `<p>` and `<div class="page-break">` between (or splitting) paragraphs. | Emits unbroken `<p>` blocks into `<w:body>`. The Word pagination engine natively lays out pages. |
| **Integrity** | Paragraphs and legal clauses could be fractured, tags cut in half, or spans split across pages. | **100% clause integrity**: Text, spans, numbers, and dates remain intact. Word splits lines naturally across pages based on strict Arial 13pt metrics. |
| **Pagination Position** | Required dynamic line spacing (`data-spacing-before-twips`) to push marker to bottom; sensitive to subtle font rendering differences. | **Fixed in footer area**: Anchored precisely inside the 6cm bottom margin (`w:pgMar w:bottom="3402" w:footer="1134"`), left-aligned (`w:jc w:val="left"`). |
| **Word Re-layout Safety** | If the notary opened the file in Word and edited or added words, the hard-coded body markers shifted into the middle of the page. | **Self-healing**: Any text addition or deletion in Word automatically updates page counts and keeps continuation markers at the true page bottom. |

---

### 2. OpenXML Footer Specifications (`word/footer1.xml`)

#### A. Document Relationship and Section Reference
In `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rIdFooter1"
  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
  Target="footer1.xml"/>
```

In `word/document.xml` section properties (`<w:sectPr>`):
```xml
<w:sectPr>
  <w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                     w:type="default"
                     r:id="rIdFooter1"/>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="567" w:right="3969" w:bottom="3402" w:left="1134" w:header="567" w:footer="1134" w:gutter="0"/>
  <w:bidi/>
</w:sectPr>
```

#### B. `word/footer1.xml` Structure
The footer contains a left-aligned, right-to-left paragraph with the strict Algerian continuation notation:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:bidi/>
      <w:jc w:val="left"/>
      <w:spacing w:before="0" w:after="0" w:line="260" w:lineRule="exact"/>
    </w:pPr>
    <!-- Separator rule above footer marker -->
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="26"/>
        <w:szCs w:val="26"/>
        <w:b/>
        <w:rtl/>
      </w:rPr>
      <w:t xml:space="preserve">الصفحة </w:t>
    </w:r>
    <!-- Dynamic Word PAGE field -->
    <w:fldSimple w:instr="PAGE">
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
          <w:sz w:val="26"/>
          <w:szCs w:val="26"/>
          <w:b/>
          <w:rtl/>
        </w:rPr>
        <w:t>1</w:t>
      </w:r>
    </w:fldSimple>
    <!-- Continuation notation ../.. or closing note -->
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="26"/>
        <w:szCs w:val="26"/>
        <w:b/>
        <w:rtl/>
      </w:rPr>
      <w:t xml:space="preserve"> ../..</w:t>
    </w:r>
  </w:p>
</w:ftr>
```

Optional Advanced Word Conditional:
If supported by the Word environment, we can also generate the native Word `IF` field:
`IF { PAGE } = { NUMPAGES } "الأخيرة" "../.."` so the final page seamlessly displays `الأخيرة` while earlier pages display `../..`.

---

### 3. Affected Modules & Implementation Steps

#### Step 1: `lib/docx-engine.ts`
- **Add Native Dynamic Notarial Footer Generator**:
  - Implement `buildNotaryDynamicFooterXml(options)`: generates `word/footer1.xml` containing left-aligned `الصفحة {PAGE} ../..` with exact 260 twip line spacing and Arial 13pt bold font.
  - In `generateNotaryDocxBlob`:
    - Register `word/footer1.xml` in `[Content_Types].xml` and `word/_rels/document.xml.rels`.
    - Set `<w:footerReference w:type="default" r:id="rIdFooter1"/>` in `<w:sectPr>` whenever `pageNumberingEnabled !== false`.
- **Bypass Body Slicing When Footer Pagination Is Active**:
  - When dynamic footer numbering is active, avoid passing the document body through height-measuring or element-splitting logic (`injectNotarialPaginationIntoHtml`), allowing paragraphs to remain pure continuous flow.
  - Only explicit manual user breaks (`<div class="page-break">`) will emit `<w:br w:type="page"/>`.

#### Step 2: `lib/notarial-pagination.ts`
- Provide a clean pass-through mode in `injectNotarialPaginationIntoHtml` that preserves raw clauses when footer mode is selected, while maintaining compatibility for visual preview in the web UI.

#### Step 3: `lib/notary-word-templates.ts` & `lib/document-io-workflows.ts`
- Update standard template exports and direct contract downloads to enable dynamic footer pagination by default (`pageNumberingEnabled: true`), keeping `headerHtml: ''` and embedding Page 1 preamble directly in Page 1 body.

---

### 4. Verification Plan

1. **Compilation**: Run `compile_applet` to ensure zero TypeScript errors.
2. **Linting**: Run `lint_applet` to verify clean code standards.
3. **OpenXML Output Check**:
   - Inspect the generated `.docx` XML structure:
     - Verify `word/footer1.xml` contains `<w:fldSimple w:instr="PAGE"/>` aligned to `left`.
     - Verify `word/document.xml` body has zero broken paragraphs or embedded continuation markers.
     - Verify no `<w:headerReference>` is created.
4. **End-to-End Verification**: Test exporting a multi-page deed and verifying clean layout.
