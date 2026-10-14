# Notarial Word Export & Column-Bottom Pagination Fix

Implementation plan to repair the Algerian notarial Word (.docx) export pipeline, prevent body text tag escaping and broken nested DOM structures, and lock notarial pagination marks (`../..`) to the bottom-left of the writing column on non-final pages while placing closing ordinals exclusively on the final page.

### User Review & Critical Decisions

> [!IMPORTANT]
> The clarifying choices confirmed by the user in Phase 1 govern the technical approach:
> - **Confirmed Decision 1**: Standard Word templates will have `[contenu_acte]` replaced by native Word body paragraphs instead of injecting HTML strings inside a single `<p>` or `<span>`.
> - **Confirmed Decision 2**: Notarial pagination continuation markers (`الصفحة ... ../..`) are strictly fixed to the bottom-left of the writing column on every non-final page, preventing vertical jumping or sentence splitting.
> - **Confirmed Invariant**: Strictly preserve Arial 13pt, exact margins, line spacing (260 twips), color palette, and layout without adding non-notarial markup or UI alterations.

---

### 1. Overview & Core Concept

- **What It Does**: Ensures standard Algerian notary contract exports (specifically *أصل العقد*, *ملف العقد المرجعي*, and *نسخة العقد*) produce clean, fully compliant `.docx` documents where the body text consists of genuine document-level paragraphs rather than escaped HTML embedded within an enclosing paragraph. In addition, notarial pagination markers (`الصفحة {رقم بالحروف} ../..`) will appear consistently at the bottom-left of the writing column for non-final pages, and the closing formula (`الصفحة {الأخيرة} والأخيرة`) appears only at the document termination.
- **Target Audience**: Algerian notary offices producing legally binding authentic deeds (*أصل العقد*, *النسخة التنفيذية*, *العقد المرجعي*) that demand strict adherence to ministerial formatting standards.
- **Key Value**: Eliminates document corruption, raw HTML string leaks (such as `<span style="...` splitting), and jumping pagination markers across exported deeds.

---

### 2. User Experience & Document Structure

- **Standard Template Paragraph Integration**:
  - In `lib/notary-word-templates.ts`, the placeholder `<p dir="rtl" style="${pBase}">[contenu_acte]</p>` in *أصل العقد*, *ملف العقد المرجعي*, and *نسخة العقد* will be replaced with native document paragraphs when generating the document body, preventing invalid DOM nesting like `<p><span><p>...</p></span></p>`.
  - The export workflow will splice the sanitized contract body blocks directly into the template's clause section as peer block-level elements.
- **Notarial Pagination & Column-Bottom Positioning**:
  - Non-final pages: Mark `الصفحة {العدد الترتيبي مؤنث} ../..` aligned left (`w:jc w:val="left"` with `w:line="260"`) positioned at the bottom of the writing column.
  - Final page: Closing marker (`الصفحة {العدد الترتيبي مؤنث} والأخيرة`) placed only once at the conclusion of the contract text.
  - Suppression of placeholder text: Remove UI debug labels like `نهاية الصفحة X — بداية الصفحة Y` from appearing anywhere in the resulting `document.xml`.

---

### 3. Key Technical Decisions & Trade-Offs

- **Decision 1: Native Block Insertion vs String-Level Placeholder Replacement**:
  - *Chosen Approach*: Replace the `[contenu_acte]` slot at the block level by parsing template blocks and inserting the sanitized contract body paragraphs directly as siblings into the document XML tree.
  - *Why*: Injecting rich contract body HTML into text placeholder nodes escapes tags or produces invalid nested paragraphs (`<p>` inside `<span>` inside `<p>`), breaking Word's XML hierarchy.
  - *Alternative Considered*: String regex replacement of `[contenu_acte]`. Rejected because standard template wrappers create enclosing `<p>` tags that cannot validly contain child `<p>` tags.

- **Decision 2: Paragraph Alignment Mapping in OpenXML**:
  - *Chosen Approach*: In `buildParagraphPropertiesXml`, ensure that pagination markers and left-aligned elements explicitly emit `<w:jc w:val="left"/>` even in RTL documents, without erroneous bidirectional inversion.
  - *Why*: Algerian notarial standards mandate that continuation marks appear at the left side of the writing column (`w:jc w:val="left"`).

- **Decision 3: Elimination of Visual Break Elements in OpenXML**:
  - *Chosen Approach*: When serializing `.page-break` elements in `docx-engine.ts`, emit only `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` and completely discard inner text nodes containing "نهاية الصفحة".

---

### 4. Technical Architecture & System Flow

```
┌────────────────────────────────────────────────────────┐
│             Contract Body (Editor Content)             │
│            Purified & Normalized Block Nodes            │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│            Template Block Splice & Expansion           │
│  - Replaces [contenu_acte] block with native <p> nodes │
│  - No string escaping inside <p> or <span> wrappers    │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│             Notarial Pagination Alignment              │
│  - Continuation mark: "الصفحة ... ../.." on non-final │
│  - Fixed to bottom-left of writing column              │
│  - Closing mark: "الصفحة ... والأخيرة" on last page   │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│            OpenXML Serialization (docx-engine)         │
│  - Emits <w:jc w:val="left"/> & <w:line="260"/>        │
│  - Emits clean <w:br w:type="page"/> (no debug text)   │
│  - Outputs valid, pristine Word .docx archive          │
└──────────────────────────┘
```

#### Affected Modules:
1. `lib/notary-word-templates.ts`: Update built-in templates so the contract body section accommodates native block insertion.
2. `lib/docx-engine.ts`:
   - Block-level substitution for `[contenu_acte]` / contract body elements.
   - Enforce exact `<w:jc w:val="left"/>` for pagination markers.
   - Ensure page breaks emit only `<w:br w:type="page"/>` without text nodes.
3. `lib/notarial-pagination.ts`: Ensure markers are placed cleanly at the column bottom without slicing sentences or creating fragmented HTML spans.
4. `lib/field-values-normalizer.ts`: Verify `contenu_acte` remains excluded from raw string escaping.
