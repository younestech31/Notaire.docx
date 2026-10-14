# Algerian Notarial Deed Structure & Single-Page Preamble Fix

Implementation plan to align document generation with genuine Algerian notarial deed practice: eliminate repeating Word page headers, place the registration box (*إطار خاص بالتسجيل*) and office preamble exclusively on the first page body, and ensure all subsequent pages contain only legal clauses and bottom-left continuation pagination (`../..`).

### User Review & Critical Decisions

> [!IMPORTANT]
> The clarifying choices confirmed by the user govern this implementation:
> - **Confirmed Decision 1**: The document header will be placed directly within the Page 1 body content. Word repeating headers (`w:headerReference w:type="default"`) will be completely eliminated so subsequent pages remain clean without repeating headers.
> - **Confirmed Decision 2**: Page 1 will feature the standard registration box (*إطار خاص بالتسجيل*) as a bordered rectangular frame at the very top of the text column, followed by the deed title, notary credentials, the centered transition marker (*حـــــــــــــضـــــــــــــر*), and party recitals.
> - **Confirmed Invariant**: Arial 13pt, 1.0 exact line spacing (`w:line="260"`), 12cm text column (7cm right, 2cm left, 1cm top, 6cm bottom margins), full justification, and continuation markers (`../..`) fixed at the bottom-left of non-final pages.

---

### 1. Overview & Core Concept

- **What It Does**: Refactors the document generation pipeline so that authentic Algerian notary deeds match the exact visual and structural reality shown in official ministerial practice:
  - **Page 1**: Customized introductory page containing the registration box (*إطار خاص بالتسجيل*), deed title (*عقد بيع ترقوي* / *أصل عقد توثيقي*), notary office preamble, dashed separator, *حـــــــــــضـــــــــــر*, and party identity statements.
  - **Pages 2+**: Pure contractual body text, terms, clauses, and tables with zero header repetition, culminating in bottom-left continuation pagination (`الصفحة الثانية ../..`) and the closing ordinal on the final page (`الصفحة ... والأخيرة`).
- **Target Audience**: Algerian notary offices requiring 100% compliance with Chamber of Notaries and land registry standards.
- **Key Value**: Eradicates the illegal repeating header bug in exported `.docx` deeds while matching the exact layout of the reference deed scan.

---

### 2. User Experience & Document Structure

#### Page 1 Architecture (First Page Only)
```
┌────────────────────────────────────────────────────────┐
│  إطار خاص بالتسجيل                                     │
│  (Registration Frame: 120mm wide x ~30-35mm height)   │
└────────────────────────────────────────────────────────┘

              عقد بيع ترقوي / أصل عقد توثيقي
            (Title: Centered, Bold, Underlined)

لدى الأستاذ: [notaire_nom] الموثق، الواقع مكتبه العمومي...
----------------------------------------------------------
                     حـــــــــــــضـــــــــــــر
                      (Centered Notarial Marker)

المؤسسة الوطنية للترقية العقارية / الأطراف المتعاقدة...
(Party Identification & Legal Capacities)
...
----------------------------------------------------------
الصفحة الاولى ../..
```

#### Subsequent Pages Architecture (Pages 2, 3... N)
```
[Continuation of contract body, clauses, terms, and tables]
[Full justification, strict Arial 13pt, line spacing 260]

...
----------------------------------------------------------
الصفحة الثانية ../.. (or final closing formula on last page)
```

---

### 3. Key Technical Decisions & Trade-Offs

- **Decision 1: Removal of Word OpenXML `<w:headerReference>`**:
  - *Chosen Approach*: Disable the generation of default Word headers (`word/header1.xml` with `w:type="default"`) for notarial deeds. Integrate the registration box and notary preamble directly into the Page 1 document body flow.
  - *Why*: In OpenXML, `w:type="default"` forces Word to display the header on every single page of the section. Notarial deeds strictly forbid headers on subsequent pages.
  - *Alternative Considered*: `w:type="first"` with `w:titlePg`. While possible, Word headers have rigid margin constraints that interfere with the registration stamp box dimensions. Embedding the frame directly in the document body guarantees exact millimeter control and seamless pagination.

- **Decision 2: Registration Box (*إطار خاص بالتسجيل*) Implementation**:
  - *Chosen Approach*: Implement as a single-cell bordered table (`border: 1px solid #000000; width: 100%; height: ~35mm;`) containing the text "إطار خاص بالتسجيل" right-aligned at the top, leaving the remaining box space clear for official tax and registration stamps.
  - *Why*: Generates clean `<w:tbl>` with exact borders in Word OpenXML, identically matching the uploaded deed sample `1.png`.

- **Decision 3: Separation of Separator Dashes and Continuation Marks**:
  - *Chosen Approach*: Add the clean notarial dashed separator (`----------------------------------------`) above the continuation marker on page breaks and above *حضر*, matching the reference deed.

---

### 4. Technical Architecture & System Flow

```
┌────────────────────────────────────────────────────────┐
│           Built-in Notary Template Definitions         │
│  - Page 1 Body: Registration Box + Preamble + Title   │
│  - Body Section: Native contract clauses and terms    │
│  - No repeating headerHtml configuration               │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│             Notarial Body Purifier & Splicer           │
│  - Purifies paragraphs to Arial 13pt, justified        │
│  - Slices [contenu_acte] into native sibling blocks    │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│             Notarial Pagination Engine                 │
│  - Page 1 accounts for registration box & preamble     │
│  - Non-final pages: "الصفحة {الترتيب} ../.." at bottom │
│  - Final page: "الصفحة {الترتيب} والأخيرة"             │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│            docx-engine Serialization (OpenXML)         │
│  - Omit <w:headerReference w:type="default"/>          │
│  - Emit registration table & title in Page 1 body      │
│  - Emit <w:line="260"/>, <w:jc w:val="left"/>          │
└────────────────────────────────────────────────────────┘
```

#### Affected Modules:
1. `lib/notary-word-templates.ts`:
   - Incorporate the "إطار خاص بالتسجيل" registration box and proper deed preamble into the Page 1 body of standard templates (`acte_original`, `contrat_reference`, `expedition`).
   - Clear `headerHtml` from standard templates to eliminate repeating page headers.
2. `lib/docx-engine.ts`:
   - Ensure `generateNotaryDocxBlob` does not inject repeating headers into OpenXML section properties for notarial templates.
   - Preserve clean table borders and paragraph properties.
3. `lib/document-io-workflows.ts`:
   - Ensure `exportEditorBodyToDocx` and template exports omit repeating headers by default unless explicitly customized.
