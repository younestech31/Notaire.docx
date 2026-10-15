# Revised Implementation Plan - Final Architecture Alignment & Template Unification

This plan addresses all specific architectural requirements, fixes file paths and canonical symbols, and establishes rigorous acceptance criteria for template unification, code purging, and UI warnings.

## User Feedback & Mandatory Corrections

1. **Correct File Path**: Modify templates in **`lib/notary-word-templates.ts`** (not `office-templates.ts`).
2. **Canonical Symbols**: Use exact canonical notary symbols (`${num_acte}`, `${date_ar}`, `${place}`, etc.), strictly avoiding non-canonical keys like `${num_repertoire}`.
3. **Template Structure for All 6 Templates** (أصل العقد، ملف مرجعي، نسخة، مستخرج، شهادة، إلخ):
   - Wrap parties section in `<div data-template-part="parties" class="parties-wrapper">...</div>`.
   - Use `${block}` ... `${titre}` ... `${clause}` ... `${/block}` for clauses.
   - Zero legacy `[contenu_acte]`, `[partie1_*]`, or `[num_repertoire]` in template `bodyHtml`.
   - Update `supportedTags` to match `${}` canonical symbols only.
4. **Purge Scope**:
   - **Purge**: In-text legacy pagination injection (`injectNotarialPaginationIntoHtml` and related tests/dead imports).
   - **Keep**: Official digital footer (`footer1.xml` with `fldSimple PAGE` and `../..`, aligned left).
5. **Editor Warning Banner**:
   - Monitor `bodyHtml` (or active clause/document state) in editor (`EditorA4Canvas` or main editor view).
   - Display a gentle amber notification banner when `${...}` template syntax is detected inside clause text, instructing the user that template tags belong in templates rather than clauses.

---

### Step-by-Step Execution Plan

#### Phase 1: Upgrade All 6 Built-In Templates in `lib/notary-word-templates.ts`
- Ensure every built-in template uses:
  - `<div data-template-part="parties" class="parties-wrapper">...</div>`
  - `${block}` loop syntax
  - Canonical `${num_acte}`, `${date_hijri}`, etc.
- Update `supportedTags` arrays for all templates.

#### Phase 2: Purge Legacy In-Text Pagination & Dead Code
- Remove `injectNotarialPaginationIntoHtml` import and usage from `docx-engine.ts`.
- Clean up dead unit tests checking in-text pagination injection.
- Keep footer generation (`footer1.xml`) intact.

#### Phase 3: Add Editor Template Syntax Warning Banner
- Add reactive check in editor UI (`EditorA4Canvas` or wrapper): `bodyHtml.includes('${')`.
- Render subtle amber alert banner above editor workspace when triggered.

#### Phase 4: Verification & Acceptance Tests
- Run `npm test` and `npm run build`.
- Verify zero legacy `[]` tags in built-in template definitions.
- Verify `footer1.xml` has `PAGE` and `../..`.
