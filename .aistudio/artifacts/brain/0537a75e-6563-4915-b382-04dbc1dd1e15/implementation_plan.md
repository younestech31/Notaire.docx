# Implementation Plan - P0 Contract Export Fix (Dual Paths & HTML Leakage Guard)

## Objective
Implement the 5 steps for robust contract export:
1. Dual export workflows: Direct body export (`exportEditorBodyToDocx`) vs Standard template export (`exportWithStandardTemplate`).
2. HTML leakage guard in DOCX serialization engine (`lib/docx-engine.ts`) to prevent inline style tags (`font-family`, `line-height`, etc.) leaking as raw text.
3. Mandatory body purification before all exports and removal of trailing empty paragraphs.
4. Safe pagination injection and exact `pageCount` extraction for `pages2lettre_acte`.
5. Comprehensive unit tests in `tests/notarial-core.test.ts`.

## Proposed Changes

### 1. Dual Export Paths (`lib/document-io-workflows.ts` & UI components)
- Implement `exportEditorBodyToDocx` (Path A): pure direct export of current editor body (`bodyHtml`), header/footer, purification, and pagination without template shell wrapping.
- Implement `exportWithStandardTemplate` (Path B): template-based export with guard against raw full-body HTML insertion into `[contenu_acte]` if parties are empty.
- Update export dropdown/buttons to provide "تصدير Word (المتن الحالي)" as primary and template export as secondary.

### 2. HTML Leakage Guard (`lib/docx-engine.ts`)
- In `serializeInlineNodesToRunsXml` and text serialization, sanitize or validate runs so CSS property names or unescaped style strings do not appear as raw text in `<w:t>`.
- Add guard check: throw an informative error if CSS strings leak into text nodes.

### 3. Mandatory Purification & Clean Up
- Ensure `generateNotaryDocxBlob` and workflow functions always run `purifyNotarialBodyHtml`.

### 4. Safe Pagination & `pageCount`
- Extract exact `pageCount` from injected `[data-notary-pagination="true"]` markers and update field values.

### 5. Unit Tests (`tests/notarial-core.test.ts`)
- Add tests covering purification, direct export vs template export guards, and pagination safety.

## Verification
- Run `npm test` and `compile_applet`.
