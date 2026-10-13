# Implementation Plan - Finalizing Documentation, Docx Pagination & Unit Tests

## Objective
Complete the remaining items from the audit report:
1. Update `README.md` to reflect current architecture (`DB_VERSION = 9`, new core files, actual modal components, block loops, and normalization engine).
2. Pass `pageCount` into `normalizeAllNotaryFieldValues` within `generateNotaryDocxBlob` in `lib/docx-engine.ts`.
3. Add comprehensive unit tests in `tests/notarial-core.test.ts` for:
   - Full 3-tier normalization (Arabic keys, standard symbols, Latin tags).
   - `expandBlockLoop` behavior (excluding disabled clauses).
   - `pages2lettre_acte` accuracy matching passed `pageCount`.

## Proposed Changes

### 1. Documentation (`README.md`)
- Update IndexedDB version references from v7 to v9.
- Document core files in `lib/`: `field-values-normalizer.ts`, `symbol-map.ts`, `notary-word-templates.ts`, `notary-word-variables.ts`.
- Document modal structure (`EditorModalsHost`, `WordTemplatesManagerModal`, `DocumentWorkflowModals`).
- Document support for `${block}...${/block}` and bidirectional symbol normalization.
- Ensure strict neutrality with no competitor mentions.

### 2. DOCX Engine (`lib/docx-engine.ts`)
- Update `generateNotaryDocxBlob` to pass `pageCount` into `normalizeAllNotaryFieldValues`.

### 3. Unit Tests (`tests/notarial-core.test.ts`)
- Add tests covering:
  - 3-tier field value normalization.
  - `expandBlockLoop` with filtered active clauses.
  - Page count tafqeet (`pages2lettre_acte`).

## Verification Plan
- Run `npm test` to verify all unit tests pass.
- Run `compile_applet` to ensure zero compilation or type errors.
