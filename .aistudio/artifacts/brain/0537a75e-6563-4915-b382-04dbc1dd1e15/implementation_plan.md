# Implementation Plan - Final Cleanup & Template Canonicalization

## Objectives
1. **Unify Templates 2–6 to Canonical Symbols**: Replace legacy extended Latin keys (`partie1_nom_complet`, `notaire_nom`, etc.) across the 5 remaining built-in templates in `lib/notary-word-templates.ts` with canonical notary symbols (`${c_1}`, `${c_2}`, `${office_name}`, `${office_addr}`, `${num_acte}`, `${date_acte}`, `${act_value}`, etc.).
2. **Purge Legacy Pagination in UI**: Remove `injectNotarialPaginationIntoHtml` calls and imports from `components/NotaryEditorApp.tsx` and related components, ensuring the app preview uses clean CSS/digital footers without in-text paragraph injection.
3. **Remove Dead Legacy Code**: Delete `spliceContractBodyIntoTemplate`, `mergePlaceholdersIntoHtml`, and obsolete references to `[contenu_acte]` from `lib/notary-word-templates.ts` and `lib/docx-engine.ts`.
4. **Verification**: Run `npm test` and `npm run build` to confirm all 18+ tests and production build pass successfully.

## Proposed Changes

### 1. `lib/notary-word-templates.ts`
- Update templates 2 through 6 (Contrat Référence, Copie, Extrait, Certificat de Vente, Tableau d'Envoi) to use canonical symbols:
  - Parties header using `${c_1}` and `${c_2}` within `parties-wrapper`.
  - Notary info using `${office_name}`, `${office_addr}`, `${office_tel}`.
  - Act details using `${num_acte}`, `${date_acte}`, `${act_value}`.
  - Body using `${block}`, `${titre}`, `${clause}`, `${/block}`.
- Remove dead legacy functions `spliceContractBodyIntoTemplate`, `mergePlaceholdersIntoHtml`, and old bracket tag patterns.
- Refine `supportedTags` to list only canonical symbols.

### 2. `components/NotaryEditorApp.tsx` (and editor UI)
- Remove `injectNotarialPaginationIntoHtml` imports and usage from the app preview/print pipeline.
- Keep the digital footer preview in `EditorA4Canvas` and document header/footer generation clean.

### 3. Tests & Build
- Update any test assertions in `tests/notarial-core.test.ts` that reference old template tags or legacy pagination functions.
- Run `npm test` and `npm run build`.

## Verification Plan
- Automated unit tests: `npm test`
- Production build: `npm run build`
