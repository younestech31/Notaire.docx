# Implementation Plan - Final Verification of Notary Clauses & Party Cards

## Goal
Verify and finalize the closed container rule, strict `#` clause recognition, ignoring stale `data-clause-id` stamps without `#`, and precise party card variable mapping (`الطرف_الأول_الاسم`, `الطرف_الثاني_الاسم`, etc.) in `Notaire DZ Office`.

## Proposed Changes

### 1. Verification of Clause Outline & Recognition (`lib/clause-outline-engine.ts`)
- Confirm that `div.clause-container` acts as a closed container.
- Confirm that paragraphs starting with `#` swallow siblings until the next `#` and preserve existing IDs.
- Confirm that any other element with `data-clause-id` lacking `#` is ignored as a clause.
- Confirm no new IDs are generated during read operations.

### 2. Verification of Party Card Variable Mapping
- Verify `applyPartyCardsToFieldValues` and `buildPartyCardsVariableMap` correctly populate variables with full prefixes (`الطرف_الأول_الاسم`, `الطرف_الأول_تاريخ_الميلاد`, `الطرف_الثاني_الاسم`, etc.).

### 3. Build & Lint Verification
- Run `compile_applet` and `lint_applet` to ensure zero compilation errors.

## Verification Plan
- Compile and test applet compilation.
- Ensure all tests pass.
