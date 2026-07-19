---
name: merge-mapped-disease-rows
description: Validate and consolidate a complete post-Mapping Excel pool without changing any medical decision. Use when Codex must join a complete source sheet to the fixed Mapping结果 sheet strictly by Mapping ID, structurally merge same-source Mapped rows with the same Disease, close every cross-source business-equivalence candidate as MERGE or KEEP SEPARATE, and generate a static full-field sheet while preserving all original audit sheets.
---

# Merge Mapped Disease Rows

Use this Skill after indication splitting, Mapping finalization, and Mapping-ID backfill, but before listed/clinical cross-pool consolidation. Keep this Skill limited to one source pool: do not derive Full/Clean views, join listed and clinical pools, or run ownership, MOA, MNC, VBP, or other downstream modules.

Before acting, read [references/input-output-contract.md](references/input-output-contract.md). Copy and complete [references/merge-config.example.json](references/merge-config.example.json) for each run. If Layer 2 candidates exist, use [references/layer2-decisions.example.json](references/layer2-decisions.example.json) to close them.

## Required workflow

1. **Freeze the input.** Confirm the complete post-Mapping workbook, source sheet, fixed `Mapping结果` sheet, source lineage key, field mapping, output sheet name, delimiter, Layer 2 identity fields, and allowed registration/history difference fields. Treat all row and column counts as run-specific.
2. **Preview read-only.** Run `Preview`. Report the input fingerprint, exact status counts, Mapping-ID coverage and order, backfill agreement, formula and external-link risks, Layer 1 groups and conflicts, Layer 2 candidates and difference fields, stable candidate IDs, potential reductions, and run fingerprint.
3. **Resolve Layer 1 blockers.** Layer 1 means the same source lineage key and the same Disease. Group by `Source key + Disease CN`; validate Entity and every source field except the approved split-varying fields. Never hide a conflict by putting Entity inside the group key.
4. **Close every Layer 2 candidate.** Layer 2 means different source records that may be the same downstream business asset. Record exactly one `MERGE` or `KEEP SEPARATE` decision with a substantive rationale for every candidate ID. A global confirmation is not a substitute for candidate-level closure.
5. **Approve execution.** Set the four workflow flags true only after the input, merge keys, Layer 2 decisions, and final execution are approved. If the input, config contract, or candidates change, rerun Preview and invalidate the old decisions.
6. **Generate once.** Run `Generate -Confirmed` with the bound run fingerprint and decision file. Merge only approved `MERGE` groups; retain `KEEP SEPARATE` groups separately. Never overwrite the input or an existing output.
7. **Reopen and verify.** Verify original sheets and fingerprints, source and Mapping formulas, exact output columns, row reductions, five-state counts, non-Mapped preservation, approved group execution, zero formulas in the static final sheet, and unchanged input SHA-256.

## Mapping contract

Accept exactly these five statuses with ASCII hyphens:

- `Mapped`
- `Manual Review Required - No TA Match`
- `Manual Review Required - Multiple Candidates`
- `Unmapped - Other TA`
- `Unmapped - Invalid Information`

Only `Mapped` rows with a nonblank whitelist Disease CN may enter either layer. Preserve every other status as one row in original relative order, with blank `Disease CN` and blank source `对应疾病`. Never normalize, reinterpret, or close a Mapping decision here.

## Merge rules

- Join the source and `Mapping结果` only by unique `Mapping ID`; require identical ID set and order.
- Layer 1 key is `Source key + Disease CN`. Join Mapping IDs and split indications in source order with the approved delimiter. All other fields must agree unless explicitly listed as Layer 1 split-varying fields.
- Layer 2 candidates are built only after Layer 1 from the approved identity fields. Require Entity, generic name, dosage form, brand name including blankness, merged indication, and Disease among those fields; allow additional run-specific identity fields.
- A Layer 2 `MERGE` is legal only when every divergent field is the source key, Mapping ID, or an approved registration/history difference field. A prohibited difference cannot be overridden by the decision file.
- For an approved Layer 2 merge, join only the source key, Mapping ID, and approved divergent fields. Preserve identical fields from the anchor row and retain source order.
- Keep all original worksheets. If the requested final sheet already exists, stop; do not delete or replace it.
- Create the final sheet from static evaluated values while preserving practical value types and anchor-row formats. Do not convert the entire output to text.

## Commands

```powershell
& .\scripts\merge_mapped_disease_rows.ps1 `
  -Mode Preview `
  -ConfigPath "C:\path\merge-config.json"
```

After every Layer 2 candidate is closed and the workflow flags are approved:

```powershell
& .\scripts\merge_mapped_disease_rows.ps1 `
  -Mode Generate `
  -ConfigPath "C:\path\merge-config.json" `
  -DecisionPath "C:\path\layer2-decisions.json" `
  -OutputPath "C:\path\mapped-disease-merged.xlsx" `
  -Confirmed
```

`DecisionPath` may be omitted only when Preview reports zero Layer 2 candidates.

## Hard stops

Stop when a required sheet or field is missing; Mapping status is outside the five-state contract; IDs are blank, duplicated, missing, or differently ordered; Mapping and source backfill disagree; Rationale is blank; a Mapped row lacks Disease; a non-Mapped row contains Disease; Layer 1 contains an unapproved field conflict; a Layer 2 `MERGE` contains a prohibited difference; a candidate is missing, duplicated, unknown, or still pending; the decision fingerprint differs; formulas contain errors or unsafe external dependencies; the final sheet already exists; the output extension differs; any original sheet changes; or approval is incomplete.

## Completion report

Report the input and output paths and fingerprints; input and output row/field counts; Layer 1 groups and reduction; Layer 2 MERGE and KEEP SEPARATE groups and reduction; five-state counts before and after; non-Mapped preservation; formula and external-link results; original-sheet preservation; final formula count; and exceptions.
