---
name: merge-mapped-disease-rows
description: Validate and consolidate a complete Excel mapping workbook after indication-to-disease mapping. Use when Codex must join a complete source sheet to the fixed Mapping结果 sheet strictly by Mapping ID, preview and, only after approval, merge Mapped rows with the same disease within a source record, and deduplicate business-equivalent cross-registration records while preserving both audit sheets and a static full-field output sheet.
---

# Merge Mapped Disease Rows

Create a final analysis pool after indication splitting and disease mapping. Do not split indications again or change any medical mapping conclusion.

Before acting, read [references/input-output-contract.md](references/input-output-contract.md).

## Required interaction

Follow these gates in order. Do not skip the confirmation gate.

1. **Request the input.** Ask for one complete post-mapping Excel workbook, its complete source-sheet name, the fixed `Mapping结果` sheet, the relevant source column names, and an output path. The source sheet must include every original business field plus `Mapping ID` and `对应疾病`; it does not need `Mapping Status` or `Rationale`.
2. **Validate read-only and preview.** Run the script in `Preview` mode. Report data rows and columns; exact 10-column Mapping-result schema; Mapping ID uniqueness, coverage, and order across both sheets; source backfill agreement; Mapping Status counts; Mapped rows lacking Disease CN; formula/external-link risk on both sheets; first-layer candidate groups; second-layer cross-registration candidate groups; conflicts; and expected row changes.
3. **Pause for approval.** Show all candidate groups when the number is manageable. Obtain explicit confirmation of the two merge keys, every cross-registration candidate, preservation of non-Mapped records, delimiter, and output name. If any conflict or business-identity uncertainty remains, stop.
4. **Generate only after approval.** Run the script in `Generate` mode with `-Confirmed`. Reopen and verify the result.
5. **Report completion.** Give the output path, layer-by-layer reductions, final status counts, confirmation that the source sheet is unchanged, and any exceptions.

## Core rules

- Join the source sheet and `Mapping结果` only by the unique `Mapping ID`; never use row position, source sequence, product name, or indication text.
- Read `Disease CN`, `Mapping Status`, and `Rationale` from `Mapping结果`. Use source `对应疾病` only to validate the Mapping backfill.
- Merge only rows whose joined status exactly equals the confirmed mapped value (normally `Mapped`) and whose joined Disease CN is nonblank.
- First layer: merge within the same source record using `SourceKey + Entity + Disease CN`.
- Second layer: merge only business-equivalent records using `Entity + Generic + Dosage form + Brand (including blankness) + merged indication + Disease CN + Mapping Status`.
- Keep distinct diseases as separate rows. Keep every non-Mapped row unchanged and in original relative order.
- Join retained IDs, indications, and different registration values in original order with the full-width semicolon `；`. Keep row-level rationales on the unchanged `Mapping结果` sheet instead of adding a new final-table column.
- Preserve the copied source and `Mapping结果` sheets unchanged. The merged sheet must contain the source columns only, as static values, never formulas that depend on an external workbook.

## Run the script

Run preview first. Supply the actual column headers used by the workbook.

```powershell
& .\scripts\merge_mapped_disease_rows.ps1 `
  -Mode Preview `
  -InputPath "C:\path\mapped-pool.xlsx" `
  -SourceSheet "<source-sheet>" `
  -MappingResultSheet "Mapping结果" `
  -SourceKeyColumn "<source-key>" `
  -MappingIdColumn "Mapping ID" `
  -EntityColumn "<product-or-entity>" `
  -IndicationColumn "<split-indication>" `
  -DiseaseColumn "对应疾病" `
  -GenericColumn "<generic-name>" `
  -DosageFormColumn "<dosage-form>" `
  -BrandColumn "<brand-name>"
```

After explicit approval, rerun with `-Mode Generate -Confirmed` and add `-OutputPath`. The script refuses to overwrite an existing file.

## Hard stops

Stop and ask for direction when either sheet is missing; the Mapping-result schema is not the fixed 10 columns; Mapping ID is blank, duplicated, absent from either sheet, or ordered differently; `对应疾病` disagrees with the joined result; a Mapped row has no disease; a non-Mapped row contains Disease CN; source-record core fields conflict; a proposed cross-registration group changes product, generic name, dosage form, indication, disease, status, or brand positioning; a formula cannot return a valid value; or approval has not been provided.

## Completion report

Report input and output row/field counts; Mapping ID join verification; first- and second-layer merge groups and reductions; Mapping Status counts before and after; non-Mapped preservation; formula checks; source- and Mapping-result-sheet preservation; exceptions; and a clickable output path.
