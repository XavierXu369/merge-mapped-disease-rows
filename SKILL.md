---
name: merge-mapped-disease-rows
description: Validate and consolidate a complete Excel mapping workbook after indication-to-disease mapping. Use when Codex must preview and, only after approval, merge Mapped rows with the same disease within a source record and deduplicate business-equivalent cross-registration records while preserving an unchanged source sheet and a static full-field output sheet.
---

# Merge Mapped Disease Rows

Create a final analysis pool after indication splitting and disease mapping. Do not split indications again or change any medical mapping conclusion.

Before acting, read [references/input-output-contract.md](references/input-output-contract.md).

## Required interaction

Follow these gates in order. Do not skip the confirmation gate.

1. **Request the input.** Ask for one complete post-mapping Excel workbook, its source-sheet name, the relevant column names, and an output path. The sheet must include every original business field, not only mapping columns.
2. **Validate read-only and preview.** Run the script in `Preview` mode. Report data rows and columns; Mapping ID uniqueness; Mapping Status counts; Mapped rows lacking Disease CN; formula/external-link risk; first-layer candidate groups; second-layer cross-registration candidate groups; conflicts; and expected row changes.
3. **Pause for approval.** Show all candidate groups when the number is manageable. Obtain explicit confirmation of the two merge keys, every cross-registration candidate, preservation of non-Mapped records, delimiter, and output name. If any conflict or business-identity uncertainty remains, stop.
4. **Generate only after approval.** Run the script in `Generate` mode with `-Confirmed`. Reopen and verify the result.
5. **Report completion.** Give the output path, layer-by-layer reductions, final status counts, confirmation that the source sheet is unchanged, and any exceptions.

## Core rules

- Merge only rows whose status exactly equals the confirmed mapped value (normally `Mapped`) and whose Disease CN is nonblank.
- First layer: merge within the same source record using `SourceKey + Entity + Disease CN`.
- Second layer: merge only business-equivalent records using `Entity + Generic + Dosage form + Brand (including blankness) + merged indication + Disease CN + Mapping Status`.
- Keep distinct diseases as separate rows. Keep every non-Mapped row unchanged and in original relative order.
- Join retained IDs, indications, rationales, and different registration values in original order with the full-width semicolon `；`.
- Preserve the copied source sheet unchanged. The merged sheet must contain static values only, never formulas that depend on an external workbook.

## Run the script

Run preview first. Supply the actual column headers used by the workbook.

```powershell
& .\scripts\merge_mapped_disease_rows.ps1 `
  -Mode Preview `
  -InputPath "C:\path\mapped-pool.xlsx" `
  -SourceSheet "<source-sheet>" `
  -SourceKeyColumn "<source-key>" `
  -MappingIdColumn "<mapping-id>" `
  -EntityColumn "<product-or-entity>" `
  -IndicationColumn "<split-indication>" `
  -DiseaseColumn "<disease-cn>" `
  -StatusColumn "<mapping-status>" `
  -RationaleColumn "<rationale>" `
  -GenericColumn "<generic-name>" `
  -DosageFormColumn "<dosage-form>" `
  -BrandColumn "<brand-name>"
```

After explicit approval, rerun with `-Mode Generate -Confirmed` and add `-OutputPath`. The script refuses to overwrite an existing file.

## Hard stops

Stop and ask for direction when a required column is missing; Mapping ID is blank or duplicated; a Mapped row has no disease; source-record core fields conflict; a proposed cross-registration group changes product, generic name, dosage form, indication, disease, status, or brand positioning; a formula cannot return a valid value; or approval has not been provided.

## Completion report

Report input and output row/field counts; first- and second-layer merge groups and reductions; Mapping Status counts before and after; non-Mapped preservation; formula checks; source-sheet preservation; exceptions; and a clickable output path.
