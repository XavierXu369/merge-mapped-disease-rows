# Input and Output Contract

## Purpose and boundary

Use this Skill only after a complete pool has already been split to one indication per row and mapped to a disease. It changes the **record granularity** of the final research pool; it does not re-split, re-map, normalize disease entities, or make new medical judgments.

The desired business unit is normally `product or generic name x Disease CN`. Different Disease CN values must remain separate because they are separate research and charging units.

## Required inputs

Provide one Excel workbook, one complete post-mapping source sheet, exact column-header names, and an output path. The required logical fields are:

| Logical field | Role |
|---|---|
| Source key | Identifies the record before indication splitting; often a repeated serial number. |
| Mapping ID | Unique technical ID for each split row; it must be nonblank and unique before merging. |
| Entity | Default main business entity, usually product name. |
| Split indication | One indication per input row. |
| Disease CN | Mapped target disease. |
| Mapping Status | Exact mapping conclusion. |
| Rationale | Audit trail for the mapping conclusion. |
| Generic name, dosage form, brand name | Core identity fields for cross-registration consolidation. |

All remaining source fields are mandatory preservation fields. Common examples are holder, manufacturer, location, group, specification, approval number, dates, VBP flag, target, and original approved indication.

## Eligibility and merge keys

Only rows with `Mapping Status = Mapped` (or the user-confirmed mapped value) and a nonblank Disease CN may enter either layer.

### Layer 1 — merge mapped indications within a source record

```text
Source key + Entity + Disease CN
```

Use this layer when one original product record has several split indications that map to the same disease. Keep the first source row as the anchor. Join Mapping IDs, split indications, and rationales in original order with `；`; retain anchor-row values for all other fields.

### Layer 2 — consolidate equivalent cross-registration records

Run only after Layer 1. The core identity key is:

```text
Entity + Generic name + Dosage form + Brand name (including blankness)
+ merged split indication + Disease CN + Mapping Status
```

Use it only when different records represent the same downstream research object and differ solely in registration/history information, such as holder, manufacturer, specification, approval number, packaging, or approval date.

Join multiple source keys, Mapping IDs, and divergent registration/history values with `；` in original order. Do not repeat an identical indication or rationale. If a core identity field differs, preserve separate rows and request a business decision.

## Validation and confirmation gate

Before generation, validate read-only:

1. Workbook, sheet, header names, and complete field structure are readable.
2. Mapping ID is unique and nonblank.
3. Every Mapped row has a nonblank Disease CN.
4. Candidate Layer 1 groups have consistent core static fields.
5. Candidate Layer 2 groups agree on the full core identity key; review any distinct rationale wording for a possible medical-conclusion conflict.
6. Formula cells and potential external workbook dependencies are reported. They are allowed only in the copied source sheet, never in the final merged sheet.

Show the candidate groups and expected row counts. Do not generate until the user explicitly confirms the keys, candidate groups, delimiter, treatment of non-Mapped rows, and output name.

## Output and quality requirements

Create a new workbook, never overwrite the input. It must contain:

1. The copied source sheet, unchanged, including its original formulas if present.
2. A merged full-field sheet with the same columns and source order, but with static displayed values only.

Before handoff, verify:

- Original source sheet equals the input sheet in dimensions and formulas.
- Output columns equal source columns.
- Output rows equal input rows minus Layer 1 and Layer 2 reductions.
- Mapped output rows equal Mapped input rows minus both reductions.
- Every non-Mapped status count is unchanged.
- All merged IDs, indications, rationales, and registration differences retain original order and use `；`.
- The merged sheet contains zero formulas.

## Stop instead of merging

Stop and request human direction if Mapping ID is blank, duplicate, or already concatenated; a Mapped row lacks Disease CN; a first-layer group has a core-field conflict; a second-layer candidate differs in generic name, dosage form, indication, disease, status, or brand positioning; the source sheet is incomplete; or formulas return errors/invalid values.
