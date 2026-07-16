# Input and Output Contract

## Purpose and boundary

Use this Skill only after the Mapping Skill has finalized one workbook containing a complete post-split source sheet and the fixed `Mapping结果` sheet. It changes the **record granularity** of the final research pool; it does not re-split indications, re-map diseases, change any Mapping status, or add a new medical judgment.

The desired business unit is normally `product or generic name × Disease CN`. Different Disease CN values remain separate research units.

## Required workbook structure

### Complete source sheet

The source sheet contains all business fields plus:

| Logical field | Role |
|---|---|
| Source key | Identifies the record before indication splitting; it may repeat after splitting. |
| `Mapping ID` | Unique technical ID for each split row. |
| Entity | Default product or business entity. |
| Split indication | One indication per source row. |
| `对应疾病` | Mapping backfill used for cross-sheet validation and final grouping. |
| Generic name, dosage form, brand name | Core identity fields for cross-registration consolidation. |

The source sheet does **not** need `Mapping Status` or `Rationale`. All remaining fields are preservation fields.

### Fixed Mapping-result sheet

`Mapping结果` contains exactly these 10 columns in order:

```text
序号
Mapping ID
药品
适应症拆分结果
标准化疾病实体
ICD-10代码
ICD-10疾病名称
Disease CN
Mapping Status
Rationale
```

`Mapping ID` is the sole join key. Both sheets must have the same nonblank unique ID set in the same order. Never join by Excel row number, source key, product name, or indication text.

For each joined row:

- `Mapped` requires nonblank `Disease CN` and source `对应疾病 = Disease CN`;
- every other status requires blank `Disease CN` and blank source `对应疾病`;
- `Rationale` is always nonblank.

## Eligibility and merge keys

Only rows whose joined `Mapping Status = Mapped` and whose joined `Disease CN` is nonblank may enter either merge layer.

### Layer 1 — merge mapped indications within a source record

```text
Source key + Entity + Disease CN
```

Keep the first source row as the anchor. Join source keys when necessary, `Mapping ID`, and split indications in original order with `；`; retain anchor-row values for other identical fields. Mapping rationales remain auditable by the retained IDs on the unchanged `Mapping结果` sheet.

### Layer 2 — consolidate equivalent cross-registration records

Run only after Layer 1. The core identity key is:

```text
Entity + Generic name + Dosage form + Brand name (including blankness)
+ merged split indication + Disease CN + Mapping Status
```

Use it only when different registrations represent the same downstream research object and differ solely in registration/history information such as holder, manufacturer, specification, approval number, packaging, or approval date.

Join multiple source keys, Mapping IDs, and divergent registration/history values with `；` in original order. If a core identity field differs, preserve separate rows and request a business decision.

## Validation and confirmation gate

Before generation, validate read-only:

1. Both sheets, header names, and complete field structures are readable.
2. `Mapping结果` uses the exact fixed 10-column schema.
3. Mapping IDs are unique, nonblank, and identical across sheets in set and order.
4. Every Mapping row has an allowed status and nonblank rationale.
5. Disease and source backfill obey the status-specific rules.
6. Candidate Layer 1 groups have consistent core static fields.
7. Candidate Layer 2 groups agree on the full core identity key; distinct rationale wording is shown for review but remains on `Mapping结果`.
8. Formula cells and possible external dependencies are reported for both sheets. They may remain on copied audit sheets, never on the final merged sheet.

Show all candidate groups and expected row counts. Do not generate until the user explicitly confirms the keys, cross-registration candidates, delimiter, treatment of non-Mapped rows, and output name.

## Output and quality requirements

Create a new workbook, never overwrite the input. Preserve every original sheet, including the complete source sheet and `Mapping结果`, unchanged. Add one final merged full-field sheet that:

- uses exactly the source-sheet columns and order;
- contains static displayed values only;
- keeps every non-Mapped row unchanged and in original relative order;
- retains the Mapping IDs required to trace every merged row back to `Mapping结果`.

Before handoff, verify:

- copied source and Mapping-result formulas equal the input sheets;
- output columns equal source columns;
- output rows equal input rows minus Layer 1 and Layer 2 reductions;
- Mapped output rows equal Mapped input rows minus both reductions;
- every non-Mapped status count is unchanged;
- merged IDs, indications, and registration differences retain original order and use `；`;
- the final merged sheet contains zero formulas.

## Stop instead of merging

Stop when either sheet is missing or incomplete; the Mapping-result schema changes; Mapping ID is blank, duplicated, missing from either sheet, or ordered differently; source backfill disagrees with Mapping results; a Mapped row lacks Disease CN; a non-Mapped row contains Disease CN; a first-layer group has a core-field conflict; a second-layer candidate differs in generic name, dosage form, indication, disease, status, or brand positioning; or formulas return invalid values.
