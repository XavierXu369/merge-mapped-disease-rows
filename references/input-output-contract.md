# Merge Mapped Disease Rows V2 — Input and Output Contract

## Contents

1. Boundary
2. Run inputs
3. Mapping contract
4. Layer 1 contract
5. Layer 2 contract
6. Candidate decisions and fingerprint
7. Preview contract
8. Generate contract
9. Output preservation
10. Quality control and stops

## 1. Boundary

Run this Skill on one complete source pool after indication splitting, Mapping finalization, and Mapping-ID backfill. It changes only record granularity. It must not split indications again, revise medical Mapping, consolidate listed and clinical pools, derive Full/Clean views, or calculate downstream module markers.

The usual downstream research unit is close to `business asset × dosage form × Disease`, but every run must freeze its own Layer 2 identity fields. Field counts, row counts, candidate counts, and TA-specific decisions are never permanent constants.

## 2. Run inputs

### 2.1 Complete source sheet

The source sheet contains every original business field plus:

| Logical field | Requirement |
|---|---|
| Source key | Nonblank lineage ID for the record before indication splitting. It may repeat after splitting. |
| `Mapping ID` | Nonblank, unique technical ID for every split row. |
| Entity | Nonblank product or business entity used for identity validation. |
| Split indication | One split indication per Mapping row. |
| `对应疾病` | Backfilled Disease for Mapped rows only. |
| Generic name, dosage form, brand name | Required Layer 2 identity inputs; values may be blank only when the business field legitimately has no value. |

The configured header row and effective data rows are determined by nonblank Mapping IDs, not by formatting-only `UsedRange` rows.

### 2.2 Fixed Mapping-result sheet

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

`Mapping ID` is the only join key. The source and Mapping sheets must have identical nonblank unique IDs in identical order.

## 3. Mapping contract

Allowed statuses are exactly:

| Status | Disease CN and source backfill | Merge eligibility |
|---|---|---|
| `Mapped` | Both nonblank and equal | Eligible |
| `Manual Review Required - No TA Match` | Both blank | Never merge |
| `Manual Review Required - Multiple Candidates` | Both blank | Never merge |
| `Unmapped - Other TA` | Both blank | Never merge |
| `Unmapped - Invalid Information` | Both blank | Never merge |

Every Mapping row requires a nonblank Rationale. This Skill preserves all Mapping values and formatting, including the red No TA Match Rationale rule. It does not convert deprecated status wording or make a new medical decision.

## 4. Layer 1 contract

Layer 1 reconstructs split indications that belong to the same original source record and now resolve to the same Disease.

```text
Layer 1 key = Source key + Disease CN
```

Do not include Entity in the key. Entity is a validation field: if one source lineage key contains conflicting Entity values, the conflict must be visible rather than silently split into separate groups.

For a multirow Layer 1 group:

- join `Mapping ID` and split indication in source order with the configured delimiter;
- retain Source key and Disease once;
- require all other source fields to have one normalized value unless the header appears in `layer1_allowed_varying_fields`;
- join explicitly approved split-varying fields in source order;
- stop before generation when any other field differs.

Layer 1 is structural and may be executed without a per-group business decision only after all consistency checks pass.

## 5. Layer 2 contract

Layer 2 compares different Layer 1 source records. Its identity fields are run configuration, but must include at least:

- Entity;
- generic name;
- dosage form;
- brand name, with blank and nonblank treated as different;
- merged split indication;
- Disease.

Additional fields such as ingredient combination, route, release form, salt form, target, or other asset-defining fields may be added for a run. Source key and Mapping ID may not be identity fields.

Every candidate reports all differing source columns. The following differences are automatically aggregatable:

- Source key;
- `Mapping ID`.

Any other difference must be named in `layer2_allowed_difference_fields`, which should contain only approved registration/history fields such as holder, manufacturer, specification, approval number, packaging, or approval date. A `MERGE` decision cannot override a prohibited difference.

Different Disease values never share a candidate. Non-Mapped records never enter Layer 2.

## 6. Candidate decisions and fingerprint

Preview creates:

- a SHA-256 input fingerprint;
- a contract hash that excludes mutable workflow approval flags;
- stable Layer 1 and Layer 2 group IDs;
- a run fingerprint covering the input, contract, groups, and differences.

When Layer 2 candidates exist, the decision file must contain the exact run fingerprint and exactly one entry per candidate:

- `MERGE`: consolidate the candidate;
- `KEEP SEPARATE`: retain each child record separately.

Each decision requires a substantive rationale. Unknown, duplicated, missing, or pending candidates block generation. A changed workbook, field contract, allowed difference list, or candidate set produces a different fingerprint and invalidates the old decision file.

## 7. Preview contract

Preview is read-only and returns at least:

- source and Mapping sheet dimensions and effective rows;
- input SHA-256 and contract hash;
- exact five-state counts;
- Mapping-ID set/order and backfill validation;
- source and Mapping formula counts, formula-error counts, and external-formula counts;
- Layer 1 groups, joined IDs/indications, differing fields, blockers, and reduction;
- Layer 2 candidate IDs, child source keys, Mapping IDs, identity values, differing fields and values, prohibited differences, merge eligibility, and potential reduction;
- final-sheet name availability;
- run fingerprint and readiness.

Preview may be rerun without changing any workbook.

## 8. Generate contract

Generate requires:

- `workflow.input_approved = true`;
- `workflow.merge_keys_approved = true`;
- `workflow.second_layer_decisions_closed = true`;
- `workflow.final_execution_approved = true`;
- CLI `-Confirmed`;
- a matching, complete decision file when Layer 2 candidates exist.

Apply every valid Layer 1 merge. Apply only Layer 2 candidates decided `MERGE`; retain every `KEEP SEPARATE` child. Sort output records by their earliest source row.

## 9. Output preservation

Create a new workbook with the same extension as the input. Preserve every original worksheet; do not delete or replace a pre-existing final-sheet name.

The new final sheet must:

- use exactly the source headers and order;
- contain static evaluated values and zero formulas;
- retain source value types where practical instead of coercing the entire matrix to text;
- copy anchor-row formats and number formats;
- join only approved aggregate fields with the configured delimiter;
- preserve the Mapping IDs needed to trace each merged record to the unchanged `Mapping结果` sheet.

The input file must remain byte-identical. Reopen the output and verify original sheet names plus value/formula fingerprints for all preserved sheets.

## 10. Quality control and stops

Stop when:

- a required workbook, sheet, header, or effective row is missing;
- the fixed Mapping schema or five-state vocabulary is violated;
- Mapping IDs are blank, duplicated, missing, or differently ordered;
- source backfill and Mapping Disease disagree;
- Rationale is blank or Disease rules are violated;
- mapped source lineage key, Entity, or split indication is blank;
- Layer 1 contains an unapproved varying field;
- a Layer 2 candidate is unclosed or a `MERGE` contains a prohibited difference;
- the decision fingerprint differs;
- formulas contain error values or external formulas are present without explicit approval;
- the final sheet already exists;
- the output exists, overwrites the input, or changes extension;
- an original worksheet changes;
- the final row arithmetic, status counts, merged IDs, or zero-formula check fails.
