# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Merge two Excel files into a single output file, using the first input file's column headers as the canonical schema. Records from the second file are tagged with a fellowship name under the `Fellowship` column.

There is also a comparison script (`compare.js`) that checks a merged fellowship file against the master `DynamicGroupReport.xlsx` export to find records that are new, unchanged, or different from the master copy.

## Running the Script

```bash
node merge.js <file1.xlsx> <file2.xlsx> <fellowship_name> [output.xlsx]
```

Example:
```bash
node merge.js "Members files sample/2026 LTG Fellowship Members List.xlsx" "Members files sample/Andrew Fellowship Members' Contact.xlsx" "Andrew"
node merge.js "Members files sample/2026 LTG Fellowship Members List.xlsx" "Members files sample/Andrew Fellowship Members' Contact.xlsx" "Andrew" "2026-Merged-List.xlsx"
node merge.js "Members files sample/2026-Merged-List.xlsx" "Members files sample/2025 (Daniel) Fellowship Directory.xlsx" "Daniel" "Members files sample/2026-Merged-Andrew-Daniel.xlsx"
```

The first three arguments are required. The fourth argument (output file) is optional and defaults to `Result.xlsx`. If required arguments are omitted, the script prints a usage hint and exits.

### compare.js

```bash
node compare.js <masterFile.xlsx> <mergedFile.xlsx> [noChangeOut.xlsx] [yesChangeOut.xlsx]
```

Example:
```bash
node compare.js "DynamicGroupReport.xlsx" "Merged-Fellowships.xlsx"
```

Compares every record in `mergedFile` against `masterFile` (the DynamicGroupReport-style export, source of truth for `Id`) and writes three output files:
- `noChangeOut.xlsx` (default `noChange.xlsx`) — merged records that match a master record with no differences, tagged with the master `Id`.
- `yesChangeOut.xlsx` (default `yesChange.xlsx`) — merged records that match a master record but differ in one or more fields, tagged with the master `Id` and a `Changed Fields` column listing what changed.
- `notFound.xlsx` — merged records that couldn't be resolved to exactly one master record, tagged with a `Reason` column (`No master record found` or `Multiple master candidates even after Chinese tiebreak`). Not requested by the original spec, but included so unmatched records aren't silently dropped.

Both input files are required; if omitted, the script prints a usage hint and exits.

Requires the `xlsx` package (already installed locally):

```bash
npm install xlsx
```

## Output

After a successful run the script prints a summary:

```
=== Merge Summary ===
Input file 1 (2026 LTG Fellowship Members List.xlsx): 518 records
Input file 2 (Andrew Fellowship Members' Contact.xlsx): 44 records
Records merged into 2026-Merged-List.xlsx: 562
```

`compare.js` prints a similar summary:

```
=== Compare Summary ===
Master file (DynamicGroupReport.xlsx): 2447 records
Merged file (Merged-Fellowships.xlsx): 1742 records
No change (noChange.xlsx): 189
Changed (yesChange.xlsx): 905
Not found (notFound.xlsx): 648 (648 no master record, 0 still ambiguous after Chinese tiebreak)
```

## File Overview

- `merge.js` — Main script; merges two input files into a single output file
- `compare.js` — Compares a merged fellowship file against the master DynamicGroupReport export to find new/unchanged/changed records
- `Members files sample/` — Sample input files used for testing
- `Fellowship files/` — Source fellowship directory files
- `Result.xlsx` — Default output file when no output name is specified

## Key Logic in merge.js

- Only the first sheet in each .XLSX file will be merged into the output file
- The first XLSX file should have the column "Fellowship" with its proper value. You may need to add it before doing the first merge action to get the proper result
- The first input file's headers define the column order and names in the output
- Header row auto-detection: within the first 5 rows, the row with the most non-empty cells is treated as the header row — description rows above it are skipped automatically
- The second file's columns are matched to the first file's headers case-insensitively, so minor casing differences are handled automatically
- All rows from file1 appear before rows from file2 in the result
- Each file2 record is tagged with the provided `fellowship_name` under the `Fellowship` column
- It is strongly recommended to have "fellowship" column and value in the file1; if `Fellowship` does not exist in file1's headers it is appended as the last column with empty value.
- Output handling: if the output path resolves to the same file as file1, no new workbook is built — file2's records are appended directly onto file1's existing sheet in place (preserving file1's original rows, formatting, and any other sheets). Otherwise, a brand new output workbook is created from the merged data.

## Key Logic in compare.js

- Matching a merged record to a master record works in stages, from most to least specific:
  1. Scope candidates to the same Fellowship — the merged file's `Fellowship` value must equal the *last word* of the master file's `Fellowship` string (e.g. merged `"Esther"` matches master `"Bayview Cantonese - 2023 Esther"`) — and the same `Last Name`.
  2. Match `First Name`/`Middle Name` tolerantly: either the plain First Name matches on both sides, or an order-independent token set of First+Middle matches. This handles master sometimes storing the full name in `First Name` alone (Middle Name blank), name-order swaps (e.g. `"Wai King Peggy"` vs `"Peggy"` + `"Wai King"`), and hyphen-vs-space compound name variants (e.g. `"Shuk-Yin"` vs `"Shuk Yin"`).
  3. If more than one master candidate remains, narrow by `Chinese Name` (master) vs `Chinese` (merged).
- SG Name matching is abbreviation-tolerant (one side is a prefix of the other, e.g. merged `"Faith"` matches master `"Faithfulness"`) and is compared against master's `SG (Compute)` column, falling back to `Small Group` when `SG (Compute)` is blank (master leaves `SG (Compute)` empty whenever `Small Group` holds a person's full name rather than a single-word role/group name).
- Once a record is matched to exactly one master record, these fields are compared to decide "changed" vs "no change": Middle Name, Chinese/Chinese Name, Home Phone, Mobile Phone/Mobile Number (compared as digits only, ignoring formatting like `(647) 778-3213` vs `647-778-3213`), Email (case-insensitive), SG Leader/SG Leader (Compute), and SG Name (using the same abbreviation-tolerant comparison as the matching step, so an abbreviation difference alone isn't flagged as a change).
- Records with zero master candidates, or more than one candidate even after the Chinese tiebreak, are written to `notFound.xlsx` rather than silently dropped — this surfaces both genuinely new people (not yet in the master file) and ambiguous/unresolvable cases for manual review.
- Known limitation: nickname mismatches (e.g. merged `"Tony"` vs master `"Anthony"`) and incomplete data (e.g. merged has only `"Paul"` where master's combined name is `"Paul Ching Hung"`) cannot be resolved by string matching alone and will show up in `notFound.xlsx`.
