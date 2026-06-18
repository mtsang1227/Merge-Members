# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Merge two Excel files into a single output file, using the first input file's column headers as the canonical schema. Records from the second file are tagged with a fellowship name under the `Fellowship` column.

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

## File Overview

- `merge.js` — Main script; merges two input files into a single output file
- `Members files sample/` — Sample input files used for testing
- `Fellowship files/` — Source fellowship directory files
- `Result.xlsx` — Default output file when no output name is specified

## Key Logic in merge.js

- The first input file's headers define the column order and names in the output
- Header row auto-detection: within the first 5 rows, the row with the most non-empty cells is treated as the header row — description rows above it are skipped automatically
- The second file's columns are matched to the first file's headers case-insensitively, so minor casing differences are handled automatically
- All rows from file1 appear before rows from file2 in the result
- Each file2 record is tagged with the provided `fellowship_name` under the `Fellowship` column
- It is strongly recommended to have "fellowship" column and value in the file1; if `Fellowship` does not exist in file1's headers it is appended as the last column with empty value.
