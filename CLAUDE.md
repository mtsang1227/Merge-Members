# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Merge two Excel files into a single `Result.xlsx`, using the first input file's column headers as the canonical schema.

## Running the Script

```bash
node merge.js <file1.xlsx> <file2.xlsx> [output.xlsx]
```

Example:
```bash
node merge.js F1.XLSX F2.XLSX
node merge.js F1.XLSX F2.XLSX MyOutput.xlsx
```

Both input file arguments are required. The third argument (output file) is optional and defaults to `Result.xlsx`. If the required arguments are omitted, the script prints a usage hint and exits, along with a note that both input files MUST have the proper column names and that the second file does not need to have the same column ordering.

Requires the `xlsx` package (already installed locally):

```bash
npm install xlsx
```

## Output

After a successful run the script prints a summary:

```
=== Merge Summary ===
Input file 1 (F1.XLSX): 3 records
Input file 2 (F2.XLSX): 3 records
Records merged into Result.xlsx: 6
```

## File Overview

- `F1.XLSX` / `F2.XLSX` — Source data files. Both have columns: `ID, Lastname, Firstname, Phone, Address`
- `merge.js` — Merges two input files (passed as arguments) into `Result.xlsx`
- `Result.xlsx` — Output file (file1 rows first, then file2 rows)

## Key Logic in merge.js

- The first input file's headers define the column order and names in the output
- The second file's columns are matched to the first file's headers case-insensitively, so minor casing differences are handled automatically
- All rows from file1 appear before rows from file2 in the result
