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
- `noChangeOut.xlsx` (default `noChange.xlsx`) — merged records that match a master record with no differences, tagged with the master `Id` and a `Match Method` column.
- `yesChangeOut.xlsx` (default `yesChange.xlsx`) — merged records that match a master record but differ in one or more fields, tagged with the master `Id`, a `Match Method` column, and a `Changed Fields` column listing what changed.
- `notFound.xlsx` — merged records that couldn't be resolved to exactly one master record, tagged with a `Reason` column (`No master record found`, `Multiple master candidates even after Chinese tiebreak`, `Multiple master candidates matched by Mobile Phone/Email`, or `Multiple master candidates across different Fellowships (same name)`). Not requested by the original spec, but included so unmatched records aren't silently dropped.

`Match Method` is `Name` for records matched by name within the same Fellowship (see Key Logic below), `Contact info (name mismatch)` when the name text didn't line up but Mobile Phone or Email uniquely identified the same person within the same Last Name + Fellowship bucket, or `Name (different Fellowship)` when the person was found under a different Fellowship than the merged file lists (people can legitimately belong to more than one Fellowship, or master may just be stale) — `Fellowship` itself then shows up in `Changed Fields`. The two fallback methods are worth reviewing with extra scrutiny: contact info can occasionally be shared between family members, and a same-name match across Fellowships could coincidentally be a different person entirely if Chinese Name isn't on file to disambiguate.

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
No change (noChange.xlsx): 215
Changed (yesChange.xlsx): 1259 (30 via name (different fellowship), 82 via contact info (name mismatch), 74 via chinese name + contact info (different last name), 16 via contact info (2 of 3 signals, no last name/chinese name match), 41 via contact info (different fellowship, name mismatch))
Not found (notFound.xlsx): 268 (all "no master record found" - none ambiguous)
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

- Matching a merged record to a master record works in stages, from most to least specific. All stages except the last two are scoped to the same Fellowship (the merged file's `Fellowship` value must equal the *last word* of master's `Fellowship` string, e.g. merged `"Esther"` matches master `"Bayview Cantonese - 2023 Esther"`) and the same `Last Name` (compared with hyphens normalized to spaces, e.g. `"Chan-Lam"` matches `"Chan Lam"`, the same tolerance already applied to First/Middle names):
  1. **Chinese Name first, if it uniquely identifies someone.** If merged's `Chinese` is non-empty and exactly one master record in the bucket has that exact `Chinese Name`, use it directly — this runs *before* the name match below, because Chinese Name is a stronger, less collision-prone signal than a bare English first name. Skipping this step let two different merged "Amy Fung" rows (different Chinese names, emails, phones) both get matched to the same wrong master record, since a bare `"Amy"` name-match found exactly one (wrong) candidate before Chinese Name was ever consulted.
  2. Otherwise, match `First Name`/`Middle Name` tolerantly: either the plain First Name matches on both sides, or an order-independent token set of First+Middle matches. This handles master sometimes storing the full name in `First Name` alone (Middle Name blank), name-order swaps (e.g. `"Wai King Peggy"` vs `"Peggy"` + `"Wai King"`), and hyphen-vs-space compound name variants (e.g. `"Shuk-Yin"` vs `"Shuk Yin"`). If more than one candidate remains, narrow by Chinese Name.
  3. If step 2 finds zero candidates (name text doesn't line up at all — incomplete data, unrecognized nickname), fall back to Mobile Phone or Email matching a master record within the same Last Name + Fellowship bucket. Home Phone is deliberately excluded from this fallback — it's often shared by a household/spouse and produced a real false-positive in testing (two different people, same last name and home phone, but different names/mobile/email). If Mobile/Email match more than one candidate, narrow first by Chinese Name, then by a loose "does the name resemble it at all" check (all whitespace/hyphens stripped, one side contains the other — catches e.g. merged `"Zu-Jie"` inside master's `"Joseph Zujie"`, which token-set matching can't since there's no shared separator to split on). This handles cases where a shared family email is cross-attributed to the wrong person but Mobile Phone still identifies the right one, or vice versa. Records matched this way are tagged `Match Method: Contact info (name mismatch)` instead of `Name` — lower-confidence than a name match, worth reviewing separately.
  4. If step 3 also finds zero candidates, search by Last Name across *all* Fellowships (people can legitimately belong to more than one, or master may simply be stale), requiring a **strict** First+Middle token-set match — the "plain First Name alone is enough" tolerance from step 2 is deliberately not used here, since dropping the Fellowship scope means common first names (e.g. "Michael") can otherwise collide with unrelated people of the same last name. Candidates whose Chinese Name actively disagrees are excluded; if more than one candidate still remains, narrow further by Mobile Phone/Email.
  5. If step 4 doesn't resolve to exactly one candidate (whether it found none, or is still stuck with several same-named-but-uncorroborated candidates), fall back further to Mobile Phone/Email matching across *all* Fellowships by Last Name — a clean, unique contact match is stronger evidence than a same-first-name coincidence that step 4's stricter requirements couldn't otherwise disambiguate. This is what finally resolves a case like "Amy Fung" appearing under Jeremiah in the merged file with no Chinese Name filled in, when master only has her under LTG (as "Amy Pui To Fung") - step 4 wrongly picks up two unrelated "Amy Fung"s under other Fellowships by name alone, neither of which has matching contact info, but this step correctly finds the real match by Mobile Phone/Email regardless of Fellowship or name text.
  6. If nothing above resolves to exactly one candidate and merged's `Chinese` is non-empty, drop the Last Name requirement entirely and search *all* of master for a Chinese Name match (exact, or either name containing the other - a compound married surname is sometimes prepended in one file's Chinese Name but not the other's, e.g. merged `"甘伍麗兒"` vs master `"伍麗兒"`) combined with a Mobile Phone/Home Phone/Email/Other Email match. This is the only tier that can find someone whose Last Name itself changed (marriage/surname change) that master hasn't been updated to reflect, or vice versa - every earlier tier is scoped by Last Name and so cannot find them at all, regardless of how strong the other signals are. Chinese Name is required as an anchor (not optional) precisely because dropping the Last Name scope removes the main safety net against a coincidental single-field match; requiring a second corroborating contact signal on top keeps it safe. Unlike every other contact-matching tier, this one also accepts Home Phone (both same-field and cross-field against the other side's Mobile Number) - the household-sharing risk that excludes Home Phone elsewhere is mitigated here since Chinese Name has already narrowed the search to a specific individual, and the two files don't always agree on which field a given number belongs in (a real match, "Feng Leung", was only findable because merged's Mobile Phone equaled master's *Home* Phone).
  7. If step 6 didn't run (Chinese Name blank) or still didn't resolve to exactly one candidate, fall back to the least specific tier: search *all* of master with no Last Name or Chinese Name anchor at all, requiring at least 2 of {Email, Mobile Phone, Home Phone} to independently agree with the same candidate (each compared to its own counterpart field, not cross-field). A single matching field alone isn't reliable with nothing else to scope the search by, but two independently coincidentally matching the wrong person is very unlikely. This is what resolves a merged record with a completely blank Last Name (e.g. "Louis", Fellowship-scoped tiers can't even build a bucket key for an empty Last Name) or blank Chinese Name (e.g. "Loretta Law" whose real master record is "Loretta Mok" - Mobile Phone and Home Phone both agree, but Chinese Name is blank so step 6 never got a chance to run).
  - Matches are tagged `Match Method: Name (different Fellowship)` (step 4), `Contact info (different Fellowship, name mismatch)` (step 5), `Chinese Name + contact info (different Last Name)` (step 6), or `Contact info (2 of 3 signals, no Last Name/Chinese Name match)` (step 7); `Fellowship` and/or `Last Name` are added to `Changed Fields` as appropriate.
- Also checks master's `Other Email` column (not just `Email`) wherever Email is used as a matching signal - a real match ("Georgina Leung") was only findable that way, since her merged Email matched master's `Other Email`, not primary `Email`.
- At every tier, candidates that share the same master `Id` are deduped to one before counting them as "multiple candidates" — master sometimes lists the same person's `Id` more than once, once per Fellowship they belong to (flagged `Duplicate (2x)` in master's own `Duplicate Flag` column), which isn't ambiguity.
- SG Name matching is abbreviation-tolerant (one side is a prefix of the other, e.g. merged `"Faith"` matches master `"Faithfulness"`) and is compared against master's `SG (Compute)` column, falling back to `Small Group` when `SG (Compute)` is blank (master leaves `SG (Compute)` empty whenever `Small Group` holds a person's full name rather than a single-word role/group name).
- Once a record is matched to exactly one master record, these fields are compared to decide "changed" vs "no change": First Name (case-insensitive — flags when the match came from token-set/fallback tolerance rather than an exact name), Middle Name, Last Name (only ever differs for a step-6/7 match), Fellowship (only ever differs for a step-4/5/6/7 match), Chinese/Chinese Name, Home Phone, Mobile Phone/Mobile Number (compared as digits only, ignoring formatting like `(647) 778-3213` vs `647-778-3213`), Email (case-insensitive), SG Leader/SG Leader (Compute), and SG Name (using the same abbreviation-tolerant comparison as the matching step, so an abbreviation difference alone isn't flagged as a change).
- Records with zero master candidates at every tier, or more than one candidate remaining after every tiebreak, are written to `notFound.xlsx` rather than silently dropped — this surfaces both genuinely new people (not yet in the master file) and ambiguous/unresolvable cases for manual review. Includes cases with only a single weak/cross-field signal (e.g. "Belinda Yeung" - merged Home Phone happens to equal a candidate's Mobile Number, but no other field corroborates and Last Name differs) that deliberately don't meet the bar for any tier, since a single coincidental field match isn't reliable enough with no name-based anchor at all.
- Known limitation: nickname mismatches (e.g. merged `"Tony"` vs master `"Anthony"`), incomplete data (e.g. merged has only `"Paul"` where master's combined name is `"Paul Ching Hung"`), and a changed Last Name with fewer than 2 corroborating contact signals available (no Chinese Name, and only one of Email/Mobile/Home Phone matches) aren't resolvable by string matching alone. Simplified vs. traditional Chinese character variants (e.g. `丽` vs `麗`) also aren't normalized, so a Chinese Name match can fail on an otherwise-correct pairing if the two files used different character sets for the same name.
- `readSheet` normalizes blank cells to `''` (not JavaScript `undefined`) so that string-building helpers like `nameTokenSet` never accidentally interpolate the literal text `"undefined"` into a comparison.
