# Description:
Currently RHCCCC have 10 fellowships. Each fellowship maintain members' update using .XLSX file. Every year, information from these .XLSX files are needed to update the members' record in Rock RMS. 
We would like to have a single .XLSX file that combined all the members' new information into a single file BEFORE we process them into Roack RMS.

## Common Issues:
1) Columns headers are not in the same order for all files. This makes it impossible to copy records among files.
2) Some .XLSX file has descriptions in the first two rows, and some .XLSX files only have headers row.
3) Some file have extra columns compare to the LTG file which is used as the standard column for the final merged file.

## Requirements
1) The file from LTG Fellowship is used as standard format for column headers in the result file.
2) Add a new column "Fellowship" into the *LTG file* if the column does not exist. The value should be "LTG" for every row of records.
3) The last column of the final merged file should have the name of the fellowship.

## How to run the script
Usage: node merge.js <file1.xlsx> <file2.xlsx> <fellowship_name> [output.xlsx]
- Both input files MUST have the proper column names. The second file does not need to have the same column ordering.
- The third argument is the fellowship name to tag each file2 record under the "Fellowship" column.
- The fourth argument is optional and specifies the output file name (default: Result.xlsx).
- If the output file resolves to the same path as file1, no new file is built — file2's records are appended directly onto file1's existing sheet in place, leaving file1's original rows, formatting, and other sheets untouched. Otherwise, a fresh output file is created.

## Comparing a merged file against the master report

Once fellowship files are merged, `compare.js` checks the merged records against the master `DynamicGroupReport.xlsx` export (the source of truth for member `Id`) to find which records are new, unchanged, or different.

Usage: node compare.js <masterFile.xlsx> <mergedFile.xlsx> [noChangeOut.xlsx] [yesChangeOut.xlsx]
- Records are matched by Last Name and Fellowship first, then Chinese Name if it uniquely identifies someone in that bucket (checked ahead of the name match, since it's a stronger signal than a bare English first name), otherwise First/Middle Name (tolerant of full-name-in-one-field and name-order differences), with Chinese Name used to break ties when more than one master record matches by name.
- If the name text doesn't line up at all (incomplete data, unrecognized nickname), matching falls back to Mobile Phone or Email within the same Last Name + Fellowship group. Home Phone is intentionally not used for this fallback since it can be shared by a household/spouse.
- If Mobile/Email match more than one candidate, ties are broken first by Chinese Name, then by a loose name-resemblance check (catches nicknames/partial names contact info alone can't disambiguate, e.g. a shared family email cross-attributed to the wrong person).
- If that still finds nothing, matching falls back further to a strict First+Middle match (no loose tolerance) by Last Name across *all* Fellowships, since people can legitimately belong to more than one, or master may be stale. Candidates whose Chinese Name disagrees are excluded, and remaining ties are broken by Mobile Phone/Email.
- If that still doesn't resolve to exactly one candidate, matching falls back once more to Mobile Phone/Email across *all* Fellowships (not just the same one) — a clean, unique contact match beats an unresolved same-name coincidence.
- Candidates sharing the same master `Id` (master sometimes lists one person once per Fellowship they belong to) are deduped before being counted as ambiguous.
- Matched records with no field differences are written to `noChangeOut.xlsx` (default `noChange.xlsx`), tagged with the master `Id` and a `Match Method` column.
- Matched records with at least one field difference are written to `yesChangeOut.xlsx` (default `yesChange.xlsx`), tagged with the master `Id`, a `Match Method` column, and a `Changed Fields` column. `Match Method` is `Name` for a normal name-based match, or one of the fallback labels above when a lower-confidence tier resolved it — worth extra scrutiny, since contact info can be shared between family members and a same-name cross-Fellowship match could coincidentally be a different person.
- Records that can't be resolved to exactly one master record (new members, nickname mismatches, incomplete data, ambiguous same-name matches, etc.) are written to `notFound.xlsx` with a `Reason` column.


