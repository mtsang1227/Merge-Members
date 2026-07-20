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
- Records are matched by Last Name, First/Middle Name (tolerant of full-name-in-one-field and name-order differences), Fellowship, and SG Name/Small Group (tolerant of abbreviations), with Chinese Name used to break ties when more than one master record matches.
- Matched records with no field differences are written to `noChangeOut.xlsx` (default `noChange.xlsx`), tagged with the master `Id`.
- Matched records with at least one field difference are written to `yesChangeOut.xlsx` (default `yesChange.xlsx`), tagged with the master `Id` and a `Changed Fields` column.
- Records that can't be resolved to exactly one master record (new members, nickname mismatches, incomplete data, etc.) are written to `notFound.xlsx` with a `Reason` column.


