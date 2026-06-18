# Description:
Currently RHCCCC have 10 fellowships. Each fellowship maintain members' update using .XLSX file. Every year, information from these .XLSX files are needed to update the members' record in Rock RMS. 
We would like to have a single .XLSX file that combined all the members' new information into a single file BEFORE we process them into Roack RMS.

## Common Issues:
1) Columns headers are not in the same order for all files. This makes it impossible to copy records among files.
2) Some .XLSX file has descriptions in the first two rows, and some .XLSX files only have headers row.
3) Some file have extra columns compare to the LTG file which is used as the standard column for the final merged file.

A merge.js file is built to speed up the merging process and resolve the issues above
