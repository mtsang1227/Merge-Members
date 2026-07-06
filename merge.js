const XLSX = require('xlsx');
const path = require('path');

// Read input file arguments
const [file1, file2, fellowshipName, file3] = process.argv.slice(2);

if (!file1 || !file2 || !fellowshipName) {
  console.error('Usage: node merge.js <file1.xlsx> <file2.xlsx> <fellowship_name> [output.xlsx]');
  console.error('Both input files MUST have the proper column names. The second file does not need to have the same column ordering.');
  console.error('The third argument is the fellowship name to tag each file2 record under the "Fellowship" column.');
  console.error('The fourth argument is optional and specifies the output file name (default: Result.xlsx).');
  process.exit(1);
}

// Load both workbooks
const wb1 = XLSX.readFile(file1);
const wb2 = XLSX.readFile(file2);

// Get the first sheet from each
const ws1 = wb1.Sheets[wb1.SheetNames[0]];
const ws2 = wb2.Sheets[wb2.SheetNames[0]];

// Within the first maxLookAhead rows, return the index of the row with the most non-empty cells.
// Description rows are typically sparse (1-2 cells); the real header row fills every column.
function findHeaderRowIndex(rawRows, maxLookAhead = 5) {
  const limit = Math.min(maxLookAhead, rawRows.length);
  let bestIdx = 0, bestCount = -1;
  for (let i = 0; i < limit; i++) {
    const count = (rawRows[i] || []).filter(c => c !== null && c !== undefined && c !== '').length;
    if (count > bestCount) { bestCount = count; bestIdx = i; }
  }
  return bestIdx;
}

// Read file1: detect header row, build row objects keyed by header name
const f1RawRows = XLSX.utils.sheet_to_json(ws1, { header: 1 });
const f1HeaderIdx = findHeaderRowIndex(f1RawRows);
const f1Headers = f1RawRows[f1HeaderIdx];
const f1Rows = f1RawRows.slice(f1HeaderIdx + 1)
  .filter(row => row.some(c => c !== null && c !== undefined && String(c).trim() !== ''))
  .map(row => {
    const obj = {};
    f1Headers.forEach((col, i) => { if (col) obj[col] = row[i]; });
    return obj;
  });

// Resolve the canonical "Fellowship" column name (case-insensitive match against file1 headers, or use as-is)
const fellowshipCol = f1Headers.find(h => h && h.toLowerCase() === 'fellowship') || 'Fellowship';

// Read file2: detect header row, re-map columns to file1's header names
const f2RawRows = XLSX.utils.sheet_to_json(ws2, { header: 1 });
const f2HeaderIdx = findHeaderRowIndex(f2RawRows);
const f2Headers = f2RawRows[f2HeaderIdx];
const f2Rows = f2RawRows.slice(f2HeaderIdx + 1)
  .filter(row => row.some(c => c !== null && c !== undefined && String(c).trim() !== ''))
  .map(row => {
    const obj = {};
    f2Headers.forEach((col, i) => {
      if (!col) return;
      const f1Col = f1Headers.find(h => h && h.toLowerCase() === col.toLowerCase());
      if (f1Col) obj[f1Col] = row[i];
    });
    obj[fellowshipCol] = fellowshipName;
    return obj;
  });

// Merge: file1 rows first, then file2 rows
const merged = [...f1Rows, ...f2Rows];

// Build output headers: file1's headers, plus Fellowship if not already present
const outHeaders = f1Headers.includes(fellowshipCol)
  ? f1Headers
  : [...f1Headers, fellowshipCol];

const outputFile = file3 || 'Result.xlsx';

// If the output path is file1 itself, avoid rebuilding a brand new workbook from scratch:
// just append file2's rows onto file1's existing sheet, leaving its original rows,
// formatting, and any other sheets untouched.
const inPlace = path.resolve(file1) === path.resolve(outputFile);

if (inPlace) {
  if (!f1Headers.includes(fellowshipCol)) {
    const headerCell = XLSX.utils.encode_cell({ r: f1HeaderIdx, c: f1Headers.length });
    XLSX.utils.sheet_add_aoa(ws1, [[fellowshipCol]], { origin: headerCell });
  }
  XLSX.utils.sheet_add_json(ws1, f2Rows, {
    header: outHeaders,
    skipHeader: true,
    origin: f1RawRows.length,
  });
  XLSX.writeFile(wb1, outputFile);
} else {
  // Build a fresh output workbook
  const wsOut = XLSX.utils.json_to_sheet(merged, { header: outHeaders });
  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, wsOut, 'Merged');
  XLSX.writeFile(wbOut, outputFile);
}

// Summary
console.log('');
console.log('=== Merge Summary ===');
console.log(`Input file 1 (${path.basename(file1)}): ${f1Rows.length} records`);
console.log(`Input file 2 (${path.basename(file2)}): ${f2Rows.length} records`);
console.log(`Records merged into ${outputFile}: ${merged.length}`);
console.log('');
