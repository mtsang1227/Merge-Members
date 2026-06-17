const XLSX = require('xlsx');
const path = require('path');

// Read input file arguments
const [file1, file2] = process.argv.slice(2);

if (!file1 || !file2) {
  console.error('Usage: node merge.js <file1.xlsx> <file2.xlsx>');
  process.exit(1);
}

// Load both workbooks
const wb1 = XLSX.readFile(file1);
const wb2 = XLSX.readFile(file2);

// Get the first sheet from each
const ws1 = wb1.Sheets[wb1.SheetNames[0]];
const ws2 = wb2.Sheets[wb2.SheetNames[0]];

// Read file1 with its headers (used as the canonical column order)
const f1Headers = XLSX.utils.sheet_to_json(ws1, { header: 1 })[0];
const f1Rows = XLSX.utils.sheet_to_json(ws1); // keyed by file1's header names

// Read file2 rows, re-mapped to file1's header names
const f2RawRows = XLSX.utils.sheet_to_json(ws2, { header: 1 });
const f2Headers = f2RawRows[0];
const f2Rows = f2RawRows.slice(1).map(row => {
  const obj = {};
  f2Headers.forEach((col, i) => {
    // Find the matching file1 header (case-insensitive)
    const f1Col = f1Headers.find(h => h.toLowerCase() === col.toLowerCase());
    if (f1Col) obj[f1Col] = row[i];
  });
  return obj;
});

// Merge: file1 rows first, then file2 rows
const merged = [...f1Rows, ...f2Rows];

// Build output workbook using file1's headers
const wsOut = XLSX.utils.json_to_sheet(merged, { header: f1Headers });
const wbOut = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbOut, wsOut, 'Merged');

const outputFile = 'Result.xlsx';
XLSX.writeFile(wbOut, outputFile);

// Summary
console.log('');
console.log('=== Merge Summary ===');
console.log(`Input file 1 (${path.basename(file1)}): ${f1Rows.length} records`);
console.log(`Input file 2 (${path.basename(file2)}): ${f2Rows.length} records`);
console.log(`Records merged into ${outputFile}: ${merged.length}`);
console.log('');
