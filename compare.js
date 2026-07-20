const XLSX = require('xlsx');
const path = require('path');

// Read input file arguments
const [masterFile, mergedFile, noChangeOut, yesChangeOut] = process.argv.slice(2);

if (!masterFile || !mergedFile) {
  console.error('Usage: node compare.js <masterFile.xlsx> <mergedFile.xlsx> [noChangeOut.xlsx] [yesChangeOut.xlsx]');
  console.error('masterFile is the DynamicGroupReport-style file (source of truth, has "Id").');
  console.error('mergedFile is the Merged-Fellowships-style file to check for differences against masterFile.');
  console.error('noChangeOut/yesChangeOut are optional and default to noChange.xlsx / yesChange.xlsx.');
  process.exit(1);
}

const noChangeFile = noChangeOut || 'noChange.xlsx';
const yesChangeFile = yesChangeOut || 'yesChange.xlsx';
const notFoundFile = 'notFound.xlsx';

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

function readSheet(file) {
  const wb = XLSX.readFile(file);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headerIdx = findHeaderRowIndex(rawRows);
  const headers = rawRows[headerIdx];
  const rows = rawRows.slice(headerIdx + 1)
    .filter(row => row.some(c => c !== null && c !== undefined && String(c).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((col, i) => { if (col) obj[col] = row[i]; });
      return obj;
    });
  return { headers, rows };
}

function getCol(headers, name) {
  return headers.find(h => h && h.toLowerCase() === name.toLowerCase());
}

// Trim + collapse internal whitespace
function normText(v) {
  return (v === undefined || v === null) ? '' : String(v).trim().replace(/\s+/g, ' ');
}

function normKey(v) {
  return normText(v).toLowerCase();
}

function digitsOnly(v) {
  return (v === undefined || v === null) ? '' : String(v).replace(/\D/g, '');
}

// Last token of the master Fellowship string, e.g. "Bayview Cantonese - 2023 Esther" -> "Esther"
function fellowshipLastWord(v) {
  const t = normText(v);
  if (!t) return '';
  const parts = t.split(' ');
  return parts[parts.length - 1];
}

// Abbreviation-tolerant SG match: one side is a prefix of the other (e.g. "Faith" / "Faithfulness")
function sgMatches(mergedSg, masterSg) {
  const a = normKey(mergedSg), b = normKey(masterSg);
  if (!a && !b) return true;
  if (!a || !b) return false;
  return a.startsWith(b) || b.startsWith(a);
}

const { headers: masterHeaders, rows: masterRows } = readSheet(masterFile);
const { headers: mergedHeaders, rows: mergedRows } = readSheet(mergedFile);

const mId = getCol(masterHeaders, 'Id');
const mLastName = getCol(masterHeaders, 'Last Name');
const mFirstName = getCol(masterHeaders, 'First Name');
const mMiddleName = getCol(masterHeaders, 'Middle Name');
const mFellowship = getCol(masterHeaders, 'Fellowship');
const mSgCompute = getCol(masterHeaders, 'SG (Compute)');
const mSmallGroup = getCol(masterHeaders, 'Small Group');
const mSgLeaderCompute = getCol(masterHeaders, 'SG Leader (Compute)');
const mChineseName = getCol(masterHeaders, 'Chinese Name');
const mEmail = getCol(masterHeaders, 'Email');
const mHomePhone = getCol(masterHeaders, 'Home Phone');
const mMobileNumber = getCol(masterHeaders, 'Mobile Number');

const gLastName = getCol(mergedHeaders, 'Last Name');
const gFirstName = getCol(mergedHeaders, 'First Name');
const gMiddleName = getCol(mergedHeaders, 'Middle Name');
const gFellowship = getCol(mergedHeaders, 'Fellowship');
const gSgName = getCol(mergedHeaders, 'SG Name');
const gSgLeader = getCol(mergedHeaders, 'SG Leader');
const gChinese = getCol(mergedHeaders, 'Chinese');
const gEmail = getCol(mergedHeaders, 'Email');
const gHomePhone = getCol(mergedHeaders, 'Home Phone');
const gMobilePhone = getCol(mergedHeaders, 'Mobile Phone');

// Master sometimes stores the full first+middle name in the First Name column alone
// (Middle Name left blank) instead of splitting it like the merged file does, and the
// name tokens are sometimes in a different order (e.g. "Wai King Peggy" vs "Peggy" / "Wai King").
// Compare as an order-independent token set of First+Middle from both sides.
function nameTokenSet(first, middle) {
  // Compound name syllables are hyphenated in some records and space-separated in others
  // (e.g. "Shuk-Yin" vs "Shuk Yin") - normalize hyphens to spaces so both tokenize the same.
  return normKey(`${first} ${middle}`.replace(/-/g, ' ')).split(' ').filter(Boolean).sort().join(' ');
}
function firstNameMatches(gFirst, gMiddle, mFirst, mMiddle) {
  // Plain First Name alone often matches even when Middle Name data differs/is missing on
  // one side, so that's checked on its own in addition to the full token-set comparison.
  return normKey(gFirst) === normKey(mFirst) || nameTokenSet(gFirst, gMiddle) === nameTokenSet(mFirst, mMiddle);
}

// Fallback for records where the name text itself doesn't line up (e.g. incomplete name data,
// unrecognized nickname) but Mobile Phone or Email uniquely identifies the same person within
// the same Last Name + Fellowship bucket. Home Phone is deliberately excluded - it's often
// shared by a household/spouse and produced a real false-positive in testing (two different
// people with the same last name and home phone number, but different names/mobile/email).
function contactMatches(gMobile, gEmail, mMobile, mEmail) {
  const gM = digitsOnly(gMobile), mM = digitsOnly(mMobile);
  const gE = normKey(gEmail), mE = normKey(mEmail);
  return (gM && mM && gM === mM) || (gE && mE && gE === mE);
}

// Index master rows by (lastName, fellowshipLastWord); first name (with the split-inconsistency
// tolerance above) and SG (abbreviation-tolerant) are filtered from each bucket per merged record.
const masterIndex = new Map();
masterRows.forEach(row => {
  const key = [normKey(row[mLastName]), normKey(fellowshipLastWord(row[mFellowship]))].join('|');
  if (!masterIndex.has(key)) masterIndex.set(key, []);
  masterIndex.get(key).push(row);
});

const noChangeRecords = [];
const yesChangeRecords = [];
const notFoundRecords = [];

// Field comparisons: [merged column, master column, comparator]
const fieldComparisons = [
  [gFirstName, mFirstName, (a, b) => normKey(a) === normKey(b)],
  [gMiddleName, mMiddleName, (a, b) => normText(a) === normText(b)],
  [gChinese, mChineseName, (a, b) => normText(a) === normText(b)],
  [gHomePhone, mHomePhone, (a, b) => digitsOnly(a) === digitsOnly(b)],
  [gMobilePhone, mMobileNumber, (a, b) => digitsOnly(a) === digitsOnly(b)],
  [gEmail, mEmail, (a, b) => normKey(a) === normKey(b)],
  // Master leaves SG Leader (Compute) blank by design for role-based small groups (e.g.
  // "Joy", "Purple" - a single-word group name with no recorded leader), so an empty
  // master value isn't a real difference to flag - only compare when master has one on file.
  [gSgLeader, mSgLeaderCompute, (a, b) => !normText(b) || normText(a) === normText(b)],
  [gSgName, null, (a, b, master) => sgMatches(a, master[mSgCompute] || master[mSmallGroup])],
];

mergedRows.forEach(mergedRow => {
  const key = [normKey(mergedRow[gLastName]), normKey(mergedRow[gFellowship])].join('|');
  const bucket = masterIndex.get(key) || [];
  let matchMethod = 'Name';
  let candidates = bucket.filter(mr => firstNameMatches(mergedRow[gFirstName], mergedRow[gMiddleName], mr[mFirstName], mr[mMiddleName]));

  if (candidates.length > 1) {
    const chineseFiltered = candidates.filter(mr => normText(mr[mChineseName]) === normText(mergedRow[gChinese]));
    if (chineseFiltered.length >= 1) candidates = chineseFiltered;
  }

  if (candidates.length === 0) {
    // Name text didn't line up (incomplete data, unrecognized nickname, etc). Fall back to
    // Mobile Phone / Email, which can uniquely identify the same person within this bucket.
    matchMethod = 'Contact info (name mismatch)';
    candidates = bucket.filter(mr => contactMatches(mergedRow[gMobilePhone], mergedRow[gEmail], mr[mMobileNumber], mr[mEmail]));
  }

  if (candidates.length === 0) {
    notFoundRecords.push({ ...mergedRow, Reason: 'No master record found' });
    return;
  }
  if (candidates.length > 1) {
    const reason = matchMethod === 'Name'
      ? 'Multiple master candidates even after Chinese tiebreak'
      : 'Multiple master candidates matched by Mobile Phone/Email';
    notFoundRecords.push({ ...mergedRow, Reason: reason });
    return;
  }

  const master = candidates[0];
  const changedFields = fieldComparisons
    .filter(([gCol, mCol, cmp]) => gCol && !cmp(mergedRow[gCol], mCol ? master[mCol] : undefined, master))
    .map(([gCol]) => gCol);

  const outRow = { Id: master[mId], ...mergedRow, 'Match Method': matchMethod };
  if (changedFields.length === 0) {
    noChangeRecords.push(outRow);
  } else {
    yesChangeRecords.push({ ...outRow, 'Changed Fields': changedFields.join(', ') });
  }
});

const noChangeHeaders = ['Id', ...mergedHeaders, 'Match Method'];
const yesChangeHeaders = ['Id', ...mergedHeaders, 'Match Method', 'Changed Fields'];

const wsNoChange = XLSX.utils.json_to_sheet(noChangeRecords, { header: noChangeHeaders });
const wbNoChange = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbNoChange, wsNoChange, 'noChange');
XLSX.writeFile(wbNoChange, noChangeFile);

const wsYesChange = XLSX.utils.json_to_sheet(yesChangeRecords, { header: yesChangeHeaders });
const wbYesChange = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbYesChange, wsYesChange, 'yesChange');
XLSX.writeFile(wbYesChange, yesChangeFile);

const notFoundHeaders = [...mergedHeaders, 'Reason'];
const wsNotFound = XLSX.utils.json_to_sheet(notFoundRecords, { header: notFoundHeaders });
const wbNotFound = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbNotFound, wsNotFound, 'notFound');
XLSX.writeFile(wbNotFound, notFoundFile);

const contactMatchCount = [...noChangeRecords, ...yesChangeRecords].filter(r => r['Match Method'] !== 'Name').length;
const reasonCounts = notFoundRecords.reduce((acc, r) => {
  acc[r.Reason] = (acc[r.Reason] || 0) + 1;
  return acc;
}, {});

console.log('');
console.log('=== Compare Summary ===');
console.log(`Master file (${path.basename(masterFile)}): ${masterRows.length} records`);
console.log(`Merged file (${path.basename(mergedFile)}): ${mergedRows.length} records`);
console.log(`No change (${noChangeFile}): ${noChangeRecords.length}`);
console.log(`Changed (${yesChangeFile}): ${yesChangeRecords.length}${contactMatchCount ? ` (${contactMatchCount} matched via Mobile Phone/Email fallback, not name)` : ''}`);
if (notFoundRecords.length) {
  const breakdown = Object.entries(reasonCounts).map(([reason, count]) => `${count} ${reason.toLowerCase()}`).join(', ');
  console.log(`Not found (${notFoundFile}): ${notFoundRecords.length} (${breakdown})`);
}
console.log('');
