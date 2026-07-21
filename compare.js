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
      // Default missing/blank cells to '' rather than leaving them undefined - template
      // literals elsewhere (e.g. nameTokenSet) would otherwise coerce undefined into the
      // literal text "undefined" and silently corrupt the comparison.
      headers.forEach((col, i) => { if (col) obj[col] = row[i] === undefined ? '' : row[i]; });
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

// Last Name comparisons tolerate hyphen-vs-space formatting (e.g. "Chan-Lam" vs "Chan Lam"),
// the same tolerance already applied to First/Middle names - used wherever Last Name is used as
// a lookup/grouping key, not for the separate "did this field's text change" comparison.
function normLastName(v) {
  return normKey(String(v === undefined || v === null ? '' : v).replace(/-/g, ' '));
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
const mOtherEmail = getCol(masterHeaders, 'Other Email');
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
// Checks merged Email against BOTH master's Email and Other Email columns - a real match
// ("Georgina Leung") was only findable via Other Email, not the primary Email field.
function contactMatches(gMobile, gEmail, mr) {
  const gM = digitsOnly(gMobile), mM = digitsOnly(mr[mMobileNumber]);
  const gE = normKey(gEmail), mE = normKey(mr[mEmail]), mOE = normKey(mr[mOtherEmail]);
  return (gM && mM && gM === mM) || (gE && mE && gE === mE) || (gE && mOE && gE === mOE);
}

// Broader version used only where Chinese Name is already a mandatory, strong anchor (the
// surname-change tier): also accepts Home Phone, and checks it/Mobile Phone cross-field (merged
// Home Phone vs master's Mobile Number, or vice versa) - the two files don't always agree on
// which field a given number belongs in. Home Phone's usual household-sharing risk is mitigated
// here since Chinese Name has already narrowed the search to a specific individual.
function contactMatchesWithHome(gHome, gMobile, gEmail, mr) {
  const gH = digitsOnly(gHome), gM = digitsOnly(gMobile);
  const mH = digitsOnly(mr[mHomePhone]), mM = digitsOnly(mr[mMobileNumber]);
  const phoneMatch = (gH && (gH === mH || gH === mM)) || (gM && (gM === mM || gM === mH));
  return phoneMatch || contactMatches(gMobile, gEmail, mr);
}

// Last resort for records with neither a usable Last Name (missing entirely) nor a Chinese
// Name (blank) to anchor on: require at least 2 of {Email, Mobile Phone, Home Phone} to
// independently agree with the same candidate, each compared to its own counterpart field (not
// cross-field - a single cross-field phone coincidence, e.g. merged Home Phone matching master's
// Mobile Number, isn't strong enough alone and was deliberately left unmatched in testing).
// A single matching field alone (especially Home Phone, shared by a household) isn't reliable
// with no Last Name/Chinese Name to scope the search, but two independent fields both
// coincidentally matching the wrong person is very unlikely.
function twoOfThreeContactMatch(gHome, gMobile, gEmail, mr) {
  const emailMatch = normKey(gEmail) && (normKey(gEmail) === normKey(mr[mEmail]) || normKey(gEmail) === normKey(mr[mOtherEmail]));
  const mobileMatch = digitsOnly(gMobile) && digitsOnly(gMobile) === digitsOnly(mr[mMobileNumber]);
  const homeMatch = digitsOnly(gHome) && digitsOnly(gHome) === digitsOnly(mr[mHomePhone]);
  return [emailMatch, mobileMatch, homeMatch].filter(Boolean).length >= 2;
}

// Loose "does this name look like the same person" check, for breaking ties when Mobile/Email
// match different candidates (e.g. a shared family email cross-attributed to the wrong person).
// Compares First+Middle with all whitespace/hyphens stripped, checking either side contains the
// other - catches cases plain/token-set matching misses, like "Zu-Jie" inside "Joseph Zujie"
// (no shared separator to tokenize on) or "Linda" inside "Woon Hing Linda".
function nameResembles(gFirst, gMiddle, mFirst, mMiddle) {
  const compact = (first, middle) => normKey(`${normText(first)}${normText(middle)}`.replace(/[\s-]/g, ''));
  const g = compact(gFirst, gMiddle), m = compact(mFirst, mMiddle);
  return !!g && !!m && (g.includes(m) || m.includes(g));
}

// Master sometimes lists the same person's Id more than once (once per Fellowship they belong
// to - e.g. Id 16649 appears under both Caleb and LTG for one member, flagged "Duplicate (2x)").
// That's the same record, not multiple candidates - collapse to one, preferring whichever row's
// Fellowship matches the merged record's own Fellowship if there's a choice.
function dedupeById(candidates, idCol, fellowshipCol, mergedFellowship) {
  const byId = new Map();
  candidates.forEach(c => {
    const id = c[idCol];
    if (!byId.has(id)) byId.set(id, []);
    byId.get(id).push(c);
  });
  return [...byId.values()].map(rows => rows.find(r => normKey(fellowshipLastWord(r[fellowshipCol])) === normKey(mergedFellowship)) || rows[0]);
}

// Mobile/Email match within a candidate pool, with the same tiebreak cascade used whether the
// pool is scoped to one Fellowship or spans all of them: dedupe by Id, then narrow by Chinese
// Name, then by loose name resemblance (handles a shared family email/mobile cross-attributed
// to the wrong person, where one candidate's name still looks like the merged record's).
function matchByContact(pool, mergedRow, mergedFellowship) {
  let candidates = dedupeById(
    pool.filter(mr => contactMatches(mergedRow[gMobilePhone], mergedRow[gEmail], mr)),
    mId, mFellowship, mergedFellowship);

  if (candidates.length > 1) {
    const chineseFiltered = candidates.filter(mr => normText(mr[mChineseName]) === normText(mergedRow[gChinese]));
    if (chineseFiltered.length >= 1) candidates = chineseFiltered;
  }
  if (candidates.length > 1) {
    const nameFiltered = candidates.filter(mr => nameResembles(mergedRow[gFirstName], mergedRow[gMiddleName], mr[mFirstName], mr[mMiddleName]));
    if (nameFiltered.length >= 1) candidates = nameFiltered;
  }
  return candidates;
}

// Index master rows by (lastName, fellowshipLastWord); first name (with the split-inconsistency
// tolerance above) and SG (abbreviation-tolerant) are filtered from each bucket per merged record.
const masterIndex = new Map();
masterRows.forEach(row => {
  const key = [normLastName(row[mLastName]), normKey(fellowshipLastWord(row[mFellowship]))].join('|');
  if (!masterIndex.has(key)) masterIndex.set(key, []);
  masterIndex.get(key).push(row);
});

// Index master rows by Last Name only, for the cross-Fellowship fallback tier below.
const masterByLastName = new Map();
masterRows.forEach(row => {
  const key = normLastName(row[mLastName]);
  if (!masterByLastName.has(key)) masterByLastName.set(key, []);
  masterByLastName.get(key).push(row);
});

const noChangeRecords = [];
const yesChangeRecords = [];
const notFoundRecords = [];

// Field comparisons: [merged column, master column, comparator]
const fieldComparisons = [
  [gFirstName, mFirstName, (a, b) => normKey(a) === normKey(b)],
  [gMiddleName, mMiddleName, (a, b) => normText(a) === normText(b)],
  // Only ever differs for the Chinese Name + contact info (different Last Name) tier - a
  // marriage/surname change master hasn't been updated to reflect, or vice versa.
  [gLastName, mLastName, (a, b) => normKey(a) === normKey(b)],
  // Master's Fellowship is a full string ("Bayview Cantonese - 2023 Esther"); merged's is just
  // the last word ("Esther"). Only ever differs for cross-Fellowship fallback matches (see
  // below), since the other tiers already scope by matching Fellowship.
  [gFellowship, mFellowship, (a, b) => normKey(a) === normKey(fellowshipLastWord(b))],
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
  const key = [normLastName(mergedRow[gLastName]), normKey(mergedRow[gFellowship])].join('|');
  const bucket = masterIndex.get(key) || [];
  let matchMethod = 'Name';

  // Chinese Name is checked first, ahead of the name match: it's a stronger, less collision-
  // prone signal than a bare English first name (which can't tell two different "Amy"s in the
  // same bucket apart). Without this, a plain-name match landing on exactly one WRONG candidate
  // (Chinese Name blank/mismatched) was silently preferred over a different bucket member whose
  // Chinese Name matched exactly - found via two different merged "Amy Fung" LTG rows both
  // resolving to the same wrong master Id, only one of which actually belonged there.
  const gChineseName = normText(mergedRow[gChinese]);
  const chineseMatches = gChineseName
    ? bucket.filter(mr => normText(mr[mChineseName]) === gChineseName)
    : [];

  let candidates = chineseMatches.length === 1
    ? chineseMatches
    : dedupeById(
        bucket.filter(mr => firstNameMatches(mergedRow[gFirstName], mergedRow[gMiddleName], mr[mFirstName], mr[mMiddleName])),
        mId, mFellowship, mergedRow[gFellowship]);

  if (candidates.length > 1) {
    const chineseFiltered = candidates.filter(mr => normText(mr[mChineseName]) === gChineseName);
    if (chineseFiltered.length >= 1) candidates = chineseFiltered;
  }

  if (candidates.length === 0) {
    // Name text didn't line up (incomplete data, unrecognized nickname, etc). Fall back to
    // Mobile Phone / Email, which can uniquely identify the same person within this bucket.
    matchMethod = 'Contact info (name mismatch)';
    candidates = matchByContact(bucket, mergedRow, mergedRow[gFellowship]);
  }

  if (candidates.length === 0) {
    // Same person may be a member of a different Fellowship than master has them scoped
    // under (people can join more than one). Search by Last Name across all Fellowships,
    // requiring a strict First+Middle token-set match - the loose plain-First-only tolerance
    // used above is only safe when also scoped by Fellowship, since dropping that scope lets
    // common first names collide with unrelated people of the same last name. Also exclude
    // candidates whose Chinese Name actively disagrees, as further protection against that.
    matchMethod = 'Name (different Fellowship)';
    const gTokenSet = nameTokenSet(mergedRow[gFirstName], mergedRow[gMiddleName]);
    const crossFellowshipPool = masterByLastName.get(normLastName(mergedRow[gLastName])) || [];

    // An exact Chinese Name match is authoritative here too, same as the within-Fellowship
    // Tier 1 check - Last Name already matches (this tier only drops the Fellowship scope), so
    // it doesn't need the strict name-token match below when Chinese Name alone is unique and
    // exact (e.g. "Billy" whose Middle Name partially matches master's "Yiu Wah Vincent" - the
    // English name text differs but the full Chinese Name matches exactly).
    const crossFellowshipChineseMatches = gChineseName
      ? dedupeById(crossFellowshipPool.filter(mr => normText(mr[mChineseName]) === gChineseName), mId, mFellowship, mergedRow[gFellowship])
      : [];

    candidates = crossFellowshipChineseMatches.length === 1
      ? crossFellowshipChineseMatches
      : dedupeById(
          crossFellowshipPool
            .filter(mr => nameTokenSet(mr[mFirstName], mr[mMiddleName]) === gTokenSet)
            .filter(mr => {
              const gC = normText(mergedRow[gChinese]), mC = normText(mr[mChineseName]);
              return !gC || !mC || gC === mC;
            }),
          mId, mFellowship, mergedRow[gFellowship]);

    if (candidates.length > 1) {
      // Name+Chinese alone didn't narrow it down (e.g. two same-named different people, no
      // Chinese on file) - Mobile/Email corroboration can still resolve it.
      const contactFiltered = candidates.filter(mr => contactMatches(mergedRow[gMobilePhone], mergedRow[gEmail], mr));
      if (contactFiltered.length >= 1) candidates = contactFiltered;
    }

    if (candidates.length !== 1) {
      // Either nothing matched by name, or multiple same-named-but-otherwise-uncorroborated
      // candidates are stuck unresolved (e.g. two different "Amy Fung"s, neither of which is
      // the true match - the strict name-token requirement above can wrongly exclude the real
      // match, like "Amy" vs master's "Amy Pui To"). A clean, unique Mobile/Email match anywhere
      // in this Last Name pool is stronger evidence than a same-name-only coincidence, so it
      // takes priority over an unresolved name-tier result.
      const contactCandidates = matchByContact(crossFellowshipPool, mergedRow, mergedRow[gFellowship]);
      if (contactCandidates.length === 1) {
        candidates = contactCandidates;
        matchMethod = 'Contact info (different Fellowship, name mismatch)';
      }
    }
  }

  if (candidates.length !== 1 && normText(mergedRow[gChinese])) {
    // Last Name itself can differ from master's (marriage/surname change - e.g. merged "Kam"
    // vs master "Kam-Ng", or merged "Lung" vs master "Lam"), which defeats every tier above
    // since they're all scoped by Last Name. With that scope gone, search all of master
    // requiring Chinese Name to match exactly (a strong, low-collision anchor) AND at least one
    // of Mobile Phone/Email/Other Email to also agree - Chinese Name alone isn't required to be
    // globally unique, so a second corroborating signal keeps this safe.
    const gChineseName = normText(mergedRow[gChinese]);
    const surnameChangeCandidates = dedupeById(
      masterRows
        // A compound married surname is sometimes prepended in one file's Chinese Name but not
        // the other's (e.g. merged "甘伍麗兒" vs master "伍麗兒" - the "甘"/Kam character), so
        // one containing the other counts as a match, not just exact equality.
        .filter(mr => {
          const mC = normText(mr[mChineseName]);
          return mC.length >= 2 && (mC === gChineseName || gChineseName.includes(mC) || mC.includes(gChineseName));
        })
        .filter(mr => contactMatchesWithHome(mergedRow[gHomePhone], mergedRow[gMobilePhone], mergedRow[gEmail], mr)),
      mId, mFellowship, mergedRow[gFellowship]);

    if (surnameChangeCandidates.length === 1) {
      candidates = surnameChangeCandidates;
      matchMethod = 'Chinese Name + contact info (different Last Name)';
    }
  }

  if (candidates.length !== 1) {
    // Neither a Last Name nor a Chinese Name match was available to anchor on (e.g. merged Last
    // Name is blank entirely, or Chinese Name is blank so the tier above couldn't run). Fall
    // back to requiring 2 independent contact signals in agreement, searched across all of
    // master with no name-based scoping at all.
    let strongContactCandidates = dedupeById(
      masterRows.filter(mr => twoOfThreeContactMatch(mergedRow[gHomePhone], mergedRow[gMobilePhone], mergedRow[gEmail], mr)),
      mId, mFellowship, mergedRow[gFellowship]);

    if (strongContactCandidates.length > 1) {
      // Phone numbers have turned out to be less reliable than email in this data - master has
      // real cases of unrelated people sharing the exact same Home/Mobile Number (likely copied
      // across rows within the same small group by mistake), which can satisfy "2 of 3" on phone
      // alone for more than one wrong candidate. Email is more personally distinctive, so an
      // exact match narrows to the right person even when phone numbers collide.
      const gE = normKey(mergedRow[gEmail]);
      const emailFiltered = gE && strongContactCandidates.filter(mr => gE === normKey(mr[mEmail]) || gE === normKey(mr[mOtherEmail]));
      if (emailFiltered && emailFiltered.length === 1) strongContactCandidates = emailFiltered;
    }

    if (strongContactCandidates.length === 1) {
      candidates = strongContactCandidates;
      matchMethod = 'Contact info (2 of 3 signals, no Last Name/Chinese Name match)';
    }
  }

  if (candidates.length === 0) {
    notFoundRecords.push({ ...mergedRow, Reason: 'No master record found' });
    return;
  }
  if (candidates.length > 1) {
    const reason = matchMethod === 'Name' ? 'Multiple master candidates even after Chinese tiebreak'
      : matchMethod === 'Contact info (name mismatch)' ? 'Multiple master candidates matched by Mobile Phone/Email'
      : matchMethod === 'Name (different Fellowship)' ? 'Multiple master candidates across different Fellowships (same name)'
      : 'Multiple master candidates across different Fellowships matched by Mobile Phone/Email';
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

const matchMethodCounts = [...noChangeRecords, ...yesChangeRecords].reduce((acc, r) => {
  acc[r['Match Method']] = (acc[r['Match Method']] || 0) + 1;
  return acc;
}, {});
const fallbackBreakdown = Object.entries(matchMethodCounts)
  .filter(([method]) => method !== 'Name')
  .map(([method, count]) => `${count} via ${method.toLowerCase()}`)
  .join(', ');
const reasonCounts = notFoundRecords.reduce((acc, r) => {
  acc[r.Reason] = (acc[r.Reason] || 0) + 1;
  return acc;
}, {});

console.log('');
console.log('=== Compare Summary ===');
console.log(`Master file (${path.basename(masterFile)}): ${masterRows.length} records`);
console.log(`Merged file (${path.basename(mergedFile)}): ${mergedRows.length} records`);
console.log(`No change (${noChangeFile}): ${noChangeRecords.length}`);
console.log(`Changed (${yesChangeFile}): ${yesChangeRecords.length}${fallbackBreakdown ? ` (${fallbackBreakdown})` : ''}`);
if (notFoundRecords.length) {
  const breakdown = Object.entries(reasonCounts).map(([reason, count]) => `${count} ${reason.toLowerCase()}`).join(', ');
  console.log(`Not found (${notFoundFile}): ${notFoundRecords.length} (${breakdown})`);
}
console.log('');
