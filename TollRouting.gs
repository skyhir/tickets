/**
 * Toll routing from the tickets uploader into toll-processing!current_tolls, plus the
 * "is this car on the right toll account?" check written to the Tolls 2026 tab.
 *
 * Why this file exists (2026-11-04):
 *  - current_tolls is in the toll pipeline's PRE-step-4 layout (the processed_tolls
 *    headers: Date (PST) / Time (PST) in F:G), but the old router wrote a fixed
 *    18-column POST-step-4 row, so every ticket-routed toll landed two columns off.
 *    Rows are now written BY HEADER NAME, so they land correctly in either layout.
 *  - Unmatched plates were written as "#N/A". The toll pipeline's convention is to
 *    put the plate in "Envoy #" when no fleet match exists, so operators can still
 *    find the car; the router now does the same.
 *  - Collections letters and tolls with no time can't be attributed to a driver, so
 *    they stay on the Tolls 2026 tab (accounting still pays the agency) and are NOT
 *    sent to current_tolls.
 */

// Header names on toll-processing!current_tolls (processed_tolls layout).
const CT_HEADERS_ = {
  FIRST: 'First Name', LAST: 'Last Name', EMAIL: 'User Email', ENVOY: 'Envoy #',
  BOOKING: 'Booking ID', DATE_PST: 'Date (PST)', TIME_PST: 'Time (PST)',
  LOCAL: 'Time and Date (Local)', TYPE: 'Type', TOLL: 'Toll Amount',
  ADMIN: 'Admin Fee', TOTAL: 'Total Due', NOTES: 'Notes'
};
// Canonical current_tolls header row (= processed_tolls row 1). Used when re-seeding.
const CT_PROCESSED_LAYOUT_ = [
  'First Name', 'Last Name', 'User Email', 'Envoy #', 'Booking ID',
  'Date (PST)', 'Time (PST)', 'Time and Date (Local)', 'Type', 'Toll Amount',
  'Admin Fee', 'Total Due', 'Confirmed?', 'Charged?', 'Notified?',
  'Total Value - Booking', '', 'Total Value: Pure Tolls', 'Total Value: Admin', 'Notes'
];
// Same formats as toll-processing/reformatting.gs so routed rows read identically.
const CT_PST_ZONE_ = 'America/Los_Angeles';
const CT_DATE_PST_FORMAT_ = 'MM/dd/yyyy';
const CT_TIME_PST_FORMAT_ = 'hh:mm a';
const CT_LOCAL_FORMAT_ = "h:mma 'on' EEEE, MMMM d, yyyy";

// Tolls 2026 tab: agency + verification columns, located by header text.
const TOLLS_AGENCY_HEADER = 'Agency Status';            // AD: toll program the car should be on (existing column)
const TOLLS_ACTION_HEADER = 'Action Required';          // AE (existing column)
const TOLLS_VERIFY_HEADER = 'Vehicle on Agency Account?'; // AF (added 2026-11-04)
const TOLLS_VERIFY_VALUES = ['Yes', 'No - needs adding', 'Cannot add - on another account', 'Unknown vehicle', 'Unverified'];

// Driver Billing Status values for tolls that are NOT sent to current_tolls.
const BILLING_STATUS_COLLECTIONS = 'Collections - No Toll Detail';
const BILLING_STATUS_NO_TIME = 'Toll - Date/Time Unknown';
const BILLING_STATUS_FEE_ONLY = 'Agency Fee - Not Billed to Driver';
const BILLING_STATUS_TRANSFERRED = 'Transferred to Toll Sheet';

// Transponder Index = the E-ZPass account roster toll-processing already uses
// (Transponder / License Plate / Envoy #). Opened as the operator running the upload.
const TRANSPONDER_INDEX_SPREADSHEET_ID = '1MyfE0ZZLr0RWRdBOkOnRpawmOQkZofSfLeqYT5NCYd0';
const TRANSPONDER_INDEX_SHEET_NAME = 'Transponder Index';

// Local zone per state where the toll road is (copy of toll-processing TOLL_NOTICE_STATE_TZ_).
const TOLL_STATE_TZ_ = {
  AL: 'America/Chicago', AZ: 'America/Phoenix', CA: 'America/Los_Angeles', CO: 'America/Denver',
  CT: 'America/New_York', DE: 'America/New_York', FL: 'America/New_York', GA: 'America/New_York',
  IL: 'America/Chicago', IN: 'America/Indiana/Indianapolis', KS: 'America/Chicago', KY: 'America/New_York',
  MA: 'America/New_York', MD: 'America/New_York', ME: 'America/New_York', MI: 'America/Detroit',
  MN: 'America/Chicago', NC: 'America/New_York', NH: 'America/New_York', NJ: 'America/New_York',
  NV: 'America/Los_Angeles', NY: 'America/New_York', OH: 'America/New_York', OK: 'America/Chicago',
  OR: 'America/Los_Angeles', PA: 'America/New_York', RI: 'America/New_York', SC: 'America/New_York',
  TX: 'America/Chicago', UT: 'America/Denver', VA: 'America/New_York', WA: 'America/Los_Angeles',
  WV: 'America/New_York', DC: 'America/New_York', HI: 'Pacific/Honolulu'
};
// Accepted model-supplied zones. Utilities.formatDate silently treats an unknown zone as GMT.
const TOLL_KNOWN_TZ_ = Object.keys(TOLL_STATE_TZ_).map(function (k) { return TOLL_STATE_TZ_[k]; }).concat([
  'America/Boise', 'America/Anchorage', 'America/Kentucky/Louisville', 'America/Indiana/Knox',
  'America/Menominee', 'America/North_Dakota/Center'
]);

// Toll program Envoy needs a car enrolled in, per toll-road state, in the vocabulary
// operators already use in "Agency Status" (CA FasTrak/express lanes = "The Toll Roads"
// account; FL/VA/NE corridor = "EZ Pass").
const TOLL_PROGRAM_BY_STATE_ = {
  CA: 'The Toll Roads', WA: 'GoodToGo', IL: 'iPass', CO: 'ExpressTolls', OK: 'PikePass',
  TX: 'TxTag', KS: 'K-TAG',
  DE: 'EZ Pass', DC: 'EZ Pass', FL: 'EZ Pass', IN: 'EZ Pass', KY: 'EZ Pass', MA: 'EZ Pass',
  MD: 'EZ Pass', ME: 'EZ Pass', MN: 'EZ Pass', NC: 'EZ Pass', NH: 'EZ Pass', NJ: 'EZ Pass',
  NY: 'EZ Pass', OH: 'EZ Pass', PA: 'EZ Pass', RI: 'EZ Pass', VA: 'EZ Pass', WV: 'EZ Pass',
  GA: 'EZ Pass'
};

/** Uppercase 2-letter state for the toll road, or '' (accepts "Virginia"). */
function tollStateCode_(v) {
  const s = String((v && v.tollState) || '').trim();
  if (!s) return '';
  const abbr = abbreviateState_(s);
  return /^[A-Z]{2}$/.test(abbr) ? abbr : '';
}

/** Local IANA zone of the toll: model's zone if known, else the toll state's zone, else null. */
function resolveTollTimeZone_(v) {
  const tz = String((v && v.tollTimeZone) || '').trim();
  if (TOLL_KNOWN_TZ_.indexOf(tz) !== -1) return tz;
  return TOLL_STATE_TZ_[tollStateCode_(v)] || null;
}

/** Toll program (Agency Status vocabulary) for a violation, or ''. */
function resolveTollProgram_(v) {
  const hay = [v && v.issuingAgency, v && v.originalCreditor, v && v.paymentWebsite, v && v.violationLocation]
    .filter(Boolean).join(' ').toLowerCase();
  if (/e-?z ?pass|ezdrive/.test(hay)) return 'EZ Pass';
  if (/i-?pass|illinois tollway/.test(hay)) return 'iPass';
  if (/good ?to ?go|wsdot/.test(hay)) return 'GoodToGo';
  if (/fastrak|thetollroads|toll roads|metro ?expresslanes|405 express|golden gate/.test(hay)) return 'The Toll Roads';
  if (/e-?470|expresstoll|northwest parkway/.test(hay)) return 'ExpressTolls';
  if (/colorado transportation investment/.test(hay)) return 'ExpressTolls';
  if (/fort bend|harris county|hctra|ntta|txtag|central texas regional/.test(hay)) return 'TxTag';
  if (/kansas turnpike|drive ?ks|k-?tag/.test(hay)) return 'K-TAG';
  if (/oklahoma turnpike|pikepass/.test(hay)) return 'PikePass';
  if (/dulles|i-?66|transurban|express lanes virginia|virginia toll|95 express|495 express|new jersey|nj turnpike|mta b|port authority|thruway|massachusetts|ezdrivema|delaware|de toll|rhode island|nc quick pass/.test(hay)) return 'EZ Pass';
  return TOLL_PROGRAM_BY_STATE_[tollStateCode_(v)] || '';
}

/** True for agency-fee line items (no trip), which are paid by accounting, not billed as tolls. */
function isAgencyFeeOnlyLine_(v) {
  const id = String((v && v.violationId) || '');
  const type = String((v && v.violationType) || '');
  // "_FEE_n" ids come from the prompt's fee-line convention; "Toll Violation Fee",
  // "Courtesy Violation Fee", "Mailed Statement Fee" are fees, "Toll Charge" is not.
  return /_FEE(_|$)/i.test(id) || /\bfees?\b/i.test(type);
}

/**
 * "Type" on current_tolls is what the driver sees in the toll email, so a bare gantry
 * code ("Route 410 411") gets the agency in front: "I-66 Express - Route 410 411".
 */
function tollTypeLabel_(v) {
  const loc = formatLocation_(String((v && v.violationLocation) || '')) || '';
  const agency = String((v && (v.originalCreditor || v.issuingAgency)) || '')
    .replace(/\b(mobility partners|partners|llc|inc\.?|authority|operations)\b/ig, '').replace(/\s+/g, ' ').trim();
  if (!loc) return agency || String((v && v.violationType) || '');
  const firstWord = (agency.split(' ')[0] || '').toLowerCase();
  if (!agency || (firstWord && loc.toLowerCase().indexOf(firstWord) !== -1)) return loc;
  return `${agency} - ${loc}`;
}

/** Plate with any leading/trailing state code removed, uppercase. */
function bareTollPlate_(plate) {
  let p = String(plate || '').trim().toUpperCase();
  p = p.replace(/[\s\-_/.]+([A-Z]{2})$/, function (m, code) { return VALID_STATE_CODES_[code] ? '' : m; });
  p = p.replace(/^([A-Z]{2})[\s\-_/.]+/, function (m, code) { return VALID_STATE_CODES_[code] ? '' : m; });
  return p.replace(/\s+/g, '');
}

/**
 * Local-time Date for a violation, built in the TOLL's zone. Returns
 * { instant: Date|null, tz: string|null, localYmdHms: string|null }.
 */
function tollInstant_(parsedDate, parsedTime, tz) {
  if (!parsedDate || !parsedTime) return { instant: null, tz: tz, localYmdHms: null };
  const scriptTz = Session.getScriptTimeZone();
  const localYmdHms = Utilities.formatDate(parsedDate, scriptTz, 'yyyy-MM-dd') + ' ' +
                      Utilities.formatDate(parsedTime, scriptTz, 'HH:mm:ss');
  if (!tz) return { instant: null, tz: null, localYmdHms: localYmdHms };
  return { instant: Utilities.parseDate(localYmdHms, tz, 'yyyy-MM-dd HH:mm:ss'), tz: tz, localYmdHms: localYmdHms };
}

/** Header-name → 0-based index map for row 1 of a sheet (trimmed, case-insensitive). */
function headerIndexMap_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const hdr = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  hdr.forEach(function (h, i) {
    const k = String(h || '').replace(/\s+/g, ' ').trim().toLowerCase();
    if (k && !(k in map)) map[k] = i;
  });
  return { map: map, width: hdr.length, headers: hdr };
}
function headerIdx_(hm, name) {
  const k = String(name).replace(/\s+/g, ' ').trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(hm.map, k) ? hm.map[k] : -1;
}

/** Ensures a header exists on row 1; appends it after the last header if missing. Returns 0-based index. */
function ensureHeader_(sheet, name) {
  let hm = headerIndexMap_(sheet);
  let idx = headerIdx_(hm, name);
  if (idx !== -1) return idx;
  let last = hm.headers.length - 1;
  while (last >= 0 && !String(hm.headers[last] || '').trim()) last--;
  idx = last + 1;
  if (idx + 1 > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), idx + 1 - sheet.getMaxColumns());
  sheet.getRange(1, idx + 1).setValue(name).setFontWeight('bold');
  return idx;
}

/**
 * Decide where a toll goes. Returns { route: boolean, billingStatus: string, reason: string }.
 * Only tolls with a usable date AND time and a real trip are billable to a driver.
 */
function tollRoutingDecision_(v, parsedDate, parsedTime) {
  if (v && v.isCollectionsNotice === true) {
    return { route: false, billingStatus: BILLING_STATUS_COLLECTIONS,
             reason: 'collections notice: no toll time on the letter, so no driver can be identified' };
  }
  if (isAgencyFeeOnlyLine_(v)) {
    return { route: false, billingStatus: BILLING_STATUS_FEE_ONLY,
             reason: 'agency fee line item (accounting pays it; toll pipeline never bills agency fees)' };
  }
  if (!parsedDate || !parsedTime) {
    return { route: false, billingStatus: BILLING_STATUS_NO_TIME,
             reason: 'no usable toll date/time on the notice, so no driver can be identified' };
  }
  return { route: true, billingStatus: BILLING_STATUS_TRANSFERRED, reason: '' };
}

/**
 * Appends one toll to toll-processing!current_tolls in whatever layout its header row
 * has (normally processed_tolls). Skips duplicates already on the sheet.
 * @returns {{status: 'appended'|'duplicate', row: number}}
 */
function routeTollToCurrentTolls_(violationData, parsedDate, parsedTime, dbInfo, vehicleNumber, pdfLink) {
  const functionName = 'routeTollToCurrentTolls_';
  const ss = SpreadsheetApp.openById(TOLL_PROCESSING_SHEET_ID);
  const sheet = ss.getSheetByName(TOLL_CURRENT_SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${TOLL_CURRENT_SHEET_NAME}" not found in toll-processing spreadsheet`);

  const notesIdx = ensureHeader_(sheet, CT_HEADERS_.NOTES);
  const hm = headerIndexMap_(sheet);
  const need = [CT_HEADERS_.ENVOY, CT_HEADERS_.LOCAL, CT_HEADERS_.TOLL];
  need.forEach(function (h) {
    if (headerIdx_(hm, h) === -1) throw new Error(`current_tolls is missing the "${h}" header; refusing to write a misaligned row.`);
  });

  const tz = resolveTollTimeZone_(violationData);
  const t = tollInstant_(parsedDate, parsedTime, tz);
  const plate = bareTollPlate_(violationData.licensePlate);
  const envoyValue = (vehicleNumber && vehicleNumber !== '#N/A') ? String(vehicleNumber) : plate;

  let first = '', last = '';
  if (dbInfo && dbInfo.driverName) {
    const parts = String(dbInfo.driverName).trim().split(/\s+/);
    first = parts.shift() || '';
    last = parts.join(' ');
  }
  const tollAmount = parseFloat(violationData.originalPenaltyAmount) || 0;
  const adminFee = calculateTollAdminFee_(tollAmount);
  const penaltyAmt = parseFloat(violationData.additionalPenaltyAmount) || 0;

  const localDisplay = t.instant ? Utilities.formatDate(t.instant, t.tz, CT_LOCAL_FORMAT_) : (t.localYmdHms || '');
  const datePst = t.instant ? Utilities.formatDate(t.instant, CT_PST_ZONE_, CT_DATE_PST_FORMAT_) : '';
  const timePst = t.instant ? Utilities.formatDate(t.instant, CT_PST_ZONE_, CT_TIME_PST_FORMAT_) : '';

  const note = [
    'From ticket PDF',
    `violationId=${violationData.violationId || 'N/A'}`,
    violationData.tripId ? `trip=${violationData.tripId}` : null,
    `agency=${violationData.issuingAgency || 'N/A'}`,
    `plate=${plate || 'N/A'}${violationData.licenseState ? ' ' + abbreviateState_(violationData.licenseState) : ''}`,
    t.tz ? `local tz=${t.tz}` : 'local tz UNKNOWN (PST not computed)',
    penaltyAmt > 0 ? `agency fees billed to accounting=$${penaltyAmt.toFixed(2)}` : null,
    'needs agency review',
    pdfLink ? `PDF=${pdfLink}` : null
  ].filter(Boolean).join(' · ');

  // --- Duplicate guard: same violation id, or same car + same PST minute + same amount. ---
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const body = sheet.getRange(2, 1, lastRow - 1, hm.width).getDisplayValues();
    const idToken = `violationId=${violationData.violationId || ''} `;
    const iEnvoy = headerIdx_(hm, CT_HEADERS_.ENVOY), iD = headerIdx_(hm, CT_HEADERS_.DATE_PST),
          iT = headerIdx_(hm, CT_HEADERS_.TIME_PST), iTo = headerIdx_(hm, CT_HEADERS_.TOLL);
    for (let r = 0; r < body.length; r++) {
      const row = body[r];
      const n = String(row[notesIdx] || '') + ' ';
      const sameId = violationData.violationId && n.indexOf(idToken) !== -1;
      const sameTrip = datePst && iD !== -1 && iT !== -1 &&
        String(row[iEnvoy]).trim() === envoyValue && row[iD] === datePst &&
        String(row[iT]).replace(/^0/, '') === timePst.replace(/^0/, '') &&
        Math.abs((parseFloat(String(row[iTo]).replace(/[$,]/g, '')) || 0) - tollAmount) < 0.005;
      if (sameId || sameTrip) {
        logToSheet_(functionName, `Skipped duplicate: Viol ID ${violationData.violationId} already on current_tolls row ${r + 2}.`, 'WARN');
        return { status: 'duplicate', row: r + 2 };
      }
    }
  }

  const row = new Array(hm.width).fill('');
  const put = function (header, value) {
    const i = headerIdx_(hm, header);
    if (i !== -1) row[i] = value;
  };
  put(CT_HEADERS_.FIRST, first);
  put(CT_HEADERS_.LAST, last);
  put(CT_HEADERS_.EMAIL, (dbInfo && dbInfo.driverEmail) || '');
  put(CT_HEADERS_.ENVOY, envoyValue);
  put(CT_HEADERS_.BOOKING, (dbInfo && dbInfo.bookingId) || '');
  put(CT_HEADERS_.DATE_PST, datePst);
  put(CT_HEADERS_.TIME_PST, timePst);
  put(CT_HEADERS_.LOCAL, localDisplay);
  put(CT_HEADERS_.TYPE, tollTypeLabel_(violationData));
  put(CT_HEADERS_.TOLL, tollAmount);
  put(CT_HEADERS_.ADMIN, adminFee);
  put(CT_HEADERS_.TOTAL, Math.round((tollAmount + adminFee) * 100) / 100);
  row[notesIdx] = note;

  const target = sheet.getLastRow() + 1;
  if (target > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), target - sheet.getMaxRows());
  // Text-format the id/date/time cells first so Sheets doesn't turn "08/10/2026" into a date serial.
  [CT_HEADERS_.ENVOY, CT_HEADERS_.BOOKING, CT_HEADERS_.DATE_PST, CT_HEADERS_.TIME_PST, CT_HEADERS_.LOCAL].forEach(function (h) {
    const i = headerIdx_(hm, h);
    if (i !== -1) sheet.getRange(target, i + 1).setNumberFormat('@');
  });
  sheet.getRange(target, 1, 1, row.length).setValues([row]);
  [CT_HEADERS_.TOLL, CT_HEADERS_.ADMIN, CT_HEADERS_.TOTAL].forEach(function (h) {
    const i = headerIdx_(hm, h);
    if (i !== -1) sheet.getRange(target, i + 1).setNumberFormat('$#,##0.00');
  });
  SpreadsheetApp.flush();
  logToSheet_(functionName, `Appended toll to current_tolls row ${target} (Envoy # ${envoyValue}, local ${localDisplay}, PST ${datePst} ${timePst}, $${tollAmount.toFixed(2)})`, 'SUCCESS');
  return { status: 'appended', row: target };
}

// ====================================================================
// Toll agency verification (Tolls 2026 tab)
// ====================================================================

/** Cached roster from the Transponder Index: { byEnvoy:{}, byPlate:{}, error:string|null }. */
function loadTransponderRoster_() {
  if (loadTransponderRoster_.cache) return loadTransponderRoster_.cache;
  const out = { byEnvoy: {}, byPlate: {}, error: null };
  try {
    const sheet = SpreadsheetApp.openById(TRANSPONDER_INDEX_SPREADSHEET_ID).getSheetByName(TRANSPONDER_INDEX_SHEET_NAME);
    if (!sheet) throw new Error(`tab "${TRANSPONDER_INDEX_SHEET_NAME}" not found`);
    const data = sheet.getDataRange().getDisplayValues();
    const hdr = (data[0] || []).map(function (h) { return String(h).trim().toLowerCase(); });
    const find = function (re, dflt) { const i = hdr.findIndex(function (h) { return re.test(h); }); return i === -1 ? dflt : i; };
    // Same defaults toll-processing uses: A Transponder, B License Plate, C Envoy #.
    const iTag = find(/transponder|tag/, 0), iPlate = find(/plate/, 1), iEnvoy = find(/envoy/, 2);
    for (let r = 1; r < data.length; r++) {
      const tag = String(data[r][iTag] || '').trim();
      const plate = bareTollPlate_(data[r][iPlate]);
      const envoy = String(data[r][iEnvoy] || '').trim();
      if (!tag && !plate) continue;
      if (envoy && !out.byEnvoy[envoy]) out.byEnvoy[envoy] = tag || plate;
      if (plate && !out.byPlate[plate]) out.byPlate[plate] = tag || plate;
    }
  } catch (e) {
    out.error = String(e && e.message || e);
  }
  loadTransponderRoster_.cache = out;
  return out;
}

/**
 * Is this car on the toll account it needs for this toll?
 * @returns {{program:string, status:string, action:string}} status ∈ TOLLS_VERIFY_VALUES
 */
function verifyVehicleTollAccount_(envoyNumber, plate, program) {
  const envoy = String(envoyNumber || '').trim();
  const bare = bareTollPlate_(plate);
  if (!envoy || envoy === '#N/A') {
    return { program: program, status: 'Unknown vehicle',
             action: `Plate ${bare || '?'} not in fleet sheet; confirm the car before adding it to ${program || 'a toll account'}` };
  }
  if (program !== 'EZ Pass') {
    return { program: program, status: 'Unverified',
             action: program ? `Check Envoy ${envoy} (${bare}) is on the ${program} account` : 'Set the toll program in Agency Status' };
  }
  const roster = loadTransponderRoster_();
  if (roster.error) {
    return { program: program, status: 'Unverified', action: `Could not read Transponder Index (${roster.error}); check EZ Pass account by hand` };
  }
  const tag = roster.byEnvoy[envoy] || roster.byPlate[bare];
  if (tag) {
    return { program: program, status: 'Yes',
             action: `On EZ Pass (transponder ${tag}) yet still violated: check tag is mounted, plate on account, balance funded` };
  }
  return { program: program, status: 'No - needs adding', action: `Add Envoy ${envoy} (${bare}) to the EZ Pass account` };
}

/**
 * Writes Agency Status / Vehicle on Agency Account? / Action Required on a Tolls 2026 row.
 * Never overwrites a value an operator already typed (only fills blanks).
 */
function writeTollAgencyCells_(sheet, rowIndex, result) {
  const verifyIdx = ensureHeader_(sheet, TOLLS_VERIFY_HEADER);
  const hm = headerIndexMap_(sheet);
  const cells = [
    [headerIdx_(hm, TOLLS_AGENCY_HEADER), result.program],
    [verifyIdx, result.status],
    [headerIdx_(hm, TOLLS_ACTION_HEADER), result.status === 'Unverified' && !result.program ? '' : result.action]
  ];
  cells.forEach(function (c) {
    if (c[0] === -1 || !c[1]) return;
    const rng = sheet.getRange(rowIndex, c[0] + 1);
    if (String(rng.getValue() || '').trim() === '') rng.setValue(c[1]);
  });
}

/**
 * Menu: fill Agency Status + "Vehicle on Agency Account?" for every Tolls 2026 row where
 * the verification cell is blank. Safe to re-run; never overwrites typed values.
 */
function verifyTollAgencyAssignments() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOLL_TARGET_SHEET_NAME);
  if (!sheet) { ui.alert(`Sheet "${TOLL_TARGET_SHEET_NAME}" not found.`); return; }
  const verifyIdx = ensureHeader_(sheet, TOLLS_VERIFY_HEADER);
  ensureTollVerifyValidation_(sheet, verifyIdx);
  const hm = headerIndexMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('No toll rows.'); return; }
  const data = sheet.getRange(2, 1, lastRow - 1, hm.width).getDisplayValues();
  const iAgencyStatus = headerIdx_(hm, TOLLS_AGENCY_HEADER);
  const counts = {};
  let done = 0;
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (String(row[verifyIdx] || '').trim()) continue;
    if (!String(row[COL.LICENSE_PLATE - 1] || '').trim() && !String(row[COL.VEHICLE - 1] || '').trim()) continue;
    // An operator-typed Agency Status wins over the derived program ("EZ Pass - Cannot Add..." → EZ Pass).
    const typed = iAgencyStatus !== -1 ? String(row[iAgencyStatus] || '').trim() : '';
    let program = typed.replace(/\s*-.*$/, '');
    if (!program) {
      program = resolveTollProgram_({
        issuingAgency: row[COL.ISSUING_AGENCY - 1], violationLocation: row[COL.VIOLATION_LOCATION - 1],
        paymentWebsite: row[COL.NOTES - 1], tollState: ''
      });
    }
    let res;
    if (/cannot add/i.test(typed)) res = { program: program, status: 'Cannot add - on another account', action: '' };
    else if (/unknown vehicle/i.test(typed)) res = { program: program, status: 'Unknown vehicle', action: '' };
    else res = verifyVehicleTollAccount_(row[COL.VEHICLE - 1], row[COL.LICENSE_PLATE - 1], program);
    writeTollAgencyCells_(sheet, r + 2, res);
    counts[res.status] = (counts[res.status] || 0) + 1;
    done++;
  }
  const summary = Object.keys(counts).map(function (k) { return `${k}: ${counts[k]}`; }).join('\n');
  logToSheet_('verifyTollAgencyAssignments', `Verified ${done} rows. ${summary.replace(/\n/g, '; ')}`, 'INFO');
  ui.alert(`Toll agency check filled ${done} row(s).\n\n${summary || 'Nothing to fill.'}`);
}

/** Dropdown on the verification column (allows typed values too, so nothing is rejected). */
function ensureTollVerifyValidation_(sheet, verifyIdx) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(TOLLS_VERIFY_VALUES, true).setAllowInvalid(true).build();
  sheet.getRange(2, verifyIdx + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}
