/**
 * Manual booking attribution for an existing citation/toll row.
 *
 * Use case: during a batch upload the MySQL booking lookup can come back empty
 * (no booking found around the violation date, unparsable date/time from the
 * scan, vehicle # not yet in the Fleet sheet, or a transient DB error). Those
 * rows land with a blank Booking ID (col H) and no Responsible Driver / Driver
 * Email / Property. This tool re-runs getDriverBookingInfo_() against the live
 * Envoy DB for the row the user has selected and backfills those fields.
 *
 * Operates on the active row of the active "Envoy Violations - New 2026" or
 * "Envoy Violations - Tolls 2026" tab (both share the same 29-column schema).
 */
function manualBookingLookupForRow() {
  const functionName = "manualBookingLookupForRow";
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // --- Guard: must be one of the violation tabs (shared COL schema) ---
  const sheetName = sheet.getName();
  if (sheetName !== TARGET_SHEET_NAME && sheetName !== TOLL_TARGET_SHEET_NAME) {
    ui.alert(
      "Wrong sheet",
      `Select a row on "${TARGET_SHEET_NAME}" or "${TOLL_TARGET_SHEET_NAME}" first, ` +
      `then run this tool. (Active sheet is "${sheetName}".)`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- Resolve the target row from the active selection ---
  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    ui.alert("Pick a data row", "Row 1 is the header. Click any cell in the citation row you want to look up, then run this tool again.", ui.ButtonSet.OK);
    return;
  }

  // Read display strings so date (MM/dd/yyyy) and time (h:mm a) come back in the
  // exact formats parseFlexibleDate_/parseFlexibleTime_ already understand,
  // regardless of the cell's underlying value type.
  const lastCol = Math.min(sheet.getLastColumn(), NUM_COLUMNS);
  const display = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  const cell = function (col) { return (col <= lastCol) ? (display[col - 1] || '').toString().trim() : ''; };

  const vehicleRaw = cell(COL.VEHICLE);
  const dateStr = cell(COL.DATE_VIOLATION);
  const timeStr = cell(COL.TIME_VIOLATION);
  const existingBookingId = cell(COL.BOOKING_ID);

  // --- Validate vehicle ---
  const vehicleNumber = vehicleRaw.replace(/^envoy\s+/i, '').trim();
  if (!vehicleNumber || vehicleNumber === '#N/A') {
    ui.alert("Missing vehicle #", `Row ${row} has no usable vehicle number in column A (found "${vehicleRaw}"). Fill the Envoy # first, then retry.`, ui.ButtonSet.OK);
    return;
  }

  // --- Validate / parse date & time ---
  const parsedDate = parseFlexibleDate_(dateStr);
  const parsedTime = parseFlexibleTime_(timeStr, parsedDate);
  if (!parsedDate || !parsedTime) {
    ui.alert(
      "Unparsable date/time",
      `Row ${row}: could not read a valid violation date AND time ` +
      `(date col D = "${dateStr}", time col F = "${timeStr}"). ` +
      `Booking attribution is time-of-day specific, so both are required. Fix them and retry.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- Warn before clobbering an existing attribution ---
  if (existingBookingId) {
    const resp = ui.alert(
      "Booking ID already present",
      `Row ${row} already has Booking ID "${existingBookingId}". ` +
      `Re-run the lookup and overwrite the driver/booking/property fields?`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      return;
    }
  }

  // --- Run the same DB lookup the batch upload uses ---
  const violationDateTimeLocalStr =
    Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd") + " " +
    Utilities.formatDate(parsedTime, Session.getScriptTimeZone(), "HH:mm:ss");
  logToSheet_(functionName, `Manual lookup for row ${row}: vehicle ${vehicleNumber}, time ${violationDateTimeLocalStr}`, "INFO");

  let dbInfo;
  try {
    dbInfo = getDriverBookingInfo_(vehicleNumber, violationDateTimeLocalStr);
  } catch (e) {
    logToSheet_(functionName, `Lookup threw for row ${row}: ${e.toString()}`, "ERROR");
    ui.alert("Lookup failed", `Database lookup threw an error:\n\n${e.message || e}`, ui.ButtonSet.OK);
    return;
  }

  if (!dbInfo || dbInfo.error) {
    const msg = dbInfo ? dbInfo.error : "No details returned.";
    logToSheet_(functionName, `No match for row ${row}: ${msg}`, "WARN");
    ui.alert(
      "No matching booking",
      `Envoy ${vehicleNumber} at ${violationDateTimeLocalStr}:\n\n${msg}\n\n` +
      `Nothing was written. The row may genuinely have no booking covering the violation time.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- Backfill the attribution fields ---
  sheet.getRange(row, COL.RESPONSIBLE_DRIVER).setValue(dbInfo.driverName || '');
  sheet.getRange(row, COL.DRIVER_EMAIL).setValue(dbInfo.driverEmail || '');
  sheet.getRange(row, COL.BOOKING_ID).setValue(dbInfo.bookingId || '');
  sheet.getRange(row, COL.PROPERTY).setValue(dbInfo.propertyName || '');

  // Leave a breadcrumb in Notes so it's clear this was attributed after upload.
  try {
    const notesCell = sheet.getRange(row, COL.NOTES);
    const existingNotes = (notesCell.getDisplayValue() || '').trim();
    const stamp = "Booking manually attributed " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (existingNotes.indexOf("Booking manually attributed") === -1) {
      notesCell.setValue(existingNotes ? existingNotes + "; " + stamp : stamp);
    }
  } catch (noteErr) {
    logToSheet_(functionName, `Could not stamp Notes for row ${row}: ${noteErr}`, "WARN");
  }

  SpreadsheetApp.flush();
  logToSheet_(functionName, `Row ${row} attributed to booking ${dbInfo.bookingId} (${dbInfo.driverName}, ${dbInfo.propertyName}).`, "SUCCESS");

  ui.alert(
    "Booking attributed",
    `Row ${row} updated:\n\n` +
    `Booking ID: ${dbInfo.bookingId}\n` +
    `Driver: ${dbInfo.driverName}\n` +
    `Email: ${dbInfo.driverEmail}\n` +
    `Property: ${dbInfo.propertyName}`,
    ui.ButtonSet.OK
  );
}
