/**
 * @OnlyCurrentDoc // Limits script authorization to the current document.
 */

// ====================================================================
// Constants & Configuration
// ====================================================================
const TARGET_SHEET_NAME = "Envoy Violations - New 2026";
const FLEET_SHEET_ID = "1E5d_YhQAQVm52XEI6dH74U1n_YGtwpjLvPmm57oDbHE";
const FLEET_SHEET_TAB_NAME = "Entire Fleet Overview";
const TICKET_DRIVE_FOLDER_NAME = "Traffic Tickets"; // Or your preferred folder name
const LOG_SHEET_NAME = "TicketProcessingLogs"; // Dedicated log sheet

// Cross-project toll routing
const TOLL_PROCESSING_SHEET_ID = "1NLp7cqIY9At7TOCDq7cQ_ucH-rC6xADaXYvt3Jv5tGg";
const TOLL_CURRENT_SHEET_NAME = "current_tolls";

// New-row notification recipient (per-row email on every upload)
const NEW_ROW_NOTIFY_EMAIL = "sky@envoythere.com";

// Column Mapping for TARGET_SHEET_NAME (1-based index) — matches "Envoy Violations - New 2026" layout
const COL = {
  VEHICLE: 1,                 // A: Envoy #
  LICENSE_PLATE: 2,           // B: License Plate
  LICENSE_PLATE_STATE: 3,     // C: State
  DATE_VIOLATION: 4,          // D: Date
  DUE_DATE: 5,                // E: Due Date
  TIME_VIOLATION: 6,          // F: Time (check)
  VIOLATION_ID: 7,            // G: Violation ID
  BOOKING_ID: 8,              // H: Booking ID
  PIN_NUMBER: 9,              // I: PIN Number
  ISSUING_AGENCY: 10,         // J: Issuing Agency
  VIOLATION_TYPE: 11,         // K: Violation Type
  VIOLATION_LOCATION: 12,     // L: Violation Location
  RESPONSIBLE_DRIVER: 13,     // M: Responsible Driver
  DRIVER_EMAIL: 14,           // N: Driver Email
  EMAIL_STATUS: 15,           // O: Driver Email Status
  DRIVER_BILLING_STATUS: 16,  // P: Driver Billing Status
  PROPERTY: 17,               // Q: Property
  ORIGINAL_PENALTY: 18,       // R: Original (Billable) Amount
  ADDITIONAL_PENALTY: 19,     // S: Penalty Amount
  PAYABLE_AMOUNT: 20,         // T: Payable Amount
  TOTAL_OWED_COLLECTIONS: 21, // U: Balance Due
  TOLL_OR_TICKET: 22,         // V: Toll or Ticket (now set by script based on Gemini isToll)
  // W: Subsidiary (not set by script)
  // X: Recorded in bill.com (Accounting)
  // Y: Paid by (Accounting)
  // Z: Dispute Status (Accounting)
  PAYMENT_STATUS: 27,         // AA: Payment Status (Sky)
  NOTES: 28,                  // AB: Notes
  PDF_LINK: 29                // AC: PDF
};
// IMPORTANT: Ensure your TARGET_SHEET_NAME actually has columns up to AC (29)
const NUM_COLUMNS = 29; // Updated to include up to column AC
// In TicketProcessingv2-NeedsDebugging.gs.txt

// In TicketProcessing.gs (THIS IS FILE B)

// --- START OF AUTHORITATIVE BATCH LOGGING SYSTEM ---
const LOG_BUFFER = [];
let BATCH_LOGGING_ACTIVE = false;
const MAX_BUFFERED_LOGS = 50;
// ── Drive folder cache ──
let TICKET_DRIVE_FOLDER = null;   // holds the DriveApp Folder object once found/created
let GEMINI_API_KEY_CACHE = null;
// LOG_SHEET_NAME is already a const in this file, so the logging functions will use it.
// In TicketProcessing.gs
function testCrossFileCall() {
  console.log("SUCCESS: testCrossFileCall from TicketProcessing.gs was reached!");
  Logger.log("SUCCESS: testCrossFileCall from TicketProcessing.gs was reached! (Logger.log)");
  // Optionally, try to use the BATCH_LOGGING_ACTIVE from this file:
  // console.log("Current BATCH_LOGGING_ACTIVE in TicketProcessing.gs: " + BATCH_LOGGING_ACTIVE);
  return "Test function in TicketProcessing.gs executed.";
}
/**
 * Enables or disables batch logging mode.
 * When disabling, it flushes any existing logs in the buffer.
 */
function setBatchLoggingEnabled(enabled) {
  const oldStatus = BATCH_LOGGING_ACTIVE;
  BATCH_LOGGING_ACTIVE = enabled;
  console.log(`Batch logging ${BATCH_LOGGING_ACTIVE ? 'ENABLED' : 'DISABLED'}. Old status: ${oldStatus}`);
  if (oldStatus && !enabled) {
    console.log("Flushing log buffer because batch logging was disabled.");
    flushLogBuffer_();
  }
}

/**
 * Writes all logs currently in LOG_BUFFER to the sheet.
 */
function flushLogBuffer_() {
  if (LOG_BUFFER.length === 0) {
    // console.log("Log buffer is empty, nothing to flush."); // Optional: for debugging
    return;
  }
  console.log(`Flushing ${LOG_BUFFER.length} log entries.`);
  _writeLogEntriesToSheet_(LOG_BUFFER.slice()); // Pass a copy
  LOG_BUFFER.length = 0; // Clear original buffer
}

/**
 * Internal helper to write an array of log entries to the sheet.
 * (This function seems to be your existing non-batch logToSheet_, adapted)
 * Ensure this is the version you want for writing the buffered logs.
 * The key is that it writes MULTIPLE entries at once if logEntriesArray has multiple.
 */
function _writeLogEntriesToSheet_(logEntriesArray) {
    if (!logEntriesArray || logEntriesArray.length === 0) return;

    // const logSheetName = LOG_SHEET_NAME; // This should be globally available in this file
    const maxLogRows = 1500;

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss) {
            logEntriesArray.forEach(logRow => Logger.log(`!!! SPREADSHEET INACCESSIBLE (BUFFERED) !!! [${logRow[1]}] ${logRow[2]}: ${logRow[3]}`));
            return;
        }
        let logSheet = ss.getSheetByName(LOG_SHEET_NAME); // Use the constant from this file

        if (!logSheet) {
            try {
                logSheet = ss.insertSheet(LOG_SHEET_NAME, 0);
                const headers = [["Timestamp", "Level", "Function", "Message"]];
                logSheet.getRange("A1:D1").setValues(headers).setFontWeight("bold").setBackground("#eeeeee");
                logSheet.setFrozenRows(1);
                logSheet.setColumnWidth(1, 180).setColumnWidth(2, 80).setColumnWidth(3, 200).setColumnWidth(4, 700);
            } catch (sheetCreateError) {
                logEntriesArray.forEach(logRow => Logger.log(`!!! ERROR CREATING LOG SHEET (BUFFERED) '${LOG_SHEET_NAME}': ${sheetCreateError} !!! [${logRow[1]}] ${logRow[2]}: ${logRow[3]}`));
                return;
            }
        }

        if (logSheet && logEntriesArray.length > 0) {
            logSheet.insertRowsBefore(2, logEntriesArray.length);
            const rangeToSet = logSheet.getRange(2, 1, logEntriesArray.length, 4);
            rangeToSet.setValues(logEntriesArray); // This writes all buffered logs at once

            for (let i = 0; i < logEntriesArray.length; i++) {
                const level = logEntriesArray[i][1];
                const levelCell = logSheet.getRange(2 + i, 2);
                if (level === "ERROR" || level === "FATAL") levelCell.setFontColor("#a50e0e").setFontWeight("bold");
                else if (level === "WARN") levelCell.setFontColor("#e67e22");
                else if (level === "SUCCESS") levelCell.setFontColor("#137333");
                else levelCell.setFontColor("black").setFontWeight("normal");
            }
            const lastRow = logSheet.getLastRow();
            if (lastRow > (maxLogRows + 1)) {
                const rowsToDelete = lastRow - (maxLogRows + 1);
                if (rowsToDelete > 0) logSheet.deleteRows(maxLogRows + 2, rowsToDelete);
            }
        }
    } catch (e) {
        logEntriesArray.forEach(logRow => Logger.log(`!!! SHEET LOGGING FAILED (_writeLogEntriesToSheet_) !!! [${logRow[1]}] ${logRow[2]}: ${logRow[3]}. Error: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : "")));
    }
}

// --- END OF AUTHORITATIVE BATCH LOGGING SYSTEM ---

// N.B. The actual logToSheet_ function defined later in your TicketProcessing.gs
// needs to be MODIFIED to use this batch system.

// In TicketProcessing.gs

/**
 * Logs messages to a dedicated sheet. If batch logging is active, messages are buffered.
 * THIS IS THE CANONICAL logToSheet_ that should be used by ALL functions.
 * @param {string} functionName The name of the calling function.
 * @param {string} message The message to log.
 * @param {string} [level='INFO'] The logging level (e.g., INFO, WARN, ERROR, DEBUG, FATAL, ENTRY).
 */
function logToSheet_(functionName, message, level) {
  level = (level || "INFO").toUpperCase();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss.SSS Z");
  let messageStr = String(message);
  if (messageStr.length > 49000) messageStr = messageStr.substring(0, 49000) + "... (truncated)";

  const logDataRow = [timestamp, level, functionName, messageStr];

  if (typeof BATCH_LOGGING_ACTIVE !== 'undefined' && BATCH_LOGGING_ACTIVE) {
    // console.log(`Batching log: [${level}] ${functionName}: ${messageStr.substring(0,100)}`); // For debugging the batching itself
    LOG_BUFFER.push(logDataRow);
    if (LOG_BUFFER.length >= MAX_BUFFERED_LOGS) {
      console.log("Max buffered logs reached, flushing buffer from logToSheet_.");
      flushLogBuffer_(); // Uses the flushLogBuffer_ defined at the top of this file
    }
    return;
  }
  // If not batch logging, write this single entry immediately.
  // console.log(`Immediate log: [${level}] ${functionName}: ${messageStr.substring(0,100)}`); // For debugging
  _writeLogEntriesToSheet_([logDataRow]); // Uses the _writeLogEntriesToSheet_ defined at the top
}

// ... (rest of your TicketProcessingv2-NeedsDebugging.gs.txt code, including uploadAndAnalyzeTicket, addMultipleTicketData, etc.)

// ====================================================================
// Menu Functions
// ====================================================================

function showAddTicketForm() {
  var html = HtmlService.createHtmlOutputFromFile('AddTicketForm')
      .setWidth(600) // Wider for more fields
      .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Parking/Traffic Ticket');
}

// ====================================================================
// File Upload and OCR Processing (Called from HTML)
// ====================================================================

// ====================================================================
// File Upload and OCR Processing (Called from HTML)
// ====================================================================

/**
 * Processes payments for tickets marked with 'Attempt' status in Column O (DRIVER_BILLING_STATUS).
 * **REVISED: Filters rows first, removed flush inside loop.**
 */
function processTicketPayments() {
  const functionName = "processTicketPayments";
  const ui = SpreadsheetApp.getUi();
  logToSheet_(functionName, "Starting payment processing run.", "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    ui.alert("Error: Target sheet '" + TARGET_SHEET_NAME + "' not found.");
    logToSheet_(functionName, "Sheet not found.", "ERROR");
    return;
  }

  const TEST_CUSTOMER_ID = "cus_QPOtpj7JZThQK1"; // Test customer ID for fee waiver

  // --- Column Indices ---
  const headerRows = 1;
  const billingStatusColIndex = COL.DRIVER_BILLING_STATUS - 1;
  const bookingIdColIndex = COL.BOOKING_ID - 1;
  const originalPenaltyColIndex = COL.ORIGINAL_PENALTY - 1;
  const totalOwedCol = COL.TOTAL_OWED_COLLECTIONS; // Column AA

  let successCount = 0;
  let collectionCount = 0;
  let errorCount = 0;

  // --- Get All Data Once ---
  const dataRange = sheet.getDataRange();
  const allData = dataRange.getValues();
  const numDataRows = allData.length - headerRows;
  logToSheet_(functionName, `Read ${allData.length} total rows (${numDataRows} data rows) from sheet.`, "DEBUG");

  // --- ** FILTER DATA FIRST ** ---
  const rowsToProcess = [];
  for (let i = headerRows; i < allData.length; i++) {
    const rowData = allData[i];
    if (!rowData || rowData.length <= billingStatusColIndex) {
        // logToSheet_(functionName, `Row ${i + 1}: Skipping pre-filter check - Row is invalid or too short.`, "WARN"); // Can be noisy
        continue;
    }
    const statusRaw = rowData[billingStatusColIndex];
    const statusProcessed = String(statusRaw).trim().toUpperCase();

    if (statusProcessed === "ATTEMPT") {
      rowsToProcess.push({ data: rowData, originalRowIndex: i + 1 });
    }
  }
  // --- ** End Filtering ** ---

  const numRowsToProcess = rowsToProcess.length;
  logToSheet_(functionName, `Found ${numRowsToProcess} row(s) with status 'Attempt' in Column ${String.fromCharCode(65 + billingStatusColIndex)}.`, "INFO");

  if (numRowsToProcess === 0) {
    logToSheet_(functionName, "No rows found requiring payment processing.", "INFO");
    ui.alert("No rows found with status 'Attempt'. Processing complete.");
    return;
  }

  // --- Get Stripe Key ---
  const stripeKey = getStripeSecretKey();
  if (!stripeKey) {
     ui.alert("Error: Could not retrieve Stripe Secret Key. Check logs and Secret Manager setup.");
     logToSheet_(functionName, "Stripe key retrieval failed.", "ERROR");
     return;
  }
  logToSheet_(functionName, "Stripe key retrieved successfully.", "INFO");

  let sheetUpdates = []; // Batch sheet updates

  // --- ** Process ONLY the Filtered Rows ** ---
  for (let j = 0; j < rowsToProcess.length; j++) {
    const item = rowsToProcess[j];
    const row = item.data;
    const actualRowNum = item.originalRowIndex;

    logToSheet_(functionName, `Processing filtered row ${j + 1} of ${numRowsToProcess} (Sheet Row ${actualRowNum}). Status: 'Attempt'.`, "INFO");

    const bookingId = row[bookingIdColIndex];
    const originalPenalty = parseFloat(row[originalPenaltyColIndex]) || 0;

    // Update status immediately visually (can remove if batching preferred)
    sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Processing Payment...");
    SpreadsheetApp.flush(); // Flush *only* for visual update here

    if (!bookingId || originalPenalty < 0) {
      logToSheet_(functionName, `Sheet Row ${actualRowNum}: Skipping payment - Missing Booking ID ('${bookingId}') or negative Original Penalty ($${originalPenalty}).`, "WARN");
      // Add update to batch
      sheetUpdates.push({ row: actualRowNum, col: COL.DRIVER_BILLING_STATUS, value: "Error - Missing/Invalid Data" });
      errorCount++;
      continue;
    }

    // Lookup Stripe Customer ID
    logToSheet_(functionName, `Sheet Row ${actualRowNum}: Looking up Stripe ID for Booking ID ${bookingId}...`, "DEBUG");
    const stripeCustomerId = getStripeCustomerIdFromBooking_(bookingId); // Assuming this doesn't need caching or is fast

    if (!stripeCustomerId || typeof stripeCustomerId !== 'string' || !stripeCustomerId.startsWith('cus_')) {
        logToSheet_(functionName, `Sheet Row ${actualRowNum}: Stripe Customer ID lookup failed or invalid for Booking ID ${bookingId}. Result: ${stripeCustomerId}`, "ERROR");
        sheetUpdates.push({ row: actualRowNum, col: COL.DRIVER_BILLING_STATUS, value: "Error - No Stripe ID" });

        let adminFeeOnFailure = 25.00;
        if (stripeCustomerId === TEST_CUSTOMER_ID) {
           adminFeeOnFailure = 0.00;
           logToSheet_(functionName, `Sheet Row ${actualRowNum}: Waiving $25 admin fee for TEST_CUSTOMER_ID during Stripe ID lookup failure.`, "INFO");
        }
        const totalOwed = originalPenalty + adminFeeOnFailure;

        if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
           sheetUpdates.push({ row: actualRowNum, col: totalOwedCol, value: totalOwed, format: "$#,##0.00"});
        } else {
           logToSheet_(functionName, `Sheet Row ${actualRowNum}: Invalid TOTAL_OWED_COLLECTIONS column index (${totalOwedCol}). Cannot stage owed amount update.`, "ERROR");
        }
        errorCount++;
        continue;
    }
    logToSheet_(functionName, `Sheet Row ${actualRowNum}: Found Stripe Customer ID: ${stripeCustomerId}`, "DEBUG");

    // Fee Waiver Logic
    let adminFee = 25.00;
    if (stripeCustomerId === TEST_CUSTOMER_ID) {
        adminFee = 0.00;
        logToSheet_(functionName, `Sheet Row ${actualRowNum}: Waiving $25 admin fee for test customer ${TEST_CUSTOMER_ID}.`, "INFO");
    }

    const amountToCharge = originalPenalty + adminFee;
    if (amountToCharge <= 0 && originalPenalty <= 0) {
       logToSheet_(functionName, `Sheet Row ${actualRowNum}: Calculated amount is $${amountToCharge.toFixed(2)}. Skipping charge attempt for non-positive amount. Marking as 'Paid (Amount $0)'.`, "WARN");
       sheetUpdates.push({ row: actualRowNum, col: COL.DRIVER_BILLING_STATUS, value: "Paid (Amount $0)" });
       if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
         sheetUpdates.push({ row: actualRowNum, col: totalOwedCol, value: 0, format: "$#,##0.00"});
       }
       successCount++;
       continue;
    }

    // Attempt Stripe Charge
    const amountInCents = Math.round(amountToCharge * 100);
    const description = `52065 Vehicle Toll Roads/Citations/Impound: Booking ${bookingId}`;
    logToSheet_(functionName, `Sheet Row ${actualRowNum}: Attempting Stripe charge for $${amountToCharge.toFixed(2)} (${amountInCents} cents)...`, "DEBUG");
    const paymentResult = attemptStripeCharge_(stripeKey, stripeCustomerId, amountInCents, description);

    // Stage Updates Based on Result
    if (paymentResult.success) {
      logToSheet_(functionName, `Sheet Row ${actualRowNum}: Payment successful. PI ID: ${paymentResult.paymentIntentId}`, "SUCCESS");
      sheetUpdates.push({ row: actualRowNum, col: COL.DRIVER_BILLING_STATUS, value: "Paid" });
      if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
        sheetUpdates.push({ row: actualRowNum, col: totalOwedCol, value: 0, format: "$#,##0.00"});
      }
      successCount++;
    } else {
      logToSheet_(functionName, `Sheet Row ${actualRowNum}: Payment failed. Reason: ${paymentResult.message}`, "WARN");
      sheetUpdates.push({ row: actualRowNum, col: COL.DRIVER_BILLING_STATUS, value: "Collections" });
      if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
          sheetUpdates.push({ row: actualRowNum, col: totalOwedCol, value: amountToCharge, format: "$#,##0.00"});
      } else {
           logToSheet_(functionName, `Sheet Row ${actualRowNum}: Invalid TOTAL_OWED_COLLECTIONS column index (${totalOwedCol}). Cannot stage owed amount update.`, "ERROR");
      }
      collectionCount++;
    }
     // --- REMOVED SpreadsheetApp.flush(); ---

  } // --- End filtered row loop ---

  // --- Apply Batch Updates ---
  if (sheetUpdates.length > 0) {
      logToSheet_(functionName, `Applying ${sheetUpdates.length} batch updates to the sheet...`, "INFO");
      sheetUpdates.forEach(update => {
          const range = sheet.getRange(update.row, update.col);
          range.setValue(update.value);
          if (update.format) {
              range.setNumberFormat(update.format);
          }
      });
      SpreadsheetApp.flush(); // Flush once after all updates are applied
      logToSheet_(functionName, "Batch updates applied.", "DEBUG");
  } else {
       logToSheet_(functionName, "No sheet updates needed.", "INFO");
  }


  // Final Summary
  logToSheet_(functionName, `Payment processing complete. Rows Found with 'Attempt': ${numRowsToProcess}, Successful Payments: ${successCount}, Sent to Collections: ${collectionCount}, Errors (Setup/Data/Lookup): ${errorCount}`, "INFO");
  ui.alert(`Payment Processing Complete.\n\nRows Found with 'Attempt': ${numRowsToProcess}\nSuccessful Payments: ${successCount}\nSent to Collections: ${collectionCount}\nErrors (Setup/Data/Lookup): ${errorCount}`);
}


// ====================================================================
// Google Drive Interaction
// ====================================================================
/**
 * Renames a file in Google Drive.
 * @param {string} fileId The ID of the file to rename.
 * @param {string} newFileName The desired new filename (including extension).
 * @returns {boolean} True if successful, false otherwise.
 */
function renameDriveFile_(fileId, newFileName) {
  const functionName = "renameDriveFile_";
  if (!fileId || !newFileName) {
    logToSheet_(functionName, "Missing fileId or newFileName for rename.", "ERROR");
    return false;
  }
  // Basic check for potentially invalid characters (though should be handled before calling)
  if (newFileName.includes('/') || newFileName.includes('\\')) {
      logToSheet_(functionName, `Invalid characters found in proposed new filename: "${newFileName}". Rename aborted.`, "ERROR");
      return false;
  }

  try {
    logToSheet_(functionName, `Attempting to rename File ID: ${fileId} to "${newFileName}"`, "INFO");
    const file = DriveApp.getFileById(fileId);
    file.setName(newFileName);
    logToSheet_(functionName, `Successfully renamed File ID: ${fileId} to "${newFileName}"`, "SUCCESS");
    return true;
  } catch (e) {
    logToSheet_(functionName, `Failed to rename File ID: ${fileId} to "${newFileName}". Error: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    // Specific check for "Document is missing" which means fileId was wrong
    if (e.message && e.message.includes("ocument is missing")) {
         logToSheet_(functionName, `Hint: Rename failed because File ID '${fileId}' was not found in Drive.`, "WARN");
    }
    return false;
  }
}
/**
 * Uploads a ticket file blob to a specific Google Drive folder.
 * @param {Blob} fileBlob The file blob to upload.
 * @returns {object} { fileUrl: string, fileId: string, initialFileName: string } // Corrected return type description
 * @throws {Error} If upload or permission setting fails.
 */
function uploadTicketToDrive_(fileBlob) {
  const functionName = "uploadTicketToDrive_";
  const folderName = TICKET_DRIVE_FOLDER_NAME;
  let fileUrl = null;
  let fileId = null;
  let initialFileName = null; // Declare the variable

  try {
    logToSheet_(functionName, `Starting upload to Drive folder: '${folderName}'`, "INFO");

    // ── use cached folder object if we already have it ──
let folder = TICKET_DRIVE_FOLDER;
if (!folder) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();               // existing folder
  } else {
    folder = DriveApp.createFolder(folderName); // first run: create it
  }
  TICKET_DRIVE_FOLDER = folder;            // cache for the rest of the batch
  logToSheet_(functionName,
    `Using folder '${folderName}' (ID: ${folder.getId()}) — cached for subsequent uploads.`,
    folders.hasNext() ? "DEBUG" : "INFO");
}


    const safeBaseName = (fileBlob.getName() || "Ticket").replace(/[^a-zA-Z0-9._-]/g, '_').replace(/\.pdf$/i, '');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
    const fileName = `${safeBaseName}_${timestamp}.pdf`;
    logToSheet_(functionName, `Generated initial filename: ${fileName}`, "DEBUG");

    // *** THIS IS THE MISSING LINE TO ADD BACK ***
    initialFileName = fileName; // Assign the generated filename to the variable
    // *********************************************

    const file = folder.createFile(fileBlob);
    file.setName(fileName); // Set name after creation using the 'fileName' variable
    fileId = file.getId();
    logToSheet_(functionName, `File created in Drive with ID: ${fileId} and Name: ${fileName}`, "INFO");

    // Set sharing: Anyone with the link can VIEW
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    logToSheet_(functionName, `File sharing set to Anyone with link can VIEW for file ID: ${fileId}`, "INFO");

    fileUrl = `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;
    logToSheet_(functionName, "Successfully uploaded. Returning URL, ID, and initial name.", "SUCCESS");

    // Now this return statement will work because initialFileName has been assigned a value
    return { fileUrl: fileUrl, fileId: fileId, initialFileName: initialFileName };

  } catch (e) {
    // If an error happens *before* initialFileName is assigned, it might still be null here.
    // The error message constructor handles this.
    logToSheet_(functionName, `Error uploading ticket: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    // Let the calling function handle the wrapping of the error message
    throw new Error(`Failed to upload ticket to Google Drive: ${e.message}`);
  }
}
/**
 * Handles file upload, saves to Drive, calls Gemini for OCR, renames the file using
 * Gemini-provided short name and violation ID (if available), and returns results.
 * @param {object} fileData Object containing name, mimeType, and base64 data.
 * @returns {object} { success: boolean, fileUrl: string|null, fileId: string|null, ocrData: object|null, message: string|null }
 */
function uploadAndAnalyzeTicket(fileData) {
  const functionName = "uploadAndAnalyzeTicket";
  // --- IMMEDIATE LOGGING ---
  logToSheet_(functionName, "Function entered.", "ENTRY"); // Use a distinct level if needed

  let fileUrl = null;
  let fileId = null;
  let ocrResult = null;
  let finalFileName = null; // Store the final name after potential rename

  try {
    // --- Log Input Data (moved inside try block) ---
    let initialLogMsg = "Received request to process ticket.";
    if (fileData && fileData.name) {
      initialLogMsg += " Name: " + fileData.name;
    } else {
      initialLogMsg += " (fileData or fileData.name seems invalid)";
    }
    logToSheet_(functionName, initialLogMsg, "INFO");
    // --- End Log Input Data ---


    if (!fileData || !fileData.data || !fileData.mimeType || !fileData.name) {
      throw new Error("Invalid fileData received.");
    }
    if (fileData.mimeType !== 'application/pdf') {
      throw new Error("Invalid file type. Only PDF is allowed.");
    }

    // 1. Upload to Drive (with initial timestamped name)
    logToSheet_(functionName, "Step 1: Uploading to Drive with temporary name...", "INFO");
    const decoded = Utilities.base64Decode(fileData.data);
    const blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.name);
    const uploadResult = uploadTicketToDrive_(blob); // Gets initial name
    fileUrl = uploadResult.fileUrl;
    fileId = uploadResult.fileId;
    finalFileName = uploadResult.initialFileName; // Store the initial name first
    logToSheet_(functionName, `Step 1 SUCCESS. File ID: ${fileId}, Initial Name: ${finalFileName}, URL: ${fileUrl}`, "SUCCESS");

    // 2. Analyze with Gemini
    logToSheet_(functionName, "Step 2: Analyzing with Gemini...", "INFO");
    ocrResult = analyzeTicketWithGemini_(fileId); // Handles empty violations array gracefully now

    if (!ocrResult.success) {
      // Gemini analysis itself failed (API error, parsing error, etc.)
      const errorMessage = ocrResult.message || "Gemini analysis failed.";
      logToSheet_(functionName, "Step 2 Failed. Gemini analysis issue: " + errorMessage, "ERROR");
      return {
        success: false,
        fileUrl: fileUrl,
        fileId: fileId,
        ocrData: null,
        message: `Upload OK (File: ${finalFileName}), but ticket analysis failed: ${errorMessage}`
      };
    }
    // Check if Gemini succeeded but found no violations/tolls
    if (ocrResult.data && ocrResult.data.violations && ocrResult.data.violations.length === 0) {
         logToSheet_(functionName, "Step 2 SUCCESS (No Violations Found). Gemini analysis complete, but no violations/tolls matched schema. Proceeding without rename.", "WARN");
         // Skip renaming, return success with the initial filename and empty OCR data for this specific part
         return {
             success: true,
             fileUrl: fileUrl,
             fileId: fileId,
             ocrData: ocrResult.data, // Return the data structure even if violations array is empty
             message: `Ticket processed. 0 violation(s)/toll(s) found matching schema. Filename: ${finalFileName}. (Rename skipped)`
         };
    }
    // If we reach here, Gemini analysis was successful AND found at least one violation/toll
    logToSheet_(functionName, "Step 2 SUCCESS. Gemini analysis complete with violations/tolls.", "SUCCESS");


    // --- ** Step 3: Construct New Filename and Rename (Only if violations exist) ** ---
    logToSheet_(functionName, "Step 3: Attempting to rename file based on OCR data...", "INFO");
    let renameSuccess = false;
    let attemptedNewName = null;
    try {
      // Use data from the *first* violation for naming convention
      // We already know ocrResult.data.violations has at least one element here
      const firstViolation = ocrResult.data.violations[0];
      const agencyShortName = firstViolation.agencyShortName;
      const violationIdRaw = firstViolation.violationId;

      if (agencyShortName && violationIdRaw) {
          const violationIdClean = String(violationIdRaw).replace(/[^A-Z0-9_-]/gi, ''); // Sanitize

          if (agencyShortName && violationIdClean) {
              attemptedNewName = `${agencyShortName}_${violationIdClean}.pdf`;
              logToSheet_(functionName, `Constructed new filename: "${attemptedNewName}"`, "DEBUG");
              renameSuccess = renameDriveFile_(fileId, attemptedNewName);
              if (renameSuccess) {
                finalFileName = attemptedNewName;
              }
          } else {
              logToSheet_(functionName, `Could not generate valid new filename (agencyShortName or cleaned ID was empty). Keeping initial name: ${finalFileName}`, "WARN");
          }
      } else {
          logToSheet_(functionName, `OCR data missing required fields for renaming (agencyShortName: '${agencyShortName}', violationId: '${violationIdRaw}'). Keeping initial name: ${finalFileName}`, "WARN");
      }
    } catch (renameError) {
       logToSheet_(functionName, `Error during filename construction or rename call: ${renameError}`, "ERROR");
       renameSuccess = false;
    }
    logToSheet_(functionName, `Step 3 ${renameSuccess ? 'SUCCESS' : 'SKIPPED/FAILED'}. File rename attempt finished. Final filename: ${finalFileName}`, renameSuccess ? "SUCCESS" : "WARN");
    // --- ** End Renaming Step ** ---

    // 4. Return combined results
    return {
      success: true,
      fileUrl: fileUrl,
      fileId: fileId,
      ocrData: ocrResult.data,
      message: `Ticket processed. ${ocrResult.data.violations.length} violation(s)/toll(s) found. Filename: ${finalFileName}.${renameSuccess ? '' : ' (Rename skipped/failed, see logs.)'}`
    };

  } catch (e) {
    // Log the error *before* returning
    const errorMsg = `Error during upload/analysis/rename: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : "");
    logToSheet_(functionName, errorMsg, "FATAL"); // Use FATAL level for caught errors

    // Return failure including file details if upload succeeded before error
    return {
      success: false,
      fileUrl: fileUrl,
      fileId: fileId,
      ocrData: null,
      message: "Failed to process ticket: " + e.message + (finalFileName ? ` (Initial file: ${finalFileName})` : "")
    };
  }
}
/**
 * Processes a batch of ticket data packages, adding each to the sheet.
 * @param {Array<object>} arrayOfDataPackages An array of dataPackage objects.
 *        Each dataPackage is { ticketPDF: string, ticketFileId: string, isLateNotice: boolean, violations: object[] }
 * @returns {object} A summary of the batch submission { total: number, successes: number, failures: number, messages: string[] }.
 */
function addMultipleTicketDataBatch(arrayOfDataPackages) {
  const functionName = "addMultipleTicketDataBatch";
  logToSheet_(functionName, `Batch submission started for ${arrayOfDataPackages.length} ticket packages.`, "INFO");
  setBatchLoggingEnabled(true); // Use batch logging for this potentially long operation

  let successCount = 0;
  let failureCount = 0;
  const overallMessages = [];

  if (!Array.isArray(arrayOfDataPackages) || arrayOfDataPackages.length === 0) {
    const msg = "No data packages received for batch submission.";
    logToSheet_(functionName, msg, "WARN");
    setBatchLoggingEnabled(false);
    return { total: 0, successes: 0, failures: 0, messages: [msg] };
  }

  arrayOfDataPackages.forEach((dataPackage, index) => {
    try {
      // Extract a unique identifier for logging, e.g., first violation ID or file name from PDF URL
      let packageIdentifier = `Package ${index + 1}`;
      if (dataPackage.ticketPDF) {
        try {
          packageIdentifier = dataPackage.ticketPDF.substring(dataPackage.ticketPDF.lastIndexOf('/') + 1);
        } catch (e) { /* ignore bad pdf url for identifier */ }
      } else if (dataPackage.violations && dataPackage.violations.length > 0 && dataPackage.violations[0].violationId) {
        packageIdentifier = `Viol. ID: ${dataPackage.violations[0].violationId}`;
      }
      
      logToSheet_(functionName, `Processing ${packageIdentifier} in batch submission.`, "DEBUG");
      
      // Call the existing single dataPackage processing function
      const resultMessage = addMultipleTicketData(dataPackage); 
      // addMultipleTicketData returns a string message, "X violation(s) submitted successfully."
      // We'll assume success if it doesn't throw an error and contains "successfully"
      if (resultMessage && resultMessage.toLowerCase().includes("successfully")) {
        successCount++;
        overallMessages.push(`${packageIdentifier}: ${resultMessage}`);
        logToSheet_(functionName, `Successfully processed ${packageIdentifier}: ${resultMessage}`, "SUCCESS");
      } else {
        failureCount++;
        overallMessages.push(`${packageIdentifier}: Failed or no success message (${resultMessage})`);
        logToSheet_(functionName, `Failed to process ${packageIdentifier}. Message: ${resultMessage}`, "WARN");
      }
    } catch (e) {
      failureCount++;
      const errorMsg = `Error processing package ${index + 1} in batch: ${e.message}`;
      logToSheet_(functionName, errorMsg + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
      overallMessages.push(errorMsg);
    }
    // Optional: Utilities.sleep(100) if many DB lookups cause issues, though unlikely needed here.
  });

  const summaryMessage = `Batch submission complete. Total Packages: ${arrayOfDataPackages.length}, Successes: ${successCount}, Failures: ${failureCount}.`;
  logToSheet_(functionName, summaryMessage, "INFO");
  setBatchLoggingEnabled(false);
  
  return {
    total: arrayOfDataPackages.length,
    successes: successCount,
    failures: failureCount,
    messages: overallMessages,
    summary: summaryMessage
  };
}
/**
 * Batch‑oriented wrapper around the existing single‑file pipeline.
 * Accepts an Array of the same `fileData` objects currently passed to
 * `uploadAndAnalyzeTicket()` and returns a parallel Array of result objects.
 */
function uploadAndAnalyzeTicketsBatch(filesData) {
  const functionName = "uploadAndAnalyzeTicketsBatch";

  // --- TEST CALL ---
  try {
    console.log("Attempting to call testCrossFileCall()...");
    const testResult = testCrossFileCall(); // Call the function from the other file
    console.log("Result from testCrossFileCall():", testResult);
    Logger.log("Result from testCrossFileCall() in Logger: " + testResult);
  } catch (e) {
    console.error("ERROR calling testCrossFileCall():", e.toString());
    Logger.log("ERROR calling testCrossFileCall(): " + e.toString());
    // If this errors, then setBatchLoggingEnabled will also error.
    // We should probably stop or handle this before proceeding.
    // For now, let's just log and let it try to call setBatchLoggingEnabled anyway.
  }
  // --- END TEST CALL ---

  setBatchLoggingEnabled(true); // Enable batch logging
  logToSheet_(functionName, `Batch processing started for ${filesData.length} files.`, "INFO");
  // ... rest of function

  if (!Array.isArray(filesData) || filesData.length === 0) {
    logToSheet_(functionName, "Received empty or invalid file array.", "ERROR");
    throw new Error("Batch endpoint expects a non‑empty Array of fileData objects.");
  }

  const summary = {
    startedAt: new Date().toISOString(),
    total: filesData.length,
    results: []
  };

  filesData.forEach((fileData, idx) => {
    let singleResult;
    logToSheet_(functionName, `Processing file ${idx + 1} of ${filesData.length}: ${fileData.name}`, "DEBUG");
    try {
      // Call the existing single-file processing function
      singleResult = uploadAndAnalyzeTicket(fileData);
    } catch (err) {
      logToSheet_(functionName, `Error processing file ${fileData.name} in batch: ${err.message}`, "ERROR");
      singleResult = {
        success: false,
        message: `Failed to process ${fileData.name}: ${err.message}`,
        fileUrl: null, // Ensure these are present for consistency
        fileId: null,
        ocrData: null
      };
    }
    summary.results.push({ index: idx, name: fileData.name, ...singleResult });

    // Optional: Add a small delay if hitting API rate limits with Gemini
    // Adjust sleep time as needed or remove if not necessary.
    // Consider if GEMINI_API_KEY is enterprise or has higher quotas.
    if (filesData.length > 1) { // Only sleep if it's a true batch
         Utilities.sleep(500); // 500 ms delay
    }
  });

  summary.finishedAt = new Date().toISOString();
  logToSheet_(functionName, `Batch processing finished. ${summary.results.filter(r => r.success).length} succeeded.`, "INFO");
  setBatchLoggingEnabled(false); // Disable batch logging (this will also flush any remaining logs)
  return summary;
}
// ====================================================================
// Gemini API Interaction
// ====================================================================

/**
 * Analyzes a ticket PDF using the Gemini API to extract specific fields,
 * including a suggested short agency name for filenames.
 * Handles cases where no violations matching the schema are found (e.g., toll statements).
 * @param {string} fileId The Google Drive File ID of the PDF ticket.
 * @returns {object} { success: boolean, data: object|null, message: string|null }
 */
function analyzeTicketWithGemini_(fileId) {
  const functionName = "analyzeTicketWithGemini_";
  logToSheet_(functionName, `Starting Gemini analysis for file ID: ${fileId}`, "INFO");

const apiKey =
  GEMINI_API_KEY_CACHE || (GEMINI_API_KEY_CACHE = getScriptProperty_('GEMINI_API_KEY'));

  if (!apiKey) {
    logToSheet_(functionName, "Gemini API Key is missing in Script Properties.", "ERROR");
    return { success: false, data: null, message: "Gemini API Key is missing." };
  }

  const model = "gemini-2.5-flash";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  try {
    // Get File Blob
    const file = DriveApp.getFileById(fileId);
    if (!file) throw new Error(`Could not find file in Drive with ID: ${fileId}`);
    const blob = file.getBlob();
    if (blob.getContentType() !== 'application/pdf') {
        logToSheet_(functionName, `File is not a PDF. MimeType: ${blob.getContentType()}`, "ERROR");
        throw new Error(`Invalid file type. Expected PDF, got ${blob.getContentType()}`);
    }
    const base64Data = Utilities.base64Encode(blob.getBytes());
    logToSheet_(functionName, `File Name: ${file.getName()}, Size: ${blob.getBytes().length} bytes.`, "DEBUG");

    // --- ** REVISED Gemini Prompt - Accepts Tolls, less strict on date/time output format ** ---
    const prompt = `Analyze the attached traffic, parking ticket, OR toll statement PDF. Identify if it contains one or multiple distinct violations/tolls.
Determine if the document appears to be a standard initial notice or a LATE / SECOND / FINAL notice.

Format the output strictly as a JSON object with the following top-level keys:
- "isLateNotice": boolean (true if it appears to be a late notice, false otherwise).
- "violations": An array of JSON objects. Each object represents ONE distinct violation or toll entry. Contain these keys:
    - "isToll": boolean. Set TRUE if this line item comes from a TOLL STATEMENT (e.g. E-ZPass, FasTrak, SunPass, I-Pass, Good To Go, TollRoads, DriveKS, any toll-authority statement — including statement fees or administrative line items on such a statement). Set FALSE for parking tickets, moving violations, red-light/bus-lane/speed-camera citations, and any non-toll municipal citation.
    - "licensePlate": Vehicle license plate number (string). Extract exactly as seen.
    - "licenseState": State of the license plate (string, e.g., "CA", "NY", "New Jersey"). Extract exactly as seen.
    - "violationDate": Date the violation/toll occurred (string). Extract the date exactly as it appears (e.g., "2024-03-15", "03/15/2024", "03/15/24").
    - "violationTime": Time the violation/toll occurred (string). Extract the time exactly as it appears (e.g., "14:35:00", "14:35", "02:35 PM", "2:35PM").
    - "violationId": The unique citation number OR a unique identifier for the toll/statement (e.g., statement ID, transaction ID if specific violation ID is absent) (string). Use Statement ID if no specific violation ID exists.
    - "pinNumber": Any PIN or access code mentioned (string, null if none).
    - "issuingAgency": Full name of the city, agency, or toll authority (string). Extract exactly as written.
    - "agencyShortName": A concise abbreviation for the issuing agency, suitable for a filename (string). Examples: "SFMTA", "NYC", "LAPD", "DRIVEKS", "TOLLROAD_AUTH_TX". Use ONLY uppercase letters (A-Z), numbers (0-9), and underscores (_). Be consistent.
    - "violationType": Brief description (e.g., "Expired Meter", "Speeding", "Toll Charge", "Mailed Statement Fee"). Expand common abbreviations if possible. (string).
    - "violationLocation": Location where violation/toll occurred (string). Extract exactly as written.
    - "originalPenaltyAmount": The initial fine or toll amount for THIS specific entry (number, e.g., 75.00, 1.88). If multiple tolls/fees are listed separately, create separate violation objects. For a statement fee, list it as its own entry.
    - "additionalPenaltyAmount": Any late fees or other penalties explicitly listed FOR THIS specific entry or broadly for the notice if it's late (number, default to 0).
    - "paymentWebsite": If a website URL for payment is provided (string, null if none).
    - "dueDate": The date payment is due (string, format YYYY-MM-DD if easily parsable, otherwise extract as seen, null if none).

**IMPORTANT**:
- If the document is a TOLL STATEMENT with multiple toll entries AND/OR fees, create a SEPARATE object in the "violations" array for EACH toll line item and EACH fee line item found. Use the Statement ID or Account number as the 'violationId' if a more specific ID per line isn't available, but ensure amounts reflect individual lines.
- If only one violation/toll is found, the "violations" array should contain one object.
- If NO specific violations or tolls are identifiable in the required structure, return an EMPTY "violations" array, e.g., {"isLateNotice": false, "violations": []}. Do NOT invent data.

Example (Toll Statement):
{
  "isLateNotice": false,
  "violations": [
    { "isToll": true, "licensePlate": "KS-D856147", "licenseState": "KS", "violationDate": "03/01/25", "violationTime": "07:10 PM", "violationId": "STMT_28795752_TOLL_1", /*Append _TOLL_# or _FEE_# if needed */ "pinNumber": null, "issuingAgency": "KTA", "agencyShortName": "DRIVEKS", "violationType": "Toll", "violationLocation": "Andover 21st/Wichita K-96", "originalPenaltyAmount": 0.38, "additionalPenaltyAmount": 0, "paymentWebsite": "DriveKS.com", "dueDate": "04/07/2025" },
    { "isToll": true, "licensePlate": "KS-D856147", "licenseState": "KS", "violationDate": "03/18/25", "violationTime": "12:19 AM", "violationId": "STMT_28795752_FEE_1", "pinNumber": null, "issuingAgency": "KTA", "agencyShortName": "DRIVEKS", "violationType": "Mailed Statement Fee", "violationLocation": null, "originalPenaltyAmount": 1.50, "additionalPenaltyAmount": 0, "paymentWebsite": "DriveKS.com", "dueDate": "04/07/2025" }
  ]
}

Example (Parking Ticket — non-toll):
{
  "isLateNotice": false,
  "violations": [
    { "isToll": false, "licensePlate": "8ABC123", "licenseState": "CA", "violationDate": "02/14/2026", "violationTime": "10:32 AM", "violationId": "SF-12345678", "pinNumber": null, "issuingAgency": "San Francisco Municipal Transportation Agency", "agencyShortName": "SFMTA", "violationType": "Expired Meter", "violationLocation": "100 Market St", "originalPenaltyAmount": 87.00, "additionalPenaltyAmount": 0, "paymentWebsite": "sfmta.com", "dueDate": "03/16/2026" }
  ]
}`;

    // --- Construct Payload ---
    const payload = {
      "contents": [{
        "parts": [
          { "text": prompt },
          { "inline_data": { "mime_type": "application/pdf", "data": base64Data } }
        ]
      }],
      "generationConfig": {
        "response_mime_type": "application/json",
        "temperature": 0.1, // Low temperature for factual extraction
      }
    };

    // --- Call Gemini API ---
    logToSheet_(functionName, `Calling Gemini API (${model})...`, "INFO");
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    logToSheet_(functionName, `Gemini API Response Code: ${responseCode}`, "DEBUG");

    if (responseCode !== 200) {
      logToSheet_(functionName, `Gemini API Error Response: ${responseText.substring(0, 1000)}`, "ERROR");
      throw new Error(`Gemini API call failed. Code: ${responseCode}. Response: ${responseText.substring(0, 500)}`);
    }

    // --- Parse Gemini Response ---
    logToSheet_(functionName, `Raw Gemini Response: ${responseText.substring(0,500)}...`, "DEBUG");
    const responseData = JSON.parse(responseText);

     if (!responseData || !responseData.candidates || !responseData.candidates.length > 0 ||
         !responseData.candidates[0].content || !responseData.candidates[0].content.parts ||
         !responseData.candidates[0].content.parts.length > 0 || !responseData.candidates[0].content.parts[0].text) {
         logToSheet_(functionName, "Gemini response structure is missing expected content.", "ERROR");
         logToSheet_(functionName, "Full Response (truncated): " + responseText.substring(0, 1000), "DEBUG");
         throw new Error("Invalid or incomplete response structure received from Gemini.");
    }

    const candidate = responseData.candidates[0];
    if (candidate.finishReason && candidate.finishReason !== "STOP") {
      logToSheet_(functionName, `Gemini generation finished with reason: ${candidate.finishReason}. Check safety ratings or response content.`, "WARN");
    }

    let generatedJsonText = candidate.content.parts[0].text;
    logToSheet_(functionName, `Gemini Generated JSON Text: ${generatedJsonText.substring(0, 500)}...`, "INFO");

    let extractedOuterData;
    try {
        generatedJsonText = generatedJsonText.replace(/^```json\s*/, '').replace(/\s*```$/, '');
        extractedOuterData = JSON.parse(generatedJsonText);
    } catch (jsonParseError) {
        logToSheet_(functionName, `Failed to parse JSON response from Gemini: ${jsonParseError}`, "ERROR");
        logToSheet_(functionName, `Cleaned text received for parsing: ${generatedJsonText.substring(0,1000)}`, "DEBUG");
        throw new Error("Failed to parse the JSON output provided by Gemini.");
    }

    // --- **REVISED Validation** ---
    // Check for the main keys, but allow violations array to be empty
    if (!extractedOuterData || typeof extractedOuterData.isLateNotice === 'undefined' || !Array.isArray(extractedOuterData.violations)) {
         logToSheet_(functionName, "Gemini response missing expected top-level structure (isLateNotice, violations array). Data: " + JSON.stringify(extractedOuterData), "ERROR");
         throw new Error("Gemini response structure is invalid (missing isLateNotice or violations array).");
    }

    // If violations array is empty, log a specific message but consider it a successful analysis
    if (extractedOuterData.violations.length === 0) {
        logToSheet_(functionName, "Gemini Analysis SUCCESSFUL, but returned 0 violations/tolls matching schema. This might be expected for certain documents (e.g., informational notices, non-violation/toll statements).", "WARN");
        // Proceed, but later steps (renaming, row adding) might be skipped if they rely on violation data.
    } else {
        logToSheet_(functionName, `Gemini Analysis SUCCESS. Late Notice: ${extractedOuterData.isLateNotice}, Violations/Tolls Found: ${extractedOuterData.violations.length}`, "SUCCESS");
        // Basic validation/cleaning for each violation in the array
        extractedOuterData.violations.forEach((violation, index) => {
            // Validate mandatory fields for renaming (log warning if missing)
            if (!violation.agencyShortName) logToSheet_(functionName, `Violation/Toll ${index+1} is missing 'agencyShortName'.`, "WARN");
            if (!violation.violationId) logToSheet_(functionName, `Violation/Toll ${index+1} is missing 'violationId'.`, "WARN");

            // Ensure amounts are numbers
            violation.originalPenaltyAmount = parseFloat(violation.originalPenaltyAmount) || 0;
            violation.additionalPenaltyAmount = parseFloat(violation.additionalPenaltyAmount) || 0;
            // Add more validation per violation if needed (e.g., date/time string format checks could go here)
        });
    }

    return {
        success: true,
        data: extractedOuterData, // Return the whole outer object
        message: `Ticket analyzed by Gemini. Found ${extractedOuterData.violations.length} violation(s)/toll(s).`
    };

  } catch (e) {
    logToSheet_(functionName, `Error during Gemini analysis: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    return {
        success: false,
        data: null,
        message: `Gemini analysis failed: ${e.message}`
    };
  }
}

// ====================================================================
// Data Submission and External Lookups (Called from HTML)
// ====================================================================

/**
 * Receives potentially multiple violations from the form, processes each one,
 * and writes them as separate rows to the target sheet.
 * @param {object} dataPackage Data submitted from the HTML form, containing
 *                   { ticketPDF: string, ticketFileId: string, isLateNotice: boolean, violations: object[] }
 * @returns {string} Success or error message summarizing the outcome.
 */
function addMultipleTicketData(dataPackage) {
  const functionName = "addMultipleTicketData";
  logToSheet_(functionName, `Starting submission for ${dataPackage.violations ? dataPackage.violations.length : 0} violations. Late Notice: ${dataPackage.isLateNotice}`, "INFO");
  // Log snippet of first violation for debugging
  if (dataPackage.violations && dataPackage.violations.length > 0) {
     logToSheet_(functionName, "First violation data snippet: " + JSON.stringify(dataPackage.violations[0]).substring(0, 300) + "...", "DEBUG");
  }


  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  try {
    // --- 1. Input Validation ---
    if (!dataPackage || !Array.isArray(dataPackage.violations) || dataPackage.violations.length === 0 || !dataPackage.ticketPDF) {
      throw new Error("Invalid data package received: Missing violations array or PDF link.");
    }

    const isLate = dataPackage.isLateNotice || false; // Get late notice flag

    // --- 2. Process Each Violation ---
    for (let i = 0; i < dataPackage.violations.length; i++) {
        const violationData = dataPackage.violations[i];
        logToSheet_(functionName, `Processing violation ${i + 1} of ${dataPackage.violations.length}: ID ${violationData.violationId || 'N/A'}`, "INFO");

        try {
            // Call helper function to handle row creation, lookups, and formatting for THIS violation
            appendAndFormatViolationRow_(violationData, dataPackage.ticketPDF, isLate);
            successCount++;
        } catch (singleError) {
            // Catch errors from the helper function for a specific violation
            logToSheet_(functionName, `Error processing violation ${i + 1} (ID: ${violationData.violationId || 'N/A'}): ${singleError.toString()}` + (singleError.stack ? ` Stack: ${singleError.stack}` : ""), "ERROR");
            errors.push(`Violation ${i + 1} (ID: ${violationData.violationId || 'N/A'}): ${singleError.message}`);
            errorCount++;
            // Continue to the next violation even if one fails
        }
    }

    // --- 3. Return Summary Message ---
    let message = "";
    if (successCount > 0) {
        message += `${successCount} violation(s) submitted successfully.`;
    }
    if (errorCount > 0) {
        message += (message ? " " : "") + `${errorCount} violation(s) failed to submit. Errors: ${errors.join("; ")}`;
    }
    if (successCount === 0 && errorCount === 0) {
       message = "No violations were processed."; // Should not happen with validation
    }

    logToSheet_(functionName, `Submission complete. Success: ${successCount}, Errors: ${errorCount}.`, errorCount > 0 ? "WARN" : "SUCCESS");
    return message;

  } catch (e) {
    // Catch errors in the main function (e.g., initial validation)
    logToSheet_(functionName, `Error adding multiple ticket data: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    throw new Error(`Server Error: Failed to process batch. ${e.message}`);
  }
}


/**
 * Helper function to append and format a single violation row.
 * Performs lookups, applies formatting, and handles late notice styling.
 * Includes robust date/time parsing.
 * @param {object} violationData The data object for a single violation.
 * @param {string} pdfLink The common PDF link for this batch.
 * @param {boolean} isLateNotice Indicates if this is a late notice.
 */
function appendAndFormatViolationRow_(violationData, pdfLink, isLateNotice) {
    const functionName = "appendAndFormatViolationRow_";
    logToSheet_(functionName, `Appending row for Violation ID: ${violationData.violationId || 'N/A'}`, "DEBUG");

    // --- Prepare Row Data Array ---
    const rowData = new Array(NUM_COLUMNS).fill('');

    // --- Populate from Violation Data & Apply Transformations ---
    rowData[COL.LICENSE_PLATE - 1] = violationData.licensePlate;
    rowData[COL.LICENSE_PLATE_STATE - 1] = abbreviateState_(violationData.licenseState);

    // --- ** Revised Date/Time Parsing ** ---
    const parsedDate = parseFlexibleDate_(violationData.violationDate);
    const parsedTime = parseFlexibleTime_(violationData.violationTime, parsedDate); // Pass parsedDate for context if needed
    rowData[COL.DATE_VIOLATION - 1] = parsedDate ? Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "MM/dd/yyyy") : violationData.violationDate; // Format if valid Date object, else keep original string
    rowData[COL.TIME_VIOLATION - 1] = parsedTime ? Utilities.formatDate(parsedTime, Session.getScriptTimeZone(), "h:mm a").toUpperCase() : violationData.violationTime; // Format if valid Date object, else keep original string
    // --- End Revised Date/Time Parsing ---

    rowData[COL.VIOLATION_ID - 1] = violationData.violationId;
    rowData[COL.PIN_NUMBER - 1] = violationData.pinNumber;
    rowData[COL.ISSUING_AGENCY - 1] = violationData.issuingAgency;
    rowData[COL.VIOLATION_TYPE - 1] = formatViolationType_(violationData.violationType);
    rowData[COL.VIOLATION_LOCATION - 1] = formatLocation_(violationData.violationLocation);
    const originalPenalty = parseFloat(violationData.originalPenaltyAmount) || 0;
    const additionalPenalty = parseFloat(violationData.additionalPenaltyAmount) || 0;
    rowData[COL.ORIGINAL_PENALTY - 1] = originalPenalty;
    rowData[COL.ADDITIONAL_PENALTY - 1] = additionalPenalty;
    const isToll = classifyIsToll_(violationData);
    rowData[COL.TOLL_OR_TICKET - 1] = isToll ? "Toll" : "Ticket";

    let notesContent = violationData.paymentWebsite || '';
    if (isLateNotice) {
        notesContent += (notesContent ? "; " : "") + "Late Notice";
    }
    if (isToll) {
        notesContent += (notesContent ? "; " : "") + "Also routed to toll-processing/current_tolls for admin agency review";
    }
    rowData[COL.NOTES - 1] = notesContent;

    // --- Due Date Parsing ---
    const parsedDueDate = parseFlexibleDate_(violationData.dueDate);
    rowData[COL.DUE_DATE - 1] = parsedDueDate ? Utilities.formatDate(parsedDueDate, Session.getScriptTimeZone(), "MM/dd/yyyy") : violationData.dueDate;
    // --- End Due Date Parsing ---

    rowData[COL.PDF_LINK - 1] = pdfLink;
    rowData[COL.PAYABLE_AMOUNT - 1] = originalPenalty + additionalPenalty;

    // --- Lookups (Fleet & DB) ---
    const vehicleNumber = getVehicleNumberFromFleetSheet_(violationData.licensePlate); // Uses cached version
    if (vehicleNumber) {
      rowData[COL.VEHICLE - 1] = vehicleNumber;
      // Perform DB lookup ONLY IF we successfully parsed violation date/time
      if (parsedDate && parsedTime) {
          // Reconstruct local time string in a standard format for DB lookup function
          // Assuming getDriverBookingInfo_ expects "YYYY-MM-DD HH:MM:SS"
          const violationDateTimeLocalStr = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd") + " " + Utilities.formatDate(parsedTime, Session.getScriptTimeZone(), "HH:mm:ss");
          logToSheet_(functionName, `Formatted DateTime for DB lookup: ${violationDateTimeLocalStr}`, "DEBUG");

          const dbInfo = getDriverBookingInfo_(vehicleNumber, violationDateTimeLocalStr);
          if (dbInfo && !dbInfo.error) {
            rowData[COL.RESPONSIBLE_DRIVER - 1] = dbInfo.driverName || '';
            rowData[COL.DRIVER_EMAIL - 1] = dbInfo.driverEmail || '';
            rowData[COL.BOOKING_ID - 1] = dbInfo.bookingId || '';
            rowData[COL.PROPERTY - 1] = dbInfo.propertyName || '';
          } else {
            logToSheet_(functionName, `DB Lookup failed for Viol ID ${violationData.violationId}: ${dbInfo ? dbInfo.error : 'No details'}`, "WARN");
          }
      } else {
          logToSheet_(functionName, `Skipping DB lookup for Viol ID ${violationData.violationId} due to unparsable date/time from Gemini ('${violationData.violationDate}', '${violationData.violationTime}').`, "WARN");
      }
    } else {
      logToSheet_(functionName, `Vehicle # not found for plate ${violationData.licensePlate} (Viol ID ${violationData.violationId}).`, "WARN");
      rowData[COL.VEHICLE - 1] = "#N/A";
    }

    // --- Set Default Statuses ---
    rowData[COL.EMAIL_STATUS - 1] = "Ready";
    rowData[COL.PAYMENT_STATUS - 1] = '';
    rowData[COL.DRIVER_BILLING_STATUS - 1] = '';

    // --- Write to Target Sheet ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Target sheet "${TARGET_SHEET_NAME}" not found.`);
    }
    // Consider batching appends if performance is still an issue
    sheet.appendRow(rowData);
    const newRowIndex = sheet.getLastRow();
    logToSheet_(functionName, `Appended row ${newRowIndex} for Violation ID: ${violationData.violationId || 'N/A'}`, "DEBUG");


    // --- Apply Formatting ---
    // Get range AFTER appending
    const formatRange = sheet.getRange(newRowIndex, 1, 1, Math.min(sheet.getLastColumn(), NUM_COLUMNS));
    const backgroundColor = isLateNotice ? "#f4cccc" : "#fff2cc";
    formatRange.setBackground(backgroundColor);
    //logToSheet_(functionName, `Set background to ${backgroundColor} (Late: ${isLateNotice}) for row ${newRowIndex}`, "DEBUG"); // Reduced logging noise

    sheet.getRange(newRowIndex, COL.VEHICLE).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange(newRowIndex, COL.ORIGINAL_PENALTY).setNumberFormat("$#,##0.00");
    sheet.getRange(newRowIndex, COL.ADDITIONAL_PENALTY).setNumberFormat("$#,##0.00");
    sheet.getRange(newRowIndex, COL.PAYABLE_AMOUNT).setNumberFormat("$#,##0.00");
    sheet.getRange(newRowIndex, COL.TOTAL_OWED_COLLECTIONS).setNumberFormat("$#,##0.00");

    // --- REMOVED SpreadsheetApp.flush() ---
    // SpreadsheetApp.flush(); // Avoid flushing inside loop
    logToSheet_(functionName, `Formatting applied for row ${newRowIndex}`, "DEBUG");

    // --- Toll routing: also append to toll-processing/current_tolls ---
    let routedToTolls = false;
    if (isToll) {
        try {
            routeTollToCurrentTolls_(violationData, parsedDate, parsedTime, {
                driverName: rowData[COL.RESPONSIBLE_DRIVER - 1] || '',
                driverEmail: rowData[COL.DRIVER_EMAIL - 1] || '',
                bookingId: rowData[COL.BOOKING_ID - 1] || ''
            }, rowData[COL.VEHICLE - 1] || '', pdfLink);
            routedToTolls = true;
        } catch (routeErr) {
            logToSheet_(functionName, `Toll routing to current_tolls failed for Viol ID ${violationData.violationId || 'N/A'}: ${routeErr.toString()}`, "ERROR");
        }
    }

    // --- New-row email notification (never blocks row append) ---
    try {
        notifyNewRow_(isToll, violationData, rowData[COL.VEHICLE - 1] || '', newRowIndex, pdfLink, routedToTolls);
    } catch (mailErr) {
        logToSheet_(functionName, `Email notification failed for Viol ID ${violationData.violationId || 'N/A'}: ${mailErr.toString()}`, "WARN");
    }
}

// ====================================================================
// Toll vs Ticket classification + cross-sheet routing + email notify
// ====================================================================

/**
 * Decide whether a Gemini violation entry should be treated as a toll.
 * Primary signal: Gemini's `isToll` boolean (added to the prompt schema).
 * Fallback: keyword match on violationType / issuingAgency / paymentWebsite
 * when Gemini omits the field or returns something non-boolean.
 */
function classifyIsToll_(violationData) {
    if (!violationData || typeof violationData !== 'object') return false;
    if (typeof violationData.isToll === 'boolean') return violationData.isToll;
    const hay = [
        violationData.violationType,
        violationData.issuingAgency,
        violationData.agencyShortName,
        violationData.paymentWebsite
    ].filter(Boolean).join(' ').toLowerCase();
    return /toll|ez-?pass|i-?pass|tollroad|good to go|drivek|statement fee|fastrak|sunpass|peach pass|mta bridges/.test(hay);
}

/**
 * Admin fee mirror of toll-processing/reformatting.gs:calculateAdminFee.
 * Kept inline to avoid a cross-project library dependency.
 */
function calculateTollAdminFee_(tollAmount) {
    if (typeof tollAmount !== 'number' || isNaN(tollAmount)) return 0;
    if (tollAmount > 3.99) return 2.00;
    if (tollAmount > 1)    return 1.00;
    return 0.50;
}

/**
 * Append a row to toll-processing!current_tolls representing a toll that
 * arrived via the ticket-PDF uploader. Customer reimbursement (base toll +
 * admin fee) will run through the existing Stripe flow on that sheet; the
 * tickets-sheet row continues to cover the agency payment via accounting.
 */
function routeTollToCurrentTolls_(violationData, parsedDate, parsedTime, dbInfo, vehicleNumber, pdfLink) {
    const functionName = "routeTollToCurrentTolls_";
    const ss = SpreadsheetApp.openById(TOLL_PROCESSING_SHEET_ID);
    if (!ss) throw new Error(`Could not open toll-processing spreadsheet ${TOLL_PROCESSING_SHEET_ID}`);
    const sheet = ss.getSheetByName(TOLL_CURRENT_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${TOLL_CURRENT_SHEET_NAME}" not found in toll-processing spreadsheet`);

    const tz = Session.getScriptTimeZone();
    let firstName = '', lastName = '';
    if (dbInfo && dbInfo.driverName) {
        const parts = String(dbInfo.driverName).trim().split(/\s+/);
        firstName = parts.shift() || '';
        lastName  = parts.join(' ');
    }

    let dateTimeLocalStr = '';
    if (parsedDate && parsedTime) {
        dateTimeLocalStr = Utilities.formatDate(parsedDate, tz, "yyyy-MM-dd") + " " + Utilities.formatDate(parsedTime, tz, "HH:mm:ss");
    } else if (parsedDate) {
        dateTimeLocalStr = Utilities.formatDate(parsedDate, tz, "yyyy-MM-dd");
    } else {
        dateTimeLocalStr = [violationData.violationDate, violationData.violationTime].filter(Boolean).join(' ');
    }

    const tollAmount = parseFloat(violationData.originalPenaltyAmount) || 0;
    const adminFee   = calculateTollAdminFee_(tollAmount);
    const totalDue   = Math.round((tollAmount + adminFee) * 100) / 100;
    const penaltyAmt = parseFloat(violationData.additionalPenaltyAmount) || 0;

    const note = [
        'From ticket PDF',
        `violationId=${violationData.violationId || 'N/A'}`,
        `agency=${violationData.issuingAgency || 'N/A'}`,
        `plate=${violationData.licensePlate || 'N/A'}`,
        penaltyAmt > 0 ? `penalty_billed_accounting=$${penaltyAmt.toFixed(2)}` : null,
        pdfLink ? `PDF=${pdfLink}` : null
    ].filter(Boolean).join(' · ');

    const row = [
        firstName,                                 // A First Name
        lastName,                                  // B Last Name
        (dbInfo && dbInfo.driverEmail) || '',      // C User Email
        vehicleNumber || '',                       // D Envoy #
        (dbInfo && dbInfo.bookingId) || '',        // E Booking ID
        dateTimeLocalStr,                          // F Time and Date (Local)
        violationData.violationType || '',         // G Type
        tollAmount,                                // H Toll Amount
        adminFee,                                  // I Admin Fee
        totalDue,                                  // J Total Due
        '',                                        // K Confirmed?
        '',                                        // L Charged?
        '',                                        // M Notified?
        '',                                        // N Booking Toll Total
        '',                                        // O (blank)
        '',                                        // P Booking Sum (from H)
        '',                                        // Q Booking Sum (from I)
        '',                                        // R Booking GL String
        note,                                      // S Ticket Source Note
        'YES'                                      // T Needs Agency Review
    ];

    sheet.appendRow(row);
    const newIdx = sheet.getLastRow();
    try {
        sheet.getRange(newIdx, 8, 1, 3).setNumberFormat("$#,##0.00"); // H, I, J currency
    } catch (fmtErr) {
        logToSheet_(functionName, `Currency format failed on routed toll row ${newIdx}: ${fmtErr}`, "DEBUG");
    }
    logToSheet_(functionName, `Appended toll to current_tolls row ${newIdx} (plate ${violationData.licensePlate}, amount $${tollAmount.toFixed(2)})`, "SUCCESS");
}

/** Per-row email notification to NEW_ROW_NOTIFY_EMAIL with subject "new toll" / "new ticket". */
function notifyNewRow_(isToll, violationData, vehicleNumber, newRowIndex, pdfLink, routedToTolls) {
    if (!NEW_ROW_NOTIFY_EMAIL) return;
    const plate  = violationData.licensePlate || 'N/A';
    const state  = violationData.licenseState || '';
    const date   = violationData.violationDate || '';
    const time   = violationData.violationTime || '';
    const agency = violationData.issuingAgency || '';
    const vtype  = violationData.violationType || '';
    const orig   = parseFloat(violationData.originalPenaltyAmount) || 0;
    const addl   = parseFloat(violationData.additionalPenaltyAmount) || 0;
    const total  = orig + addl;
    const amtStr = total.toFixed(2);

    const subject = isToll
        ? `new toll — ${plate}${state ? ' ' + state : ''} · ${date} · $${amtStr}`
        : `new ticket — ${plate}${state ? ' ' + state : ''} · ${agency || vtype || 'citation'}`;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss ? ss.getId() : '';
    const rowUrl = (ssId && newRowIndex) ? `https://docs.google.com/spreadsheets/d/${ssId}/edit#gid=0&range=A${newRowIndex}` : '';

    const lines = [
        `<p><b>${isToll ? 'Toll' : 'Ticket'}</b> logged to "${TARGET_SHEET_NAME}" (row ${newRowIndex})</p>`,
        `<p>Plate: ${plate}${state ? ' (' + state + ')' : ''} · Envoy #: ${vehicleNumber || 'N/A'}</p>`,
        `<p>When: ${date} ${time}</p>`,
        agency ? `<p>Issuing agency: ${agency}</p>` : '',
        vtype  ? `<p>Type: ${vtype}</p>` : '',
        `<p>Original: $${orig.toFixed(2)} · Penalty: $${addl.toFixed(2)} · Total: $${amtStr}</p>`,
        pdfLink ? `<p><a href="${pdfLink}">PDF</a></p>` : '',
        rowUrl  ? `<p><a href="${rowUrl}">Open row in tickets sheet</a></p>` : '',
        isToll
            ? `<p>Also routed to toll-processing <i>current_tolls</i>: ${routedToTolls ? 'yes' : 'NO — see logs'}</p>`
            : ''
    ];
    const htmlBody = lines.filter(Boolean).join('\n');

    MailApp.sendEmail({
        to: NEW_ROW_NOTIFY_EMAIL,
        subject: subject,
        htmlBody: htmlBody
    });
}

/**
 * One-time setup: ensure current_tolls has headers for columns S and T.
 * Run manually from the Apps Script editor after first deploy. Idempotent.
 */
function seedCurrentTollsHeaders_() {
    const ss = SpreadsheetApp.openById(TOLL_PROCESSING_SHEET_ID);
    const sheet = ss.getSheetByName(TOLL_CURRENT_SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${TOLL_CURRENT_SHEET_NAME}" not found`);
    const hdrRange = sheet.getRange(1, 19, 1, 2); // S1:T1
    const existing = hdrRange.getValues()[0];
    if (existing[0] !== 'Ticket Source Note' || existing[1] !== 'Needs Agency Review') {
        hdrRange.setValues([['Ticket Source Note', 'Needs Agency Review']]).setFontWeight('bold');
    }
}

/** Helper to format values as plain text for sheet to avoid auto-formatting issues */
function formatValueForSheet_(value) {
  if (value === null || value === undefined) return '';
  // Prepend apostrophe to force text interpretation for dates/times/IDs
  if (typeof value === 'string' && (value.match(/^\d{4}-\d{2}-\d{2}$/) || value.match(/^\d{2}:\d{2}(:\d{2})?$/) || value.match(/^[A-Z0-9]+$/i)) ) {
     return "'" + value;
  }
  return value;
}

// ====================================================================
// External Sheet Lookup (Fleet Overview)
// ====================================================================

// ====================================================================
// External Sheet Lookup (Fleet Overview)
// ====================================================================

/**
 * Looks up the Vehicle Number (Col A) in the Fleet Sheet based on License Plate.
 * First checks Column F, then falls back to checking Column I.
 * Cleans license plate input (removes trailing state).
 * Implements caching to improve performance.
 * @param {string} licensePlate The license plate to search for (may include state).
 * @returns {string|null} The vehicle number string, or null if not found/error.
 */
function getVehicleNumberFromFleetSheet_(licensePlate) {
  const functionName = "getVehicleNumberFromFleetSheet_";
  if (!licensePlate) {
    logToSheet_(functionName, "License plate parameter is missing.", "WARN");
    return null;
  }

  // --- Plate Cleaning ---
  let cleanedPlate = licensePlate.trim().toUpperCase();
  // Regex: Matches a space followed by exactly 2 capital letters at the end of the string
  cleanedPlate = cleanedPlate.replace(/ ([A-Z]{2})$/, '');
  logToSheet_(functionName, `Original Plate: "${licensePlate}", Cleaned Plate for Lookup: "${cleanedPlate}"`, "DEBUG");
  // --- End Plate Cleaning ---


  // --- Caching Logic ---
  const cache = CacheService.getScriptCache();
  const cacheKey = `fleet_plate_${cleanedPlate}`; // Use cleaned plate for cache key
  const cachedVehicleNumber = cache.get(cacheKey);

  if (cachedVehicleNumber !== null) {
    logToSheet_(functionName, `Cache HIT for plate "${cleanedPlate}". Vehicle #: ${cachedVehicleNumber || '#N/A or Blank'}`, "DEBUG");
    // Return null if cache explicitly stored 'not_found', otherwise return the cached number
    return cachedVehicleNumber === 'not_found' ? null : cachedVehicleNumber;
  }
  logToSheet_(functionName, `Cache MISS for plate "${cleanedPlate}". Performing sheet lookup.`, "DEBUG");
  // --- End Caching Logic ---

  try {
    const fleetSheet = SpreadsheetApp.openById(FLEET_SHEET_ID);
    const tab = fleetSheet.getSheetByName(FLEET_SHEET_TAB_NAME);
    if (!tab) {
      logToSheet_(functionName, `Fleet sheet tab "${FLEET_SHEET_TAB_NAME}" not found in sheet ID ${FLEET_SHEET_ID}.`, "ERROR");
      return null; // Cannot proceed
    }

    const data = tab.getDataRange().getValues();
    const colFIndex = 5; // Column F (0-based)
    const colIIndex = 8; // Column I (0-based)
    const colAIndex = 0; // Column A (0-based) Vehicle #
    let foundVehicleNumber = null;

    logToSheet_(functionName, `Read ${data.length} rows from Fleet sheet. Searching for CLEANED plate: ${cleanedPlate}`, "DEBUG");

    for (let i = 1; i < data.length; i++) { // Skip header row (i=0)
      const row = data[i];
      let vehicleNumberInRow = row[colAIndex] ? String(row[colAIndex]).trim() : ''; // Get vehicle # first

      // Check Column F
      const plateInF = row[colFIndex] ? String(row[colFIndex]).trim().toUpperCase() : '';
      if (plateInF === cleanedPlate) {
        if (vehicleNumberInRow) {
          logToSheet_(functionName, `Match found in Col F on sheet row ${i + 1}. Vehicle #: ${vehicleNumberInRow}`, "SUCCESS");
          foundVehicleNumber = vehicleNumberInRow;
          break; // Found a match with a vehicle number, exit loop
        } else {
          logToSheet_(functionName, `Match found in Col F on sheet row ${i + 1} for plate ${cleanedPlate}, but Vehicle # in Col A is blank. Checking Col I.`, "WARN");
          // Continue checking Col I for this row, just in case
        }
      }

      // Check Column I (only if not already found in F with a number)
      if (foundVehicleNumber === null) {
         const plateInI = row[colIIndex] ? String(row[colIIndex]).trim().toUpperCase() : '';
         if (plateInI === cleanedPlate) {
            if (vehicleNumberInRow) {
               logToSheet_(functionName, `Match found in Col I on sheet row ${i + 1}. Vehicle #: ${vehicleNumberInRow}`, "SUCCESS");
               foundVehicleNumber = vehicleNumberInRow;
               break; // Found a match with a vehicle number, exit loop
            } else {
               // Avoid double logging if F also matched blank
               if (plateInF !== cleanedPlate) {
                   logToSheet_(functionName, `Match found in Col I on sheet row ${i + 1} for plate ${cleanedPlate}, but Vehicle # in Col A is blank.`, "WARN");
               }
            }
         }
      }
    } // End loop

    // --- Update Cache ---
    if (foundVehicleNumber !== null) {
      cache.put(cacheKey, foundVehicleNumber, 21600); // Cache for 6 hours
      logToSheet_(functionName, `Cached result for "${cleanedPlate}": ${foundVehicleNumber}`, "DEBUG");
    } else {
      cache.put(cacheKey, 'not_found', 21600); // Cache the 'not found' status
      logToSheet_(functionName, `Cached result for "${cleanedPlate}": not_found`, "DEBUG");
    }
    // --- End Update Cache ---

    if (foundVehicleNumber === null) {
       logToSheet_(functionName, `Cleaned plate "${cleanedPlate}" not found with a valid Vehicle # in Columns F or I.`, "WARN");
    }

    return foundVehicleNumber; // Return the found number or null

  } catch (e) {
    logToSheet_(functionName, `Error accessing Fleet Sheet (ID: ${FLEET_SHEET_ID}): ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    return null; // Return null on error
  }
}

// ====================================================================
// Database Interaction (MySQL)
// ====================================================================

/**
 * Connects to the Envoy DB and finds booking/driver details based on vehicle and time.
 * Handles timezone conversion.
 * @param {string} vehicleNumber The vehicle number (e.g., "636").
 * @param {string} violationDateTimeLocalStr The violation date and time as a string ("YYYY-MM-DD HH:MM" or "YYYY-MM-DD HH:MM:SS") in the *local* timezone of the violation.
 * @returns {object} { driverName: string, driverEmail: string, bookingId: string, propertyName: string, error: string|null }
 */
function getDriverBookingInfo_(vehicleNumber, violationDateTimeLocalStr) {
  const functionName = "getDriverBookingInfo_";
  if (!vehicleNumber || !violationDateTimeLocalStr) {
    logToSheet_(functionName, "Missing vehicleNumber or violationDateTimeLocalStr for DB lookup.", "WARN");
    return { error: "Missing required parameters (vehicleNumber, violationDateTimeLocalStr)." };
  }

  const vehicleNameDbFormat = `Envoy ${vehicleNumber}`;
  logToSheet_(functionName, `Querying DB for vehicle: ${vehicleNameDbFormat}, local time: ${violationDateTimeLocalStr}`, "DEBUG");

  let conn = null;
  let stmt = null;
  let rs = null;
  let propStmt = null;
  let propRs = null;
  let bookingResults = []; // Store potential booking matches

  try {
    // --- Get DB Credentials ---
    const dbHost = getScriptProperty_('DB_HOST');
    const dbPort = getScriptProperty_('DB_PORT');
    const dbName = getScriptProperty_('DB_NAME');
    const dbUser = getScriptProperty_('DB_USER');
    const dbPassword = getScriptProperty_('DB_PASSWORD');

    if (!dbHost || !dbPort || !dbName || !dbUser || !dbPassword) {
      throw new Error("Database credentials missing in Script Properties.");
    }

    const dbUrl = `jdbc:mysql://${dbHost}:${dbPort}/${dbName}`;
    logToSheet_(functionName, `Connecting to DB: ${dbUrl} User: ${dbUser}`, "DEBUG");
    conn = Jdbc.getConnection(dbUrl, dbUser, dbPassword);
    conn.setAutoCommit(false); // Good practice for reads
    logToSheet_(functionName, "DB Connection successful.", "INFO");

    // --- Query 1: Find potential bookings around the violation date ---
    const violationDateOnly = violationDateTimeLocalStr.substring(0, 10); // YYYY-MM-DD
    const bookingSql = `
      SELECT ref_number, first_name, last_name, Email, Property, start_date, end_date
      FROM Envoy_Booking_View
      WHERE vehicle_name = ?
      AND DATE(start_date) <= ?
      AND DATE(end_date) >= ?
      ORDER BY start_date DESC`;

    stmt = conn.prepareStatement(bookingSql);
    stmt.setString(1, vehicleNameDbFormat);
    stmt.setString(2, violationDateOnly);
    stmt.setString(3, violationDateOnly);

    logToSheet_(functionName, `Executing Booking Query for vehicle ${vehicleNameDbFormat}, date ${violationDateOnly}`, "DEBUG");
    rs = stmt.executeQuery();

    while (rs.next()) {
       bookingResults.push({
         ref_number: rs.getString("ref_number"),
         first_name: rs.getString("first_name"),
         last_name: rs.getString("last_name"),
         Email: rs.getString("Email"),
         Property: rs.getString("Property"),
         start_date_str: rs.getString("start_date"), // PST String e.g., "2024-09-30 16:38:00" or "2024-09-30 16:38:00.0"
         end_date_str: rs.getString("end_date")     // PST String
       });
    }
    logToSheet_(functionName, `Found ${bookingResults.length} potential booking(s) around the violation date.`, "INFO");

    if (bookingResults.length === 0) {
      return { error: `No bookings found for vehicle ${vehicleNameDbFormat} around ${violationDateOnly}.` };
    }

    // --- Iterate through potential bookings to find the correct one using timezone ---
    let matchedBooking = null;
    for (const booking of bookingResults) {
       logToSheet_(functionName, `Checking booking ID: ${booking.ref_number}, Property: ${booking.Property}`, "DEBUG");

       // --- Query 2: Get Timezone for the booking's property ---
       const propertySql = "SELECT tz_id FROM properties WHERE name = ?";
       propStmt = conn.prepareStatement(propertySql);
       propStmt.setString(1, booking.Property);
       propRs = propStmt.executeQuery();

       let propertyTzId = null;
       if (propRs.next()) {
         propertyTzId = propRs.getString("tz_id");
         logToSheet_(functionName, `Found timezone '${propertyTzId}' for property '${booking.Property}'.`, "DEBUG");
       } else {
         logToSheet_(functionName, `Timezone not found for property '${booking.Property}'. Skipping this booking check.`, "WARN");
         closeQuietly_(propRs);
         closeQuietly_(propStmt);
         continue;
       }
       closeQuietly_(propRs);
       closeQuietly_(propStmt);

       if (!propertyTzId) continue;

       // --- Timezone Conversion & Comparison ---
       try {
         // --- **FIX 1: Determine correct format for violationDateTimeLocalStr** ---
         let violationDateFormat = "yyyy-MM-dd HH:mm"; // Assume HH:mm by default
         if (violationDateTimeLocalStr.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
            violationDateFormat = "yyyy-MM-dd HH:mm:ss"; // Use HH:mm:ss if seconds are present
         }
         logToSheet_(functionName, `Parsing local violation time '${violationDateTimeLocalStr}' using TZ: '${propertyTzId}' and format: '${violationDateFormat}'`, "DEBUG");
         const violationDateObj = Utilities.parseDate(violationDateTimeLocalStr, propertyTzId, violationDateFormat);
         const violationTimestamp = violationDateObj.getTime();
         logToSheet_(functionName, `Absolute violation timestamp (ms): ${violationTimestamp}`, "DEBUG");


         // --- **FIX 2: Clean DB date strings before parsing** ---
         // Remove trailing ".0" or ".<any_digits>" if present
         const cleanStartDateStr = booking.start_date_str.replace(/\.\d+$/, '');
         const cleanEndDateStr = booking.end_date_str.replace(/\.\d+$/, '');
         const dbDateFormat = "yyyy-MM-dd HH:mm:ss"; // The DB format without milliseconds

         // Parse the cleaned PST booking start/end times
         logToSheet_(functionName, `Parsing cleaned StartStr '${cleanStartDateStr}' using TZ: 'PST' and format: '${dbDateFormat}'`, "DEBUG");
         const bookingStartDateObj = Utilities.parseDate(cleanStartDateStr, "PST", dbDateFormat);
         logToSheet_(functionName, `Parsing cleaned EndStr '${cleanEndDateStr}' using TZ: 'PST' and format: '${dbDateFormat}'`, "DEBUG");
         const bookingEndDateObj = Utilities.parseDate(cleanEndDateStr, "PST", dbDateFormat);

         const bookingStartTimestamp = bookingStartDateObj.getTime();
         const bookingEndTimestamp = bookingEndDateObj.getTime();

         logToSheet_(functionName, `Booking ${booking.ref_number} Start (ms): ${bookingStartTimestamp} (Orig: ${booking.start_date_str} PST)`, "DEBUG");
         logToSheet_(functionName, `Booking ${booking.ref_number} End (ms): ${bookingEndTimestamp} (Orig: ${booking.end_date_str} PST)`, "DEBUG");

         // Compare the absolute timestamps
         if (violationTimestamp >= bookingStartTimestamp && violationTimestamp <= bookingEndTimestamp) {
           logToSheet_(functionName, `MATCH FOUND! Violation time falls within booking ${booking.ref_number}.`, "SUCCESS");
           matchedBooking = booking;
           break;
         } else {
           logToSheet_(functionName, `Violation time does not fall within booking ${booking.ref_number}.`, "DEBUG");
         }

       } catch (parseError) {
          // Log the cleaned strings as well in case of error
          const cleanedStart = booking.start_date_str.replace(/\.\d+$/, '');
          const cleanedEnd = booking.end_date_str.replace(/\.\d+$/, '');
          logToSheet_(functionName, `Error parsing dates/times for booking ${booking.ref_number}: ${parseError}. ViolationStr: '${violationDateTimeLocalStr}', TZ: '${propertyTzId}', CleanedStartStr: '${cleanedStart}', CleanedEndStr: '${cleanedEnd}'. Skipping check.`, "ERROR");
          continue;
       }
    } // End loop through potential bookings

    // --- Return results ---
    if (matchedBooking) {
       return {
         driverName: `${matchedBooking.first_name} ${matchedBooking.last_name}`,
         driverEmail: matchedBooking.Email,
         bookingId: matchedBooking.ref_number,
         propertyName: matchedBooking.Property,
         error: null
       };
    } else {
       return { error: `No matching booking found for vehicle ${vehicleNameDbFormat} at the specified time after checking ${bookingResults.length} potential(s).` };
    }

  } catch (e) {
    logToSheet_(functionName, `Database Error: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    return { error: `Database query failed: ${e.message}` };
  } finally {
    closeQuietly_(rs);
    closeQuietly_(stmt);
    closeQuietly_(propRs);
    closeQuietly_(propStmt);
    closeQuietly_(conn);
    logToSheet_(functionName, "DB resources closed.", "DEBUG");
  }
}
/** Close JDBC resources quietly */
function closeQuietly_(resource) {
  try {
    if (resource) resource.close();
  } catch (e) { /* ignore */ }
}


// ====================================================================
// Utility and Logging Functions
// ====================================================================
/**
 * Attempts to parse various common date string formats into a valid Date object.
 * Handles YYYY-MM-DD, MM/DD/YYYY, MM/DD/YY.
 * @param {string} dateString The date string extracted by Gemini.
 * @returns {Date|null} A valid Date object representing the start of the day in script timezone, or null if parsing fails.
 */
function parseFlexibleDate_(dateString) {
    if (!dateString || typeof dateString !== 'string') return null;
    dateString = dateString.trim();
    let parsedDate = null;
    const functionName = "parseFlexibleDate_";

    try {
        // Try YYYY-MM-DD
        if (/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
            // Split and create date to avoid timezone issues with Utilities.parseDate assuming UTC midnight
            const parts = dateString.split('-');
            parsedDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
        }
        // Try MM/DD/YYYY
        else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateString)) {
            parsedDate = Utilities.parseDate(dateString, Session.getScriptTimeZone(), "MM/dd/yyyy");
        }
        // Try MM/DD/YY (assume 20xx)
        else if (/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(dateString)) {
            const parts = dateString.split('/');
            const year = parseInt(parts[2]) + 2000; // Assume 20xx
            const month = parseInt(parts[0]) - 1;
            const day = parseInt(parts[1]);
            parsedDate = new Date(year, month, day);
        }
        // Add more formats here if needed (e.g., "Month Day, Year")

        if (parsedDate && !isNaN(parsedDate)) {
             logToSheet_(functionName, `Successfully parsed date string "${dateString}" to Date object.`, "DEBUG");
             return parsedDate;
        } else {
             logToSheet_(functionName, `Could not parse date string "${dateString}" into a known format.`, "WARN");
             return null;
        }
    } catch (e) {
        logToSheet_(functionName, `Error parsing date string "${dateString}": ${e}`, "ERROR");
        return null;
    }
}

/**
 * Attempts to parse various common time string formats into a valid Date object (using a dummy date).
 * Handles HH:MM:SS, HH:MM (24hr), h:mmA/PM, h:mm A/PM.
 * @param {string} timeString The time string extracted by Gemini.
 * @param {Date} [contextDate=null] Optional Date object for context (currently unused but could be useful).
 * @returns {Date|null} A valid Date object (on Jan 1, 1970) representing the time, or null if parsing fails.
 */
function parseFlexibleTime_(timeString, contextDate = null) {
    if (!timeString || typeof timeString !== 'string') return null;
    timeString = timeString.trim().toUpperCase(); // Normalize case for AM/PM
    let parsedTime = null;
    const functionName = "parseFlexibleTime_";

    try {
        let hours = 0, minutes = 0, seconds = 0;
        let match;

        // Try HH:MM:SS (24hr)
        match = timeString.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
        if (match) {
            hours = parseInt(match[1]);
            minutes = parseInt(match[2]);
            seconds = parseInt(match[3]);
            parsedTime = new Date(1970, 0, 1, hours, minutes, seconds);
        }
        // Try HH:MM (24hr)
        else if ((match = timeString.match(/^(\d{1,2}):(\d{2})$/))) {
            hours = parseInt(match[1]);
            minutes = parseInt(match[2]);
            parsedTime = new Date(1970, 0, 1, hours, minutes, 0);
        }
        // Try h:mm[ ]AM/PM or hh:mm[ ]AM/PM
        else if ((match = timeString.match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/))) {
            hours = parseInt(match[1]);
            minutes = parseInt(match[2]);
            const ampm = match[3];

            if (ampm === "PM" && hours < 12) hours += 12;
            if (ampm === "AM" && hours === 12) hours = 0; // Midnight case
            parsedTime = new Date(1970, 0, 1, hours, minutes, 0);
        }
         // Add more formats here if needed

        if (parsedTime && !isNaN(parsedTime)) {
            logToSheet_(functionName, `Successfully parsed time string "${timeString}" to Date object.`, "DEBUG");
            return parsedTime;
        } else {
            logToSheet_(functionName, `Could not parse time string "${timeString}" into a known format.`, "WARN");
            return null;
        }
    } catch (e) {
        logToSheet_(functionName, `Error parsing time string "${timeString}": ${e}`, "ERROR");
        return null;
    }
}
/**
 * Retrieves a value from Script Properties.
 * @param {string} key The property key.
 * @returns {string|null} The property value or null if not found.
 */
function getScriptProperty_(key) {
  const functionName = "getScriptProperty_";
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const value = scriptProps.getProperty(key);
    if (!value) {
      logToSheet_(functionName, `Property key "${key}" not found in Script Properties.`, "WARN");
      return null;
    }
    // Avoid logging sensitive values like passwords
    const logValue = key.toUpperCase().includes("PASSWORD") ? "*** HIDDEN ***" : value.substring(0,50) + (value.length > 50 ? "..." : "");
    logToSheet_(functionName, `Retrieved property "${key}": ${logValue}`, "DEBUG");
    return value;
  } catch (e) {
    logToSheet_(functionName, `Error retrieving script property "${key}": ${e.toString()}`, "ERROR");
    return null;
  }
}


/**
 * Logs messages to a dedicated sheet with fallback to Logger.log.
 * @param {string} functionName The name of the calling function.
 * @param {string} message The message to log.
 * @param {string} [level='INFO'] The logging level (e.g., INFO, WARN, ERROR, DEBUG, FATAL, ENTRY).
 */
function logToSheetLegacy_(functionName, message, level) {
  level = (level || "INFO").toUpperCase();
  const logSheetName = LOG_SHEET_NAME;
  const maxLogRows = 1500;
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss.SSS Z");
  let messageStr = String(message);
  if (messageStr.length > 49000) messageStr = messageStr.substring(0, 49000) + "... (truncated)";

  const logEntry = `[${level}] ${functionName}: ${messageStr}`;

  try {
    // --- Core Sheet Logging ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      // Fallback if spreadsheet is totally inaccessible
      Logger.log(`!!! SPREADSHEET INACCESSIBLE !!! ${logEntry}`);
      return;
    }

    let logSheet = ss.getSheetByName(logSheetName);

    // Check if sheet needs creation *after* confirming ss exists
    if (!logSheet) {
      try {
        logSheet = ss.insertSheet(logSheetName, 0); // Attempt creation
        if (!logSheet) {
           Logger.log(`!!! FAILED TO CREATE LOG SHEET '${logSheetName}' !!! ${logEntry}`);
           return;
        }
        // Set up headers only if creation succeeded
        const headers = [["Timestamp", "Level", "Function", "Message"]];
        logSheet.getRange("A1:D1").setValues(headers).setFontWeight("bold").setBackground("#eeeeee");
        logSheet.setFrozenRows(1);
        logSheet.setColumnWidth(1, 180); // Wider Timestamp
        logSheet.setColumnWidth(2, 80);  // Level
        logSheet.setColumnWidth(3, 200); // Function
        logSheet.setColumnWidth(4, 700); // Message
        // Log the initial message again now that the sheet exists
        logSheet.insertRowBefore(2);
        logSheet.getRange(2, 1, 1, 4).setValues([[timestamp, level, functionName, messageStr]]);
        // Apply formatting for the first message
         const levelCell = logSheet.getRange(2, 2);
         if (level === "ERROR" || level === "FATAL") levelCell.setFontColor("#a50e0e").setFontWeight("bold");
         else if (level === "WARN") levelCell.setFontColor("#e67e22");
         else if (level === "SUCCESS") levelCell.setFontColor("#137333");
         else levelCell.setFontColor("black").setFontWeight("normal");
        // Don't trim rows on the very first entry
        return; // Exit after setting up and logging first message

      } catch (sheetCreateError) {
         Logger.log(`!!! ERROR CREATING LOG SHEET '${logSheetName}': ${sheetCreateError} !!! ${logEntry}`);
         return; // Stop if sheet creation fails
      }
    }

    // --- Insert Log Row (if sheet exists) ---
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1, 1, 4).setValues([[timestamp, level, functionName, messageStr]]);

    // Apply formatting
    const levelCell = logSheet.getRange(2, 2);
    if (level === "ERROR" || level === "FATAL") levelCell.setFontColor("#a50e0e").setFontWeight("bold");
    else if (level === "WARN") levelCell.setFontColor("#e67e22");
    else if (level === "SUCCESS") levelCell.setFontColor("#137333");
    else levelCell.setFontColor("black").setFontWeight("normal");


    // --- Trim Old Logs ---
    // Check last row *after* inserting the new row
    const lastRow = logSheet.getLastRow();
    if (lastRow > (maxLogRows + 1)) { // +1 for the header row
         const rowsToDelete = lastRow - (maxLogRows + 1);
         // Check if rowsToDelete is valid before attempting deletion
         if (rowsToDelete > 0) {
             logSheet.deleteRows(maxLogRows + 2, rowsToDelete); // Delete from row *after* max rows + header
         }
    }
  // --- End Core Sheet Logging ---

  } catch (e) {
    // --- Fallback Logging ---
    // Log the original message AND the error that occurred during sheet logging
    Logger.log(`!!! SHEET LOGGING FAILED !!! ${logEntry}. Error during logging: ${e.toString()}` + (e.stack ? ` Stack: ${e.stack}` : ""));
  }
}
/**
 * Converts state names to abbreviations (add more as needed).
 * @param {string} stateName Full or abbreviated state name.
 * @returns {string} Uppercase state abbreviation.
 */
function abbreviateState_(stateName) {
  if (!stateName) return '';
  const name = stateName.trim().toUpperCase();
  // Add more mappings if Gemini returns full names often
  const stateMap = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR", "CALIFORNIA": "CA",
    "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE", "FLORIDA": "FL", "GEORGIA": "GA",
    "HAWAII": "HI", "IDAHO": "ID", "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA",
    "KANSAS": "KS", "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME", "MARYLAND": "MD",
    "MASSACHUSETTS": "MA", "MICHIGAN": "MI", "MINNESOTA": "MN", "MISSISSIPPI": "MS",
    "MISSOURI": "MO", "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV", "NEW HAMPSHIRE": "NH",
    "NEW JERSEY": "NJ", "NEW MEXICO": "NM", "NEW YORK": "NY", "NORTH CAROLINA": "NC",
    "NORTH DAKOTA": "ND", "OHIO": "OH", "OKLAHOMA": "OK", "OREGON": "OR", "PENNSYLVANIA": "PA",
    "RHODE ISLAND": "RI", "SOUTH CAROLINA": "SC", "SOUTH DAKOTA": "SD", "TENNESSEE": "TN",
    "TEXAS": "TX", "UTAH": "UT", "VERMONT": "VT", "VIRGINIA": "VA", "WASHINGTON": "WA",
    "WEST VIRGINIA": "WV", "WISCONSIN": "WI", "WYOMING": "WY",
    "DISTRICT OF COLUMBIA": "DC",
    // Canadian provinces
    "ALBERTA": "AB", "BRITISH COLUMBIA": "BC", "MANITOBA": "MB", "NEW BRUNSWICK": "NB",
    "NEWFOUNDLAND AND LABRADOR": "NL", "NOVA SCOTIA": "NS", "ONTARIO": "ON",
    "PRINCE EDWARD ISLAND": "PE", "QUEBEC": "QC", "SASKATCHEWAN": "SK"
  };
  // Return abbreviation if found in map, otherwise return original (if already abbr) or original
  return stateMap[name] || (name.length === 2 ? name : stateName.toUpperCase()); // Ensure return is uppercase
}


/**
 * Formats a date string (YYYY-MM-DD) or Date object into MM/DD/YYYY for the sheet.
 * **REVISED to prevent timezone shift for YYYY-MM-DD strings.**
 * @param {string|Date} dateInput The date value. Expected string format is YYYY-MM-DD.
 * @returns {string} Formatted date string "MM/DD/YYYY" or original if invalid.
 */
function formatSheetDate_(dateInput) {
  if (!dateInput) return '';
  try {
    // Prioritize handling the YYYY-MM-DD string directly to avoid timezone issues
    if (typeof dateInput === 'string' && dateInput.match(/^\d{4}-\d{2}-\d{2}$/)) {
      const parts = dateInput.split('-'); // [YYYY, MM, DD]
      // Directly rearrange the parts into MM/DD/YYYY format.
      // This avoids creating a Date object based on UTC midnight,
      // which can shift to the previous day when formatted in local timezones behind UTC.
      return `${parts[1]}/${parts[2]}/${parts[0]}`;
    }
    // Handle cases where a Date object might genuinely be passed
    else if (Object.prototype.toString.call(dateInput) === "[object Date]" && !isNaN(dateInput)) {
      // If it's already a valid Date object, format it using the script's timezone.
      // This assumes the Date object correctly represents the intended point in time.
      return Utilities.formatDate(dateInput, Session.getScriptTimeZone(), "MM/dd/yyyy");
    }
    // If input is neither the expected string format nor a valid Date object
    else {
      logToSheet_("formatSheetDate_", `Input '${dateInput}' is not a valid YYYY-MM-DD string or Date object. Returning original.`, "WARN");
      return String(dateInput); // Return original as string
    }
  } catch (e) {
    logToSheet_("formatSheetDate_", `Error formatting date '${dateInput}': ${e}`, "WARN");
    return String(dateInput); // Return original on error
  }
}


/**
 * Formats a time string (HH:MM:SS or HH:MM) or Date object into h:mm A/PM format.
 * @param {string|Date} timeInput The time value.
 * @returns {string} Formatted time string or original if invalid.
 */
function formatSheetTime_(timeInput) {
  if (!timeInput) return '';
  try {
    let dateObj;
     // Handle potential Date object input first
    if (Object.prototype.toString.call(timeInput) === "[object Date]" && !isNaN(timeInput)) {
      dateObj = timeInput;
    } else if (typeof timeInput === 'string') {
      // Attempt to parse HH:MM:SS or HH:MM string
      const timeParts = timeInput.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/); // Match HH:MM or HH:MM:SS
      if (timeParts) {
         // Create a dummy date object just for time formatting
         dateObj = new Date(1970, 0, 1, parseInt(timeParts[1]), parseInt(timeParts[2]), timeParts[3] ? parseInt(timeParts[3]) : 0);
      } else {
         logToSheet_("formatSheetTime_", `Input string '${timeInput}' is not a valid HH:MM(:SS) format.`, "WARN");
         return timeInput; // Invalid time string format
      }
    } else {
       logToSheet_("formatSheetTime_", `Input '${timeInput}' is not a valid string or Date object.`, "WARN");
       return timeInput; // Return original if not string or Date object
    }
     // Format the time using the script's timezone
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "h:mm a").toUpperCase(); // Format as H:MM AM/PM
  } catch (e) {
    logToSheet_("formatSheetTime_", `Error formatting time '${timeInput}': ${e}`, "WARN");
    return timeInput; // Return original on error
  }
}


/**
 * Applies Proper Case to a string, handling specific exceptions like directions.
 * @param {string} str The input string.
 * @param {string[]} exceptions Lowercase strings to keep uppercase (e.g., ['eb', 'nb', 'sb', 'wb', 'sw', 'se', 'nw', 'ne']).
 * @returns {string} The formatted string.
 */
function toProperCase_(str, exceptions = []) {
  if (!str) return '';
  const exceptionsUpper = exceptions.map(e => e.toUpperCase()); // Ensure exceptions are uppercase for comparison
  return str.toLowerCase().replace(/([^\w']|_)+/g, ' ').trim() // Normalize spaces and handle non-word chars
     .split(' ').map(word => {
    if (exceptionsUpper.includes(word.toUpperCase())) {
      return word.toUpperCase(); // Keep specified exceptions uppercase
    }
    if (word.length > 0) {
      // Handle words with internal caps (like McDonald's) - this logic keeps them as is after lowercasing
      // Basic Proper Case:
      return word.charAt(0).toUpperCase() + word.slice(1);
    }
    return word;
  }).join(' ');
}


/**
 * Formats the violation type: Applies Proper Case and attempts simple expansions.
 * Relies on Gemini prompt for initial expansion.
 * @param {string} typeStr The raw violation type string (potentially expanded by Gemini).
 * @returns {string} The formatted string.
 */
function formatViolationType_(typeStr) {
    if (!typeStr) return '';
    // Apply Proper Case after Gemini's potential expansion
    return toProperCase_(typeStr);
}


/**
 * Formats the violation location using Proper Case with directional exceptions.
 * @param {string} locationStr The raw location string.
 * @returns {string} The formatted string.
 */
function formatLocation_(locationStr) {
    if (!locationStr) return '';
    const directions = ['EB', 'NB', 'SB', 'WB', 'SW', 'SE', 'NW', 'NE', 'N', 'S', 'E', 'W']; // Include single letters too
    // Use the existing toProperCase_ but pass the directions as exceptions
    return toProperCase_(locationStr, directions);
}




