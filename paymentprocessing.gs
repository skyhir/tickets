/**
 * Helper to URL-encode an object for Stripe API payloads.
 * Handles simple key-value pairs and basic metadata.
 */
function toUrlEncoded_(obj) {
    var params = [];
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) {
            var value = obj[key];
            if (typeof value === 'object' && value !== null && key === 'metadata') {
                // Handle metadata specifically
                for (var metaKey in value) {
                    if (value.hasOwnProperty(metaKey)) {
                        params.push(encodeURIComponent('metadata[' + metaKey + ']') + '=' + encodeURIComponent(value[metaKey]));
                    }
                }
            } else if (value !== undefined && value !== null) { // Only include defined, non-null values
                params.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
            }
        }
    }
    return params.join('&');
}

// Cascade limits for attemptStripeCharge_. A driver with a long wallet history can
// have a dozen saved cards; walking all of them turns one unpaid citation into a
// dozen declines on the same customer, which is what trips Stripe's card-testing
// protection. Five is enough to clear a stale primary card without that risk.
const MAX_CARD_ATTEMPTS = 5;

// Declines where the issuer is telling us to stop, not to try harder. Continuing
// down the wallet after one of these is what gets a customer flagged.
const HARD_DECLINE_CODES = [
  'stolen_card',
  'lost_card',
  'fraudulent',
  'pickup_card',
  'restricted_card',
  'revocation_of_all_authorizations',
  'merchant_blacklist'
];

/**
 * Reads the customer's default payment method so the cascade starts with the card
 * the driver actually expects to be charged.
 *
 * Stripe's /v1/payment_methods list is ordered by creation date, NOT by default, so
 * without this the "first" card is simply the most recently added one. A failure here
 * is non-fatal — we fall back to creation order.
 *
 * @param {string} stripeKey The Stripe Secret Key.
 * @param {string} stripeCustomerId The Stripe Customer ID ('cus_...').
 * @returns {string|null} The default payment method id, or null if unset/unavailable.
 */
function getDefaultPaymentMethodId_(stripeKey, stripeCustomerId) {
  const functionName = "getDefaultPaymentMethodId_";
  try {
    const response = UrlFetchApp.fetch(`https://api.stripe.com/v1/customers/${encodeURIComponent(stripeCustomerId)}`, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + stripeKey },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      logToSheet_(functionName, `Could not read customer ${stripeCustomerId} (Code: ${response.getResponseCode()}). Falling back to creation order.`, "WARN");
      return null;
    }
    const customer = JSON.parse(response.getContentText());
    const defaultPmId = (customer.invoice_settings && customer.invoice_settings.default_payment_method) || null;
    logToSheet_(functionName, defaultPmId
      ? `Default payment method for ${stripeCustomerId}: ${defaultPmId}`
      : `No default payment method set for ${stripeCustomerId}. Using creation order.`, "DEBUG");
    return defaultPmId;
  } catch (e) {
    logToSheet_(functionName, `Error reading default payment method for ${stripeCustomerId}: ${e.message}. Falling back to creation order.`, "WARN");
    return null;
  }
}

/**
 * Turns the raw Stripe payment-method list into the ordered list of cards worth charging.
 *
 * Order: default card first, then newest-first. Duplicates are collapsed on
 * card.fingerprint — drivers routinely re-add the same physical card, and charging it
 * three times only produces three identical declines. Cards already past their
 * expiry are dropped, unless every card on file is expired, in which case the newest
 * one is still attempted so the row lands in Collections with a real Stripe decline
 * behind it rather than no record at all.
 *
 * @param {Array} pmData The `data` array from /v1/payment_methods.
 * @param {string|null} defaultPmId The customer's default payment method id, if any.
 * @param {string} callerName Function name to attribute log lines to.
 * @returns {Array<{id: string, label: string, isDefault: boolean}>} Cards to attempt, in order.
 */
function selectChargeCandidates_(pmData, defaultPmId, callerName) {
  const functionName = callerName || "selectChargeCandidates_";
  const now = new Date();
  const nowYear = now.getFullYear();
  const nowMonth = now.getMonth() + 1; // 1-based, to match Stripe's exp_month

  const describe = function (pm) {
    const card = pm.card || {};
    const brand = card.brand ? String(card.brand) : 'card';
    const last4 = card.last4 ? `••${card.last4}` : pm.id;
    const exp = (card.exp_month && card.exp_year) ? ` exp ${card.exp_month}/${card.exp_year}` : '';
    return `${brand} ${last4}${exp}`;
  };

  const isExpired = function (pm) {
    const card = pm.card || {};
    if (!card.exp_year || !card.exp_month) return false; // unknown expiry: give it a shot
    if (card.exp_year < nowYear) return true;
    return card.exp_year === nowYear && card.exp_month < nowMonth;
  };

  // Newest-first (Stripe already returns this order; sorting makes it explicit),
  // then hoist the default card to the front.
  const ordered = pmData.slice().sort(function (a, b) { return (b.created || 0) - (a.created || 0); });
  if (defaultPmId) {
    const defaultIdx = ordered.findIndex(function (pm) { return pm.id === defaultPmId; });
    if (defaultIdx > 0) ordered.unshift(ordered.splice(defaultIdx, 1)[0]);
  }

  // Collapse duplicates. Keeping the FIRST occurrence means the default card wins
  // over its own re-added copies. Cards with no fingerprint key on their own id.
  const seenFingerprints = {};
  const unique = [];
  let duplicatesDropped = 0;
  for (const pm of ordered) {
    const key = (pm.card && pm.card.fingerprint) ? `fp:${pm.card.fingerprint}` : `id:${pm.id}`;
    if (seenFingerprints[key]) { duplicatesDropped++; continue; }
    seenFingerprints[key] = true;
    unique.push(pm);
  }
  if (duplicatesDropped > 0) {
    logToSheet_(functionName, `Collapsed ${duplicatesDropped} duplicate card(s) by fingerprint.`, "DEBUG");
  }

  const fresh = unique.filter(function (pm) { return !isExpired(pm); });
  const expired = unique.filter(isExpired);

  let chosen = fresh;
  if (fresh.length === 0 && expired.length > 0) {
    logToSheet_(functionName, `Every card on file is expired. Attempting the newest one (${describe(expired[0])}) anyway so the decline is recorded in Stripe.`, "WARN");
    chosen = [expired[0]];
  } else if (expired.length > 0) {
    logToSheet_(functionName, `Skipping ${expired.length} expired card(s).`, "DEBUG");
  }

  return chosen.map(function (pm) {
    return { id: pm.id, label: describe(pm), isDefault: pm.id === defaultPmId };
  });
}

/**
 * Resolves the "Date/Time Driver Paid" column (AD) for a given sheet.
 *
 * Prefers a header-text match on row 1 so the stamp follows the column if it is
 * ever moved or inserted at a different index on one of the tabs; falls back to
 * the declared COL.PAID_TIMESTAMP index. Returns null (and logs) when the sheet
 * has no such column, so callers can skip stamping instead of writing to the
 * wrong cell.
 *
 * @param {Sheet} sheet The violations tab.
 * @returns {number|null} 1-based column index, or null if unavailable.
 */
function resolvePaidTimestampCol_(sheet) {
  const functionName = "resolvePaidTimestampCol_";
  const maxCols = sheet.getMaxColumns();

  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      const wanted = PAID_TIMESTAMP_HEADER.replace(/\s+/g, ' ').trim().toLowerCase();
      for (let c = 0; c < headers.length; c++) {
        const h = (headers[c] || '').toString().replace(/\s+/g, ' ').trim().toLowerCase();
        if (h === wanted) return c + 1;
      }
    }
  } catch (e) {
    logToSheet_(functionName, `Header scan failed on "${sheet.getName()}": ${e}. Falling back to index ${COL.PAID_TIMESTAMP}.`, "WARN");
  }

  if (COL.PAID_TIMESTAMP <= maxCols) {
    logToSheet_(functionName, `No "${PAID_TIMESTAMP_HEADER}" header on "${sheet.getName()}"; falling back to column ${COL.PAID_TIMESTAMP}.`, "WARN");
    return COL.PAID_TIMESTAMP;
  }

  logToSheet_(functionName, `"${sheet.getName()}" has no "${PAID_TIMESTAMP_HEADER}" column and only ${maxCols} columns. Paid timestamps will not be recorded for this tab.`, "WARN");
  return null;
}

/**
 * Processes payments for tickets marked with 'Attempt' status in Column O (DRIVER_BILLING_STATUS).
 * **REVISED: Filters rows first for efficiency and reduced logging.**
 * Successful charges stamp the time of the charge into "Date/Time Driver Paid" (AD).
 */
function processTicketPayments() {
  const functionName = "processTicketPayments";
  const ui = SpreadsheetApp.getUi();
  logToSheet_(functionName, "Starting payment processing run.", "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = getViolationsSheets_(ss);
  if (sheets.length === 0) {
    ui.alert(`Error: Neither target sheet ('${TARGET_SHEET_NAME}' or '${TOLL_TARGET_SHEET_NAME}') was found.`);
    logToSheet_(functionName, "No target sheets found.", "ERROR");
    return;
  }

  const TEST_CUSTOMER_ID = "cus_QPOtpj7JZThQK1"; // Test customer ID for fee waiver

  // --- Column Indices ---
  const headerRows = 1;
  const billingStatusColIndex = COL.DRIVER_BILLING_STATUS - 1; // Column O (0-based index 14)
  const bookingIdColIndex = COL.BOOKING_ID - 1;
  const originalPenaltyColIndex = COL.ORIGINAL_PENALTY - 1;
  const totalOwedCol = COL.TOTAL_OWED_COLLECTIONS; // Column AA

  // "Date/Time Driver Paid" (AD) resolved once per tab, keyed by sheet name.
  const paidTsColBySheet = {};
  const paidTsColFor = function (sheet) {
    const key = sheet.getName();
    if (!(key in paidTsColBySheet)) {
      paidTsColBySheet[key] = resolvePaidTimestampCol_(sheet);
    }
    return paidTsColBySheet[key];
  };

  // --- Filter Attempt rows across every violations tab first ---
  // Each item carries its sheet so updates write back to the tab the row came from.
  const rowsToProcess = [];
  for (const sheet of sheets) {
    const allData = sheet.getDataRange().getValues();
    logToSheet_(functionName, `Read ${allData.length} total rows from "${sheet.getName()}".`, "DEBUG");
    for (let i = headerRows; i < allData.length; i++) {
      const rowData = allData[i];
      if (!rowData || rowData.length <= billingStatusColIndex) {
        logToSheet_(functionName, `"${sheet.getName()}" Row ${i + 1}: Skipping pre-filter check - Row is invalid or too short.`, "WARN");
        continue;
      }
      const statusRaw = rowData[billingStatusColIndex];
      const statusProcessed = String(statusRaw).trim().toUpperCase();
      if (statusProcessed === "ATTEMPT") {
        rowsToProcess.push({ sheet: sheet, data: rowData, originalRowIndex: i + 1 });
      }
    }
  }

  const numRowsToProcess = rowsToProcess.length;
  logToSheet_(functionName, `Found ${numRowsToProcess} row(s) with status 'Attempt' across ${sheets.length} tab(s).`, "INFO");

  if (numRowsToProcess === 0) {
    logToSheet_(functionName, "No rows found requiring payment processing.", "INFO");
    ui.alert("No rows found with status 'Attempt'. Processing complete.");
    return;
  }

  // --- Get Stripe Key (only if rows need processing) ---
  const stripeKey = getStripeSecretKey();
  if (!stripeKey) {
     ui.alert("Error: Could not retrieve Stripe Secret Key. Check logs and Secret Manager setup.");
     logToSheet_(functionName, "Stripe key retrieval failed.", "ERROR");
     return;
  }
  logToSheet_(functionName, "Stripe key retrieved successfully.", "INFO");

  let successCount = 0;
  let collectionCount = 0;
  let errorCount = 0;

  // --- ** Process ONLY the Filtered Rows ** ---
  for (let j = 0; j < rowsToProcess.length; j++) {
    const item = rowsToProcess[j];
    const sheet = item.sheet;
    const row = item.data; // The actual row data array
    const actualRowNum = item.originalRowIndex; // The original row number on the sheet

    logToSheet_(functionName, `Processing filtered row ${j + 1} of ${numRowsToProcess} ("${sheet.getName()}" Row ${actualRowNum}). Status: 'Attempt'.`, "INFO");

    const bookingId = row[bookingIdColIndex];
    const originalPenalty = parseFloat(row[originalPenaltyColIndex]) || 0;

    // Validate required data
    if (!bookingId || originalPenalty < 0) {
      logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Skipping payment - Missing Booking ID ('${bookingId}') or negative Original Penalty ($${originalPenalty}).`, "WARN");
      sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Error - Missing/Invalid Data");
      errorCount++;
      continue; // Go to next filtered row
    }

    // Lookup Stripe Customer ID
    logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Looking up Stripe ID for Booking ID ${bookingId}...`, "DEBUG");
    sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Processing Payment..."); // Update status on sheet
    SpreadsheetApp.flush();

    const stripeCustomerId = getStripeCustomerIdFromBooking_(bookingId);

    if (!stripeCustomerId || typeof stripeCustomerId !== 'string' || !stripeCustomerId.startsWith('cus_')) {
        logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Stripe Customer ID lookup failed or invalid for Booking ID ${bookingId}. Result: ${stripeCustomerId}`, "ERROR");
        sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Error - No Stripe ID");

        let adminFeeOnFailure = 25.00;
        if (stripeCustomerId === TEST_CUSTOMER_ID) { // Check even if format is wrong, maybe ID was returned
           adminFeeOnFailure = 0.00;
           logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Waiving $25 admin fee for TEST_CUSTOMER_ID during Stripe ID lookup failure.`, "INFO");
        }
        const totalOwed = originalPenalty + adminFeeOnFailure;

        if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
           sheet.getRange(actualRowNum, totalOwedCol).setValue(totalOwed).setNumberFormat("$#,##0.00");
        } else {
           logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Invalid TOTAL_OWED_COLLECTIONS column index (${totalOwedCol}). Cannot write owed amount.`, "ERROR");
        }
        errorCount++;
        continue; // Go to next filtered row
    }
    logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Found Stripe Customer ID: ${stripeCustomerId}`, "DEBUG");

    // Fee Waiver Logic
    let adminFee = 25.00;
    if (stripeCustomerId === TEST_CUSTOMER_ID) {
        adminFee = 0.00;
        logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Waiving $25 admin fee for test customer ${TEST_CUSTOMER_ID}.`, "INFO");
    }

    // Calculate amount and handle $0 cases
    const amountToCharge = originalPenalty + adminFee;
     if (amountToCharge <= 0 && originalPenalty <= 0) {
       logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Calculated amount is $${amountToCharge.toFixed(2)}. Skipping charge attempt for non-positive amount. Marking as 'Paid (Amount $0)'.`, "WARN");
       sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Paid (Amount $0)");
       if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
         sheet.getRange(actualRowNum, totalOwedCol).setValue(0).setNumberFormat("$#,##0.00");
       }
       successCount++; // Count $0 charge as success
       continue; // Go to next filtered row
     }

    // Attempt Stripe Charge
    const amountInCents = Math.round(amountToCharge * 100);
    const description = `52065 Vehicle Toll Roads/Citations/Impound: Booking ${bookingId}`;
    logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Attempting Stripe charge for $${amountToCharge.toFixed(2)} (${amountInCents} cents)...`, "DEBUG");
    const paymentResult = attemptStripeCharge_(stripeKey, stripeCustomerId, amountInCents, description);

    // Update Sheet Based on Result
    if (paymentResult.success) {
      logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Payment successful. PI ID: ${paymentResult.paymentIntentId}`, "SUCCESS");
      sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Paid");
      if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
        sheet.getRange(actualRowNum, totalOwedCol).setValue(0).setNumberFormat("$#,##0.00");
      }
      // Stamp when the charge actually cleared. Only successful Stripe charges get
      // a timestamp — declines and $0 rows are left blank on purpose.
      const paidTsCol = paidTsColFor(sheet);
      if (paidTsCol) {
        sheet.getRange(actualRowNum, paidTsCol)
             .setValue(new Date())
             .setNumberFormat(PAID_TIMESTAMP_FORMAT);
      }
      successCount++;
    } else {
      logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Payment failed. Reason: ${paymentResult.message}`, "WARN");
      sheet.getRange(actualRowNum, COL.DRIVER_BILLING_STATUS).setValue("Collections");
      // Clear any stale paid timestamp: a row re-run as 'Attempt' after an earlier
      // success must not keep claiming the driver paid when this charge declined.
      const failedTsCol = paidTsColFor(sheet);
      if (failedTsCol) {
        const tsCell = sheet.getRange(actualRowNum, failedTsCol);
        if ((tsCell.getDisplayValue() || '').trim()) {
          logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Clearing stale paid timestamp after declined charge.`, "WARN");
          tsCell.clearContent();
        }
      }
       if (totalOwedCol && totalOwedCol > 0 && totalOwedCol <= sheet.getMaxColumns()) {
          sheet.getRange(actualRowNum, totalOwedCol).setValue(amountToCharge).setNumberFormat("$#,##0.00");
       } else {
           logToSheet_(functionName, `"${sheet.getName()}" Row ${actualRowNum}: Invalid TOTAL_OWED_COLLECTIONS column index (${totalOwedCol}). Cannot write owed amount.`, "ERROR");
       }
      collectionCount++;
    }
     SpreadsheetApp.flush(); // Flush after each processed row

  } // --- End filtered row loop ---

  // Final Summary
  logToSheet_(functionName, `Payment processing complete. Rows Found with 'Attempt': ${numRowsToProcess}, Successful Payments: ${successCount}, Sent to Collections: ${collectionCount}, Errors (Setup/Data/Lookup): ${errorCount}`, "INFO");
  ui.alert(`Payment Processing Complete.\n\nRows Found with 'Attempt': ${numRowsToProcess}\nSuccessful Payments: ${successCount}\nSent to Collections: ${collectionCount}\nErrors (Setup/Data/Lookup): ${errorCount}`);
}
/**
 * Looks up the Stripe Customer ID (wallet_id) from the database using the booking ref_number.
 * @param {string} bookingId The Envoy booking ref_number.
 * @returns {string|null} The Stripe Customer ID ('cus_...') or null if not found/error.
 */
function getStripeCustomerIdFromBooking_(bookingId) {
  const functionName = "getStripeCustomerIdFromBooking_";
  if (!bookingId) {
    logToSheet_(functionName, "Booking ID is missing.", "WARN");
    return null;
  }

  let conn = null;
  let stmt1 = null, rs1 = null;
  let stmt2 = null, rs2 = null;
  let userId = null;
  let stripeCustomerId = null;

  try {
    // --- Get DB Credentials --- (Ensure getScriptProperty_ works)
    const dbHost = getScriptProperty_('DB_HOST');
    const dbPort = getScriptProperty_('DB_PORT');
    const dbName = getScriptProperty_('DB_NAME');
    const dbUser = getScriptProperty_('DB_USER');
    const dbPassword = getScriptProperty_('DB_PASSWORD');
    if (!dbHost || !dbPort || !dbName || !dbUser || !dbPassword) {
      throw new Error("DB credentials missing.");
    }

    const dbUrl = `jdbc:mysql://${dbHost}:${dbPort}/${dbName}`;
    conn = Jdbc.getConnection(dbUrl, dbUser, dbPassword);
    conn.setAutoCommit(false);
    logToSheet_(functionName, `DB connected for Stripe ID lookup (Booking: ${bookingId})`, "DEBUG");

    // --- Query 1: Get User_ID from Envoy_Booking_View ---
    const sql1 = "SELECT User_ID FROM Envoy_Booking_View WHERE ref_number = ?";
    stmt1 = conn.prepareStatement(sql1);
    stmt1.setString(1, bookingId);
    rs1 = stmt1.executeQuery();

    if (rs1.next()) {
      userId = rs1.getString("User_ID");
      logToSheet_(functionName, `Found User_ID ${userId} for Booking ${bookingId}`, "DEBUG");
    } else {
      logToSheet_(functionName, `No User_ID found for Booking ID ${bookingId} in Envoy_Booking_View.`, "WARN");
      return null; // Can't proceed without User_ID
    }

    // --- Query 2: Get wallet_id from users table using User_ID ---
    // CRITICAL: Verify the column name in the 'users' table that corresponds to User_ID from Envoy_Booking_View.
    // Assuming it's 'user_id' as per your description. Adjust if different.
    const usersTableLookupColumn = 'user_id'; // <<< VERIFY AND CHANGE IF NEEDED
    if (userId) {
      const sql2 = `SELECT wallet_id FROM users WHERE ${usersTableLookupColumn} = ?`;
      stmt2 = conn.prepareStatement(sql2);
      stmt2.setString(1, userId); // Use the User_ID obtained from query 1
      logToSheet_(functionName, `Querying users table: ${sql2} with param: ${userId}`, "DEBUG");
      rs2 = stmt2.executeQuery();

      if (rs2.next()) {
        stripeCustomerId = rs2.getString("wallet_id");
        if (stripeCustomerId && String(stripeCustomerId).trim().startsWith('cus_')) {
           stripeCustomerId = String(stripeCustomerId).trim(); // Trim whitespace
           logToSheet_(functionName, `Found Stripe Customer ID (wallet_id) ${stripeCustomerId} for User_ID ${userId}`, "SUCCESS");
        } else {
           logToSheet_(functionName, `Found wallet_id for User_ID ${userId}, but it's invalid or not a Stripe Customer ID: '${stripeCustomerId}'`, "WARN");
           stripeCustomerId = null; // Invalid ID
        }
      } else {
        logToSheet_(functionName, `No record found in 'users' table for ${usersTableLookupColumn} = ${userId}.`, "WARN");
      }
    }

    return stripeCustomerId;

  } catch (e) {
    logToSheet_(functionName, `DB Error looking up Stripe ID for Booking ${bookingId}: ${e.toString()}`+ (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
    return null;
  } finally {
    closeQuietly_(rs1); closeQuietly_(stmt1);
    closeQuietly_(rs2); closeQuietly_(stmt2);
    closeQuietly_(conn);
     logToSheet_(functionName, `DB resources closed for Stripe ID lookup (Booking: ${bookingId})`, "DEBUG");
  }
}

/**
 * Attempts to create and confirm a Stripe Payment Intent off_session, cascading
 * through the customer's saved cards until one succeeds.
 *
 * Cards are tried default-first then newest-first (see selectChargeCandidates_),
 * deduped by fingerprint and filtered of expired cards. A soft decline moves to the
 * next card; a hard decline, a non-card error, or an exception of unknown outcome
 * stops the cascade so the row can fall to Collections. Returning success:false only
 * after the wallet is exhausted keeps the caller's Paid/Collections logic unchanged.
 *
 * @param {string} stripeKey The Stripe Secret Key.
 * @param {string} stripeCustomerId The Stripe Customer ID ('cus_...').
 * @param {number} amountInCents The amount to charge in cents.
 * @param {string} description The description for the charge.
 * @returns {object} { success: boolean, message: string, paymentIntentId: string|null, cardsTried: number }
 */
function attemptStripeCharge_(stripeKey, stripeCustomerId, amountInCents, description) {
    const functionName = "attemptStripeCharge_";
    let paymentSuccess = false;
    let paymentResultMessage = "Payment initialization failed.";
    let paymentIntentId = null; // Store the ID of the intent, successful or not

    if (!stripeKey || !stripeCustomerId || !stripeCustomerId.startsWith('cus_')) {
         logToSheet_(functionName, "Invalid parameters: Missing Stripe Key or invalid Customer ID.", "ERROR");
         return { success: false, message: "Internal Error: Invalid Stripe Key or Customer ID provided.", paymentIntentId: null };
    }
    if (amountInCents <= 0) {
        logToSheet_(functionName, `Invalid amount: ${amountInCents} cents. Amount must be positive.`, "ERROR");
        // If penalty was 0 and fee was waived, amount could be 0. Allow this? Or treat as failure?
        // Let's treat amount <= 0 as something not to charge, maybe success=true? Or failure?
        // For now, let's consider 0 amount a failure to charge, consistent with original logic.
        // If originalPenalty could be negative, this needs more thought. Assuming non-negative penalty.
        return { success: false, message: "Amount must be positive.", paymentIntentId: null };
    }

    try {
        // --- Fetch Customer's Saved Card Payment Methods ---
        logToSheet_(functionName, `Fetching 'card' payment methods for customer: ${stripeCustomerId}`, "DEBUG");
        const pmFetchUrl = `https://api.stripe.com/v1/payment_methods?customer=${stripeCustomerId}&type=card`;
        const pmOptions = {
            method: 'get',
            headers: { Authorization: 'Bearer ' + stripeKey },
            muteHttpExceptions: true
        };
        const pmResponse = UrlFetchApp.fetch(pmFetchUrl, pmOptions);
        const pmResponseCode = pmResponse.getResponseCode();
        const pmResponseText = pmResponse.getContentText();

        if (pmResponseCode !== 200) {
            logToSheet_(functionName, `Error fetching Stripe PMs (Code: ${pmResponseCode}, Customer: ${stripeCustomerId}). Response: ${pmResponseText.substring(0,300)}`, "ERROR");
            if (pmResponseCode === 404) throw new Error(`Stripe Customer ID '${stripeCustomerId}' not found in Stripe.`);
            throw new Error(`Error fetching Stripe payment methods (Code: ${pmResponseCode}).`);
        }

        const pmResult = JSON.parse(pmResponseText);
        if (!pmResult.data || pmResult.data.length === 0) {
            logToSheet_(functionName, `No saved cards found for Stripe customer ${stripeCustomerId}.`, "WARN");
            return { success: false, message: "No saved cards found for customer.", paymentIntentId: null };
        }

        // --- Order the wallet: default card first, then newest-first ---
        const defaultPmId = getDefaultPaymentMethodId_(stripeKey, stripeCustomerId);
        const candidates = selectChargeCandidates_(pmResult.data, defaultPmId, functionName);

        if (candidates.length === 0) {
            logToSheet_(functionName, `Customer ${stripeCustomerId} has cards on file but none were usable.`, "WARN");
            return { success: false, message: "No usable cards found for customer.", paymentIntentId: null };
        }

        const attempts = candidates.slice(0, MAX_CARD_ATTEMPTS);
        if (candidates.length > attempts.length) {
            logToSheet_(functionName, `Customer ${stripeCustomerId} has ${candidates.length} usable card(s); capping this run at ${MAX_CARD_ATTEMPTS}. Cards not attempted: ${candidates.length - attempts.length}.`, "WARN");
        }

        // --- Attempt Payment Intent Creation & Confirmation, card by card ---
        logToSheet_(functionName, `Attempting Payment Intent (${amountInCents} cents) for Customer ${stripeCustomerId} across up to ${attempts.length} card(s).`, "INFO");

        let cardsTried = 0;
        let lastFailureReason = "No card produced a result.";

        for (let a = 0; a < attempts.length; a++) {
            const candidate = attempts[a];
            const pmId = candidate.id;
            cardsTried++;
            logToSheet_(functionName, `Card ${cardsTried} of ${attempts.length}: attempting charge with ${candidate.label}${candidate.isDefault ? ' [default]' : ''}.`, "INFO");
            paymentIntentId = null; // Reset PI ID for this attempt

            try {
                const piPayload = {
                    amount: amountInCents,
                    currency: 'usd',
                    customer: stripeCustomerId,
                    payment_method: pmId,
                    description: description,
                    confirm: true,
                    off_session: true
                };

                const piOptions = {
                    method: 'post',
                    headers: {
                        Authorization: 'Bearer ' + stripeKey,
                        'Content-Type': 'application/x-www-form-urlencoded',
                        'Stripe-Version': '2024-04-10' // Specify API version
                    },
                    payload: toUrlEncoded_(piPayload),
                    muteHttpExceptions: true
                };

                const piResponse = UrlFetchApp.fetch("https://api.stripe.com/v1/payment_intents", piOptions);
                const piResponseCode = piResponse.getResponseCode();
                const piResponseText = piResponse.getContentText();
                const piResult = JSON.parse(piResponseText);

                // Extract PI ID regardless of success/failure for logging
                if (piResult && piResult.id) {
                   paymentIntentId = piResult.id;
                } else if (piResult && piResult.error && piResult.error.payment_intent && piResult.error.payment_intent.id) {
                   paymentIntentId = piResult.error.payment_intent.id; // ID might be in error object
                }
                logToSheet_(functionName, `PI Attempt (PM: ${pmId}, PI: ${paymentIntentId || 'N/A'}) - Code: ${piResponseCode}, Status: ${piResult.status || (piResult.error ? 'Error' : 'Unknown')}`, "DEBUG");

                if (piResult.status === "succeeded") {
                    paymentSuccess = true;
                    paymentResultMessage = `Payment successful on card ${cardsTried} of ${attempts.length} (${candidate.label}).`;
                    logToSheet_(functionName, `${paymentResultMessage} PI ID: ${paymentIntentId}`, "SUCCESS");
                    return { success: true, message: paymentResultMessage, paymentIntentId: paymentIntentId, cardsTried: cardsTried };
                }

                // --- Not successful. Decide whether the NEXT card is worth trying. ---
                // Stripe reports a declined card either as a 402 `error` (card_error) or,
                // on a 200, as `last_payment_error` alongside a non-succeeded status.
                const err = piResult.error || piResult.last_payment_error || null;
                const errType = err ? err.type : null;
                const declineCode = err ? err.decline_code : null;

                let failureReason = "Unknown status or error";
                if (err) {
                    failureReason = err.message || `Code: ${err.code || 'n/a'}`;
                    if (declineCode) failureReason += ` (decline_code: ${declineCode})`;
                } else if (piResult.status) {
                    failureReason = `Status: ${piResult.status}`; // e.g. requires_action, requires_payment_method
                }
                lastFailureReason = failureReason;
                logToSheet_(functionName, `Card ${cardsTried} (${candidate.label}) declined: ${failureReason}${paymentIntentId ? ` (PI ID: ${paymentIntentId})` : ""}`, "WARN");

                // Hard declines mean the card is compromised or the issuer wants it taken.
                // Continuing to hammer the same customer risks tripping Stripe's card-testing
                // block, so stop the cascade for this row and let it fall to Collections.
                if (declineCode && HARD_DECLINE_CODES.indexOf(declineCode) !== -1) {
                    logToSheet_(functionName, `Hard decline (${declineCode}) on ${candidate.label}. Stopping cascade for customer ${stripeCustomerId}.`, "WARN");
                    break;
                }

                // A non-card error (bad amount, bad key, customer gone, rate limit) will fail
                // identically on every other card. Trying the rest just burns quota.
                if (errType && errType !== 'card_error') {
                    logToSheet_(functionName, `Non-card error (${errType}) on ${candidate.label}. Stopping cascade for customer ${stripeCustomerId}.`, "ERROR");
                    break;
                }

                // Soft decline (insufficient_funds, expired_card, authentication_required, ...)
                // — move on to the next card on file.

            } catch (e) {
                // muteHttpExceptions swallows HTTP errors, so reaching here means the request
                // or the response parse itself blew up and we do NOT know whether Stripe
                // processed the charge. Stop rather than risk charging a second card.
                lastFailureReason = `Exception during payment attempt for ${candidate.label}: ${e.message}`;
                logToSheet_(functionName, lastFailureReason + (e.stack ? ` Stack: ${e.stack}` : ""), "ERROR");
                logToSheet_(functionName, `Charge outcome for ${candidate.label} is UNKNOWN. Stopping cascade for customer ${stripeCustomerId} to avoid a double charge.`, "ERROR");
                break;
            }
        }
        // --- End card cascade: every attempted card failed ---

        paymentSuccess = false;
        paymentResultMessage = `No card succeeded after ${cardsTried} attempt(s) of ${attempts.length} available. Last reason: ${lastFailureReason}`;
        logToSheet_(functionName, `${paymentResultMessage} (Customer: ${stripeCustomerId})`, "WARN");
        return { success: false, message: paymentResultMessage, paymentIntentId: paymentIntentId, cardsTried: cardsTried };

    } catch (outerError) {
        // Error fetching payment methods or other setup issue
        logToSheet_(functionName, `Error in payment attempt setup for ${stripeCustomerId}: ${outerError.message}`+ (outerError.stack ? ` Stack: ${outerError.stack}`: ""), "ERROR");
        return { success: false, message: `Setup Error: ${outerError.message}`, paymentIntentId: null };
    }
}
/**
 * Retrieves the Stripe Secret Key from Google Secret Manager.
 * **Revised with enhanced logging.**
 * @returns {string|null} The Stripe secret key value or null on failure.
 */
function getStripeSecretKey() {
  const functionName = "getStripeSecretKey";
  logToSheet_(functionName, "Attempting to retrieve Stripe Secret Key...", "INFO");
  var projectId = 'skilful-union-451921-m9'; // Ensure this is correct
  var secretName = 'STRIPE_SECRET_KEY';     // Ensure this is correct
  var version = 'latest'; // Or specify a version number if needed
  var url = `https://secretmanager.googleapis.com/v1/projects/${projectId}/secrets/${secretName}/versions/${version}:access`;
  logToSheet_(functionName, `Target Secret Manager URL: ${url}`, "DEBUG");

  let token = null;
  try {
    // --- Step 1: Get the Service Account Token ---
    logToSheet_(functionName, "Requesting service account token...", "INFO");
    token = getServiceAccountToken(); // Call the token function

    if (!token) {
      logToSheet_(functionName, "getServiceAccountToken() returned null. Cannot proceed.", "ERROR");
      return null; // Exit early if token acquisition failed
    }
    logToSheet_(functionName, "Service account token received (length: " + (token ? token.length : 0) + ").", "DEBUG");

    // --- Step 2: Prepare options for Secret Manager API call ---
    var options = {
      method: 'get', // Changed to GET as per API spec for :access
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true // Capture errors manually
    };
    logToSheet_(functionName, "Prepared options for Secret Manager API call.", "DEBUG");
    // Don't log the full options here as it contains the token.

    // --- Step 3: Call the Secret Manager API ---
    logToSheet_(functionName, "Calling Secret Manager API...", "INFO");
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    logToSheet_(functionName, `Secret Manager API Response Code: ${responseCode}`, "DEBUG");

    // --- Step 4: Process the response ---
    if (responseCode !== 200) {
      logToSheet_(functionName, `Failed to retrieve secret. Code: ${responseCode}. Response: ${responseText.substring(0, 1000)}`, "ERROR");
       // Log common potential issues based on code
      if (responseCode === 403) {
          logToSheet_(functionName, "Hint: Response code 403 (Forbidden) usually means the service account (" + getScriptProperty_('AUTHENTICATOR_SERVICE_ACCOUNT_EMAIL') + ") lacks the 'Secret Manager Secret Accessor' role on the secret '" + secretName + "' or the project.", "WARN");
      } else if (responseCode === 401) {
           logToSheet_(functionName, "Hint: Response code 401 (Unauthorized) often indicates an invalid or expired OAuth token.", "WARN");
      } else if (responseCode === 404) {
           logToSheet_(functionName, "Hint: Response code 404 (Not Found) means the secret name ('"+secretName+"') or project ID ('"+projectId+"') might be incorrect in the URL.", "WARN");
      }
      return null; // Indicate failure
    }

    logToSheet_(functionName, "Secret Manager API call successful (Code 200). Processing payload...", "DEBUG");
    var responseData = JSON.parse(responseText);

    if (!responseData.payload || !responseData.payload.data) {
      logToSheet_(functionName, "Secret Manager response successful, but payload or payload.data is missing.", "ERROR");
      logToSheet_(functionName, `Response Text: ${responseText.substring(0, 500)}`, "DEBUG");
      return null; // Indicate failure
    }
    logToSheet_(functionName, "Payload received. Decoding secret data...", "DEBUG");

    // --- Step 5: Decode and return the secret ---
    var secretValue = Utilities.newBlob(Utilities.base64Decode(responseData.payload.data)).getDataAsString();
    logToSheet_(functionName, "Successfully retrieved and decoded Stripe Secret Key (length: " + (secretValue ? secretValue.length : 0) + ").", "SUCCESS");
    // Be cautious about logging the actual key value, length is safer.
    // logToSheet_(functionName, "Retrieved Key starts with: " + (secretValue ? secretValue.substring(0, 8) : "N/A"), "DEBUG");
    return secretValue;

  } catch (e) {
    logToSheet_(functionName, "CRITICAL Error during Stripe secret key retrieval: " + e.toString() + (e.stack ? " Stack: " + e.stack : ""), "ERROR");
    // Log token status if error happened after getting it
    logToSheet_(functionName, "Token status at time of error: " + (token ? "Obtained" : "Not Obtained or Failed"), "DEBUG");
    return null; // Indicate failure
  }
}
/**
 * Creates a signed JWT and exchanges it for a Google OAuth2 access token
 * using a service account's credentials stored in Script Properties.
 * **Revised with enhanced logging.**
 * @returns {string|null} An access token or null on failure.
 */
function getServiceAccountToken() {
  var functionName = "getServiceAccountToken";
  logToSheet_(functionName, "Attempting to get service account token...", "INFO");

  var clientEmail = null;
  var privateKeyRaw = null;
  var privateKey = null;

  try {
    // --- Step 1: Get Authenticator Credentials ---
    logToSheet_(functionName, "Reading credentials from Script Properties...", "DEBUG");
    var scriptProps = PropertiesService.getScriptProperties();
    clientEmail = scriptProps.getProperty('AUTHENTICATOR_SERVICE_ACCOUNT_EMAIL');
    privateKeyRaw = scriptProps.getProperty('AUTHENTICATOR_SERVICE_ACCOUNT_PRIVATE_KEY');

    // --- **Enhanced Logging: Check Credentials** ---
    if (!clientEmail || typeof clientEmail !== 'string' || clientEmail.trim() === '') {
      logToSheet_(functionName, "Authenticator Service Account Email invalid or missing in Script Properties (Key: AUTHENTICATOR_SERVICE_ACCOUNT_EMAIL). Value: '" + clientEmail + "'", "ERROR");
      throw new Error("Authenticator client email not configured correctly.");
    } else {
      logToSheet_(functionName, "Retrieved clientEmail: " + clientEmail.substring(0, 15) + "...", "DEBUG");
    }

    if (!privateKeyRaw || typeof privateKeyRaw !== 'string' || privateKeyRaw.trim() === '') {
      logToSheet_(functionName, "Authenticator Service Account Private Key invalid or missing in Script Properties (Key: AUTHENTICATOR_SERVICE_ACCOUNT_PRIVATE_KEY). Key is empty or not a string.", "ERROR");
      throw new Error("Authenticator private key not configured correctly.");
    } else {
       logToSheet_(functionName, "Retrieved raw private key (length: " + privateKeyRaw.length + "). Checking format...", "DEBUG");
       if (!privateKeyRaw.includes("-----BEGIN PRIVATE KEY-----") || !privateKeyRaw.includes("-----END PRIVATE KEY-----")) {
         logToSheet_(functionName, "WARNING: Raw private key seems incomplete. It's missing '-----BEGIN PRIVATE KEY-----' or '-----END PRIVATE KEY-----'. Ensure the entire key was copied.", "WARN");
       }
    }

    // --- Explicitly fix newline characters ---
    privateKey = privateKeyRaw.replace(/\\n/g, '\n');
    logToSheet_(functionName, "Private key after newline correction (length: " + privateKey.length + ").", "DEBUG");
    if (privateKey.length < 500) { // Arbitrary short length check
        logToSheet_(functionName, "WARNING: Corrected private key seems unusually short. Double-check the value in Script Properties.", "WARN");
    }
     // --- **End Enhanced Logging** ---


    // --- Step 2: Define JWT Claims ---
    var scope = 'https://www.googleapis.com/auth/cloud-platform'; // Correct scope for Secret Manager & other Cloud APIs
    var tokenUrl = 'https://oauth2.googleapis.com/token';
    var header = { alg: 'RS256', typ: 'JWT' };
    var nowSeconds = Math.floor(Date.now() / 1000);
    var expirationSeconds = nowSeconds + 3600; // Token valid for 1 hour
    var claimSet = {
      iss: clientEmail, scope: scope, aud: tokenUrl,
      exp: expirationSeconds, iat: nowSeconds
    };
    logToSheet_(functionName, `JWT Claim Set prepared: ${JSON.stringify(claimSet)}`, "DEBUG");

    // --- Step 3: Encode JWT Header and Claim Set ---
    var encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header));
    var encodedClaimSet = Utilities.base64EncodeWebSafe(JSON.stringify(claimSet));
    logToSheet_(functionName, "JWT Header and Claim Set encoded.", "DEBUG");

    // --- Define unsignedJwt ---
    var unsignedJwt = encodedHeader + '.' + encodedClaimSet;
    logToSheet_(functionName, "Unsigned JWT created (length: " + unsignedJwt.length + ").", "DEBUG");

    // --- Step 4: Create Signature ---
    logToSheet_(functionName, "Attempting to create RSA SHA256 signature...", "DEBUG");
    var signatureBytes = Utilities.computeRsaSha256Signature(unsignedJwt, privateKey);
    var encodedSignature = Utilities.base64EncodeWebSafe(signatureBytes);
    logToSheet_(functionName, "JWT Signature created successfully.", "DEBUG");

    // --- Step 5: Assemble Signed JWT ---
    var signedJwt = unsignedJwt + '.' + encodedSignature;
    logToSheet_(functionName, "Signed JWT assembled (length: " + signedJwt.length + ").", "DEBUG");

    // --- Step 6: Exchange JWT for Access Token ---
    var payload = { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: signedJwt };
    var options = { method: 'post', contentType: 'application/x-www-form-urlencoded', payload: payload, muteHttpExceptions: true };
    logToSheet_(functionName, `Requesting access token from: ${tokenUrl} with grant_type: jwt-bearer`, "INFO");
    var response = UrlFetchApp.fetch(tokenUrl, options); // Make the call
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    logToSheet_(functionName, `Token exchange response Code: ${responseCode}`, "DEBUG");


    // --- Step 7: Process Response ---
    if (responseCode === 200) {
      var responseData = JSON.parse(responseText);
      var accessToken = responseData.access_token;
      if (accessToken) {
        logToSheet_(functionName, "Successfully obtained service account access token (length: " + accessToken.length + ").", "SUCCESS");
        return accessToken; // SUCCESS!
      } else {
        logToSheet_(functionName, "Token exchange successful (Code 200), BUT 'access_token' field missing in response!", "ERROR");
        logToSheet_(functionName, `OAuth Response Text (Code 200): ${responseText.substring(0, 1000)}`, "DEBUG"); // Log response even on success if token missing
        throw new Error("Access token missing in successful OAuth response.");
      }
    } else {
      // Log error details extensively
      logToSheet_(functionName, `Token exchange FAILED. Code: ${responseCode}`, "ERROR");
      logToSheet_(functionName, `OAuth2 Error Response Text (Code ${responseCode}): ${responseText.substring(0, 1500)}`, "ERROR"); // Log more of the error
      // Provide hints based on common errors
      if (responseText.includes("invalid_grant")) {
         logToSheet_(functionName, "Hint 'invalid_grant': Often means clock skew between Apps Script server & Google auth, invalid JWT format/signature, incorrect private key, or service account disabled/deleted.", "WARN");
      } else if (responseText.includes("invalid_request") && responseText.includes("private key")) {
         logToSheet_(functionName, "Hint 'invalid_request' mentioning key: Suggests the private key format might be corrupted or incorrect in Script Properties.", "WARN");
      } else if (responseText.includes("Invalid assertion format")) {
         logToSheet_(functionName, "Hint 'Invalid assertion format': Suggests the signed JWT structure was wrong.", "WARN");
      } else if (responseCode === 403) {
          logToSheet_(functionName, "Hint: Response code 403 (Forbidden) during token exchange is less common but could indicate project-level restrictions.", "WARN");
      } else if (responseCode === 400 && responseText.includes("deleted service account")) {
         logToSheet_(functionName, "Hint 'deleted service account': The service account used ("+clientEmail+") has been deleted in GCP.", "WARN");
      }
      throw new Error("OAuth2 token exchange failed. Code: " + responseCode + ". Check logs for response text.");
    }

  } catch (e) {
    logToSheet_(functionName, "CRITICAL Error during token acquisition: " + e.toString() + (e.stack ? " Stack: " + e.stack : ""), "ERROR");
    // Log state of key variables if they were assigned
    logToSheet_(functionName, `State at failure: clientEmail obtained: ${!!clientEmail}, privateKey processed: ${!!privateKey}`, "DEBUG");
    return null; // Return null to signify failure
  }
}