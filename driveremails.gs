/**
 * Sends email notifications for tickets marked 'Ready' in Column Z.
 * **REVISED: Explicitly formats date/time, removed flush inside loop.**
 */
function emailDriversTicketNotifications() {
  const functionName = "emailDriversTicketNotifications";
  const ui = SpreadsheetApp.getUi();
  logToSheet_(functionName, "Starting email notification run.", "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    ui.alert("Error: Target sheet '" + TARGET_SHEET_NAME + "' not found.");
    logToSheet_(functionName, "Sheet not found.", "ERROR");
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headerRows = 1;
  const emailStatusColIndex = COL.EMAIL_STATUS - 1;
  const driverEmailColIndex = COL.DRIVER_EMAIL - 1;
  const driverNameColIndex = COL.RESPONSIBLE_DRIVER - 1;
  const bookingIdColIndex = COL.BOOKING_ID - 1;
  const citationIdColIndex = COL.VIOLATION_ID - 1;
  const originalPenaltyColIndex = COL.ORIGINAL_PENALTY - 1;
  const violationTypeColIndex = COL.VIOLATION_TYPE - 1;
  const locationColIndex = COL.VIOLATION_LOCATION - 1;
  const timeColIndex = COL.TIME_VIOLATION - 1;
  const dateColIndex = COL.DATE_VIOLATION - 1;
  const pdfLinkColIndex = COL.PDF_LINK - 1;
  const paymentStatusColIndex = COL.DRIVER_BILLING_STATUS - 1;
  const totalOwedColIndex = COL.TOTAL_OWED_COLLECTIONS - 1;
  const vehicleNumColIndex = COL.VEHICLE - 1;

  // --- Email Configuration ---
  const enLogoUrl = "https://lh3.googleusercontent.com/d/1HZlt2B4qSveYMHQZ62TlpGwn30VFXMjZ=s220?authuser=3";
  let enLogoBlob = null;
  try {
      enLogoBlob = UrlFetchApp.fetch(enLogoUrl).getBlob().setName("enLogoBlob");
  } catch(e) {
      logToSheet_(functionName, `Failed to fetch logo blob: ${e}`, "WARN");
  }
  const pic = enLogoBlob ? "<a href='http://www.envoythere.com'><img src='cid:enLogo' style='width:100px; height:75px;' alt='Envoy Logo'/></a>" : "";
  const bcc = "sky@envoythere.com";
  const adminFee = 25.00;
  // --- End Email Configuration ---

  let processedCount = 0;
  let sentCount = 0;
  let errorCount = 0;
  let sheetUpdates = []; // Batch status updates

  // --- Iterate through rows (skip header) ---
  for (let i = headerRows; i < data.length; i++) {
    const row = data[i];
    const actualRowNum = i + 1;

    const emailStatus = String(row[emailStatusColIndex]).trim();
    if (emailStatus.toUpperCase() !== "READY") {
      continue;
    }

    // --- Gather Data for Email ---
    const emailAddress = String(row[driverEmailColIndex]).trim();
    const fullName = String(row[driverNameColIndex]).trim();
    const bookingId = String(row[bookingIdColIndex]).trim();
    const citationId = String(row[citationIdColIndex]).trim();
    const originalPenalty = parseFloat(row[originalPenaltyColIndex]) || 0;
    const violationDescription = String(row[violationTypeColIndex]).trim();
    const location = String(row[locationColIndex]).trim();
    const pdfUrl = String(row[pdfLinkColIndex]).trim();
    const paymentStatus = String(row[paymentStatusColIndex]).trim().toUpperCase(); // Normalize for check
    const totalOwedCollections = parseFloat(row[totalOwedColIndex]) || (originalPenalty + adminFee);
    const envoyNumber = String(row[vehicleNumColIndex]).trim();

    // --- Get RAW Date/Time Values from Sheet ---
    const rawDateValue = row[dateColIndex];
    const rawTimeValue = row[timeColIndex];
    // logToSheet_(functionName, `Row ${actualRowNum}: Raw Date Value='${rawDateValue}' (Type: ${typeof rawDateValue}), Raw Time Value='${rawTimeValue}' (Type: ${typeof rawTimeValue})`, "DEBUG"); // Can be noisy

    // --- Explicitly Format Date and Time for Email ---
    // Use the *formatted* values from the sheet if they are already dates/times
    let emailDateFormatted = rawDateValue instanceof Date ? Utilities.formatDate(rawDateValue, Session.getScriptTimeZone(), "MMM dd yyyy") : String(rawDateValue);
    let emailTimeFormatted = rawTimeValue instanceof Date ? Utilities.formatDate(rawTimeValue, Session.getScriptTimeZone(), "h:mm a").toUpperCase() : String(rawTimeValue);
    // Add fallback parsing if the sheet contains strings instead of Date objects
    if (!(rawDateValue instanceof Date)) {
        let parsedD = parseFlexibleDate_(rawDateValue); // Use helper
        if (parsedD) emailDateFormatted = Utilities.formatDate(parsedD, Session.getScriptTimeZone(), "MMM dd yyyy");
    }
     if (!(rawTimeValue instanceof Date)) {
        let parsedT = parseFlexibleTime_(rawTimeValue); // Use helper
        if (parsedT) emailTimeFormatted = Utilities.formatDate(parsedT, Session.getScriptTimeZone(), "h:mm a").toUpperCase();
    }

    // Basic validation
    if (!emailAddress || !fullName || !bookingId || !citationId) {
      logToSheet_(functionName, `Row ${actualRowNum}: Skipping email - Missing required data (Email, Name, Booking ID, Citation ID).`, "WARN");
      sheetUpdates.push({ row: actualRowNum, col: COL.EMAIL_STATUS, value: "Error - Missing Data" });
      errorCount++;
      continue;
    }
    // Use normalized paymentStatus for check
    if (paymentStatus !== "PAID" && paymentStatus !== "COLLECTIONS" && paymentStatus !== "PAID (AMOUNT $0)") {
       logToSheet_(functionName, `Row ${actualRowNum}: Skipping email - Invalid payment status '${row[paymentStatusColIndex]}'. Expected 'Paid', 'Paid (Amount $0)' or 'Collections'.`, "WARN");
       sheetUpdates.push({ row: actualRowNum, col: COL.EMAIL_STATUS, value: `Error - Invalid Status '${row[paymentStatusColIndex]}'` });
       errorCount++;
       continue;
    }

    // Mark as Processing Visually (Optional - can remove if pure batching preferred)
    logToSheet_(functionName, `Processing email for row ${actualRowNum}, Email: ${emailAddress}. Date: ${emailDateFormatted}, Time: ${emailTimeFormatted}`, "INFO");
    sheet.getRange(actualRowNum, COL.EMAIL_STATUS).setValue("Processing");
    SpreadsheetApp.flush(); // Flush *only* for visual update

    processedCount++;
    const firstName = fullName.split(' ')[0] || 'Valued Customer';
    const totalAmountBilled = originalPenalty + adminFee;

    // Fetch PDF Attachment
    let pdfBlob = null;
    if (pdfUrl) {
        pdfBlob = getBlobFromDriveUrl_(pdfUrl); // Uses helper
        if (!pdfBlob) {
           logToSheet_(functionName, `Row ${actualRowNum}: Could not retrieve PDF blob from URL: ${pdfUrl}`, "WARN");
        }
    }

    // --- Build Email Content ---
    const subject = `Citation Notification for Booking ${bookingId}`;
    let emailBodyHtml = "";
    // Define message templates (as before)
        const messagePaid =
      `<font face="Arial, sans-serif" size="2">Hello ${firstName},<br><br>` +
      `This email is to notify you that a balance of <b>$${totalAmountBilled.toFixed(2)}</b> ` +
      `has been collected from your payment method on file ` +
      `in relation to an infraction that occurred during your booking with Envoy #${envoyNumber}, ` +
      `Booking ID ${bookingId}.<br><br>` +
      `These fees are comprised of the following:<br>` +
      `Violation Fee (${citationId}): $${originalPenalty.toFixed(2)}<br>` +
      `Administrative Processing Fee: $${adminFee.toFixed(2)}<br>` +
      `Total balance collected: $${totalAmountBilled.toFixed(2)}<br><br>` +
      `This amount is in relation to a ${violationDescription} violation at ${location} ` +
      `at ${emailTimeFormatted} on ${emailDateFormatted}.<br><br>` + // USE FORMATTED VALUES
      (pdfBlob ? "A digitized version of the ticket is attached to this email.<br><br>" : "Please contact support if you require a copy of the ticket.<br><br>") +
      `Thank you,<br>--<br></font>`;

    const messageCollections =
       `<font face="Arial, sans-serif" size="2">Hello ${firstName},<br><br>` +
       `We are reaching out regarding an infraction that occurred during your booking with Envoy #${envoyNumber}, ` +
       `Booking ID ${bookingId}. ` + (pdfBlob ? "A digitized invoice is attached for your reference.<br><br>" : "Please contact support if you require a copy of the ticket.<br><br>") +
       `Please see info regarding the infraction below:<br><br>` +
       `Citation ID: ${citationId}<br>` +
       `Violation Description: ${violationDescription}<br>` +
       `Date & Time: ${emailDateFormatted} at ${emailTimeFormatted}<br>` + // USE FORMATTED VALUES
       `Location: ${location}<br><br>` +
       `The violation fee plus an administrative fee of $${adminFee.toFixed(2)} was attempted to be charged to your payment method on file, and was declined. This has resulted in a current balance due of <b>$${totalOwedCollections.toFixed(2)}</b>.<br><br>` +
       `Please settle this overdue balance in the Wallet section of the Envoy app at your earliest convenience to reinstate your account access.<br><br>` +
       `If you have any questions, please reply to this email or call our support line at the phone number listed below.<br><br>` +
       `Thank you,<br>--<br></font>`;

        const messageSignature =
      `<br><br><hr style="border: none; border-top: 1px solid #eee;"/><br>`+
      `<b><font color='#3900FF' size='1'>Envoy</font></b><br>` +
      `<b><font color='#3900FF' size='1'>C:</font></b><font size='1'> 888-610-0506 or text us at 424-404-6512</font><br>` +
      `<b><font color='#3900FF' size='1'>E:</font></b><font size='1'> <a href='mailto:info@envoythere.com'>info@envoythere.com</a></font><br>` +
      `<b><font color='#3900FF' size='1'>W:</font></b><font size='1'> <a href='http://www.envoythere.com'>www.envoythere.com</a></font><br>` +
      `<b><font color='#3900FF' size='1'>A:</font></b><font size='1'> 8575 Washington Blvd., Culver City, CA, 90232</font><br>`;

    // Choose message based on normalized status
    if (paymentStatus === "PAID" || paymentStatus === "PAID (AMOUNT $0)") {
      emailBodyHtml = messagePaid + pic + messageSignature;
    } else if (paymentStatus === "COLLECTIONS") {
      emailBodyHtml = messageCollections + pic + messageSignature;
    } else {
      // Should not happen due to validation above, but safety check
       logToSheet_(functionName, `Row ${actualRowNum}: Internal error - Unexpected payment status '${paymentStatus}' for email body selection.`, "ERROR");
       sheetUpdates.push({ row: actualRowNum, col: COL.EMAIL_STATUS, value: "Error - Internal Status" });
       errorCount++;
       continue;
    }

    // --- Send Email ---
    try {
      const mailOptions = {
        htmlBody: emailBodyHtml,
        bcc: bcc,
        name: "Envoy Support"
      };
      if (enLogoBlob) {
        mailOptions.inlineImages = { enLogo: enLogoBlob };
      }
      if (pdfBlob) {
        // Try to use a more descriptive name if available from PDF metadata or original upload
        let attachmentName = pdfBlob.getName() || citationId || bookingId || "Ticket.pdf";
        // Ensure it ends with .pdf
         if (!attachmentName.toLowerCase().endsWith('.pdf')) {
             attachmentName += '.pdf';
         }
        mailOptions.attachments = [pdfBlob.setName(attachmentName)];
         logToSheet_(functionName, `Row ${actualRowNum}: Attaching PDF with name: ${attachmentName}`, "DEBUG");
      }

      MailApp.sendEmail(emailAddress, subject, "", mailOptions);
      logToSheet_(functionName, `Row ${actualRowNum}: Successfully sent email to ${emailAddress}`, "SUCCESS");
      sheetUpdates.push({ row: actualRowNum, col: COL.EMAIL_STATUS, value: "Sent" });
      sentCount++;

    } catch (mailError) {
      logToSheet_(functionName, `Row ${actualRowNum}: Error sending email to ${emailAddress}: ${mailError}`, "ERROR");
      sheetUpdates.push({ row: actualRowNum, col: COL.EMAIL_STATUS, value: `Error - ${mailError.message.substring(0,100)}` });
      errorCount++;
    }
     // --- REMOVED SpreadsheetApp.flush(); ---

  } // --- End row loop ---

  // --- Apply Batch Updates ---
   if (sheetUpdates.length > 0) {
      logToSheet_(functionName, `Applying ${sheetUpdates.length} batch status updates to the sheet...`, "INFO");
      sheetUpdates.forEach(update => {
          sheet.getRange(update.row, update.col).setValue(update.value);
      });
      SpreadsheetApp.flush(); // Flush once after all updates are applied
      logToSheet_(functionName, "Batch status updates applied.", "DEBUG");
  } else {
       logToSheet_(functionName, "No email status updates needed.", "INFO");
  }


   logToSheet_(functionName, `Email notification run complete. Processed: ${processedCount}, Sent: ${sentCount}, Errors: ${errorCount}`, "INFO");
   ui.alert(`Email Notification Run Complete.\n\nRows Processed: ${processedCount}\nEmails Sent: ${sentCount}\nErrors: ${errorCount}`);
}
/**
 * Helper function to get a Blob from a Google Drive URL.
 * @param {string} url The Google Drive file URL.
 * @returns {GoogleAppsScript.Base.Blob | null} The file blob or null if error/not found.
 */
function getBlobFromDriveUrl_(url) {
    const functionName = "getBlobFromDriveUrl_";
    if (!url || typeof url !== 'string') return null;
    try {
        // More robust regex to handle various Drive URL formats
        const fileIdMatch = url.match(/[-\w]{25,}/); // Finds sequences of 25+ letters, numbers, -, _
        if (fileIdMatch && fileIdMatch[0]) {
            const fileId = fileIdMatch[0];
            logToSheet_(functionName, `Extracted File ID: ${fileId} from URL: ${url}`, "DEBUG");
            const file = DriveApp.getFileById(fileId);
            return file.getBlob();
        } else {
            logToSheet_(functionName, `Could not extract File ID from URL: ${url}`, "WARN");
            return null;
        }
    } catch (e) {
        // Catch specific errors like "File not found"
        if (e.message.includes("ocument is missing")) { // DriveApp often gives "Document is missing" for not found
             logToSheet_(functionName, `File not found in Drive for URL ${url}. Error: ${e.message}`, "WARN");
        } else {
             logToSheet_(functionName, `Error fetching blob for URL ${url}: ${e}`, "ERROR");
        }
        return null;
    }
}
