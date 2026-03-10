
// --- END OF BATCH-AWARE LOGGING SYSTEM ---
function showBatchTicketForm() {
  const html = HtmlService.createHtmlOutputFromFile('BatchTicketForm')
      .setWidth(1200)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Batch PDF Upload & Review');
}
/**
 * Batch‑oriented wrapper around the existing single‑file pipeline.
 * Accepts an Array of the same `fileData` objects currently passed to
 * `uploadAndAnalyzeTicket()` and returns a parallel Array of result objects.
 *
 * NOTE: ▸ No changes were made to OCR parsing, spreadsheet writes, or
 *        downstream admin/collection workflows – we simply invoke the
 *        well‑tested single‑item path in a loop and stream the responses.
 *      ▸ A 500 ms pause guards Gemini API quotas; adjust if you have higher
 *        quota.
 */
// PLACE NEAR OTHER FILE PROCESSING FUNCTIONS
