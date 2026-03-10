function downloadTollFilesOnly() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const startRow = 2; // assuming row 1 is headers
  const lastRow = sheet.getLastRow();

  // Column indexes (1-based)
  const LINK_COL = 22; // V
  const FILTER_COL = 24; // X

  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, FILTER_COL).getValues();

  const folder = DriveApp.createFolder("Toll Files " + new Date().toISOString());

  data.forEach((row, i) => {
    const link = row[LINK_COL - 1];
    const filterValue = row[FILTER_COL - 1];

    if (filterValue !== "Toll" || !link) return;

    const match = link.match(/[-\w]{25,}/);
    if (!match) return;

    try {
      const file = DriveApp.getFileById(match[0]);
      file.makeCopy(file.getName(), folder);
    } catch (e) {
      Logger.log(`Row ${startRow + i}: failed to copy`);
    }
  });

  Logger.log("Done. Folder URL: " + folder.getUrl());
}
