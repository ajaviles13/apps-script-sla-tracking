function showAllHiddenSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  sheets.forEach(sheet => {
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
  });
}