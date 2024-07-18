function moveCompletedProjects() {
  // Open the spreadsheet and get the sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackerSheet = ss.getSheetByName("Tracker");
  var completedSheet = ss.getSheetByName("Completed Projects");

  // Get the last row in column C
  var lastRow = trackerSheet.getLastRow();

  // Get the data from the Tracker sheet
  var range = trackerSheet.getRange("B3:B" + lastRow);
  var values = range.getValues();

  // Loop through the values and find rows with TRUE
  var rowsToMove = [];
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === true) {
      var rowNumber = i + 3;
      var rowData = trackerSheet.getRange(rowNumber, 3, 1, 11).getValues()[0];

      rowsToMove.push(rowData);
      trackerSheet.deleteRow(rowNumber);
    }
  }

  // Append rows to the Completed Projects sheet
  if (rowsToMove.length > 0) {
    completedSheet
      .getRange(
        completedSheet.getLastRow() + 1,
        1,
        rowsToMove.length,
        rowsToMove[0].length
      )
      .setValues(rowsToMove);
  }
}
