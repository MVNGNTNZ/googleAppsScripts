function calculateBuyouts() {
  // Open spreadsheet and the specific sheets
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var purchaseOrders = spreadsheet.getSheetByName("PO's");
  var buyoutBaseline = spreadsheet.getSheetByName("Buyout Baseline");

  // Get today's date
  var today = new Date();

  // Get the range of row 1 of Buyout Baseline
  var row = buyoutBaseline
    .getRange(1, 1, 1, buyoutBaseline.getMaxColumns())
    .getValues()[0];

  // Find next empty column in row 1 of Buyout Baseline
  var nextEmptyColumn = row.indexOf("") + 1;

  // If no empty column found, set empty column to first after last filled column
  if (nextEmptyColumn === 0) {
    nextEmptyColumn = row.length + 1;
  }

  // Find the last empty row in column B of Buyout Baseline
  var lastRow = buyoutBaseline.getLastRow();
  var jobs = buyoutBaseline
    .getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat();

  // Create a map of jobs to their respective rows in Buyout Baseline
  var jobRowMap = {};
  for (var i = 0; i < jobs.length; i++) {
    if (jobs[i]) {
      var jobName = jobs[i].toString().toLowerCase();
      if (!jobRowMap[jobName]) {
        jobRowMap[jobName] = [];
      }
      jobRowMap[jobName].push(i + 2); // Row index in the sheet (1-based)
    }
  }

  // Get all POs data and filter POs to include only those that match a job in Buyout Baseline
  var poData = purchaseOrders
    .getRange(2, 1, purchaseOrders.getLastRow() - 1, 7)
    .getValues();
  var filteredPOs = poData.filter(function (row) {
    return (
      jobRowMap.hasOwnProperty(row[0].toString().toLowerCase()) &&
      row[6] !== "Unreleased"
    );
  });

  // Create a map to store the total sums for each job
  var jobTotals = {};
  for (var job in jobRowMap) {
    jobTotals[job] = 0;
  }

  // Calculate the totals for each job
  filteredPOs.forEach(function (row) {
    var job = row[0].toString().toLowerCase();
    var amount = row[5];
    if (jobTotals[job] !== undefined) {
      jobTotals[job] += amount;
    }
  });

  // Output the totals into the corresponding rows and set formatting
  for (var job in jobTotals) {
    var totalSum = jobTotals[job];
    var rowIndices = jobRowMap[job];
    rowIndices.forEach(function (rowIndex) {
      var cell = buyoutBaseline.getRange(rowIndex, nextEmptyColumn);
      cell.setValue(totalSum);
      cell.setNumberFormat(
        '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
      );
    });
  }

  // Set column borders
  buyoutBaseline
    .getRange(1, nextEmptyColumn, buyoutBaseline.getMaxRows(), 1)
    .setBorder(false, true, false, true, false, false);

  // Find rows with "Totals" in column A and sum the values above
  var data = buyoutBaseline
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();
  var startRow = 2;
  for (var i = 0; i < data.length; i++) {
    if (data[i] === "Totals") {
      var endRow = i + 1; // Row index in the sheet (1-based)
      var sumRange = buyoutBaseline.getRange(
        startRow,
        nextEmptyColumn,
        endRow - startRow,
        1
      );
      var sum = sumRange.getValues().reduce(function (acc, val) {
        return acc + (val[0] || 0);
      }, 0);

      var totalCell = buyoutBaseline.getRange(endRow + 1, nextEmptyColumn);
      totalCell.setValue(sum);
      totalCell.setNumberFormat(
        '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
      );
      totalCell.setBorder(true, true, true, true, true, true);
      startRow = endRow + 1;
    }
  }

  // Set today's date and formatting, autofit column width
  var dateCell = buyoutBaseline.getRange(1, nextEmptyColumn);
  dateCell.setValue(today);
  dateCell.setNumberFormat("MM/DD/YY");
  dateCell.setHorizontalAlignment("center");
  dateCell.setVerticalAlignment("middle");
  dateCell.setFontWeight("bold");
  dateCell.setBorder(true, true, true, true, true, true);

  buyoutBaseline.autoResizeColumn(nextEmptyColumn);
  var currentWidth = buyoutBaseline.getColumnWidth(nextEmptyColumn);
  buyoutBaseline.setColumnWidth(nextEmptyColumn, currentWidth + 5);
}
