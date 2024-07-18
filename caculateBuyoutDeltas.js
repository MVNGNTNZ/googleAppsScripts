function calculateBuyoutDeltas() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var scoreboardJobs = spreadsheet.getSheetByName("Scoreboard - Jobs");
  var buyoutBaseline = spreadsheet.getSheetByName("Buyout Baseline");

  // Get the jobs from Scoreboard
  var lastRowSB = scoreboardJobs.getLastRow();
  var jobsSB = scoreboardJobs
    .getRange(2, 2, lastRowSB - 1, 1)
    .getValues()
    .flat()
    .map((job) => job.toLowerCase());

  // Determine the current report and previous report
  var lastFilledColumn = buyoutBaseline.getLastColumn();
  var currentReport = lastFilledColumn;
  var previousReport = lastFilledColumn - 1;

  // Get the jobs and sums from Buyout Baseline
  var lastRowBB = buyoutBaseline.getLastRow();
  var jobsBB = buyoutBaseline
    .getRange(2, 2, lastRowBB - 1, 1)
    .getValues()
    .flat()
    .map((job) => job.toLowerCase());

  // Get the sums from current and previous report in Buyout Baseline
  var currentSums = buyoutBaseline
    .getRange(2, currentReport, lastRowBB - 1, 1)
    .getValues()
    .flat();
  var previousSums = buyoutBaseline
    .getRange(2, previousReport, lastRowBB - 1, 1)
    .getValues()
    .flat();

  // Create a map of jobs to sums for current and previous report
  var currentSumsMap = {};
  var previousSumsMap = {};

  for (var i = 0; i < jobsBB.length; i++) {
    var job = jobsBB[i];
    if (job) {
      currentSumsMap[job] = currentSums[i] || 0;
      previousSumsMap[job] = previousSums[i] || 0;
    }
  }

  // Create a map for deltas
  var deltasMap = {};

  // Calculate the deltas
  for (var job in currentSumsMap) {
    var currentSum = currentSumsMap[job] || 0;
    var previousSum = previousSumsMap[job] || 0;
    deltasMap[job] = currentSum - previousSum;
  }

  // Prepare the deltas for output
  var deltas = jobsSB.map(function (job) {
    if (job) {
      return [deltasMap[job] || 0];
    } else {
      return [""];
    }
  });

  // Set column borders, output the deltas to column S in Scoreboard, and set formatting
  scoreboardJobs
    .getRange(1, 19, scoreboardJobs.getMaxRows(), 1)
    .setBorder(false, true, false, true, false, false);

  var outputRange = scoreboardJobs.getRange(2, 19, deltas.length, 1);
  outputRange.setValues(deltas);
  outputRange.setNumberFormat(
    '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
  );
  outputRange.setBorder(false, true, false, true, false, false);

  var headerCell = scoreboardJobs.getRange(1, 19);
  headerCell.setBorder(false, true, true, true, false, false);

  scoreboardJobs.autoResizeColumn(19);
  var currentWidth = scoreboardJobs.getColumnWidth(19);
  scoreboardJobs.setColumnWidth(19, currentWidth + 5);

  // Find rows with "Totals" in column A and sum the values above
  var data = scoreboardJobs
    .getRange(2, 1, lastRowSB - 1, 1)
    .getValues()
    .flat();
  var startRow = 2;
  for (var i = 0; i < data.length; i++) {
    if (data[i] === "Totals") {
      var endRow = i + 1; // Row index in the sheet (1-based)
      var sumRange = scoreboardJobs.getRange(
        startRow,
        19,
        endRow - startRow,
        1
      );
      var sum = sumRange.getValues().reduce(function (acc, val) {
        return acc + (val[0] || 0);
      }, 0);

      var totalCell = scoreboardJobs.getRange(endRow + 1, 19);
      totalCell.setValue(sum);
      totalCell.setNumberFormat(
        '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
      );
      totalCell.setBorder(true, true, true, true, true, true);
      startRow = endRow + 1;
    }
  }
}
