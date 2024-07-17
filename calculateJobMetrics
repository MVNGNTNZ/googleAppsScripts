function calculateJobMetrics() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var scoreboardJobs = spreadsheet.getSheetByName("Scoreboard - Jobs");
  var toDo = spreadsheet.getSheetByName("To Do");
  var buyoutBaseline = spreadsheet.getSheetByName("Buyout Baseline");

  // Get the most recent report date from Buyout Baseline
  var lastFilledColumnBB = buyoutBaseline.getLastColumn();
  var reportDate = buyoutBaseline.getRange(1, lastFilledColumnBB).getValue();

  // Get the jobs from Scoreboard
  var lastRowSB = scoreboardJobs.getLastRow();
  var jobsInScoreboard = scoreboardJobs.getRange(2, 2, lastRowSB - 1, 1).getValues().flat().map(job => job.toLowerCase());

  // Get the jobs, statuses, and scheduled dates from To Do
  var lastRowTD = toDo.getLastRow();
  var jobsInToDo = toDo.getRange(2, 1, lastRowTD - 1, 1).getValues().flat().map(job => job.toLowerCase());
  var statuses = toDo.getRange(2, 3, lastRowTD - 1, 1).getValues().flat();
  var scheduledDates = toDo.getRange(2, 8, lastRowTD - 1, 1).getValues().flat();

  // Initialize arrays to store the results
  var countsNegativeDays = [];
  var percentages0To30 = [];
  var percentages30To60 = [];
  var percentages60To90 = [];
  var percentagesOver90 = [];

  // Process each job in Scoreboard
  jobsInScoreboard.forEach(function(job) {
    if (job) { // Only process if job is not empty
      var countNegativeDays = 0;
      var count0To30 = 0;
      var count30To60 = 0;
      var count60To90 = 0;
      var countOver90 = 0;
      var totalComplete = 0;
      var totalTasks = 0;

      // Find all matching jobs in To Do
      jobsInToDo.forEach(function(toDoJob, index) {
        if (toDoJob === job) {
          var status = statuses[index];
          var scheduledDate = new Date(scheduledDates[index]);
          var daysDelta = (scheduledDate - reportDate) / (1000 * 60 * 60 * 24); // Convert milliseconds to days

          totalTasks++;
          if (status === "Complete") {
            totalComplete++;
            if (daysDelta < 0) {
              countNegativeDays++;
            } else if (daysDelta <= 30) {
              count0To30++;
            } else if (daysDelta <= 60) {
              count30To60++;
            } else if (daysDelta <= 90) {
              count60To90++;
            } else {
              countOver90++;
            }
          }
        }
      });

      // Calculate percentages based on total tasks
      var percentage0To30 = totalTasks > 0 ? (count0To30 / totalTasks) : 0;
      var percentage30To60 = totalTasks > 0 ? (count30To60 / totalTasks) : 0;
      var percentage60To90 = totalTasks > 0 ? (count60To90 / totalTasks) : 0;
      var percentageOver90 = totalTasks > 0 ? (countOver90 / totalTasks) : 0;

      // Store the results
      countsNegativeDays.push([countNegativeDays]);
      percentages0To30.push([percentage0To30]);
      percentages30To60.push([percentage30To60]);
      percentages60To90.push([percentage60To90]);
      percentagesOver90.push([percentageOver90]);
    } else {
      // Push empty values for rows with no job
      countsNegativeDays.push([""]);
      percentages0To30.push([""]);
      percentages30To60.push([""]);
      percentages60To90.push([""]);
      percentagesOver90.push([""]);
    }
  });

  // Output the results to Scoreboard
  scoreboardJobs.getRange(2, 24, countsNegativeDays.length, 1).setValues(countsNegativeDays); // Column X
  scoreboardJobs.getRange(2, 23, percentages0To30.length, 1).setValues(percentages0To30); // Column W
  scoreboardJobs.getRange(2, 22, percentages30To60.length, 1).setValues(percentages30To60); // Column V
  scoreboardJobs.getRange(2, 21, percentages60To90.length, 1).setValues(percentages60To90); // Column U
  scoreboardJobs.getRange(2, 20, percentagesOver90.length, 1).setValues(percentagesOver90); // Column T

  // Set number format to percentage / count for the relevant columns
  scoreboardJobs.getRange(2, 24, countsNegativeDays.length, 1).setNumberFormat("0"); // Column X
  scoreboardJobs.getRange(2, 23, percentages0To30.length, 1).setNumberFormat("0%"); // Column W
  scoreboardJobs.getRange(2, 22, percentages30To60.length, 1).setNumberFormat("0%"); // Column V
  scoreboardJobs.getRange(2, 21, percentages60To90.length, 1).setNumberFormat("0%"); // Column U
  scoreboardJobs.getRange(2, 20, percentagesOver90.length, 1).setNumberFormat("0%"); // Column T

  // Center-align the outputs
  scoreboardJobs.getRange(2, 20, jobsInScoreboard.length, 5).setHorizontalAlignment("center").setVerticalAlignment("middle"); // Columns T to X
}