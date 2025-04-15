/**
 * This function reads the input 大会名 from "dashboard" sheet cell C4,
 * scans the "answers" sheet to find rows where either column E or I equals the input,
 * collects the 申請した人の名前 from column C, and outputs them on the "dashboard" sheet
 * starting below cell C5.
 */
function updateDashboard() {
  // Get the active spreadsheet and target sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName("dashboard");
  var answersSheet = ss.getSheetByName("[編集注意]申込フォーム回答");
  
  // Get the input value from cell C4 on the "dashboard" sheet
  var query = dashboardSheet.getRange("C4").getValue();
  
  if (!query) {
    SpreadsheetApp.getUi().alert("大会名 is empty in cell C4. Please enter a valid 大会名.");
    return;
  }
  
  // Get all data from the "answers" sheet
  var dataRange = answersSheet.getDataRange();
  var data = dataRange.getValues();
  
  // Array to hold matching names.
  // We store each name as a single-item array, which is useful for setValues later.
  var matches = [];
  
  // Loop through the data starting after the header row (assumed row 1)
  for (var i = 1; i < data.length; i++) {
    // Column indexes (0-indexed): C => 2, E => 4, I => 8.
    var name = data[i][2];    // 申請した人の名前
    var tournamentE = data[i][4];  // 大会名 from column E
    var tournamentI = data[i][8];  // 大会名 from column I
    
    // If either 大会名 column matches the query, collect the name.
    if (tournamentE === query || tournamentI === query) {
      matches.push([name]);
    }
  }
  
  // Decide where on the dashboard to display results.
  // Now, we output the names starting at cell C6 (i.e. below C5).
  var outputStartRow = 6;
  var outputColumn = 3; // Column C
  
  // Clear previous output from column C starting from row 6 down to the last row.
  var lastRow = dashboardSheet.getLastRow();
  if (lastRow >= outputStartRow) {
    dashboardSheet.getRange(outputStartRow, outputColumn, lastRow - outputStartRow + 1, 1).clearContent();
  }
  
  // Write the results: either the list of names or a "No matches found" message.
  if (matches.length > 0) {
    dashboardSheet.getRange(outputStartRow, outputColumn, matches.length, 1).setValues(matches);
  } else {
    dashboardSheet.getRange(outputStartRow, outputColumn).setValue("No matches found.");
  }
}

/**
 * This function adds a custom menu to your spreadsheet so you can run the script easily.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Update Dashboard', 'updateDashboard')
    .addToUi();
}