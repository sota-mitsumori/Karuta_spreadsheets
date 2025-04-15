/**
 * Convert grade string to a numerical sort value.
 * Recognized grades: A級, B級, C級, D級, E級.
 * Unknown or missing grades are given a high default value (999).
 */
function gradeSortValue(grade) {
  var mapping = {
    "A級": 1,
    "B級": 2,
    "C級": 3,
    "D級": 4,
    "E級": 5
  };
  return mapping[grade] || 999;
}

/**
 * Retrieve the custom order from the "dashboard" sheet column A.
 * This function assumes that the list starts at A2 and continues down
 * to the last non-empty cell in column A.
 */
function getCustomOrder(dashboardSheet) {
  // Adjust the start row if you have a header on row 1.
  var lastRow = dashboardSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  var range = dashboardSheet.getRange("A2:A" + lastRow);
  // Flatten the array and filter out empty values.
  var order = range.getValues().flat().filter(function(item) {
    return item !== "";
  });
  return order;
}

/**
 * For a given name, find its index in the custom order array.
 * If the name is not found, return a number larger than any index,
 * so that it will be placed at the bottom of its grade group.
 */
function getCustomOrderIndex(name, customOrder) {
  var index = customOrder.indexOf(name);
  return index === -1 ? customOrder.length : index;
}

/**
 * This function reads the input 大会名 from cell C4 on the "dashboard" sheet,
 * scans the "answers" sheet for rows where either column E or I equals the input,
 * collects the applicant's name from column C along with the grade from column L,
 * and sorts the matches first by grade (A級 to E級, with unknown/absent grades last)
 * then by the custom order from column A of the "dashboard" sheet.
 * The sorted names are output in column C starting at row 6 (i.e. below C5).
 */
function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the target sheets.
  var dashboardSheet = ss.getSheetByName("dashboard");
  var answersSheet = ss.getSheetByName("[編集注意]申込フォーム回答");
  
  // Check if the required sheets exist.
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('The sheet "dashboard" was not found. Please check the sheet name.');
    return;
  }
  if (!answersSheet) {
    SpreadsheetApp.getUi().alert('The sheet "answers" was not found. Please check the sheet name.');
    return;
  }
  
  // Read the input tournament name from cell C4 on the "dashboard" sheet.
  var query = dashboardSheet.getRange("C4").getValue();
  if (!query) {
    SpreadsheetApp.getUi().alert("大会名 is empty in cell C4. Please enter a valid 大会名.");
    return;
  }
  
  // Retrieve the custom order of names from column A of the dashboard.
  var customOrder = getCustomOrder(dashboardSheet);
  
  // Get all data from the "answers" sheet.
  var data = answersSheet.getDataRange().getValues();
  
  // Array to hold matching entries as objects { name, grade }.
  var matches = [];
  
  // Loop through the data starting after the header row (assumed row 1).
  for (var i = 1; i < data.length; i++) {
    // Column indexes: C => 2, E => 4, I => 8, L => 11.
    var name = data[i][2];           // 申請した人の名前.
    var tournamentE = data[i][4];     // 大会名 from column E.
    var tournamentI = data[i][8];     // 大会名 from column I.
    var grade = data[i][11];          // 級 from column L.
    
    // If the row matches the queried 大会名, add it to our matches.
    if (tournamentE === query || tournamentI === query) {
      matches.push({ name: name, grade: grade });
    }
  }
  
  // First sort by grade, then within the same grade sort by the custom ordering.
  matches.sort(function(a, b) {
    var gradeDiff = gradeSortValue(a.grade) - gradeSortValue(b.grade);
    if (gradeDiff !== 0) {
      return gradeDiff;
    }
    // Both have the same grade, so sort by custom order.
    var aIndex = getCustomOrderIndex(a.name, customOrder);
    var bIndex = getCustomOrderIndex(b.name, customOrder);
    return aIndex - bIndex;
  });
  
// Prepare the output array with blank rows between different grade groups.
  var output = [];
  var previousGrade = null;
  
  matches.forEach(function(entry, idx) {
    // If the grade has changed (and it is not the very first entry), add a blank row.
    if (previousGrade !== null && entry.grade !== previousGrade) {
      output.push([""]); // Blank row for spacing.
    }
    output.push([entry.name]);
    previousGrade = entry.grade;
  });
  
  // Determine where to output the names.
  // We output starting at cell C6 (column C, row 6).
  var outputStartRow = 6;
  var outputColumn = 3; // Column C
  
  // Clear previous output in column C from row 6 downward.
  var lastRow = dashboardSheet.getLastRow();
  if (lastRow >= outputStartRow) {
    dashboardSheet.getRange(outputStartRow, outputColumn, lastRow - outputStartRow + 1, 1).clearContent();
  }
  
  // Write the sorted results: either the list of names or a "No matches found" message.
  if (output.length > 0) {
    dashboardSheet.getRange(outputStartRow, outputColumn, output.length, 1).setValues(output);
  } else {
    dashboardSheet.getRange(outputStartRow, outputColumn).setValue("No matches found.");
  }
}

/**
 * Adds a custom menu to the spreadsheet so that you can run the script easily.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Update Dashboard', 'updateDashboard')
    .addToUi();
}