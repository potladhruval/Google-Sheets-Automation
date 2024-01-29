function onEdit(e) {
  if (e) {
    handleEditTesterSheet(e);
    handleHyperlinkTesterSheet(e);
    //getDataFromMainSheet();
  }
}

function handleEditTesterSheet(e) {
  if (e && e.range) {
    var editedSheetId = e.source.getId();
    var mainSheetId = '1BsVThfiuDzZ-oGp8O6BPaN_nBf85uuc-IgN3fPF7Ngc';

    if (editedSheetId !== mainSheetId) {
      var mainSpreadsheet = SpreadsheetApp.openById(mainSheetId);
      var mainSheet = mainSpreadsheet.getSheetByName('Dallas_DT');
      var editedRange = e.range;
      var editedValue = e.value;
      var editedA1Notation = editedRange.getA1Notation();

      // Only update until column 31 in the main sheet
      if (editedRange.getColumn() <= 31) {
        var mainRange = mainSheet.getRange(editedA1Notation);
        mainRange.setValue(editedValue);
      }
    }
  }
}

function handleHyperlinkTesterSheet(e) {
  if (e && e.range) {
    var editedSheetId = e.source.getId();
    var mainSheetId = '1BsVThfiuDzZ-oGp8O6BPaN_nBf85uuc-IgN3fPF7Ngc';

    if (editedSheetId !== mainSheetId) {
      var range = e.range;
      var col = range.getColumn();
      var targetColumn = 32;
      var fullUrlColumn = 33;

      if (col === targetColumn) {
        var url = range.getValue();
        if (url && url.startsWith("http")) {
          var mainSheet = SpreadsheetApp.openById(mainSheetId).getSheetByName('Dallas_DT');
          var valuesInColumn = mainSheet.getRange(2, fullUrlColumn, mainSheet.getLastRow() - 1, 1).getValues();
          var isDuplicate = false;

          for (var i = 0; i < valuesInColumn.length; i++) {
            if (valuesInColumn[i][0] === url) {
              isDuplicate = true;
              break;
            }
          }

          if (isDuplicate) {
            Browser.msgBox('This link has already been pasted.');
            range.clearContent();
          } else {
            var timezone = "CST";
            var timestamp = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd HH:mm:ss");
            var hyperlink = '=HYPERLINK("' + url + '", "Image Link ' + timestamp + '")';
            range.setValue(hyperlink);

            // Add timestamp to column 32 in the main sheet
            var mainSheet = SpreadsheetApp.openById(mainSheetId).getSheetByName('Dallas_DT');
            var mainRange = mainSheet.getRange(e.range.getA1Notation());
            mainSheet.getRange(mainRange.getRow(), targetColumn).setValue(timestamp);

            // Add complete link to column 33 in the main sheet
            var fullUrlRange = mainSheet.getRange(mainRange.getRow(), fullUrlColumn);
            fullUrlRange.setValue(url);
          }
        }
      }
    }
  }
}

function getDataFromMainSheet(startRow, endRow) {
  try {
    var mainSpreadsheetId = '1BsVThfiuDzZ-oGp8O6BPaN_nBf85uuc-IgN3fPF7Ngc';
    var mainSheetName = 'Dallas_DT';
    var testerSpreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var testerSheetName = 'Sheet1';

    console.log("Starting getDataFromMainSheet...");

    var mainSpreadsheet = SpreadsheetApp.openById(mainSpreadsheetId);
    var mainSheet = mainSpreadsheet.getSheetByName(mainSheetName);

    var testerSpreadsheet = SpreadsheetApp.openById(testerSpreadsheetId);
    var testerSheet = testerSpreadsheet.getSheetByName(testerSheetName);

    startRow = 1630 || 1; // Default to row 1 if not provided
    endRow = 1653 || mainSheet.getLastRow(); // Default to the last row if not provided

    console.log("Fetching rows from main sheet: " + startRow + " to " + endRow);

    // Get values as strings for the specified range
    var mainRange = mainSheet.getRange(startRow, 1, endRow - startRow + 1, mainSheet.getLastColumn());
    var mainValues = mainRange.getValues().map(row => row.map(cell => cell instanceof Date ? cell.toString() : (cell !== null ? cell.toString() : '')));

    // Log the values before updating the tester sheet
    console.log("Main values to set:", JSON.stringify(mainValues));

    // Update tester sheet with new data
    var testerRange = testerSheet.getRange(startRow, 1, mainValues.length, mainValues[0].length);
    testerRange.clearContent(); // Clear existing content in the specified range
    testerRange.setValues(mainValues);

    console.log("getDataFromMainSheet completed.");

  } catch (error) {
    console.error("Error in getDataFromMainSheet: " + error.message);
    console.error("Error stack: " + error.stack);
  }
}




// Uncomment the line below to set up triggers initially
// setUpTriggers();
