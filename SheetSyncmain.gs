function onEdit(e) {
  if (e && e.range) {
    var sheet = e.range.getSheet();
    var range = e.range;
    var col = range.getColumn();

    // Replace YOUR_COLUMN_NUMBER with the actual column numbers
    var targetColumn = 32;
    var fullUrlColumn = 33;

    if (col === targetColumn) {
      var url = range.getValue();
      if (url && url.startsWith("http")) {
        // Check if the URL is already used in the same column (column 33)
        var valuesInColumn = sheet.getRange(2, fullUrlColumn, sheet.getLastRow() - 1, 1).getValues();
        var isDuplicate = false;

        for (var i = 0; i < valuesInColumn.length; i++) {
          if (valuesInColumn[i][0] === url) {
            isDuplicate = true;
            break;
          }
        }

        if (isDuplicate) {
          // Display a message, or you can remove the next line if you don't want the prompt.
          Browser.msgBox('This link is already pasted in the column.');
          range.clearContent();
        } else {
          var timezone = "CST"; // Set the desired timezone
          var timestamp = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd HH:mm:ss");
          var hyperlink = '=HYPERLINK("' + url + '", "Image Link ' + timestamp + '")';
          range.setValue(hyperlink);

          // Set the full URL in the next column (column 33)
          var fullUrlRange = sheet.getRange(range.getRow(), fullUrlColumn);
          fullUrlRange.setValue(url);
        }
      }
    }
  }
}
