function addEditResponseLinks() {
  try {
    // automatically get the form id
    var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
    
    if (!formUrl) {
      SpreadsheetApp.getUi().alert('Cannot find');
      return;
    }
    
    var form = FormApp.openByUrl(formUrl);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var responses = form.getResponses();
    
    // setup a map based on timestamp
    var editLinks = {};
    for (var i = 0; i < responses.length; i++) {
      editLinks[responses[i].getTimestamp().toString()] = responses[i].getEditResponseUrl();
    }
    
    // find the index of last column
    var lastColumn = sheet.getLastColumn();
    
    // interate the record
    for (var j = 1; j < data.length; j++) {
      var timestamp = data[j][0].toString(); // assume the first column is timestamp
      if (editLinks[timestamp]) {
        // add a link at last column
        sheet.getRange(j + 1, lastColumn + 1).setValue(editLinks[timestamp]);
      }
    }
    
    SpreadsheetApp.getUi().alert('Task Done! Check the last column');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error：' + e.message);
  }
}
