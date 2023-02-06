function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Email Menu')
      .addItem('Email Preview', 'concatEmailBody')
      //.addItem('Send Email', '')
      .addToUi();
}


function concatEmailBody() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Sheet1";
  const sheet = spreadsheet.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const startRow = 6
  const endRow = sheet.getLastRow();
  const headerCol = 2;
  const contentCol = 3;
  const previewCell = "G8"

  let bodyPreview = "";
  for(i=startRow; i<data.length; i++){
    console.log("header:", data[i][headerCol]);
    let header = data[i][headerCol];
    bodyPreview += header + "\n";

    console.log("contents:",data[i][contentCol]);
    let content = data[i][contentCol];
    bodyPreview += content + "\n\n" 
  }  
  console.log(bodyPreview);
  sheet.getRange(previewCell).setValue(bodyPreview);
}
