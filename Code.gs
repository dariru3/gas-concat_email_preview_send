function concatEmailBody() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Sheet1";
  const sheet = spreadsheet.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const startRow = 9
  const endRow = sheet.getLastRow();
  let bodyPreview = "";
  for(i=startRow; i<data.length; i++){
    console.log(data[i][0])
    for(j=0; j<2; j++){
      //console.log(data[i][j]);
      bodyPreview += data[i][j]+'\n';
    }
  }  
  //console.log(bodyPreview);
  //sheet.getRange("E9").setValue(bodyPreview);
}
