function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Menu')
      .addItem('Email Preview', 'concatEmailBody')
      //.addItem('Send Email', '')
      .addToUi();
}

function concatEmailBody() {
  // connect to spreadsheet and values
  const sheetName = "Sheet1";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues()
  const startRow = 6;
  const headerCol = 2;
  const contentCol = 3;
  const lastContentRow = sheet.getRange("D" + sheet.getMaxRows());
  const lastRow = lastContentRow.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const previewCell = "G3"

  // loop through column C and D for email content
  let bodyPreview = "";
  bodyPreview += getDefaultGreeting2()[0] + "\n\n"; // add default greeting; later: overwrite with custom greeting
  for(i=startRow; i<lastRow; i++){
    let header = data[i][headerCol];
    bodyPreview += header + "\n";
    let content = data[i][contentCol];
    if(!content) {
      content = "None"
    }
    bodyPreview += content + "\n\n" 
  }  
  bodyPreview += getDefaultGreeting2()[1] + "\n\n" // add default closing greeting; later: overwrite with custom closing

  console.log(bodyPreview);
  sheet.getRange(previewCell).setValue(bodyPreview);
}

function getDefaultGreeting_() {
  const sheetName = "logic, pull down"
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const greetingCell = "K2"
  const closingCell = "K10"

  const defaultGreeting = sheet.getRange(greetingCell).getValue();
  const greetingSplit = defaultGreeting.split("\n")
  const firstLine = greetingSplit[0];
  const greetingMain = greetingSplit[2];
  // console.log(firstLine);
  // console.log(greetingMain);
  
  const defaultClosing = sheet.getRange(closingCell).getValue();
  
  return [defaultGreeting, defaultClosing]
}

function getDefaultGreeting2(){
  let greeting = "";
  const names = concatToNames();
  console.log(names[0])
  console.log(names[1])
  const myName = getNameFromEmail()
  console.log(myName)
  greeting += names[0] + "\n";
  greeting += names[1] + "\n\n";
  greeting += "お疲れ様です。" + myName + "です。"

  console.log(greeting)

  let closing = "何卒よろしくお願いいたします。\n\n";
  closing += myName;
  return [greeting, closing]

}
function getNameFromEmail() {
  const sheetName = "logic, pull down";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const emailNameList = sheet.getRange("H:I").getValues()
  // console.log(emailNameList)
  const myEmail = Session.getActiveUser().getEmail();
  let myName;
  for(i=0; i < emailNameList.length; i++){
    if(emailNameList[i][0] == myEmail){
      console.log(emailNameList[i][1])
      myName = emailNameList[i][1]
    }
  }
  return myName
}

function concatToNames() {
  const sheetName = "logic, pull down";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const names = sheet.getRange("O2:P6").getValues();
  console.log(names);
  let toNames = "";
  let ccNames = "(";
  for(i=0; i<names.length; i++){
    if(names[i][0] != ''){
      toNames += names[i][0] + "さん、"
    }
    if(names[i][1] != '' && names[i][1] != "Editors"){
      ccNames += names[i][1] + "さん、"
    }
  }
  ccNames = ccNames.slice(0,-1)
  ccNames += ")";
  console.log("To:", toNames);
  console.log("CC:", ccNames);
  return [toNames, ccNames]
}