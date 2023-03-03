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
  const data = sheet.getDataRange().getValues();
  const startRow = 6; // spreadsheet row 7
  const headerCol = 2; // column C
  const contentCol = 3; // column D
  const lastContentRow = sheet.getRange(sheet.getLastRow(), contentCol+1);
  const lastRow = lastContentRow.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const previewCell = "G3"
  const projectType = ["新規", "既存", "更新"]
  const removeHeader = '挨拶（任意）';
  const removeContent = ["-", "ー"];
  const characterCounter = "字";

  // loop through column C and D for email content
  bodyPreview = getDefaultGreeting_()[0];
  for(let i=startRow; i<lastRow; i++){
    let header = data[i][headerCol];
    let content = data[i][contentCol];

    if(content) {
      if(removeContent.includes(content)){
        content = "";
      }
      switch(header){
        case removeHeader:
          bodyPreview += `\n\n${content}`;
          break
        case header instanceof Date:
          header = formatDate(header);
          bodyPreview += `\n${header} ${content}`;
          break
        case projectType.includes(header):
          content += characterCounter;
          bodyPreview += `\n${header} ${content}`;
          break
        default:
          bodyPreview += `\n\n${header}\n${content}`;
          break
      }
    }
  }
  bodyPreview += `\n\n${getDefaultGreeting_()[1]}`;

  console.log(bodyPreview);
  sheet.getRange(previewCell).setValue(bodyPreview);
}

function formatDate(date) {
  const d = new Date(date);
  const month = d.getMonth() + 1;
  const day = d.getDate();
  const dayShort = new Intl.DateTimeFormat("ja-JP", { weekday: "narrow" }).format(d);
  return `${month}/${day} (${dayShort})`
}

function getDefaultGreeting_(){
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
  const sheetName = "logic";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const emailNameList = sheet.getRange("H:J").getValues()
  // console.log(emailNameList)
  const myEmail = Session.getActiveUser().getEmail();
  let myName;
  for(i=0; i < emailNameList.length; i++){
    if(emailNameList[i][0] == myEmail){
      console.log(emailNameList[i][2])
      myName = emailNameList[i][2]
    }
  }
  return myName
}

function concatToNames() {
  const sheetName = "logic";
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