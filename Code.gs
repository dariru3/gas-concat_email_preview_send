function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Menu')
      .addItem('Email Preview', 'concatEmailBody')
      //.addItem('Send Email', '')
      .addToUi();
}

/**
 * Function to display email preview in spreadsheet.
 */
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
  const projectType = new Set(["新規", "既存", "更新"]);
  const removeHeader = new Set(['挨拶（任意）']);
  const removeContent = new Set(["-", "ー"]);
  const characterCounter = "字";

  // loop through column C and D for email content
  bodyPreview = getDefaultGreeting_()[0];
  for(let i=startRow; i<lastRow; i++){
    let header = data[i][headerCol];
    let content = data[i][contentCol];
    // add to email body if there is content
    if(content) {
      if(removeContent.has(content)){
        content = "";
      }
      // format different headers
      if(removeHeader.has(header)){
        bodyPreview += `\n\n${content}`;
      } else if(projectType.has(header)){
        content += characterCounter;
        bodyPreview += `\n${header} ${content}`;
      } else if(header instanceof Date){
        header = formatDate_(header);
        bodyPreview += `\n${header} ${content}`;
      } else {
        bodyPreview += `\n\n${header}\n${content}`;
      }
    }
  }
  bodyPreview += `\n\n${getDefaultGreeting_()[1]}`;

  console.log(bodyPreview);
  sheet.getRange(previewCell).setValue(bodyPreview);
}

/**
 * Helper function to formate date value
 * with Japanese text.
 * @param {date} date Date value. 
 * @returns Date formatted with Japanese day.
 */
function formatDate_(date) {
  const d = new Date(date);
  const month = d.getMonth() + 1;
  const day = d.getDate();
  const dayShort = new Intl.DateTimeFormat("ja-JP", { weekday: "narrow" }).format(d);
  return `${month}/${day} (${dayShort})`
}

/**
 * Helper function to compile default
 * opening and closing greetings.
 * @returns Opening and closing greetings.
 */
function getDefaultGreeting_(){
  let greeting = "";
  const toNames = getNamesFromAddresses_()[0];
  const ccNames = getNamesFromAddresses_()[1];
  const myName = getNamesFromAddresses_("Daryl");
  greeting += `${toNames}\n`;
  greeting += `${ccNames}\n\n`;
  greeting += `お疲れ様です。${myName}です。`;

  const closing = `何卒よろしくお願いいたします。\n\n${myName}`;
  return [greeting, closing]
}

/**
 * Helper function to get preferred names from email address.
 * @param {any} getMyName Optional: add to get user's name.
 * @returns Either user's name or to and cc names.
 */
function getNamesFromAddresses_(getMyName) {
  // connect to source sheet and reference sheet
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const emailAddresses = sourceSheet.getRange("A3:B11").getValues();

  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logic");
  const emailNameList = referenceSheet.getRange("H:J").getValues();
  const noNames = new Set(["edit_all@link-cc.co.jp"]);

  const nameLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => nameLookup[row[0]] = row[2]);

  if(getMyName){
    getMyName = Session.getActiveUser().getEmail();
    return nameLookup[getMyName];
  }

  let toNames = "";
  let ccNames = "(";
  for(let i=0; i<emailAddresses.length; i++){
    let toAddress = emailAddresses[i][0];
    let ccAddress = emailAddresses[i][1];
    if(toAddress != "" && toAddress in nameLookup){
      toNames += `${nameLookup[toAddress]}さん、`
    }
    if(ccAddress != "" && !noNames.has(ccAddress) && ccAddress in nameLookup){
      ccNames += `${nameLookup[ccAddress]}さん、`
    }
  }
  ccNames = ccNames.slice(0,-1); // remove the final comma
  ccNames += ")";

  return [toNames, ccNames]
}