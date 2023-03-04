function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Menu')
      .addItem('Email Preview', 'concatEmailBody')
      //.addItem('Send Email', '')
      .addToUi();
}

/**
 * Function to put together email body
 * and display preview in spreadsheet.
 * @returns Email body text.
 */
function concatEmailBody() {
  // connect to spreadsheet and values
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const data = sheet.getDataRange().getValues();
  const startRow = 6; // spreadsheet row 7
  const headerCol = 2; // column C
  const contentCol = 3; // column D
  const lastContentRow = sheet.getRange(sheet.getLastRow(), contentCol+1);
  const lastRow = lastContentRow.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  const previewCell = "G3"

  const openingGreeting = getDefaultGreeting_()[0];
  const closingGreeting = getDefaultGreeting_()[1];
  let emailBody = openingGreeting
  // loop through column C and D for email content
  for(let i=startRow; i<lastRow; i++){
    let header = data[i][headerCol];
    let content = data[i][contentCol];
    // add to email body if there is content
    if(content) {
      emailBody += formatHeaderContent_(header, content);
    }
  }
  emailBody += `\n\n${closingGreeting}`;

  console.log(emailBody);
  sheet.getRange(previewCell).setValue(emailBody);
  return emailBody
}

/**
 * Helper function to format headers and content.
 * @param {string} header Header text
 * @param {string} content Content text
 * @returns Header and content text formatted.
 */
function formatHeaderContent_(header, content){
  const removeContent = new Set(["-", "ー"]);
  const removeHeader = new Set(['挨拶（任意）']);
  const projectType = new Set(["新規", "既存", "更新"]);
  const characterCounter = "字";

  let headersContent = "";
  if(removeContent.has(content)){
    content = "";
  }
  // format headers according to type
  if(removeHeader.has(header)){
    headersContent += `\n\n${content}`;
  } else if(projectType.has(header)){
    content += characterCounter;
    headersContent += `\n${header} ${content}`;
  } else if(header instanceof Date){
    header = formatDate_(header);
    headersContent += `\n${header} ${content}`;
  } else {
    headersContent += `\n\n${header}\n${content}`;
  }

  return headersContent
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
  let opening = "";
  const toNames = getNamesFromAddresses_()[0];
  const ccNames = getNamesFromAddresses_()[1];
  const myName = getNamesFromAddresses_("Daryl");
  opening += `${toNames}\n`;
  opening += `${ccNames}\n\n`;
  opening += `お疲れ様です。${myName}です。`;

  const closing = `何卒よろしくお願いいたします。\n\n${myName}`;
  return [opening, closing]
}

/**
 * Helper function to get preferred names from email address.
 * @param {any} getMyName Optional: add to get user's name.
 * @returns Either user's name or to and cc names.
 */
function getNamesFromAddresses_(getMyName) {
  if(getMyName){
    const myEmailAddress = Session.getActiveUser().getEmail();
    return getNameFromAddress_(myEmailAddress)
  }
  const toAddresses = getEmailAddresses_()[0];
  const ccAddresses = getEmailAddresses_()[1];

  let toNames = "";
  let ccNames = "";

  const loopAddressList = (list, names) => {
    for(let i=0; i<list.length; i++){
      names += `${getNameFromAddress_(list[i])}さん、`;
    }  
  }
  loopAddressList(toAddresses, toNames);
  loopAddressList(ccAddresses, ccNames);
  
  if(ccNames){
    ccNames = ccNames.slice(0,-1); // remove the final comma
    ccNames = `(${ccNames})`;
  }

  return [toNames, ccNames]
}

/**
 * Helper function to get name from an email address.
 * @param {string} address Email address to look up.
 * @returns Preferred name when sending emails.
 */
function getNameFromAddress_(address){
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("logic");
  const emailNameList = referenceSheet.getRange("H:J").getValues();
  const noNames = new Set(["edit_all@link-cc.co.jp"]);

  const nameLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => nameLookup[row[0]] = row[2]);

  if(address != "" && !noNames.has(address) && address in nameLookup){
    return nameLookup[address]
  } else {
    console.error("Email address not found.")
  }
}

function getEmailAddresses_(){
  // connect to sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  const emailAddresses = sheet.getRange("A3:B11").getValues();

  let toAddresses = [];
  let ccAddresses = [];
  for(let i=0; i<emailAddresses.length; i++){
    toAddresses.push(emailAddresses[i][0]);
    ccAddresses.push(emailAddresses[i][1]);
  }

  return [toAddresses, ccAddresses]
}