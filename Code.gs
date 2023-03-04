function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Menu')
      .addItem('Email Preview', 'concatEmailBody')
      .addItem('Send Email', 'showEmailAlerts')
      .addToUi();
}

// global variables
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
// end of global variables

/**
 * Function to put together email body
 * and display preview in spreadsheet.
 * @returns Email body text.
 */
function concatEmailBody() {
  // connect to spreadsheet and values
  const data = SHEET.getDataRange().getValues();
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
  const toNames = concatNames_()[0];
  const ccNames = concatNames_()[1];
  const myName = concatNames_("Daryl");
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
function concatNames_(getMyName) {
  if(getMyName){
    const myEmailAddress = Session.getActiveUser().getEmail();
    return getNameFromEmailAddress_(myEmailAddress)
  }
  const toAddresses = getEmailAddress_()[0];
  const ccAddresses = getEmailAddress_()[1];

  let toNames = "";
  let ccNames = "";

  const addSanToNames = (list, concatString) => {
    for(let i=0; i<list.length; i++){
      concatString += `${getNameFromEmailAddress_(list[i])}さん、`;
    }  
  }
  addSanToNames(toAddresses, toNames);
  addSanToNames(ccAddresses, ccNames);
  
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
function getNameFromEmailAddress_(address){
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

/**
 * Helper function to gather to and cc email addresses from spreadsheet
 * @returns Array with two lists: to addresses and cc addresses
 */
function getEmailAddress_(){
  // connect to sheet
  let emailAddresses = SHEET.getRange("A3:B11").getValues();
  emailAddresses = emailAddresses.filter(a => a);

  let toAddresses = [];
  let ccAddresses = [];
  for(let i=0; i<emailAddresses.length; i++){
    toAddresses.push(emailAddresses[i][0]);
    ccAddresses.push(emailAddresses[i][1]);
  }

  return [toAddresses, ccAddresses]
}

/**
 * Function to send email.
 */
function sendEmail(){
  const subject = SHEET.getRange("F3").getValue();

  const body = concatEmailBody();
  const toAddresses = getEmailAddress_()[0].join();
  const ccAddresses = getEmailAddress_()[1].join();
  const options = {
    cc: ccAddresses
  }

  try {
    GmailApp.sendEmail(toAddresses, subject, body, options)
    console.log("Success: email sent");
    showEmailAlerts("confirm");
  }
  catch(e){
    throw e
  }
}

function showEmailAlerts(confirm) {
  const ui = SpreadsheetApp.getUi();
  
  if(confirm){
    ui.alert(
      "Email sent.",
      ui.ButtonSet.OK);
    return
  }

  const boxAlert = ui.alert(
     'Is the Box link shared?',
     'Can everyone in this email access the Box folder?',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (boxAlert == ui.Button.YES) {
    // User clicked "Yes".
    const emailAlert = ui.alert(
      'Send email?',
      ui.ButtonSet.YES_NO);
    if (emailAlert == ui.Button.YES) {
      ui.alert('Sending email...')
      sendEmail();
    }
    else {
      ui.alert('Email cancelled.')
    }
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Please check Box folder settings.');
  }
}