/**
 * Function to put together email body and display preview in spreadsheet.
 */
function concatEmailBody() {
  const START_ROW = 6; // spreadsheet row 7
  const HEADER_COL = 4; // column C
  const CONTENT_COL = 5; // column D
  const data = SHEET.getDataRange().getValues();
  const lastRow = SHEET.getLastRow();
  concatEmailSubject();

  const [openingGreeting, closingGreeting] = getDefaultGreeting_();
  let emailBody = openingGreeting
  // loop through column C and D for email content
  for(let i=START_ROW; i<lastRow; i++){
    let header = data[i][HEADER_COL];
    let content = data[i][CONTENT_COL];
    // add to email body if there is content
    if(content) {
      emailBody += formatHeader_Content_(header, content);
    }
  }
  emailBody += `\n\n${closingGreeting}`;

  SHEET.getRange(BODY.cell).setValue(emailBody);
}

/**
 * Helper function to compile default opening and closing greetings.
 * @returns Opening and closing greetings.
 */
function getDefaultGreeting_(){
  const myName = getNameFromEmailAddress_(MY_EMAIL)
  let [opening, ccNames] = concatNames_();
  if(ccNames){
    ccNames = ccNames.slice(0,-1); // remove the final comma
    opening += `\n(${ccNames})`;
  }
  opening += `\n\nお疲れ様です。${myName}です。`;

  const closing = `何卒よろしくお願いいたします。\n\n${myName}`;

  return [opening, closing]
}