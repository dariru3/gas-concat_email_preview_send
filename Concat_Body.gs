/**
 * Function to put together email body and display preview in spreadsheet.
 */
function concatEmailBody() {
  concatEmailSubject_();
  // connect to spreadsheet and values
  const data = SHEET.getDataRange().getValues();
  const lastRow = SHEET.getLastRow();

  const openingGreeting = getDefaultGreeting_().openingGreeting;
  const closingGreeting = getDefaultGreeting_().closingGreeting;
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

  console.log("Email body:", emailBody);
  SHEET.getRange(PREVIEW_BODY_CELL).setValue(emailBody);
}