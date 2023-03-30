// global variables
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();
// end of global variables

/**
 * Function to put together email body and display preview in spreadsheet.
 */
function concatEmailBody() {
  concatEmailSubject();
  // connect to spreadsheet and values
  const data = SHEET.getDataRange().getValues();
  const startRow = 6; // spreadsheet row 7
  const headerCol = 2; // column C
  const contentCol = 3; // column D
  const lastRow = SHEET.getLastRow();
  const previewBodyCell = "G3"

  const openingGreeting = getDefaultGreeting_().openingGreeting;
  const closingGreeting = getDefaultGreeting_().closingGreeting;
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

  console.log("Email body:", emailBody);
  SHEET.getRange(previewBodyCell).setValue(emailBody);
}

/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject() {
  const previewSubjectCell = "G2";
  const [m2, d4, m4, c6, d6] = SHEET.getRangeList(['M2','D4','M4','C6','D6']).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(d6);
  const subjectLine = `【${m2}】 ${d4} ${m4}字 ${c6} ${formattedDate}`;
  console.log(subjectLine);
  SHEET.getRange(previewSubjectCell).setValue(subjectLine);
}
