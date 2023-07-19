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

/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject_() {
  const clientNameCell = 'F3';
  const assignTitleCell = 'F4';
  const characterNumberCell = '04';
  const dueHeaderCell = 'E6';
  const dueDateCell = 'F6';
  const [m2, d4, m4, c6, d6] = SHEET.getRangeList(['M2','D4','M4','C6','D6']).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(d6);
  const subjectLine = `【${m2}】 ${d4} ${m4}字 ${c6} ${formattedDate}`;
  console.log("Subject line:", subjectLine);
  SHEET.getRange(PREVIEW_SUBJECT_CELL).setValue(subjectLine);
}
