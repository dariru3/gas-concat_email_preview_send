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
  const characterNumberCell = 'O4';
  const dueHeaderCell = 'E6';
  const dueDateCell = 'F6';
  const [clientName, assignTitle, characterNumber, dueHeader, dueDate] = SHEET.getRangeList([clientNameCell, assignTitleCell, characterNumberCell, dueHeaderCell, dueDateCell]).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(dueDate);
  const subjectLine = `【${clientName}】 ${assignTitle} ${characterNumber}字 ${dueHeader} ${formattedDate}`;
  console.log("Subject line:", subjectLine);
  SHEET.getRange(PREVIEW_SUBJECT_CELL).setValue(subjectLine);
}

function getTaskName() {
  const taskFooter = SHEET.getRange('A2').getValue();
  const taskValues = SHEET.getRange('A3:B5').getValues();
  let taskTitle = "";
  for(let i = 0; i < taskValues.length; i++) {
    console.log(taskValues[i])
    if(taskValues[i][1] == true){
      taskTitle = taskValues[i][0]
      break
    }
  }
  if(taskTitle == ""){
    console.error("No task chosen!")
  } else {
    console.log(`Return: ${taskTitle}${taskFooter}`)
  }
}
