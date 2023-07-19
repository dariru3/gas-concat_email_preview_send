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
  const taskTitle = getTaskTitle();
  const [clientName, assignTitle, characterNumber, dueHeader, dueDate] = SHEET.getRangeList([clientNameCell, assignTitleCell, characterNumberCell, dueHeaderCell, dueDateCell]).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(dueDate);
  const subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${characterNumber}字 ${dueHeader} ${formattedDate}`;
  console.log("Subject line:", subjectLine);
  SHEET.getRange(PREVIEW_SUBJECT_CELL).setValue(subjectLine);
}

function getTaskTitle() {
  const taskFooter = SHEET.getRange('A2').getValue();
  const taskValues = SHEET.getRange('A3:B5').getValues();
  let taskTitle = "";
  let checkboxCounter = 0;
  for(let i = 0; i < taskValues.length; i++) {
    console.log(taskValues[i])
    if(taskValues[i][1] == true){
      checkboxCounter += 1;
      taskTitle = taskValues[i][0]
    }
  }
  if(taskTitle == ""){
    console.error("No task chosen!");
    taskNameAlert_(1);
  } else if(checkboxCounter >= 2){
    console.error("Too many tasks chosen!");
    taskNameAlert_(2);
  }
   else {
    return `${taskTitle}${taskFooter}`
  }
}

function taskNameAlert_(alertNumber) {
  const messageNoTask = "Please choose a task in Column B";
  const messageTooManyTasks = "Please choose only 1 task in Column B";
  let message = "";
  if(alertNumber == 1){
    message = messageNoTask
  } else {
    message = messageTooManyTasks
  }
  UI.alert(
    message,
    UI.ButtonSet.OK
  );
}
