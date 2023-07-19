/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject_() {
  const clientNameCell = 'F3';
  const assignTitleCell = 'F4';
  const dueHeaderCell = 'E6';
  const dueDateCell = 'F6';
  const taskTitle = getTaskTitle();
  const characterCount = getCharacterCount();
  const [clientName, assignTitle, dueHeader, dueDate] = SHEET.getRangeList([clientNameCell, assignTitleCell, dueHeaderCell, dueDateCell]).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(dueDate);
  let subjectLine = "";
  if(taskTitle == "翻訳依頼" || taskTitle == "追つかせ依頼") {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${characterCount}字 ${dueHeader} ${formattedDate}`;
  } else if (taskTitle == "レイアウトチェック依頼") {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${dueHeader} ${formattedDate}`;
  } else {
    console.error("Task title error!")
  }
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
    return "ERROR"
  } else if(checkboxCounter >= 2){
    console.error("Too many tasks chosen!");
    taskNameAlert_(2);
    return "ERROR"
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

function getCharacterCount() {
  const characterCountValues = SHEET.getRange('F9:F10').getValues();
  let count = 0;
  for(let i = 0; i < characterCountValues.length; i++){
    count += characterCountValues[i][0];
  }
  console.log(count);
}