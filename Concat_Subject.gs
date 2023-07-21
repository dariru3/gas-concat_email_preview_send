/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject_() {
  let subjectLine = "";
  const translateTask = "翻訳依頼";
  const additionalTask = "追つかせ依頼";
  const layoutCheckTask = "レイアウトチェック依頼"
  const taskTitle = getTaskTitle();
  const characterCount = getCharacterCount();
  const clientName = checkForSama(CLIENT_CELL);
  const [assignTitle, dueHeader, dueDate] = SHEET.getRangeList([ASSIGN_CELL, DUE_HEADER_CELL, DUE_DATE_CELL]).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(dueDate);
  
  if(taskTitle == translateTask || taskTitle == additionalTask) {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${characterCount}字 ${dueHeader} ${formattedDate}`;
  } else if (taskTitle == layoutCheckTask) {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${dueHeader} ${formattedDate}`;
  } else {
    console.error("Task title error!")
  }
  if(taskTitle == layoutCheckTask && characterCount > 0){
    subjectLine = "What are you asking for?"
  }
  console.log("Subject line:", subjectLine);
  SHEET.getRange(SUBJECT_CELL).setValue(subjectLine);
}

function getTaskTitle() {
  let taskTitle = "";
  let checkboxCounter = 0;
  for(let i = 0; i < TASK_VALUES.length; i++) {
    console.log(TASK_VALUES[i])
    if(TASK_VALUES[i][1] == true){
      checkboxCounter += 1;
      taskTitle = TASK_VALUES[i][0]
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
    return `${taskTitle}${TASK_FOOTER}`
  }
}

function taskNameAlert_(alertNumber) {
  const messageNoTask = "Please choose a task in Column B";
  const messageTooManyTasks = "Please choose only 1 task in Column B";
  let message = "";
  if(alertNumber == 1){
    message = messageNoTask
  } else if(alertNumber == 2) {
    message = messageTooManyTasks
  } else {
    console.error("taskNameAlert error!")
  }
  UI.alert(
    message,
    UI.ButtonSet.OK
  );
}

function getCharacterCount() {
  let count = 0;
  for(let i = 0; i < CHAR_COUNT_VALUES.length; i++){
    count += CHAR_COUNT_VALUES[i][0];
  }
  console.log(count);
  return count
}

function checkForSama(cell) {
  let clientName = SHEET.getRange(cell).getValue();
  if(clientName.slice(-1) !== "様") {
    clientName += "様";
  }
  return clientName
}