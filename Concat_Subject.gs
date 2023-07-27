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
    taskNameAlert_(3);
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
    // return "ERROR"
  } else if(checkboxCounter >= 2){
    console.error("Too many tasks chosen!");
    taskNameAlert_(2);
    // return "ERROR"
  }
   else {
    return `${taskTitle}${TASK_FOOTER}`
  }
}

/**
 * Triggered when an edit occurs in the spreadsheet.
 * @param {Object} e - Event parameter that can contain information about the context
 *                       that triggered the event (e.g., user, range, value).
 */
function onEdit(e) {
  const range = e.range;
  const newValue = e.value;
  const sheet = e.source.getActiveSheet();
  const checkboxRange = sheet.getRange("B3:B5");
  Logger.log(`newValue: ${newValue}`)
  Logger.log(`log 1: ${checkboxRange.getValues()}`);

  if (range.getColumn() === checkboxRange.getColumn() && range.getRow() >= checkboxRange.getRow() 
      && range.getRow() <= checkboxRange.getRow() + checkboxRange.getHeight() - 1) {

    if (newValue === "TRUE") {
      const values = checkboxRange.getValues().map((row, i) => {
        row[0] = i === range.getRow() - checkboxRange.getRow();
        return row;
      });
      Logger.log(`log 2: ${values}`);
      checkboxRange.setValues(values);
      SpreadsheetApp.flush();
    }
  }
}

function taskNameAlert_(alertNumber) {
  const messageNoTask = "B欄のタスクを選択してください";
  const messageTooManyTasks = "B欄のタスクを1つだけ選択してください";
  const messageCharCountError = "レイアウトチェックなので文字数を削除してください";
  let message = "";
  if(alertNumber == 1){
    message = messageNoTask
  } else if(alertNumber == 2) {
    message = messageTooManyTasks
  } else if(alertNumber == 3){
    message = messageCharCountError
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