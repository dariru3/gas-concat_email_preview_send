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
    subjectLine = taskTitle;
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
  if(taskTitle == "") {
  console.error("No task chosen!");
  taskNameAlert_(1);
  return "エラー: B欄のタスクを選択してください";
  } else if(checkboxCounter >= 2){
    console.error("Too many tasks chosen!");
    taskNameAlert_(2);
  }
  else {
    return `${taskTitle}${TASK_FOOTER}`
  }
}

function onEdit(e) {
  const range = e.range;
  const newValue = e.value;
  Logger.log(`newValue: ${newValue}`)
  Logger.log(`log 1: ${CHECKBOX_RANGE.getValues()}`);

  if (range.getColumn() === CHECKBOX_RANGE.getColumn() && range.getRow() >= CHECKBOX_RANGE.getRow() 
      && range.getRow() <= CHECKBOX_RANGE.getRow() + CHECKBOX_RANGE.getHeight() - 1) {

    if (newValue === "TRUE") {
      const values = CHECKBOX_RANGE.getValues().map((row, i) => {
        row[0] = i === range.getRow() - CHECKBOX_RANGE.getRow();
        return row;
      });
      Logger.log(`log 2: ${values}`);
      CHECKBOX_RANGE.setValues(values);
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