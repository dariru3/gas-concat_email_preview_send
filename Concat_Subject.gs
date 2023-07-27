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
  
  if((taskTitle == translateTask || taskTitle == additionalTask) && characterCount > 0) {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${characterCount}字 ${dueHeader} ${formattedDate}`;
  } else if (taskTitle == layoutCheckTask) {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${dueHeader} ${formattedDate}`;
  } else {
    console.error("Task title error!")
    subjectLine = taskTitle;
  }
  if(taskTitle == layoutCheckTask && characterCount > 0){
    taskNameAlert_('characters and pages mixed');
  }
  console.log("Subject line:", subjectLine);
  SHEET.getRange(SUBJECT_CELL).setValue(subjectLine);
}

function getTaskTitle() {
  let taskTitle = "";
  for(let i = 0; i < TASK_VALUES.length; i++) {
    console.log(TASK_VALUES[i])
    if(TASK_VALUES[i][1] == true){
      taskTitle = TASK_VALUES[i][0]
    }
  }
  if(taskTitle == "") {
  console.error("No task chosen!");
  taskNameAlert_('no task');
  return "エラー: B欄のタスクを選択してください";
  } else {
    return `${taskTitle}${TASK_FOOTER}`
  }
}

function onEdit(e) {
  const range = e.range;
  const newValue = e.value;

  if (range.getColumn() === CHECKBOX_RANGE.getColumn() && range.getRow() >= CHECKBOX_RANGE.getRow() 
      && range.getRow() <= CHECKBOX_RANGE.getRow() + CHECKBOX_RANGE.getHeight() - 1) {

    if (newValue === "TRUE") {
      const values = CHECKBOX_RANGE.getValues().map((row, i) => {
        row[0] = i === range.getRow() - CHECKBOX_RANGE.getRow();
        return row;
      });
      CHECKBOX_RANGE.setValues(values);
      SpreadsheetApp.flush();
    }
  }
}

function taskNameAlert_(alertType) {
  const messageNoTask = "B欄のタスクを選択してください";
  const messageCharCountError = "レイアウトチェックなので文字数を削除してください";
  let message = "";
  if(alertType == 'no task'){
    message = messageNoTask
  } else if(alertType == 'characters and pages mixed'){
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