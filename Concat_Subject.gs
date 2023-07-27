/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject_() {
  let subjectLine = "";
  const translateTask = "翻訳依頼";
  const additionalTask = "追つかせ依頼";
  const layoutCheckTask = "レイアウトチェック依頼"
  const taskTitle = getTaskTitle();
  Logger.log(`taskTitle: ${taskTitle}`);
  if(taskTitle == "" || taskTitle === undefined) {
    taskNameAlert_('no task');
    subjectLine = "エラー： B欄のタスクを選択してください"
    SHEET.getRange(SUBJECT_CELL).setValue(subjectLine);
    return
  }
  const [characterCount, pageCount] = getCharacterCount();
  const clientName = checkForSama(CLIENT_CELL);
  const [assignTitle, dueHeader, dueDate] = SHEET.getRangeList([ASSIGN_CELL, DUE_HEADER_CELL, DUE_DATE_CELL]).getRanges().map(range => range.getValues().flat())
  const formattedDate = formatDate_(dueDate);
  
  if((taskTitle == translateTask || taskTitle == additionalTask) && characterCount > 0) {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${characterCount}字 ${dueHeader} ${formattedDate}`;
  } else if (taskTitle == layoutCheckTask && pageCount > 0) {
    subjectLine = `【${clientName}】 ${assignTitle} ${taskTitle} ${dueHeader} ${formattedDate}`;
  } else {
    console.error("Task title error!")
    subjectLine = `エラー： 字数かページ数を入力してくさい`
    taskNameAlert_("no characters or pages count")
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
  // taskNameAlert_('no task');
  return // "エラー: B欄のタスクを選択してください";
  } else {
    return `${taskTitle}${TASK_FOOTER}`
  }
}

function onEdit(e) {
  const range = e.range;
  const newValue = e.value;
  const sheet = e.source.getActiveSheet();

  if (range.getColumn() === CHECKBOX_RANGE.getColumn() && range.getRow() >= CHECKBOX_RANGE.getRow() 
      && range.getRow() <= CHECKBOX_RANGE.getRow() + CHECKBOX_RANGE.getHeight() - 1) {

    if (newValue === "TRUE") {
      const values = CHECKBOX_RANGE.getValues().map((row, i) => {
        row[0] = i === range.getRow() - CHECKBOX_RANGE.getRow();
        return row;
      });
      CHECKBOX_RANGE.setValues(values);

      // Depending on the row that was edited, show or hide the 12th row.
      if (range.getRow() === 5) {
        // Unhide row 12 if B5 was checked
        sheet.showRows(12);
        sheet.hideRows(9, 3);
      } else if (range.getRow() === 3 || range.getRow() === 4) {
        // Hide row 12 if B3 or B4 was checked
        sheet.hideRows(12);
        sheet.showRows(9, 3);
      }
      SpreadsheetApp.flush();
    } else {
      sheet.showRows(9, 4)
    } 
  } 
}

function taskNameAlert_(alertType) {
  const messageNoTask = "B欄のタスクを選択してください";
  const messageCharCountError = "レイアウトチェックなので文字数を削除してください";
  const messageNoCharPageCount = "字数かページ数を入力してくさい";
  let message = "";

  switch(alertType) {
    case 'no task':
      message = messageNoTask;
      break;
    case 'characters and pages mixed':
      message = messageCharCountError;
      break;
    case 'no characters or pages count':
      message = messageNoCharPageCount;
      break;
    default:
      console.error("taskNameAlert error!");
      message = "エラー：不明なアラートタイプ"
  }

  UI.alert(
    message,
    UI.ButtonSet.OK
  );
}


function getCharacterCount() {
  let charCount = 0;
  for(let i = 0; i < CHAR_COUNT_VALUES.length; i++){
    charCount += CHAR_COUNT_VALUES[i][0];
  }

  return [charCount, PAGE_COUNT_VALUE]
}

function checkForSama(cell) {
  let clientName = SHEET.getRange(cell).getValue();
  if(clientName.slice(-1) !== "様") {
    clientName += "様";
  }
  return clientName
}