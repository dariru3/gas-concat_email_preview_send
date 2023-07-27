// const SUBJECT_CELL = "I2"
// const SUBJECT_VALUE = SHEET.getRange(SUBJECT_CELL).getValue();
const SUBJECT = {
  cell: "I2",
  value: SHEET.getRange("I2").getValue()
}

/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject_() {
  const translateTask = "翻訳依頼";
  const additionalTask = "追つかせ依頼";
  const layoutCheckTask = "レイアウトチェック依頼"
  let subjectLine = "";
  const taskTitle = getTaskTitle();
  if(taskTitle == "" || taskTitle === undefined) {
    showAlert_('no task');
    subjectLine = "エラー： B欄のタスクを選択してください"
    SHEET.getRange(SUBJECT.cell).setValue(subjectLine);
    return
  }
  const [characterCount, pageCount] = getCharacterPageCount();
  
  if((taskTitle == translateTask || taskTitle == additionalTask) && characterCount > 0) {
    subjectLine = updateSubjectLine(taskTitle, characterCount);
  } else if (taskTitle == layoutCheckTask && pageCount > 0) {
    subjectLine = updateSubjectLine(taskTitle, characterCount);
  } else {
    subjectLine = `エラー： 字数かページ数を入力してくさい`
    showAlert_("no characters or pages count")
  }
  if(taskTitle == layoutCheckTask && characterCount > 0){
    subjectLine = `エラー： 字数かページ数を入力してくさい`
    showAlert_('characters and pages mixed');
  }
  SHEET.getRange(SUBJECT_CELL).setValue(subjectLine);
}

function updateSubjectLine(taskTitle, characterCount) {
  const fileAndSectionNamesCell = 'F4';
  const dueDateCell = 'F6';
  const dueHeaderCell = 'E6'; // '〆切'
  const [assignTitle, dueHeader, dueDate] = SHEET.getRangeList([fileAndSectionNamesCell, dueHeaderCell, dueDateCell]).getRanges().map(range => range.getValues().flat())

  const clientCell = 'F3';
  const clientName = checkForSama(clientCell);
  const formattedDate = formatDate_(dueDate);
  if(characterCount == 0){
    return `【${clientName}】 ${assignTitle} ${taskTitle} ${dueHeader} ${formattedDate}`;
  } else {
    return `【${clientName}】 ${assignTitle} ${taskTitle} ${characterCount}字 ${dueHeader} ${formattedDate}`;
  }
}

function getTaskTitle() {
  const taskFooter = SHEET.getRange('A2').getValue(); // "依頼"
  const taskValues = SHEET.getRange('A3:B5').getValues();
  let taskTitle = "";
  for(let i = 0; i < taskValues.length; i++) {
    console.log(taskValues[i])
    if(taskValues[i][1] == true){
      taskTitle = taskValues[i][0]
    }
  }
  return `${taskTitle}${taskFooter}`
}

function onEdit(e) {
  const checkboxRange = SHEET.getRange("B3:B5");
  const range = e.range;
  const newValue = e.value;
  const sheet = e.source.getActiveSheet();
  
  if (range.getColumn() === checkboxRange.getColumn() && range.getRow() >= checkboxRange.getRow() 
      && range.getRow() <= checkboxRange.getRow() + checkboxRange.getHeight() - 1) {

    if (newValue === "TRUE") {
      const values = checkboxRange.getValues().map((row, i) => {
        row[0] = i === range.getRow() - checkboxRange.getRow();
        return row;
      });
      checkboxRange.setValues(values);

      // Depending on the row that was edited, show or hide rows.
      if (range.getRow() === 5) {
        sheet.showRows(12);
        sheet.hideRows(9, 3);
      } else if (range.getRow() === 3 || range.getRow() === 4) {
        sheet.hideRows(12);
        sheet.showRows(9, 3);
      }
      SpreadsheetApp.flush();
    } else {
      sheet.showRows(9, 4)
    } 
  } 
}

function showAlert_(alertType) {
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
      message = "不明なエラー"
  }

  UI.alert(
    message,
    UI.ButtonSet.OK
  );
}


function getCharacterPageCount() {
  const charCountValues = SHEET.getRange('F9:F10').getValues();
  const pageCountValue = SHEET.getRange('F12').getValue();
  let charCount = 0;
  for(let i = 0; i < charCountValues.length; i++){
    charCount += charCountValues[i][0];
  }

  return [charCount, pageCountValue]
}

function checkForSama(cell) {
  let clientName = SHEET.getRange(cell).getValue();
  if(clientName.slice(-1) !== "様") {
    clientName += "様";
  }

  return clientName
}