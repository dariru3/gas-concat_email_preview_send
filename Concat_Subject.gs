// const SUBJECT_CELL = "I2"
// const SUBJECT_VALUE = SHEET.getRange(SUBJECT_CELL).getValue();
const SUBJECT = {
  cell: "I2",
  value: function() {
    return SHEET.getRange("I2").getValue();
  }
};

/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject_() {
  const taskTranslate = "翻訳依頼";
  const taskAddition = "追つかせ依頼";
  const taskLayoutCheck = "レイアウトチェック依頼"
  const errorHeader = "依頼エラー：";
  const errorNoTaskSelected = "B欄のタスクを選択してください";
  const errorNoCharNoPageCount = "字数かページ数を入力してくさい";
  const errorNoPageCount = "ページ数のみを入力してくさい";
  const taskTitle = getTaskTitle();
  console.log("Task:", taskTitle);
  const [characterCount, pageCount] = getCharacterPageCount();
  let subjectLine = "";
  subjectLine += errorHeader;
  if(taskTitle == "依頼" || taskTitle === undefined | taskTitle == "") {
    showAlert_('no task');
    subjectLine += errorNoTaskSelected;
  } else if(characterCount > 0 && pageCount > 0) {
    subjectLine += errorNoCharNoPageCount;
    showAlert_("no characters or pages count")
  } else if(taskTitle == taskLayoutCheck && characterCount > 0){
    subjectLine += errorNoPageCount;
    showAlert_('character count with layout check mixed');
  } else if((taskTitle == taskTranslate || taskTitle == taskAddition) && characterCount > 0) {
    subjectLine = updateSubjectLine(taskTitle, characterCount);
  } else if(taskTitle == taskLayoutCheck && pageCount > 0) {
    subjectLine = updateSubjectLine(taskTitle, pageCount);
  } else {
    subjectLine = "依頼エラー：不明";
    showAlert_('unknown task error');
  }
  SHEET.getRange(SUBJECT.cell).setValue(subjectLine);
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
  const messageNoTask = "B欄のタスクを選択してください"; // 'no task'
  const messageNoCharPageCount = "字数かページ数を入力してくさい"; // 'no characters or pages count'
  const messageCharCountError = "レイアウトチェックなので文字数を削除してください"; // 'character count with layout check mixed'
  let message = "";

  switch(alertType) {
    case 'no task':
      message = messageNoTask;
      break;
    case 'character count with layout check mixed':
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