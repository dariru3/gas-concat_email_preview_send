/**
 * Function to put together email subject and display preview in spreadsheet.
 */
function concatEmailSubject() {
  const taskTitle = getTaskTitle();
  console.log("Task:", taskTitle);
  const [characterCount, pageCount] = getCharacterPageCount();
  let subjectLine = "";
  subjectLine += ERRORS.HEADER;
  if(taskTitle === undefined || taskTitle == "") {
    subjectLine += ERRORS.NO_TASK;
    showAlert_(ERRORS.NO_TASK);
  } else if(characterCount == 0 && pageCount == 0) {
    subjectLine += ERRORS.NO_COUNT;
    showAlert_(ERRORS.NO_COUNT)
  } else if((taskTitle == TASK_TYPES.TRANSLATE || taskTitle == TASK_TYPES.ADDITION) && characterCount > 0) {
    subjectLine = updateSubjectLine(taskTitle, characterCount);
  } else if(taskTitle == TASK_TYPES.LAYOUT_CHECK && pageCount > 0) {
    subjectLine = updateSubjectLine(taskTitle, pageCount);
  } else {
    subjectLine = ERRORS.UNKNOWN_ERR;
    showAlert_(ERRORS.UNKNOWN_ERR);
  }
  SHEET.getRange(SUBJECT.cell).setValue(subjectLine);
}

function updateSubjectLine(taskTitle, characterCount) {
  const clientCell = 'F3';
  const fileAndSectionNamesCell = 'F4';
  const dueDateCell = 'F6';
  const dueHeaderCell = 'E6'; // '〆切'
  const taskFooter = "依頼";
  const [assignTitle, dueHeader, dueDate] = SHEET.getRangeList([fileAndSectionNamesCell, dueHeaderCell, dueDateCell]).getRanges().map(range => range.getValues().flat())
  const clientName = checkForSama(clientCell);
  const formattedDate = formatDate_(dueDate);
  const counterType = checkCounterType(taskTitle);

  const count = characterCount > 0 ? `${characterCount}${counterType}` : "";
  return `【${clientName}】 ${assignTitle} ${taskTitle}${taskFooter} ${count} ${dueHeader} ${formattedDate}`;
}

function checkCounterType(taskTitle) {
  const characterCounter = "字";
  const pageCounter = "ページ数";
  let counterType = "";

  switch(taskTitle) {
    case "翻訳":
    case "追つかせ":
      counterType = characterCounter;
      break;
    case "レイアウトチェック":
      counterType = pageCounter;
      break;
    default:
      counterType = "依頼エラー：不明";
  }

  return counterType
}

function getTaskTitle() {
  const taskValues = SHEET.getRange('A3:B5').getValues();
  let taskTitle = "";
  for(let i = 0; i < taskValues.length; i++) {
    console.log(taskValues[i])
    if(taskValues[i][1] == true){
      taskTitle = taskValues[i][0]
    }
  }
  return `${taskTitle}`
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