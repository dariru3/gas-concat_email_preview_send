// global variables
// connect to sheet
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();
const EMAIL_ADDRESSES = SHEET.getRange("C3:D11").getValues();
const START_ROW = 6; // spreadsheet row 7
const HEADER_COL = 4; // column C
const CONTENT_COL = 5; // column D
const CHAR_COUNT_VALUES = SHEET.getRange('F9:F10').getValues();
const PAGE_COUNT_VALUE = SHEET.getRange('F12').getValue();


// subject line
const SUBJECT_CELL = "I2"
const SUBJECT_VALUE = SHEET.getRange(SUBJECT_CELL).getValue();
const CLIENT_CELL = 'F3';
const ASSIGN_CELL = 'F4';
const DUE_HEADER_CELL = 'E6';
const DUE_DATE_CELL = 'F6';

// email body
const BODY_CELL = "I3"
const BODY_VALUE = SHEET.getRange(BODY_CELL).getValue();

// tasks
const TASK_FOOTER = SHEET.getRange('A2').getValue();
const TASK_VALUES = SHEET.getRange('A3:B5').getValues();
const CHECKBOX_RANGE = SHEET.getRange("B3:B5");
// end of global variables