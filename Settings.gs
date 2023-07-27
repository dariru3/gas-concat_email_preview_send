// global variables
// connect to sheet
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();
const EMAIL_ADDRESSES = SHEET.getRange("C3:D7").getValues();
const START_ROW = 6; // spreadsheet row 7
const HEADER_COL = 4; // column C
const CONTENT_COL = 5; // column D

// email body
const BODY_CELL = "I3"
const BODY_VALUE = SHEET.getRange(BODY_CELL).getValue();

// end of global variables