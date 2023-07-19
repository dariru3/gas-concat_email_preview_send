// global variables
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();
const START_ROW = 6; // spreadsheet row 7
const HEADER_COL = 4; // column C
const CONTENT_COL = 5; // column D
const PREVIEW_SUBJECT_CELL = "I2"
const PREVIEW_BODY_CELL = "I3"
const EMAIL_ADDRESSES = SHEET.getRange("C3:D11").getValues();
const SUBJECT_CELL = SHEET.getRange("I2").getValue();
const BODY_CELL = SHEET.getRange("I3").getValue();
const TRANSLATE_ASSIGN = SHEET.getRange("B3").getValue();
const ADDITION_ASSIGN = SHEET.getRange("B4").getValue();
const LAYOUT_CHECK_ASSIGN = SHEET.getRange("B5").getValue();
// end of global variables

function test(){
  console.log(TRANSLATE_ASSIGN, ADDITION_ASSIGN, LAYOUT_CHECK_ASSIGN)
}