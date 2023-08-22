// global variables
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();

const BODY = {
  cell: "I3",
  value: function() {
      return SHEET.getRange(this.cell).getValue();
  }
}

const SUBJECT = {
    cell: "I2",
    value: function() {
      return SHEET.getRange(this.cell).getValue();
    }
  };

const TASK_TYPES = {
  translate: "翻訳",
  addition: "追つかせ",
  layoutCheck: "レイアウトチェック"
}
// end of global variables