// global variables
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();

const SUBJECT = {
    cell: "I2",
    value: function() {
      return SHEET.getRange(this.cell).getValue();
    }
  };

const TASK_TYPES = {
  TRANSLATE: "翻訳",
  ADDITION: "追つかせ",
  LAYOUT_CHECK: "レイアウトチェック"
}
// end of global variables