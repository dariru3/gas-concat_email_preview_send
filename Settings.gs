// global variables
// connect to sheet
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート");
const MY_EMAIL = Session.getActiveUser().getEmail();
const EMAIL_ADDRESSES = SHEET.getRange("C3:D7").getValues();
const START_ROW = 6; // spreadsheet row 7
const HEADER_COL = 4; // column C
const CONTENT_COL = 5; // column D

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
    TRANSLATE: "翻訳",
    ADDITION: "追つかせ",
    LAYOUT_CHECK: "レイアウトチェック"
  }

  const ERRORS = {
    HEADER: "依頼エラー：",
    NO_TASK: "B欄のタスクを選択してください",
    NO_COUNT: "字数かページ数を入力してくさい",
    UNKNOWN_ERR: "不明"
  }
// end of global variables