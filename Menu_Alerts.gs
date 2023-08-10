function onOpen() {
  UI.createMenu('GASメール')
      .addItem('メールプレビュー', 'concatEmailBody')
      .addItem('メール送信', 'emailAlertsHandler_')
      .addItem('Gmail 下書きへ', 'draftsAlertsHandler_')
      .addToUi();
}

class AlertOptions {
  constructor(type, message, status, emailParam) {
    this.type = type;
    this.message = message;
    this.status = status;
    this.emailParam = emailParam
  }
}

function emailAlertsHandler_() {
  const options = new AlertOptions(
    "email", 
    "メールを送信して良いですか？",
    'sending',
    'immediate'
  );
  showAlert(options);
}

function draftsAlertsHandler_() {
  const options = new AlertOptions(
    "email",
    "メール原稿の作成して良いですか？",
    "drafting",
    "draft"
  );
  showAlert(options);
}

function showAlert(options) {
  const UI = SpreadsheetApp.getUi();
  if(options.type == "email") {
    const boxAlert = UI.alert(
      '相手が開けるBOXリンクですか',
      'プロジェクトフォルダのリンクになっていませんか？\n共有フォルダの権限は「リンクを知っている全員」に設定されていますか？',
       UI.ButtonSet.YES_NO);
 
   if (boxAlert == UI.Button.YES) {
     const emailAlert = UI.alert(options.message, UI.ButtonSet.YES_NO);
     if (emailAlert == UI.Button.YES) {
       emailStatusToast_(options.status);
       prepareEmail(options.emailParam);
     } else { // emailAlert == NO
       UI.alert('中止しました');
     }
   } else { // boxAlert == NO
     UI.alert('BoxフォルダのURL/アクセス権限を変更してください');
   }
  } else if(options.type == "undefined") {
    UI.alert(
      'あなたの名前が見つかりませんでした。',
      '「メールリスト」のシートの、A～C列に入力してください',
      UI.ButtonSet.OK
    );
  } else if(options.type == "message") {
    UI.alert(
      options.message,
      UI.ButtonSet.OK
    );
  }
}

/**
 * Helper function to send out toast message on email status
 * @param {string} status "Sending" or "Sent"
 * @returns Triggers toast message on bottom right corner
 */
function emailStatusToast_(status) {
  let message = "";

  switch (status) {
    case "sending":
      message = "メール送信中...";
      break;
    case "sent":
      message = "メールが送信されました";
      break;
    case "cancel":
      message = "メール配信を中止しました";
      break;
    case "drafting":
      message = "下書きメールの作成中";
      break;
    case "draft":
      message = "下書きメールが作成されました。下書きボックスを確認してください。";
      break;
    default:
      console.error("Toast message error!")
  }

  return SpreadsheetApp.getActiveSpreadsheet().toast(message);
}