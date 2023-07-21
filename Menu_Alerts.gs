// global ui variable
const UI = SpreadsheetApp.getUi();
// end of global ui variable

function onOpen() {
  UI.createMenu('GASメール')
      .addItem('メールプレビュー', 'concatEmailBody')
      .addItem('メール送信', 'emailAlertsHandler_')
      .addItem('Create email draft', 'draftsAlertsHandler_')
      .addToUi();
}

function emailAlertsHandler_() {
  showEmailAlerts_('メールを送信して良いですか？', 'sending', 'immediate');
}

function draftsAlertsHandler_() {
  showEmailAlerts_('Create email draft?', 'drafting', 'draft');
}

function showEmailAlerts_(emailAlertMessage, emailStatusMessage, emailParam) {
  const boxAlert = UI.alert(
     '相手が開けるBOXリンクですか',
     'プロジェクトフォルダのリンクになっていませんか？\n共有フォルダの権限は「リンクを知っている全員」に設定されていますか？',
      UI.ButtonSet.YES_NO);

  if (boxAlert == UI.Button.YES) {
    const emailAlert = UI.alert(emailAlertMessage, UI.ButtonSet.YES_NO);
    if (emailAlert == UI.Button.YES) {
      emailStatusToast_(emailStatusMessage);
      sendEmail(emailParam);
    } else { // emailAlert == NO
      UI.alert('中止しました');
      console.warn("Email action cancelled");
    }
  } else { // boxAlert == NO
    UI.alert('BoxフォルダのURL/アクセス権限を変更してください');
    console.warn("Cancelled at Box alert");
  }
}

function undefinedNameAlert_() {
  UI.alert(
    'あなたの名前が見つかりませんでした。',
    '「メールリスト」のシートの、A～C列に入力してください',
    UI.ButtonSet.OK
  );
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
      message = "Creating email draft";
      break;
    case "draft":
      message = "Email draft created. Check draft box.";
      break;
    default:
      console.error("Toast message error!")
  }

  return SpreadsheetApp.getActiveSpreadsheet().toast(message);
}