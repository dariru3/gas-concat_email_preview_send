function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('GASメール')
      .addItem('メールプレビュー', 'concatEmailBody')
      .addItem('メール送信', 'showEmailAlerts_')
      .addToUi();
}

function showEmailAlerts_() {
  const ui = SpreadsheetApp.getUi();

  const boxAlert = ui.alert(
     '相手が開けるBOXリンクですか',
     'プロジェクトフォルダのリンクになっていませんか？\n共有フォルダの権限は「リンクを知っている全員」に設定されていますか？',
      ui.ButtonSet.YES_NO);

  if (boxAlert == ui.Button.YES) {
    const emailAlert = ui.alert(
      'メールを送信してよいですか？',
      ui.ButtonSet.YES_NO);
    if (emailAlert == ui.Button.YES) {
      emailStatusToast_("sending");
      sendEmail_();
    } else {
      ui.alert('メール配信を中止しました');
      console.warn("Email cancellled");
    }
  } else {
    ui.alert('BoxフォルダのURL/アクセス権限を変更してください');
    console.warn("Cancelled at Box alert");
  }
}

function emailStatusToast_(status) {
  let message = "";

  switch (status) {
    case "sending":
      message = "メール送信中...";
      break;
    case "sent":
      message = "メールが送信されました";
      break;
    default:
      console.error("Toast message error!")
  }

  return SpreadsheetApp.getActiveSpreadsheet().toast(message);
}
