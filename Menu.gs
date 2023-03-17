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
     'Boxリンクは共有されているのでしょうか？',
     '共有フォルダの権限は「リンクを知っている全員」に設定されていますか？',
      ui.ButtonSet.YES_NO);

  if (boxAlert == ui.Button.YES) {
    const emailAlert = ui.alert(
      'メールを送信してよいですか？',
      ui.ButtonSet.YES_NO);
    if (emailAlert == ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('メール送信中...');
      sendEmail_();
    } else {
      ui.alert('メール配信停止中');
      console.warn("Email cancellled");
    }
  } else {
    ui.alert('Boxフォルダのアクセス権限を変更してください');
    console.warn("Cancelled at Box alert");
  }
}