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
     'このメールの全員がBoxフォルダーにアクセスできるのでしょうか？',
      ui.ButtonSet.YES_NO);

  if (boxAlert == ui.Button.YES) {
    const emailAlert = ui.alert(
      'メールを送信してよろしでしょうか？',
      ui.ButtonSet.YES_NO);
    if (emailAlert == ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('メール送信中...');
      sendEmail_();
    } else {
      ui.alert('メール送信中止');
      console.warn("Email cancellled");
    }
  } else {
    ui.alert('Boxフォルダの設定をご確認ください');
    console.warn("Cancelled at Box alert");
  }
}