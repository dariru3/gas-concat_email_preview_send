function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Menu')
      .addItem('Email Preview', 'concatEmailBody')
      .addItem('Send Email', 'showEmailAlerts_')
      .addToUi();
}

function showEmailAlerts_(confirm) {
  const ui = SpreadsheetApp.getUi();

  const boxAlert = ui.alert(
     'Is the Box link shared?',
     'Can everyone in this email access the Box folder?',
      ui.ButtonSet.YES_NO);

  if (boxAlert == ui.Button.YES) {
    const emailAlert = ui.alert(
      'Send email?',
      ui.ButtonSet.YES_NO);
    if (emailAlert == ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Sending email...');
      sendEmail_();
    } else {
      ui.alert('Email cancelled.');
      console.warn("Email cancellled");
    }
  } else {
    ui.alert('Please check Box folder settings.');
    console.warn("Cancelled at Box alert.");
  }
}