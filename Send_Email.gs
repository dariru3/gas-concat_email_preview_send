/**
 * Function to send email.
 */
function sendEmail_(){
  const subject = SHEET.getRange("G2").getValue();
  const body = concatEmailBody();
  const toAddresses = getEmailAddress_().toAddresses.join();
  const ccAddresses = getEmailAddress_().ccAddresses.join();
  console.log("Send emails to:", toAddresses, ccAddresses);
  const options = {
    cc: ccAddresses
  }

  try {
    GmailApp.sendEmail(toAddresses, subject, body, options)
    console.log("Success: email sent");
    SpreadsheetApp.getActiveSpreadsheet().toast('メールが送信されました');
  }
  catch(e){
    throw "Email error:", e
  }
}