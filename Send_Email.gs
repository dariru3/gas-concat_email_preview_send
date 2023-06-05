/**
 * Function to send email.
 */
function sendEmail(){
  const myName = getNameFromEmailAddress_(MY_EMAIL, 'B');
  if(checkMyNameExists_(myName) == false){
    return;
  }
  const subject = SHEET.getRange("G2").getValue();
  const body = SHEET.getRange("G3").getValue();
  const toAddresses = getEmailAddress_().toAddresses.join();
  const ccAddresses = getEmailAddress_().ccAddresses.join();
  console.log("Send emails to:", toAddresses, ccAddresses);
  const options = {
    cc: ccAddresses,
    name: myName
  }

  try {
    GmailApp.sendEmail(toAddresses, subject, body, options)
    console.log("Success: email sent");
    emailStatusToast_("sent");
  }
  catch(e){
    throw "Email error:", e
  }
}

function checkMyNameExists_(name){
  let myNameExists = true
  const myPreferredName = getNameFromEmailAddress_(MY_EMAIL);
  if(!name || !myPreferredName){ // checks if `name` and `myPreferredName` are not (!) "undefined" or blank
    undefinedNameAlert()
    emailStatusToast_("cancel")
    console.warn("Cancellled at myName alert");
    myNameExists = false
  }
  return myNameExists
}