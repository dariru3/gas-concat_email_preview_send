/**
 * Function to send email.
 */
function sendEmail(){
  const myName = getNameFromEmailAddress_(MY_EMAIL, 'B');
  console.log("myName:", myName)
  if(checkMyNameExists(myName) == false){
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
    // GmailApp.sendEmail(toAddresses, subject, body, options)
    console.log("Email subject:", subject)
    console.log("Email body:", body)
    console.log("Success: email sent");
    emailStatusToast_("sent");
  }
  catch(e){
    throw "Email error:", e
  }
}

function checkMyNameExists(name){
  let myNameExists = true
  if(name == undefined || name == ""){
    undefinedNameAlert()
    emailStatusToast_("cancel")
    myNameExists = false
  }
  return myNameExists
}