/**
 * Function to send email.
 */
function prepareEmail(emailParam){
  const emailNameCol = 'B'
  const myName = getNameFromEmailAddress_(MY_EMAIL, emailNameCol);

  if(checkMyNameExists_(myName) == false){
    return;
  }

  const {toAddresses, ccAddresses} = getEmailAddress_();
  // console.log("Send emails to:", toAddresses, ccAddresses);
  const options = {
    cc: ccAddresses.join(),
    name: myName
  }

  try {
    sendEmail_(emailParam, toAddresses, options);
  } catch (e) {
    console.error("Email error:", e);
  }
}

function sendEmail_(emailParam, toAddresses, options) {
  if(emailParam == "immediate") {
    GmailApp.sendEmail(toAddresses, SUBJECT.value(), BODY.value(), options)
    // console.log("Success: email sent");
    emailStatusToast_("sent");
  }

  if(emailParam == "draft") {
    GmailApp.createDraft(toAddresses, SUBJECT.value(), BODY.value(), options)
    // console.log("Success: draft created");
    emailStatusToast_("draft");
  }
}

function checkMyNameExists_(name){
  let myNameExists = true
  const myPreferredName = getNameFromEmailAddress_(MY_EMAIL);
  if(!name || !myPreferredName){ // checks if `name` and `myPreferredName` are not (!) "undefined" or blank
    const undefinedAlertOptions = new AlertOptions("undefined");
    showAlert(undefinedAlertOptions);
    // undefinedNameAlert_()
    emailStatusToast_("cancel")
    console.warn("Cancelled at myName alert");
    myNameExists = false
  }
  return myNameExists
}