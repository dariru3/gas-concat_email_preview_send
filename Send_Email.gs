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
    emailStatusToast_("sent");
  }

  if(emailParam == "draft") {
    GmailApp.createDraft(toAddresses, SUBJECT.value(), BODY.value(), options)
    emailStatusToast_("draft");
  }
}

function checkMyNameExists_(name){
  let myNameExists = true
  const myPreferredName = getNameFromEmailAddress_(MY_EMAIL);
  if(!name || !myPreferredName){ // checks if `name` and `myPreferredName` are not (!) "undefined" or blank
    showAlert({ type: "undefined" });
    emailStatusToast_("cancel")
    myNameExists = false
  }
  return myNameExists
}