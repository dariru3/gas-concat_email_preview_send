/**
 * Function to send email.
 */
function sendEmail(type){
  const emailNameCol = 'B'
  const myName = getNameFromEmailAddress_(MY_EMAIL, emailNameCol);
  if(checkMyNameExists_(myName) == false){
    return;
  }
  const toAddresses = getEmailAddress_().toAddresses.join();
  const ccAddresses = getEmailAddress_().ccAddresses.join();
  console.log("Send emails to:", toAddresses, ccAddresses);
  const options = {
    cc: ccAddresses,
    name: myName
  }

  if(type == "immediate"){
    try {
    GmailApp.sendEmail(toAddresses, SUBJECT_CELL, BODY_CELL, options)
    console.log("Success: email sent");
    emailStatusToast_("sent");
    }
    catch(e){
      throw "Email error:", e
    }
  }

  if(type == "draft"){
    try {
    GmailApp.createDraft(toAddresses, SUBJECT_CELL, BODY_CELL, options)
    console.log("Success: draft created");
    emailStatusToast_("draft");
    }
    catch(e){
      throw "Email error:", e
    }
  }
  
}

function checkMyNameExists_(name){
  let myNameExists = true
  const myPreferredName = getNameFromEmailAddress_(MY_EMAIL);
  if(!name || !myPreferredName){ // checks if `name` and `myPreferredName` are not (!) "undefined" or blank
    undefinedNameAlert_()
    emailStatusToast_("cancel")
    console.warn("Cancelled at myName alert");
    myNameExists = false
  }
  return myNameExists
}