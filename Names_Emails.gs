/**
 * Helper function to get name from an email address.
 * @param {string} address Email address to look up.
 * @returns Preferred name when sending emails.
 */
function getNameFromEmailAddress_(address){
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("メールリスト");
  const emailNameList = referenceSheet.getRange("A:C").getValues();
  const noNames = new Set(["edit_all@link-cc.co.jp"]);

  const nameLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => nameLookup[row[0]] = row[2]);

  if(address != "" && !noNames.has(address) && address in nameLookup){
    console.log("Name:", nameLookup[address]);
    return nameLookup[address]
  } else {
    console.error("Email address not found.");
  }
}

/**
 * Helper function to gather to and cc email addresses from spreadsheet
 * @returns Array with two lists: to addresses and cc addresses
 */
function getEmailAddress_(){
  // connect to sheet
  let emailAddresses = SHEET.getRange("A3:B11").getValues();

  let toAddresses = [];
  let ccAddresses = [];
  for(let i=0; i<emailAddresses.length; i++){
    toAddresses.push(emailAddresses[i][0]);
    ccAddresses.push(emailAddresses[i][1]);
  }
  toAddresses = toAddresses.filter(a => a);
  ccAddresses = ccAddresses.filter(b => b);

  console.log("Get email addresses", toAddresses, ccAddresses);
  return { toAddresses: toAddresses, ccAddresses: ccAddresses }
}