/**
 * Helper function to get name from an email address.
 * @param {string} address Email address to look up.
 * @returns Preferred name when sending emails.
 */
function getNameFromEmailAddress_(address, lookupColumn = 'C') {
  const referenceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("メールリスト");
  const emailNameList = referenceSheet.getRange("A:C").getValues();
  const noNames = new Set(["edit_all@link-cc.co.jp"]);

  const nameLookup = {}; // dictionary for easier lookup
  emailNameList.forEach(row => {
    switch (lookupColumn) {
      case 'B':
        nameLookup[row[0]] = row[1];
        break;
      case 'C':
      default:
        nameLookup[row[0]] = row[2];
        break;
    }
  });

  if (address != "" && !noNames.has(address) && address in nameLookup) {
    return nameLookup[address];
  } else {
    console.error("Email address not found.");
  }
}

/**
 * Helper function to gather to and cc email addresses from spreadsheet
 * @returns Array with two lists: to addresses and cc addresses
 */
function getEmailAddress_(){
  let toAddresses = [];
  let ccAddresses = [];
  for(let i=0; i<EMAIL_ADDRESSES.length; i++){
    toAddresses.push(EMAIL_ADDRESSES[i][0]);
    ccAddresses.push(EMAIL_ADDRESSES[i][1]);
  }
  toAddresses = toAddresses.filter(a => a);
  ccAddresses = ccAddresses.filter(b => b);

  return { toAddresses: toAddresses, ccAddresses: ccAddresses }
}