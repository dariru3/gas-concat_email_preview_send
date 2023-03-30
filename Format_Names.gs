/**
 * Helper function to get preferred names from email address.
 * @param {any} getMyName Optional: add to get user's name.
 * @returns Either user's name or to and cc names.
 */
function concatNames_() {
  const toAddresses = getEmailAddress_().toAddresses;
  const ccAddresses = getEmailAddress_().ccAddresses;
  const toNames = addSanToNames_(toAddresses);
  const ccNames = addSanToNames_(ccAddresses);
  
  console.log("All names:", toNames, ccNames);
  return { toNames: toNames,ccNames: ccNames }
}

/**
 * Helper function to add -さん after every name.
 * @param {list} list List of names. 
 * @returns String of names.
 */
function addSanToNames_(list){
  let concatString = ""
  for(let i=0; i<list.length; i++){
    const preferredName = getNameFromEmailAddress_(list[i])
    if(preferredName){
      concatString += `${preferredName}さん、`;
    }
  }
  return concatString
}