/**
 * Helper function to compile default opening and closing greetings.
 * @returns Opening and closing greetings.
 */
function getDefaultGreeting_(){
  const myName = getNameFromEmailAddress_(MY_EMAIL)
  let opening = concatNames_().toNames;
  let ccNames = concatNames_().ccNames;
  if(ccNames){
    ccNames = ccNames.slice(0,-1); // remove the final comma
    opening += `\n(${ccNames})`;
  }
  opening += `\n\nお疲れ様です。${myName}です。`;

  const closing = `何卒よろしくお願いいたします。\n\n${myName}`;

  return { openingGreeting: opening, closingGreeting: closing }
}