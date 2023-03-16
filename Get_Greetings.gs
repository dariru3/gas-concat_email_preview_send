/**
 * Helper function to compile default opening and closing greetings.
 * @returns Opening and closing greetings.
 */
function getDefaultGreeting_(){
  const myName = concatNames_("Daryl");
  let opening = concatNames_().toNames;
  let ccNames = concatNames_().ccNames;
  if(ccNames){
    ccNames = ccNames.slice(0,-1); // remove the final comma
    opening += `\n(${ccNames})`;
  }
  opening += `\n\nお疲れ様です。${myName}です。`;

  const closing = `何卒よろしくお願いいたします。\n\n${myName}`;
  console.log("Greetings:", opening, closing);
  return { openingGreeting: opening, closingGreeting: closing }
}