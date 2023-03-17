/**
 * Helper function to format headers and content.
 * @param {string} header Header text
 * @param {string} content Content text
 * @returns Header and content text formatted.
 */
function formatHeaderContent_(header, content){
  const removeContent = new Set(["-", "ー"]);
  const removeHeader = new Set(['挨拶（任意）']);
  const projectType = new Set(["新規", "既存", "更新"]);
  const characterCounter = "字";

  let headersContent = "";
  if(removeContent.has(content)){
    content = "";
  }
  // format headers according to type
  if(removeHeader.has(header)){
    headersContent += `\n\n${content}`; // add content without header
  } else if(projectType.has(header)){
    content += characterCounter;
    headersContent += `\n${header} ${content}`; // put header and content+"字" side-by-side
  } else if(header instanceof Date){
    header = formatDate_(header);
    headersContent += `\n${header} ${content}`; // put header(date) and content side-by-side
  } else {
    headersContent += `\n\n${header}\n${content}`; // add header, underneath add content
  }
  console.log("Headers and content:", headersContent);
  return headersContent
}

/**
 * Helper function to format date value into Japanese text.
 * @param {date} date Date value. 
 * @returns Date formatted.
 */
function formatDate_(date) {
  const d = new Date(date);
  const month = d.getMonth() + 1;
  const day = d.getDate();
  const dayShort = new Intl.DateTimeFormat("ja-JP", { weekday: "narrow" }).format(d);
  return `${month}/${day} (${dayShort})`
}