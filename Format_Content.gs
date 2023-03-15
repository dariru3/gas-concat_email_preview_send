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
    headersContent += `\n\n${content}`;
  } else if(projectType.has(header)){
    content += characterCounter;
    headersContent += `\n${header} ${content}`;
  } else if(header instanceof Date){
    header = formatDate_(header);
    headersContent += `\n${header} ${content}`;
  } else {
    headersContent += `\n\n${header}\n${content}`;
  }
  console.log("Headers and content:", headersContent);
  return headersContent
}

/**
 * Helper function to formate date value
 * with Japanese text.
 * @param {date} date Date value. 
 * @returns Date formatted with Japanese day.
 */
function formatDate_(date) {
  const d = new Date(date);
  const month = d.getMonth() + 1;
  const day = d.getDate();
  const dayShort = new Intl.DateTimeFormat("ja-JP", { weekday: "narrow" }).format(d);
  return `${month}/${day} (${dayShort})`
}