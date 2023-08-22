/**
 * Helper function to format headers and content.
 * @param {string} header Header text
 * @param {string} content Content text
 * @returns Header and content text formatted.
 */
function formatHeader_Content_(header, content){
  const removeContent = new Set(["-", "ー"]);
  const removeHeader = new Set(['挨拶（任意）']);
  const quantityType = new Set(["新規", "既存", "更新", "ゲラ"]);
  const characterCounter = "字";
  const pageCounter = "ページ数";
  const taskType = getTaskTitle();

  if(removeContent.has(content)){
    content = "";
  }

  if(removeHeader.has(header)){
    return `\n\n${content}`;
  } else if(quantityType.has(header)) {
    if((taskType == TASK_TYPES.layoutCheck && header != "ゲラ") || (taskType != TASK_TYPES.layoutCheck && header == "ゲラ")) {
      return "";
    } 
    content += (header == "ゲラ") ? pageCounter : characterCounter;
    return `\n${header} ${content}`;
  } else if(header instanceof Date){
    return `\n${formatDate_(header)} ${content}`;
  } else {
    return `\n\n${header}\n${content}`; // default: header over content
  }
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