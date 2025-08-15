// given a date string, calculates the difference in days 
function calculateDaysDifference(dateString) {
  const today = new Date()
  const [month, day, year] = dateString.split("/").map(Number);
  const parsedDate = new Date(year, month - 1, day);
  return (today.getTime() - parsedDate.getTime()) / (1000 * 60 * 60 * 24)
}

// formats a date to display in spreadsheet
function formatDate(date) {
  var day = String(date.getDate()).padStart(2, '0'); 
  var month = String(date.getMonth() + 1).padStart(2, '0'); 
  var year = date.getFullYear(); 
  return month + '/' + day + '/' + year; 
}

// handles retreiving value from input or default fields 
function getField(ui, prompt) {
  if (prompt in promptToDefault) return promptToDefault[prompt];
  
  var fieldPrompt = ui.prompt(prompt, ui.ButtonSet.OK_CANCEL);
  if (fieldPrompt.getSelectedButton() != ui.Button.OK) throw new Error("Cancel Invoked"); 
  var fieldName = fieldPrompt.getResponseText();
  if (!fieldName) return "";
  return fieldName
}

// checks if a spreadsheet is completely empty (no styling + values)
function isSpreadsheetEmpty() {
  const values = sheet.getDataRange().getValues();
  const isEmpty = values.flat().every(cell => cell === "");
  return isEmpty;
}

// adds border, color, and background styles for header row 
function formatHeader(headerRange) {
  headerRange.setBackground(DARK_BLUE)
  headerRange.setFontWeight("bold")
  headerRange.setFontColor("white")
}