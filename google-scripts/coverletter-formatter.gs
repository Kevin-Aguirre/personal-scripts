

function getDate() {
  var date = new Date()
  var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  var day = date.getDate();
  var month = months[date.getMonth()];  
  var year = date.getFullYear();  

  var formattedDate = month + " " + day + ", " + year;
  return formattedDate
}

function getField(ui, prompt, defaultValue) {
  var fieldPrompt = ui.prompt(prompt, ui.ButtonSet.OK_CANCEL);
  if (fieldPrompt.getSelectedButton() != ui.Button.OK) return; 
  var fieldName = fieldPrompt.getResponseText();
  if (!fieldName) return defaultValue;
  return fieldName
}

function change(body, oldText, newText) {
  body.replaceText(oldText, newText)
}

function formatTemplate() {
  var ui = DocumentApp.getUi();
  var doc = DocumentApp.getActiveDocument()
  var body = doc.getBody()

  var date = getDate()
  var companyName = getField(ui, "Enter Company Name", "[COMPANY NAME]")
  var addressLine1 = getField(ui, "Enter Company Address Line 1", "[COMPANY ADDRESS LINE 1]")
  var addressLine2 = getField(ui, "Enter Company Address Line 2", "[COMPANY ADDRESS LINE 2]")
  var position = getField(ui, "Enter Position", "[POSITION]")

  change(body, "\\[DATE\\]", date)
  change(body, "\\[COMPANY NAME\\]", companyName)
  change(body, "\\[COMPANY ADDRESS LINE 1\\]", addressLine1)
  change(body, "\\[COMPANY ADDRESS LINE 2\\]", addressLine2)
  change(body, "\\[POSITION\\]", position)
}

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('Cover Letter Formatter')
      .addItem('Format Template', 'formatTemplate')
      .addToUi();
}
