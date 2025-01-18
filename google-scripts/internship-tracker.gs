const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

const positionCol = 1;
const employerCol = 2;
const linksCol = 3;
const statusCol = 4; 
const actionReqCol = 5;
const timelineCol = 6;
const dateCol = 7;
const locationCol = 8;
const payRateCol = 9;
const contactCol = 10;
const notesCol = 11;

const dropdownRule = SpreadsheetApp.newDataValidation()
  .requireValueInList(
    ["Pending", "In Progress", "Rejected", "Offer Received", "Action Required", "Ghosted"], 
    true
  )
  .setAllowInvalid(false)
  .build()

const statusColumnRange = sheet.getRange(`D2:D${sheet.getLastRow()}`)
statusColumnRange.clearDataValidations()
statusColumnRange.setDataValidation(dropdownRule)

const dropdownOptions = ["Pending", "In Progress", "Rejected", "Offer Received", "Action Required", "Ghosted"]
const dropdownColors = ["#ebebeb", "#cce5ff", "#f8d7da","#d4edda", "#f2c2f1", "#ffe5b4"]

const rules = sheet.getConditionalFormatRules();

for (var i = 0; i < dropdownOptions.length; ++i) {
  const currRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(dropdownOptions[i])
    .setBackground(dropdownColors[i])
    .setRanges([statusColumnRange])
    .build()
  
  rules.push(currRule)
}

sheet.setConditionalFormatRules(rules)

const positionShortcuts = {
  "1" : "Software Development Intern",
  "2" : "Data Science Intern",
  "3" : "IT Intern",
  "4" : "Quantitative Intern"
}

const locationShortcuts = {
  "1" : "New York, NY"
}

function calculateDaysDifference(dateString) {
  const today = new Date()
  return (today.getTime() - dateString.getTime()) / (1000 * 60 * 60 * 24)
}

function formatShortcuts(sc) {
  res = ""
  for (s in sc) res += `\n   ${s} - ${sc[s]}`
  return res 
}

function formatDate(date) {
  var day = String(date.getDate()).padStart(2, '0'); 
  var month = String(date.getMonth() + 1).padStart(2, '0'); 
  var year = date.getFullYear(); 
  return day + '/' + month + '/' + year; 
}

function getField(ui, prompt, defaultValue, shortcuts = {}) {
  var fieldPrompt = ui.prompt(prompt, ui.ButtonSet.OK_CANCEL);
  if (fieldPrompt.getSelectedButton() != ui.Button.OK) throw new Error("Cancel Invoked"); 
  var fieldName = fieldPrompt.getResponseText();
  if (shortcuts[fieldName]) return shortcuts[fieldName]
  if (!fieldName) return defaultValue;
  return fieldName
}

function addApplicationRowQuick() {addApplicationRow(true)}
function addApplicationRowDetailed() {addApplicationRow(false)}
function addApplicationRow(quick) {
  var ui = SpreadsheetApp.getUi();
  
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()); // Adjust this if needed

  let defaultValues = []
  try {
    defaultValues = [
      getField(ui, "Enter Position Name: " + formatShortcuts(positionShortcuts), "null", positionShortcuts),
      getField(ui, "Enter Employer Name", "null"),
      quick ? "null" : getField(ui, "Enter Relevant Link", "null"),
      quick ? "Pending" : getField(ui, "Enter Application Status", "null"),
      quick ? "null" : getField(ui, "Enter Action Required", "null"),
      quick ? 
        "Pending" 
        : 
        getField(ui, "Enter Interviews Completed", "null"),
      formatDate(new Date()),
      getField(ui, "Enter Location", "null", locationShortcuts),
      quick ? "null" : getField(ui, "Enter Pay Rate", "null"),
      quick ? "null" : getField(ui, "Enter Contact", "null"),
      quick ? "null" : getField(ui, "Enter Notes", "null")
    ]
  } catch (e) { // cancel invoked 
    return
  }

  sheet.appendRow(defaultValues);
  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Pending", "In Progress", "Rejected", "Offer Received", "Action Required", "Ghosted"], 
      true
    )
    .setAllowInvalid(false)
    .build()

  const statusColumnRange = sheet.getRange(`D2:D${sheet.getLastRow()}`)
  statusColumnRange.clearDataValidations()
  statusColumnRange.setDataValidation(dropdownRule)

  const dropdownOptions = ["Pending", "In Progress", "Rejected", "Offer Received", "Action Required", "Ghosted"]
  const dropdownColors = ["#ebebeb", "#cce5ff", "#f8d7da","#d4edda", "#f2c2f1", "#ffe5b4"]

  const rules = sheet.getConditionalFormatRules();

  for (var i = 0; i < dropdownOptions.length; ++i) {
    const currRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(dropdownOptions[i])
      .setBackground(dropdownColors[i])
      .setRanges([statusColumnRange])
      .build()
    
    rules.push(currRule)
  }

  sheet.setConditionalFormatRules(rules)

  // const foo = sheet.getRange(`D2:D${sheet.getLastRow()}`)
  // foo.clearDataValidations()
  // foo.setDataValidation(dropdownRule)

  sheet.getRange(sheet.getLastRow(), statusCol).setDataValidation(dropdownRule)

  var lastRowRange = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  lastRowRange.setBackground("#ebebeb")
  range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}

function formatSheet() {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastCol); // Adjust this if needed

  // Reset all formatting first
  range.setBackground("#ebebeb");
  range.setFontColor("black");
  range.setFontSize(10);
  

  for (var row = 2; row <= lastRow; row++) {
    var status = sheet.getRange(row, statusCol).getValue();
    var rowRange = sheet.getRange(row, 1, 1, lastCol);
    
    switch (status) {
      case "Pending":
        const dateApplied = sheet.getRange(row, dateCol).getValue()
        
        const difference = calculateDaysDifference(dateApplied)

        if (difference >= 90) {
          sheet.getRange(row, statusCol).setValue(String("Ghosted"))
          sheet.getRange(row, statusCol).setValue('=')
          rowRange.setBackground("#ffe5b4"); // Light orange
        } else {
          rowRange.setBackground("#ebebeb"); // Light gray
        }

        break;
      case "In Progress":
        rowRange.setBackground("#cce5ff"); // Light blue
        break;
      case "Rejected":
        rowRange.setBackground("#f8d7da"); // Light red
        break;
      case "Offer Received":
        rowRange.setBackground("#d4edda"); // Light green
        break;
      case "Action Required":
        rowRange.setBackground("#f2c2f1"); // Light purple
        break;
      case "Ghosted":
        rowRange.setBackground("#ffe5b4"); // Light orange
        break;
      default:
        rowRange.setBackground("white"); // Default background
        break;
    }
  }

  var header = sheet.getRange(1, 1, 1, lastCol)
  header.setBackground("#313b6b")
  header.setFontWeight("bold")
  header.setFontColor("white")
  // You can add other formatting rules here, like font sizes, bold text, borders, etc.

  range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID)
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Internship Tracker')
      .addItem('Add Application (Quick)', 'addApplicationRowQuick')
      .addItem('Add Application (Detailed)', 'addApplicationRowDetailed')
      .addItem('Format Spreadsheet', 'formatSheet')
      .addToUi();
} 

