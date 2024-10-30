function formatShortcuts(sc) {
  if (Object.keys(sc).length === 0) return ""; // Corrected empty object check
  res = ""
  for (s in sc) {
    res += `\n   ${s} - ${sc[s]}`
  } 
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()); // Adjust this if needed
  

  const positionShortcuts = {
    "1" : "Software Development Intern",
    "2" : "Data Science Intern",
    "3" : `"Technology" Intern`,
    "4" : "Quantitative Intern"
  }

  const locationShortcuts = {
    "1" : "New York, NY"
  }

  let defaultValues = []
  try {
    defaultValues = [
      getField(ui, "Enter Position Name: " + formatShortcuts(positionShortcuts), "null", positionShortcuts),
      getField(ui, "Enter Employer Name", "null"),
      quick ? "null" : getField(ui, "Enter Relevant Link", "null"),
      quick ? "Pending" : getField(ui, "Enter ApplicationStatus", "null"),
      quick ? "null" : getField(ui, "Enter Action Required", "null"),
      quick ? "Applied" : getField(ui, "Enter Interviews Completed", "null"),
      formatDate(new Date()),
      getField(ui, "Enter Location", "null", locationShortcuts),
      quick ? "null" : getField(ui, "Enter Pay Rate", "null"),
      quick ? "null" : getField(ui, "Enter Contact", "null"),
      quick ? "null" : getField(ui, "Enter Notes", "null")
    ]
  } catch (e) { // cancel invoked 
    return
  }
  /*
  positionName = getField(ui, "Enter Position Name: " + formatShortcuts(positionShortcuts), "null", positionShortcuts);
  employerName = getField(ui, "Enter Employer Name", "null");
  link = quick ? "null" : getField(ui, "Enter Relevant Link", "null");
  status = quick ? "Pending" : getField(ui, "Enter ApplicationStatus", "null");
  actionReq = quick ? "null" : getField(ui, "Enter Action Required", "null")
  timeline = quick ? "Applied" : getField(ui, "Enter Interviews Completed", "null");
  date = formatDate(new Date());
  location = getField(ui, "Enter Location", "null", locationShortcuts);
  payRate = quick ? "null" : getField(ui, "Enter Pay Rate", "null");
  contact = quick ? "null" : getField(ui, "Enter Contact", "null");
  notes = quick ? "null" : getField(ui, "Enter Notes", "null");
  */

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.appendRow(defaultValues);

  var lastRowRange = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  lastRowRange.setBackground("#ebebeb")
  range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}

function formatSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  var range = sheet.getRange(1, 1, lastRow, lastCol); // Adjust this if needed
  // Reset all formatting first
  range.setBackground("#ebebeb");
  range.setFontColor("black");
  range.setFontSize(10);
  
  // Apply specific formatting to certain columns (for example, status)
  var statusCol = 4; // Adjust this to your "Status" column number
  for (var i = 2; i <= lastRow; i++) {
    var status = sheet.getRange(i, statusCol).getValue();
    var rowRange = sheet.getRange(i, 1, 1, lastCol);
    
    switch (status) {
      case "Pending":
        rowRange.setBackground("#ebebeb"); // Light gray
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

