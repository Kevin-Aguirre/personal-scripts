const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId()
const rules = sheet.getConditionalFormatRules();
sheet.setConditionalFormatRules(rules)

function addApplicationRow() {

  var ui = SpreadsheetApp.getUi();
  
  let defaultValues = []
  try {
    defaultValues = [
      getField(ui, promptPosition),
      getField(ui, promptEmployer),
      getField(ui, promptRelevantLink),
      getField(ui, promptApplicationStatus),
      getField(ui, promptActionRequired),
      getField(ui, promptInterviewsCompleted),
      formatDate(new Date()),
      getField(ui, promptLocation),
      getField(ui, promptPay),
      getField(ui, promptContact),
      getField(ui, promptNotes)
    ]
  } catch (e) {
    return
  }

  if (isSpreadsheetEmpty()) {
    sheet.appendRow(headers)
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    formatHeader(header)
  }
  
  sheet.appendRow(defaultValues);

   // --- Add folder creation logic ---
  const company = defaultValues[headers.indexOf(HEADER_EMPLOYER)];  
  const parentFolder = DriveApp.getFolderById(spreadsheetFolderId);
  let companyFolder; 
  const folders = parentFolder.getFoldersByName(company)
  if (folders.hasNext()) {
    companyFolder = folders.next()
  } else {
    companyFolder = parentFolder.createFolder(company)
  }

  const location = defaultValues[headers.indexOf(HEADER_LOCATION)]
  const position = defaultValues[headers.indexOf(HEADER_POSITION)]; 
  const folderName = `${position} - ${location}`;

  const newFolder = companyFolder.createFolder(folderName);

  // Copy resume template files into the new folder
  const templateFolder = DriveApp.getFolderById(resumeTemplateFolderId);
  const resumeFiles = templateFolder.getFilesByName("resume-master");

  if (resumeFiles.hasNext()) {
    resumeFiles.next().makeCopy("resume-master", newFolder)
  }

  // --- end ---

  const lastRowWithData = sheet.getLastRow();
  const statusColumnRange = sheet.getRange(2, headers.indexOf(HEADER_STATUS) + 1, lastRowWithData - 1, 1); 

  statusColumnRange.clearDataValidations()
  statusColumnRange.setDataValidation(dropdownRule)

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

  sheet.getRange(sheet.getLastRow(), headers.indexOf(HEADER_STATUS) + 1).setDataValidation(dropdownRule)

  var lastRowRange = sheet.getRange(sheet.getLastRow(), 1, 1, sheet.getLastColumn());
  lastRowRange.setBackground(LIGHT_GRAY)
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
}

function formatSheet() {

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastCol); // Adjust this if needed

  // Reset all formatting first
  range.setBackground(LIGHT_GRAY);
  range.setFontColor("black");
  range.setFontSize(10);
  

  for (var row = 2; row <= lastRow; row++) {
    var status = sheet.getRange(row, headers.indexOf(HEADER_STATUS) + 1).getValue();
    var rowRange = sheet.getRange(row, 1, 1, lastCol);
    
    switch (status) {
      case "Pending":
        const dateApplied = sheet.getRange(row, headers.indexOf(HEADER_DATE_APPLIED)+1).getValue()
        const difference = calculateDaysDifference(dateApplied)

        if (difference >= 90) {
          sheet.getRange(row, headers.indexOf(HEADER_STATUS) + 1).setValue(String("Ghosted"))
          sheet.getRange(row, headers.indexOf(HEADER_STATUS) + 1).setValue('=')
          rowRange.setBackground(LIGHT_ORANGE);
        } else {
          rowRange.setBackground(LIGHT_GRAY);
        }

        break;
      case "In Progress":
        rowRange.setBackground(LIGHT_BLUE); 
        break;
      case "Rejected":
        rowRange.setBackground(LIGHT_RED); 
        break;
      case "Offer Received":
        rowRange.setBackground(LIGHT_GREEN);
        break;
      case "Action Required":
        rowRange.setBackground(LIGHT_PURPLE); 
        break;
      case "Ghosted":
        rowRange.setBackground(LIGHT_ORANGE); 
        break;
      default:
        rowRange.setBackground("white");
        break;
    }
  }

  var header = sheet.getRange(1, 1, 1, lastCol)
  formatHeader(header)
  range.setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID)
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Internship Tracker')
      .addItem('Add Application', 'addApplicationRow')
      .addItem('Format Spreadsheet', 'formatSheet')
      .addToUi();
} 

