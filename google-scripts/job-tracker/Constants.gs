const resumeTemplateFolderId = # folder id of resume master here 
const spreadsheetFolderId = # folder id of spreadsheet here 

const HEADER_POSITION = "Position"
const HEADER_EMPLOYER = "Employer"
const HEADER_LINKS = "Relevant Links"
const HEADER_STATUS = "Status"
const HEADER_ACTION_REQ = "Action Required"
const HEADER_INTERVIEWS = "Interviews"
const HEADER_DATE_APPLIED = "Date Applied"
const HEADER_LOCATION = "Location"
const HEADER_PAYRATE = "Pay Rate"
const HEADER_CONTACT = "Contact"
const HEADER_NOTES = "Notes"

const headers = [
  HEADER_POSITION,
  HEADER_EMPLOYER,
  HEADER_LINKS,
  HEADER_STATUS,
  HEADER_ACTION_REQ,
  HEADER_INTERVIEWS,
  HEADER_DATE_APPLIED,
  HEADER_LOCATION,
  HEADER_PAYRATE,
  HEADER_CONTACT,
  HEADER_NOTES,
]

DARK_BLUE = "#313b6b"
LIGHT_GRAY = "#ebebeb"
LIGHT_BLUE = "#cce5ff"
LIGHT_RED = "#f8d7da"
LIGHT_GREEN = "#d4edda"
LIGHT_PURPLE = "#f2c2f1"
LIGHT_YELLOW = "#ffe5b4"

const dropdownOptions = [
  "Pending", 
  "In Progress", 
  "Rejected", 
  "Offer Received", 
  "Action Required",
  "Ghosted"
]

const dropdownColors = [
  LIGHT_GRAY,
  LIGHT_BLUE,
  LIGHT_RED,
  LIGHT_GREEN,
  LIGHT_PURPLE,
  LIGHT_YELLOW 
]

const dropdownRule = SpreadsheetApp.newDataValidation()
  .requireValueInList(
    dropdownOptions, 
    true
  )
  .setAllowInvalid(false)
  .build()

const promptPosition = "Enter Position Name"
const promptEmployer = "Enter Employer Name"
const promptRelevantLink = "Enter Relevant Link"
const promptApplicationStatus = "Enter Application Status"
const promptActionRequired = "Enter Action Required"
const promptInterviewsCompleted = "Enter Interviews Completed"
const promptLocation = "Enter Location"
const promptPay = "Enter Pay Rate"
const promptContact = "Enter Contact"
const promptNotes = "Enter Notes"

const promptToDefault = {
  [promptRelevantLink]: "",
  [promptApplicationStatus]: "Pending",
  [promptActionRequired]: "false",
  [promptInterviewsCompleted]: "0",
  [promptPay]: "",
  [promptContact]: "",
  [promptNotes]: ""
}
