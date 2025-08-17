# Job Application Automation Script

## Overview
Keeping track of job applications can be time consuming yet it is necessary to assess progress or changes that need to be made. Creating document copies or adding data to a spreadsheet may take a minute or two but this adds up very quickly over many applications. This Google Apps Script automates the management of job applications in Google Sheets and Google Drive. When a new application is added via a button in the spreadsheet:

1. A new row is inserted with structured job application details.
2. A folder is automatically created in Google Drive for the application.
3. A copy of a predefined resume template (`resume-master`) is placed in the folder (in case you want to tailor your resume before applying).

---

## Features
- Adds a formatted row to a Google Sheet for each new job application.
- Creates a hierarchical folder structure in Google Drive:
- Copies the resume template into the appropriate folder.
- Supports multiple templates and easy configuration via folder IDs.
- Applies conditional formatting and dropdown validation for application status tracking.

Example Directory Output: 
```bash
.
├── Application SpreadSheet 
├── Google/
│   ├── SWE - NY /
│   │   └── resume-master
│   └── IT Specialist - SF /
│       └── resume-master
├── Meta /
│   └── ML Engineer - NY/
│       └── resume-master
└── Nvidia/
    └── Systems Engineer - Seattle/
        └── resume-master
```

## Setup
1. Open your Google Spreadsheet (note that your spreadsheet must be open for the script to work).
2. Go to `Extensions → Apps Script`.
3. Create three files `Main.gs`, `Constants.gs`, and `Utils.gs` (you could also paste everything in one file since imports/exports are not necessary in Apps Script)
4. Update the following constants in the script:

```javascript
const resumeTemplateFolderID = 'YOUR_TEMPLATE_FOLDER_ID';
const spreadsheetFolderId = 'YOUR_APPLICATIONS_ROOT_FOLDER_ID';
```

5. Go back to your spreadsheet, and you should see a new tab on the header 

## Considerations
For the most part, most changes to this script will be made to code in `Constants.gs`. Folder ids, colors, dropdown options, prompt names, and default values (values that are used instead of prompting the user) are stored here. For instance, these default values 
```
const promptToDefault = {
  [promptApplicationStatus]: "Pending",
  [promptInterviewsCompleted]: "0",
}
```
...make it so that the user isnt prompted the application status or how many interviews have been compeleted, instead it is assumed that the status is 'Pending' and 0 interviews have been completed, a reasonable assumption for new applications. 
## TODO 
* Make new row creation more dynamic (static number of headers, will break if any headers are added or removed)
* Add support for different tailored resumes for different role types
* Add some meaningful charts (pie chart showing ghosted / in progress / rejected / offer / etc. )
* Add a link to new folder created in each row, easy access to newly created foler after adding application 

## Contributing 
If you have any suggestions for this script, feel free to reach out. 