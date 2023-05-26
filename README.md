# gas-concat_email_preview
## Project Overview
gas-concat_email_preview is a Google Apps Script (GAS) project designed to streamline the process of creating and sending templated emails through Google Sheets. It simplifies the task of addressing recipients by using a database of email addresses and preferred names, thereby ensuring consistency in communication and reducing the potential for errors. This tool previews the email body and subject line next to the data entry cells, allowing for easy review and editing if necessary.

## Detailed Description
The project relies on three key technologies: Google Apps Script, Google Sheets, and the Gmail API. Google Sheets is used as the interface where users can enter email details and preview the content. Google Apps Script powers the functionality, and the Gmail API is used to send the emails. This script dynamically fills the subject line and content based on the data entered in the Google Sheet.

## Prerequisites
To use gas-concat_email_preview, you need:
- A Google account with access to Google Sheets and Gmail
- Permission to run scripts on your Google account

## Installation Instructions
1. Create a new Google Sheet.
2. Click on Extensions > Apps Script.
3. Delete any code in the script editor and replace it with the code from this repository.
4. Save and close the script editor.

## How to Use
1. Enter data in the "content" row of your Google Sheet.
2. If there is data, it will be automatically added to the email body.
3. Optional information should be left blank without deleting or editing the "heading" row.

## License
This project is licensed under the MIT License.
