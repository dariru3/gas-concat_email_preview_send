# gas-concat_email_preview

## Description
A Google Sheet is used to created a templated email for a translation request. Next to the data entry cells are preview cells to view and edit (if necessary) the subject line and email body. The goal of this project is to help with consistency and insuring that all necessary information is entered and filled in accurately.

## Features
- Simplifies addressing recipients using a database of email address and preferred names.
- Subject line and content order consistent with Western, low-context expectations.
- Headings and content are dynamicly filled.

## How to use
The file looks for data in the "content" row. If there is data, then it gets added to the body. Optional information should be left blank without deleted or editing the "heading" row.

## Technologies
- Google Apps Script (GAS)
- Google Sheets
- Gmail API