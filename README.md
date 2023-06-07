# gas-concat_email_preview
## Project Overview
gas-concat_email_preview is a Google Apps Script (GAS) project that leverages Google Sheets and the Gmail API to enhance the creation and sending of templated emails. This project ensures the inclusion of essential work details such as specific Box links and other work-related information to promote consistency and reduces errors in email communication. Importantly, this tool is designed to bridge the gap between high context and low context cultures, particularly assisting high context Japanese writers to craft emails that low context English readers can easily understand.

## Intercultural Communication
The goal of this project is not just to streamline the email creation process, but also to facilitate effective intercultural communication. The Google Sheets template includes specific fields for both high and low context communication needs. For low context readers, fields for specific details like exact dates and explicit background information are included. High context writers, on the other hand, are provided fields for including elements like historical context and implicit details, which they often find important. The script then intelligently arranges these details in an order that aligns with the expectations of a low context reader, thereby fostering smoother intercultural communication.

## Detailed Description
This project is powered by three key technologies: Google Apps Script, Google Sheets, and the Gmail API. Google Sheets serves as the user interface where email details can be entered and previewed. The functionality is driven by Google Apps Script, while the Gmail API is responsible for sending the emails.

## Prerequisites
To use gas-concat_email_preview, you need:
- A Google account with access to Google Sheets and Gmail
- Permission to run scripts on your Google account

## Installation Instructions
1. Obtain a copy of the Google Sheets template provided in this project.
2. Alternatively, create a new Google Sheet to match the layout shown in the screenshot provided in this documentation, or adjust the script to match your own sheet layout.
3. Click on Extensions > Apps Script.
4. Delete any code in the script editor and replace it with the code from this repository.
6. Save and close the script editor.

## How to Use
1. Enter data in the respective fields in your Google Sheet, ensuring all necessary information is included for both high and low context communication.
2. If there is data, it will be automatically arranged and added to the email body in an order that matches low context communication expectations.
3. Optional information should be left blank without deleting or editing the "heading" row.

## Screenshot
![Screenshot 2023-06-07 at 10 14 36](https://github.com/dariru3/gas-concat_email_preview_send/assets/107824734/b6afed4c-0120-4232-8000-17ff72516d56)
