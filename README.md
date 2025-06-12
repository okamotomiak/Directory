# Leadership Directory Tools

This Apps Script project manages a leadership directory and provides tools for mail merges and meeting reminders.

## Features
- Create a formatted leadership directory with sample data
- Send targeted mail merges using saved templates
- Manage meeting reminders via a sidebar UI
- Import new leader information from a Google Form

## Usage
1. Open the spreadsheet and run **Create Directory** from the custom menu to generate sheets.
2. Use **Create Leader Signup Form** to create a form for collecting leader info.
3. Use **Send Mail Merge…** to send personalised emails.
4. Configure recurring reminders with **Meeting Reminder Setup…**.

This repository includes HTML files for dialogs and JavaScript files with the Apps Script functions.

## Styling
The HTML dialogs load `styles.css` at runtime using
`HtmlService.createHtmlOutputFromFile('styles.css').getContent()`. This keeps
styles in a separate file while remaining fully compatible with Google Apps
Script's HTML service.
