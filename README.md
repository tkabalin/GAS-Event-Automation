# GAS Event Automation

This project is a **Google Apps Script (GAS)** solution designed to add administrative functionality to a Google Sheet used for **event booking and tracking**. It was originally intended for the Cannons Creek High School sound and lighting team so that teachers could fill in event information and requirements, and the head of the team can assign members to each event. 

## Features

* **Custom Admin Menu:** Adds a top-level menu in the Spreadsheet UI for administrative tasks (i.e. running the below functions).
* **Sort Events:** Sorts event data on the "Booking Form" and Archive sheets by date.
* **Archive Old Events:** Moves old event records from the "Booking Form" to the Archive sheet.
* **Notify:** Allows a user to select a cell and send an email notification with that cell's value to a specified address.

## Detailed Usage and Setup Guide

### 1. Adding the Script to Your Google Sheet

1.  **Open your Google Sheet.** This should be the spreadsheet you use for tracking event bookings.
2.  Go to **Extensions** > **Apps Script**. This will open the script editor in a new tab.
3.  In the Apps Script editor, you'll see a default file named `Code.gs`. **Copy and paste the entire JavaScript code** into this file, overwriting any existing content.
4.  **Save the script** by clicking the floppy disk icon (💾) or pressing `Ctrl/Cmd + S`. You can name the script project, e.g., "EventManagerScript."
5.  **Refresh your Google Sheet.** After a few moments, the custom **"Admin"** menu will appear in the spreadsheet toolbar, confirming the script is loaded.

---

### 2. Customising Sheet Names and Cell Indexes

The script relies on specific sheet names and cell references. You'll need to **update the names in the code** if your spreadsheet structure differs.

| Function | Code Line Example | Purpose |
| :--- | :--- | :--- |
| `sortSheet` | `.getSheetByName("Booking Form")` | **Main event sheet** name. |
| `sortSheet` | `.getSheetByName("Archive (2022)")` | **Historical archive** sheet name. |
| `sortSheet` | `.getRange("A4:K400")` | Adjust the **sorting range** to cover all your event data rows. |
| `notifyOther` | `.getSheetByName('Notifications')` | **Temporary sheet** used for composing and sending emails. |
| `archiveEvent` | `.getSheetByName('Archive (2023)')` | **Current archive** sheet name where old events are moved. |
| `archiveEvent` | `.getRange('A4:K4')` | The **single event row** to be archived (usually the first data row). |
| `archiveEvent` | `.getRange('H1:K1').getValue()` | The cell (`H1` in this case) that contains the **current count of archived events**. This count determines the next available row for pasting the new event. |

---

### 3. Setting Up Midnight Automation (Triggers)

To ensure old events are consistently archived, you can set up a time-driven trigger to run the `archiveEvent()` function every night.

1.  In the Apps Script editor, click the **Triggers** icon (which looks like a clock ⏱️) in the sidebar on the left.
2.  Click the **"+ Add Trigger"** button in the bottom right corner.
3.  Configure the trigger with the following settings:
    * **Choose which function to run:** Select `archiveEvent`.
    * **Select event source:** Choose **Time-driven**.
    * **Select type of time-based trigger:** Choose **Day timer**.
    * **Select time of day:** Choose a time band (e.g., **"12am to 1am"**) when you want the function to execute.
4.  Click **Save**. You'll likely need to review and **authorise the script's permissions** (to access the spreadsheet and send emails) the first time you set a trigger.

The `archiveEvent` function will now automatically run every night, moving events with a date older than today's date from the "Booking Form" sheet to the archive.

The above steps can be repeated to sort the events every night. 

## Dependencies and Layout

This script relies on a specific Google Sheet structure, including sheets named "Booking Form," "Archive (2022)," etc.

➡️ **For a visual reference of the required column and cell layout, please see the dedicated [Layout Reference Document](LAYOUT.md).**

*This README was created with assistance from Google Gemini.*
