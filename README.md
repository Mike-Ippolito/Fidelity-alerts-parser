# Fidelity-alerts-parser
Google Script to read your gmail, parse trade notifications and create a csv file. 

# Fidelity Trade Parser for Google Sheets

## Overview

This Google Apps Script automates the process of extracting trade confirmation data from Fidelity emails directly into Google Sheets. It provides a user-friendly interface to select a date range and then parses both stock and options trade confirmations, outputting the data in three convenient ways: appending to the active sheet, creating a new archive spreadsheet, and providing a downloadable CSV file.

This tool is designed to save time and eliminate manual data entry for tracking trades.

## Features

*   **Parses Stock & Options Trades:** Accurately extracts data from Fidelity trade confirmation emails for both `BUY`/`SELL` orders on stocks and `BUY CALL`/`SELL PUT` orders on options.
*   **Custom UI with Date Picker:** A user-friendly dialog box with a calendar widget allows you to easily select a specific date range for processing.
*   **Multiple Outputs:** For every run, the script provides three outputs:
    1.  **Appends to Active Sheet:** Adds the newly found trades to the bottom of your current spreadsheet for a master list.
    2.  **Creates Archive Sheet:** Generates a new, separate Google Sheet in your Drive with a timestamp, containing only the trades from that specific run.
    3.  **CSV Download:** Provides a direct download link for a CSV file of the processed trades, named with the date and time of the export.
*   **Automated Menu:** Automatically creates a custom "Fidelity Parser" menu in your Google Sheet for easy access.
*   **Robust Parsing:** Reads the HTML body of emails and uses reliable text-matching patterns (regex) to handle variations in Fidelity's email format.

## Setup Instructions

Follow these steps to install and configure the script in your Google Sheets environment.

### Step 1: Create a Google Sheet

Create a new Google Sheet or open an existing one where you want to manage your trade data. This will be your main trade log.

### Step 2: Open the Apps Script Editor

1.  In your Google Sheet, click on **Extensions > Apps Script**.
2.  A new browser tab will open with the Apps Script editor. Delete any default code in the `Code.gs` file.

### Step 3: Create the Script Files

You will need two files for this project: `Code.gs` and `DatePicker.html`.

**A. Create `Code.gs`:**
1.  Copy the entire JavaScript code provided in the final script version.
2.  Paste it into the `Code.gs` file in the Apps Script editor.

**B. Create `DatePicker.html`:**
1.  In the Apps Script editor, click the **`+`** icon next to "Files".
2.  Select **HTML**.
3.  Name the file **`DatePicker`** (the capitalization is important) and press Enter.
4.  Delete the default content in the new `DatePicker.html` file.
5.  Copy the entire HTML code provided in the final script version and paste it into the `DatePicker.html` file.

### Step 4: Save and Set Up Gmail Label

1.  Click the **Save project** icon (floppy disk) in the Apps Script editor.
2.  Go to your Gmail account.
3.  Create a new label named **`Fidelity Trades`** (the name must match exactly).
4.  Create a filter in Gmail to automatically apply this label to all incoming trade confirmation emails from Fidelity. An example filter might be `from:(fidelity.com) subject:("Trade Confirmation" OR "Order Partially Filled")`.

### Step 5: Authorize the Script

The first time you run the script, Google will ask for permissions.

1.  Reload your Google Sheet. The **"Fidelity Parser"** menu should appear.
2.  Click **Fidelity Parser > Process & Download Trades...**.
3.  A dialog box titled "Authorization Required" will appear. Click **Continue**.
4.  Choose your Google account.
5.  You may see a warning that "Google hasn’t verified this app." This is normal for personal scripts. Click **Advanced**, then click **Go to [Your Script Name] (unsafe)**.
6.  Click **Allow** to grant the script permission to read your Gmail and manage your spreadsheets.

## How to Use

Once set up, using the script is simple:

1.  **Open your main Google Sheet.**
2.  Click the **Fidelity Parser** menu at the top.
3.  Select **Process & Download Trades...**.
4.  A dialog box will appear. Use the calendar widgets to select the **Start Date** and **End Date** for the emails you want to process.
5.  Click the **Process & Download** button.
6.  The script will run. When it's finished:
    *   The new trades will be added to the bottom of your active sheet.
    *   A dialog box will appear with two links: one to **download the CSV report** and another to **open the new archive sheet** created in your Google Drive.

## File Structure

Your Apps Script project should have the following two files:

```
├── Code.gs
└── DatePicker.html
```
