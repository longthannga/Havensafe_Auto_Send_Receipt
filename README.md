# Havensafe Donation Receipt Automation

> This project was created during my volunteer internship at [Havensafe](https://havensafe.org)

This Google Apps Script automates the process of scanning incoming donation emails, extracting donor details, generating PDF receipts from a Google Docs template, emailing receipts to donors, and logging donations into a Google Sheet.

## ✨ Features

* Automatically scans **unread Gmail messages** for PayPal and generic donation notifications.
* Extracts donor **name, email, amount, and date** from email body.
* Generates a PDF receipt from a **Google Docs template** with placeholders.
* Emails the donor their receipt.
* Logs each donation into a **Google Sheets spreadsheet** with confirmation tracking.
* Moves processed receipts into a **Google Drive folder** for record-keeping.

## 🛠 Setup Instructions

### 1. Clone / Add Script

Copy this script into a [Google Apps Script project](https://script.google.com/).

### 2. Configure Variables

Update the following constants in the script:

```js
const TEMPLATE_ID    = 'YOUR TEMPLATE_ID';       // Google Docs template ID  
const SPREADSHEET_ID = 'YOUR SPREADSHEET_ID';   // Google Sheets ID  
const SHEET_NAME     = 'YOUR SHEET_NAME';       // Target sheet name  
const REP_NAME       = 'YOUR REP_NAME';         // Representative name  
const REP_TITLE      = 'YOUR REP_TITLE';        // Representative title  
```

### 3. Create Your Google Docs Template

Your template should include placeholders in double curly braces, for example:

```
{{DATE}}  
{{NAME}}  
{{AMOUNT}}  
{{DONATE_DATE}}  
{{REP_NAME}}  
{{REP_TITLE}}  
```

### 4. Set Up Google Sheet

Prepare a sheet with these columns:

| Donor Name | Donation Date | Amount | Payment Type | Sent To | Receipt Link | Receipt Date | Assigned To | Thank You Written? | Thank You Sent? | Note |
| ---------- | ------------- | ------ | ------------ | ------- | ------------ | ------------ | ----------- | ------------------ | --------------- | ---- |

### 5. Configure Gmail Query

By default, the script looks for unread donation emails with subjects:

* `"Payment received from"` (PayPal)
* `"Donation Received"` (Generic)

You can edit the `QUERY` variable to fit your email provider.

### 6. Set Up Triggers

* In **Apps Script editor → Triggers**, set `processDonations` to run:

  * Time-driven (e.g., every 10 minutes), or
  * On Gmail events (if available).

## 📂 Folder Organization

* Processed receipts are automatically moved into a Google Drive folder:

  * **Processed Receipts** (auto-created if not found).

## 📧 Email Receipt Example

Subject: *Your Havensafe Donation Receipt*
Body:

```
Dear [Donor Name],  

Thank you for your generous donation of $[Amount] on [Date].  
Please find your donation receipt attached as a PDF.  

With gratitude,  
[REP_NAME]  
[REP_TITLE]  
```

## ⚠️ Notes

* Only supports PayPal and simple generic donation formats. You can extend `parseDonationEmail()` to handle other email structures.
* Make sure Gmail, Docs, Drive, and Sheets APIs are enabled for your Google account.
* Test thoroughly with sample emails before enabling automation.
