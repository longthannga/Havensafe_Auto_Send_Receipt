/******************************************************
 * Havensafe Donation Receipt Automation
 ******************************************************/

// —— CONFIGURATION ——

// Replace with your Google Docs template's ID
const TEMPLATE_ID    = 'YOUR TEMPLATE_ID';

// Your spreadsheet ID and sheet name
const SPREADSHEET_ID = 'YOUR SPREADSHEET_ID';
const SHEET_NAME     = 'YOUR SHEET_NAME';

// Your signature block
const REP_NAME       = 'YOUR REP_NAME';
const REP_TITLE      = 'YOUR REP_TITLE';

// Search unread emails matching either format
const QUERY = [
  'subject:"Payment received from"',
  'subject:"Donation Received"'
].join(' OR ') + ' is:unread';


/** 
 * Main entry point: scans unread donation emails,
 * parses donor info, generates & sends receipt.
 */
function processDonations() {
  GmailApp.search(QUERY).forEach(thread => {
    const msg  = thread.getMessages()[0];
    const body = msg.getPlainBody();
    const d    = parseDonationEmail(body);

    if (d && d.email && d.amount) {
      const receiptDate = Utilities.formatDate(new Date(), Session.getTimeZone(), 'yyyy-MM-dd');
      const docUrl = createReceipt(d.name, d.email, d.amount, d.forwardedDate, receiptDate);
      addToSpreadsheet(d.name, d.transactionDate || d.forwardedDate, d.amount, docUrl);
      thread.markRead();
    } else {
      Logger.log('Unparsed donation email:\n' + body);
    }
  });
}


/**
 * Extracts donor details from email body.
 * Supports two formats: PayPal "Payment received" and generic "Donation Received."
 *
 * @param {string} body - Plain-text email content.
 * @returns {{name:string, email:string, amount:string, date:string}|null}
 */
function parseDonationEmail(body) {
  // — Format A: PayPal "Payment received" —
  if (/You received a payment of/i.test(body)) {
    const amountMatch = /You received a payment of.*?\$([\d,\.]+)\s+USD/i.exec(body);
    const txDateMatch = /\*Transaction date\*[\s\S]*?([A-Za-z]{3,}\s+\d{1,2},\s+\d{4}\s+\d{2}:\d{2}:\d{2}\s+\w+)/i.exec(body);

    // capture "name"
    const nameMatch = /\*Buyer information\*[\s\S]*?\n([^\n]+)/i.exec(body);
    // capture the line _after_ the name -- that's the donor's email
    const emailMatch = /\*Buyer information\*[\s\S]*?\n[^\n]+\n([^\n]+)/i.exec(body);

    // grab the forwarded header's Date: line
    const fwdDateMatch = /^Date:\s*(.+)$/m.exec(body);

    return {
      name:            nameMatch   ? nameMatch[1].trim()        : 'Friend',
      email:           emailMatch  ? emailMatch[1].trim()       : '',
      amount:          amountMatch ? amountMatch[1]             : '',
      transactionDate: txDateMatch ? txDateMatch[1].trim()      : '',
      forwardedDate:   fwdDateMatch? fwdDateMatch[1].trim()     : ''
    };
  }

  // — Format B: Generic Donation Received —
  m = {
    name:  /from\s+(.+)\s+\(/i.exec(body),
    email: /from\s+.+\((.+@.+\..+)\)/i.exec(body),
    amount: /donation of \$([\d,\.]+)/i.exec(body),
    forwardedDate: /Forwarded message ---------\s*From:[\s\S]*?Date:\s*(.+)/.exec(body)
  };

  if (m.email && m.amount) {
    return {
      name:   m.name  ? m.name[1]   : 'Friend',
      email:  m.email[1],
      amount: m.amount[1],
      forwardedDate:  m.forwardedDate[1]
    };
  }

  // No known format matched
  return null;
}


/**
 * Copies the Docs template, replaces placeholders,
 * converts to PDF, and emails the receipt.
 *
 * @param {string} name   - Donor's name.
 * @param {string} email  - Donor's email address.
 * @param {string} amount - Donation amount.
 * @param {string} date   - Donation date.
 */
function createReceipt(name, email, amount, forwardedDate, date) {
  // 1. Copy the template
  const copyName = `DonationReceipt_${name.replace(/\s+/g, '_')}`;
  const copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(copyName);
  const docId    = copyFile.getId();
  const doc      = DocumentApp.openById(docId);
  const body     = doc.getBody();

  // 2. Replace all placeholders
  body.replaceText('{{DATE}}',     date);
  body.replaceText('{{NAME}}',     name);
  body.replaceText('{{AMOUNT}}',   amount);
  body.replaceText('{{DONATE_DATE}}',forwardedDate);
  body.replaceText('{{REP_NAME}}', REP_NAME);
  body.replaceText('{{REP_TITLE}}',REP_TITLE);

  doc.saveAndClose();
  copyFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);


  // 3. Convert to PDF and send
  const pdf     = copyFile.getAs('application/pdf');
  const subject = 'Your Havensafe Donation Receipt';
  const message = [
    `Dear ${name},`,
    '',
    `Thank you for your generous donation of ${amount} on ${forwardedDate}.`,
    'Please find your donation receipt attached as a PDF.',
    '',
    'With gratitude,',
    REP_NAME,
    REP_TITLE
  ].join('\n');

  const processedFolder = getOrCreateFolder('Processed Receipts');
  copyFile.moveTo(processedFolder);

  MailApp.sendEmail({
    to:          email,
    subject:     subject,
    body:        message,
    attachments: [pdf]
  });

  // GmailApp.createDraft(email, subject, message, {
  //     attachments: [pdf]
  //   });
  
  // Return the document URL for spreadsheet linking
  return copyFile.getUrl();
}


/**
 * Adds donation information to the Google Sheets spreadsheet.
 *
 * @param {string} donorName - Name of the donor
 * @param {string} donationDate - Date of the donation
 * @param {string} amount - Donation amount
 * @param {string} email - Donor's email address
 */
function addToSpreadsheet(donorName, donationDate, amount, docUrl) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log(`Sheet "${SHEET_NAME}" not found in spreadsheet`);
      return;
    }
    
    // Format the donation date to match your spreadsheet format (M/D/YY)
    let formattedDate = '';
    if (donationDate) {
      try {
        // Parse the date and format it
        const dateObj = new Date(donationDate);
        if (!isNaN(dateObj.getTime())) {
          formattedDate = Utilities.formatDate(dateObj, Session.getTimeZone(), 'M/d/yy');
        } else {
          formattedDate = donationDate; // Use original if parsing fails
        }
      } catch (e) {
        formattedDate = donationDate; // Use original if formatting fails
      }
    }
    
    // Remove $ sign and format amount as number
    const numericAmount = parseFloat(amount.replace(/[$,]/g, ''));
    
    // Get current date for confirmation letter date
    const today = new Date();
    const confirmationDate = Utilities.formatDate(today, Session.getTimeZone(), 'M/d/yy');
    
    // Determine payment type based on context (you might want to enhance this logic)
    const paymentType = 'PayPal'; // Default to PayPal, you can enhance this based on email parsing
    
    // Add new row with donor information
    // Based on your spreadsheet columns: A=Donor Name, B=Donation Date, C=Amount, D=Payment Type, E=Sent To, F=Link to custom Donation Confirmation Letter, G=Donation Confirmation Letter date, H=Who assigned to write thank you card?, I=Thank you card written?, J=Thank you card sent?, K=Note
    const data = [
      donorName,
      formattedDate,
      numericAmount,
      paymentType,
      '',
      docUrl,
      confirmationDate,
      'Auto-processed',
      'yes',
      'yes',
      ''
    ];

    // Insert at row 2 (below headers)
    sheet.insertRowBefore(2);
    sheet.getRange(2, 1, 1, data.length).setValues([data]);
    
    Logger.log(`Added donation record for ${donorName}: $${amount}`);
    
  } catch (error) {
    Logger.log(`Error adding to spreadsheet: ${error.toString()}`);
  }
}


function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}