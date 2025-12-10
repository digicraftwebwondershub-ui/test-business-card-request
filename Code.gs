// ===============================================================
// FILE: Code.gs
// ===============================================================

// --- CONFIGURATION ---
const sheetName = 'BCR';
const companiesSheetName = 'Companies';
const HROD_MANAGER_EMAILS = ['michelleann.delacerna@uratex.com.ph', 'corporate.training@uratex.com.ph'];
const PURCHASING_EMAILS = ['corporate.training@uratex.com.ph'];
const SENDER_NAME = 'Business Card Request';
const LOGO_FILE_ID = '1eu6GN_iqD5d2aFvVjpvj8AP9b1WKPQGF'; // <-- New Logo ID
const QR_CODE_FILE_ID = '1xlTDHDf_syY2x4Gbo9--AWXGVjee170w'; // <-- New QR Code ID

/**
 * Serves the HTML for the web app.
 */
function doGet(e) {
  if (e.parameter.v === 'manager') {
    return HtmlService.createTemplateFromFile('Manager').evaluate().setTitle('Approval Dashboard');
  }
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Business Card Request Form');
}

/**
 * Checks if the current user is a view-only user.
 */
function isViewOnlyUser() {
  try {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const purchasingEmailsSet = new Set(PURCHASING_EMAILS.map(email => email.toLowerCase()));
    return purchasingEmailsSet.has(userEmail);
  } catch (e) {
    Logger.log(`Could not determine user role: ${e.toString()}`);
    return false;
  }
}

/**
 * Gets a Google Sheet by name.
 */
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" could not be found.`);
  return sheet;
}

/**
 * Processes the form submission from the requestor.
 */
function submitRequest(formData) {
  try {
    const sheet = getSheet(sheetName);
    const requestID = `BCR-${sheet.getLastRow()}`;
    const newRow = [
      new Date(), requestID, formData.requestorName, formData.requestorEmail,
      formData.fullName, formData.cardName, formData.department, formData.position,
      formData.telephone, formData.localNumber, formData.cellphone, formData.email,
      formData.website, formData.companyName, formData.companyAddress,
      formData.includeYears,
      'Pending', '', ''
    ];
    sheet.appendRow(newRow);

    // --- Email Notifications ---
    const requestorSubmitEmail = formData.requestorEmail;
    const employeeEmail = formData.email;
    const subjectRequestor = `Business Card Request Submitted (ID: ${requestID})`;
    const emailBodyRequestor = `
      <p>Gandang gising, ${formData.requestorName}!</p>
      <p>Your request for a new business card for <strong>${formData.fullName}</strong> has been successfully submitted with ID: <strong>${requestID}</strong>.</p>
      <p>It has been forwarded for review. You will receive another email once your request has been reviewed.</p>
      <p>Thank you,</p><p>Corporate HROD</p>`;
    MailApp.sendEmail({to: requestorSubmitEmail, subject: subjectRequestor, htmlBody: emailBodyRequestor, name: SENDER_NAME});

    if (requestorSubmitEmail.toLowerCase() !== employeeEmail.toLowerCase()) {
      const emailBodyEmployee = `
        <p>Gandang gising, ${formData.fullName}!</p>
        <p>For your information, a business card request was submitted on your behalf by <strong>${formData.requestorName}</strong> (Request ID: ${requestID}).</p>
        <p>You will be notified once the request is approved or disapproved.</p><p>Thank you,</p><p>Corporate HROD</p>`;
      MailApp.sendEmail({to: employeeEmail, subject: `FYI: Business Card Request Submitted For You (ID: ${requestID})`, htmlBody: emailBodyEmployee, name: SENDER_NAME});
    }

    const managerUrl = ScriptApp.getService().getUrl() + '?v=manager';
    const emailBodyManager = `
      <p>Gandang gising!</p>
      <p>A new business card request has been submitted by <strong>${formData.requestorName}</strong> for <strong>${formData.fullName}</strong> and requires approval.</p>
      <p><strong>Request ID:</strong> ${requestID}</p>
      <p><a href="${managerUrl}" style="padding: 10px 15px; background-color: #1877f2; color: white; text-decoration: none; border-radius: 5px;">Open Approval Dashboard</a></p>
      <p>Thank you.</p>`;
    MailApp.sendEmail({to: HROD_MANAGER_EMAILS.join(','), subject: `New Business Card Request for Approval (ID: ${requestID})`, htmlBody: emailBodyManager, name: SENDER_NAME});

    const emailBodyPurchasing = `
      <p>Gandang gising!</p>
      <p>This is an automated notification that a new business card request for <strong>${formData.fullName}</strong> (ID: <strong>${requestID}</strong>) has been submitted and is awaiting approval.</p>
      <p>Thank you.</p>`;
    MailApp.sendEmail({to: PURCHASING_EMAILS.join(','), subject: `FYI: New Business Card Request Submitted (ID: ${requestID})`, htmlBody: emailBodyPurchasing, name: SENDER_NAME});

    return `Request ${requestID} submitted successfully!`;
  } catch (error) {
    Logger.log(`Submit Request Error: ${error.toString()}`);
    return `An error occurred: ${error.message}`;
  }
}

/**
 * Updates the status of a request and sends notifications.
 */
function updateStatus(rowNum, status, approverEmail, reason = '') {
  try {
    const sheet = getSheet(sheetName);
    sheet.getRange(rowNum, 17).setValue(status);
    sheet.getRange(rowNum, 18).setValue(approverEmail);
    if (status === 'Disapproved' && reason) {
      sheet.getRange(rowNum, 19).setValue(reason);
    }

    const requestData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    const [, requestID, requestorName, requestorEmail, fullName, , , , , , , employeeEmail, , companyName] = requestData;
    const subject = `Update on Business Card Request (ID: ${requestID})`;
    const reasonHtml = status === 'Disapproved' && reason ? `<p style="border-left: 4px solid #fa3e3e; padding-left: 10px; background-color: #ffebee;"><strong>Reason:</strong> ${reason}</p>` : '';

    const requestorEmailBody = `
      <p>Gandang gising!</p>
      <p>Update on the business card request for <strong>${fullName}</strong> (ID: <strong>${requestID}</strong>).</p>
      <p>The request has been <strong>${status.toUpperCase()}</strong> by <strong>${approverEmail}</strong>.</p>
      ${reasonHtml}
      <p>${status === 'Approved' ? 'You may now submit your PR for approval. Please attach this email as proof to your PR. Kindly coordinate with Purchasing for printing status.' : 'Please contact Corporate HROD for more details.'}</p>
      <p>Thank you,</p><p>Corporate HROD</p>`;

    const emailOptions = {name: SENDER_NAME, subject, htmlBody: requestorEmailBody};
    MailApp.sendEmail({...emailOptions, to: requestorEmail});
    if (requestorEmail.toLowerCase() !== employeeEmail.toLowerCase()) {
      MailApp.sendEmail({...emailOptions, to: employeeEmail});
    }

    let purchasingSubject, purchasingBody;
    if (status === 'Approved') {
      purchasingSubject = `For Printing: Business Card Request for ${fullName} (ID: ${requestID})`;
      purchasingBody = `
        <p>Gandang gising!</p>
        <p>The business card request for <strong>${fullName}</strong> of <strong>${companyName}</strong> has been <strong>APPROVED</strong> by Corporate HROD.</p>
        <p>The requestor, ${requestorName}, will coordinate with your department for the PR and printing process.</p>
        <p>Thank you.</p>`;
    } else {
      purchasingSubject = `Disapproved Business Card Request (ID: ${requestID})`;
      purchasingBody = `
        <p>Gandang gising!</p>
        <p>This is to inform you that the business card request for <strong>${fullName}</strong> (ID: <strong>${requestID}</strong>) has been <strong>DISAPPROVED</strong>.</p>
        <p><strong>Reason:</strong> ${reason}</p>
        <p>No further action is required.</p>
        <p>Thank you.</p>`;
    }
    MailApp.sendEmail({to: PURCHASING_EMAILS.join(','), subject: purchasingSubject, htmlBody: purchasingBody, name: SENDER_NAME});

    const otherManagers = HROD_MANAGER_EMAILS.filter(email => email.toLowerCase() !== approverEmail.toLowerCase());
    if (otherManagers.length > 0) {
      const fyiBody = `
        <p>Gandang gising!</p>
        <p>FYI: The business card request for <strong>${fullName}</strong> (ID: <strong>${requestID}</strong>) has been <strong>${status.toUpperCase()}</strong> by <strong>${approverEmail}</strong>.</p>
        ${status === 'Disapproved' && reason ? `<p><strong>Reason:</strong> ${reason}</p>` : ''}
        <p>No further action is required.</p>`;
      MailApp.sendEmail({to: otherManagers.join(','), subject: `FYI: Request ${status} by ${approverEmail} (ID: ${requestID})`, htmlBody: fyiBody, name: SENDER_NAME});
    }

    return `Request ${requestID} has been ${status}.`;
  } catch (error) {
    Logger.log(`Update Status Error: ${error.toString()}`);
    return `An error occurred: ${error.message}`;
  }
}

/**
 * Retrieves all requests from the spreadsheet.
 */
function getRequests() {
  const sheet = getSheet(sheetName);
  if (sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return data.map((row, index) => ({
    rowNum: index + 2,
    requestID: row[1],
    requestorName: row[2],
    fullName: row[4],
    cardName: row[5],
    department: row[6],
    position: row[7],
    telephone: row[8],
    localNumber: row[9],
    cellphone: row[10],
    email: row[11],
    website: row[12],
    companyName: row[13],
    companyAddress: row[14],
    includeYears: row[15],
    status: row[16],
    approvedBy: row[17],
    disapprovalReason: row[18] || ''
  }));
}

/**
 * Gets a unique list of company names.
 */
function getCompanyList() {
  try {
    const sheet = getSheet(companiesSheetName);
    if (sheet.getLastRow() < 2) return [];
    return [...new Set(sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat())];
  } catch (e) {
    Logger.log(`Failed to get company list: ${e.toString()}`);
    return [];
  }
}

/**
 * Gets all addresses for a specific company.
 */
function getAddressesForCompany(companyName) {
  try {
    const sheet = getSheet(companiesSheetName);
    if (!companyName || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return data.filter(row => row[0] === companyName).map(row => row[1]);
  } catch (e) {
    Logger.log(`Failed to get addresses: ${e.toString()}`);
    return [];
  }
}

/**
 * Fetches an image from Drive and encodes it as a Base64 string.
 */
function getImageAsBase64(fileId) {
  const placeholder = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch (e) {
    Logger.log(`Failed to get image (ID: ${fileId}): ${e.toString()}`);
    return placeholder;
  }
}

/**
 * Wrapper function for the logo.
 */
function getLogoAsBase64() {
  return getImageAsBase64(LOGO_FILE_ID);
}

/**
 * Fetches and encodes the QR code image.
 */
function getQrCodeAsBase64() {
  return getImageAsBase64(QR_CODE_FILE_ID); // <-- Changed to use new ID
}

/**
 * Gets the email of the active user.
 */
function getManagerEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Sends reminder emails for pending requests older than 24 hours.
 */
function sendPendingReminders() {
  try {
    const sheet = getSheet(sheetName);
    if (sheet.getLastRow() < 2) return;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const now = new Date();
    const managerUrl = ScriptApp.getService().getUrl() + '?v=manager';
    const pendingRequests = data.filter(row => {
      const status = row[16];
      const timestamp = new Date(row[0]);
      const timeDifferenceHours = (now - timestamp) / (1000 * 3600);
      return status === 'Pending' && timeDifferenceHours >= 24;
    });

    if (pendingRequests.length === 0) {
      Logger.log('No pending requests needed a reminder.');
      return;
    }

    pendingRequests.forEach(row => {
      const [timestamp, requestID, , , fullName] = row;
      const subject = `Reminder: Pending Business Card Request (ID: ${requestID})`;
      const emailBody = `
        <p>Gandang gising!</p>
        <p>Reminder: A business card request for <strong>${fullName}</strong> (ID: <strong>${requestID}</strong>), submitted on ${timestamp.toLocaleString('en-US', {timeZone: 'Asia/Manila'})}, is still awaiting approval.</p>
        <p><a href="${managerUrl}" style="padding: 10px 15px; background-color: #f7b924; color: white; text-decoration: none; border-radius: 5px;">Open Approval Dashboard</a></p>
        <p>Thank you.</p>`;
      MailApp.sendEmail({to: HROD_MANAGER_EMAILS.join(','), subject, htmlBody: emailBody, name: SENDER_NAME});
    });
    Logger.log(`Sent ${pendingRequests.length} reminder(s).`);
  } catch (error) {
    Logger.log(`Failed to send pending reminders: ${error.toString()}`);
  }
}
