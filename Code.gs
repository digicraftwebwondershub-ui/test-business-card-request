// ===============================================================
// FILE: Code.gs
// ===============================================================

// IMPORTANT: Replace with your actual deployed Web App URL
const WEB_APP_URL = 'https://script.google.com/a/macros/uratex.com.ph/s/AKfycbybTFROWNydAN12rXcL7viTQwkrNvPhQgx5isd2bdubh7vT7Cl8Ehq32JQZZ8t7OHNQ/exec'; 

const SHEET_NAME = 'BCR';
const LOGS_SHEET_NAME = 'AuditTrail';
const COMPANIES_SHEET_NAME = 'Companies';
const ROUTING_SHEET_NAME = 'PurchasingRouting';
const HROD_MANAGER_EMAILS = ['michelleann.delacerna@uratex.com.ph'];
const PURCHASING_EMAILS = ['corporate.training@uratex.com.ph']; // Default fallback
const SENDER_NAME = 'Business Card Request';

const LOGO_FILE_ID = '1eu6GN_iqD5d2aFvVjpvj8AP9b1WKPQGF';
const QR_CODE_FILE_ID = '1xlTDHDf_syY2x4Gbo9--AWXGVjee170w';

/**
 * SERVE HTML
 */
function doGet(e) {
  const page = e.parameter.v;
  let template;

  if (page === 'manager') {
    template = HtmlService.createTemplateFromFile('Manager');
    template.data = {}; 
  } else if (page === 'edit') {
    template = HtmlService.createTemplateFromFile('EditRequest');
    template.data = { requestId: e.parameter.id || '' };
  } else {
    template = HtmlService.createTemplateFromFile('Index');
    template.data = { requestId: e.parameter.id || '' };
  }
  
  return template.evaluate()
    .setTitle(page === 'manager' ? 'Approval Dashboard' : 'Business Card Request')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ... (Keep existing Utilities: getIsPurchasing, getRequestById, getRequests, getHistory, getSheet, logAudit, getCompanyList, getAddressesForCompany, getLogoAsBase64, getQrCodeAsBase64, getManagerEmail, getImageAsBase64) ...
// (These functions remain unchanged from the previous correct version)
function getIsPurchasing() {
  const email = Session.getActiveUser().getEmail().toLowerCase();
  const baseEmails = PURCHASING_EMAILS.map(e => e.toLowerCase());

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ROUTING_SHEET_NAME);
    if (sheet && sheet.getLastRow() > 1) {
      const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat();
      const sheetEmails = data.join(',').split(',').map(e => e.trim().toLowerCase()).filter(e => e !== '');
      if (sheetEmails.includes(email)) return true;
    }
  } catch (e) {
    Logger.log("Error checking purchasing permissions: " + e.message);
  }

  return baseEmails.includes(email);
}
function getRequestById(id) {
  const requests = getRequests();
  return requests.find(r => r.requestID === id) || null;
}
function getRequests() {
  try {
    const sheet = getSheet(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const data = range.getValues();
    return data.map((row, index) => {
      const get = (i) => (row[i] === undefined ? '' : row[i]);
      let dateStr = '';
      try { 
        if (get(0) instanceof Date) {
           dateStr = get(0).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
        } else {
           dateStr = String(get(0)); 
        }
      } catch(e) { dateStr = 'N/A'; }
      return {
        rowNum: index + 2,
        date: dateStr,
        requestID: String(get(1)),
        requestorName: String(get(2)),
        requestorEmail: String(get(3)),
        fullName: String(get(4)),
        cardName: String(get(5)),
        department: String(get(6)),
        position: String(get(7)),
        telephone: String(get(8)),
        localNumber: String(get(9)),
        cellphone: String(get(10)),
        email: String(get(11)),
        website: String(get(12)),
        companyName: String(get(13)),
        companyAddress: String(get(14)),
        includeYears: Boolean(get(15)),
        requestReason: String(get(16)), 
        status: String(get(17) || 'Pending'), 
        approvedBy: String(get(18)),    
        reason: String(get(19))         
      };
    });
  } catch (e) { return []; }
}
function getMyRequests() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const all = getRequests();
    return all.filter(r => 
      String(r.requestorEmail).toLowerCase() === email || 
      String(r.email).toLowerCase() === email
    );
  } catch (e) { return []; }
}
function getHistory(requestID) {
  try {
    const sheet = getSheet(LOGS_SHEET_NAME);
    if (sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
    return data.filter(row => String(row[1]) === String(requestID)).map(row => {
        let tsStr = '';
        try { tsStr = row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "MMM d, yyyy h:mm a") : String(row[0]); } catch(e) { tsStr = 'N/A'; }
        return { timestamp: tsStr, user: String(row[2]), action: String(row[3]), field: String(row[4]), oldVal: String(row[5]), newVal: String(row[6]) };
      }).reverse();
  } catch (e) { return []; }
}
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    if (name === LOGS_SHEET_NAME) { sheet = ss.insertSheet(LOGS_SHEET_NAME); sheet.appendRow(['Timestamp', 'Request ID', 'User', 'Action', 'Field', 'Old Value', 'New Value']); } 
    else { throw new Error(`Sheet "${name}" not found.`); }
  }
  return sheet;
}
function logAudit(requestId, user, action, field, oldVal, newVal) {
  const sheet = getSheet(LOGS_SHEET_NAME);
  sheet.appendRow([new Date(), requestId, user, action, field, oldVal, newVal]);
}
function getCompanyList() {
  const s = getSheet(COMPANIES_SHEET_NAME);
  return s.getLastRow() > 1 ? [...new Set(s.getRange(2, 1, s.getLastRow()-1).getValues().flat())] : [];
}
function getAddressesForCompany(c) {
  const s = getSheet(COMPANIES_SHEET_NAME);
  if(s.getLastRow() < 2) return [];
  const d = s.getRange(2, 1, s.getLastRow()-1, 2).getValues();
  return d.filter(r => r[0] === c).map(r => r[1]);
}

/**
 * FETCH PURCHASING EMAILS BASED ON COMPANY AND ADDRESS
 */
function getPurchasingEmails(company, address) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ROUTING_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return PURCHASING_EMAILS;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

    // Clean address by removing " Philippines" if present for better matching
    const cleanAddress = String(address).replace(/\s*Philippines$/i, '').trim().toLowerCase();
    const cleanCompany = String(company).trim().toLowerCase();

    // 1. Try Exact match (Company + Address)
    const match = data.find(row => {
      const sheetCompany = String(row[0]).trim().toLowerCase();
      const sheetAddress = String(row[1]).trim().toLowerCase();
      return sheetCompany === cleanCompany && sheetAddress === cleanAddress;
    });

    if (match && match[2]) {
      return match[2].split(',').map(e => e.trim()).filter(e => e !== '');
    }

    // 2. Try partial address match if the sheet address is contained within the submitted address
    const partialMatch = data.find(row => {
      const sheetCompany = String(row[0]).trim().toLowerCase();
      const sheetAddress = String(row[1]).trim().toLowerCase();
      return sheetCompany === cleanCompany && sheetAddress !== '' && cleanAddress.includes(sheetAddress);
    });

    if (partialMatch && partialMatch[2]) {
      return partialMatch[2].split(',').map(e => e.trim()).filter(e => e !== '');
    }

    // 3. Fallback search: match company only
    const companyMatch = data.find(row =>
      String(row[0]).trim().toLowerCase() === cleanCompany &&
      (!row[1] || String(row[1]).trim() === '')
    );

    if (companyMatch && companyMatch[2]) {
      return companyMatch[2].split(',').map(e => e.trim()).filter(e => e !== '');
    }

  } catch (e) {
    Logger.log("Error in getPurchasingEmails: " + e.message);
  }
  return PURCHASING_EMAILS;
}
function getLogoAsBase64() { return getImageAsBase64(LOGO_FILE_ID); }
function getQrCodeAsBase64() { return getImageAsBase64(QR_CODE_FILE_ID); }
function getManagerEmail() { return Session.getActiveUser().getEmail(); }
function getImageAsBase64(id) {
  try { return `data:${DriveApp.getFileById(id).getBlob().getContentType()};base64,${Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes())}`; } 
  catch (e) { return ''; }
}

/**
 * FORM PROCESSING
 */
function processForm(formData) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); 

  try {
    const sheet = getSheet(SHEET_NAME);
    const userEmail = Session.getActiveUser().getEmail();
    let requestID = formData.requestID;
    let isEdit = false;
    let rowNum = -1;

    if (requestID) {
      const ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
      const index = ids.indexOf(requestID);
      if (index !== -1) { rowNum = index + 2; isEdit = true; }
    } else {
      requestID = `BCR-${sheet.getLastRow() + 1}`; 
    }

    const newData = [
      formData.requestorName, formData.requestorEmail,
      formData.fullName, formData.cardName, formData.department, formData.position,
      formData.telephone, formData.localNumber, formData.cellphone, formData.email,
      formData.website, formData.companyName, formData.companyAddress,
      formData.includeYears, formData.requestReason
    ];

    let changeLog = [];

    if (isEdit) {
      const range = sheet.getRange(rowNum, 3, 1, 15); 
      const currentValues = range.getValues()[0];
      const fields = ['Requestor', 'Req Email', 'Full Name', 'Card Name', 'Dept', 'Position', 'Tel', 'Local', 'Cell', 'Email', 'Web', 'Company', 'Address', 'Years', 'Req Reason'];
      
      for (let i = 0; i < fields.length; i++) {
        if (String(currentValues[i]) !== String(newData[i])) {
          logAudit(requestID, userEmail, 'Edit', fields[i], currentValues[i], newData[i]);
          changeLog.push(`<strong>${fields[i]}:</strong> ${currentValues[i]} &rarr; ${newData[i]}`);
        }
      }
      range.setValues([newData]);
      
      const currentStatus = sheet.getRange(rowNum, 18).getValue();
      if (String(currentStatus).includes('Returned')) {
        sheet.getRange(rowNum, 18).setValue('Pending'); 
        logAudit(requestID, userEmail, 'Status Change', 'Status', 'Returned', 'Pending');
        notifyManagers(requestID, formData, 'Pending', changeLog);
      }
    } else {
      // DUPLICATE CHECK
      if (sheet.getLastRow() > 1) {
        const allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
        const duplicate = allData.find(r => {
          const name = String(r[4]).trim().toLowerCase();
          const position = String(r[7]).trim().toLowerCase();
          const company = String(r[13]).trim().toLowerCase();
          const status = String(r[17]);
          const isSameDetails = (name === formData.fullName.trim().toLowerCase() && position === formData.position.trim().toLowerCase() && company === formData.companyName.trim().toLowerCase());
          const isActive = (status === 'Pending' || status === 'Resubmitted');
          return isSameDetails && isActive;
        });
        if (duplicate) {
          return { success: false, message: `A pending request already exists for ${formData.fullName} as ${formData.position} at ${formData.companyName}.` };
        }
      }

      const newRow = [new Date(), requestID, ...newData, 'Pending', '', ''];
      sheet.appendRow(newRow);
      logAudit(requestID, userEmail, 'Create', 'All', '', 'Initial Submission');
      sendSubmissionEmails(requestID, formData);
    }

    // *** FIX: RETURN THE URL SO THE FRONTEND KNOWS WHERE TO GO ***
    return { 
      success: true, 
      message: `Request ${requestID} successfully ${isEdit ? 'updated' : 'submitted'}!`,
      redirectUrl: WEB_APP_URL // Pass the URL constant to the frontend
    };

  } catch (e) {
    Logger.log(e);
    return { success: false, message: `Error: ${e.message}` };
  } finally { lock.releaseLock(); }
}

// ... (Keep existing updateStatus, emails, and reminders) ...
function updateStatus(rowNum, newStatus, reason) {
  try {
    const sheet = getSheet(SHEET_NAME);
    const currentUser = Session.getActiveUser().getEmail();
    const range = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
    const rowValues = range.getValues()[0];
    const requestID = rowValues[1];
    const oldStatus = rowValues[17];
    sheet.getRange(rowNum, 18).setValue(newStatus); 
    sheet.getRange(rowNum, 19).setValue(currentUser); 
    sheet.getRange(rowNum, 20).setValue(reason || ''); 
    logAudit(requestID, currentUser, 'Decision', 'Status', oldStatus, newStatus);
    if (reason) logAudit(requestID, currentUser, 'Comment', 'Reason', '', reason);
    sendStatusUpdateEmail(rowValues, newStatus, currentUser, reason);
    return `Request ${requestID} updated.`;
  } catch (e) { throw new Error(e.message); }
}

function getRequestDetailsHtml(d) {
  return `
    <div style="background:#f8f9fa; padding:15px; border-radius:8px; border:1px solid #e9ecef; margin:15px 0; color:#333; font-family:Arial, sans-serif; font-size:13px; line-height:1.5;">
      <strong style="color:#1877f2; font-size:14px; display:block; margin-bottom:10px; border-bottom:1px solid #ddd; padding-bottom:5px;">REQUEST DETAILS:</strong>
      <table style="width:100%; border-collapse:collapse;">
        <tr><td style="width:140px; color:#555; font-weight:bold; padding:3px 0;">Requestor:</td><td>${d.requestorName}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Employee:</td><td>${d.fullName}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Card Name:</td><td>${d.cardName}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Position:</td><td>${d.position}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Department:</td><td>${d.department}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Company:</td><td>${d.companyName}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Address:</td><td>${d.companyAddress}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Telephone:</td><td>${d.telephone} ${d.localNumber ? 'loc. '+d.localNumber : ''}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Cellphone:</td><td>${d.cellphone}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Email:</td><td>${d.email}</td></tr>
        <tr><td style="color:#555; font-weight:bold; padding:3px 0;">Reason:</td><td>${d.requestReason || 'N/A'}</td></tr>
      </table>
    </div>
  `;
}

function sendSubmissionEmails(requestID, formData) {
  const detailsHtml = getRequestDetailsHtml(formData);
  const emailBodyRequestor = `
    <p>Gandang gising, ${formData.requestorName}!</p>
    <p>Your request for a new business card for <strong>${formData.fullName}</strong> has been successfully submitted with ID: <strong>${requestID}</strong>.</p>
    ${detailsHtml}
    <p>It has been forwarded for review. You will receive another email once your request has been reviewed.</p>
    <p>Thank you,</p><p>Corporate HROD</p>`;
  MailApp.sendEmail({to: formData.requestorEmail, subject: `Business Card Request Submitted (ID: ${requestID})`, htmlBody: emailBodyRequestor, name: SENDER_NAME});

  if (formData.requestorEmail.toLowerCase() !== formData.email.toLowerCase()) {
    const emailBodyEmployee = `
      <p>Gandang gising, ${formData.fullName}!</p>
      <p>For your information, a business card request was submitted on your behalf by <strong>${formData.requestorName}</strong> (Request ID: ${requestID}).</p>
      ${detailsHtml}
      <p>You will be notified once the request is approved or disapproved.</p><p>Thank you,</p><p>Corporate HROD</p>`;
    MailApp.sendEmail({to: formData.email, subject: `FYI: Business Card Request Submitted For You (ID: ${requestID})`, htmlBody: emailBodyEmployee, name: SENDER_NAME});
  }
  notifyManagers(requestID, formData, 'New');
  const purchasingEmails = getPurchasingEmails(formData.companyName, formData.companyAddress);
  const emailBodyPurchasing = `
    <p>Gandang gising!</p>
    <p>This is an automated notification that a new business card request for <strong>${formData.fullName}</strong> (ID: <strong>${requestID}</strong>) has been submitted and is awaiting approval.</p>
    ${detailsHtml}
    <p>Thank you.</p>`;
  MailApp.sendEmail({to: purchasingEmails.join(','), subject: `FYI: New Business Card Request Submitted (ID: ${requestID})`, htmlBody: emailBodyPurchasing, name: SENDER_NAME});
}

function notifyManagers(requestID, formData, statusLabel, changes = []) {
  const managerUrl = `${WEB_APP_URL}?v=manager`;
  const detailsHtml = getRequestDetailsHtml(formData);
  let changeHtml = '';
  if (changes.length > 0) {
    changeHtml = `
      <div style="background-color:#e3f2fd; border-left: 5px solid #1877f2; padding: 15px; margin: 15px 0;">
        <strong style="color:#1877f2; font-size:14px;">SUMMARY OF CHANGES (RESUBMISSION):</strong>
        <ul style="margin:10px 0 0 20px; color:#333;">${changes.map(c => `<li>${c}</li>`).join('')}</ul>
      </div>`;
  }
  const emailBodyManager = `
    <p>Gandang gising!</p>
    <p>A business card request for <strong>${formData.fullName}</strong> is <strong>${statusLabel.toUpperCase()}</strong> and requires your action.</p>
    <p><strong>Request ID:</strong> ${requestID}</p>
    ${changeHtml}
    ${detailsHtml}
    <p><a href="${managerUrl}" style="padding: 10px 15px; background-color: #1877f2; color: white; text-decoration: none; border-radius: 5px;">Open Approval Dashboard</a></p>
    <p>Thank you.</p>`;
  let subject = '';
  if (statusLabel === 'New') { subject = `ACTION REQUIRED: New Business Card Request (${requestID})`; }
  else if (statusLabel === 'Pending' && changes.length > 0) { subject = `RESUBMITTED: Business Card Request (${requestID})`; }
  else { subject = `ACTION REQUIRED: ${statusLabel} Business Card Request (${requestID})`; }
  MailApp.sendEmail({to: HROD_MANAGER_EMAILS.join(','), subject: subject, htmlBody: emailBodyManager, name: SENDER_NAME});
}

function sendStatusUpdateEmail(row, status, approverEmail, reason) {
  const get = (i) => (row[i] === undefined ? '' : row[i]);
  const formData = {
    requestID: get(1), requestorName: get(2), requestorEmail: get(3), fullName: get(4),
    cardName: get(5), department: get(6), position: get(7), telephone: get(8), 
    localNumber: get(9), cellphone: get(10), email: get(11), website: get(12),
    companyName: get(13), companyAddress: get(14), requestReason: get(16)
  };
  let subject = `Update on Business Card Request (ID: ${formData.requestID})`;
  let reasonHtml = reason ? `<p style="border-left: 4px solid #fa3e3e; padding-left: 10px; background-color: #ffebee;"><strong>Reason:</strong> ${reason}</p>` : '';
  let nextSteps = '<p>Please contact Corporate HROD for more details.</p>';
  const managerUrl = `${WEB_APP_URL}?v=manager`;
  if (status === 'Approved') { nextSteps = '<p>You may now submit your PR for approval. Please attach this email as proof to your PR. Kindly coordinate with Purchasing for printing status.</p>'; }
  else if (status.includes('Returned')) {
    subject = `ACTION REQUIRED: Request Returned (${formData.requestID})`;
    const editLink = `${WEB_APP_URL}?v=edit&id=${formData.requestID}`;
    nextSteps = `<p>Please edit your request using the link below:</p><p><a href="${editLink}" style="padding: 10px 15px; background-color: #1877f2; color: white; text-decoration: none; border-radius: 5px;">Edit Request</a></p>`;
  }
  const detailsHtml = getRequestDetailsHtml(formData);
  const requestorEmailBody = `
    <p>Gandang gising!</p>
    <p>Update on the business card request for <strong>${formData.fullName}</strong> (ID: <strong>${formData.requestID}</strong>).</p>
    <p>The request has been <strong>${status.toUpperCase()}</strong> by <strong>${approverEmail}</strong>.</p>
    ${reasonHtml}
    ${detailsHtml}
    ${nextSteps}
    <p>Thank you,</p><p>Corporate HROD</p>`;
  MailApp.sendEmail({to: formData.requestorEmail, subject: subject, htmlBody: requestorEmailBody, name: SENDER_NAME});
  if (formData.requestorEmail.toLowerCase() !== formData.email.toLowerCase()) { MailApp.sendEmail({to: formData.email, subject: subject, htmlBody: requestorEmailBody, name: SENDER_NAME}); }
  const dashboardLinkHtml = `<p><a href="${managerUrl}" style="padding: 10px 15px; background-color: #1877f2; color: white; text-decoration: none; border-radius: 5px;">Open Approval Dashboard</a></p>`;
  const purchasingEmails = getPurchasingEmails(formData.companyName, formData.companyAddress);
  if (status === 'Approved') {
    const purchasingBody = `<p>Gandang gising!</p><p>The business card request for <strong>${formData.fullName}</strong> of <strong>${formData.companyName}</strong> has been <strong>APPROVED</strong> by Corporate HROD.</p><p>The requestor, ${formData.requestorName}, will coordinate with your department for the PR and printing process.</p>${detailsHtml}${dashboardLinkHtml}<p>Thank you.</p>`;
    MailApp.sendEmail({to: purchasingEmails.join(','), subject: `For Printing: Business Card Request for ${formData.fullName} (ID: ${formData.requestID})`, htmlBody: purchasingBody, name: SENDER_NAME});
  } else if (status === 'Disapproved') {
    const purchasingBody = `<p>Gandang gising!</p><p>This is to inform you that the business card request for <strong>${formData.fullName}</strong> (ID: <strong>${formData.requestID}</strong>) has been <strong>DISAPPROVED</strong>.</p>${reasonHtml}${detailsHtml}${dashboardLinkHtml}<p>No further action is required.</p><p>Thank you.</p>`;
    MailApp.sendEmail({to: purchasingEmails.join(','), subject: `Disapproved Business Card Request (ID: ${formData.requestID})`, htmlBody: purchasingBody, name: SENDER_NAME});
  }
  const otherManagers = HROD_MANAGER_EMAILS.filter(email => email.toLowerCase() !== approverEmail.toLowerCase());
  if (otherManagers.length > 0) {
    const fyiBody = `<p>Gandang gising!</p><p>FYI: The business card request for <strong>${formData.fullName}</strong> (ID: <strong>${formData.requestID}</strong>) has been <strong>${status.toUpperCase()}</strong> by <strong>${approverEmail}</strong>.</p>${reasonHtml}<p>No further action is required.</p>`;
    MailApp.sendEmail({to: otherManagers.join(','), subject: `FYI: Request ${status} by ${approverEmail} (ID: ${formData.requestID})`, htmlBody: fyiBody, name: SENDER_NAME});
  }
}

function sendPendingReminders() {
  try {
    const sheet = getSheet(SHEET_NAME);
    if (sheet.getLastRow() < 2) return;
    
    // Get all data
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 20).getValues();
    
    // Get "Start of Today" to compare dates only, ignoring time
    const today = new Date();
    today.setHours(0, 0, 0, 0); 

    const managerUrl = `${WEB_APP_URL}?v=manager`;
    const get = (row, i) => (row[i] === undefined ? '' : row[i]);

    const pendingRequests = data.filter(row => {
      const status = get(row, 17); // Status Column
      const timestampRaw = get(row, 0); // Timestamp Column
      
      if (!timestampRaw || status !== 'Pending') return false;

      const requestDate = new Date(timestampRaw);
      requestDate.setHours(0, 0, 0, 0); // Normalize to midnight

      // Logic: Send reminder if the request date is strictly BEFORE today
      // This captures everything from yesterday and older
      return requestDate < today; 
    });

    if (pendingRequests.length === 0) return;

    // Consolidate emails to prevent spamming (Optional, but recommended)
    // Or keep sending individual emails as per your previous preference:
    pendingRequests.forEach(row => {
      const requestID = get(row, 1);
      const fullName = get(row, 4);
      
      const subject = `Reminder: Pending Business Card Request (ID: ${requestID})`;
      const emailBody = `
        <p>Gandang gising!</p>
        <p>Reminder: A business card request for <strong>${fullName}</strong> (ID: <strong>${requestID}</strong>) is still awaiting approval.</p>
        <p><a href="${managerUrl}" style="padding: 10px 15px; background-color: #f7b924; color: white; text-decoration: none; border-radius: 5px;">Open Approval Dashboard</a></p>
        <p>Thank you.</p>`;
      
      MailApp.sendEmail({
        to: HROD_MANAGER_EMAILS.join(','), 
        subject: subject, 
        htmlBody: emailBody, 
        name: SENDER_NAME
      });
    });
    
    Logger.log(`Sent ${pendingRequests.length} reminders.`);

  } catch (error) { 
    Logger.log("Error sending reminders: " + error.toString()); 
  }
}
