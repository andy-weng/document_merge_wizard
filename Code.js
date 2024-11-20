// Code.gs

/**
 * Adds a custom menu to the Google Sheet upon opening.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Document Merge Tools')
    .addItem('Start Merge Wizard', 'showMergeSidebar')
    .addToUi();
}

/**
 * Displays the custom sidebar for the merge wizard.
 */
function showMergeSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Document Merge Wizard')
    .setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Retrieves a list of Gmail draft subjects and their IDs for the dropdown.
 * @returns {Array} Array of objects containing draft IDs and subjects.
 */
function getGmailDrafts() {
  const drafts = GmailApp.getDrafts();
  return drafts.map(draft => {
    const message = draft.getMessage();
    return {
      id: draft.getId(),
      subject: message.getSubject()
    };
  });
}

/**
 * Retrieves the column headers from the active sheet.
 * @returns {Array} Array of column header names.
 */
function getSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getDataRange().getValues()[0];
  return headers;
}

/**
 * Counts the number of rows that have data to be processed.
 * @returns {number} The number of rows to be processed.
 */
function getRowsToProcess() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const rowsToProcess = rows.filter(row => row.some(cell => cell.toString().trim() !== '')).length;
  return rowsToProcess;
}

/**
 * Sanitizes the filename by replacing illegal characters with underscores.
 * @param {string} name - The original filename.
 * @returns {string} The sanitized filename.
 */
function sanitizeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * Executes the merge and email/save actions based on user input.
 * @param {Object} options - The options selected by the user.
 * @returns {Object} Result object indicating success or failure.
 */
function executeMergeAction(options) {
  try {
    const startTime = new Date();
    // Destructure options
    const {
      action,
      templateLink,
      folderLink,
      emailDraftId,
      testMode,
      testEmail,
      emailField,
      ccField,
      bccField,
      attachmentsField,
      filenameFormat
    } = options;

    // Extract Document ID from Template Link
    const docId = extractDocumentId(templateLink);
    if (!docId) {
      throw new Error('Invalid Template Document URL.');
    }

    // Extract Folder ID from Folder Link
    const folderId = extractFolderId(folderLink);
    if (!folderId) {
      throw new Error('Invalid Destination Folder URL.');
    }

    // Validate inputs
    if (!action || !docId || !folderId || !filenameFormat) {
      throw new Error('Please provide all required fields.');
    }

    if (action === 'email' && !emailDraftId) {
      throw new Error('Please select an email draft.');
    }

    // Get the active sheet and data
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0];
    const rows = data.slice(1);

    // Convert column letters to indices (0-based)
    const emailIndex = columnLetterToIndex(emailField);
    const ccIndex = ccField ? columnLetterToIndex(ccField) : -1;
    const bccIndex = bccField ? columnLetterToIndex(bccField) : -1;
    const attachmentsIndex = attachmentsField ? columnLetterToIndex(attachmentsField) : -1;

    if (action === 'email') {
      if (emailIndex === -1 || emailIndex >= headers.length) {
        throw new Error(`Email column "${emailField}" is out of range.`);
      }
    }

    // Access the template document
    const templateDoc = DriveApp.getFileById(docId);
    if (!templateDoc) {
      throw new Error('Template document not found. Please check the Template URL.');
    }

    // Access the destination folder
    const destFolder = DriveApp.getFolderById(folderId);
    if (!destFolder) {
      throw new Error('Destination folder not found. Please check the Folder URL.');
    }

    // Retrieve the selected email draft if action is 'email'
    let emailDraft = null;
    if (action === 'email') {
      const drafts = GmailApp.getDrafts();
      const selectedDraft = drafts.find(draft => draft.getId() === emailDraftId);
      if (!selectedDraft) {
        throw new Error('Selected email draft not found.');
      }
      emailDraft = selectedDraft.getMessage();
    }

    // Estimate Gmail sending limits
    const remainingQuota = getRemainingGmailQuota();
    if (action === 'email' && !testMode) {
      if (rows.length > remainingQuota) {
        throw new Error(`You are about to send ${rows.length} emails, but only ${remainingQuota} remain for today. Please reduce the number or wait for the quota to reset.`);
      }
    }

    // Count the number of rows to be processed
    const rowsToProcess = rows.filter(row => row.some(cell => cell.toString().trim() !== '')).length;

    // Loop through each row and perform merge
    const mergeResults = [];
    rows.forEach((row, index) => {
      // Check if the row has data in any of the specified columns to determine if it's populated
      if (!row.some(cell => cell.toString().trim() !== '')) {
        return; // Skip empty rows
      }

      const rowData = {};
      headers.forEach((header, i) => {
        rowData[header] = row[i];
      });

      try {
        // Generate file name based on user-defined format
        let fileName = filenameFormat;
        const placeholderRegex = /{{\s*([^}]+)\s*}}/g;
        let match;
        while ((match = placeholderRegex.exec(filenameFormat)) !== null) {
          const placeholder = match[1];
          if (placeholder.toLowerCase() === "today's date") {
            const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
            fileName = fileName.replace(match[0], today);
          } else if (rowData.hasOwnProperty(placeholder)) {
            fileName = fileName.replace(match[0], rowData[placeholder]);
          } else {
            throw new Error(`Placeholder "{{${placeholder}}}" does not match any column headers.`);
          }
        }

        // Sanitize the filename to remove illegal characters
        fileName = sanitizeFileName(fileName);

        // Log the processing of the current row
        Logger.log(`Processing Row ${index + 2}: Filename - ${fileName}.pdf`);

        // Create a copy of the template in the destination folder
        const copy = templateDoc.makeCopy(`${fileName}`, destFolder);
        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();

        // Replace placeholders in the document
        for (const key in rowData) {
          body.replaceText(`{{${key}}}`, rowData[key]);
        }
        doc.saveAndClose();

        // Log the completion of placeholder replacement
        Logger.log(`Completed placeholder replacement for Row ${index + 2}`);

        // Optionally send email
        if (action === 'email') {
          let emailAddress = row[emailIndex];
          let ccAddress = ccIndex !== -1 ? row[ccIndex] : '';
          let bccAddress = bccIndex !== -1 ? row[bccIndex] : '';

          if (testMode && index === 0) {
            emailAddress = testEmail;
            ccAddress = '';
            bccAddress = '';
          }

          if (emailAddress) {
            const subjectTemplate = emailDraft.getSubject();
            let subject = subjectTemplate;
            // Replace placeholders in the subject
            for (const key in rowData) {
              subject = subject.replace(new RegExp(`{{${key}}}`, 'g'), rowData[key]);
            }

            let bodyContent = emailDraft.getBody();
            // Replace placeholders in the email body
            for (const key in rowData) {
              bodyContent = bodyContent.replace(new RegExp(`{{${key}}}`, 'g'), rowData[key]);
            }

            // Handle attachments
            let attachments = [];
            if (attachmentsIndex !== -1 && row[attachmentsIndex]) {
              const attachmentLinks = row[attachmentsIndex].split(',').map(link => link.trim());
              attachmentLinks.forEach(link => {
                const fileId = extractFileIdFromUrl(link);
                if (fileId) {
                  const file = DriveApp.getFileById(fileId);
                  if (file) {
                    attachments.push(file.getBlob());
                  }
                }
              });
            }

            // Attach the merged document as PDF if emailing
            const pdf = copy.getAs('application/pdf');
            attachments.push(pdf);

            // Send the email
            GmailApp.sendEmail(emailAddress, subject, '', {
              htmlBody: bodyContent,
              attachments: attachments,
              cc: ccAddress,
              bcc: bccAddress
            });

            mergeResults.push({
              row: index + 2,
              status: 'Success',
              emailSentTo: String(emailAddress)
            });

            // Log the successful email sending
            Logger.log(`Email sent to ${emailAddress} for Row ${index + 2}`);
          } else {
            mergeResults.push({
              row: index + 2,
              status: 'Error: Email address missing.',
              emailSentTo: 'N/A'
            });

            // Log the missing email address
            Logger.log(`Missing email address for Row ${index + 2}`);
          }
        } else {
          // Save as PDF if action is 'save'
          Logger.log(`Exporting PDF for Row ${index + 2}`);
          const pdfBlob = Drive.Files.export(copy.getId(), 'application/pdf');
          const pdfFile = destFolder.createFile(pdfBlob, `${fileName}.pdf`);
          
          // Log the PDF blob size
          Logger.log(`PDF Blob Size for Row ${index + 2}: ${pdfBlob.getBytes().length} bytes`);

          // Check if PDF was created successfully
          if (pdfFile && pdfFile.getId()) {
            mergeResults.push({
              row: index + 2,
              status: 'Saved as PDF',
              fileName: `${fileName}.pdf`
            });

            // Log the successful PDF saving
            Logger.log(`Successfully saved PDF: ${pdfFile.getName()} (ID: ${pdfFile.getId()}) for Row ${index + 2}`);
          } else {
            mergeResults.push({
              row: index + 2,
              status: 'Error: PDF not saved.',
              fileName: 'N/A'
            });

            // Log the failure to save PDF
            Logger.log(`Failed to save PDF for Row ${index + 2}`);
          }

          // Temporarily Comment Out Deletion to Verify PDF Creation
          // copy.setTrashed(true);
        }
      } catch (rowError) {
        mergeResults.push({
          row: index + 2,
          status: `Error: ${rowError.message}`,
          emailSentTo: 'N/A'
        });

        // Log the row-specific error
        Logger.log(`Error in Row ${index + 2}: ${rowError.message}`);
      }
    });

    // Summary Report Generation Removed

    return { success: true, results: mergeResults /*, summary: summaryFile.getUrl() */ };
  } catch (error) {
    Logger.log(`General Error: ${error.message}`);
    return { success: false, message: error.message };
  }
}

/**
 * Extracts the Document ID from a Google Docs URL.
 * @param {string} url - The URL of the Google Docs document.
 * @returns {string|null} The Document ID or null if not found.
 */
function extractDocumentId(url) {
  const regex = /\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

/**
 * Extracts the Folder ID from a Google Drive Folder URL.
 * @param {string} url - The URL of the Google Drive folder.
 * @returns {string|null} The Folder ID or null if not found.
 */
function extractFolderId(url) {
  const regex = /\/folders\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

/**
 * Extracts the file ID from a Google Drive URL.
 * @param {string} url - The URL of the Google Drive file.
 * @returns {string|null} The file ID or null if not found.
 */
function extractFileIdFromUrl(url) {
  const regex = /[-\w]{25,}/;
  const match = url.match(regex);
  return match ? match[0] : null;
}

/**
 * Converts a column letter to a 0-based index.
 * @param {string} letter - The column letter (e.g., 'A').
 * @returns {number} The 0-based column index.
 */
function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    const charCode = letter.toUpperCase().charCodeAt(i);
    if (charCode < 65 || charCode > 90) return -1; // Invalid character
    column = column * 26 + (charCode - 64);
  }
  return column - 1;
}

/**
 * Estimates the remaining Gmail sending quota for the day.
 * Note: This is a simple estimation and may not be accurate.
 * @returns {number} The estimated remaining emails that can be sent.
 */
function getRemainingGmailQuota() {
  // Google doesn't provide a direct way to get remaining quota.
  // This function can be expanded with heuristics or user input.
  // For now, we return a high number to allow most operations.
  return 1000;
}