/**
 * Google Apps Script - Invoice PDF Generator
 *
 * This script automatically generates invoice PDFs from spreadsheet data.
 *
 * Features:
 * - Header-based data mapping (column order independent)
 * - Automatic date formatting
 * - Batch processing
 * - Error handling
 *
 * Setup Instructions:
 * 1. Update the configuration section below with your IDs
 * 2. Ensure your data spreadsheet has headers in Row 1
 * 3. Use {{header_name}} placeholders in your template
 * 4. Run createInvoices() function
 */

// ============================================================================
// CONFIGURATION SECTION
// Update these values with your own spreadsheet and folder IDs
// ============================================================================

/**
 * Spreadsheet ID containing invoice data
 * Find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 */
const DATA_SPREADSHEET_ID = "YOUR_DATA_SPREADSHEET_ID";

/**
 * Spreadsheet ID containing invoice template
 * Find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 */
const TEMPLATE_SPREADSHEET_ID = "YOUR_TEMPLATE_SPREADSHEET_ID";

/**
 * Google Drive Folder ID where PDFs will be saved
 * Find it in the URL: https://drive.google.com/drive/folders/FOLDER_ID
 */
const OUTPUT_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID";

/**
 * Name of the sheet containing invoice data in the data spreadsheet
 * Default is usually 'Sheet1'
 */
const DATA_SHEET_NAME = "Sheet1";

/**
 * Name of the sheet containing the invoice template
 * Default is usually 'Sheet1'
 */
const TEMPLATE_SHEET_NAME = "Sheet1";

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

function createInvoices() {
  // --- Get spreadsheets and folder ---

  const dataSpreadsheet = SpreadsheetApp.openById(DATA_SPREADSHEET_ID);

  const dataSheet = dataSpreadsheet.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    console.error(`Data sheet "${DATA_SHEET_NAME}" not found.`);
    return;
  }

  const templateSpreadsheet = SpreadsheetApp.openById(TEMPLATE_SPREADSHEET_ID);

  const templateSheet = templateSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);

  if (!templateSheet) {
    console.error(`Template sheet "${TEMPLATE_SHEET_NAME}" not found.`);
    return;
  }

  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID); // --- Get invoice data from data sheet --- // getValues() retrieves all sheet data as a 2D array

  const data = dataSheet.getDataRange().getValues(); // Exclude header row (row 1) and start processing from row 2

  const invoiceDataRows = data.slice(1);

  // Track success and failure counts
  let successCount = 0;
  let failureCount = 0;
  let skippedCount = 0;

  console.log(`Starting invoice generation for ${invoiceDataRows.length} row(s)...`);

  // --- Process invoice data row by row ---

  invoiceDataRows.forEach((row, index) => {
    // Store data from columns A to F in respective variables

    const [
      invoiceDate, // Invoice Date (Column A)
      invoiceNumber, // Invoice Number (Column B)
      companyName, // Company Name (Column C)
      subject, // Subject (Column D)
      amount, // Amount (Column E)
      pdfFileName, // PDF Filename (Column F)
    ] = row; // Skip rows where company name is empty

    if (!companyName) {
      skippedCount++;
      console.log(`Row ${index + 2}: Skipped (no company name)`);
      return;
    } // --- Invoice creation process --- // 1. Copy template sheet to create a working sheet // Working sheet is created in the data spreadsheet

    const copiedSheet = templateSheet.copyTo(dataSpreadsheet);

    const newSheetName = `temp_${invoiceNumber}`; // Temporary sheet name

    copiedSheet.setName(newSheetName); // 2. Replace placeholders with actual data // Format date to 'yyyy/MM/dd' format

    const formattedDate = Utilities.formatDate(new Date(invoiceDate), "JST", "yyyy/MM/dd");

    copiedSheet.createTextFinder("{{請求日}}").replaceAllWith(formattedDate);

    copiedSheet.createTextFinder("{{請求書番号}}").replaceAllWith(invoiceNumber);

    copiedSheet.createTextFinder("{{会社名}}").replaceAllWith(companyName);

    copiedSheet.createTextFinder("{{件名}}").replaceAllWith(subject); // Amount is transferred as a number (template formatting will be applied)

    copiedSheet.createTextFinder("{{金額}}").replaceAllWith(amount); // 3. Flush changes to spreadsheet immediately

    SpreadsheetApp.flush(); // 4. Create PDF and save to Google Drive

    try {
      createPdfInDrive(dataSpreadsheet, copiedSheet.getSheetId(), outputFolder, pdfFileName);
      successCount++;
      console.log(`Row ${index + 2}: Successfully created ${pdfFileName}.pdf`);
    } catch (e) {
      failureCount++;
      console.error(`Row ${index + 2}: Error creating ${pdfFileName}.pdf - ${e.message}`);
    } // 5. Delete the temporary working sheet

    dataSpreadsheet.deleteSheet(copiedSheet);
  });

  // Log completion summary
  console.log("\n=== Invoice Generation Complete ===");
  console.log(`Total rows processed: ${invoiceDataRows.length}`);
  console.log(`✓ Successful: ${successCount}`);
  console.log(`✗ Failed: ${failureCount}`);
  console.log(`⊘ Skipped: ${skippedCount}`);
  console.log("===================================");
}

/**

 * Function to save a specified sheet as PDF to Google Drive

 * @param {Spreadsheet} spreadsheet - The spreadsheet containing the sheet to convert to PDF

 * @param {string} sheetId - The ID of the sheet to convert to PDF

 * @param {Folder} folder - The Google Drive folder where the PDF will be saved

 * @param {string} fileName - The filename for the PDF (without extension)

 */
function createPdfInDrive(spreadsheet, sheetId, folder, fileName) {
  const spreadsheetId = spreadsheet.getId();

  // Create URL for PDF export
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    "format=pdf" +
    "&gid=" +
    sheetId + // Specify sheet ID
    "&size=A4" + // Paper size (A3/A4/A5/letter/legal, etc.)
    "&portrait=true" + // Portrait orientation (false = landscape)
    "&scale=3" + // Scale setting:
    // 1 = Normal (100%)
    // 2 = Fit to width
    // 3 = Fit to height ★ Recommended
    // 4 = Fit to page
    "&top_margin=0.2" + // Top margin (inches)
    "&bottom_margin=0.2" + // Bottom margin (inches)
    "&left_margin=0.2" + // Left margin (inches)
    "&right_margin=0.2" + // Right margin (inches)
    "&sheetnames=false" + // Hide sheet name
    "&printtitle=false" + // Hide spreadsheet name
    "&gridlines=false" + // Hide grid lines
    "&fzr=false" + // Ignore frozen rows
    "&horizontal_alignment=CENTER" + // Horizontal alignment (LEFT/CENTER/RIGHT)
    "&vertical_alignment=TOP"; // Vertical alignment (TOP/MIDDLE/BOTTOM)

  const token = ScriptApp.getOAuthToken();
  const options = {
    headers: {
      Authorization: "Bearer " + token,
    },
    muteHttpExceptions: true,
  };

  // Fetch PDF data from URL
  const response = UrlFetchApp.fetch(url, options);
  const blob = response.getBlob().setName(fileName + ".pdf");

  // Create PDF file in folder
  folder.createFile(blob);
}
