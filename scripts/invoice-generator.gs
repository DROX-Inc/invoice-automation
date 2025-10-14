/**
 * Google Apps Script - Invoice PDF Generator
 *
 * This script automatically generates invoice PDFs from Google Sheets data.
 *
 * Features:
 * - Header-based data mapping (column order independent)
 * - Automatic date formatting
 * - Batch processing
 * - Error handling
 * - Dynamic seller information support
 *
 * Setup Instructions:
 * 1. Create a data spreadsheet with invoice information (see invoice-data.csv for structure)
 * 2. Create a template spreadsheet with placeholders (see invoice-template.csv)
 * 3. Update DATA_SPREADSHEET_ID, TEMPLATE_SPREADSHEET_ID, and OUTPUT_FOLDER_ID
 * 4. Run createInvoices() function
 */

// ============================================================================
// CONFIGURATION SECTION
// Update these values with your own spreadsheet and folder IDs
// ============================================================================

/**
 * Spreadsheet ID containing invoice data
 * Find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 *
 * The data spreadsheet should have the following columns (in any order):
 * - invoice_date: Invoice date
 * - seller_name: Seller company name
 * - seller_address: Seller address
 * - seller_phone_number: Seller phone
 * - item_1_name: Item description
 * - item_1_number: Item quantity
 * - item_1_price: Item price
 * - seller_bank_name: Bank name
 * - seller_bank_type: Account type
 * - seller_bank_number: Account number
 * - seller_bank_holder_name: Account holder name
 * - pdfFileName: PDF filename (without extension)
 */
const DATA_SPREADSHEET_ID = "YOUR_DATA_SPREADSHEET_ID";

/**
 * Name of the sheet containing invoice data in the data spreadsheet
 * Default is usually 'Sheet1'
 */
const DATA_SHEET_NAME = "Sheet1";

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
 * Name of the sheet containing the invoice template
 * Default is usually 'Sheet1'
 */
const TEMPLATE_SHEET_NAME = "Sheet1";

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

function createInvoices() {
  // --- Get data spreadsheet, template spreadsheet, and output folder ---

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

  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);

  // --- Read invoice data from spreadsheet ---
  const data = dataSheet.getDataRange().getValues();

  // First row contains headers
  const headers = data[0];
  const dataRows = data.slice(1);

  // Convert rows to invoice objects based on headers
  const invoiceDataRows = dataRows.map(row => {
    const invoice = {};
    headers.forEach((header, index) => {
      invoice[header] = row[index];
    });
    return invoice;
  });

  // Track success and failure counts
  let successCount = 0;
  let failureCount = 0;
  let skippedCount = 0;

  console.log(`Starting invoice generation for ${invoiceDataRows.length} row(s)...`);

  // --- Process invoice data row by row ---

  invoiceDataRows.forEach((invoice, index) => {
    // Skip invoices without required data
    if (!invoice.pdfFileName) {
      skippedCount++;
      console.log(`Invoice ${index + 1}: Skipped (no PDF filename)`);
      return;
    }

    // --- Invoice creation process ---
    // 1. Copy template sheet to create a working sheet
    // Working sheet is created in the template spreadsheet

    const copiedSheet = templateSheet.copyTo(templateSpreadsheet);

    const newSheetName = `temp_${index + 1}`; // Temporary sheet name

    copiedSheet.setName(newSheetName);

    // 2. Replace placeholders with actual data
    // Format date to 'yyyy/MM/dd' format

    const formattedDate = Utilities.formatDate(new Date(invoice.invoice_date), "JST", "yyyy/MM/dd");

    // Replace invoice details
    copiedSheet.createTextFinder("{{invoice_date}}").replaceAllWith(formattedDate);

    // Replace seller information
    copiedSheet.createTextFinder("{{seller_name}}").replaceAllWith(invoice.seller_name);
    copiedSheet.createTextFinder("{{seller_address}}").replaceAllWith(invoice.seller_address);
    copiedSheet.createTextFinder("{{seller_phone_number}}").replaceAllWith(invoice.seller_phone_number);

    // Replace item details
    copiedSheet.createTextFinder("{{item_1_name}}").replaceAllWith(invoice.item_1_name);
    copiedSheet.createTextFinder("{{item_1_number}}").replaceAllWith(invoice.item_1_number);
    copiedSheet.createTextFinder("{{item_1_price}}").replaceAllWith(invoice.item_1_price);

    // Replace bank information
    copiedSheet.createTextFinder("{{seller_bank_name}}").replaceAllWith(invoice.seller_bank_name);
    copiedSheet.createTextFinder("{{seller_bank_type}}").replaceAllWith(invoice.seller_bank_type);
    copiedSheet.createTextFinder("{{seller_bank_number}}").replaceAllWith(invoice.seller_bank_number);
    copiedSheet.createTextFinder("{{seller_bank_holder_name}}").replaceAllWith(invoice.seller_bank_holder_name);

    // 3. Flush changes to spreadsheet immediately

    SpreadsheetApp.flush();

    // 4. Create PDF and save to Google Drive

    try {
      createPdfInDrive(templateSpreadsheet, copiedSheet.getSheetId(), outputFolder, invoice.pdfFileName);
      successCount++;
      console.log(`Invoice ${index + 1}: Successfully created ${invoice.pdfFileName}.pdf`);
    } catch (e) {
      failureCount++;
      console.error(`Invoice ${index + 1}: Error creating ${invoice.pdfFileName}.pdf - ${e.message}`);
    }

    // 5. Delete the temporary working sheet

    templateSpreadsheet.deleteSheet(copiedSheet);
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
