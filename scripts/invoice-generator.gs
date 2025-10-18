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
 * 3. Update DATA_SPREADSHEET_ID, TEMPLATE_SPREADSHEET_ID, and NOTION_API_KEY below
 * 4. Run createInvoices() function
 */

// ============================================================================
// ENVIRONMENT VARIABLES - UPDATE THESE FIRST
// ============================================================================


const DATA_SPREADSHEET_ID = "";

const TEMPLATE_SPREADSHEET_ID = "";

const NOTION_API_KEY = "";

// ============================================================================
// CONFIGURATION SECTION
// ============================================================================

/**
 * The data spreadsheet should have the following columns (in any order):
 * - seller_name: Seller company name
 * - seller_address: Seller address
 * - seller_email: Email address for invoice delivery
 * - item_1_name: Item description
 * - item_1_number: Item quantity
 * - item_1_price: Item price
 * - seller_bank_name: Bank name
 * - seller_bank_type: Account type
 * - seller_bank_number: Account number
 * - seller_bank_holder_name: Account holder name
 * - output_folder_id: (MANDATORY) Google Drive Folder ID for this invoice
 *
 * Note: Invoice date and PDF filename are automatically generated when creating the PDF
 * PDF filename format: YYYYMMDD_株式会社DROX様_請求書.pdf
 */

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
// EMAIL CONFIGURATION
// ============================================================================

/**
 * Email Settings
 *
 * Configure email sending functionality
 * Set TEST_MODE to true to avoid actual email sending during testing
 */
const EMAIL_ENABLED = true; // Set to false to disable email sending completely
const TEST_MODE = false; // Set to true to only log emails without sending
const SENDER_EMAIL = "koki-hata@drox-inc.com"; // Sender email address
const SENDER_NAME = "株式会社DROX"; // Sender display name

// ============================================================================
// NOTION API CONFIGURATION
// ============================================================================

/**
 * Notion Database Configuration
 *
 * Make sure to:
 * 1. Set NOTION_API_KEY in the environment variables section above
 * 2. Share your Notion database with the integration
 */
const NOTION_DATABASE_ID = "25143e83afc180cfaa2bff97b74538eb";
const NOTION_API_VERSION = "2022-06-28";

// ============================================================================
// NOTION API FUNCTIONS
// ============================================================================

/**
 * Get Notion database schema to check property names and types
 * This helper function can be used to debug property configuration
 *
 * @returns {Object} Database schema with properties
 */
function getNotionDatabaseSchema() {
  try {
    const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}`;

    const options = {
      method: "get",
      headers: {
        "Authorization": `Bearer ${NOTION_API_KEY}`,
        "Notion-Version": NOTION_API_VERSION
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200) {
      console.log("=== Notion Database Properties ===");
      for (const [propName, propConfig] of Object.entries(data.properties)) {
        console.log(`Property: "${propName}" - Type: ${propConfig.type}`);
      }
      console.log("=================================");
      return data.properties;
    } else {
      console.error("Failed to get database schema:", data);
      return null;
    }
  } catch (error) {
    console.error("Error getting database schema:", error.message);
    return null;
  }
}

/**
 * Fetch hours from Notion database with date range filter
 * Sums up all "hours" property values from entries within the specified date range
 *
 * @param {string} startDate - Start date in YYYY-MM-DD format (e.g., "2025-10-01")
 * @param {string} endDate - End date in YYYY-MM-DD format (e.g., "2025-10-31")
 * @returns {number} Total hours summed from matching Notion entries
 */
function fetchNotionHours(startDate, endDate) {
  try {
    // Validate input parameters
    if (!startDate || !endDate) {
      console.error(`Invalid date parameters: startDate="${startDate}", endDate="${endDate}"`);
      throw new Error(`Date parameters are required. Received: startDate="${startDate}", endDate="${endDate}"`);
    }

    console.log(`Fetching hours from Notion for period: ${startDate} to ${endDate}`);

    const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`;

    // Construct the filter for date range
    // Using "and" to filter by both Start Time and End Time
    const payload = {
      filter: {
        and: [
          {
            property: "Start Time",
            date: {
              on_or_after: startDate
            }
          },
          {
            property: "End Time",
            date: {
              on_or_before: endDate
            }
          }
        ]
      },
      page_size: 100 // Maximum allowed by Notion API
    };

    // Debug: Log the filter structure
    console.log("Notion API payload:", JSON.stringify(payload, null, 2));

    const options = {
      method: "post",
      headers: {
        "Authorization": `Bearer ${NOTION_API_KEY}`,
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    let totalHours = 0;
    let hasMore = true;
    let nextCursor = null;

    // Handle pagination if there are more than 100 results
    while (hasMore) {
      if (nextCursor) {
        payload.start_cursor = nextCursor;
        options.payload = JSON.stringify(payload);
      }

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        console.error(`Notion API error: ${response.getContentText()}`);
        throw new Error(`Notion API returned error code: ${responseCode}`);
      }

      const data = JSON.parse(response.getContentText());

      // Process each entry to sum up hours
      if (data.results && data.results.length > 0) {
        data.results.forEach((page) => {
          try {
            // Access the hours property (adjust property name if different in your database)
            if (page.properties && page.properties.hours) {
              let hoursValue = 0;

              // Handle different property types (number, formula, etc.)
              if (page.properties.hours.type === "number") {
                hoursValue = page.properties.hours.number || 0;
              } else if (page.properties.hours.type === "formula") {
                hoursValue = page.properties.hours.formula.number || 0;
              }

              totalHours += hoursValue;
            }
          } catch (e) {
            console.error(`Error processing entry: ${e.message}`);
          }
        });
      }

      // Check if there are more pages
      hasMore = data.has_more || false;
      nextCursor = data.next_cursor || null;
    }

    console.log(`Total hours fetched from Notion: ${totalHours}`);
    return totalHours;

  } catch (error) {
    console.error(`Error fetching data from Notion: ${error.message}`);
    throw error;
  }
}

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
  let emailsSentCount = 0;
  let emailsFailedCount = 0;

  console.log(`Starting invoice generation for ${invoiceDataRows.length} row(s)...`);

  // --- Process invoice data row by row ---

  invoiceDataRows.forEach((invoice, index) => {
    // --- Invoice creation process ---
    // 1. Copy template sheet to create a working sheet
    // Working sheet is created in the template spreadsheet

    const copiedSheet = templateSheet.copyTo(templateSpreadsheet);

    const newSheetName = `temp_${index + 1}`; // Temporary sheet name

    copiedSheet.setName(newSheetName);

    // 2. Replace placeholders with actual data
    // Use current date and generate PDF filename

    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, "JST", "yyyy/MM/dd");
    const pdfFileName = Utilities.formatDate(currentDate, "JST", "yyyyMMdd") + "_株式会社DROX様_請求書";

    // --- Fetch hours from Notion and override item_1_number ---
    // This will replace the Google Sheets value with the summed hours from Notion
    let itemNumber = invoice.item_1_number; // Default to Google Sheets value

    try {
      // For now, using fixed dates. You can make these dynamic later
      const startDate = "2025-10-01";
      const endDate = "2025-10-31";

      const notionHours = fetchNotionHours(startDate, endDate);

      if (notionHours !== null && notionHours !== undefined) {
        itemNumber = notionHours;
        console.log(`Invoice ${index + 1}: Using Notion hours: ${notionHours}`);
      } else {
        console.log(`Invoice ${index + 1}: No Notion hours found, using Google Sheets value: ${itemNumber}`);
      }
    } catch (notionError) {
      console.error(`Invoice ${index + 1}: Failed to fetch Notion data, falling back to Google Sheets value`);
      console.error(`Error details: ${notionError.message}`);
      // Keep using the original Google Sheets value
    }

    // Replace invoice details
    copiedSheet.createTextFinder("{{invoice_date}}").replaceAllWith(formattedDate);

    // Replace seller information
    copiedSheet.createTextFinder("{{seller_name}}").replaceAllWith(invoice.seller_name);
    copiedSheet.createTextFinder("{{seller_address}}").replaceAllWith(invoice.seller_address);

    // Replace item details (using itemNumber from Notion or Google Sheets)
    copiedSheet.createTextFinder("{{item_1_name}}").replaceAllWith(invoice.item_1_name);
    copiedSheet.createTextFinder("{{item_1_number}}").replaceAllWith(itemNumber);
    copiedSheet.createTextFinder("{{item_1_price}}").replaceAllWith(invoice.item_1_price);

    // Replace bank information
    copiedSheet.createTextFinder("{{seller_bank_name}}").replaceAllWith(invoice.seller_bank_name);
    copiedSheet.createTextFinder("{{seller_bank_type}}").replaceAllWith(invoice.seller_bank_type);
    copiedSheet.createTextFinder("{{seller_bank_number}}").replaceAllWith(invoice.seller_bank_number);
    copiedSheet.createTextFinder("{{seller_bank_holder_name}}").replaceAllWith(invoice.seller_bank_holder_name);

    // 3. Flush changes to spreadsheet immediately

    SpreadsheetApp.flush();

    // 4. Validate and get output folder (MANDATORY)

    if (!invoice.output_folder_id || invoice.output_folder_id.trim() === "") {
      failureCount++;
      console.error(`Invoice ${index + 1}: Error - output_folder_id is required but not provided`);
      templateSpreadsheet.deleteSheet(copiedSheet);
      return;
    }

    let targetFolder;
    try {
      targetFolder = DriveApp.getFolderById(invoice.output_folder_id);
      console.log(`Invoice ${index + 1}: Using folder ID: ${invoice.output_folder_id}`);
    } catch (e) {
      failureCount++;
      console.error(`Invoice ${index + 1}: Error - Invalid output_folder_id: ${invoice.output_folder_id}. ${e.message}`);
      templateSpreadsheet.deleteSheet(copiedSheet);
      return;
    }

    // 5. Create PDF and save to Google Drive

    let pdfFile = null;
    try {
      pdfFile = createPdfInDrive(templateSpreadsheet, copiedSheet.getSheetId(), targetFolder, pdfFileName);
      successCount++;
      console.log(`Invoice ${index + 1}: Successfully created ${pdfFileName}.pdf`);
    } catch (e) {
      failureCount++;
      console.error(`Invoice ${index + 1}: Error creating ${pdfFileName}.pdf - ${e.message}`);
      templateSpreadsheet.deleteSheet(copiedSheet);
      return;
    }

    // 6. Send email if enabled and email address is available

    if (EMAIL_ENABLED && pdfFile) {
      const recipientEmail = invoice.seller_email;
      const recipientName = invoice.seller_name;

      if (recipientEmail && recipientEmail.includes('@')) {
        const emailSent = sendInvoiceEmail(recipientEmail, recipientName, pdfFile, formattedDate);

        if (emailSent) {
          emailsSentCount++;
          console.log(`Invoice ${index + 1}: Email sent to ${recipientEmail}`);
        } else {
          emailsFailedCount++;
          console.error(`Invoice ${index + 1}: Failed to send email to ${recipientEmail}`);
        }
      } else {
        console.warn(`Invoice ${index + 1}: No valid email address provided, skipping email sending`);
      }
    }

    // 7. Delete the temporary working sheet

    templateSpreadsheet.deleteSheet(copiedSheet);
  });

  // Log completion summary
  console.log("\n=== Invoice Generation Complete ===");
  console.log(`Total rows processed: ${invoiceDataRows.length}`);
  console.log(`✓ PDFs created: ${successCount}`);
  console.log(`✗ PDFs failed: ${failureCount}`);
  console.log(`⊘ Skipped: ${skippedCount}`);

  if (EMAIL_ENABLED) {
    console.log("\n--- Email Summary ---");
    console.log(`✓ Emails sent: ${emailsSentCount}`);
    console.log(`✗ Emails failed: ${emailsFailedCount}`);

    if (TEST_MODE) {
      console.log("※ TEST MODE was enabled - no emails were actually sent");
    }
  } else {
    console.log("\n※ Email sending is disabled");
  }

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
  const pdfFile = folder.createFile(blob);

  // Return the created PDF file for further processing
  return pdfFile;
}

// ============================================================================
// EMAIL FUNCTIONS
// ============================================================================

/**
 * Send invoice email with PDF link
 *
 * @param {string} recipientEmail - The recipient's email address
 * @param {string} recipientName - The recipient's name
 * @param {File} pdfFile - The PDF file object from Google Drive
 * @param {string} invoiceDate - The invoice date in YYYY/MM/DD format
 * @returns {boolean} True if email sent successfully, false otherwise
 */
function sendInvoiceEmail(recipientEmail, recipientName, pdfFile, invoiceDate) {
  try {
    // Validate email address
    if (!recipientEmail || !recipientEmail.includes('@')) {
      console.error(`Invalid email address: ${recipientEmail}`);
      return false;
    }

    // Get the sharing URL for the PDF file
    const pdfUrl = getPdfShareableUrl(pdfFile);

    // Extract year and month from invoice date
    const dateObj = new Date(invoiceDate);
    const year = dateObj.getFullYear();
    const month = dateObj.getMonth() + 1; // getMonth() returns 0-11

    // Email subject
    const subject = `請求書送付のご案内 [${year}年${month}月分]`;

    // Email body (Japanese business email format)
    const body = createEmailBody(recipientName, year, month, pdfUrl);

    // Prepare email options
    const options = {
      name: SENDER_NAME,
      htmlBody: body,
      from: SENDER_EMAIL
    };

    // Send email or log if in test mode
    if (TEST_MODE) {
      console.log("=== TEST MODE: Email would be sent ===");
      console.log(`To: ${recipientEmail}`);
      console.log(`Subject: ${subject}`);
      console.log(`Body: ${body}`);
      console.log("=====================================");
      return true;
    } else {
      GmailApp.sendEmail(recipientEmail, subject, body, options);
      console.log(`Email sent successfully to: ${recipientEmail}`);
      return true;
    }

  } catch (error) {
    console.error(`Failed to send email to ${recipientEmail}: ${error.message}`);
    return false;
  }
}

/**
 * Generate shareable URL for PDF file in Google Drive
 *
 * @param {File} pdfFile - The PDF file object
 * @returns {string} Shareable URL for the PDF
 */
function getPdfShareableUrl(pdfFile) {
  try {
    // Set sharing permissions - Anyone with the link can view
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Get the download URL
    const fileId = pdfFile.getId();
    const shareableUrl = `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;

    return shareableUrl;
  } catch (error) {
    console.error(`Failed to generate shareable URL: ${error.message}`);
    throw error;
  }
}

/**
 * Create email body with Japanese business format
 *
 * @param {string} recipientName - The recipient's name
 * @param {number} year - Invoice year
 * @param {number} month - Invoice month
 * @param {string} pdfUrl - URL to the PDF file
 * @returns {string} Formatted email body in HTML
 */
function createEmailBody(recipientName, year, month, pdfUrl) {
  const htmlBody = `
    <div style="font-family: 'Hiragino Sans', 'Meiryo', sans-serif; line-height: 1.6; color: #333;">
      <p>${recipientName} 様</p>

      <p>いつもお世話になっております。<br>
      株式会社DROXです。</p>

      <p>${year}年${month}月分の請求書をお送りいたします。<br>
      下記リンクよりご確認ください。</p>

      <p style="margin: 20px 0;">
        <a href="${pdfUrl}" style="background-color: #4CAF50; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">
          請求書を確認する
        </a>
      </p>

      <p><strong>請求書リンク：</strong><br>
      <a href="${pdfUrl}" style="color: #0066cc;">${pdfUrl}</a></p>

      <p style="margin-top: 30px;">
      ご不明な点がございましたら、お気軽にお問い合わせください。<br>
      今後ともよろしくお願いいたします。</p>

      <hr style="margin: 30px 0; border: none; border-top: 1px solid #ddd;">

      <p style="font-size: 12px; color: #666;">
      株式会社DROX<br>
      ${SENDER_EMAIL}<br>
      ※このメールは自動送信されています。
      </p>
    </div>
  `;

  return htmlBody;
}
