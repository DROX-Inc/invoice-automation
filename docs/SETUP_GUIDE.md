# Google Apps Script Invoice PDF Generator - Setup Guide

This guide provides step-by-step instructions for setting up an automated invoice PDF generation system using Google Apps Script. The script reads data from a Google Spreadsheet, applies it to an invoice template, and generates PDF files automatically.

## Table of Contents

1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Step 1: Prepare Your Files](#step-1-prepare-your-files)
4. [Step 2: Set Up Google Apps Script](#step-2-set-up-google-apps-script)
5. [Step 3: Configure the Script](#step-3-configure-the-script)
6. [Step 4: Grant Permissions](#step-4-grant-permissions)
7. [Step 5: Run and Verify](#step-5-run-and-verify)
8. [Troubleshooting](#troubleshooting)
9. [Advanced Usage](#advanced-usage)

---

## Overview

This automation system consists of three main components:

1. **Data Spreadsheet**: Contains invoice data (one invoice per row)
2. **Template Spreadsheet**: Contains the invoice design with placeholders
3. **Google Apps Script**: Automates the process of merging data into templates and generating PDFs

### How It Works

1. The script reads invoice data from the data spreadsheet
2. For each row of data, it creates a copy of the template
3. It replaces placeholders (like `{{company_name}}`) with actual data
4. It converts the populated template to PDF
5. The PDF is saved to a specified Google Drive folder
6. The temporary working sheet is deleted

### Key Features

- **Header-based data mapping**: Uses column headers to match data, so column order doesn't matter
- **Flexible template design**: Customize your invoice layout freely
- **Batch processing**: Generate multiple invoices in one execution
- **Date formatting**: Automatically formats dates correctly
- **Error handling**: Skips rows with missing critical data

---

## Prerequisites

Before you begin, ensure you have:

- A Google account
- Basic familiarity with Google Sheets
- Access to Google Drive
- Permission to create and run Google Apps Scripts

---

## Step 1: Prepare Your Files

### 1.1 Create the Data Spreadsheet

1. Create a new Google Spreadsheet
2. Name it something descriptive (e.g., "Invoice Data" or "demo-data")
3. Set up your data with headers in **Row 1**

**Example structure:**

| Invoice Date | Invoice Number | Company Name | Description | Amount | PDF Filename |
|--------------|----------------|--------------|-------------|--------|--------------|
| 2025/10/01   | INV-001        | ABC Corp     | Web Development | 100000 | 20251001_ABC_Corp |
| 2025/10/02   | INV-002        | XYZ Ltd      | Design Services | 50000  | 20251002_XYZ_Ltd |

**Important Notes:**

- **Row 1 must contain headers** - these will be used as placeholder names
- **Amount column**: Enter numbers only (e.g., `100000` not `100,000`)
- **Header names**: Must match exactly with placeholders in your template
- **Keep the spreadsheet URL handy** - you'll need the Spreadsheet ID later

**To find your Spreadsheet ID:**
```
URL format: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
Example:    https://docs.google.com/spreadsheets/d/1A2B3C4D5E6F7G8H9I0J/edit
                                                    └─────────┬────────┘
                                                    This is your ID
```

### 1.2 Create the Template Spreadsheet

1. Create another new Google Spreadsheet
2. Name it (e.g., "Invoice Template" or "invoice-template")
3. Design your invoice layout
4. Use placeholders in double curly braces: `{{header_name}}`

**Example template structure:**

```
Row 1:                         INVOICE

Row 3:  {{Company Name}}    To                      Invoice #: {{Invoice Number}}

Row 6:  Date: {{Invoice Date}}                      Your Company Name
                                                     Tax ID: T1234567890
Row 8:  Please find our invoice details below:

Row 14: Item    Description         Qty    Price      Total
Row 15: 1       {{Description}}     1      {{Amount}} =E15*D15

Row 34:                                     Subtotal:  =E15
Row 35: Bank Details:                       Tax (10%): =G34*0.1
Row 36: Bank Name: Example Bank             Total:     =G34+G35
        Account: 1234567890
```

**Important Notes:**

- **Placeholders must match headers** in your data spreadsheet exactly
  - If your data has "Company Name", use `{{Company Name}}`
  - If your data has "company_name", use `{{company_name}}`
- **Formulas work**: Use Excel/Sheets formulas for calculations
- **Design freely**: Style your invoice however you like
- **Save the Spreadsheet ID** from the URL

### 1.3 Create a Google Drive Folder

1. Go to Google Drive
2. Create a new folder (e.g., "Generated Invoices")
3. Open the folder
4. Copy the Folder ID from the URL

**To find your Folder ID:**
```
URL format: https://drive.google.com/drive/folders/FOLDER_ID_HERE
Example:    https://drive.google.com/drive/folders/1A2B3C4D5E6F7G8H9I0J
                                                    └─────────┬────────┘
                                                    This is your Folder ID
```

---

## Step 2: Set Up Google Apps Script

1. Open your **Data Spreadsheet** (the one with invoice data)
2. Click on **Extensions** in the menu bar
3. Select **Apps Script**
4. A new tab will open with the Google Apps Script editor
5. Delete any default code in the editor

---

## Step 3: Configure the Script

1. Copy the script from the `scripts/invoice-generator.gs` file in this repository
2. Paste it into the Apps Script editor
3. Update the configuration section at the top of the script:

```javascript
// --- Configuration ---
const DATA_SPREADSHEET_ID = 'YOUR_DATA_SPREADSHEET_ID';     // From Step 1.1
const TEMPLATE_SPREADSHEET_ID = 'YOUR_TEMPLATE_SPREADSHEET_ID'; // From Step 1.2
const OUTPUT_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID';            // From Step 1.3
const DATA_SHEET_NAME = 'Sheet1';        // Name of the sheet in your data spreadsheet
const TEMPLATE_SHEET_NAME = 'Sheet1';    // Name of the sheet in your template
```

4. Replace the placeholder IDs with your actual IDs:
   - Replace `YOUR_DATA_SPREADSHEET_ID` with the ID from Step 1.1
   - Replace `YOUR_TEMPLATE_SPREADSHEET_ID` with the ID from Step 1.2
   - Replace `YOUR_DRIVE_FOLDER_ID` with the ID from Step 1.3

5. Update sheet names if different from "Sheet1"

6. Click the **Save** icon (💾) or press `Ctrl+S` (Windows) / `Cmd+S` (Mac)

7. Give your project a name when prompted (e.g., "Invoice PDF Generator")

---

## Step 4: Grant Permissions

On first execution, Google needs your permission to access your spreadsheets and Drive.

1. In the Apps Script editor, ensure `createInvoices` is selected in the function dropdown
2. Click the **Run** button (▶️)
3. A dialog will appear: "Authorization required"
4. Click **Review permissions**
5. Select your Google account
6. You may see "This app isn't verified" - this is normal for personal scripts
   - Click **Advanced**
   - Click **Go to [Project Name] (unsafe)**
7. Review the permissions requested:
   - View and manage your spreadsheets
   - View and manage your Google Drive files
8. Click **Allow**

The script will now execute for the first time.

---

## Step 5: Run and Verify

### Running the Script

1. In the Apps Script editor, click the **Run** button (▶️)
2. The script will process each row in your data spreadsheet
3. You'll see a completion message when finished

### Verifying the Output

1. Go to your Google Drive folder (from Step 1.3)
2. You should see PDF files named according to your "PDF Filename" column
3. Open a PDF to verify:
   - All placeholders have been replaced with actual data
   - Formulas have calculated correctly
   - Layout looks as expected

### Expected Results

If your data spreadsheet has 2 rows of data (excluding the header), you should see 2 PDF files:
- `20251001_ABC_Corp.pdf`
- `20251002_XYZ_Ltd.pdf`

---

## Troubleshooting

### Common Issues and Solutions

#### Issue: "Sheet not found" error

**Cause**: The sheet name in the configuration doesn't match the actual sheet name

**Solution**:
1. Check the exact name of your sheets (they're case-sensitive)
2. Update `DATA_SHEET_NAME` and `TEMPLATE_SHEET_NAME` in the configuration
3. Remember: The default is usually "Sheet1" in English or "シート1" in Japanese

#### Issue: Placeholders not being replaced

**Cause**: Mismatch between header names and placeholder names

**Solution**:
1. Open your data spreadsheet and check the exact header text in Row 1
2. Open your template and verify placeholders match exactly
3. Example: If your header is "Company Name", use `{{Company Name}}` (not `{{company_name}}` or `{{CompanyName}}`)

#### Issue: Date appears as a number

**Cause**: Date formatting needs to be specified

**Solution**: The script includes automatic date formatting for the "Invoice Date" column. If you have other date columns, modify the script to format them similarly:

```javascript
if (header === '請求日' && value instanceof Date) {
  value = Utilities.formatDate(value, 'JST', 'yyyy/MM/dd');
}
```

#### Issue: Permission denied errors

**Cause**: Script doesn't have necessary permissions

**Solution**:
1. Go to Apps Script editor
2. Click **Run** → **Run function** → `createInvoices`
3. Follow the authorization process again
4. If issues persist, check your Google Account security settings

#### Issue: PDFs are blank or incomplete

**Cause**: Script is creating PDFs before spreadsheet changes are saved

**Solution**: The script includes `SpreadsheetApp.flush()` to ensure changes are saved. If issues persist:
1. Add a small delay: `Utilities.sleep(1000);` after the flush
2. Check that your template formulas are correct

#### Issue: Some invoices are skipped

**Cause**: Missing required data (likely the company name)

**Solution**: The script skips rows where the company name is empty. Check your data for:
1. Empty cells in critical columns
2. Spaces or non-visible characters
3. Verify data starts from Row 2 (Row 1 is headers)

### Checking Execution Logs

To see detailed information about what the script is doing:

1. In Apps Script editor, click **View** → **Logs** (or **Execution log**)
2. Run the script
3. Check the logs for error messages or warnings
4. Console.log statements in the script will appear here

---

## Advanced Usage

### Adding New Data Columns

The script is designed to be flexible with columns:

1. Add a new column header in your data spreadsheet (e.g., "Notes")
2. Add the corresponding placeholder in your template: `{{Notes}}`
3. No script changes needed - it will automatically map the data

### Customizing PDF Export Settings

The `createPdfInDrive` function includes PDF export parameters. You can customize:

```javascript
const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
  'format=pdf' +
  '&gid=' + sheetId +
  '&portrait=true' +        // Change to false for landscape
  '&fitw=true' +            // Fit to page width
  '&sheetnames=false' +     // Hide sheet name
  '&printtitle=false' +     // Hide spreadsheet title
  '&gridlines=false' +      // Hide gridlines
  '&fzr=false';             // Ignore frozen rows
```

### Batch Processing Tips

For large batches:

1. **Test with a few rows first**: Add only 2-3 rows of data initially
2. **Monitor execution time**: Scripts have a 6-minute execution limit
3. **Break into chunks**: If you have 100+ invoices, process them in batches
4. **Add logging**: Uncomment or add `console.log()` statements to track progress

### Automation with Triggers

To run the script automatically:

1. In Apps Script editor, click the **clock icon** (Triggers)
2. Click **+ Add Trigger**
3. Choose function: `createInvoices`
4. Choose event source: **Time-driven**
5. Select frequency (e.g., daily, weekly)
6. Save

**Note**: Be cautious with automatic triggers - ensure your data spreadsheet only contains invoices you want to generate.

### Preventing Duplicate PDFs

To avoid regenerating the same invoices:

1. Add a "Status" column to your data spreadsheet
2. Modify the script to check this column
3. Skip rows where status is "Generated"
4. Update status after successful PDF creation

Example modification:

```javascript
invoiceDataObjects.forEach(invoiceData => {
  // Skip if already generated
  if (invoiceData['Status'] === 'Generated') {
    return;
  }

  // ... rest of the code ...

  // After successful PDF creation:
  // Update the status column in the original spreadsheet
});
```

---

## Best Practices

1. **Test thoroughly**: Always test with sample data before processing real invoices
2. **Backup your data**: Keep a copy of your data spreadsheet
3. **Version control**: Save copies of your template before making major changes
4. **Naming conventions**: Use clear, consistent naming for files and columns
5. **Documentation**: Keep notes about your customizations
6. **Error handling**: Monitor execution logs regularly
7. **Security**: Keep your spreadsheet IDs and folder IDs private

---

## Support and Resources

- **Google Apps Script Documentation**: https://developers.google.com/apps-script
- **Spreadsheet Service Reference**: https://developers.google.com/apps-script/reference/spreadsheet
- **Drive Service Reference**: https://developers.google.com/apps-script/reference/drive

---

## License

This project is provided as-is for educational and personal use.

---

## Changelog

### Version 1.0
- Initial release
- Header-based data mapping
- Automatic date formatting
- Batch PDF generation
- Error handling for missing data
