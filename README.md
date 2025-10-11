# Invoice Automation with Google Apps Script

Automatically generate invoice PDFs from Google Sheets data using Google Apps Script. This project provides a flexible, header-based system that maps spreadsheet data to invoice templates and generates professional PDF invoices.

## Features

- **Header-Based Mapping**: Column order doesn't matter - uses header names to map data
- **Batch Processing**: Generate multiple invoices in a single execution
- **Automatic Date Formatting**: Properly formats date values
- **Error Handling**: Skips rows with missing critical data
- **Customizable Templates**: Design your invoice layout freely
- **PDF Export**: Automatic conversion and saving to Google Drive

## Quick Start

1. **Prepare Your Spreadsheets**
   - Create a data spreadsheet with invoice information
   - Create a template spreadsheet with your invoice design
   - Create a Google Drive folder for PDFs

2. **Set Up the Script**
   - Open Apps Script from your data spreadsheet
   - Copy the script from `scripts/invoice-generator.gs`
   - Update the configuration with your IDs

3. **Run the Script**
   - Execute `createInvoices()` function
   - Grant necessary permissions
   - Check your Drive folder for generated PDFs

## Documentation

📖 **[Complete Setup Guide](docs/SETUP_GUIDE.md)** - Detailed step-by-step instructions

## Project Structure

```
invoice-automation/
├── README.md                    # This file
├── docs/
│   └── SETUP_GUIDE.md          # Comprehensive setup instructions
└── scripts/
    └── invoice-generator.gs     # Main Google Apps Script code
```

## Requirements

- Google Account
- Google Sheets
- Google Drive
- Google Apps Script (included with Google Account)

## Example Data Structure

**Data Spreadsheet (Row 1 = Headers):**

| Invoice Date | Invoice Number | Company Name | Description | Amount | PDF Filename |
|--------------|----------------|--------------|-------------|--------|--------------|
| 2025/10/01   | INV-001        | ABC Corp     | Web Dev     | 100000 | 20251001_ABC |
| 2025/10/02   | INV-002        | XYZ Ltd      | Design      | 50000  | 20251002_XYZ |

**Template Spreadsheet (with placeholders):**

```
INVOICE

{{Company Name}}                           Invoice #: {{Invoice Number}}
Date: {{Invoice Date}}

Description: {{Description}}
Amount: {{Amount}}
```

## How It Works

1. Script reads data from your spreadsheet using header names
2. For each row, creates a copy of the template
3. Replaces `{{header_name}}` placeholders with actual data
4. Converts the populated sheet to PDF
5. Saves PDF to specified Google Drive folder
6. Cleans up temporary working sheets

## Configuration

Update these values in the script:

```javascript
const DATA_SPREADSHEET_ID = 'your-data-spreadsheet-id';
const TEMPLATE_SPREADSHEET_ID = 'your-template-spreadsheet-id';
const OUTPUT_FOLDER_ID = 'your-drive-folder-id';
const DATA_SHEET_NAME = 'Sheet1';
const TEMPLATE_SHEET_NAME = 'Sheet1';
```

## Testing Functions

The script includes helpful testing functions:

- `testConfiguration()` - Verify all IDs are correct
- `testFirstInvoice()` - Test with just the first row of data

## Troubleshooting

Common issues and solutions are covered in the [Setup Guide](docs/SETUP_GUIDE.md#troubleshooting).

## License

This project is provided as-is for educational and personal use.

## Contributing

Feel free to fork this project and adapt it to your needs. If you have improvements or bug fixes, pull requests are welcome.

## Support

For detailed instructions, please refer to the [Complete Setup Guide](docs/SETUP_GUIDE.md).

---

Made with ❤️ for automating tedious invoice generation tasks
