# Invoice Automation with Google Apps Script

Automatically generate invoice PDFs from hardcoded data using Google Apps Script. This project provides a simple, easy-to-edit system that generates professional PDF invoices from a template spreadsheet.

## Features

- **Easy-to-Edit Data**: Invoice data stored directly in the script as an array
- **Batch Processing**: Generate multiple invoices in a single execution
- **Automatic Date Formatting**: Properly formats date values
- **Error Handling**: Skips rows with missing critical data
- **Customizable Templates**: Design your invoice layout freely
- **PDF Export**: Automatic conversion and saving to Google Drive

## Quick Start

1. **Prepare Your Resources**
   - Create a template spreadsheet with your invoice design (see `google-spread-sheets/invoice-template.csv` for reference)
   - Create a Google Drive folder for generated PDFs

2. **Set Up the Script**
   - Open Google Apps Script editor
   - Copy the script from `scripts/invoice-generator.gs`
   - Update `INVOICE_DATA` array with your invoice information
   - Update `TEMPLATE_SPREADSHEET_ID` and `OUTPUT_FOLDER_ID` with your IDs

3. **Run the Script**
   - Execute `createInvoices()` function
   - Grant necessary permissions
   - Check your Drive folder for generated PDFs

## Documentation

- 📖 **[Complete Setup Guide](docs/SETUP_GUIDE.md)** - Detailed step-by-step instructions
- 📋 **[Template Reference](google-spread-sheets/README.md)** - Invoice template documentation and structure

## Project Structure

```
invoice-automation/
├── README.md                           # This file
├── docs/
│   ├── SETUP_GUIDE.md                 # Comprehensive setup instructions
│   └── QUICK_REFERENCE.md             # Quick reference guide
├── google-spread-sheets/
│   ├── README.md                      # Template documentation
│   └── invoice-template.csv           # Template structure reference
└── scripts/
    └── invoice-generator.gs            # Main Google Apps Script code
```

## Requirements

- Google Account
- Google Sheets
- Google Drive
- Google Apps Script (included with Google Account)

## Example Data Structure

**INVOICE_DATA Array in Script:**

```javascript
const INVOICE_DATA = [
  {
    invoice_date: "2025/11/01",
    seller_name: "株式会社サンプル商事",
    seller_address: "東京都渋谷区サンプル町1-2-3 サンプルビル4階",
    item_1_name: "Webシステム開発費",
    item_1_number: 1,
    item_1_price: 150000,
    seller_bank_name: "サンプル銀行 東京支店(123)",
    seller_bank_type: "普通",
    seller_bank_number: "1234567",
    seller_bank_holder_name: "株式会社サンプル商事",
    pdfFileName: "20251101_株式会社DROX様_請求書"
  }
];
```

**Template Spreadsheet (with placeholders):**

```
請求書 (INVOICE)

株式会社 DROX 御中
Date: {{invoice_date}}

Seller: {{seller_name}}
Address: {{seller_address}}

Item: {{item_1_name}}
Quantity: {{item_1_number}}
Price: {{item_1_price}}

Bank: {{seller_bank_name}}
```

See [Template Reference](google-spread-sheets/README.md) for complete template structure.

## How It Works

1. Script reads invoice data from the `INVOICE_DATA` array
2. For each invoice entry, creates a copy of the template sheet
3. Replaces `{{placeholder}}` values with actual data
4. Converts the populated sheet to PDF
5. Saves PDF to specified Google Drive folder
6. Cleans up temporary working sheets

## Configuration

Update these values in the script:

```javascript
// Add your invoice data here
const INVOICE_DATA = [
  {
    invoice_date: "2025/11/01",
    seller_name: "株式会社サンプル商事",
    seller_address: "東京都渋谷区サンプル町1-2-3 サンプルビル4階",
    item_1_name: "Webシステム開発費",
    item_1_number: 1,
    item_1_price: 150000,
    seller_bank_name: "サンプル銀行 東京支店(123)",
    seller_bank_type: "普通",
    seller_bank_number: "1234567",
    seller_bank_holder_name: "株式会社サンプル商事",
    pdfFileName: "20251101_株式会社DROX様_請求書"
  }
];

const TEMPLATE_SPREADSHEET_ID = 'your-template-spreadsheet-id';
const OUTPUT_FOLDER_ID = 'your-drive-folder-id';
const TEMPLATE_SHEET_NAME = 'Sheet1';
```

## Adding New Invoices

Simply add new invoice objects to the `INVOICE_DATA` array:

```javascript
const INVOICE_DATA = [
  {
    invoice_date: "2025/11/01",
    seller_name: "株式会社サンプル商事",
    // ... other fields ...
    pdfFileName: "20251101_株式会社DROX様_請求書"
  },
  {
    invoice_date: "2025/11/15",
    seller_name: "株式会社サンプル商事",
    // ... other fields ...
    pdfFileName: "20251115_株式会社DROX様_請求書"
  }
];
```

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
