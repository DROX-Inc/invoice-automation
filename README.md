# Invoice Automation with Google Apps Script

Automatically generate invoice PDFs from hardcoded data using Google Apps Script. This project provides a simple, easy-to-edit system that generates professional PDF invoices from a template spreadsheet.

## Features

- **CSV Data Source**: Invoice data read from CSV file with email addresses
- **Email Integration**: Automatically sends invoice emails with PDF links after generation
- **Batch Processing**: Generate and send multiple invoices in a single execution
- **Notion Integration**: Fetches hours data from Notion database for billing
- **Automatic Date Formatting**: Auto-generates invoice dates and filenames
- **Error Handling**: Robust error handling for PDF generation and email sending
- **Customizable Templates**: Design your invoice layout freely
- **PDF Export**: Automatic conversion and saving to Google Drive
- **Test Mode**: Option to preview email content without actually sending

## Quick Start

1. **Prepare Your Resources**
   - Create a data spreadsheet with invoice data (see `google-spread-sheets/invoice-data.csv`)
   - Create a template spreadsheet with your invoice design (see `google-spread-sheets/invoice-template.csv`)
   - Create Google Drive folders for each recipient's PDFs

2. **Set Up the Script**
   - Open Google Apps Script editor
   - Copy the script from `scripts/invoice-generator.gs`
   - Update the configuration IDs:
     - `DATA_SPREADSHEET_ID`: Your data spreadsheet ID
     - `TEMPLATE_SPREADSHEET_ID`: Your template spreadsheet ID
     - `NOTION_API_KEY`: Your Notion integration token (optional)
   - Configure email settings:
     - `SENDER_EMAIL`: Your business email address
     - `TEST_MODE`: Set to `true` for testing without sending emails

3. **Run the Script**
   - Execute `createInvoices()` function
   - Grant necessary permissions (Gmail, Drive, Sheets)
   - Check your Drive folders for generated PDFs
   - Monitor console logs for email sending status

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
│   ├── invoice-data.csv               # Invoice data with email addresses
│   └── invoice-template.csv           # Template structure reference
└── scripts/
    └── invoice-generator.gs            # Main Google Apps Script with email integration
```

## Requirements

- Google Account
- Google Sheets
- Google Drive
- Google Apps Script (included with Google Account)

## Example Data Structure

**CSV Data Source (invoice-data.csv):**

```csv
seller_name,seller_address,seller_phone_number,seller_email,item_1_name,item_1_number,item_1_price,seller_bank_name,seller_bank_type,seller_bank_number,seller_bank_holder_name,output_folder_id
與儀源喜知,〒161-0031 東京都新宿区西落合4丁目17-4 パークレジデンス西落合401,03-0000-0000,yogi@example.com,Webシステム開発費,1,150000,三井住友銀行 テスト支店（111）,普通,1111111,ヨギゲンキチ,14iWEOIgVM6iX6NCGn5CP1MLU22ykiTYQ
小林大樹,〒108-0023 東京都港区芝浦4-12-36-1002,03-0000-0000,kobayashi@example.com,コンサルティング費用,1,200000,三井住友銀行 オリーブDILL支店（864）,普通,0104482,コバヤシダイキ,14f_-ygHrs6FfX_uVM8KGezcnvkHHkKRf
```

**Email Configuration in Script:**

```javascript
// Email Settings
const EMAIL_ENABLED = true;      // Enable/disable email sending
const TEST_MODE = false;          // Test mode (logs only, no actual sending)
const SENDER_EMAIL = "koki-hata@drox-inc.com";
const SENDER_NAME = "株式会社DROX";
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

1. Script reads invoice data from the data spreadsheet (CSV)
2. Optionally fetches hours data from Notion database
3. For each invoice entry, creates a copy of the template sheet
4. Replaces `{{placeholder}}` values with actual data
5. Converts the populated sheet to PDF
6. Saves PDF to specified Google Drive folder
7. Generates shareable link for the PDF file
8. Sends email to recipient with PDF link (if email address provided)
9. Cleans up temporary working sheets
10. Logs summary of PDFs created and emails sent

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
