# Invoice Template Reference

## Overview

This directory contains the Google Sheets invoice template used to generate PDF invoices.

## Template Files

- **File**: `invoice-template.csv` - Current template with dynamic seller information
- **File**: `invoice-template-old.csv` - Previous template version for reference

## Template Placeholders

The template uses the following placeholders that get replaced with actual invoice data:

### Invoice Details
| Placeholder | Description | Example Value |
|------------|-------------|---------------|
| `{{invoice_date}}` | Invoice date | 2025/10/01 |

### Seller Information
| Placeholder | Description | Example Value |
|------------|-------------|---------------|
| `{{seller_name}}` | Seller company name | 株式会社サンプル商事 |
| `{{seller_address}}` | Seller address | 東京都渋谷区サンプル町1-2-3 |
| `{{seller_phone_number}}` | Seller phone | 03-1234-5678 |

### Item Details
| Placeholder | Description | Example Value |
|------------|-------------|---------------|
| `{{item_1_name}}` | Item description | システム開発費 |
| `{{item_1_number}}` | Item quantity | 1 |
| `{{item_1_price}}` | Item price | 100000 |

### Bank Information
| Placeholder | Description | Example Value |
|------------|-------------|---------------|
| `{{seller_bank_name}}` | Bank name | サンプル銀行 東京支店(123) |
| `{{seller_bank_type}}` | Account type | 普通 |
| `{{seller_bank_number}}` | Account number | 1234567 |
| `{{seller_bank_holder_name}}` | Account holder | 株式会社サンプル商事 |

## Template Structure

The template includes the following sections:

1. **Header**
   - Title: "請求書" (Invoice)
   - Client name: 株式会社 DROX (hardcoded) with "御中" (To)

2. **Invoice Information**
   - Issue date (発行日): `{{invoice_date}}`
   - Seller company: `{{seller_name}}`
   - Seller address: `{{seller_address}}`
   - Seller phone: `{{seller_phone_number}}`

3. **Items Table**
   - Line items (No. 1-9)
   - Item 1: `{{item_1_name}}`, `{{item_1_number}}`, `{{item_1_price}}`
   - Columns: 摘要 (Description), 数量 (Quantity), 単価 (Unit Price), 金額 (Amount)
   - Subtotal, tax, and total fields (calculated automatically)

4. **Payment Information**
   - Bank name: `{{seller_bank_name}}`
   - Account type: `{{seller_bank_type}}`
   - Account number: `{{seller_bank_number}}`
   - Account holder: `{{seller_bank_holder_name}}`

5. **Notes**
   - Remarks section (備考)

## How It Works

1. The Google Apps Script (`scripts/invoice-generator.gs`) reads invoice data
2. It copies this template sheet
3. Replaces all `{{placeholder}}` values with actual data
4. Exports the completed sheet as a PDF
5. Saves the PDF to Google Drive

## Template Setup in Google Sheets

To use this template:

1. Create a new Google Sheets document
2. Import this CSV structure or recreate the layout
3. Use the placeholders exactly as shown: `{{placeholder_name}}`
4. Copy the Spreadsheet ID and add it to `TEMPLATE_SPREADSHEET_ID` in the script

## Customization

To customize the template:

1. Edit your Google Sheets template directly
2. Keep the placeholder format: `{{placeholder_name}}`
3. Match placeholder names with the data fields in `INVOICE_DATA` array
4. Update formatting, colors, and layout as needed

## Notes

- The template must remain in CSV/Google Sheets format
- Placeholders are case-sensitive
- Ensure placeholder names match the data field names in the script
- Template formatting (currency, date formats) is preserved when generating PDFs
