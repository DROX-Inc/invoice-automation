# Invoice Generator - Quick Reference Card

## Setup Checklist

- [ ] Create data spreadsheet with headers in Row 1
- [ ] Create template spreadsheet with `{{placeholders}}`
- [ ] Create Google Drive folder for PDFs
- [ ] Copy IDs from URLs
- [ ] Open Apps Script from data spreadsheet
- [ ] Paste script code
- [ ] Update configuration IDs
- [ ] Save script
- [ ] Run `testConfiguration()` first
- [ ] Run `createInvoices()` to generate PDFs

## Finding IDs

### Spreadsheet ID
```
https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
                                        └──────┬──────┘
                                          Copy this
```

### Folder ID
```
https://drive.google.com/drive/folders/FOLDER_ID_HERE
                                        └────┬────┘
                                        Copy this
```

## Configuration Template

```javascript
const DATA_SPREADSHEET_ID = 'paste-your-data-sheet-id';
const TEMPLATE_SPREADSHEET_ID = 'paste-your-template-id';
const OUTPUT_FOLDER_ID = 'paste-your-folder-id';
const DATA_SHEET_NAME = 'Sheet1';  // Update if different
const TEMPLATE_SHEET_NAME = 'Sheet1';  // Update if different
```

## Data Spreadsheet Format

**Row 1 = Headers (Required)**

| Invoice Date | Invoice Number | Company Name | Description | Amount | PDF Filename |
|--------------|----------------|--------------|-------------|--------|--------------|
| 2025/10/01   | INV-001        | ABC Corp     | Service     | 100000 | 20251001_ABC |

**Notes:**
- Headers must match template placeholders
- Amount should be numbers only (no commas)
- Company Name is required (rows without it are skipped)

## Template Format

Use `{{Header Name}}` exactly as it appears in Row 1:

```
                    INVOICE

{{Company Name}}                    Invoice #: {{Invoice Number}}

Date: {{Invoice Date}}

Description: {{Description}}
Amount: {{Amount}}

Total: =C15*D15  (You can use formulas)
```

## Placeholder Rules

✅ **Correct:**
- Data header: "Company Name"
- Template: `{{Company Name}}`

❌ **Incorrect:**
- Data header: "Company Name"
- Template: `{{CompanyName}}` or `{{company_name}}`

**They must match exactly!**

## Testing Functions

Run these before full batch processing:

1. **testConfiguration()** - Verifies all IDs are correct
2. **testFirstInvoice()** - Tests with first data row only
3. **createInvoices()** - Generates all invoices

## Common Issues

| Problem | Solution |
|---------|----------|
| "Sheet not found" | Check sheet names match configuration |
| Placeholders not replaced | Header names must match exactly |
| Date shows as number | Script handles "Invoice Date" automatically |
| No PDFs generated | Check folder ID and permissions |
| Some invoices skipped | Ensure Company Name field has data |

## Execution Steps

1. Open data spreadsheet
2. Go to **Extensions** → **Apps Script**
3. Select function from dropdown
4. Click **Run** button (▶️)
5. Grant permissions on first run
6. Check Drive folder for PDFs

## Permission Dialog (First Run Only)

1. "Authorization required" → Click **Review permissions**
2. Select your Google account
3. "App isn't verified" → Click **Advanced**
4. Click **Go to [Project Name] (unsafe)**
5. Click **Allow**

## After Setup

### To Generate New Invoices:
1. Add new rows to data spreadsheet
2. Run `createInvoices()` again

### To Change Template:
1. Edit template spreadsheet
2. No code changes needed
3. Run `createInvoices()` to use new template

### To Add New Data Fields:
1. Add column header to data sheet
2. Add `{{Header}}` to template
3. No code changes needed

## File Locations

```
Your Google Account
├── Data Spreadsheet (contains invoice data)
├── Template Spreadsheet (contains invoice design)
└── Google Drive
    └── Invoice Folder (contains generated PDFs)
```

## Script Functions Reference

| Function | Purpose |
|----------|---------|
| `createInvoices()` | Main function - generates all invoices |
| `testConfiguration()` | Verify IDs and permissions |
| `testFirstInvoice()` | Test with one invoice |

## Quick Troubleshooting

**Script won't run:**
- Check configuration IDs are updated
- Run `testConfiguration()` first

**PDFs look wrong:**
- Check template formatting
- Verify placeholders match headers

**Some invoices missing:**
- Check for empty Company Name cells
- Review execution logs for errors

## Execution Logs

To view logs:
1. Apps Script editor
2. **View** → **Execution log**
3. Check for errors or warnings

## Support Resources

- 📖 [Full Setup Guide](SETUP_GUIDE.md)
- 📝 Main README.md
- 💻 Script: `scripts/invoice-generator.gs`

---

**Pro Tip:** Always test with a few rows first before processing large batches!
