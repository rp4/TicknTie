# AI Agent Instructions for Creating TicknTie-Compatible Zip Files

This document provides instructions for AI agents (Claude Desktop, ChatGPT, etc.) to create zip files that can be imported into TicknTie.

## Quick Start Prompt

Copy and paste this prompt to your AI agent:

```
Create a zip file with the following structure for a TicknTie audit project:

1. An Excel file named "workbook.xlsx" with:
   - Sample data in cells (financial data, audit notes, etc.)
   - Hyperlinks in specific cells pointing to evidence files
   - Each hyperlink should:
     - Link to: evidence/filename.ext (relative path)
     - Display as: 📌 filename.ext

2. An "evidence" folder containing:
   - Sample images (receipts, invoices, screenshots)
   - Sample PDFs (documents, reports)
   - All files under 50MB each

3. Structure:
   project.zip
   ├── workbook.xlsx
   ├── evidence/
   │   ├── receipt_001.jpg
   │   ├── invoice_002.pdf
   │   └── ...
   └── README.txt (optional)

Important: The hyperlinks must be actual Excel hyperlinks (Insert > Hyperlink), not just text.
```

## Detailed Requirements

### Zip File Structure

```
your-project.zip
├── workbook.xlsx       # Excel file with hyperlinks to evidence
├── evidence/           # Folder containing all evidence files
│   ├── receipt_001.jpg
│   ├── invoice_002.pdf
│   ├── screenshot_003.png
│   └── ...
└── README.txt          # Optional project information
```

### 1. Excel File Requirements

- **Filename**: Must be exactly `workbook.xlsx`
- **Content**: Any spreadsheet data (financial records, audit trails, etc.)
- **Hyperlinks**:
  - Use Excel's hyperlink feature (not plain text)
  - Link format: `evidence/filename.ext` (relative path)
  - Display text: `📌 filename.ext` (includes pushpin emoji)
  - Example: Cell A1 hyperlinks to `evidence/receipt.jpg` displaying `📌 receipt.jpg`

### 2. Evidence Folder Requirements

- **Folder name**: Must be exactly `evidence` (lowercase)
- **File placement**: All files directly in this folder (no subfolders)
- **Supported formats**:
  - Images: JPG, PNG, GIF, WebP
  - Documents: PDF
- **Size limits**: Maximum 50MB per file
- **Naming conventions**:
  - Remove special characters: `< > : " / \ | ? *`
  - Keep file extensions intact
  - Use underscores or hyphens for spaces
  - Ensure unique names (add _1, _2 if duplicates)

### 3. Sample Excel Content Ideas

For a realistic audit workbook, consider including:

```
Column A: Transaction ID
Column B: Date
Column C: Description
Column D: Amount
Column E: Evidence (hyperlinks to files)

Row 1: Headers
Row 2: TRX001 | 2024-01-15 | Office Supplies | $245.50 | 📌 receipt_001.pdf
Row 3: TRX002 | 2024-01-16 | Travel Expense | $1,200.00 | 📌 boarding_pass.jpg
Row 4: TRX003 | 2024-01-17 | Client Dinner | $180.75 | 📌 restaurant_receipt.png
```

### 4. Creating Hyperlinks in Excel

For AI agents that can create Excel files:

1. **Using Excel formulas**:
   ```
   =HYPERLINK("evidence/receipt_001.jpg", "📌 receipt_001.jpg")
   ```

2. **Using Excel UI**:
   - Select cell
   - Insert > Hyperlink
   - Address: `evidence/filename.ext`
   - Text to display: `📌 filename.ext`

### 5. Sample Evidence Files

Create diverse evidence files such as:
- `receipt_001.jpg` - Store receipt
- `invoice_002.pdf` - Vendor invoice
- `bank_statement_003.pdf` - Financial document
- `email_confirmation_004.png` - Screenshot
- `contract_005.pdf` - Legal document
- `expense_report_006.jpg` - Scanned form

## Testing Your Zip File

Before delivering the zip file, verify:

1. ✅ File is named with .zip extension
2. ✅ Contains `workbook.xlsx` at root level
3. ✅ Contains `evidence/` folder at root level
4. ✅ All evidence files are inside the `evidence` folder
5. ✅ Excel file opens correctly
6. ✅ Hyperlinks in Excel use relative paths (`evidence/...`)
7. ✅ Hyperlink display text includes 📌 emoji
8. ✅ All filenames are properly sanitized

## Common Issues to Avoid

❌ **Don't**: Use absolute paths in hyperlinks (C:\Users\...)
✅ **Do**: Use relative paths (evidence/filename.ext)

❌ **Don't**: Create subfolders inside evidence folder
✅ **Do**: Place all files directly in evidence folder

❌ **Don't**: Use plain text that looks like links
✅ **Do**: Use actual Excel hyperlink feature

❌ **Don't**: Forget the 📌 emoji in display text
✅ **Do**: Include it to help TicknTie identify linked cells

## Example Prompt for Specific Industries

### For Financial Audit:
```
Create a TicknTie zip file for a financial audit with:
- Excel: Q1 2024 expense report with columns for date, vendor, amount, category, and evidence
- Evidence: 10 sample receipts/invoices as JPG/PDF files
- Link evidence files to corresponding expense rows
```

### For Legal Review:
```
Create a TicknTie zip file for contract review with:
- Excel: Contract tracking sheet with columns for party, date, value, status, and evidence
- Evidence: 5 sample contracts and amendments as PDF files
- Link each contract to its tracking row
```

### For Quality Assurance:
```
Create a TicknTie zip file for QA testing with:
- Excel: Bug tracking sheet with columns for ID, description, severity, screenshot
- Evidence: 8 screenshots showing various issues as PNG files
- Link screenshots to corresponding bug entries
```

## Notes for AI Agents

- If you cannot create actual Excel hyperlinks, inform the user that manual linking may be required
- When generating sample evidence files, ensure they are under 50MB each
- Use realistic but generic filenames that indicate content type
- The 📌 emoji is Unicode U+1F4CC and should be preserved in the display text

## Import Instructions

Once created, the zip file can be imported into TicknTie by:
1. Opening TicknTie in a web browser
2. Clicking "Open Zip Bundle" button
3. Selecting the generated zip file
4. TicknTie will automatically load the Excel file and map all evidence files