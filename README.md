# TicknTie

An open-source web application for evidence organization using and excel like interface. TicknTie allows auditors and analysts to link images to Excel cells, view full-resolution images in a dedicated panel, and export workbooks as zip bundles containing both the Excel file and associated images.

## Features

- **Side Panel Viewer**: View full-resolution images in a dedicated panel
- **Zip Bundle Support**: Import/export workbooks with images as zip files
- **Browser-Based**: No installation required, runs entirely in your browser

## Quick Start

### Option 1: Use Hosted Version
Visit [https://evidence.audittoolbox.com/](https://evidence.audittoolbox.com/)

### Option 2: Run Locally

1. Clone the repository:
```bash
git clone https://github.com/rp4/TicknTie.git
cd tickntie
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm run dev
```

## Usage

### Working with Excel Files

1. **Open Excel File**: Click "Open Excel" or drag-drop an .xlsx file
2. **Edit Cells**: Click on any cell to edit its value
3. **Add Sheets**: Click the "+" button in the sheet tabs
4. **Save Work**: Click "Export Excel" to download the modified file

### Working with Images

1. **Add Image to Cell**: 
   - Select a cell
   - Click "Add Image" in the toolbar
   - Choose image file(s)

2. **View Images**:
   - Click on a cell with an image to preview in side panel
   - Use Gallery tab to see all images
   - Click images for full-screen view

3. **Remove Images**:
   - Select cell with image
   - Click "Remove from Cell" in the side panel

### Zip Bundle Format

TicknTie uses a special zip format for preserving image-cell relationships:

```
workbook.zip
├── workbook.xlsx       # Excel file with image references
├── images/            # Folder containing all images
│   ├── img_001.png
│   ├── img_002.jpg
│   └── ...
├── mappings.json      # Metadata about image-cell relationships
└── README.txt         # Bundle information
```

To export as zip: Click "Export as Zip"
To import zip: Click "Open Zip Bundle" or drag-drop a .zip file

## Architecture

### Key Technologies

- **Univer**: Modern Excel-like spreadsheet framework
- **ExcelJS**: Excel file reading/writing with full formatting support
- **JSZip**: Zip file creation and extraction
- **Vite**: Fast build tool and development server
- **Vanilla JavaScript**: No framework dependencies

## Development

### Building from Source

```bash
# Install dependencies
npm install

# Run development server
npm run dev

# Run tests (when available)
npm test
```

### Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Limitations

- All data is processed client-side (no server required)
- Large files may impact browser performance
- Excel formulas have limited support
- Images are stored in browser memory during session

## Security & Privacy

- **100% Client-Side**: No data is sent to any server
- **No Tracking**: No analytics or tracking code
- **Open Source**: Full code transparency
- **Local Storage**: Optional local browser storage

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

<p align="center">
  Made with ❤️ by <a href="https://www.audittoolbox.com/">AuditToolbox</a> • 
</p>