# TicknTie

An open-source web application for auditing Excel files with image attachments, inspired by DataSnipper. TicknTie allows auditors and analysts to link images to Excel cells, view full-resolution images in a dedicated panel, and export workbooks as zip bundles containing both the Excel file and associated images.

## Features

- **Excel File Management**: Load, edit, and save Excel files (.xlsx, .xls)
- **Image Linking**: Attach images to specific cells with URL references instead of embedding
- **Side Panel Viewer**: View full-resolution images in a dedicated panel
- **Image Gallery**: Browse all images in your workbook
- **Zip Bundle Support**: Import/export workbooks with images as zip files
- **Cell Formatting**: Apply formatting to cells (bold, italic, colors, alignment)
- **Multiple Sheets**: Work with multiple worksheets
- **Keyboard Shortcuts**: Efficient navigation and operations
- **Browser-Based**: No installation required, runs entirely in your browser

## Quick Start

### Option 1: Use Hosted Version
Visit [https://yourdomain.com/tickntie](https://yourdomain.com/tickntie) (when deployed)

### Option 2: Run Locally

1. Clone the repository:
```bash
git clone https://github.com/yourusername/tickntie.git
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

4. Open your browser to `http://localhost:8080`

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

### Project Structure

```
TicknTie/
├── src/
│   ├── css/
│   │   └── styles.css          # Application styles
│   ├── js/
│   │   ├── app.js              # Main application
│   │   ├── excel-handler.js    # Excel operations
│   │   ├── image-manager.js    # Image management
│   │   ├── ui-components.js    # UI rendering
│   │   └── zip-handler.js      # Zip operations
│   └── index.html              # Main HTML file
├── package.json                # Dependencies
└── README.md                   # This file
```

### Key Technologies

- **SheetJS (xlsx)**: Excel file processing
- **JSZip**: Zip file creation and extraction
- **Vanilla JavaScript**: No framework dependencies
- **HTML5 APIs**: File, Blob, and Canvas APIs

## Development

### Prerequisites

- Node.js (v14 or higher)
- npm or yarn
- Modern web browser

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

## Browser Compatibility

- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+

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

## Acknowledgments

- Inspired by [DataSnipper](https://www.datasnipper.com/)
- Built with [SheetJS](https://sheetjs.com/)
- Zip handling by [JSZip](https://stuk.github.io/jszip/)

## Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/tickntie/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/tickntie/discussions)
- **Email**: support@tickntie.com

## Roadmap

- [ ] Cloud storage integration (Google Drive, Dropbox)
- [ ] Collaborative editing features
- [ ] Advanced formula support
- [ ] PDF export with images
- [ ] Mobile responsive design
- [ ] OCR for image text extraction
- [ ] Audit trail and versioning
- [ ] Custom plugins/extensions

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history.

---

Made with ❤️ by the TicknTie community