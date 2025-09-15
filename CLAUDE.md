# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TicknTie is an open-source web application for auditing Excel files with image attachments, designed in the style of DataSnipper. It allows users to link images to Excel cells and view full-resolution images in a side panel.

## Application Architecture

### Project Structure
```
TicknTie/
├── src/
│   ├── css/
│   │   └── styles.css          # All application styles
│   ├── js/
│   │   ├── app.js              # Main application logic
│   │   ├── excel-handler.js    # Excel file processing
│   │   ├── image-manager.js    # Image storage and linking
│   │   ├── ui-components.js    # UI rendering and events
│   │   └── zip-handler.js      # Zip file import/export
│   └── index.html              # Main HTML file
├── public/
│   └── images/                 # Temporary image storage
├── package.json                # Dependencies and scripts
├── README.md                   # Project documentation
├── LICENSE                     # MIT License
└── .gitignore
```

### Core Data Structures
- `workbook`: Main object storing all Excel data
- `sheets`: Array of sheet objects with cells and metadata
- `currentSheet`: Index of active sheet
- `imageStore`: Map of image IDs to image data (replaces cellImages)
  - Format: `{ "img_001.png": { url: "blob:...", data: "base64...", name: "original.png" } }`
- `cellImageRefs`: Map of cell references to image IDs
  - Format: `{ "A1": "img_001.png", "B2": "img_002.png" }`
- `selectedCells`: Array tracking multi-cell selection

### Key Functionality Areas
1. **Excel Processing**: Uses SheetJS library (installed via npm)
2. **Image Handling**: 
   - Images stored separately from Excel data
   - Cells contain image references (e.g., "img_001.png")
   - Full images displayed in side panel
3. **Zip Support**: JSZip for bundling Excel files with images
4. **Cell Rendering**: Dynamic grid with image thumbnails
5. **Side Panel**: Full image preview and image gallery

## Development Commands

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production (if build system added later)
npm run build

# Run tests (when added)
npm test
```

## Key Modules and Their Responsibilities

### app.js
- Main application initialization
- Event coordination between modules
- Global state management

### excel-handler.js
- `loadExcelFile(file)`: Process uploaded Excel files
- `exportExcel()`: Generate Excel with image references
- `parseWorkbook(data)`: Convert Excel data to internal format
- `updateCellValue(sheet, cell, value)`: Update cell data

### image-manager.js
- `storeImage(file)`: Store image and return ID
- `linkImageToCell(cellRef, imageId)`: Create cell-image association
- `getImageData(imageId)`: Retrieve image for display
- `generateThumbnail(imageId)`: Create cell thumbnail
- `clearImageStore()`: Clean up memory

### ui-components.js
- `renderSheet()`: Draw spreadsheet grid
- `renderSidePanel()`: Display image preview panel
- `updateCellDisplay(cellRef)`: Update individual cell
- `handleCellSelection(cellRef)`: Process cell clicks
- `toggleSidePanel()`: Show/hide image panel

### zip-handler.js
- `loadZipFile(file)`: Extract Excel and images from zip
- `createZipBundle()`: Package Excel with images folder
- `extractImages(zip)`: Process images from zip
- `generateImageReferences()`: Create Excel-compatible image links

## Working with Features

### Adding Image to Cell
1. User selects cell and clicks "Add Image"
2. Image stored in `imageStore` with unique ID
3. Cell value set to image reference
4. Thumbnail displayed in cell
5. Full image shown in side panel on cell selection

### Exporting with Images
1. User clicks "Export"
2. System creates zip file containing:
   - Excel file with image references
   - Images folder with all image files
3. Image references in Excel use relative paths

### Importing Zip Bundle
1. User uploads .zip file
2. System extracts Excel and images
3. Images loaded into `imageStore`
4. Excel cells mapped to correct images
5. Application ready for editing

## Important Considerations

1. **Image Storage**: Images kept in memory during session, not embedded in Excel
2. **File Size**: Separate image storage prevents Excel file bloat
3. **Performance**: Lazy loading for images, thumbnails for cells
4. **Compatibility**: Excel shows image references, full images only in app
5. **Browser Limits**: Consider memory usage with many/large images
6. **Security**: Validate image files, sanitize filenames

## API/Library Dependencies

- **SheetJS (xlsx)**: Excel file processing
- **JSZip**: Zip file creation and extraction
- **No framework**: Vanilla JavaScript for simplicity
- **No build tools initially**: Direct module loading

## Testing Guidelines

When adding features:
1. Test with various Excel formats (.xlsx, .xls)
2. Verify image formats (PNG, JPG, GIF)
3. Test large files (>50MB)
4. Check memory usage with many images
5. Validate zip export/import cycle