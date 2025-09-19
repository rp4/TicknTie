# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TicknTie is an open-source web application for auditing Excel files with image attachments, designed in the style of DataSnipper. It provides a simple image viewer sidebar plugin for Univer spreadsheets, allowing users to attach files to cells via "pushpin" markers and view previews in a side panel.

## Development Approach

This project uses **Vite + Univer** for the best balance of simplicity and functionality:
- **Vite** provides instant hot-reload development and optimized production builds
- **Univer** handles all spreadsheet functionality (formulas, Excel import/export, etc.)
- **Image Sidebar** is a lightweight plugin that adds file preview capabilities with smart caching

## Project Structure

```
TicknTie/
├── src/
│   ├── main.js              # Application entry point & Univer initialization
│   ├── image-plugin.js      # Hyperlink-based image sidebar plugin
│   ├── styles.css           # Application styles
│   └── index.html           # Main HTML template
├── public/                  # Static assets
├── package.json            # Dependencies and scripts
├── vite.config.js          # Vite configuration
├── README.md               # Project documentation
└── .gitignore
```

## Core Architecture

### Data Model (Hyperlink-Based)
- **Univer Instance**: Handles all spreadsheet operations
- **Hyperlink Storage**: Files stored as data URLs or web links
  ```javascript
  cellHyperlinks = new Map()  // cellRef -> { url, displayText }
  previewCache = new Map()    // url -> { preview, timestamp, size }
  preloadQueue = []           // Queue for preloading adjacent cells
  ```

### Key Components

1. **Univer Spreadsheet** (Left Side)
   - Full Excel-like functionality
   - Formula support
   - Import/Export capabilities
   - Cells display 📌 icon for attached files

2. **Image Sidebar Plugin** (Right Side)
   - One-click file upload (opens native file picker)
   - Displays image/PDF previews
   - Smart caching with LRU eviction
   - Adjacent cell preloading (3x3 grid)
   - Resizable and collapsible sidebar

## Development Commands

```bash
# Initial setup (one time)
npm create vite@latest . -- --template vanilla
npm install @univerjs/presets

# Development
npm run dev          # Start dev server with hot reload (http://localhost:5173)

# Production
npm run build        # Build for production
npm run preview     # Preview production build
```

## Implementation Guide

### main.js - Application Entry
```javascript
import { createUniver, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/presets/sheets'
import { ImagePlugin } from './image-plugin'

// Initialize Univer
const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  presets: [UniverSheetsCorePreset({ container: 'spreadsheet' })]
})

// Initialize Image Plugin
const imagePlugin = new ImagePlugin(univerAPI)
imagePlugin.init()
```

### image-plugin.js - Hyperlink Image Sidebar
```javascript
export class ImagePlugin {
  constructor(univerAPI) {
    this.univerAPI = univerAPI
    this.cellHyperlinks = new Map()  // cellRef -> { url, displayText }
    this.previewCache = new Map()    // url -> { preview, timestamp, size }
    this.preloadQueue = []
    this.maxCacheSize = 100          // Max cached previews
    this.maxCacheAge = 30 * 60 * 1000 // 30 minutes
  }

  init() {
    this.createSidebar()
    this.listenToSelection()
    this.setupHyperlinkInput()
    this.startCacheCleanup()
  }

  // Direct file upload experience
  setupHyperlinkInput() {
    // Click button → Open file picker → Auto-process file
  }

  // Smart preloading for instant preview
  preloadAdjacentCells(cellRef) {
    // Preload 3x3 grid around selected cell
  }
}
```

## Working with Features

### Adding File to Cell
1. Select a cell in Univer spreadsheet
2. Click "➕ Add File to Cell" button
3. Native file picker opens immediately
4. Select image or PDF file
5. File converts to data URL and stores as hyperlink
6. Cell displays 📌 icon with filename
7. Preview appears instantly in sidebar

### File Storage Approach
- **Uploaded files**: Converted to data URLs (base64 encoded)
- **Web links**: Can be added via URL (future feature)
- **Excel export**: Hyperlinks preserved as text with 📌 icon
- **Session-based**: Files stored in browser memory

### Smart Caching System
- **Preview cache**: Up to 100 previews cached with LRU eviction
- **Adjacent preloading**: 3x3 grid around selected cell
- **Time-based cleanup**: Cache entries expire after 30 minutes
- **Instant display**: Cached previews show without loading

### Cell Selection Detection
Multiple detection methods for reliability:
1. Univer API worksheet selection methods
2. Keyboard navigation (arrow keys, Tab, Enter)
3. Mouse click and focus events
4. Internal service discovery
5. DOM observation fallback

## Important Design Decisions

1. **Hyperlink Architecture**: Files stored as data URLs, enabling future web link support
2. **Direct Upload UX**: One-click file picker without intermediate forms
3. **Smart Caching**: Preview cache with preloading for instant display
4. **No Size Limits**: Since only previews are cached, not full files
5. **Session Storage**: Files exist only during browser session

## Development Best Practices

### When Adding Features
1. **Let Univer handle** all spreadsheet operations
2. **Keep plugin simple** - just file handling and preview
3. **Use caching** for performance optimization
4. **Test selection detection** with mouse and keyboard

### Performance Considerations
- **Lazy load previews** only when needed
- **Preload adjacent cells** in background
- **Use LRU cache** to manage memory
- **Convert to data URLs** for instant access
- **Clean up old cache** entries periodically

## Deployment

```bash
# Build for production
npm run build

# Output in dist/ folder
# - Optimized and minified
# - Ready for static hosting
# - Can be served from any web server
```

## Why This Approach?

### Hyperlink Benefits
- ✅ **Future-proof**: Ready for web links and network files
- 📎 **Excel compatible**: Links preserved in exports
- 🚀 **Fast previews**: Data URLs load instantly
- 💾 **Flexible storage**: Supports both local and remote files

### Caching Benefits
- ⚡ **Instant display**: No fetch delay for cached previews
- 🎯 **Smart preloading**: Adjacent cells ready before selection
- 🧹 **Automatic cleanup**: Memory managed with LRU eviction
- 📊 **Scalable**: Handles many files efficiently

### UX Benefits
- 🖱️ **One-click upload**: Direct file picker access
- 📌 **Visual indicators**: Pushpin icons in cells
- 🔍 **Auto-expand sidebar**: Opens when cell with file selected
- ↔️ **Resizable panel**: Adjust sidebar width as needed

## Common Tasks

### Start Fresh Development
```bash
git checkout univer-simple-plugin
npm install
npm run dev
# Open http://localhost:5173 (or shown port)
```

### Debug Cell Selection
1. Open browser console (F12)
2. Click cells or use arrow keys
3. Look for: "Selection changed to: B3"
4. Check debug output for API availability

### Add New File Type Support
1. Update `accept` attribute in file input
2. Add type detection in `getFileType()`
3. Add preview handler for new type
4. Test with sample files

### Deploy to Production
```bash
npm run build
# Upload dist/ folder to hosting service
```

## Testing Guidelines

1. **Cell Selection**: Test with mouse clicks and keyboard navigation
2. **File Upload**: Verify immediate file picker opening
3. **Image Formats**: Test PNG, JPG, GIF, WebP support
4. **PDF Preview**: Verify first page preview generation
5. **Cache Performance**: Test with 50+ files attached
6. **Browser Compatibility**: Test in Chrome, Firefox, Safari
7. **Excel Export**: Verify cells show 📌 icon with filename

## Known Limitations

1. **Session Storage**: Files lost on page refresh (by design)
2. **Excel Hyperlinks**: Currently stored as text, not true hyperlinks
3. **CORS**: Web URLs may have cross-origin restrictions
4. **Large Files**: Very large files may impact browser memory

## Future Enhancements

- Support for web URLs alongside file uploads
- True Excel hyperlink export when Univer API supports it
- Persistent storage option (IndexedDB)
- Multi-page PDF navigation
- Batch file upload support