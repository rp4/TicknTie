# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TicknTie is an open-source web application for auditing Excel files with image attachments, designed in the style of DataSnipper. It provides a simple image viewer sidebar plugin for Univer spreadsheets, allowing users to attach images to cells via "pushpin" markers and view full-resolution images in a side panel.

## Development Approach

This project uses **Vite + Univer** for the best balance of simplicity and functionality:
- **Vite** provides instant hot-reload development and optimized production builds
- **Univer** handles all spreadsheet functionality (formulas, Excel import/export, etc.)
- **Image Sidebar** is a lightweight plugin that adds image viewing capabilities

## Project Structure

```
TicknTie/
├── src/
│   ├── main.js              # Application entry point & Univer initialization
│   ├── image-plugin.js      # Image sidebar plugin for Univer
│   ├── styles.css           # Application styles
│   └── index.html           # Main HTML template
├── public/                  # Static assets
├── package.json            # Dependencies and scripts
├── vite.config.js          # Vite configuration
├── README.md               # Project documentation
└── .gitignore
```

## Core Architecture

### Simplified Data Model
- **Univer Instance**: Handles all spreadsheet operations
- **Image Store**: Simple Map storing image data
  ```javascript
  imageStore = new Map() // imageId -> { url, name, size, type }
  cellImages = new Map() // cellRef -> imageId
  ```

### Key Components

1. **Univer Spreadsheet** (Left Side)
   - Full Excel-like functionality
   - Formula support
   - Import/Export capabilities
   - Handled entirely by Univer

2. **Image Sidebar Plugin** (Right Side)
   - Simple image viewer
   - Upload images to cells
   - Display selected cell's image
   - Remove images from cells

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

### image-plugin.js - Image Sidebar
```javascript
export class ImagePlugin {
  constructor(univerAPI) {
    this.univerAPI = univerAPI
    this.imageStore = new Map()
    this.cellImages = new Map()
  }

  init() {
    this.createSidebar()
    this.listenToSelection()
    this.setupImageUpload()
  }

  // Simple sidebar implementation
  createSidebar() { /* Create HTML elements */ }

  // Listen for cell selection
  listenToSelection() {
    this.univerAPI.onSelectionChange(selection => {
      this.updateSidebar(selection)
    })
  }
}
```

## Working with Features

### Adding Image to Cell
1. Select a cell in Univer spreadsheet
2. Click "Add Image" button in sidebar
3. Image stored in memory with unique ID
4. Cell displays 📌 icon with filename
5. Image appears in sidebar

### Cell-Image Association
- Images linked to cells via simple Map
- No complex data structures needed
- Univer handles cell references

### Image Display
- Click cell with image → Shows in sidebar
- Display image info (name, size, type)
- Remove button to clear image

## Important Design Decisions

1. **Keep It Simple**: Leverage Univer for all spreadsheet complexity
2. **Minimal State**: Just two Maps for image management
3. **No Custom Rendering**: Use cell values for visual indicators (📌 emoji)
4. **Memory Only**: Images stored in browser memory during session
5. **Clean Separation**: Spreadsheet logic (Univer) vs Image logic (Plugin)

## Development Best Practices

### When Adding Features
1. **Let Univer handle it** if it's spreadsheet-related
2. **Keep image logic simple** - just storage and display
3. **Use Vite's HMR** for fast development cycles
4. **Test with real Excel files** to ensure compatibility

### Performance Considerations
- **Lazy load images** only when cells are selected
- **Use object URLs** for efficient image display
- **Clean up URLs** when images are removed
- **Consider file size limits** for browser memory

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

### Vite Benefits
- ⚡ **Instant updates** during development (HMR)
- 📦 **Optimized builds** for production
- 🎯 **Zero config** to start
- 🔧 **Modern tooling** out of the box

### Univer Benefits
- 📊 **Full spreadsheet** functionality
- 📁 **Excel compatibility** built-in
- 🔢 **Formula support** included
- 🎨 **Professional UI** ready to use

### Simplicity Benefits
- 📝 **Minimal code** to maintain
- 🚀 **Fast to develop** new features
- 🐛 **Easy to debug** issues
- 👥 **Simple to understand** for new developers

## Common Tasks

### Start Fresh Development
```bash
git checkout univer-simple-plugin
npm install
npm run dev
# Open http://localhost:5173
```

### Add New Feature
1. Modify `image-plugin.js` for image-related features
2. Let Univer handle spreadsheet features
3. Save file → See changes instantly (HMR)

### Deploy to Production
```bash
npm run build
# Upload dist/ folder to hosting service
```

## Testing Guidelines

1. **Excel Files**: Test import/export with real .xlsx files
2. **Image Formats**: Verify PNG, JPG, GIF support
3. **Large Files**: Test with files >10MB
4. **Multiple Images**: Test with 50+ images attached
5. **Browser Compatibility**: Test in Chrome, Firefox, Safari