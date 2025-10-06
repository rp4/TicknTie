# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TicknTie is an open-source web application for auditing Excel files with image attachments. It provides a simple image viewer sidebar plugin for Univer spreadsheets, allowing users to attach files to cells via "pushpin" markers and view previews in a side panel.

## Development Commands

```bash
# Install dependencies
npm install

# Start development server (http://localhost:5173)
npm run dev

# Build for production (outputs to dist/)
npm run build

# Preview production build
npm run preview
```

## Architecture

### Technology Stack
- **Vite**: Build tool and development server
- **Univer**: Excel-like spreadsheet functionality (@univerjs/presets)
- **PDF.js**: PDF preview rendering (loaded from CDN)
- **Vanilla JavaScript**: No framework dependencies

### Core Components

1. **main.js**: Application entry point
   - Initializes Univer spreadsheet with 100 columns × 1000 rows
   - Sets up English locale and default theme
   - Creates ImagePlugin instance and makes it globally available at `window.ticknTie`

2. **image-plugin.js**: Hyperlink-based image sidebar plugin
   - Stores files as data URLs in `cellHyperlinks` Map (cellRef → {url, displayText})
   - Implements smart caching with LRU eviction (100 preview limit, 30-minute TTL)
   - Preloads adjacent cells (3×3 grid) for instant preview
   - Handles multiple selection detection methods (API, keyboard, mouse, DOM observation)
   - Supports image compression and PDF first-page preview

3. **index.html**: Main template
   - Two-panel layout: `#spreadsheet` (left) and `#sidebar` (right)
   - Preloads PDF.js library from CDN

### Key Features

- **File Attachment**: Click "➕ Add Evidence" → native file picker → data URL storage
- **Cell Markers**: Attached files show as "📌 filename.ext" in cells
- **Preview Caching**: Up to 100 previews cached with automatic cleanup
- **Adjacent Preloading**: Preloads 3×3 grid around selected cell
- **Resizable Sidebar**: Drag handle or double-click to resize/collapse
- **Auto-expand**: Sidebar expands when selecting cell with attachment
- **Cell Deletion Detection**: Removes hyperlinks when cells are cleared

### Data Storage

- **Session-based**: All data stored in browser memory during session
- **Cell References**: Standard Excel notation (A1, B2, etc.)
- **File Format Support**: Images (JPG, PNG, GIF, WebP) and PDFs
- **Size Limits**: 50MB per file, images compressed to 1920×1080 max

## Working with Univer API

The Univer spreadsheet instance is accessed via `this.univerAPI` in the plugin. Common operations:

```javascript
// Get active workbook and sheet
const workbook = univerAPI.getActiveWorkbook()
const sheet = workbook.getActiveSheet()

// Get/set cell values
const range = sheet.getRange(row, col)
range.setValue(value)
const value = range.getValue()

// Cell selection detection attempts multiple methods due to API limitations
```

## Performance Optimizations

- **Preview Cache**: Map with LRU eviction and time-based cleanup
- **Preload Queue**: Background loading of adjacent cell previews
- **Image Compression**: Resizes to max 1920×1080, JPEG 85% quality
- **DOM Caching**: Frequently accessed elements cached in `domCache` Map
- **Smart Polling**: Selection detection only when document visible

## Current Limitations

- Hyperlinks stored as text with 📌 icon (true hyperlink API not yet available in Univer)
- Files lost on page refresh (session storage only)
- Excel export shows text markers, not embedded images
- CORS restrictions may affect external image URLs