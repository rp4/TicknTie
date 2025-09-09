# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

TicknTie is a single-page web application for auditing Excel files with embedded images, designed in the style of DataSnipper. The entire application is contained in `audit-ticktie-enhanced.html`.

## Application Architecture

### Structure
- **Single File Architecture**: Everything is in `audit-ticktie-enhanced.html`
  - CSS: Lines 8-527
  - HTML: Lines 529-664
  - JavaScript: Lines 665-1848

### Core Data Structures
- `workbook`: Main object storing all Excel data
- `sheets`: Array of sheet objects with cells and metadata
- `currentSheet`: Index of active sheet
- `cellImages`: Map of cell references to image data
- `selectedCells`: Array tracking multi-cell selection

### Key Functionality Areas
1. **Excel Processing**: Uses SheetJS library loaded from CDN
2. **Image Handling**: Base64 encoding for embedded images in cells
3. **Cell Rendering**: Dynamic grid generation with formatting support
4. **Event Management**: Click, keyboard, and drag events for cell interaction

## Development Commands

Since this is a standalone HTML file with no build system:

```bash
# Open the application
open audit-ticktie-enhanced.html  # macOS
xdg-open audit-ticktie-enhanced.html  # Linux
start audit-ticktie-enhanced.html  # Windows

# No build, test, or lint commands - direct browser development
```

## Key Functions and Their Locations

- `handleFileUpload()` (line ~700): Processes Excel file uploads
- `renderSheet()` (line ~800): Renders spreadsheet grid
- `updateCell()` (line ~950): Handles cell value changes
- `embedImage()` (line ~1100): Embeds images in cells
- `exportFile()` (line ~1400): Exports modified Excel with images
- `createNewSheet()` (line ~1300): Adds new worksheet
- `switchSheet()` (line ~1350): Changes active worksheet

## Working with Features

### Adding New Cell Formatting
Cell formatting functions are in the toolbar event handlers section (lines ~1500-1700). Format changes update the global `sheets` array and trigger `renderSheet()`.

### Modifying Excel Import/Export
Excel processing uses SheetJS library. Import logic starts at `handleFileUpload()`, export at `exportFile()`. Images are handled separately through `cellImages` map.

### Extending Toolbar Functions
Toolbar buttons are defined in HTML (lines ~530-580) with event listeners added in JavaScript initialization (lines ~1650-1750).

## Important Considerations

1. **Browser-Only**: No server-side code or API endpoints
2. **Memory Usage**: All data stored in browser memory - large files may cause performance issues
3. **Image Storage**: Images stored as base64 strings in `cellImages` map
4. **No Persistence**: Refreshing loses all unsaved work
5. **Excel Compatibility**: Uses SheetJS library limitations apply to formula support