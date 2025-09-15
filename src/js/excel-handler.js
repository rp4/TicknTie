/**
 * Excel Handler Module
 * Manages Excel file operations with image references
 */

class ExcelHandler {
    constructor() {
        this.workbook = null;
        this.sheets = [];
        this.currentSheet = 0;
    }

    /**
     * Load an Excel file
     */
    async loadExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    this.workbook = XLSX.read(data, { type: 'array' });
                    this.parseWorkbook();
                    resolve(this.sheets);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Parse workbook into internal format
     */
    parseWorkbook() {
        this.sheets = [];
        
        this.workbook.SheetNames.forEach((sheetName) => {
            const worksheet = this.workbook.Sheets[sheetName];
            const sheetData = this.parseSheet(worksheet);
            
            this.sheets.push({
                name: sheetName,
                data: sheetData.cells,
                range: sheetData.range,
                formatting: sheetData.formatting
            });
        });
    }

    /**
     * Parse a single sheet
     */
    parseSheet(worksheet) {
        const cells = {};
        const formatting = {};
        let minRow = Infinity, maxRow = -Infinity;
        let minCol = Infinity, maxCol = -Infinity;
        
        // Parse all cells
        for (const cellRef in worksheet) {
            if (cellRef[0] === '!') continue; // Skip special properties
            
            const cell = worksheet[cellRef];
            const { col, row } = this.parseCellRef(cellRef);
            
            // Track range
            minRow = Math.min(minRow, row);
            maxRow = Math.max(maxRow, row);
            minCol = Math.min(minCol, col);
            maxCol = Math.max(maxCol, col);
            
            // Store cell value
            cells[cellRef] = {
                value: cell.v,
                type: cell.t,
                formula: cell.f,
                format: cell.z,
                style: cell.s
            };
            
            // Check if this is a hyperlink to an image
            if (worksheet[cellRef + '!l'] && worksheet[cellRef + '!l'].Target) {
                const target = worksheet[cellRef + '!l'].Target;
                if (this.isImagePath(target)) {
                    const imageId = target.split('/').pop();
                    cells[cellRef].imageRef = imageId;
                    cells[cellRef].hyperlink = target;
                }
            }
            
            // Store formatting if present
            if (cell.s) {
                formatting[cellRef] = this.parseStyle(cell.s);
            }
        }
        
        return {
            cells,
            formatting,
            range: {
                minRow: minRow === Infinity ? 1 : minRow,
                maxRow: maxRow === -Infinity ? 100 : maxRow,
                minCol: minCol === Infinity ? 1 : minCol,
                maxCol: maxCol === -Infinity ? 26 : maxCol
            }
        };
    }

    /**
     * Check if a path is an image path
     */
    isImagePath(path) {
        const ext = path.split('.').pop().toLowerCase();
        return ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'].includes(ext);
    }

    /**
     * Parse cell reference to column and row
     */
    parseCellRef(cellRef) {
        const match = cellRef.match(/^([A-Z]+)(\d+)$/);
        if (!match) return { col: 1, row: 1 };
        
        const col = this.columnToNumber(match[1]);
        const row = parseInt(match[2]);
        
        return { col, row };
    }

    /**
     * Convert column letter to number
     */
    columnToNumber(column) {
        let result = 0;
        for (let i = 0; i < column.length; i++) {
            result = result * 26 + (column.charCodeAt(i) - 64);
        }
        return result;
    }

    /**
     * Convert number to column letter
     */
    numberToColumn(num) {
        let column = '';
        while (num > 0) {
            const remainder = (num - 1) % 26;
            column = String.fromCharCode(65 + remainder) + column;
            num = Math.floor((num - 1) / 26);
        }
        return column;
    }

    /**
     * Create cell reference from row and column
     */
    createCellRef(row, col) {
        return this.numberToColumn(col) + row;
    }

    /**
     * Parse style object
     */
    parseStyle(style) {
        return {
            bold: style.font?.bold || false,
            italic: style.font?.italic || false,
            underline: style.font?.underline || false,
            fontSize: style.font?.sz || 12,
            fontColor: style.font?.color?.rgb || '000000',
            bgColor: style.fill?.fgColor?.rgb || 'FFFFFF',
            alignment: style.alignment?.horizontal || 'left'
        };
    }

    /**
     * Update cell value
     */
    updateCellValue(sheetIndex, cellRef, value) {
        if (!this.sheets[sheetIndex]) return;
        
        if (!this.sheets[sheetIndex].data[cellRef]) {
            this.sheets[sheetIndex].data[cellRef] = {};
        }
        
        // Note: Image references are now handled via hyperlinks, not cell values
        
        this.sheets[sheetIndex].data[cellRef].value = value;
    }

    /**
     * Set image hyperlink for a cell
     */
    setCellImageRef(sheetIndex, cellRef, imageId) {
        if (!this.sheets[sheetIndex]) return;

        if (!this.sheets[sheetIndex].data[cellRef]) {
            this.sheets[sheetIndex].data[cellRef] = {};
        }

        // Store as hyperlink reference
        this.sheets[sheetIndex].data[cellRef].imageRef = imageId;
        this.sheets[sheetIndex].data[cellRef].hyperlink = `images/${imageId}`;
        this.sheets[sheetIndex].data[cellRef].value = imageId; // Display filename in cell
    }

    /**
     * Export to Excel with image references
     */
    exportExcel() {
        const wb = XLSX.utils.book_new();
        
        this.sheets.forEach((sheet) => {
            const ws = this.createWorksheet(sheet);
            XLSX.utils.book_append_sheet(wb, ws, sheet.name);
        });
        
        // Generate Excel file
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        return new Blob([wbout], { type: 'application/octet-stream' });
    }

    /**
     * Create worksheet from sheet data
     */
    createWorksheet(sheet) {
        const ws = {};
        
        // Add cells
        for (const cellRef in sheet.data) {
            const cellData = sheet.data[cellRef];
            const cell = {
                v: cellData.value,
                t: cellData.type || 's'
            };
            
            // Add formula if present
            if (cellData.formula) {
                cell.f = cellData.formula;
            }
            
            // Add hyperlink to image if present
            if (cellData.imageRef) {
                cell.v = cellData.imageRef; // Display filename
                cell.t = 's';
                // Add hyperlink
                cell.l = {
                    Target: `images/${cellData.imageRef}`,
                    Tooltip: `Image: ${cellData.imageRef}`
                };
            }
            
            ws[cellRef] = cell;
        }
        
        // Set range
        const maxCell = this.createCellRef(
            sheet.range.maxRow || 100,
            sheet.range.maxCol || 26
        );
        ws['!ref'] = `A1:${maxCell}`;
        
        return ws;
    }

    /**
     * Create a new sheet
     */
    createNewSheet(name) {
        const sheetName = name || `Sheet${this.sheets.length + 1}`;
        
        this.sheets.push({
            name: sheetName,
            data: {},
            range: {
                minRow: 1,
                maxRow: 100,
                minCol: 1,
                maxCol: 26
            },
            formatting: {}
        });
        
        return this.sheets.length - 1;
    }

    /**
     * Delete a sheet
     */
    deleteSheet(index) {
        if (this.sheets.length <= 1) {
            throw new Error('Cannot delete the last sheet');
        }
        
        this.sheets.splice(index, 1);
        
        if (this.currentSheet >= this.sheets.length) {
            this.currentSheet = this.sheets.length - 1;
        }
    }

    /**
     * Get sheet by index
     */
    getSheet(index) {
        return this.sheets[index] || null;
    }

    /**
     * Get current sheet
     */
    getCurrentSheet() {
        return this.sheets[this.currentSheet] || null;
    }

    /**
     * Set current sheet
     */
    setCurrentSheet(index) {
        if (index >= 0 && index < this.sheets.length) {
            this.currentSheet = index;
            return true;
        }
        return false;
    }

    /**
     * Clear all data
     */
    clear() {
        this.workbook = null;
        this.sheets = [];
        this.currentSheet = 0;
    }
}

// Export as singleton
const excelHandler = new ExcelHandler();

// Make it available globally
window.ExcelHandler = excelHandler;