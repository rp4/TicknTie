/**
 * Main Application Module
 * Coordinates all modules and handles application flow
 */

class TicknTieApp {
    constructor() {
        this.initialized = false;
        this.currentFile = null;
    }

    /**
     * Initialize the application
     */
    async init() {
        if (this.initialized) return;

        console.log('Initializing TicknTie Application...');

        // Check for pending data from landing page
        await this.checkForPendingData();

        // Set up event listeners
        this.setupEventListeners();

        // Initialize with empty spreadsheet if no data was loaded
        if (!window.ExcelHandler.sheets || window.ExcelHandler.sheets.length === 0) {
            this.initializeEmptySpreadsheet();
        }

        // Render UI components
        window.uiComponents.renderImagePanel();

        this.initialized = true;
        console.log('TicknTie Application initialized successfully');
    }

    /**
     * Check for pending data from landing page
     */
    async checkForPendingData() {
        // Check for pending zip file
        const pendingZipData = sessionStorage.getItem('pendingZipFile');
        const pendingZipFileName = sessionStorage.getItem('pendingZipFileName');

        if (pendingZipData && pendingZipFileName) {
            // Clear from session storage
            sessionStorage.removeItem('pendingZipFile');
            sessionStorage.removeItem('pendingZipFileName');

            // Convert data URL back to blob
            const response = await fetch(pendingZipData);
            const blob = await response.blob();
            const file = new File([blob], pendingZipFileName, { type: 'application/zip' });

            // Load the zip file
            await this.loadZipFile(file);
            return;
        }

        // Check for sample data flag
        const sampleDataStr = sessionStorage.getItem('loadSampleData');

        if (sampleDataStr) {
            // Clear from session storage
            sessionStorage.removeItem('loadSampleData');

            const sampleData = JSON.parse(sampleDataStr);

            // Create sample workbook
            if (sampleData.type === 'sample' && sampleData.sheets) {
                // Clear any existing sheets
                window.ExcelHandler.sheets = [];

                // Create sheets from sample data
                sampleData.sheets.forEach((sheetData, index) => {
                    const sheetIndex = window.ExcelHandler.createNewSheet(sheetData.name || `Sheet${index + 1}`);
                    const sheet = window.ExcelHandler.sheets[sheetIndex];

                    // Populate cells with sample data
                    if (sheetData.data && Array.isArray(sheetData.data)) {
                        sheetData.data.forEach((row, rowIndex) => {
                            row.forEach((value, colIndex) => {
                                if (value) {
                                    const colLetter = String.fromCharCode(65 + colIndex);
                                    const cellRef = `${colLetter}${rowIndex + 1}`;
                                    window.ExcelHandler.updateCellValue(sheetIndex, cellRef, value);
                                }
                            });
                        });
                    }
                });

                // Set first sheet as current
                window.ExcelHandler.currentSheet = 0;

                // Hide upload area and show spreadsheet
                const container = document.getElementById('spreadsheetContent');
                const uploadArea = document.getElementById('uploadArea');

                if (container && uploadArea) {
                    uploadArea.style.display = 'none';
                    container.style.display = 'block';
                }

                // Render first sheet
                this.refreshCurrentSheet();
                this.updateSheetTabs();

                window.uiComponents.updateStatusBar('Sample data loaded');
            }
        }
    }

    /**
     * Set up all event listeners
     */
    setupEventListeners() {
        // File upload handlers
        const fileInput = document.getElementById('fileInput');
        const imageInput = document.getElementById('imageInput');
        const zipInput = document.getElementById('zipInput');
        const uploadArea = document.getElementById('uploadArea');
        
        if (fileInput) {
            fileInput.addEventListener('change', (e) => this.handleFileUpload(e));
        }
        
        if (imageInput) {
            imageInput.addEventListener('change', (e) => this.handleImageUpload(e));
        }
        
        if (zipInput) {
            zipInput.addEventListener('change', (e) => this.handleZipUpload(e));
        }
        
        // Drag and drop for upload area
        if (uploadArea) {
            uploadArea.addEventListener('click', () => {
                const input = document.getElementById('fileInput') || document.getElementById('zipInput');
                if (input) input.click();
            });
            
            uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                uploadArea.classList.add('dragover');
            });
            
            uploadArea.addEventListener('dragleave', () => {
                uploadArea.classList.remove('dragover');
            });
            
            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('dragover');
                this.handleFileDrop(e);
            });
        }
        
        // Formula bar
        const formulaBar = document.getElementById('formulaBar');
        if (formulaBar) {
            formulaBar.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    this.updateCurrentCellFromFormulaBar();
                }
            });
        }
        
        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => this.handleKeyboardShortcuts(e));
    }

    /**
     * Initialize empty spreadsheet
     */
    initializeEmptySpreadsheet() {
        // Create default sheet
        window.ExcelHandler.createNewSheet('Sheet1');
        
        // Render the sheet
        const container = document.getElementById('spreadsheetContent');
        const uploadArea = document.getElementById('uploadArea');
        
        if (container && uploadArea) {
            uploadArea.style.display = 'none';
            container.style.display = 'block';
            
            const sheet = window.ExcelHandler.getCurrentSheet();
            window.uiComponents.renderSheet(sheet, container);
        }
        
        window.uiComponents.updateStatusBar('Ready');
    }

    /**
     * Handle file upload (Excel or Zip)
     */
    async handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        window.uiComponents.showLoading();
        
        try {
            if (file.name.endsWith('.zip')) {
                await this.loadZipFile(file);
            } else {
                await this.loadExcelFile(file);
            }
        } catch (error) {
            console.error('Error loading file:', error);
            alert('Error loading file: ' + error.message);
        } finally {
            window.uiComponents.hideLoading();
        }
    }

    /**
     * Handle image upload
     */
    async handleImageUpload(event) {
        const files = Array.from(event.target.files);
        if (!files.length) return;
        
        window.uiComponents.showLoading();
        
        try {
            for (const file of files) {
                const imageId = await window.ImageManager.storeImage(file);
                
                // If we have a current cell selected, link the image
                if (window.uiComponents.currentCell) {
                    window.ImageManager.linkImageToCell(window.uiComponents.currentCell, imageId);
                    window.ExcelHandler.setCellImageRef(
                        window.ExcelHandler.currentSheet,
                        window.uiComponents.currentCell,
                        imageId
                    );
                    
                    // Re-render the sheet to show the image
                    this.refreshCurrentSheet();
                }
            }
            
            // Update image panel
            window.uiComponents.renderImagePanel();
            window.uiComponents.updateStatusBar(`Added ${files.length} image(s)`);
        } catch (error) {
            console.error('Error uploading images:', error);
            alert('Error uploading images: ' + error.message);
        } finally {
            window.uiComponents.hideLoading();
        }
    }

    /**
     * Handle zip file upload
     */
    async handleZipUpload(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        await this.loadZipFile(file);
    }

    /**
     * Handle file drop
     */
    async handleFileDrop(event) {
        const files = Array.from(event.dataTransfer.files);
        if (!files.length) return;
        
        const file = files[0];
        
        if (file.name.endsWith('.zip')) {
            await this.loadZipFile(file);
        } else if (file.name.match(/\.(xlsx?|xlsm|xlsb)$/i)) {
            await this.loadExcelFile(file);
        } else if (file.type.startsWith('image/')) {
            await this.handleImageUpload({ target: { files } });
        }
    }

    /**
     * Load Excel file
     */
    async loadExcelFile(file) {
        window.uiComponents.showLoading();
        
        try {
            await window.ExcelHandler.loadExcelFile(file);
            this.currentFile = file.name;
            
            // Hide upload area and show spreadsheet
            const container = document.getElementById('spreadsheetContent');
            const uploadArea = document.getElementById('uploadArea');
            
            if (container && uploadArea) {
                uploadArea.style.display = 'none';
                container.style.display = 'block';
            }
            
            // Render first sheet
            this.refreshCurrentSheet();
            this.updateSheetTabs();
            
            window.uiComponents.updateStatusBar(`Loaded: ${file.name}`);
        } catch (error) {
            console.error('Error loading Excel file:', error);
            throw error;
        } finally {
            window.uiComponents.hideLoading();
        }
    }

    /**
     * Load zip file
     */
    async loadZipFile(file) {
        window.uiComponents.showLoading();
        
        try {
            const result = await window.ZipHandler.loadZipFile(file);
            
            if (!result.excelFile) {
                throw new Error('No Excel file found in zip');
            }
            
            // Load Excel file
            await this.loadExcelFile(result.excelFile);
            
            // Import images
            if (result.imageFiles && Object.keys(result.imageFiles).length > 0) {
                const imageMappings = await window.ImageManager.importImages(result.imageFiles);
                
                // Apply mappings to cells if provided
                if (result.imageMappings) {
                    for (const [cellRef, originalImageId] of Object.entries(result.imageMappings)) {
                        const newImageId = imageMappings[originalImageId];
                        if (newImageId) {
                            window.ImageManager.linkImageToCell(cellRef, newImageId);
                        }
                    }
                }
                
                window.uiComponents.renderImagePanel();
            }
            
            window.uiComponents.updateStatusBar(`Loaded zip: ${file.name}`);
        } catch (error) {
            console.error('Error loading zip file:', error);
            alert('Error loading zip file: ' + error.message);
        } finally {
            window.uiComponents.hideLoading();
        }
    }

    /**
     * Export as Excel file
     */
    async exportExcel() {
        window.uiComponents.showLoading();
        
        try {
            const blob = window.ExcelHandler.exportExcel();
            const filename = this.currentFile || 'workbook.xlsx';
            
            this.downloadFile(blob, filename);
            window.uiComponents.updateStatusBar('Excel file exported');
        } catch (error) {
            console.error('Error exporting Excel:', error);
            alert('Error exporting Excel: ' + error.message);
        } finally {
            window.uiComponents.hideLoading();
        }
    }

    /**
     * Export as zip bundle
     */
    async exportZipBundle() {
        window.uiComponents.showLoading();
        
        try {
            // Get Excel blob
            const excelBlob = window.ExcelHandler.exportExcel();
            
            // Get all images
            const images = await window.ImageManager.exportImages();
            
            // Get cell mappings
            const cellMappings = window.ImageManager.getAllCellMappings();
            
            // Create zip bundle
            const zipBlob = await window.ZipHandler.createZipBundle(
                excelBlob,
                images,
                cellMappings
            );
            
            // Download zip
            const filename = window.ZipHandler.generateFilename(
                this.currentFile ? this.currentFile.replace(/\.[^.]+$/, '') : 'tickntie'
            );
            
            window.ZipHandler.downloadZip(zipBlob, filename);
            window.uiComponents.updateStatusBar('Zip bundle exported');
        } catch (error) {
            console.error('Error exporting zip bundle:', error);
            alert('Error exporting zip bundle: ' + error.message);
        } finally {
            window.uiComponents.hideLoading();
        }
    }

    /**
     * Refresh current sheet display
     */
    refreshCurrentSheet() {
        const container = document.getElementById('spreadsheetContent');
        const sheet = window.ExcelHandler.getCurrentSheet();
        
        if (container && sheet) {
            window.uiComponents.renderSheet(sheet, container);
        }
    }

    /**
     * Update sheet tabs
     */
    updateSheetTabs() {
        const tabsContainer = document.getElementById('sheetTabs');
        if (!tabsContainer) return;
        
        tabsContainer.innerHTML = '';
        
        window.ExcelHandler.sheets.forEach((sheet, index) => {
            const tab = document.createElement('div');
            tab.className = 'sheet-tab';
            if (index === window.ExcelHandler.currentSheet) {
                tab.classList.add('active');
            }
            tab.textContent = sheet.name;
            tab.onclick = () => this.switchSheet(index);
            
            tabsContainer.appendChild(tab);
        });
        
        // Add new sheet button
        const addButton = document.createElement('button');
        addButton.textContent = '+';
        addButton.className = 'sheet-tab';
        addButton.style.minWidth = '30px';
        addButton.onclick = () => this.addNewSheet();
        tabsContainer.appendChild(addButton);
    }

    /**
     * Switch to a different sheet
     */
    switchSheet(index) {
        if (window.ExcelHandler.setCurrentSheet(index)) {
            this.refreshCurrentSheet();
            this.updateSheetTabs();
        }
    }

    /**
     * Add a new sheet
     */
    addNewSheet() {
        const name = prompt('Enter sheet name:');
        if (name) {
            const index = window.ExcelHandler.createNewSheet(name);
            this.switchSheet(index);
        }
    }

    /**
     * Update current cell from formula bar
     */
    updateCurrentCellFromFormulaBar() {
        const formulaBar = document.getElementById('formulaBar');
        if (!formulaBar || !window.uiComponents.currentCell) return;
        
        const value = formulaBar.value;
        const cellRef = window.uiComponents.currentCell;
        
        window.ExcelHandler.updateCellValue(
            window.ExcelHandler.currentSheet,
            cellRef,
            value
        );
        
        // Update cell display
        const input = document.querySelector(`input[data-cell="${cellRef}"]`);
        if (input) {
            input.value = value;
        }
    }

    /**
     * Handle keyboard shortcuts
     */
    handleKeyboardShortcuts(event) {
        // Ctrl/Cmd + S: Save/Export
        if ((event.ctrlKey || event.metaKey) && event.key === 's') {
            event.preventDefault();
            this.exportZipBundle();
        }
        
        // Ctrl/Cmd + O: Open file
        if ((event.ctrlKey || event.metaKey) && event.key === 'o') {
            event.preventDefault();
            const input = document.getElementById('fileInput') || document.getElementById('zipInput');
            if (input) input.click();
        }
    }

    /**
     * Download file helper
     */
    downloadFile(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    /**
     * Close image modal
     */
    closeImageModal() {
        const modal = document.getElementById('imageModal');
        if (modal) {
            modal.style.display = 'none';
        }
    }
}

// Create and initialize app instance
const app = new TicknTieApp();

// Initialize when DOM is ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => app.init());
} else {
    app.init();
}

// Export for global access
window.TicknTieApp = app;