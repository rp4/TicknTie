/**
 * UI Components Module
 * Handles all UI rendering and interactions
 */

class UIComponents {
    constructor() {
        this.currentCell = null;
        this.selectedCells = [];
        this.imagePanelTab = 'preview'; // 'preview' or 'gallery'
        // Resize tracking
        this.isResizingColumn = false;
        this.isResizingRow = false;
        this.resizeStartX = 0;
        this.resizeStartY = 0;
        this.resizingColumn = null;
        this.resizingRow = null;
        this.originalWidth = 0;
        this.originalHeight = 0;
    }

    /**
     * Render the spreadsheet grid
     */
    renderSheet(sheet, container) {
        if (!sheet) return;
        
        const table = document.createElement('table');
        table.className = 'spreadsheet-table';
        
        // Calculate grid size
        const maxRow = sheet.range?.maxRow || 100;
        const maxCol = sheet.range?.maxCol || 26;
        
        // Create header row
        const headerRow = document.createElement('tr');
        headerRow.appendChild(this.createHeaderCell('', 'corner'));
        
        for (let col = 1; col <= maxCol; col++) {
            const colHeader = this.createHeaderCell(
                window.ExcelHandler.numberToColumn(col),
                'col-header',
                col
            );
            headerRow.appendChild(colHeader);
        }
        table.appendChild(headerRow);
        
        // Create data rows
        for (let row = 1; row <= maxRow; row++) {
            const tr = document.createElement('tr');
            
            // Row header
            const rowHeader = this.createHeaderCell(row.toString(), 'row-header', row);
            tr.appendChild(rowHeader);

            // Set row height
            const height = window.TicknTieApp.rowHeights[row] || window.TicknTieApp.defaultRowHeight;
            tr.style.height = height + 'px';
            
            // Data cells
            for (let col = 1; col <= maxCol; col++) {
                const cellRef = window.ExcelHandler.createCellRef(row, col);
                const cellData = sheet.data[cellRef] || {};
                const td = this.createDataCell(cellRef, cellData);

                // Set column width
                const width = window.TicknTieApp.columnWidths[col] || window.TicknTieApp.defaultColumnWidth;
                td.style.width = width + 'px';
                td.style.minWidth = width + 'px';
                td.style.maxWidth = width + 'px';

                tr.appendChild(td);
            }
            
            table.appendChild(tr);
        }
        
        // Clear container and add table
        container.innerHTML = '';
        container.appendChild(table);
    }

    /**
     * Create header cell
     */
    createHeaderCell(content, className, index = null) {
        const th = document.createElement('th');
        th.className = className;

        // Create wrapper for content and resize handle
        const wrapper = document.createElement('div');
        wrapper.className = 'header-wrapper';

        // Add content
        const contentSpan = document.createElement('span');
        contentSpan.textContent = content;
        wrapper.appendChild(contentSpan);

        // Add resize handle for column headers
        if (className === 'col-header' && index !== null) {
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'resize-handle resize-handle-col';
            resizeHandle.dataset.column = index;

            // Set width based on stored or default value
            const width = window.TicknTieApp.columnWidths[index] || window.TicknTieApp.defaultColumnWidth;
            th.style.width = width + 'px';
            th.style.minWidth = width + 'px';
            th.style.maxWidth = width + 'px';

            // Add resize event handlers
            resizeHandle.addEventListener('mousedown', (e) => this.startColumnResize(e, index));
            wrapper.appendChild(resizeHandle);
        }

        // Add resize handle for row headers
        if (className === 'row-header' && index !== null) {
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'resize-handle resize-handle-row';
            resizeHandle.dataset.row = index;

            // Set height based on stored or default value
            const height = window.TicknTieApp.rowHeights[index] || window.TicknTieApp.defaultRowHeight;
            th.style.height = height + 'px';
            th.parentElement && (th.parentElement.style.height = height + 'px');

            // Add resize event handlers
            resizeHandle.addEventListener('mousedown', (e) => this.startRowResize(e, index));
            wrapper.appendChild(resizeHandle);
        }

        th.appendChild(wrapper);
        return th;
    }

    /**
     * Create data cell
     */
    createDataCell(cellRef, cellData) {
        const td = document.createElement('td');
        td.dataset.cell = cellRef;

        // Check for image reference
        const imageData = window.ImageManager.getImageForCell(cellRef);

        // Create a container for cell content
        const cellContent = document.createElement('div');
        cellContent.className = 'cell-content-wrapper';
        cellContent.style.display = 'flex';
        cellContent.style.alignItems = 'center';
        cellContent.style.width = '100%';
        cellContent.style.height = '100%';

        if (imageData) {
            td.className = 'cell-with-image-ref';

            // Add pushpin icon on the LEFT
            const pushpin = this.createPushpinIcon(cellRef, imageData);
            pushpin.style.position = 'relative';
            pushpin.style.marginRight = '4px';
            cellContent.appendChild(pushpin);
        }

        // Add input element for text (shows filename or regular cell value)
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell-input';

        // If there's an image, show the filename as text
        if (imageData && cellData.imageRef) {
            input.value = cellData.imageRef || cellData.value || '';
        } else {
            input.value = cellData.value || '';
        }

        input.dataset.cell = cellRef;
        input.style.flex = '1';
        input.style.border = 'none';
        input.style.background = 'transparent';
        input.style.outline = 'none';

        // Apply formatting
        if (cellData.formatting) {
            this.applyCellFormatting(input, cellData.formatting);
        }

        // Add event listeners
        input.addEventListener('focus', () => this.handleCellFocus(cellRef));
        input.addEventListener('blur', () => this.handleCellBlur(cellRef, input.value));
        input.addEventListener('input', () => this.handleCellInput(cellRef, input.value));

        // Add click handler to cell to ensure panel opens when clicking anywhere in the cell
        td.addEventListener('click', (e) => {
            // Focus the input if not already focused
            if (document.activeElement !== input) {
                input.focus();
            }
            // If cell has image, ensure panel shows it
            if (imageData) {
                this.showImageInPanel(imageData);
            }
        });

        cellContent.appendChild(input);
        td.appendChild(cellContent);

        return td;
    }

    /**
     * Create pushpin icon for cell with image
     */
    createPushpinIcon(cellRef, imageData) {
        const container = document.createElement('div');
        container.className = 'pushpin-container';

        // Get stored color for this cell or use default
        const color = window.ImageManager.getPushpinColor(cellRef) || '#FF4444';

        // Create SVG pushpin icon
        const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        svg.setAttribute('width', '20');
        svg.setAttribute('height', '20');
        svg.setAttribute('viewBox', '0 0 24 24');
        svg.setAttribute('class', 'pushpin-icon');

        // Pushpin path
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('d', 'M16,12V4H17V2H7V4H8V12L6,14V16H11.2V22H12.8V16H18V14L16,12Z');
        path.setAttribute('fill', color);
        path.setAttribute('class', 'pushpin-path');

        svg.appendChild(path);
        container.appendChild(svg);

        // Click handler to show image
        container.onclick = (e) => {
            e.stopPropagation();
            this.showImageInPanel(imageData);
        };

        // Right-click to change color
        container.oncontextmenu = (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.showPushpinColorPicker(cellRef, path, color);
        };

        return container;
    }

    /**
     * Show color picker for pushpin
     */
    showPushpinColorPicker(cellRef, pathElement, currentColor) {
        // Remove any existing color picker
        const existingPicker = document.querySelector('.pushpin-color-picker');
        if (existingPicker) {
            existingPicker.remove();
        }

        // Create color picker popup
        const picker = document.createElement('div');
        picker.className = 'pushpin-color-picker';

        // Predefined colors
        const colors = [
            '#FF4444', // Red
            '#44FF44', // Green
            '#4444FF', // Blue
            '#FFAA00', // Orange
            '#FF00FF', // Magenta
            '#00FFFF', // Cyan
            '#FFFF00', // Yellow
            '#8B4513', // Brown
            '#800080', // Purple
            '#000000'  // Black
        ];

        colors.forEach(color => {
            const colorBtn = document.createElement('button');
            colorBtn.className = 'color-option';
            colorBtn.style.backgroundColor = color;
            if (color === currentColor) {
                colorBtn.classList.add('selected');
            }

            colorBtn.onclick = () => {
                // Update the icon color
                pathElement.setAttribute('fill', color);

                // Store the color preference
                window.ImageManager.setPushpinColor(cellRef, color);

                // Remove picker
                picker.remove();
            };

            picker.appendChild(colorBtn);
        });

        // Position near the cell
        const cell = document.querySelector(`[data-cell="${cellRef}"]`);
        const rect = cell.getBoundingClientRect();
        picker.style.position = 'absolute';
        picker.style.left = rect.left + 'px';
        picker.style.top = (rect.bottom + 5) + 'px';
        picker.style.zIndex = '1000';

        document.body.appendChild(picker);

        // Close on click outside
        setTimeout(() => {
            document.addEventListener('click', function closePickerHandler() {
                picker.remove();
                document.removeEventListener('click', closePickerHandler);
            }, { once: true });
        }, 100);
    }

    /**
     * Apply cell formatting
     */
    applyCellFormatting(element, formatting) {
        if (formatting.bold) element.style.fontWeight = 'bold';
        if (formatting.italic) element.style.fontStyle = 'italic';
        if (formatting.underline) element.style.textDecoration = 'underline';
        if (formatting.fontSize) element.style.fontSize = formatting.fontSize + 'px';
        if (formatting.fontColor) element.style.color = '#' + formatting.fontColor;
        if (formatting.bgColor) element.style.backgroundColor = '#' + formatting.bgColor;
        if (formatting.alignment) element.style.textAlign = formatting.alignment;
    }

    /**
     * Handle cell focus
     */
    handleCellFocus(cellRef) {
        this.currentCell = cellRef;
        this.updateFormulaBar(cellRef);
        this.updateCellReference(cellRef);
        
        // Check if cell has image and show in panel
        const imageData = window.ImageManager.getImageForCell(cellRef);
        if (imageData) {
            this.showImageInPanel(imageData);
        }
    }

    /**
     * Handle cell blur
     */
    handleCellBlur(cellRef, value) {
        const sheetIndex = window.ExcelHandler.currentSheet;
        window.ExcelHandler.updateCellValue(sheetIndex, cellRef, value);
    }

    /**
     * Handle cell input
     */
    handleCellInput(cellRef, value) {
        this.updateFormulaBar(cellRef, value);
    }

    /**
     * Update formula bar
     */
    updateFormulaBar(cellRef, value) {
        const formulaBar = document.getElementById('formulaBar');
        if (formulaBar) {
            if (value !== undefined) {
                formulaBar.value = value;
            } else {
                const sheet = window.ExcelHandler.getCurrentSheet();
                const cellData = sheet?.data[cellRef];
                formulaBar.value = cellData?.value || '';
            }
        }
    }

    /**
     * Update cell reference display
     */
    updateCellReference(cellRef) {
        const cellRefDisplay = document.getElementById('cellRef');
        if (cellRefDisplay) {
            cellRefDisplay.textContent = cellRef;
        }
    }

    /**
     * Render image panel
     */
    renderImagePanel() {
        const panel = document.getElementById('imagePanel');
        if (!panel) return;

        // Check if panel is currently collapsed
        const isCollapsed = panel.classList.contains('collapsed');
        const iconDirection = isCollapsed ? '◀' : '▶';

        panel.innerHTML = `
            <button class="image-panel-toggle" onclick="UIComponents.toggleImagePanel()">
                <span id="imagePanelToggleIcon">${iconDirection}</span>
            </button>

            <div class="image-panel-header">
                <div class="image-panel-title">Image Viewer</div>
                <div class="image-panel-info" id="imagePanelInfo">
                    ${window.ImageManager.getStatistics().totalImages} images
                </div>
            </div>

            <div class="image-panel-tabs">
                <button class="image-panel-tab ${this.imagePanelTab === 'preview' ? 'active' : ''}"
                        onclick="UIComponents.switchPanelTab('preview')">
                    Preview
                </button>
                <button class="image-panel-tab ${this.imagePanelTab === 'gallery' ? 'active' : ''}"
                        onclick="UIComponents.switchPanelTab('gallery')">
                    Gallery
                </button>
            </div>

            <div class="image-panel-content" id="imagePanelContent">
                ${this.imagePanelTab === 'preview' ? this.renderImagePreview() : this.renderImageGallery()}
            </div>
        `;
    }

    /**
     * Render image preview
     */
    renderImagePreview(imageData) {
        if (!imageData) {
            return `
                <div style="text-align: center; padding: 40px; color: #95a5a6;">
                    <div style="font-size: 48px; margin-bottom: 15px;">🖼️</div>
                    <div>Select a cell with an image to preview</div>
                </div>
            `;
        }
        
        return `
            <div class="image-preview-container">
                <img src="${imageData.url}" alt="${imageData.originalName}" class="image-preview">
                <div class="image-details">
                    <div class="image-detail-row">
                        <span class="image-detail-label">Name:</span>
                        <span class="image-detail-value">${imageData.originalName}</span>
                    </div>
                    <div class="image-detail-row">
                        <span class="image-detail-label">Size:</span>
                        <span class="image-detail-value">${this.formatFileSize(imageData.size)}</span>
                    </div>
                    <div class="image-detail-row">
                        <span class="image-detail-label">Dimensions:</span>
                        <span class="image-detail-value">${imageData.width} × ${imageData.height}</span>
                    </div>
                    <div class="image-detail-row">
                        <span class="image-detail-label">Type:</span>
                        <span class="image-detail-value">${imageData.type}</span>
                    </div>
                </div>
                <div style="margin-top: 15px; display: flex; gap: 10px;">
                    <button class="format-button" onclick="UIComponents.viewFullImage('${imageData.id}')">
                        View Full Size
                    </button>
                    <button class="format-button" onclick="UIComponents.removeImageFromCell()">
                        Remove from Cell
                    </button>
                </div>
            </div>
        `;
    }

    /**
     * Render image gallery
     */
    renderImageGallery() {
        const images = window.ImageManager.getAllImages();
        
        if (images.length === 0) {
            return `
                <div style="text-align: center; padding: 40px; color: #95a5a6;">
                    <div style="font-size: 48px; margin-bottom: 15px;">📁</div>
                    <div>No images in workbook</div>
                    <div style="font-size: 12px; margin-top: 10px;">Add images to cells to see them here</div>
                </div>
            `;
        }
        
        let html = '<div class="image-gallery">';
        
        images.forEach(image => {
            // Find which cells use this image
            const cells = [];
            window.ImageManager.cellImageRefs.forEach((imgId, cellRef) => {
                if (imgId === image.id) cells.push(cellRef);
            });
            
            html += `
                <div class="gallery-item" onclick="UIComponents.selectGalleryImage('${image.id}')">
                    <img src="${image.url}" alt="${image.originalName}">
                    ${cells.length > 0 ? `<div class="gallery-item-cell">${cells.join(', ')}</div>` : ''}
                </div>
            `;
        });
        
        html += '</div>';
        return html;
    }

    /**
     * Show image in panel
     */
    showImageInPanel(imageData) {
        // Switch to preview tab
        this.imagePanelTab = 'preview';

        // Ensure panel is visible first
        const panel = document.getElementById('imagePanel');
        if (panel && panel.classList.contains('collapsed')) {
            UIComponents.toggleImagePanel();
        }

        // Update the tab buttons to show preview as active
        const tabs = panel.querySelectorAll('.image-panel-tab');
        tabs.forEach((tab, index) => {
            if (index === 0) { // Preview tab
                tab.classList.add('active');
            } else { // Gallery tab
                tab.classList.remove('active');
            }
        });

        // Update content to show the image
        const content = document.getElementById('imagePanelContent');
        if (content) {
            content.innerHTML = this.renderImagePreview(imageData);
        }
    }

    /**
     * Toggle image panel
     */
    static toggleImagePanel() {
        const panel = document.getElementById('imagePanel');
        const icon = document.getElementById('imagePanelToggleIcon');
        
        if (panel) {
            panel.classList.toggle('collapsed');
            if (icon) {
                icon.textContent = panel.classList.contains('collapsed') ? '◀' : '▶';
            }
        }
    }

    /**
     * Switch panel tab
     */
    static switchPanelTab(tab) {
        const ui = window.uiComponents || new UIComponents();
        ui.imagePanelTab = tab;
        ui.renderImagePanel();
    }

    /**
     * Select gallery image
     */
    static selectGalleryImage(imageId) {
        const imageData = window.ImageManager.getImageData(imageId);
        if (imageData) {
            const ui = window.uiComponents || new UIComponents();
            ui.showImageInPanel(imageData);
        }
    }

    /**
     * View full image
     */
    static viewFullImage(imageId) {
        const imageData = window.ImageManager.getImageData(imageId);
        if (imageData) {
            const modal = document.getElementById('imageModal');
            const modalImg = document.getElementById('modalImage');
            
            if (modal && modalImg) {
                modalImg.src = imageData.url;
                modal.style.display = 'flex';
            }
        }
    }

    /**
     * Remove image from current cell
     */
    static removeImageFromCell() {
        const ui = window.uiComponents || new UIComponents();
        if (ui.currentCell) {
            window.ImageManager.removeImageFromCell(ui.currentCell);
            // Re-render the cell
            const cell = document.querySelector(`[data-cell="${ui.currentCell}"]`);
            if (cell) {
                cell.className = '';
                // Remove pushpin and reset cell content
                const wrapper = cell.querySelector('.cell-content-wrapper');
                if (wrapper) {
                    const pushpin = wrapper.querySelector('.pushpin-container');
                    if (pushpin) pushpin.remove();
                }
            }
            // Update panel
            ui.renderImagePanel();
        }
    }

    /**
     * Format file size
     */
    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
    }

    /**
     * Start column resize
     */
    startColumnResize(e, column) {
        e.preventDefault();
        e.stopPropagation();

        this.isResizingColumn = true;
        this.resizingColumn = column;
        this.resizeStartX = e.pageX;
        this.originalWidth = window.TicknTieApp.columnWidths[column] || window.TicknTieApp.defaultColumnWidth;

        // Add resize cursor to body
        document.body.style.cursor = 'col-resize';

        // Add global mouse handlers
        document.addEventListener('mousemove', this.handleColumnResize);
        document.addEventListener('mouseup', this.endColumnResize);
    }

    /**
     * Handle column resize
     */
    handleColumnResize = (e) => {
        if (!this.isResizingColumn) return;

        const deltaX = e.pageX - this.resizeStartX;
        const newWidth = Math.max(30, this.originalWidth + deltaX); // Minimum width of 30px

        // Update column width in state
        window.TicknTieApp.columnWidths[this.resizingColumn] = newWidth;

        // Update all cells in this column
        const colHeaders = document.querySelectorAll(`.col-header:nth-child(${this.resizingColumn + 1})`);
        colHeaders.forEach(header => {
            header.style.width = newWidth + 'px';
            header.style.minWidth = newWidth + 'px';
            header.style.maxWidth = newWidth + 'px';
        });

        const cells = document.querySelectorAll(`td:nth-child(${this.resizingColumn + 1})`);
        cells.forEach(cell => {
            cell.style.width = newWidth + 'px';
            cell.style.minWidth = newWidth + 'px';
            cell.style.maxWidth = newWidth + 'px';
        });
    }

    /**
     * End column resize
     */
    endColumnResize = () => {
        if (!this.isResizingColumn) return;

        this.isResizingColumn = false;
        this.resizingColumn = null;
        document.body.style.cursor = '';

        // Remove global mouse handlers
        document.removeEventListener('mousemove', this.handleColumnResize);
        document.removeEventListener('mouseup', this.endColumnResize);
    }

    /**
     * Start row resize
     */
    startRowResize(e, row) {
        e.preventDefault();
        e.stopPropagation();

        this.isResizingRow = true;
        this.resizingRow = row;
        this.resizeStartY = e.pageY;
        this.originalHeight = window.TicknTieApp.rowHeights[row] || window.TicknTieApp.defaultRowHeight;

        // Add resize cursor to body
        document.body.style.cursor = 'row-resize';

        // Add global mouse handlers
        document.addEventListener('mousemove', this.handleRowResize);
        document.addEventListener('mouseup', this.endRowResize);
    }

    /**
     * Handle row resize
     */
    handleRowResize = (e) => {
        if (!this.isResizingRow) return;

        const deltaY = e.pageY - this.resizeStartY;
        const newHeight = Math.max(20, this.originalHeight + deltaY); // Minimum height of 20px

        // Update row height in state
        window.TicknTieApp.rowHeights[this.resizingRow] = newHeight;

        // Update the row
        const rows = document.querySelectorAll(`tr:nth-child(${this.resizingRow + 1})`);
        rows.forEach(row => {
            row.style.height = newHeight + 'px';
        });
    }

    /**
     * End row resize
     */
    endRowResize = () => {
        if (!this.isResizingRow) return;

        this.isResizingRow = false;
        this.resizingRow = null;
        document.body.style.cursor = '';

        // Remove global mouse handlers
        document.removeEventListener('mousemove', this.handleRowResize);
        document.removeEventListener('mouseup', this.endRowResize);
    }

    /**
     * Update status bar
     */
    updateStatusBar(message) {
        const statusText = document.getElementById('statusText');
        if (statusText) {
            statusText.textContent = message;
        }
    }

    /**
     * Show loading spinner
     */
    showLoading() {
        const loading = document.getElementById('loading');
        if (loading) {
            loading.style.display = 'block';
        }
    }

    /**
     * Hide loading spinner
     */
    hideLoading() {
        const loading = document.getElementById('loading');
        if (loading) {
            loading.style.display = 'none';
        }
    }
}

// Create singleton instance
const uiComponents = new UIComponents();

// Make it available globally
window.UIComponents = UIComponents;
window.uiComponents = uiComponents;