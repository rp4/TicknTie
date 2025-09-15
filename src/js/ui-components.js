/**
 * UI Components Module
 * Handles all UI rendering and interactions
 */

class UIComponents {
    constructor() {
        this.currentCell = null;
        this.selectedCells = [];
        this.imagePanelTab = 'preview'; // 'preview' or 'gallery'
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
                'col-header'
            );
            headerRow.appendChild(colHeader);
        }
        table.appendChild(headerRow);
        
        // Create data rows
        for (let row = 1; row <= maxRow; row++) {
            const tr = document.createElement('tr');
            
            // Row header
            tr.appendChild(this.createHeaderCell(row.toString(), 'row-header'));
            
            // Data cells
            for (let col = 1; col <= maxCol; col++) {
                const cellRef = window.ExcelHandler.createCellRef(row, col);
                const cellData = sheet.data[cellRef] || {};
                const td = this.createDataCell(cellRef, cellData);
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
    createHeaderCell(content, className) {
        const th = document.createElement('th');
        th.className = className;
        th.textContent = content;
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
        
        if (imageData) {
            td.className = 'cell-with-image-ref';
            
            // Add thumbnail
            const thumbnail = document.createElement('img');
            thumbnail.className = 'cell-image-thumbnail';
            thumbnail.src = imageData.url;
            thumbnail.onclick = (e) => {
                e.stopPropagation();
                this.showImageInPanel(imageData);
            };
            td.appendChild(thumbnail);
        }
        
        // Add input element
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell-input';
        input.value = cellData.value || '';
        input.dataset.cell = cellRef;
        
        // Apply formatting
        if (cellData.formatting) {
            this.applyCellFormatting(input, cellData.formatting);
        }
        
        // Add event listeners
        input.addEventListener('focus', () => this.handleCellFocus(cellRef));
        input.addEventListener('blur', () => this.handleCellBlur(cellRef, input.value));
        input.addEventListener('input', () => this.handleCellInput(cellRef, input.value));
        
        td.appendChild(input);
        
        return td;
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
        
        panel.innerHTML = `
            <button class="image-panel-toggle" onclick="UIComponents.toggleImagePanel()">
                <span id="imagePanelToggleIcon">◀</span>
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
        this.imagePanelTab = 'preview';
        const content = document.getElementById('imagePanelContent');
        if (content) {
            content.innerHTML = this.renderImagePreview(imageData);
        }
        
        // Ensure panel is visible
        const panel = document.getElementById('imagePanel');
        if (panel && panel.classList.contains('collapsed')) {
            this.toggleImagePanel();
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
                const thumbnail = cell.querySelector('.cell-image-thumbnail');
                if (thumbnail) thumbnail.remove();
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