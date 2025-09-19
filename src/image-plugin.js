/**
 * Image Sidebar Plugin for Univer
 * Provides image attachment and viewing functionality
 */

export class ImagePlugin {
  constructor(univerAPI) {
    this.univerAPI = univerAPI
    this.imageStore = new Map()  // imageId -> imageData
    this.cellImages = new Map()  // cellRef -> imageId
    this.selectedCell = null
  }

  async init() {
    this.createSidebar()
    this.listenToSelection()
    this.setupImageUpload()
    console.log('📌 Image plugin initialized')

    // Debug: Log available methods
    console.log('Available UniversAPI methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(this.univerAPI)))
    const workbook = this.univerAPI.getActiveWorkbook()
    if (workbook) {
      console.log('Workbook methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(workbook)))
      const sheet = workbook.getActiveSheet()
      if (sheet) {
        console.log('Sheet methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(sheet)))
      }
    }
  }

  createSidebar() {
    const sidebar = document.getElementById('sidebar')

    sidebar.innerHTML = `
      <div class="sidebar-header">
        <h3>📌 Image Viewer</h3>
        <input
          type="file"
          id="imageInput"
          accept="image/*"
          style="display: none;"
        >
        <button id="addImageBtn" class="btn-primary">
          ➕ Add Image to Cell
        </button>
      </div>

      <div class="sidebar-content">
        <div id="cellInfo" class="cell-info">
          <small>Selected Cell:</small>
          <strong id="cellRef">None</strong>
        </div>

        <div id="imageContainer" class="image-container">
          <div class="empty-state">
            <div class="empty-icon">🖼️</div>
            <p>No image in this cell</p>
            <small>Select a cell and add an image</small>
          </div>
        </div>
      </div>
    `
  }

  listenToSelection() {
    // Default to A1 if no selection
    this.selectCell('A1')

    // Listen to click events on the spreadsheet to detect selection
    const spreadsheetContainer = document.getElementById('spreadsheet')
    if (spreadsheetContainer) {
      spreadsheetContainer.addEventListener('click', (e) => {
        setTimeout(() => {
          this.detectCurrentSelection()
        }, 100)
      })
    }

    // Also poll for selection changes
    setInterval(() => {
      this.detectCurrentSelection()
    }, 500)
  }

  detectCurrentSelection() {
    try {
      // First try to get from Univer's internal state
      const univer = this.univerAPI._univer || this.univerAPI.getUniver?.()
      if (univer) {
        const selectionManagerService = univer.getInjector?.()?.get?.('SelectionManagerService')
        if (selectionManagerService) {
          const selections = selectionManagerService.getSelections?.()
          if (selections && selections.length > 0) {
            const range = selections[0].range
            if (range) {
              const row = range.startRow
              const col = range.startColumn
              const cellRef = this.getCellRef(row, col)
              if (cellRef && cellRef !== this.selectedCell) {
                console.log('Selection changed to:', cellRef)
                this.selectCell(cellRef)
              }
              return
            }
          }
        }
      }

      // Try the workbook/worksheet API
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) return

      const worksheet = workbook.getActiveSheet()
      if (!worksheet) return

      // Check if there's a getActiveCell method
      if (worksheet.getActiveCell) {
        const activeCell = worksheet.getActiveCell()
        if (activeCell) {
          const row = activeCell.row ?? activeCell.getRow?.() ?? 0
          const col = activeCell.column ?? activeCell.getColumn?.() ?? 0
          const cellRef = this.getCellRef(row, col)
          if (cellRef && cellRef !== this.selectedCell) {
            console.log('Selection changed to:', cellRef)
            this.selectCell(cellRef)
          }
          return
        }
      }

      // Try to get from selections
      const selections = worksheet.getSelections?.() || worksheet.getSelection?.()
      if (selections) {
        const range = selections.getRange?.() || selections.getCurrentRange?.() || selections[0]
        if (range) {
          const row = range.startRow ?? range.row ?? 0
          const col = range.startColumn ?? range.column ?? 0
          const cellRef = this.getCellRef(row, col)
          if (cellRef && cellRef !== this.selectedCell) {
            console.log('Selection changed to:', cellRef)
            this.selectCell(cellRef)
          }
        }
      }
    } catch (e) {
      console.debug('Selection detection error:', e)
    }
  }

  setupImageUpload() {
    // Add image button click
    document.getElementById('addImageBtn').addEventListener('click', () => {
      if (!this.selectedCell) {
        alert('Please select a cell first')
        return
      }
      document.getElementById('imageInput').click()
    })

    // File input change
    document.getElementById('imageInput').addEventListener('change', (e) => {
      const file = e.target.files[0]
      if (file) {
        this.handleImageUpload(file)
        e.target.value = '' // Clear input
      }
    })
  }

  async handleImageUpload(file) {
    if (!this.selectedCell) return

    // Validate file
    if (!file.type.startsWith('image/')) {
      alert('Please select a valid image file')
      return
    }

    if (file.size > 10 * 1024 * 1024) { // 10MB limit
      alert('Image size must be less than 10MB')
      return
    }

    try {
      // Create image data
      const imageId = `img_${Date.now()}`
      const imageUrl = URL.createObjectURL(file)

      const imageData = {
        id: imageId,
        name: file.name,
        size: file.size,
        type: file.type,
        url: imageUrl
      }

      // Store image
      this.imageStore.set(imageId, imageData)
      this.cellImages.set(this.selectedCell, imageId)

      // Update cell value
      await this.updateCellValue(this.selectedCell, `📌 ${file.name}`)

      // Show image
      this.displayImage(imageData)

      console.log(`✅ Image added to cell ${this.selectedCell}`)
    } catch (error) {
      console.error('Failed to upload image:', error)
      alert('Failed to upload image. Please try again.')
    }
  }

  selectCell(cellRef) {
    this.selectedCell = cellRef

    // Update UI
    document.getElementById('cellRef').textContent = cellRef || 'None'

    // Check for image
    if (this.cellImages.has(cellRef)) {
      const imageId = this.cellImages.get(cellRef)
      const imageData = this.imageStore.get(imageId)
      if (imageData) {
        this.displayImage(imageData)
      }
    } else {
      this.displayEmptyState()
    }
  }

  displayImage(imageData) {
    const container = document.getElementById('imageContainer')

    container.innerHTML = `
      <div class="image-preview">
        <img src="${imageData.url}" alt="${imageData.name}">

        <div class="image-details">
          <div class="detail-row">
            <span>Name:</span>
            <strong>${imageData.name}</strong>
          </div>
          <div class="detail-row">
            <span>Size:</span>
            <strong>${this.formatFileSize(imageData.size)}</strong>
          </div>
          <div class="detail-row">
            <span>Type:</span>
            <strong>${imageData.type}</strong>
          </div>
        </div>

        <button class="btn-danger" onclick="window.ticknTie.imagePlugin.removeImage()">
          🗑️ Remove Image
        </button>
      </div>
    `
  }

  displayEmptyState() {
    const container = document.getElementById('imageContainer')

    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">🖼️</div>
        <p>No image in this cell</p>
        <small>Click "Add Image to Cell" to attach one</small>
      </div>
    `
  }

  async removeImage() {
    if (!this.selectedCell || !this.cellImages.has(this.selectedCell)) return

    try {
      const imageId = this.cellImages.get(this.selectedCell)
      const imageData = this.imageStore.get(imageId)

      // Clean up
      if (imageData) {
        URL.revokeObjectURL(imageData.url)
        this.imageStore.delete(imageId)
      }
      this.cellImages.delete(this.selectedCell)

      // Clear cell value
      await this.updateCellValue(this.selectedCell, '')

      // Update UI
      this.displayEmptyState()

      console.log(`✅ Image removed from cell ${this.selectedCell}`)
    } catch (error) {
      console.error('Failed to remove image:', error)
    }
  }

  async updateCellValue(cellRef, value) {
    try {
      const { row, col } = this.parseCellRef(cellRef)

      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) return

      const sheet = workbook.getActiveSheet()
      if (!sheet) return

      const range = sheet.getRange(row, col)
      if (range && range.setValue) {
        range.setValue(value)
      }
    } catch (e) {
      console.warn('Failed to update cell value:', e)
    }
  }

  getCellRef(row, col) {
    if (row < 0 || col < 0) return null
    const colLetter = String.fromCharCode(65 + col)
    return `${colLetter}${row + 1}`
  }

  parseCellRef(cellRef) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/)
    if (!match) return { row: 0, col: 0 }

    const col = match[1].charCodeAt(0) - 65
    const row = parseInt(match[2]) - 1
    return { row, col }
  }

  formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
  }

  // Public API for debugging/extensions
  getImageCount() {
    return this.imageStore.size
  }

  getCellsWithImages() {
    return Array.from(this.cellImages.keys())
  }

  clearAll() {
    // Clean up all object URLs
    this.imageStore.forEach(imageData => {
      URL.revokeObjectURL(imageData.url)
    })

    // Clear maps
    this.imageStore.clear()
    this.cellImages.clear()

    // Update UI
    this.displayEmptyState()

    console.log('✅ All images cleared')
  }
}