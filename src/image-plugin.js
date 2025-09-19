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
    this.setupResizeHandler()
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

    // Add resize handle
    const resizeHandle = document.createElement('div')
    resizeHandle.className = 'sidebar-resize-handle'
    resizeHandle.innerHTML = '<div class="resize-grip"></div>'
    sidebar.parentNode.insertBefore(resizeHandle, sidebar)

    // Add collapse button OUTSIDE the sidebar (as a sibling)
    const collapseBtn = document.createElement('button')
    collapseBtn.id = 'collapseBtn'
    collapseBtn.className = 'collapse-btn'
    collapseBtn.title = 'Collapse/Expand sidebar'
    collapseBtn.innerHTML = '<span class="collapse-icon">▶</span>'
    sidebar.parentNode.insertBefore(collapseBtn, sidebar)

    sidebar.innerHTML += `
      <div class="sidebar-header">
        <h3>📌 Image Viewer</h3>
        <input
          type="file"
          id="imageInput"
          accept="image/*,application/pdf"
          style="display: none;"
        >
        <button id="addImageBtn" class="btn-primary">
          ➕ Add File to Cell
        </button>
      </div>

      <div class="sidebar-content">
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

    // Also poll for selection changes and check for deleted content
    setInterval(() => {
      this.detectCurrentSelection()
      this.checkForDeletedCells()
    }, 500)
  }

  checkForDeletedCells() {
    try {
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) return

      const worksheet = workbook.getActiveSheet()
      if (!worksheet) return

      // Check each cell that has an image linked
      for (const [cellRef, imageId] of this.cellImages.entries()) {
        const { row, col } = this.parseCellRef(cellRef)
        const range = worksheet.getRange(row, col)

        if (range) {
          const value = range.getValue?.() || range.value

          // If cell is empty or doesn't contain the image reference, remove the link
          if (!value || value === '') {
            console.log(`Cell ${cellRef} was cleared, removing linked image`)
            this.removeImageLink(cellRef)
          }
        }
      }
    } catch (e) {
      console.debug('Error checking for deleted cells:', e)
    }
  }

  removeImageLink(cellRef) {
    if (!this.cellImages.has(cellRef)) return

    const imageId = this.cellImages.get(cellRef)
    const imageData = this.imageStore.get(imageId)

    // Clean up image data
    if (imageData && imageData.url) {
      URL.revokeObjectURL(imageData.url)
      this.imageStore.delete(imageId)
    }

    // Remove the link
    this.cellImages.delete(cellRef)

    // Update UI if this is the selected cell
    if (this.selectedCell === cellRef) {
      this.displayEmptyState()
    }

    console.log(`✅ Image link removed from cell ${cellRef}`)
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
    if (!file.type.startsWith('image/') && file.type !== 'application/pdf') {
      alert('Please select a valid image or PDF file')
      return
    }

    if (file.size > 10 * 1024 * 1024) { // 10MB limit
      alert('File size must be less than 10MB')
      return
    }

    try {
      // Create file data
      const fileId = `file_${Date.now()}`
      const fileUrl = URL.createObjectURL(file)

      let previewUrl = fileUrl

      // For PDFs, generate a preview of the first page
      if (file.type === 'application/pdf') {
        previewUrl = await this.generatePDFPreview(file)
      }

      const imageData = {
        id: fileId,
        name: file.name,
        size: file.size,
        type: file.type,
        url: fileUrl,
        previewUrl: previewUrl,
        isPDF: file.type === 'application/pdf'
      }

      // Store file
      this.imageStore.set(fileId, imageData)
      this.cellImages.set(this.selectedCell, fileId)

      // Update cell value with appropriate icon
      const icon = imageData.isPDF ? '📄' : '📌'
      await this.updateCellValue(this.selectedCell, `${icon} ${file.name}`)

      // Show image
      this.displayImage(imageData)

      console.log(`✅ File added to cell ${this.selectedCell}`)
    } catch (error) {
      console.error('Failed to upload file:', error)
      alert('Failed to upload file. Please try again.')
    }
  }

  selectCell(cellRef) {
    this.selectedCell = cellRef

    // Check for image
    if (this.cellImages.has(cellRef)) {
      const imageId = this.cellImages.get(cellRef)
      const imageData = this.imageStore.get(imageId)
      if (imageData) {
        // Auto-expand sidebar if collapsed
        this.expandSidebar()
        this.displayImage(imageData)
      }
    } else {
      this.displayEmptyState()
    }
  }

  displayImage(imageData) {
    const container = document.getElementById('imageContainer')

    // Use preview URL for display (for PDFs, this will be the canvas image)
    const displayUrl = imageData.previewUrl || imageData.url

    container.innerHTML = `
      <div class="image-preview">
        <img src="${displayUrl}" alt="${imageData.name}">
        ${imageData.isPDF ? '<div class="pdf-notice">📄 PDF Preview (Page 1)</div>' : ''}

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

        <button class="btn-primary" onclick="window.open('${imageData.url}', '_blank')">
          📂 Open File
        </button>
      </div>
    `
  }

  displayEmptyState() {
    const container = document.getElementById('imageContainer')

    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📁</div>
        <p>No file in this cell</p>
        <small>Click "Add File to Cell" to attach one</small>
      </div>
    `
  }

  async removeImage() {
    if (!this.selectedCell || !this.cellImages.has(this.selectedCell)) return

    try {
      // Clear cell value first
      await this.updateCellValue(this.selectedCell, '')

      // Remove the image link
      this.removeImageLink(this.selectedCell)
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
    // Convert column number to Excel-style column letters (A, B, ..., Z, AA, AB, ...)
    let colLetter = ''
    let colNum = col
    while (colNum >= 0) {
      colLetter = String.fromCharCode(65 + (colNum % 26)) + colLetter
      colNum = Math.floor(colNum / 26) - 1
    }
    return `${colLetter}${row + 1}`
  }

  parseCellRef(cellRef) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/)
    if (!match) return { row: 0, col: 0 }

    // Convert Excel-style column letters to column number
    const colLetters = match[1]
    let col = 0
    for (let i = 0; i < colLetters.length; i++) {
      col = col * 26 + (colLetters.charCodeAt(i) - 64)
    }
    col = col - 1 // Convert to 0-based index

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

  async generatePDFPreview(file) {
    try {
      // Get PDF.js from global scope
      const pdfjsLib = window.pdfjsLib
      if (!pdfjsLib) {
        console.error('PDF.js library not loaded')
        return URL.createObjectURL(file)
      }

      // Configure worker
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'

      // Load the PDF
      const fileUrl = URL.createObjectURL(file)
      const loadingTask = pdfjsLib.getDocument(fileUrl)
      const pdf = await loadingTask.promise

      // Get the first page
      const page = await pdf.getPage(1)

      // Set scale for preview
      const scale = 1.5
      const viewport = page.getViewport({ scale })

      // Create canvas for rendering
      const canvas = document.createElement('canvas')
      const context = canvas.getContext('2d')
      canvas.height = viewport.height
      canvas.width = viewport.width

      // Render the page
      const renderContext = {
        canvasContext: context,
        viewport: viewport
      }
      await page.render(renderContext).promise

      // Convert canvas to blob and create URL
      return new Promise((resolve) => {
        canvas.toBlob((blob) => {
          const previewUrl = URL.createObjectURL(blob)
          resolve(previewUrl)
        }, 'image/png')
      })
    } catch (error) {
      console.error('Failed to generate PDF preview:', error)
      // Fallback to using the original file URL
      return URL.createObjectURL(file)
    }
  }

  setupResizeHandler() {
    const sidebar = document.getElementById('sidebar')
    const resizeHandle = document.querySelector('.sidebar-resize-handle')
    let isResizing = false
    let startX = 0
    let startWidth = 0

    // Get saved width from localStorage or use default
    const savedWidth = localStorage.getItem('sidebarWidth')
    const collapseBtn = document.getElementById('collapseBtn')

    if (savedWidth && savedWidth !== '0') {
      sidebar.style.width = savedWidth + 'px'
      if (collapseBtn) {
        // Position on outer (left) edge, flush with sidebar
        collapseBtn.style.right = `${savedWidth}px`
      }
    } else if (collapseBtn) {
      // Position on outer (left) edge with default width
      collapseBtn.style.right = '350px'
    }

    // Check if sidebar was collapsed
    const wasCollapsed = localStorage.getItem('sidebarCollapsed') === 'true'
    if (wasCollapsed) {
      sidebar.classList.add('collapsed')
      sidebar.style.width = '40px'
      if (collapseBtn) {
        collapseBtn.style.display = 'none'  // Hide button when collapsed
        const collapseIcon = collapseBtn.querySelector('.collapse-icon')
        if (collapseIcon) collapseIcon.textContent = '◀'
      }
    }

    // Setup collapse button
    this.setupCollapseButton()

    resizeHandle.addEventListener('mousedown', (e) => {
      isResizing = true
      startX = e.clientX
      startWidth = sidebar.offsetWidth

      // Add classes for styling during resize
      document.body.classList.add('resizing')
      resizeHandle.classList.add('resizing')

      e.preventDefault()
    })

    document.addEventListener('mousemove', (e) => {
      if (!isResizing) return

      // Don't allow resizing if sidebar is collapsed
      if (sidebar.classList.contains('collapsed')) {
        isResizing = false
        return
      }

      const deltaX = startX - e.clientX
      const newWidth = startWidth + deltaX

      // Set min and max width constraints
      const minWidth = 250
      const maxWidth = 800

      if (newWidth >= minWidth && newWidth <= maxWidth) {
        sidebar.style.width = newWidth + 'px'
        sidebar.classList.remove('collapsed')
        // Update button position to track with sidebar (outer edge)
        const collapseBtn = document.getElementById('collapseBtn')
        if (collapseBtn) {
          collapseBtn.style.right = `${newWidth}px`
        }
      }
    })

    document.addEventListener('mouseup', () => {
      if (isResizing) {
        isResizing = false

        // Remove styling classes
        document.body.classList.remove('resizing')
        resizeHandle.classList.remove('resizing')

        // Save width to localStorage
        const finalWidth = sidebar.offsetWidth
        localStorage.setItem('sidebarWidth', finalWidth)
        // Update button position (outer edge)
        const collapseBtn = document.getElementById('collapseBtn')
        if (collapseBtn) {
          collapseBtn.style.right = `${finalWidth}px`
        }
      }
    })

    // Handle double-click to collapse/expand
    resizeHandle.addEventListener('dblclick', () => {
      this.toggleSidebar()
    })
  }

  setupCollapseButton() {
    const sidebar = document.getElementById('sidebar')
    const collapseBtn = document.getElementById('collapseBtn')

    if (!collapseBtn) return

    collapseBtn.addEventListener('click', () => {
      this.toggleSidebar()
    })

    // Make entire sidebar clickable when collapsed
    sidebar.addEventListener('click', (e) => {
      if (sidebar.classList.contains('collapsed')) {
        // Prevent triggering if clicking on something inside when partially visible
        e.stopPropagation()
        this.toggleSidebar()
      }
    })
  }

  toggleSidebar() {
    const sidebar = document.getElementById('sidebar')
    const collapseBtn = document.getElementById('collapseBtn')
    const collapseIcon = collapseBtn?.querySelector('.collapse-icon')
    const collapsedWidth = 40  // Keep a narrow sidebar visible
    const defaultWidth = 350

    if (sidebar.classList.contains('collapsed')) {
      // Expand
      sidebar.style.width = defaultWidth + 'px'
      sidebar.classList.remove('collapsed')
      collapseBtn.style.right = `${defaultWidth}px`  // Outer edge
      collapseBtn.style.display = 'flex'  // Show button
      if (collapseIcon) collapseIcon.textContent = '▶'
      localStorage.setItem('sidebarWidth', defaultWidth)
      localStorage.removeItem('sidebarCollapsed')
    } else {
      // Collapse
      sidebar.style.width = collapsedWidth + 'px'
      sidebar.classList.add('collapsed')
      collapseBtn.style.display = 'none'  // Hide button
      if (collapseIcon) collapseIcon.textContent = '◀'
      localStorage.setItem('sidebarWidth', collapsedWidth)
      localStorage.setItem('sidebarCollapsed', 'true')
    }
  }

  expandSidebar() {
    const sidebar = document.getElementById('sidebar')
    const collapseBtn = document.getElementById('collapseBtn')
    const collapseIcon = collapseBtn?.querySelector('.collapse-icon')

    if (sidebar.classList.contains('collapsed')) {
      const defaultWidth = 350
      sidebar.style.width = defaultWidth + 'px'
      sidebar.classList.remove('collapsed')
      collapseBtn.style.right = `${defaultWidth}px`  // Outer edge
      collapseBtn.style.display = 'flex'  // Show button
      if (collapseIcon) collapseIcon.textContent = '▶'
      localStorage.setItem('sidebarWidth', defaultWidth)
      localStorage.removeItem('sidebarCollapsed')
    }
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