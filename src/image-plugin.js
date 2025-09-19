/**
 * Hyperlink Image Sidebar Plugin for Univer
 * Provides hyperlink attachment and preview functionality with smart caching
 */

export class ImagePlugin {
  constructor(univerAPI) {
    this.univerAPI = univerAPI
    this.cellHyperlinks = new Map()  // cellRef -> { url, displayText }
    this.previewCache = new Map()    // url -> { preview, timestamp, size }
    this.selectedCell = null
    this.preloadQueue = []
    this.isPreloading = false
    this.maxCacheSize = 100 // Maximum number of cached previews
    this.maxCacheAge = 30 * 60 * 1000 // 30 minutes
  }

  async init() {
    this.createSidebar()
    this.setupResizeHandler()
    this.listenToSelection()
    this.setupHyperlinkInput()
    this.startCacheCleanup()
    console.log('📌 Hyperlink image plugin initialized')

    // Debug: Log available Univer APIs
    console.log('=== Debugging Univer API ===')
    console.log('univerAPI:', this.univerAPI)

    const workbook = this.univerAPI.getActiveWorkbook()
    if (workbook) {
      console.log('Workbook available, methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(workbook)))

      const sheet = workbook.getActiveSheet()
      if (sheet) {
        console.log('Sheet available, methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(sheet)))

        // Try to get selection
        const selection = sheet.getSelection?.()
        if (selection) {
          console.log('Selection available, methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(selection)))
        }
      }
    }

    // Check for internal Univer instance
    const univer = this.univerAPI._univer || this.univerAPI.getUniver?.() || window.univer
    if (univer) {
      console.log('Internal Univer found:', univer)
      const injector = univer.getInjector?.()
      if (injector) {
        console.log('Injector found, trying to get services...')
        // Try to list available services
        if (injector.get) {
          try {
            // Common Univer service names
            const services = [
              'SelectionManagerService',
              'ICommandService',
              'IUniverInstanceService',
              'IRenderManagerService'
            ]
            services.forEach(name => {
              try {
                const service = injector.get(name)
                if (service) console.log(`Found service: ${name}`, service)
              } catch (e) {
                // Service not found
              }
            })
          } catch (e) {
            console.log('Could not enumerate services:', e)
          }
        }
      }
    }
    console.log('=== End Debug ===')
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

        <!-- File Upload (hidden input) -->
        <input
          type="file"
          id="fileInput"
          accept="image/*,application/pdf"
          style="display: none;"
        >

        <button id="addFileBtn" class="btn-primary">
          ➕ Add Evidence
        </button>
      </div>

      <div class="sidebar-content">
        <div id="imageContainer" class="image-container">
          <div class="empty-state">
            <div class="empty-icon">📁</div>
            <p>No evidence in this cell</p>
            <small>Select a cell and add evidence</small>
          </div>
        </div>
      </div>
    `

    // Add styles for new elements
    const style = document.createElement('style')
    style.textContent = `
      .loading-spinner {
        display: inline-block;
        width: 20px;
        height: 20px;
        border: 3px solid rgba(0,0,0,.1);
        border-radius: 50%;
        border-top-color: #007bff;
        animation: spin 1s linear infinite;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
      .preview-loading {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 300px;
        color: #666;
      }
      .preview-error {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        height: 200px;
        color: #d32f2f;
        text-align: center;
        padding: 20px;
      }
      .preview-error .error-icon {
        font-size: 48px;
        margin-bottom: 10px;
      }
      .retry-btn {
        margin-top: 10px;
        padding: 8px 16px;
        background: #007bff;
        color: white;
        border: none;
        border-radius: 4px;
        cursor: pointer;
      }
    `
    document.head.appendChild(style)
  }

  listenToSelection() {
    // Default to A1 if no selection
    this.selectCell('A1')

    // Listen to multiple events on the spreadsheet to detect selection
    const spreadsheetContainer = document.getElementById('spreadsheet')
    if (spreadsheetContainer) {
      // Click events
      spreadsheetContainer.addEventListener('click', (e) => {
        setTimeout(() => {
          this.detectCurrentSelection()
        }, 100)
      })

      // Keyboard navigation and deletion
      spreadsheetContainer.addEventListener('keyup', (e) => {
        // Arrow keys, Tab, Enter
        if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab', 'Enter'].includes(e.key)) {
          setTimeout(() => {
            this.detectCurrentSelection()
          }, 100)
        }

        // Delete or Backspace key - check for cell deletion
        if (e.key === 'Delete' || e.key === 'Backspace') {
          setTimeout(() => {
            this.checkForDeletedContent()
          }, 100)
        }
      })

      // Focus events
      spreadsheetContainer.addEventListener('focusin', (e) => {
        setTimeout(() => {
          this.detectCurrentSelection()
        }, 100)
      })

      // Input events to detect content changes
      spreadsheetContainer.addEventListener('input', (e) => {
        setTimeout(() => {
          this.checkForDeletedContent()
        }, 100)
      })
    }

    // Poll for selection changes and content deletion
    setInterval(() => {
      this.detectCurrentSelection()
      this.checkForDeletedContent()
    }, 500)

    // Initial detection after a short delay
    setTimeout(() => {
      this.detectCurrentSelection()
    }, 1000)
  }

  // Check if any cells with hyperlinks have been cleared
  checkForDeletedContent() {
    try {
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) return

      const worksheet = workbook.getActiveSheet()
      if (!worksheet) return

      // Check each cell that has a hyperlink
      const cellsToRemove = []

      for (const [cellRef, linkData] of this.cellHyperlinks.entries()) {
        const { row, col } = this.parseCellRef(cellRef)

        try {
          const range = worksheet.getRange(row, col)
          if (range) {
            const value = range.getValue?.() || range.value

            // If cell is empty or doesn't contain the pushpin icon, remove the hyperlink
            if (!value || value === '' || !value.includes('📌')) {
              console.log(`Cell ${cellRef} was cleared, removing hyperlink`)
              cellsToRemove.push(cellRef)
            }
          }
        } catch (e) {
          console.debug('Error checking cell:', cellRef, e)
        }
      }

      // Remove hyperlinks for cleared cells
      cellsToRemove.forEach(cellRef => {
        this.removeHyperlinkFromCell(cellRef)
      })

    } catch (e) {
      console.debug('Error checking for deleted content:', e)
    }
  }

  // Remove hyperlink from a specific cell
  removeHyperlinkFromCell(cellRef) {
    if (!this.cellHyperlinks.has(cellRef)) return

    // Remove from storage
    this.cellHyperlinks.delete(cellRef)

    // Update UI if this is the currently selected cell
    if (this.selectedCell === cellRef) {
      this.displayEmptyState()
    }

    console.log(`✅ Hyperlink removed from cell ${cellRef}`)
  }

  detectCurrentSelection() {
    try {
      // Method 1: Try to get selection from Univer API
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) {
        console.debug('No active workbook')
        return
      }

      const worksheet = workbook.getActiveSheet()
      if (!worksheet) {
        console.debug('No active worksheet')
        return
      }

      // Method 2: Try getting selection from worksheet
      const selection = worksheet.getSelection?.()
      if (selection) {
        const activeRange = selection.getActiveRange?.() || selection.getRange?.()
        if (activeRange) {
          const row = activeRange.getRow?.() ?? activeRange.startRow ?? activeRange.row ?? 0
          const col = activeRange.getColumn?.() ?? activeRange.startColumn ?? activeRange.column ?? 0
          const cellRef = this.getCellRef(row, col)

          if (cellRef && cellRef !== this.selectedCell) {
            console.log('Selection changed to:', cellRef)
            this.selectCell(cellRef)
          }
          return
        }
      }

      // Method 3: Try internal Univer services
      const univer = this.univerAPI._univer || this.univerAPI.getUniver?.() || window.univer
      if (univer) {
        // Try to get injector
        const injector = univer.getInjector?.() || univer.__injector || univer._injector
        if (injector) {
          // Try various selection service names
          const serviceNames = [
            'SelectionManagerService',
            'selection.manager.service',
            'ISelectionManagerService',
            'SelectionService'
          ]

          for (const serviceName of serviceNames) {
            try {
              const selectionService = injector.get?.(serviceName)
              if (selectionService) {
                const selections = selectionService.getSelections?.() ||
                                 selectionService.getCurrentSelections?.() ||
                                 selectionService.getActiveSelection?.()

                if (selections) {
                  const sel = Array.isArray(selections) ? selections[0] : selections
                  if (sel?.range) {
                    const row = sel.range.startRow ?? sel.range.row ?? 0
                    const col = sel.range.startColumn ?? sel.range.column ?? 0
                    const cellRef = this.getCellRef(row, col)

                    if (cellRef && cellRef !== this.selectedCell) {
                      console.log('Selection changed to:', cellRef, '(from service)')
                      this.selectCell(cellRef)
                    }
                    return
                  }
                }
              }
            } catch (e) {
              // Continue to next service name
            }
          }
        }
      }

      // Method 4: Try to intercept DOM events or observe selection changes
      // This is a fallback - look for selected cell visual indicators
      const selectedCellElement = document.querySelector('.univer-cell-selected, .cell-selected, [data-selected="true"]')
      if (selectedCellElement) {
        const row = parseInt(selectedCellElement.getAttribute('data-row') || '0')
        const col = parseInt(selectedCellElement.getAttribute('data-col') || '0')
        if (!isNaN(row) && !isNaN(col)) {
          const cellRef = this.getCellRef(row, col)
          if (cellRef && cellRef !== this.selectedCell) {
            console.log('Selection changed to:', cellRef, '(from DOM)')
            this.selectCell(cellRef)
          }
        }
      }

    } catch (e) {
      console.debug('Selection detection error:', e)
    }
  }

  setupHyperlinkInput() {
    const addFileBtn = document.getElementById('addFileBtn')
    const fileInput = document.getElementById('fileInput')

    // Add file button - directly opens file picker
    addFileBtn.addEventListener('click', () => {
      if (!this.selectedCell) {
        alert('Please select a cell first')
        return
      }
      fileInput.click()
    })

    // File input change - immediately process the file
    fileInput.addEventListener('change', async (e) => {
      const file = e.target.files[0]
      if (file) {
        await this.handleFileUpload(file)
        // Reset input for next use
        fileInput.value = ''
      }
    })
  }

  async handleFileUpload(file) {
    if (!this.selectedCell) return

    try {
      // Convert file to data URL
      const url = await this.fileToDataUrl(file)
      const displayText = file.name

      // Add to cell
      await this.addHyperlinkToCell(url, displayText)

      console.log(`✅ File added to cell ${this.selectedCell}`)
    } catch (error) {
      console.error('Failed to upload file:', error)
      alert('Failed to upload file. Please try again.')
    }
  }

  // Convert file to data URL
  async fileToDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = (e) => resolve(e.target.result)
      reader.onerror = reject
      reader.readAsDataURL(file)
    })
  }

  async addHyperlinkToCell(url, displayText) {
    if (!this.selectedCell) return

    try {
      // Store hyperlink data
      this.cellHyperlinks.set(this.selectedCell, { url, displayText })

      // Update cell with hyperlink (for now, just text with icon)
      // TODO: Use proper Univer hyperlink API when available
      await this.updateCellValue(this.selectedCell, `📌 ${displayText}`)

      // Preload this image and adjacent cells
      await this.displayPreview(url, displayText)
      this.preloadAdjacentCells(this.selectedCell)

      console.log(`✅ Hyperlink added to cell ${this.selectedCell}`)
    } catch (error) {
      console.error('Failed to add hyperlink:', error)
      alert('Failed to add hyperlink. Please try again.')
    }
  }

  async selectCell(cellRef) {
    this.selectedCell = cellRef

    // Check for hyperlink
    if (this.cellHyperlinks.has(cellRef)) {
      const link = this.cellHyperlinks.get(cellRef)
      // Auto-expand sidebar if collapsed
      this.expandSidebar()
      await this.displayPreview(link.url, link.displayText)
      // Preload adjacent cells
      this.preloadAdjacentCells(cellRef)
    } else {
      this.displayEmptyState()
    }
  }

  async displayPreview(url, displayText) {
    const container = document.getElementById('imageContainer')

    // Check cache first
    if (this.previewCache.has(url)) {
      const cached = this.previewCache.get(url)
      // Update access time for LRU
      cached.timestamp = Date.now()
      this.showPreviewContent(cached.preview, url, displayText)
      return
    }

    // Show loading state
    container.innerHTML = `
      <div class="preview-loading">
        <div class="loading-spinner"></div>
        <p>Loading preview...</p>
        <small>${displayText}</small>
      </div>
    `

    try {
      const preview = await this.fetchPreview(url)
      if (preview) {
        // Cache the preview
        this.addToCache(url, preview)
        this.showPreviewContent(preview, url, displayText)
      }
    } catch (error) {
      console.error('Failed to load preview:', error)
      this.showErrorState(url, displayText, error.message)
    }
  }

  async fetchPreview(url) {
    // Check if it's a data URL
    if (url.startsWith('data:')) {
      return await this.fetchDataUrlPreview(url)
    }

    // Determine file type from URL
    const fileType = this.getFileType(url)

    if (fileType === 'pdf') {
      return await this.fetchPDFPreview(url)
    } else if (fileType === 'image') {
      return await this.fetchImagePreview(url)
    } else {
      throw new Error('Unsupported file type')
    }
  }

  async fetchDataUrlPreview(dataUrl) {
    // Determine type from data URL
    const isPdf = dataUrl.startsWith('data:application/pdf')

    if (isPdf) {
      // For PDF data URLs, we need to convert to blob first
      const response = await fetch(dataUrl)
      const blob = await response.blob()
      const objectUrl = URL.createObjectURL(blob)

      try {
        const preview = await this.fetchPDFPreview(objectUrl)
        URL.revokeObjectURL(objectUrl)
        return preview
      } catch (error) {
        URL.revokeObjectURL(objectUrl)
        throw error
      }
    } else {
      // For image data URLs, use directly
      return new Promise((resolve) => {
        const img = new Image()

        img.onload = () => {
          // Create canvas for preview (max 800px wide/tall)
          const canvas = document.createElement('canvas')
          const ctx = canvas.getContext('2d')
          const maxSize = 800

          let width = img.width
          let height = img.height

          if (width > maxSize || height > maxSize) {
            const scale = Math.min(maxSize / width, maxSize / height)
            width *= scale
            height *= scale
          }

          canvas.width = width
          canvas.height = height
          ctx.drawImage(img, 0, 0, width, height)

          resolve({
            dataUrl: canvas.toDataURL('image/jpeg', 0.9),
            width,
            height,
            type: 'image'
          })
        }

        img.src = dataUrl
      })
    }
  }

  async fetchImagePreview(url) {
    return new Promise((resolve, reject) => {
      const img = new Image()
      img.crossOrigin = 'anonymous' // Handle CORS if possible

      img.onload = () => {
        // Create canvas for preview (max 800px wide/tall)
        const canvas = document.createElement('canvas')
        const ctx = canvas.getContext('2d')
        const maxSize = 800

        let width = img.width
        let height = img.height

        if (width > maxSize || height > maxSize) {
          const scale = Math.min(maxSize / width, maxSize / height)
          width *= scale
          height *= scale
        }

        canvas.width = width
        canvas.height = height
        ctx.drawImage(img, 0, 0, width, height)

        // Convert to data URL for caching
        resolve({
          dataUrl: canvas.toDataURL('image/jpeg', 0.9),
          width,
          height,
          type: 'image'
        })
      }

      img.onerror = () => {
        reject(new Error('Failed to load image'))
      }

      img.src = url
    })
  }

  async fetchPDFPreview(url) {
    try {
      // Load PDF.js if not already loaded
      if (!window.pdfjsLib) {
        await this.loadPDFJS()
      }

      const pdfjsLib = window.pdfjsLib
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'

      // Fetch PDF
      const loadingTask = pdfjsLib.getDocument(url)
      const pdf = await loadingTask.promise

      // Get first page
      const page = await pdf.getPage(1)
      const scale = 1.5
      const viewport = page.getViewport({ scale })

      // Render to canvas
      const canvas = document.createElement('canvas')
      const context = canvas.getContext('2d')
      canvas.height = viewport.height
      canvas.width = viewport.width

      await page.render({
        canvasContext: context,
        viewport: viewport
      }).promise

      // Convert to data URL
      return {
        dataUrl: canvas.toDataURL('image/png'),
        width: viewport.width,
        height: viewport.height,
        type: 'pdf',
        pageCount: pdf.numPages
      }
    } catch (error) {
      throw new Error('Failed to load PDF preview')
    }
  }

  async loadPDFJS() {
    return new Promise((resolve, reject) => {
      if (window.pdfjsLib) {
        resolve()
        return
      }

      const script = document.createElement('script')
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js'
      script.onload = resolve
      script.onerror = reject
      document.head.appendChild(script)
    })
  }

  showPreviewContent(preview, url, displayText) {
    const container = document.getElementById('imageContainer')

    container.innerHTML = `
      <div class="image-preview">
        <img src="${preview.dataUrl}" alt="${displayText}" style="max-width: 100%; height: auto;">
        ${preview.type === 'pdf' ? `<div class="pdf-notice">📄 PDF Preview (Page 1 of ${preview.pageCount || '?'})</div>` : ''}

        <div class="image-details">
          <div class="detail-row">
            <span>Name:</span>
            <strong>${displayText}</strong>
          </div>
          <div class="detail-row">
            <span>Type:</span>
            <strong>${preview.type}</strong>
          </div>
          <div class="detail-row">
            <span>Dimensions:</span>
            <strong>${preview.width} × ${preview.height}px</strong>
          </div>
        </div>

        <button class="btn-primary" onclick="window.open('${url}', '_blank')" style="width: 100%; margin-top: 10px;">
          🔗 Open Link
        </button>
      </div>
    `
  }

  showErrorState(url, displayText, errorMessage) {
    const container = document.getElementById('imageContainer')

    container.innerHTML = `
      <div class="preview-error">
        <div class="error-icon">⚠️</div>
        <p><strong>Failed to load preview</strong></p>
        <small>${displayText}</small>
        <small style="color: #666;">${errorMessage}</small>
        <button class="retry-btn" onclick="window.ticknTie.imagePlugin.retryPreview('${url}', '${displayText}')">
          Retry
        </button>
        <button class="btn-primary" onclick="window.open('${url}', '_blank')" style="margin-top: 8px; width: 100%;">
          🔗 Open Link Anyway
        </button>
      </div>
    `
  }

  async retryPreview(url, displayText) {
    // Clear from cache to force re-fetch
    this.previewCache.delete(url)
    await this.displayPreview(url, displayText)
  }

  displayEmptyState() {
    const container = document.getElementById('imageContainer')

    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📁</div>
        <p>No evidence in this cell</p>
        <small>Click "Add Evidence" to attach a file</small>
      </div>
    `
  }

  async removeHyperlink() {
    if (!this.selectedCell || !this.cellHyperlinks.has(this.selectedCell)) return

    try {
      // Clear cell value
      await this.updateCellValue(this.selectedCell, '')

      // Use the common removal method
      this.removeHyperlinkFromCell(this.selectedCell)
    } catch (error) {
      console.error('Failed to remove hyperlink:', error)
    }
  }

  // Smart adjacent preloading
  preloadAdjacentCells(cellRef) {
    const { row, col } = this.parseCellRef(cellRef)
    const adjacentCells = []

    // Get 3x3 grid around current cell
    for (let r = row - 1; r <= row + 1; r++) {
      for (let c = col - 1; c <= col + 1; c++) {
        if (r >= 0 && c >= 0 && (r !== row || c !== col)) {
          const adjacentRef = this.getCellRef(r, c)
          if (adjacentRef && this.cellHyperlinks.has(adjacentRef)) {
            adjacentCells.push(adjacentRef)
          }
        }
      }
    }

    // Add to preload queue
    adjacentCells.forEach(ref => {
      const link = this.cellHyperlinks.get(ref)
      if (link && !this.previewCache.has(link.url)) {
        this.preloadQueue.push({ url: link.url, priority: 1 })
      }
    })

    // Start preloading if not already running
    if (!this.isPreloading) {
      this.processPreloadQueue()
    }
  }

  async processPreloadQueue() {
    if (this.preloadQueue.length === 0) {
      this.isPreloading = false
      return
    }

    this.isPreloading = true
    const item = this.preloadQueue.shift()

    try {
      if (!this.previewCache.has(item.url)) {
        const preview = await this.fetchPreview(item.url)
        if (preview) {
          this.addToCache(item.url, preview)
        }
      }
    } catch (error) {
      console.debug('Preload failed:', item.url, error)
    }

    // Continue with next item
    setTimeout(() => this.processPreloadQueue(), 100)
  }

  // Cache management
  addToCache(url, preview) {
    // Estimate size (rough approximation)
    const size = preview.dataUrl.length

    // Check cache size and evict if necessary
    if (this.previewCache.size >= this.maxCacheSize) {
      this.evictOldestCache()
    }

    this.previewCache.set(url, {
      preview,
      timestamp: Date.now(),
      size
    })
  }

  evictOldestCache() {
    let oldest = null
    let oldestTime = Date.now()

    for (const [url, data] of this.previewCache) {
      if (data.timestamp < oldestTime) {
        oldest = url
        oldestTime = data.timestamp
      }
    }

    if (oldest) {
      this.previewCache.delete(oldest)
    }
  }

  startCacheCleanup() {
    // Clean old cache entries every 5 minutes
    setInterval(() => {
      const now = Date.now()
      for (const [url, data] of this.previewCache) {
        if (now - data.timestamp > this.maxCacheAge) {
          this.previewCache.delete(url)
        }
      }
    }, 5 * 60 * 1000)
  }

  // Utility methods
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

    const colLetters = match[1]
    let col = 0
    for (let i = 0; i < colLetters.length; i++) {
      col = col * 26 + (colLetters.charCodeAt(i) - 64)
    }
    col = col - 1

    const row = parseInt(match[2]) - 1
    return { row, col }
  }

  getFileNameFromUrl(url) {
    try {
      const urlObj = new URL(url)
      const pathname = urlObj.pathname
      return pathname.substring(pathname.lastIndexOf('/') + 1) || 'Link'
    } catch {
      return url.substring(url.lastIndexOf('/') + 1) || 'Link'
    }
  }

  getFileType(url) {
    // Handle data URLs
    if (url.startsWith('data:')) {
      if (url.startsWith('data:application/pdf')) {
        return 'pdf'
      }
      return 'image'
    }

    const extension = url.toLowerCase().split('.').pop().split('?')[0]

    if (['pdf'].includes(extension)) {
      return 'pdf'
    } else if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'svg', 'bmp'].includes(extension)) {
      return 'image'
    }

    // Default to image if no clear extension
    return 'image'
  }

  // Sidebar UI methods (resize, collapse, etc.)
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
        collapseBtn.style.right = `${savedWidth}px`
      }
    } else if (collapseBtn) {
      collapseBtn.style.right = '350px'
    }

    // Check if sidebar was collapsed
    const wasCollapsed = localStorage.getItem('sidebarCollapsed') === 'true'
    if (wasCollapsed) {
      sidebar.classList.add('collapsed')
      sidebar.style.width = '40px'
      if (collapseBtn) {
        collapseBtn.style.display = 'none'
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

      document.body.classList.add('resizing')
      resizeHandle.classList.add('resizing')

      e.preventDefault()
    })

    document.addEventListener('mousemove', (e) => {
      if (!isResizing) return

      if (sidebar.classList.contains('collapsed')) {
        isResizing = false
        return
      }

      const deltaX = startX - e.clientX
      const newWidth = startWidth + deltaX

      const minWidth = 250
      const maxWidth = 800

      if (newWidth >= minWidth && newWidth <= maxWidth) {
        sidebar.style.width = newWidth + 'px'
        sidebar.classList.remove('collapsed')
        const collapseBtn = document.getElementById('collapseBtn')
        if (collapseBtn) {
          collapseBtn.style.right = `${newWidth}px`
        }
      }
    })

    document.addEventListener('mouseup', () => {
      if (isResizing) {
        isResizing = false

        document.body.classList.remove('resizing')
        resizeHandle.classList.remove('resizing')

        const finalWidth = sidebar.offsetWidth
        localStorage.setItem('sidebarWidth', finalWidth)
        const collapseBtn = document.getElementById('collapseBtn')
        if (collapseBtn) {
          collapseBtn.style.right = `${finalWidth}px`
        }
      }
    })

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

    sidebar.addEventListener('click', (e) => {
      if (sidebar.classList.contains('collapsed')) {
        e.stopPropagation()
        this.toggleSidebar()
      }
    })
  }

  toggleSidebar() {
    const sidebar = document.getElementById('sidebar')
    const collapseBtn = document.getElementById('collapseBtn')
    const collapseIcon = collapseBtn?.querySelector('.collapse-icon')
    const collapsedWidth = 40
    const defaultWidth = 350

    if (sidebar.classList.contains('collapsed')) {
      sidebar.style.width = defaultWidth + 'px'
      sidebar.classList.remove('collapsed')
      collapseBtn.style.right = `${defaultWidth}px`
      collapseBtn.style.display = 'flex'
      if (collapseIcon) collapseIcon.textContent = '▶'
      localStorage.setItem('sidebarWidth', defaultWidth)
      localStorage.removeItem('sidebarCollapsed')
    } else {
      sidebar.style.width = collapsedWidth + 'px'
      sidebar.classList.add('collapsed')
      collapseBtn.style.display = 'none'
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
      collapseBtn.style.right = `${defaultWidth}px`
      collapseBtn.style.display = 'flex'
      if (collapseIcon) collapseIcon.textContent = '▶'
      localStorage.setItem('sidebarWidth', defaultWidth)
      localStorage.removeItem('sidebarCollapsed')
    }
  }

  // Public API for debugging
  getCacheInfo() {
    return {
      size: this.previewCache.size,
      urls: Array.from(this.previewCache.keys()),
      totalSize: Array.from(this.previewCache.values()).reduce((sum, item) => sum + item.size, 0)
    }
  }

  clearCache() {
    this.previewCache.clear()
    console.log('✅ Preview cache cleared')
  }
}