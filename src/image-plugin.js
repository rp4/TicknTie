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
    this.maxFileSize = 50 * 1024 * 1024 // 50MB limit
    this.cleanupInterval = null
    this.selectionCheckInterval = null
    this.eventListeners = [] // Track all event listeners for cleanup
    this.domCache = new Map() // Cache DOM queries
  }

  async init() {
    this.createSidebar()
    this.setupResizeHandler()
    this.listenToSelection()
    this.setupHyperlinkInput()
    this.startCacheCleanup()
  }

  destroy() {
    // Clean up all intervals
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval)
      this.cleanupInterval = null
    }
    if (this.selectionCheckInterval) {
      clearInterval(this.selectionCheckInterval)
      this.selectionCheckInterval = null
    }

    // Remove all event listeners
    this.eventListeners.forEach(({ element, event, handler }) => {
      element.removeEventListener(event, handler)
    })
    this.eventListeners = []

    // Clear caches
    this.previewCache.clear()
    this.domCache.clear()
  }

  getDOMElement(id) {
    if (!this.domCache.has(id)) {
      this.domCache.set(id, document.getElementById(id))
    }
    return this.domCache.get(id)
  }

  createSidebar() {
    const sidebar = this.getDOMElement('sidebar')

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

    // Use textContent and createElement for security
    const headerDiv = document.createElement('div')
    headerDiv.className = 'sidebar-header'

    // Header with title and download button
    const headerTop = document.createElement('div')
    headerTop.style.cssText = 'display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;'

    const h3 = document.createElement('h3')
    h3.textContent = 'Evidence Viewer'
    h3.style.margin = '0'
    headerTop.appendChild(h3)

    // Download project button
    const downloadBtn = document.createElement('button')
    downloadBtn.id = 'downloadProjectBtn'
    downloadBtn.style.cssText = `
      background: rgba(255, 255, 255, 0.2);
      border: 1px solid rgba(255, 255, 255, 0.3);
      color: white;
      padding: 6px 12px;
      border-radius: 4px;
      cursor: pointer;
      font-size: 13px;
      display: flex;
      align-items: center;
      gap: 5px;
      transition: background 0.2s;
    `
    downloadBtn.innerHTML = '💾 Export'
    downloadBtn.title = 'Download project as ZIP'
    downloadBtn.onmouseover = () => { downloadBtn.style.background = 'rgba(255, 255, 255, 0.3)' }
    downloadBtn.onmouseout = () => { downloadBtn.style.background = 'rgba(255, 255, 255, 0.2)' }
    downloadBtn.onclick = () => this.downloadProject()
    headerTop.appendChild(downloadBtn)

    headerDiv.appendChild(headerTop)

    const fileInput = document.createElement('input')
    fileInput.type = 'file'
    fileInput.id = 'fileInput'
    fileInput.accept = 'image/*,application/pdf'
    fileInput.style.display = 'none'
    headerDiv.appendChild(fileInput)

    const addFileBtn = document.createElement('button')
    addFileBtn.id = 'addFileBtn'
    addFileBtn.className = 'btn-primary'
    addFileBtn.textContent = '➕ Add Evidence'
    headerDiv.appendChild(addFileBtn)

    // Import project button
    const importBtn = document.createElement('button')
    importBtn.id = 'importProjectBtn'
    importBtn.className = 'btn-primary'
    importBtn.textContent = '📥 Import Project'
    importBtn.style.marginTop = '10px'
    importBtn.onclick = () => this.showImportDialog()
    headerDiv.appendChild(importBtn)

    const contentDiv = document.createElement('div')
    contentDiv.className = 'sidebar-content'

    const imageContainer = document.createElement('div')
    imageContainer.id = 'imageContainer'
    imageContainer.className = 'image-container'

    const emptyState = document.createElement('div')
    emptyState.className = 'empty-state'
    const emptyIcon = document.createElement('div')
    emptyIcon.className = 'empty-icon'
    emptyIcon.textContent = '📁'
    const emptyP = document.createElement('p')
    emptyP.textContent = 'No evidence in this cell'
    const emptySmall = document.createElement('small')
    emptySmall.textContent = 'Select a cell and add evidence'
    emptyState.appendChild(emptyIcon)
    emptyState.appendChild(emptyP)
    emptyState.appendChild(emptySmall)

    imageContainer.appendChild(emptyState)
    contentDiv.appendChild(imageContainer)

    sidebar.appendChild(headerDiv)
    sidebar.appendChild(contentDiv)

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

    // Smart polling - only when document is visible
    this.selectionCheckInterval = setInterval(() => {
      if (!document.hidden) {
        this.detectCurrentSelection()
        this.checkForDeletedContent()
      }
    }, 1000) // Reduced frequency

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
              cellsToRemove.push(cellRef)
            }
          }
        } catch (e) {
          // Silent error handling
        }
      }

      // Remove hyperlinks for cleared cells
      cellsToRemove.forEach(cellRef => {
        this.removeHyperlinkFromCell(cellRef)
      })

    } catch (e) {
      // Silent error handling
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

    // Silent removal
  }

  detectCurrentSelection() {
    try {
      // Method 1: Try to get selection from Univer API
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) {
        return
      }

      const worksheet = workbook.getActiveSheet()
      if (!worksheet) {
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
            this.selectCell(cellRef)
          }
        }
      }

    } catch (e) {
      // Silent error handling
    }
  }

  setupHyperlinkInput() {
    const addFileBtn = this.getDOMElement('addFileBtn')
    const fileInput = this.getDOMElement('fileInput')

    // Add file button - directly opens file picker
    const clickHandler = () => {
      if (!this.selectedCell) {
        alert('Please select a cell first')
        return
      }
      fileInput.click()
    }
    addFileBtn.addEventListener('click', clickHandler)
    this.eventListeners.push({ element: addFileBtn, event: 'click', handler: clickHandler })

    // File input change with validation
    const changeHandler = async (e) => {
      const file = e.target.files[0]
      if (file) {
        // Validate file type first
        const validTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'application/pdf']
        if (!validTypes.includes(file.type) && !file.name.match(/\.(jpg|jpeg|png|gif|webp|pdf)$/i)) {
          alert('Invalid file type. Please upload an image or PDF.')
          fileInput.value = ''
          return
        }

        // Only enforce size limit for images (PDFs only render first page anyway)
        if (!file.type.includes('pdf') && file.size > this.maxFileSize) {
          alert(`Image file too large. Maximum size is ${this.maxFileSize / 1024 / 1024}MB`)
          fileInput.value = ''
          return
        }

        await this.handleFileUpload(file)
        fileInput.value = ''
      }
    }
    fileInput.addEventListener('change', changeHandler)
    this.eventListeners.push({ element: fileInput, event: 'change', handler: changeHandler })
  }

  async handleFileUpload(file) {
    if (!this.selectedCell) return

    try {
      // Convert file to data URL
      const url = await this.fileToDataUrl(file)
      const displayText = file.name

      // Add to cell
      await this.addHyperlinkToCell(url, displayText)

      // Silent success
    } catch (error) {
      // Silent error
      alert('Failed to upload file. Please try again.')
    }
  }

  // Convert file to compressed data URL
  async fileToDataUrl(file) {
    // For images, compress them
    if (file.type.startsWith('image/')) {
      return this.compressImage(file)
    }

    // For PDFs and other files, use regular data URL
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = (e) => resolve(e.target.result)
      reader.onerror = reject
      reader.readAsDataURL(file)
    })
  }

  async compressImage(file) {
    return new Promise((resolve, reject) => {
      const img = new Image()
      const reader = new FileReader()

      reader.onload = (e) => {
        img.src = e.target.result
      }

      img.onload = () => {
        const canvas = document.createElement('canvas')
        const ctx = canvas.getContext('2d')

        // Max dimensions for storage
        const maxWidth = 1920
        const maxHeight = 1080

        let width = img.width
        let height = img.height

        if (width > maxWidth || height > maxHeight) {
          const scale = Math.min(maxWidth / width, maxHeight / height)
          width *= scale
          height *= scale
        }

        canvas.width = width
        canvas.height = height
        ctx.drawImage(img, 0, 0, width, height)

        // Convert to JPEG with compression
        resolve(canvas.toDataURL('image/jpeg', 0.85))
      }

      img.onerror = reject
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

      // Silent success
    } catch (error) {
      // Silent error
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
    const container = this.getDOMElement('imageContainer')

    // Check cache first
    if (this.previewCache.has(url)) {
      const cached = this.previewCache.get(url)
      // Update access time for LRU
      cached.timestamp = Date.now()
      this.showPreviewContent(cached.preview, url, displayText)
      return
    }

    // Show loading state safely
    const loadingDiv = document.createElement('div')
    loadingDiv.className = 'preview-loading'

    const spinner = document.createElement('div')
    spinner.className = 'loading-spinner'
    loadingDiv.appendChild(spinner)

    const loadingText = document.createElement('p')
    loadingText.textContent = 'Loading preview...'
    loadingDiv.appendChild(loadingText)

    const fileName = document.createElement('small')
    fileName.textContent = displayText
    loadingDiv.appendChild(fileName)

    container.innerHTML = ''
    container.appendChild(loadingDiv)

    try {
      const preview = await this.fetchPreview(url)
      if (preview) {
        // Cache the preview
        this.addToCache(url, preview)
        this.showPreviewContent(preview, url, displayText)
      }
    } catch (error) {
      // Silent error
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
    const isPdf = dataUrl.startsWith('data:application/pdf')

    if (isPdf) {
      return this.createPDFPreviewFromDataUrl(dataUrl)
    } else {
      return this.createImagePreview(dataUrl)
    }
  }

  async fetchImagePreview(url) {
    return this.createImagePreview(url).catch(() => {
      throw new Error('Failed to load image')
    })
  }

  // Unified image preview creation function
  async createImagePreview(url) {
    return new Promise((resolve) => {
      const img = new Image()
      if (!url.startsWith('data:')) {
        img.crossOrigin = 'anonymous' // Handle CORS if possible
      }

      img.onload = () => {
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

      img.onerror = () => {
        resolve(null)
      }

      img.src = url
    })
  }

  async createPDFPreviewFromDataUrl(dataUrl) {
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
    const container = this.getDOMElement('imageContainer')
    container.innerHTML = ''

    const previewDiv = document.createElement('div')
    previewDiv.className = 'image-preview'

    // Make image clickable for fullscreen view
    const img = document.createElement('img')
    img.src = preview.dataUrl
    img.alt = displayText
    img.style.maxWidth = '100%'
    img.style.height = 'auto'
    img.style.cursor = 'zoom-in'
    img.title = 'Click to view fullscreen'
    img.onclick = () => this.showFullscreenView(url, displayText, preview.type)
    previewDiv.appendChild(img)

    if (preview.type === 'pdf') {
      const pdfNotice = document.createElement('div')
      pdfNotice.className = 'pdf-notice'
      pdfNotice.textContent = `📄 PDF Preview (Page 1 of ${preview.pageCount || '?'})`
      previewDiv.appendChild(pdfNotice)
    }

    // View fullscreen button
    const viewBtn = document.createElement('button')
    viewBtn.className = 'btn-primary'
    viewBtn.style.width = '100%'
    viewBtn.style.marginTop = '10px'
    viewBtn.textContent = '🔍 View Fullscreen'
    viewBtn.onclick = () => this.showFullscreenView(url, displayText, preview.type)
    previewDiv.appendChild(viewBtn)

    container.appendChild(previewDiv)
  }

  showFullscreenView(url, filename, type) {
    // Create fullscreen overlay
    const overlay = document.createElement('div')
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(0, 0, 0, 0.95);
      z-index: 10000;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 20px;
    `

    // Header with filename and close button
    const header = document.createElement('div')
    header.style.cssText = `
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      background: rgba(0, 0, 0, 0.8);
      color: white;
      padding: 15px 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    `

    const title = document.createElement('span')
    title.textContent = filename
    title.style.fontSize = '16px'
    header.appendChild(title)

    const closeBtn = document.createElement('button')
    closeBtn.textContent = '✕'
    closeBtn.style.cssText = `
      background: none;
      border: none;
      color: white;
      font-size: 24px;
      cursor: pointer;
      padding: 0 10px;
    `
    closeBtn.onclick = () => document.body.removeChild(overlay)
    header.appendChild(closeBtn)

    overlay.appendChild(header)

    // Content area
    const content = document.createElement('div')
    content.style.cssText = `
      max-width: 90vw;
      max-height: 85vh;
      overflow: auto;
      margin-top: 60px;
      position: relative;
    `

    if (type === 'pdf') {
      // Show loading spinner for PDFs
      const loadingContainer = document.createElement('div')
      loadingContainer.style.cssText = `
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        text-align: center;
        color: white;
      `

      const spinner = document.createElement('div')
      spinner.style.cssText = `
        width: 50px;
        height: 50px;
        border: 4px solid rgba(255, 255, 255, 0.3);
        border-top-color: white;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin: 0 auto 20px;
      `

      const loadingText = document.createElement('div')
      loadingText.textContent = 'Loading PDF...'
      loadingText.style.fontSize = '16px'

      loadingContainer.appendChild(spinner)
      loadingContainer.appendChild(loadingText)
      content.appendChild(loadingContainer)

      // For PDFs, show embedded viewer
      const embed = document.createElement('embed')
      embed.src = url
      embed.type = 'application/pdf'
      embed.style.cssText = `
        width: 90vw;
        height: 85vh;
      `

      // Hide loading when PDF loads
      embed.onload = () => {
        if (loadingContainer.parentNode) {
          content.removeChild(loadingContainer)
        }
      }

      // Also hide loading after a timeout (fallback for browsers that don't fire onload for embed)
      setTimeout(() => {
        if (loadingContainer.parentNode) {
          content.removeChild(loadingContainer)
        }
      }, 2000)

      content.appendChild(embed)
    } else {
      // Show loading for large images too
      const loadingContainer = document.createElement('div')
      loadingContainer.style.cssText = `
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        text-align: center;
        color: white;
      `

      const spinner = document.createElement('div')
      spinner.style.cssText = `
        width: 50px;
        height: 50px;
        border: 4px solid rgba(255, 255, 255, 0.3);
        border-top-color: white;
        border-radius: 50%;
        animation: spin 1s linear infinite;
        margin: 0 auto 20px;
      `

      const loadingText = document.createElement('div')
      loadingText.textContent = 'Loading image...'
      loadingText.style.fontSize = '16px'

      loadingContainer.appendChild(spinner)
      loadingContainer.appendChild(loadingText)
      content.appendChild(loadingContainer)

      // For images, show full resolution
      const img = document.createElement('img')
      img.style.cssText = `
        max-width: 90vw;
        max-height: 85vh;
        object-fit: contain;
        display: none;
      `

      img.onload = () => {
        if (loadingContainer.parentNode) {
          content.removeChild(loadingContainer)
        }
        img.style.display = 'block'
      }

      img.onerror = () => {
        if (loadingContainer.parentNode) {
          loadingText.textContent = 'Failed to load image'
          spinner.style.display = 'none'
        }
      }

      img.src = url
      content.appendChild(img)
    }

    overlay.appendChild(content)

    // Close on ESC key
    const handleEsc = (e) => {
      if (e.key === 'Escape') {
        document.body.removeChild(overlay)
        document.removeEventListener('keydown', handleEsc)
      }
    }
    document.addEventListener('keydown', handleEsc)

    // Close on overlay click (but not content)
    overlay.onclick = (e) => {
      if (e.target === overlay) {
        document.body.removeChild(overlay)
        document.removeEventListener('keydown', handleEsc)
      }
    }

    document.body.appendChild(overlay)
  }

  showErrorState(url, displayText, errorMessage) {
    const container = this.getDOMElement('imageContainer')
    container.innerHTML = ''

    const errorDiv = document.createElement('div')
    errorDiv.className = 'preview-error'

    const errorIcon = document.createElement('div')
    errorIcon.className = 'error-icon'
    errorIcon.textContent = '⚠️'
    errorDiv.appendChild(errorIcon)

    const errorTitle = document.createElement('p')
    const strong = document.createElement('strong')
    strong.textContent = 'Failed to load preview'
    errorTitle.appendChild(strong)
    errorDiv.appendChild(errorTitle)

    const fileName = document.createElement('small')
    fileName.textContent = displayText
    errorDiv.appendChild(fileName)

    const errorMsg = document.createElement('small')
    errorMsg.style.color = '#666'
    errorMsg.textContent = errorMessage
    errorDiv.appendChild(errorMsg)

    const retryBtn = document.createElement('button')
    retryBtn.className = 'retry-btn'
    retryBtn.textContent = 'Retry'
    retryBtn.onclick = () => this.retryPreview(url, displayText)
    errorDiv.appendChild(retryBtn)

    const viewBtn = document.createElement('button')
    viewBtn.className = 'btn-primary'
    viewBtn.style.marginTop = '8px'
    viewBtn.style.width = '100%'
    viewBtn.textContent = '🔍 Try Fullscreen View'
    viewBtn.onclick = () => this.showFullscreenView(url, displayText, 'image')
    errorDiv.appendChild(viewBtn)

    container.appendChild(errorDiv)
  }

  async retryPreview(url, displayText) {
    // Clear from cache to force re-fetch
    this.previewCache.delete(url)
    await this.displayPreview(url, displayText)
  }

  displayEmptyState() {
    const container = this.getDOMElement('imageContainer')
    container.innerHTML = ''

    const emptyDiv = document.createElement('div')
    emptyDiv.className = 'empty-state'

    const emptyIcon = document.createElement('div')
    emptyIcon.className = 'empty-icon'
    emptyIcon.textContent = '📁'
    emptyDiv.appendChild(emptyIcon)

    const emptyText = document.createElement('p')
    emptyText.textContent = 'No evidence in this cell'
    emptyDiv.appendChild(emptyText)

    const emptyHelp = document.createElement('small')
    emptyHelp.textContent = 'Click "Add Evidence" to attach a file'
    emptyDiv.appendChild(emptyHelp)

    container.appendChild(emptyDiv)
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
      // Silent preload failure
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
    this.cleanupInterval = setInterval(() => {
      const now = Date.now()
      const toDelete = []

      for (const [url, data] of this.previewCache) {
        if (now - data.timestamp > this.maxCacheAge) {
          toDelete.push(url)
        }
      }

      toDelete.forEach(url => this.previewCache.delete(url))
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
      // Silent error
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
    const sidebar = this.getDOMElement('sidebar')
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
      // Style download button for collapsed state
      const downloadBtn = document.getElementById('downloadProjectBtn')
      if (downloadBtn) {
        downloadBtn.innerHTML = '💾'
        downloadBtn.style.cssText = `
          background: transparent;
          border: none;
          color: white;
          padding: 5px;
          cursor: pointer;
          font-size: 18px;
          position: absolute;
          top: 10px;
          left: 50%;
          transform: translateX(-50%);
          transition: transform 0.2s;
        `
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
    const sidebar = this.getDOMElement('sidebar')
    const collapseBtn = this.getDOMElement('collapseBtn')

    if (!collapseBtn) return

    const toggleHandler = () => this.toggleSidebar()
    collapseBtn.addEventListener('click', toggleHandler)
    this.eventListeners.push({ element: collapseBtn, event: 'click', handler: toggleHandler })

    const sidebarHandler = (e) => {
      if (sidebar.classList.contains('collapsed')) {
        e.stopPropagation()
        this.toggleSidebar()
      }
    }
    sidebar.addEventListener('click', sidebarHandler)
    this.eventListeners.push({ element: sidebar, event: 'click', handler: sidebarHandler })
  }

  toggleSidebar() {
    const sidebar = this.getDOMElement('sidebar')
    const collapseBtn = this.getDOMElement('collapseBtn')
    const collapseIcon = collapseBtn?.querySelector('.collapse-icon')
    const downloadBtn = document.getElementById('downloadProjectBtn')
    const collapsedWidth = 40
    const defaultWidth = 350

    if (sidebar.classList.contains('collapsed')) {
      // Expand
      sidebar.style.width = defaultWidth + 'px'
      sidebar.classList.remove('collapsed')
      collapseBtn.style.right = `${defaultWidth}px`
      collapseBtn.style.display = 'flex'  // Show button when expanded
      if (collapseIcon) collapseIcon.textContent = '▶'
      if (downloadBtn) {
        downloadBtn.innerHTML = '💾 Export'
        downloadBtn.style.cssText = `
          background: rgba(255, 255, 255, 0.2);
          border: 1px solid rgba(255, 255, 255, 0.3);
          color: white;
          padding: 6px 12px;
          border-radius: 4px;
          cursor: pointer;
          font-size: 13px;
          display: flex;
          align-items: center;
          gap: 5px;
          transition: background 0.2s;
          position: static;
          top: auto;
          right: auto;
        `
      }
      localStorage.setItem('sidebarWidth', defaultWidth)
      localStorage.removeItem('sidebarCollapsed')
    } else {
      // Collapse
      sidebar.style.width = collapsedWidth + 'px'
      sidebar.classList.add('collapsed')
      collapseBtn.style.display = 'none'  // Hide button when collapsed
      if (collapseIcon) collapseIcon.textContent = '◀'
      if (downloadBtn) {
        downloadBtn.innerHTML = '💾'
        downloadBtn.style.cssText = `
          background: transparent;
          border: none;
          color: white;
          padding: 5px;
          cursor: pointer;
          font-size: 18px;
          position: absolute;
          top: 10px;
          left: 50%;
          transform: translateX(-50%);
          transition: transform 0.2s;
        `
      }
      localStorage.setItem('sidebarWidth', collapsedWidth)
      localStorage.setItem('sidebarCollapsed', 'true')
    }
  }

  expandSidebar() {
    const sidebar = this.getDOMElement('sidebar')
    const collapseBtn = this.getDOMElement('collapseBtn')
    const collapseIcon = collapseBtn?.querySelector('.collapse-icon')
    const downloadBtn = document.getElementById('downloadProjectBtn')

    if (sidebar.classList.contains('collapsed')) {
      const defaultWidth = 350
      sidebar.style.width = defaultWidth + 'px'
      sidebar.classList.remove('collapsed')
      collapseBtn.style.right = `${defaultWidth}px`
      collapseBtn.style.display = 'flex'
      if (collapseIcon) collapseIcon.textContent = '▶'
      if (downloadBtn) {
        downloadBtn.innerHTML = '💾 Export'
        downloadBtn.style.cssText = `
          background: rgba(255, 255, 255, 0.2);
          border: 1px solid rgba(255, 255, 255, 0.3);
          color: white;
          padding: 6px 12px;
          border-radius: 4px;
          cursor: pointer;
          font-size: 13px;
          display: flex;
          align-items: center;
          gap: 5px;
          transition: background 0.2s;
          position: static;
          top: auto;
          right: auto;
        `
      }
      localStorage.setItem('sidebarWidth', defaultWidth)
      localStorage.removeItem('sidebarCollapsed')
    }
  }

  async downloadProject() {
    try {
      // Import JSZip dynamically
      const JSZip = (await import('jszip')).default
      const zip = new JSZip()

      // Create project name with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
      const projectName = `TicknTie_Project_${timestamp}`

      // Get spreadsheet data from Univer
      const workbook = this.univerAPI.getActiveWorkbook()
      if (!workbook) {
        alert('No workbook found to export')
        return
      }

      // Create evidence folder and prepare file mapping FIRST
      const evidenceFolder = zip.folder('evidence')
      const fileMapping = new Map() // Map cellRef to evidence file path
      const usedFileNames = new Set()
      let fileCounter = 1

      // Process each hyperlink and create evidence files
      for (const [cellRef, linkData] of this.cellHyperlinks) {
        const { url, displayText } = linkData

        if (url.startsWith('data:')) {
          // Use original filename if available, otherwise generate one
          let fileName = displayText

          // Sanitize filename (remove invalid characters)
          fileName = fileName.replace(/[<>:"/\\|?*]/g, '_')

          // Ensure unique filename
          let finalFileName = fileName
          let counter = 1
          while (usedFileNames.has(finalFileName)) {
            const nameParts = fileName.split('.')
            if (nameParts.length > 1) {
              const ext = nameParts.pop()
              const base = nameParts.join('.')
              finalFileName = `${base}_${counter}.${ext}`
            } else {
              finalFileName = `${fileName}_${counter}`
            }
            counter++
          }
          usedFileNames.add(finalFileName)

          // Extract base64 data and convert to blob
          const base64Data = url.split(',')[1]
          const mimeType = url.match(/data:([^;]+)/)?.[1] || 'application/octet-stream'
          const byteCharacters = atob(base64Data)
          const byteNumbers = new Array(byteCharacters.length)
          for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i)
          }
          const byteArray = new Uint8Array(byteNumbers)
          const blob = new Blob([byteArray], { type: mimeType })

          // Add file to evidence folder
          evidenceFolder.file(finalFileName, blob)
          fileMapping.set(cellRef, `evidence/${finalFileName}`)

          fileCounter++
        }
      }

      // Export Excel file with hyperlinks
      console.log('Exporting Excel file with hyperlinks...')
      const excelBlob = await this.exportExcelWithHyperlinks(workbook, fileMapping)
      if (excelBlob) {
        // Always use .xlsx extension now that we're using SheetJS
        const filename = 'workbook.xlsx'
        zip.file(filename, excelBlob)
        console.log(`Added ${filename} to ZIP with ${fileMapping.size} hyperlinks`)
      } else {
        console.error('Failed to generate Excel file - excelBlob is null')
        alert('Warning: Excel file could not be exported. The ZIP will contain evidence files only.')
      }

      // Create README file
      const readme = `# ${projectName}

## Contents
- workbook.xlsx: The main spreadsheet with evidence references (Excel-compatible)
- evidence/: Folder containing all attached evidence files

## Evidence Files
${fileCounter > 1 ? `This project contains ${fileCounter - 1} evidence files.` : 'No evidence files attached.'}

## Instructions
1. Extract this ZIP file to a folder
2. Open workbook.xlsx in Excel or any spreadsheet application
3. Evidence files are referenced in cells with 📌 markers
4. View evidence files in the evidence/ folder

## Created
${new Date().toLocaleString()}

## Created with
TicknTie - Open source audit evidence management
https://github.com/yourusername/tickntie
`
      zip.file('README.txt', readme)

      // Generate and download the ZIP file
      const zipBlob = await zip.generateAsync({ type: 'blob' })

      // Create download link
      const link = document.createElement('a')
      link.href = URL.createObjectURL(zipBlob)
      link.download = `${projectName}.zip`
      document.body.appendChild(link)
      link.click()
      document.body.removeChild(link)
      URL.revokeObjectURL(link.href)

      // Show success message
      this.showExportSuccess(fileCounter - 1)

    } catch (error) {
      alert('Failed to export project: ' + error.message)
    }
  }

  async exportExcelWithHyperlinks(workbook, fileMapping) {
    try {
      console.log('Starting Excel export with hyperlinks...')

      // Use SheetJS to create a proper XLSX file
      const xlsxModule = await import('xlsx')
      const XLSX = xlsxModule.default || xlsxModule
      console.log('SheetJS loaded for hyperlink export')

      const sheet = workbook.getActiveSheet()
      if (!sheet) {
        console.error('No active sheet found')
        return null
      }

      // First, collect all data as before
      const data = []
      const maxRow = 100
      const maxCol = 26

      for (let r = 0; r < maxRow; r++) {
        const row = []
        let hasData = false
        for (let c = 0; c < maxCol; c++) {
          try {
            const range = sheet.getRange(r, c)
            const value = range?.getValue?.() || range?.value || ''
            row.push(value)
            if (value) hasData = true
          } catch (e) {
            row.push('')
          }
        }
        if (hasData) {
          data.push(row)
        }
      }

      // If no data, at least add headers
      if (data.length === 0) {
        data.push(['']) // Add at least one cell
      }

      // Create workbook and worksheet
      const wb = XLSX.utils.book_new()
      const ws = XLSX.utils.aoa_to_sheet(data)

      // Add hyperlinks to specific cells
      let hyperlinkCount = 0
      for (const [cellRef, evidencePath] of fileMapping.entries()) {
        const linkData = this.cellHyperlinks.get(cellRef)
        if (linkData) {
          const displayName = linkData.displayText || 'Evidence'

          // Set the cell to have both value and hyperlink using the cell object format
          if (ws[cellRef]) {
            // Create hyperlink using HYPERLINK formula
            ws[cellRef] = {
              f: `HYPERLINK("${evidencePath}","📌 ${displayName}")`,
              v: `📌 ${displayName}`,
              t: 's'
            }
            hyperlinkCount++
            console.log(`Added hyperlink to cell ${cellRef}: ${evidencePath}`)
          } else {
            // Cell doesn't exist in worksheet, need to create it
            const { row, col } = this.parseCellRef(cellRef)
            if (row < data.length) {
              ws[cellRef] = {
                f: `HYPERLINK("${evidencePath}","📌 ${displayName}")`,
                v: `📌 ${displayName}`,
                t: 's'
              }
              hyperlinkCount++
            }
          }
        }
      }

      console.log(`Added ${hyperlinkCount} hyperlinks to worksheet using HYPERLINK formula`)

      // Update the range if we added cells
      if (!ws['!ref']) {
        ws['!ref'] = 'A1:Z100'
      }

      XLSX.utils.book_append_sheet(wb, ws, "Sheet1")

      // Generate XLSX binary string with formulas
      const xlsxData = XLSX.write(wb, {
        bookType: 'xlsx',
        type: 'array',
        cellFormula: true,
        bookSST: true
      })
      console.log(`Generated XLSX with hyperlinks: ${xlsxData.byteLength} bytes`)

      // Return as blob
      return new Blob([xlsxData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    } catch (error) {
      console.error('Export with hyperlinks failed:', error)
      // Fall back to regular export
      return this.exportExcel(workbook)
    }
  }

  async exportExcel(workbook) {
    try {
      console.log('Starting Excel export...')
      console.log('Workbook object:', workbook)
      console.log('Workbook methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(workbook || {})))

      // Try to export using Univer API first
      if (workbook.export) {
        console.log('Using workbook.export method')
        return await workbook.export('xlsx')
      } else if (workbook.saveAsBlob) {
        console.log('Using workbook.saveAsBlob method')
        return await workbook.saveAsBlob()
      } else {
        console.log('Using SheetJS for XLSX export')
        // Use SheetJS to create a proper XLSX file
        const xlsxModule = await import('xlsx')
        const XLSX = xlsxModule.default || xlsxModule
        console.log('SheetJS loaded successfully', XLSX)

        const sheet = workbook.getActiveSheet()
        console.log('Sheet object:', sheet)
        if (!sheet) {
          console.error('No active sheet found')
          return null
        }
        console.log('Sheet methods:', Object.getOwnPropertyNames(Object.getPrototypeOf(sheet || {})))

        // Get sheet data
        const data = []
        const maxRow = 100 // Limit for export
        const maxCol = 26 // A-Z columns
        console.log(`Reading data from sheet (max ${maxRow} rows, ${maxCol} cols)...`)

        // Try different methods to get data
        try {
          // Method 1: Using getRange
          for (let r = 0; r < maxRow; r++) {
            const row = []
            let hasData = false
            for (let c = 0; c < maxCol; c++) {
              try {
                const range = sheet.getRange(r, c)
                const value = range?.getValue?.() || range?.value || ''
                row.push(value)
                if (value) hasData = true
              } catch (e) {
                row.push('')
              }
            }
            // Only add non-empty rows
            if (hasData) {
              data.push(row)
            }
          }
        } catch (error) {
          console.error('Failed to read sheet data with getRange:', error)

          // Method 2: Try to get all data at once
          try {
            const allData = sheet.getDataRange?.()?.getValues?.() ||
                           sheet.getData?.() ||
                           sheet.getValues?.()
            if (allData && Array.isArray(allData)) {
              data.push(...allData.slice(0, maxRow))
              console.log('Used alternative method to get sheet data')
            }
          } catch (e) {
            console.error('Alternative method also failed:', e)
          }
        }

        console.log(`Collected ${data.length} rows of data`)

        // If no data, at least add headers
        if (data.length === 0) {
          console.log('No data found, adding empty row')
          data.push(['']) // Add at least one cell
        }

        // Create workbook and worksheet
        const wb = XLSX.utils.book_new()
        const ws = XLSX.utils.aoa_to_sheet(data)
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1")
        console.log('Created SheetJS workbook')

        // Generate XLSX binary string
        const xlsxData = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
        console.log(`Generated XLSX data: ${xlsxData.byteLength} bytes`)

        // Return as blob
        const blob = new Blob([xlsxData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        console.log(`Created blob: ${blob.size} bytes`)
        return blob
      }
    } catch (error) {
      console.error('Export failed with error:', error)
      console.error('Error stack:', error.stack)
      // Return null if export fails
      return null
    }
  }

  getFileExtension(dataUrl, displayText) {
    // Try to get extension from display text first
    if (displayText) {
      const match = displayText.match(/\.[^.]+$/)
      if (match) return match[0]
    }

    // Fallback to MIME type
    if (dataUrl.includes('image/jpeg')) return '.jpg'
    if (dataUrl.includes('image/png')) return '.png'
    if (dataUrl.includes('image/gif')) return '.gif'
    if (dataUrl.includes('image/webp')) return '.webp'
    if (dataUrl.includes('application/pdf')) return '.pdf'

    return '.dat' // Generic extension
  }

  showExportSuccess(fileCount) {
    // Create success notification
    const notification = document.createElement('div')
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #28a745;
      color: white;
      padding: 15px 20px;
      border-radius: 6px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
      z-index: 10001;
      animation: slideIn 0.3s ease;
    `
    notification.innerHTML = `
      <div style="display: flex; align-items: center; gap: 10px;">
        <span style="font-size: 20px;">✅</span>
        <div>
          <div style="font-weight: 600;">Project exported successfully!</div>
          <div style="font-size: 13px; opacity: 0.9;">${fileCount} evidence file${fileCount !== 1 ? 's' : ''} included</div>
        </div>
      </div>
    `

    // Add slide-in animation
    const style = document.createElement('style')
    style.textContent = `
      @keyframes slideIn {
        from { transform: translateX(400px); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
      }
    `
    document.head.appendChild(style)

    document.body.appendChild(notification)

    // Auto-remove after 5 seconds
    setTimeout(() => {
      notification.style.animation = 'slideIn 0.3s ease reverse'
      setTimeout(() => {
        document.body.removeChild(notification)
        document.head.removeChild(style)
      }, 300)
    }, 5000)
  }

  showImportDialog() {
    // Create file input for ZIP
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.zip'
    input.onchange = async (e) => {
      const file = e.target.files[0]
      if (file) {
        await this.importProject(file)
      }
    }
    input.click()
  }

  async importProject(zipFile) {
    try {
      // Show loading indicator
      const loadingDiv = document.createElement('div')
      loadingDiv.id = 'importLoading'
      loadingDiv.style.cssText = `
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        background: white;
        padding: 30px;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
        z-index: 10002;
        text-align: center;
      `
      loadingDiv.innerHTML = `
        <div class="loading-spinner" style="margin: 0 auto 20px;"></div>
        <div>Importing project...</div>
      `
      document.body.appendChild(loadingDiv)

      // Import JSZip
      const JSZip = (await import('jszip')).default
      const zip = await JSZip.loadAsync(zipFile)

      // Clear existing data
      const confirmClear = confirm('This will replace current spreadsheet data. Continue?')
      if (!confirmClear) {
        document.body.removeChild(loadingDiv)
        return
      }

      // Clear current hyperlinks
      this.cellHyperlinks.clear()

      // First, load all evidence files and create a mapping
      const evidenceMapping = new Map() // filename -> dataUrl
      const evidenceFolder = zip.folder('evidence')
      const evidenceFiles = []

      if (evidenceFolder) {
        evidenceFolder.forEach((relativePath, file) => {
          if (!file.dir) { // Skip directories
            evidenceFiles.push({ path: relativePath, file })
          }
        })

        // Process each evidence file and store in mapping
        for (const { path, file } of evidenceFiles) {
          const blob = await file.async('blob')
          const dataUrl = await this.blobToDataUrl(blob)
          const fileName = path.split('/').pop()
          evidenceMapping.set(fileName, dataUrl)
        }
      }

      // Now import the workbook and restore cells with evidence
      let workbookFile = zip.file('workbook.xlsx') || zip.file('workbook.csv')

      if (workbookFile) {
        console.log(`Found workbook file: ${workbookFile.name}`)

        if (workbookFile.name.endsWith('.xlsx')) {
          // Import XLSX using SheetJS
          const arrayBuffer = await workbookFile.async('arraybuffer')
          await this.importXLSXWithEvidence(arrayBuffer, evidenceMapping)
        } else if (workbookFile.name.endsWith('.csv')) {
          // Fallback to CSV import
          const content = await workbookFile.async('string')
          await this.importCSVWithEvidence(content, evidenceMapping)
        }
      } else {
        console.warn('No workbook.xlsx or workbook.csv found in ZIP')
      }

      // Remove loading indicator
      document.body.removeChild(loadingDiv)

      // Trigger selection update to refresh sidebar if needed
      this.detectCurrentSelection()

      // Show success
      this.showImportSuccess(evidenceFiles.length)

    } catch (error) {
      // Remove loading if it exists
      const loading = document.getElementById('importLoading')
      if (loading) document.body.removeChild(loading)

      alert('Failed to import project: ' + error.message)
    }
  }

  async importXLSXWithEvidence(arrayBuffer, evidenceMapping) {
    try {
      const xlsxModule = await import('xlsx')
      const XLSX = xlsxModule.default || xlsxModule

      // Read the XLSX file
      const workbook = XLSX.read(arrayBuffer, { type: 'array' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      // Convert to array of arrays
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

      // Get Univer workbook and sheet
      const univerWorkbook = this.univerAPI.getActiveWorkbook()
      if (!univerWorkbook) return

      const sheet = univerWorkbook.getActiveSheet()
      if (!sheet) return

      // Import data and restore hyperlinks
      data.forEach((row, rowIndex) => {
        row.forEach((cellValue, colIndex) => {
          if (cellValue && cellValue.toString().trim()) {
            const range = sheet.getRange(rowIndex, colIndex)
            if (range?.setValue) {
              const cellStr = cellValue.toString()
              // Check if this cell contains an evidence reference
              const evidenceMatch = cellStr.match(/\[\[evidence\/([^\]]+)\]\]/)

              if (evidenceMatch) {
                const fileName = evidenceMatch[1]
                const cleanValue = cellStr.replace(/\[\[evidence\/[^\]]+\]\]/, '').trim()

                // Set the cell value without the reference marker
                range.setValue(cleanValue)

                // If we have this evidence file, restore the hyperlink
                if (evidenceMapping.has(fileName)) {
                  const cellRef = this.getCellRef(rowIndex, colIndex)
                  this.cellHyperlinks.set(cellRef, {
                    url: evidenceMapping.get(fileName),
                    displayText: fileName
                  })

                  console.log(`Restored evidence link for cell ${cellRef}: ${fileName}`)
                }
              } else {
                // Regular cell without evidence
                range.setValue(cellStr)
              }
            }
          }
        })
      })

      console.log(`Imported ${data.length} rows from XLSX with ${this.cellHyperlinks.size} evidence links`)
    } catch (error) {
      console.error('XLSX import failed:', error)
      // Fall back to basic import without evidence restoration
      throw error
    }
  }

  async importCSVWithEvidence(csvContent, evidenceMapping) {
    // Parse CSV properly handling quotes and commas
    const rows = []
    const lines = csvContent.split('\n')

    for (const line of lines) {
      if (line.trim()) { // Skip empty lines
        const cells = []
        let current = ''
        let inQuotes = false

        for (let i = 0; i < line.length; i++) {
          const char = line[i]
          const nextChar = line[i + 1]

          if (char === '"') {
            if (inQuotes && nextChar === '"') {
              current += '"'
              i++ // Skip next quote
            } else {
              inQuotes = !inQuotes
            }
          } else if (char === ',' && !inQuotes) {
            cells.push(current)
            current = ''
          } else {
            current += char
          }
        }
        cells.push(current) // Add last cell
        rows.push(cells)
      }
    }

    // Clear existing sheet first
    const workbook = this.univerAPI.getActiveWorkbook()
    if (!workbook) return

    const sheet = workbook.getActiveSheet()
    if (!sheet) return

    // Clear existing content (optional - you might want to keep this)
    // for (let row = 0; row < 100; row++) {
    //   for (let col = 0; col < 26; col++) {
    //     const range = sheet.getRange(row, col)
    //     if (range?.setValue) {
    //       range.setValue('')
    //     }
    //   }
    // }

    // Import data and restore hyperlinks
    rows.forEach((row, rowIndex) => {
      row.forEach((cellValue, colIndex) => {
        if (cellValue && cellValue.trim()) {
          const range = sheet.getRange(rowIndex, colIndex)
          if (range?.setValue) {
            // Check if this cell contains an evidence reference
            const evidenceMatch = cellValue.match(/\[\[evidence\/([^\]]+)\]\]/)

            if (evidenceMatch) {
              const fileName = evidenceMatch[1]
              const cleanValue = cellValue.replace(/\[\[evidence\/[^\]]+\]\]/, '').trim()

              // Set the cell value without the reference marker
              range.setValue(cleanValue)

              // If we have this evidence file, restore the hyperlink
              if (evidenceMapping.has(fileName)) {
                const cellRef = this.getCellRef(rowIndex, colIndex)
                this.cellHyperlinks.set(cellRef, {
                  url: evidenceMapping.get(fileName),
                  displayText: fileName
                })

                console.log(`Restored evidence link for cell ${cellRef}: ${fileName}`)
              }
            } else {
              // Regular cell without evidence
              range.setValue(cellValue)
            }
          }
        }
      })
    })

    console.log(`Imported ${rows.length} rows with ${this.cellHyperlinks.size} evidence links`)
  }


  async blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = (e) => resolve(e.target.result)
      reader.onerror = reject
      reader.readAsDataURL(blob)
    })
  }

  showImportSuccess(fileCount) {
    const notification = document.createElement('div')
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #17a2b8;
      color: white;
      padding: 15px 20px;
      border-radius: 6px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
      z-index: 10001;
      animation: slideIn 0.3s ease;
    `
    notification.innerHTML = `
      <div style="display: flex; align-items: center; gap: 10px;">
        <span style="font-size: 20px;">📥</span>
        <div>
          <div style="font-weight: 600;">Project imported successfully!</div>
          <div style="font-size: 13px; opacity: 0.9;">${fileCount} evidence file${fileCount !== 1 ? 's' : ''} restored</div>
        </div>
      </div>
    `

    document.body.appendChild(notification)

    setTimeout(() => {
      notification.style.animation = 'slideIn 0.3s ease reverse'
      setTimeout(() => {
        document.body.removeChild(notification)
      }, 300)
    }, 5000)
  }

}