/**
 * TicknTie - Main Application Entry
 * Initializes Univer spreadsheet and image sidebar plugin
 */

import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver, defaultTheme, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/lib/locales/en-US.js'
import { ImagePlugin } from './image-plugin'

// Initialize application
async function init() {
  try {
    // Check if it's a default export
    const locale = UniverPresetSheetsCoreEnUS.default || UniverPresetSheetsCoreEnUS

    // Create Univer instance with spreadsheet
    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: locale
      },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({
          container: 'spreadsheet'
        })
      ]
    })

    // Create initial workbook with expanded column count
    // Default Univer only creates 20 columns (A-T), we need more for proper Excel-like functionality
    const workbook = univerAPI.createUniverSheet({
      name: 'TicknTie Workbook',
      sheets: {
        'sheet-01': {
          id: 'sheet-01',
          name: 'Sheet1',
          columnCount: 100,  // Support columns A through CV (100 columns)
          rowCount: 1000     // Keep default row count
        }
      }
    })

    // Get the active sheet and try to set column count
    try {
      const activeWorkbook = univerAPI.getActiveWorkbook()
      if (activeWorkbook) {
        const sheet = activeWorkbook.getActiveSheet()
        if (sheet) {
          // Try to expand the sheet if possible
          const sheetData = sheet.getSheetData?.()
        }
      }
    } catch (e) {
      // Silent error
    }

    // Initialize image sidebar plugin
    const imagePlugin = new ImagePlugin(univerAPI)
    await imagePlugin.init()

    // Application ready

    // Make plugin available globally for debugging
    window.ticknTie = {
      univerAPI,
      imagePlugin
    }

  } catch (error) {
    const errorDiv = document.createElement('div')
    errorDiv.style.padding = '40px'
    errorDiv.style.textAlign = 'center'

    const errorTitle = document.createElement('h2')
    errorTitle.textContent = 'Failed to initialize application'
    errorDiv.appendChild(errorTitle)

    const errorMsg = document.createElement('p')
    errorMsg.textContent = error.message
    errorDiv.appendChild(errorMsg)

    const errorHelp = document.createElement('p')
    errorHelp.textContent = 'Please refresh the page to try again.'
    errorDiv.appendChild(errorHelp)

    document.getElementById('app').innerHTML = ''
    document.getElementById('app').appendChild(errorDiv)
  }
}

// Start application when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init)
} else {
  init()
}