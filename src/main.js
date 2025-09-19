/**
 * TicknTie - Main Application Entry
 * Initializes Univer spreadsheet and image sidebar plugin
 */

import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver, defaultTheme, LocaleType, Tools } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/lib/locales/en-US.js'
import { ImagePlugin } from './image-plugin'

// Initialize application
async function init() {
  try {
    console.log('🚀 Initializing TicknTie...')

    // Debug locale loading
    console.log('Loading preset locale...')
    console.log('Preset locale type:', typeof UniverPresetSheetsCoreEnUS)
    console.log('Preset locale keys (if object):', UniverPresetSheetsCoreEnUS ? Object.keys(UniverPresetSheetsCoreEnUS).slice(0, 10) : 'Not an object')

    // Check if it's a default export
    const locale = UniverPresetSheetsCoreEnUS.default || UniverPresetSheetsCoreEnUS
    console.log('Using locale:', typeof locale)
    if (typeof locale === 'object') {
      console.log('Locale sample keys:', Object.keys(locale).slice(0, 10))
    }

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
          console.log('Sheet initialized successfully')
          // Try to expand the sheet if possible
          const sheetData = sheet.getSheetData?.()
          if (sheetData) {
            console.log('Current sheet dimensions:', {
              rowCount: sheetData.rowCount,
              columnCount: sheetData.columnCount
            })
          }
        }
      }
    } catch (e) {
      console.log('Could not access sheet data:', e)
    }

    // Initialize image sidebar plugin
    const imagePlugin = new ImagePlugin(univerAPI)
    await imagePlugin.init()

    console.log('✅ TicknTie ready!')

    // Make plugin available globally for debugging
    window.ticknTie = {
      univerAPI,
      imagePlugin
    }

  } catch (error) {
    console.error('❌ Failed to initialize TicknTie:', error)
    document.getElementById('app').innerHTML = `
      <div style="padding: 40px; text-align: center;">
        <h2>Failed to initialize application</h2>
        <p>${error.message}</p>
        <p>Please refresh the page to try again.</p>
      </div>
    `
  }
}

// Start application when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init)
} else {
  init()
}