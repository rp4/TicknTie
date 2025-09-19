/**
 * TicknTie - Main Application Entry
 * Initializes Univer spreadsheet and image sidebar plugin
 */

import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver, defaultTheme, LocaleType, Tools } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import enUS from '@univerjs/ui/lib/locale/en-US.js'
import { ImagePlugin } from './image-plugin'

// Initialize application
async function init() {
  try {
    console.log('🚀 Initializing TicknTie...')

    // Create Univer instance with spreadsheet
    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        'en-US': enUS
      },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({
          container: 'spreadsheet'
        })
      ]
    })

    // Create initial workbook
    univerAPI.createUniverSheet({
      name: 'TicknTie Workbook'
    })

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