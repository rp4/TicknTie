/**
 * TicknTie - Main Application Entry
 * Initializes Univer spreadsheet and image sidebar plugin
 */

import '@univerjs/preset-sheets-core/lib/index.css'
import { createUniver, defaultTheme, LocaleType } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/lib/locales/en-US.js'
import { ImagePlugin } from './image-plugin'
import { LandingPage } from './LandingPage.js'
import { inject } from '@vercel/analytics'

// Initialize spreadsheet application
async function initSpreadsheet() {
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

// Initialize landing page
function init() {
  // Initialize Vercel Analytics
  inject()

  const landing = new LandingPage({
    icon: '🕵️',
    title: 'Tick-n-Tie',
    subtitle: 'An excel like evidence organization tool with full privacy',
    showInfoButton: true,
    infoPopup: {
      title: 'Your Privacy Matters',
      icon: '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>',
      sections: [
        {
          icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0110 0v4"/></svg>',
          title: 'Local Processing Only',
          content: 'All processing happens in your browser. Your data and files never leave your device unless you explicitly export them.'
        },
        {
          icon: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>',
          title: 'No Data Collection',
          bullets: [
            'No tracking or analytics',
            'No cookies or local storage',
            'No server uploads',
            'Complete privacy'
          ]
        }
      ]
    },
    actions: [
      {
        label: 'Load Project',
        variant: 'secondary'
      },
      {
        label: 'Try Demo',
        onClick: async () => {
          landing.hide()
          await initSpreadsheet()

          // Load demo data
          setTimeout(async () => {
            const imagePlugin = window.ticknTie?.imagePlugin
            if (imagePlugin) {
              try {
                const response = await fetch('/demo_data.zip')
                const blob = await response.blob()
                const file = new File([blob], 'demo_data.zip', { type: 'application/zip' })
                await imagePlugin.importProject(file)
              } catch (error) {
                console.error('Failed to load demo data:', error)
              }
            }
          }, 500)
        },
        variant: 'primary'
      }
    ],
    fileUpload: {
      accept: '.xlsx,.xls,.csv,.zip',
      onFileSelect: async (file) => {
        // Hide landing page and initialize spreadsheet first
        landing.hide()
        await initSpreadsheet()

        // Wait a moment for the spreadsheet to fully initialize
        setTimeout(async () => {
          // Get the image plugin instance that was just created
          const imagePlugin = window.ticknTie?.imagePlugin

          if (imagePlugin) {
            const fileName = file.name.toLowerCase()

            if (fileName.endsWith('.zip')) {
              // Handle ZIP file with evidence
              await imagePlugin.importProject(file)
            } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
              // Handle Excel file
              const reader = new FileReader()
              reader.onload = async (e) => {
                await imagePlugin.importXLSXWithEvidence(e.target.result, new Map())
              }
              reader.readAsArrayBuffer(file)
            } else if (fileName.endsWith('.csv')) {
              // Handle CSV file
              const reader = new FileReader()
              reader.onload = async (e) => {
                await imagePlugin.importCSVWithEvidence(e.target.result, new Map())
              }
              reader.readAsText(file)
            }
          }
        }, 500)
      },
      dragDropEnabled: true
    },
    footerLinks: [
      {
        icon: '<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor"><path fill-rule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.531 1.032 1.531 1.032.892 1.53 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z" clip-rule="evenodd"/></svg>',
        href: 'https://github.com/rp4/TicknTie',
        title: 'GitHub Repository'
      },
      {
        icon: '🧰',
        href: 'https://audittoolbox.com',
        title: 'Find other audit tools here'
      }
    ],
    containerClassName: 'bg-white/90 backdrop-blur-md'
  })

  // Show the landing page
  landing.show()

  // Change background to use TicknTie.png after element is created
  if (landing.element) {
    landing.element.style.background = 'url(/TicknTie.png) center/cover'
  }
}

// Start application when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init)
} else {
  init()
}