export class LandingPage {
  constructor(options = {}) {
    this.options = options
    this.element = null
    this.createLandingPage()
  }

  createLandingPage() {
    // Create container
    const container = document.createElement('div')
    container.className = `landing-page ${this.options.containerClassName || ''}`
    container.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 10000;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    `

    // Create content wrapper
    const content = document.createElement('div')
    content.className = 'landing-content'
    content.style.cssText = `
      background: rgba(255, 255, 255, 0.75);
      backdrop-filter: blur(12px);
      -webkit-backdrop-filter: blur(12px);
      border: 1px solid rgba(255, 255, 255, 0.3);
      border-radius: 16px;
      padding: 48px;
      max-width: 600px;
      width: 90%;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      text-align: center;
    `

    // Add icon and title
    if (this.options.icon || this.options.title) {
      const header = document.createElement('div')
      header.style.marginBottom = '24px'

      if (this.options.title) {
        const titleContainer = document.createElement('div')
        titleContainer.style.cssText = 'display: flex; align-items: center; justify-content: center; gap: 12px;'

        if (this.options.icon) {
          const icon = document.createElement('span')
          icon.style.cssText = 'font-size: 56px;'
          icon.textContent = this.options.icon
          titleContainer.appendChild(icon)
        }

        const title = document.createElement('h1')
        title.style.cssText = 'margin: 0; font-size: 48px; color: #1f2937; font-weight: 600;'
        title.textContent = this.options.title
        titleContainer.appendChild(title)

        header.appendChild(titleContainer)
      }

      if (this.options.subtitle) {
        const subtitleContainer = document.createElement('div')
        subtitleContainer.style.cssText = 'display: flex; align-items: center; justify-content: center; gap: 8px; margin-top: 12px;'

        const subtitle = document.createElement('p')
        subtitle.style.cssText = 'margin: 0; font-size: 16px; color: #6b7280;'

        // Split subtitle to add info button next to "privacy" word
        if (this.options.showInfoButton && this.options.infoPopup) {
          const subtitleParts = this.options.subtitle.split('privacy')
          if (subtitleParts.length > 1) {
            subtitle.textContent = subtitleParts[0] + 'privacy'
            subtitleContainer.appendChild(subtitle)

            // Add info button inline
            const infoButton = document.createElement('button')
            infoButton.innerHTML = `
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <circle cx="12" cy="12" r="10"/>
                <path d="M12 16v-4"/>
                <path d="M12 8h.01"/>
              </svg>
            `
            infoButton.style.cssText = `
              background: none;
              border: none;
              cursor: pointer;
              color: #6b7280;
              padding: 4px;
              border-radius: 50%;
              transition: background 0.3s;
              display: inline-flex;
              align-items: center;
            `

            let popupElement = null
            let hoverTimeout = null
            let leaveTimeout = null

            const showPopup = () => {
              if (!popupElement) {
                popupElement = this.showInfoPopup(true, infoButton)

                // Allow mouse to move to popup
                if (popupElement) {
                  const popupContent = popupElement.querySelector('.info-popup-content')
                  if (popupContent) {
                    popupContent.onmouseenter = () => {
                      if (leaveTimeout) {
                        clearTimeout(leaveTimeout)
                        leaveTimeout = null
                      }
                    }
                    popupContent.onmouseleave = () => {
                      hidePopup()
                    }
                  }
                }
              }
            }

            const hidePopup = () => {
              leaveTimeout = setTimeout(() => {
                if (popupElement) {
                  popupElement.remove()
                  popupElement = null
                }
              }, 100)
            }

            infoButton.onmouseenter = () => {
              infoButton.style.background = '#f3f4f6'
              if (leaveTimeout) {
                clearTimeout(leaveTimeout)
                leaveTimeout = null
              }
              // Small delay to prevent accidental triggers
              hoverTimeout = setTimeout(showPopup, 200)
            }

            infoButton.onmouseleave = () => {
              infoButton.style.background = 'none'
              if (hoverTimeout) {
                clearTimeout(hoverTimeout)
                hoverTimeout = null
              }
              hidePopup()
            }

            subtitleContainer.appendChild(infoButton)

            // Add the rest of the subtitle if any
            if (subtitleParts[1]) {
              const restText = document.createElement('span')
              restText.style.cssText = 'margin: 0; font-size: 16px; color: #6b7280;'
              restText.textContent = subtitleParts[1]
              subtitleContainer.appendChild(restText)
            }
          } else {
            subtitle.textContent = this.options.subtitle
            subtitleContainer.appendChild(subtitle)
          }
        } else {
          subtitle.textContent = this.options.subtitle
          subtitleContainer.appendChild(subtitle)
        }

        header.appendChild(subtitleContainer)
      }

      content.appendChild(header)
    }

    // Add actions
    if (this.options.actions && this.options.actions.length > 0) {
      const actionsContainer = document.createElement('div')
      actionsContainer.style.cssText = 'margin-top: 32px; display: flex; gap: 16px; justify-content: center; flex-wrap: wrap;'

      this.options.actions.forEach(action => {
        const button = document.createElement('button')
        button.textContent = action.label
        button.style.cssText = `
          padding: 12px 24px;
          border-radius: 8px;
          font-size: 16px;
          cursor: pointer;
          transition: all 0.3s;
          border: 2px solid transparent;
          ${action.variant === 'primary' ?
            'background: #3b82f6; color: white; border-color: #3b82f6;' :
            'background: rgba(255, 255, 255, 0.8); color: #3b82f6; border-color: #3b82f6;'}
        `

        button.onmouseover = () => {
          button.style.transform = 'translateY(-2px)'
          button.style.boxShadow = '0 4px 12px rgba(0,0,0,0.15)'
        }
        button.onmouseout = () => {
          button.style.transform = 'translateY(0)'
          button.style.boxShadow = 'none'
        }

        if (action.onClick) {
          button.onclick = action.onClick
        }

        // Handle file upload for Load Project button
        if (action.label === 'Load Project' && this.options.fileUpload) {
          button.onclick = () => this.triggerFileUpload()
        }

        actionsContainer.appendChild(button)
      })

      content.appendChild(actionsContainer)
    }

    // Add file upload input (hidden)
    if (this.options.fileUpload) {
      this.fileInput = document.createElement('input')
      this.fileInput.type = 'file'
      this.fileInput.accept = this.options.fileUpload.accept || '*'
      this.fileInput.style.display = 'none'
      this.fileInput.onchange = (e) => {
        const file = e.target.files[0]
        if (file && this.options.fileUpload.onFileSelect) {
          this.options.fileUpload.onFileSelect(file)
        }
      }
      content.appendChild(this.fileInput)

      // Add drag and drop support
      if (this.options.fileUpload.dragDropEnabled) {
        this.addDragDropSupport(content)
      }
    }

    // Add footer links
    if (this.options.footerLinks && this.options.footerLinks.length > 0) {
      const footer = document.createElement('div')
      footer.style.cssText = 'margin-top: 48px; display: flex; gap: 24px; justify-content: center; flex-wrap: wrap;'

      this.options.footerLinks.forEach(link => {
        const a = document.createElement('a')
        a.href = link.href
        a.target = '_blank'
        a.rel = 'noopener noreferrer'
        a.title = link.title || ''
        a.style.cssText = 'color: #6b7280; transition: color 0.3s; text-decoration: none; display: flex; align-items: center; gap: 8px; font-size: 32px;'

        if (link.icon) {
          if (link.icon.startsWith('<svg')) {
            const iconContainer = document.createElement('div')
            iconContainer.style.cssText = 'width: 32px; height: 32px;'
            iconContainer.innerHTML = link.icon.replace(/width="24"/, 'width="32"').replace(/height="24"/, 'height="32"')
            a.appendChild(iconContainer)
          } else {
            const iconSpan = document.createElement('span')
            iconSpan.textContent = link.icon
            a.appendChild(iconSpan)
          }
        }

        a.onmouseover = () => a.style.color = '#3b82f6'
        a.onmouseout = () => a.style.color = '#6b7280'

        footer.appendChild(a)
      })

      content.appendChild(footer)
    }

    container.appendChild(content)
    this.element = container
  }


  showInfoPopup(isHover = false) {
    const popup = document.createElement('div')
    popup.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100vh;
      background: ${isHover ? 'rgba(0,0,0,0.3)' : 'rgba(0,0,0,0.5)'};
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 10001;
      ${isHover ? 'pointer-events: none;' : ''}
    `

    const popupContent = document.createElement('div')
    popupContent.className = 'info-popup-content'
    popupContent.style.cssText = `
      background: white;
      border-radius: 12px;
      padding: 32px;
      max-width: 500px;
      width: 90%;
      max-height: 80vh;
      overflow-y: auto;
      ${isHover ? 'pointer-events: auto;' : ''}
    `

    // Add popup header
    const header = document.createElement('div')
    header.style.cssText = 'display: flex; align-items: center; gap: 12px; margin-bottom: 24px;'

    if (this.options.infoPopup.icon) {
      const iconContainer = document.createElement('div')
      iconContainer.innerHTML = this.options.infoPopup.icon
      header.appendChild(iconContainer)
    }

    const title = document.createElement('h2')
    title.style.cssText = 'margin: 0; color: #1f2937;'
    title.textContent = this.options.infoPopup.title || 'Information'
    header.appendChild(title)

    popupContent.appendChild(header)

    // Add sections
    if (this.options.infoPopup.sections) {
      this.options.infoPopup.sections.forEach(section => {
        const sectionDiv = document.createElement('div')
        sectionDiv.style.cssText = 'margin-bottom: 24px;'

        if (section.icon || section.title) {
          const sectionHeader = document.createElement('div')
          sectionHeader.style.cssText = 'display: flex; align-items: center; gap: 8px; margin-bottom: 12px;'

          if (section.icon) {
            const iconDiv = document.createElement('div')
            iconDiv.innerHTML = section.icon
            sectionHeader.appendChild(iconDiv)
          }

          if (section.title) {
            const sectionTitle = document.createElement('h3')
            sectionTitle.style.cssText = 'margin: 0; font-size: 18px; color: #374151;'
            sectionTitle.textContent = section.title
            sectionHeader.appendChild(sectionTitle)
          }

          sectionDiv.appendChild(sectionHeader)
        }

        if (section.content) {
          const contentP = document.createElement('p')
          contentP.style.cssText = 'margin: 0; color: #6b7280; line-height: 1.6;'
          contentP.textContent = section.content
          sectionDiv.appendChild(contentP)
        }

        if (section.bullets) {
          const ul = document.createElement('ul')
          ul.style.cssText = 'margin: 8px 0 0; padding-left: 24px; color: #6b7280;'
          section.bullets.forEach(bullet => {
            const li = document.createElement('li')
            li.textContent = bullet
            li.style.marginBottom = '4px'
            ul.appendChild(li)
          })
          sectionDiv.appendChild(ul)
        }

        popupContent.appendChild(sectionDiv)
      })
    }

    // Add close button only if not in hover mode
    if (!isHover) {
      const closeButton = document.createElement('button')
      closeButton.textContent = 'Got it'
      closeButton.style.cssText = `
        width: 100%;
        padding: 12px;
        background: #3b82f6;
        color: white;
        border: none;
        border-radius: 8px;
        font-size: 16px;
        cursor: pointer;
        margin-top: 24px;
      `
      closeButton.onclick = () => popup.remove()
      popupContent.appendChild(closeButton)
    }

    popup.appendChild(popupContent)

    if (!isHover) {
      popup.onclick = (e) => {
        if (e.target === popup) popup.remove()
      }
    }

    document.body.appendChild(popup)
    return popup
  }

  addDragDropSupport(content) {
    // Add drag and drop support to the entire content area
    content.ondragover = (e) => {
      e.preventDefault()
      content.style.boxShadow = '0 20px 60px rgba(59, 130, 246, 0.5)'
    }

    content.ondragleave = () => {
      content.style.boxShadow = '0 20px 60px rgba(0,0,0,0.3)'
    }

    content.ondrop = (e) => {
      e.preventDefault()
      content.style.boxShadow = '0 20px 60px rgba(0,0,0,0.3)'

      const file = e.dataTransfer.files[0]
      if (file && this.options.fileUpload.onFileSelect) {
        this.options.fileUpload.onFileSelect(file)
      }
    }
  }

  triggerFileUpload() {
    if (this.fileInput) {
      this.fileInput.click()
    }
  }

  show() {
    if (this.element && !this.element.parentNode) {
      document.body.appendChild(this.element)
    }
  }

  hide() {
    if (this.element && this.element.parentNode) {
      this.element.remove()
    }
  }
}