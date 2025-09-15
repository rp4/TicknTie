export class LandingPage {
  constructor(config = {}) {
    // Support both direct config and callback-based initialization
    if (typeof config === 'function' || (arguments.length === 2 && typeof arguments[0] === 'function')) {
      // Legacy SwimLaneLandingPage compatibility mode
      const onStart = arguments[0];
      const onFileLoad = arguments[1];
      this.config = this.getLegacyConfig(onStart, onFileLoad);
    } else {
      this.config = {
        icon: '🏊',
        title: 'Swim Lane Diagram',
        subtitle: 'Pool Planner',
        showInfoButton: false,
        infoPopup: null,
        actions: [],
        fileUpload: null,
        footerLinks: [],
        className: '',
        containerClassName: '',
        ...config
      };
    }

    this.isDragging = false;
    this.showPrivacyInfo = false;
    this.element = null;
    this.container = null;
  }

  // Legacy compatibility method to create config from callbacks
  getLegacyConfig(onStart, onFileLoad) {
    return {
      icon: '🏊',
      title: 'SwimLanes',
      subtitle: 'Interactive Process Diagram Visualizer',
      showInfoButton: true,
      infoPopup: {
        title: 'Your Privacy Matters',
        icon: this.createShieldIcon(),
        sections: [
          {
            icon: this.createLockIcon(),
            title: 'Local Processing Only',
            content: 'All processing happens in your browser. Your diagrams never leave your device unless you explicitly export them.'
          },
          {
            icon: this.createDatabaseIcon(),
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
          label: 'Start New Diagram',
          onClick: onStart,
          variant: 'primary',
          icon: this.createFileTextIcon()
        },
        {
          label: 'Load from File',
          variant: 'secondary',
          icon: this.createUploadIcon()
        }
      ],
      fileUpload: {
        accept: '.json',
        onFileSelect: (file) => {
          const reader = new FileReader();
          reader.onload = (e) => {
            try {
              const jsonData = JSON.parse(e.target.result);
              if (onFileLoad) {
                onFileLoad(jsonData);
              }
            } catch (error) {
              alert('Invalid JSON file. Please select a valid swim lane diagram file.');
            }
          };
          reader.readAsText(file);
        },
        dragDropEnabled: true
      },
      footerLinks: [
        {
          icon: this.createGithubIcon(),
          href: 'https://github.com/rp4/SwimLanes',
          title: 'GitHub Repository'
        },
        {
          icon: this.createChatGPTIcon(),
          href: 'https://chatgpt.com',
          title: 'Run the custom GPT to create your inputs here'
        },
        {
          icon: '🏆',
          href: 'https://scoreboard.audittoolbox.com',
          title: 'See the prompt to create your inputs here'
        },
        {
          icon: '🧰',
          href: 'https://audittoolbox.com',
          title: 'Find other audit tools here'
        }
      ],
      containerClassName: 'bg-white/90 backdrop-blur-md'
    };
  }

  // Icon creation methods for legacy compatibility
  createShieldIcon() {
    return `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
    </svg>`;
  }

  createLockIcon() {
    return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2">
      <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
      <path d="M7 11V7a5 5 0 0110 0v4"/>
    </svg>`;
  }

  createDatabaseIcon() {
    return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6b7280" stroke-width="2">
      <ellipse cx="12" cy="5" rx="9" ry="3"/>
      <path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/>
      <path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>
    </svg>`;
  }

  createFileTextIcon() {
    return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
      <polyline points="14 2 14 8 20 8"/>
      <line x1="16" y1="13" x2="8" y2="13"/>
      <line x1="16" y1="17" x2="8" y2="17"/>
      <polyline points="10 9 9 9 8 9"/>
    </svg>`;
  }

  createUploadIcon() {
    return `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/>
      <polyline points="17 8 12 3 7 8"/>
      <line x1="12" y1="3" x2="12" y2="15"/>
    </svg>`;
  }

  createGithubIcon() {
    return `<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
      <path fill-rule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.531 1.032 1.531 1.032.892 1.53 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z" clip-rule="evenodd"/>
    </svg>`;
  }

  createChatGPTIcon() {
    return `<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">
      <path d="M22.282 9.821a5.985 5.985 0 0 0-.516-4.91 6.046 6.046 0 0 0-6.51-2.9A6.065 6.065 0 0 0 4.981 4.18a5.985 5.985 0 0 0-3.998 2.9 6.046 6.046 0 0 0 .743 7.097 5.98 5.98 0 0 0 .51 4.911 6.051 6.051 0 0 0 6.515 2.9A5.985 5.985 0 0 0 13.26 24a6.056 6.056 0 0 0 5.772-4.206 5.99 5.99 0 0 0 3.997-2.9 6.056 6.056 0 0 0-.747-7.073zM13.26 22.43a4.476 4.476 0 0 1-2.876-1.04l.141-.081 4.779-2.758a.795.795 0 0 0 .392-.681v-6.737l2.02 1.168a.071.071 0 0 1 .038.052v5.583a4.504 4.504 0 0 1-4.494 4.494zM3.6 18.304a4.47 4.47 0 0 1-.535-3.014l.142.085 4.783 2.759a.771.771 0 0 0 .78 0l5.843-3.369v2.332a.08.08 0 0 1-.033.062L9.74 19.95a4.5 4.5 0 0 1-6.14-1.646zM2.34 7.896a4.485 4.485 0 0 1 2.366-1.973V11.6a.766.766 0 0 0 .388.676l5.815 3.355-2.02 1.168a.076.076 0 0 1-.071 0l-4.83-2.786A4.504 4.504 0 0 1 2.34 7.872zm16.597 3.855l-5.833-3.387L15.119 7.2a.076.076 0 0 1 .071 0l4.83 2.791a4.494 4.494 0 0 1-.676 8.105v-5.678a.79.79 0 0 0-.407-.667zm2.01-3.023l-.141-.085-4.774-2.782a.776.776 0 0 0-.785 0L9.409 9.23V6.897a.066.066 0 0 1 .028-.061l4.83-2.787a4.5 4.5 0 0 1 6.68 4.66zm-12.64 4.135l-2.02-1.164a.08.08 0 0 1-.038-.057V6.075a4.5 4.5 0 0 1 7.375-3.453l-.142.08L8.704 5.46a.795.795 0 0 0-.393.681v6.737zm1.097-2.365l2.602-1.5 2.607 1.5v2.999l-2.597 1.5-2.607-1.5z"/>
    </svg>`;
  }

  render() {
    const container = document.createElement('div');
    container.id = 'landing-page-container';
    container.className = `landing-page-wrapper ${this.config.className}`;
    container.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100vh;
      background: url('/pool.png') center/cover;
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
      backdrop-filter: blur(2px);
    `;

    const innerContainer = document.createElement('div');
    innerContainer.className = `landing-container ${this.config.containerClassName}`;
    innerContainer.style.cssText = `
      width: 100%;
      max-width: 42rem;
      padding: 3rem;
      border: 2px dashed;
      border-radius: 1rem;
      transition: all 0.3s ease;
      background: rgba(255, 255, 255, 0.95);
      backdrop-filter: blur(10px);
      box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
      ${this.isDragging
        ? 'border-color: #5b21b6; background: rgba(139, 92, 246, 0.1);'
        : 'border-color: #d1d5db; background: rgba(255, 255, 255, 0.95);'}
    `;

    innerContainer.innerHTML = `
      <div style="text-align: center;">
        <!-- Title with icon -->
        <h2 style="font-size: 3rem; font-weight: bold; color: #1f2937; margin-bottom: 1rem; display: flex; align-items: center; justify-content: center; gap: 0.75rem;">
          <span style="font-size: 3.75rem;">${this.config.icon}</span>
          ${this.config.title}
        </h2>

        <!-- Subtitle with optional info button -->
        <div style="display: flex; align-items: center; justify-content: center; gap: 0.5rem; margin-bottom: 1.5rem; position: relative;">
          <p style="font-size: 1.25rem; color: #4b5563;">${this.config.subtitle}</p>
          ${this.config.showInfoButton ? `
            <button class="info-button" style="
              background: none;
              border: none;
              cursor: help;
              padding: 0;
              display: flex;
              align-items: center;
              justify-content: center;
            ">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color: #9ca3af; transition: color 0.2s;">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="12" y1="16" x2="12" y2="12"></line>
                <line x1="12" y1="8" x2="12.01" y2="8"></line>
              </svg>
            </button>
            <div class="info-popup" style="display: none; position: absolute;">
              <!-- Info popup content would go here -->
            </div>
          ` : ''}
        </div>

        <!-- File input (hidden) -->
        ${this.config.fileUpload ? `
          <input type="file"
            accept="${this.config.fileUpload.accept || '.json'}"
            style="display: none;"
            id="landing-file-input"
          />
        ` : ''}

        <!-- Action buttons -->
        ${this.config.actions.length > 0 ? `
          <div style="display: flex; align-items: center; justify-content: center; gap: 1rem; flex-wrap: wrap;">
            ${this.config.actions.map((action, idx) => {
              const isFileUploadButton = (action.label === 'Choose File' || action.label === 'Load from File') && this.config.fileUpload;
              const buttonStyles = this.getButtonStyles(action.variant);

              if (isFileUploadButton) {
                return `
                  <label for="landing-file-input"
                    style="
                      display: inline-flex;
                      align-items: center;
                      justify-content: center;
                      width: 13rem;
                      height: 5rem;
                      padding: 0 1.5rem;
                      font-size: 1.125rem;
                      font-weight: 500;
                      border-radius: 0.5rem;
                      cursor: pointer;
                      transition: all 0.3s;
                      ${buttonStyles}
                      ${action.className || ''}
                    "
                    title="${action.tooltip || ''}"
                  >
                    ${action.icon || ''}
                    ${action.label}
                  </label>
                `;
              }

              return `
                <button
                  data-action-index="${idx}"
                  style="
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    width: 13rem;
                    height: 5rem;
                    padding: 0 1.5rem;
                    font-size: 1.125rem;
                    font-weight: 500;
                    border-radius: 0.5rem;
                    transition: all 0.3s;
                    cursor: pointer;
                    ${buttonStyles}
                    ${action.className || ''}
                  "
                  title="${action.tooltip || ''}"
                >
                  ${action.icon || ''}
                  ${action.label}
                </button>
              `;
            }).join('')}
          </div>
        ` : ''}

        <!-- Footer links -->
        ${this.config.footerLinks.length > 0 ? `
          <div style="display: flex; align-items: center; justify-content: center; gap: 1.5rem; margin-top: 2rem;">
            ${this.config.footerLinks.map(link => {
              // Check if icon is an emoji (not SVG)
              const isEmoji = !link.icon.includes('<svg');
              // For SVGs, wrap in a larger container
              const iconContent = link.icon.includes('<svg')
                ? link.icon.replace('width="24" height="24"', 'width="32" height="32"')
                : link.icon;
              return `
                <a href="${link.href}"
                  target="${link.target || '_blank'}"
                  rel="noopener noreferrer"
                  style="
                    color: #4b5563;
                    transition: all 0.3s;
                    display: inline-block;
                    text-decoration: none;
                    ${isEmoji ? 'font-size: 32px; line-height: 1;' : ''}
                  "
                  title="${link.title || ''}"
                  onmouseover="this.style.color='#1f2937'; this.style.transform='scale(1.1)'"
                  onmouseout="this.style.color='#4b5563'; this.style.transform='scale(1)'"
                >
                  ${iconContent}
                </a>
              `;
            }).join('')}
          </div>
        ` : ''}
      </div>
    `;

    container.appendChild(innerContainer);

    // Setup event listeners
    this.setupEventListeners(container, innerContainer);

    this.element = container;
    this.container = container;
    return container;
  }

  setupEventListeners(container, innerContainer) {
    // Handle drag and drop if enabled
    if (this.config.fileUpload?.dragDropEnabled) {
      innerContainer.addEventListener('dragover', (e) => {
        e.preventDefault();
        this.isDragging = true;
        innerContainer.style.borderColor = '#5b21b6';
        innerContainer.style.background = 'rgba(139, 92, 246, 0.1)';
      });

      innerContainer.addEventListener('dragleave', (e) => {
        e.preventDefault();
        this.isDragging = false;
        innerContainer.style.borderColor = '#d1d5db';
        innerContainer.style.background = 'rgba(255, 255, 255, 0.95)';
      });

      innerContainer.addEventListener('drop', (e) => {
        e.preventDefault();
        this.isDragging = false;
        innerContainer.style.borderColor = '#d1d5db';
        innerContainer.style.background = 'rgba(255, 255, 255, 0.95)';

        const files = Array.from(e.dataTransfer.files);
        if (files[0]) {
          this.config.fileUpload.onFileSelect(files[0]);
        }
      });
    }

    // Handle file input change
    const fileInput = container.querySelector('#landing-file-input');
    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        if (e.target.files?.[0]) {
          this.config.fileUpload.onFileSelect(e.target.files[0]);
        }
      });
    }

    // Handle action button clicks
    container.querySelectorAll('button[data-action-index]').forEach(button => {
      button.addEventListener('click', () => {
        const index = parseInt(button.dataset.actionIndex);
        const action = this.config.actions[index];
        if (action?.onClick) {
          action.onClick();
        }
      });
    });

    // Handle info button hover
    const infoButton = container.querySelector('.info-button');
    const infoPopup = container.querySelector('.info-popup');
    if (infoButton && infoPopup && this.config.infoPopup) {
      infoButton.addEventListener('mouseenter', () => {
        this.showPrivacyInfo = true;
        this.renderInfoPopup(infoPopup);
      });

      infoButton.addEventListener('mouseleave', () => {
        this.showPrivacyInfo = false;
        infoPopup.style.display = 'none';
      });

      // Keep popup open when hovering over it
      infoPopup.addEventListener('mouseenter', () => {
        this.showPrivacyInfo = true;
        infoPopup.style.display = 'block';
      });

      infoPopup.addEventListener('mouseleave', () => {
        this.showPrivacyInfo = false;
        infoPopup.style.display = 'none';
      });
    }
  }

  renderInfoPopup(popupElement) {
    if (!this.config.infoPopup) return;

    // Position the popup to the side with fixed positioning to stay in viewport
    popupElement.style.cssText = `
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      max-width: min(90vw, 500px);
      width: 500px;
      max-height: 80vh;
      overflow-y: auto;
      background: #f9fafb;
      border-radius: 0.75rem;
      box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04);
      border: 1px solid #e5e7eb;
      padding: 2rem;
      z-index: 10000;
      display: block;
    `;

    // Render popup content with icons
    popupElement.innerHTML = `
      <!-- Header with shield icon -->
      <div style="display: flex; align-items: center; gap: 1rem; margin-bottom: 2rem;">
        ${this.config.infoPopup.icon ? `
          <div style="width: 3rem; height: 3rem; background: #dbeafe; border-radius: 0.5rem; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
            ${this.config.infoPopup.icon}
          </div>
        ` : ''}
        <h3 style="font-size: 1.5rem; font-weight: 600; color: #111827; margin: 0;">
          ${this.config.infoPopup.title}
        </h3>
      </div>

      <!-- Sections -->
      ${this.config.infoPopup.sections.map((section, idx) => `
        <div style="display: flex; align-items: flex-start; gap: 1rem; margin-bottom: ${idx < this.config.infoPopup.sections.length - 1 ? '1.5rem' : '0'};">
          ${section.icon ? `
            <div style="width: 2.5rem; height: 2.5rem; background: #e5e7eb; border-radius: 0.5rem; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 0.25rem;">
              ${section.icon}
            </div>
          ` : ''}
          <div style="flex: 1;">
            <h4 style="font-size: 1.125rem; font-weight: 600; color: #111827; margin: 0 0 0.5rem 0;">
              ${section.title}
            </h4>
            ${section.content ? `
              <p style="color: #6b7280; margin: 0 0 0.5rem 0; line-height: 1.5;">
                ${section.content}
              </p>
            ` : ''}
            ${section.bullets ? `
              <ul style="list-style: none; padding: 0; margin: 0;">
                ${section.bullets.map(bullet => `
                  <li style="color: #6b7280; margin-bottom: 0.5rem; padding-left: 1.25rem; position: relative;">
                    <span style="position: absolute; left: 0;">•</span>
                    ${bullet}
                  </li>
                `).join('')}
              </ul>
            ` : ''}
          </div>
        </div>
      `).join('')}
    `;
  }

  getButtonStyles(variant = 'primary') {
    switch (variant) {
      case 'secondary':
        return 'border: 2px solid #6366f1; color: #4f46e5; background: white;';
      case 'success':
        return 'background: #10b981; color: white; border: 2px solid #10b981;';
      case 'primary':
      default:
        return 'background: #6366f1; color: white; border: 2px solid #6366f1;';
    }
  }

  show() {
    if (!this.element) {
      this.element = this.render();
      document.body.appendChild(this.element);
    } else if (this.element.style.display === 'none') {
      this.element.style.display = 'flex';
    }
  }

  hide() {
    if (this.element) {
      this.element.style.display = 'none';
    }
    this.remove();
  }

  remove() {
    if (this.element && this.element.parentNode) {
      this.element.parentNode.removeChild(this.element);
    }
    this.element = null;
    this.container = null;
  }
}