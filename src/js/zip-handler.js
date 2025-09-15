/**
 * Zip Handler Module
 * Manages zip file operations for bundling Excel files with images
 */

class ZipHandler {
    constructor() {
        this.zip = null;
    }

    /**
     * Load a zip file containing Excel and images
     */
    async loadZipFile(file) {
        try {
            const zip = new JSZip();
            const contents = await zip.loadAsync(file);
            
            let excelFile = null;
            const imageFiles = {};
            const imageMappings = {};
            
            // Process each file in the zip
            for (const [path, zipEntry] of Object.entries(contents.files)) {
                const filename = path.split('/').pop();
                
                if (this.isExcelFile(filename)) {
                    // Found Excel file
                    excelFile = await zipEntry.async('arraybuffer');
                    excelFile = new File([excelFile], filename, {
                        type: this.getExcelMimeType(filename)
                    });
                } else if (this.isImageFile(filename)) {
                    // Found image file
                    const imageData = await zipEntry.async('blob');
                    const imageFile = new File([imageData], filename, {
                        type: this.getImageMimeType(filename)
                    });
                    
                    // Store in images folder structure
                    const imagePath = path.replace(/^images\//, '');
                    imageFiles[imagePath] = imageFile;
                }
            }
            
            // Check if we have a mappings.json file
            if (contents.files['mappings.json']) {
                const mappingsData = await contents.files['mappings.json'].async('string');
                const mappings = JSON.parse(mappingsData);
                Object.assign(imageMappings, mappings);
            }
            
            return {
                excelFile,
                imageFiles,
                imageMappings
            };
        } catch (error) {
            console.error('Error loading zip file:', error);
            throw new Error('Failed to load zip file: ' + error.message);
        }
    }

    /**
     * Create a zip bundle with Excel file and images
     */
    async createZipBundle(excelBlob, images, cellMappings) {
        const zip = new JSZip();
        
        // Add Excel file
        zip.file('workbook.xlsx', excelBlob);
        
        // Create images folder
        const imagesFolder = zip.folder('images');
        
        // Add all images
        for (const [imageId, imageData] of Object.entries(images)) {
            if (imageData.file) {
                imagesFolder.file(imageId, imageData.file);
            }
        }
        
        // Create mappings file for cell-image relationships
        const mappings = {
            version: '1.0',
            created: new Date().toISOString(),
            cellMappings: cellMappings,
            images: Object.keys(images).reduce((acc, imageId) => {
                acc[imageId] = {
                    originalName: images[imageId].originalName,
                    size: images[imageId].file.size,
                    type: images[imageId].file.type
                };
                return acc;
            }, {})
        };
        
        zip.file('mappings.json', JSON.stringify(mappings, null, 2));
        
        // Create README for the zip structure
        const readme = `TicknTie Audit Workpaper Bundle
================================

This zip file contains:
- workbook.xlsx: The Excel workbook with image references
- images/: Folder containing all referenced images
- mappings.json: Metadata about image-cell relationships

To view with full image support, import this zip file into TicknTie.

Generated: ${new Date().toLocaleString()}
`;
        
        zip.file('README.txt', readme);
        
        // Generate zip file
        const zipBlob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: {
                level: 6
            }
        });
        
        return zipBlob;
    }

    /**
     * Extract images from a zip file
     */
    async extractImages(zipFile) {
        const zip = new JSZip();
        const contents = await zip.loadAsync(zipFile);
        const images = {};
        
        for (const [path, zipEntry] of Object.entries(contents.files)) {
            if (zipEntry.dir) continue;
            
            const filename = path.split('/').pop();
            if (this.isImageFile(filename)) {
                const blob = await zipEntry.async('blob');
                const file = new File([blob], filename, {
                    type: this.getImageMimeType(filename)
                });
                images[filename] = file;
            }
        }
        
        return images;
    }

    /**
     * Check if file is an Excel file
     */
    isExcelFile(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        return ['xlsx', 'xls', 'xlsm', 'xlsb'].includes(ext);
    }

    /**
     * Check if file is an image file
     */
    isImageFile(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        return ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'webp'].includes(ext);
    }

    /**
     * Get Excel MIME type
     */
    getExcelMimeType(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        const mimeTypes = {
            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'xls': 'application/vnd.ms-excel',
            'xlsm': 'application/vnd.ms-excel.sheet.macroenabled.12',
            'xlsb': 'application/vnd.ms-excel.sheet.binary.macroenabled.12'
        };
        return mimeTypes[ext] || 'application/octet-stream';
    }

    /**
     * Get image MIME type
     */
    getImageMimeType(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        const mimeTypes = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'svg': 'image/svg+xml',
            'webp': 'image/webp'
        };
        return mimeTypes[ext] || 'image/png';
    }

    /**
     * Download zip file
     */
    downloadZip(zipBlob, filename = 'tickntie-export.zip') {
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    /**
     * Generate export filename with timestamp
     */
    generateFilename(prefix = 'tickntie') {
        const now = new Date();
        const timestamp = now.toISOString().replace(/[:.]/g, '-').slice(0, -5);
        return `${prefix}_${timestamp}.zip`;
    }

    /**
     * Validate zip structure
     */
    async validateZipStructure(file) {
        try {
            const zip = new JSZip();
            const contents = await zip.loadAsync(file);
            
            let hasExcel = false;
            let hasImages = false;
            
            for (const path of Object.keys(contents.files)) {
                const filename = path.split('/').pop();
                if (this.isExcelFile(filename)) hasExcel = true;
                if (this.isImageFile(filename)) hasImages = true;
            }
            
            return {
                valid: hasExcel,
                hasExcel,
                hasImages,
                fileCount: Object.keys(contents.files).length
            };
        } catch (error) {
            return {
                valid: false,
                error: error.message
            };
        }
    }
}

// Export as singleton
const zipHandler = new ZipHandler();

// Make it available globally
window.ZipHandler = zipHandler;