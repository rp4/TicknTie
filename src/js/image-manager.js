/**
 * Image Manager Module
 * Handles image storage, linking, and retrieval using URL references
 */

class ImageManager {
    constructor() {
        this.imageStore = new Map(); // Map of imageId -> {url, file, metadata}
        this.cellImageRefs = new Map(); // Map of cellRef -> imageId
        this.imageIdCounter = 1;
    }

    /**
     * Generate a unique image ID
     */
    generateImageId(filename) {
        const ext = filename.split('.').pop().toLowerCase();
        const id = `img_${String(this.imageIdCounter).padStart(3, '0')}.${ext}`;
        this.imageIdCounter++;
        return id;
    }

    /**
     * Store an image file and return its ID
     */
    async storeImage(file) {
        const imageId = this.generateImageId(file.name);
        const url = URL.createObjectURL(file);
        
        // Get image dimensions
        const dimensions = await this.getImageDimensions(url);
        
        this.imageStore.set(imageId, {
            id: imageId,
            url: url,
            file: file,
            originalName: file.name,
            size: file.size,
            type: file.type,
            width: dimensions.width,
            height: dimensions.height,
            addedAt: new Date().toISOString()
        });
        
        return imageId;
    }

    /**
     * Store image from base64 data (for compatibility)
     */
    async storeImageFromBase64(base64Data, filename = 'image.png') {
        // Convert base64 to blob
        const response = await fetch(base64Data);
        const blob = await response.blob();
        const file = new File([blob], filename, { type: blob.type });
        
        return this.storeImage(file);
    }

    /**
     * Get image dimensions
     */
    getImageDimensions(url) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                resolve({ width: img.width, height: img.height });
            };
            img.onerror = () => {
                resolve({ width: 0, height: 0 });
            };
            img.src = url;
        });
    }

    /**
     * Link an image to a cell
     */
    linkImageToCell(cellRef, imageId) {
        if (!this.imageStore.has(imageId)) {
            throw new Error(`Image ${imageId} not found in store`);
        }
        
        // Remove any existing image from this cell
        if (this.cellImageRefs.has(cellRef)) {
            const oldImageId = this.cellImageRefs.get(cellRef);
            // Optionally clean up unused images
            this.checkAndCleanupImage(oldImageId);
        }
        
        this.cellImageRefs.set(cellRef, imageId);
        return imageId;
    }

    /**
     * Get image data by ID
     */
    getImageData(imageId) {
        return this.imageStore.get(imageId) || null;
    }

    /**
     * Get image ID for a cell
     */
    getImageForCell(cellRef) {
        const imageId = this.cellImageRefs.get(cellRef);
        return imageId ? this.getImageData(imageId) : null;
    }

    /**
     * Generate a thumbnail URL for an image
     */
    async generateThumbnail(imageId, maxSize = 50) {
        const imageData = this.getImageData(imageId);
        if (!imageData) return null;
        
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                
                // Calculate thumbnail dimensions
                let width = img.width;
                let height = img.height;
                
                if (width > height) {
                    if (width > maxSize) {
                        height = (maxSize / width) * height;
                        width = maxSize;
                    }
                } else {
                    if (height > maxSize) {
                        width = (maxSize / height) * width;
                        height = maxSize;
                    }
                }
                
                canvas.width = width;
                canvas.height = height;
                ctx.drawImage(img, 0, 0, width, height);
                
                canvas.toBlob((blob) => {
                    const thumbnailUrl = URL.createObjectURL(blob);
                    resolve(thumbnailUrl);
                });
            };
            img.onerror = () => resolve(null);
            img.src = imageData.url;
        });
    }

    /**
     * Remove image from cell
     */
    removeImageFromCell(cellRef) {
        const imageId = this.cellImageRefs.get(cellRef);
        if (imageId) {
            this.cellImageRefs.delete(cellRef);
            this.checkAndCleanupImage(imageId);
        }
    }

    /**
     * Check if an image is still in use and clean up if not
     */
    checkAndCleanupImage(imageId) {
        // Check if any cell still references this image
        const stillInUse = Array.from(this.cellImageRefs.values()).includes(imageId);
        
        if (!stillInUse) {
            const imageData = this.imageStore.get(imageId);
            if (imageData && imageData.url) {
                URL.revokeObjectURL(imageData.url);
            }
            this.imageStore.delete(imageId);
        }
    }

    /**
     * Get all images in the store
     */
    getAllImages() {
        return Array.from(this.imageStore.values());
    }

    /**
     * Get all cell-image mappings
     */
    getAllCellMappings() {
        const mappings = {};
        this.cellImageRefs.forEach((imageId, cellRef) => {
            mappings[cellRef] = imageId;
        });
        return mappings;
    }

    /**
     * Clear all images and mappings
     */
    clearAll() {
        // Revoke all object URLs to free memory
        this.imageStore.forEach((imageData) => {
            if (imageData.url) {
                URL.revokeObjectURL(imageData.url);
            }
        });
        
        this.imageStore.clear();
        this.cellImageRefs.clear();
        this.imageIdCounter = 1;
    }

    /**
     * Export images for zip download
     */
    async exportImages() {
        const images = {};
        
        for (const [imageId, imageData] of this.imageStore) {
            images[imageId] = {
                file: imageData.file,
                originalName: imageData.originalName
            };
        }
        
        return images;
    }

    /**
     * Import images from zip extraction
     */
    async importImages(imageFiles) {
        const mappings = {};
        
        for (const [filename, file] of Object.entries(imageFiles)) {
            const imageId = await this.storeImage(file);
            mappings[filename] = imageId;
        }
        
        return mappings;
    }

    /**
     * Get statistics about stored images
     */
    getStatistics() {
        const totalSize = Array.from(this.imageStore.values())
            .reduce((sum, img) => sum + img.size, 0);
        
        return {
            totalImages: this.imageStore.size,
            totalMappings: this.cellImageRefs.size,
            totalSizeBytes: totalSize,
            totalSizeMB: (totalSize / (1024 * 1024)).toFixed(2)
        };
    }
}

// Export as singleton
const imageManager = new ImageManager();

// Make it available globally for the app
window.ImageManager = imageManager;