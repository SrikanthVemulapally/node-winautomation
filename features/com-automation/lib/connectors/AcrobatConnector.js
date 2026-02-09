/**
 * @fileoverview Adobe Acrobat Connector
 * Provides basic PDF automation via Acrobat COM API
 * Note: Acrobat's COM API is limited. For advanced PDF operations,
 * consider using pdf-lib or pdf-parse npm packages.
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * @class AcrobatConnector
 * @description Basic API for Adobe Acrobat automation
 */
class AcrobatConnector {
    constructor() {
        try {
            this.app = new COMObject('AcroExch.App');
        } catch (e) {
            throw new Error('Adobe Acrobat not installed or COM interface not available: ' + e.message);
        }
    }

    /**
     * Show Acrobat window
     */
    show() {
        this.app.invoke('Show');
    }

    /**
     * Hide Acrobat window
     */
    hide() {
        this.app.invoke('Hide');
    }

    /**
     * Open PDF document
     * @param {string} path - Full path to PDF file
     * @returns {PDFDocument} PDF document wrapper
     */
    openDocument(path) {
        const avDoc = new COMObject('AcroExch.AVDoc');
        const result = avDoc.invoke('Open', path, '');
        if (!result) {
            throw new Error('Failed to open PDF: ' + path);
        }
        return new PDFDocument(avDoc);
    }

    /**
     * Exit Acrobat application
     */
    exit() {
        this.app.invoke('Exit');
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

/**
 * @class PDFDocument
 * @description Wrapper for Acrobat PDF Document
 */
class PDFDocument {
    constructor(avDoc) {
        this.avDoc = avDoc;
        this.pdDoc = null;
    }

    /**
     * Get PDDoc object (internal PDF document)
     * @returns {COMObject} PDDoc object
     */
    getPDDoc() {
        if (!this.pdDoc) {
            this.pdDoc = this.avDoc.invoke('GetPDDoc');
        }
        return this.pdDoc;
    }

    /**
     * Get page count
     * @returns {number} Number of pages
     */
    getPageCount() {
        const pdDoc = this.getPDDoc();
        return pdDoc.invoke('GetNumPages');
    }

    /**
     * Get file name
     * @returns {string} File path
     */
    getFileName() {
        const pdDoc = this.getPDDoc();
        return pdDoc.invoke('GetFileName');
    }

    /**
     * Save PDF
     * @param {string} path - Path to save (optional, saves to current location if not specified)
     */
    save(path = null) {
        const pdDoc = this.getPDDoc();
        if (path) {
            pdDoc.invoke('Save', 1, path);
        } else {
            pdDoc.invoke('Save', 0, '');
        }
    }

    /**
     * Close document
     * @param {boolean} noSave - True to close without saving
     */
    close(noSave = true) {
        this.avDoc.invoke('Close', noSave ? 1 : 0);
    }

    /**
     * Release COM objects
     */
    release() {
        if (this.pdDoc) {
            this.pdDoc.release();
        }
        this.avDoc.release();
    }
}

module.exports = {
    AcrobatConnector,
    PDFDocument,
};
