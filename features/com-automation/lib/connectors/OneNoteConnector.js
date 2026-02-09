/**
 * @fileoverview Microsoft OneNote Connector
 * Provides API for OneNote automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * @class OneNoteConnector
 * @description API for Microsoft OneNote automation
 */
class OneNoteConnector {
    constructor() {
        this.app = new COMObject('OneNote.Application');
    }

    /**
     * Get OneNote version
     * @returns {string} OneNote version
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Get notebooks XML
     * @returns {string} XML string containing notebook hierarchy
     */
    getNotebooksXML() {
        let xml = '';
        this.app.invoke('GetHierarchy', '', 2, xml);
        return xml;
    }

    /**
     * Create new page
     * @param {string} sectionId - Section ID where page will be created
     * @returns {string} New page ID
     */
    createNewPage(sectionId) {
        let pageId = '';
        this.app.invoke('CreateNewPage', sectionId, pageId);
        return pageId;
    }

    /**
     * Get page content
     * @param {string} pageId - Page ID
     * @returns {string} Page content as XML
     */
    getPageContent(pageId) {
        let xml = '';
        this.app.invoke('GetPageContent', pageId, xml);
        return xml;
    }

    /**
     * Update page content
     * @param {string} xml - Page content as XML
     */
    updatePageContent(xml) {
        this.app.invoke('UpdatePageContent', xml);
    }

    /**
     * Delete page
     * @param {string} pageId - Page ID to delete
     */
    deletePage(pageId) {
        this.app.invoke('DeletePageContent', pageId);
    }

    /**
     * Navigate to page
     * @param {string} pageId - Page ID to navigate to
     */
    navigateToPage(pageId) {
        this.app.invoke('NavigateTo', pageId);
    }

    /**
     * Publish page
     * @param {string} pageId - Page ID to publish
     * @param {string} targetPath - Target file path
     * @param {number} format - Publish format (0=OneNote, 1=PDF, 2=XPS, 3=Word, 4=EMF, 5=HTML)
     */
    publish(pageId, targetPath, format) {
        this.app.invoke('Publish', pageId, targetPath, format);
    }

    /**
     * Open hierarchy
     * @param {string} path - Path to notebook
     * @returns {string} Notebook ID
     */
    openHierarchy(path) {
        let notebookId = '';
        this.app.invoke('OpenHierarchy', path, '', notebookId);
        return notebookId;
    }

    /**
     * Close notebook
     * @param {string} notebookId - Notebook ID to close
     */
    closeNotebook(notebookId) {
        this.app.invoke('CloseNotebook', notebookId);
    }

    /**
     * Insert image into page
     * @param {string} pageId - Page ID
     * @param {string} imagePath - Full path to image file
     * @param {Object} position - Position {x, y} (optional)
     */
    insertImage(pageId, imagePath, position = {x: 0, y: 0}) {
        const fs = require('fs');
        const imageData = fs.readFileSync(imagePath, 'base64');
        const imageExt = imagePath.split('.').pop().toLowerCase();
        
        const imageXml = `
            <one:Image>
                <one:Position x="${position.x}" y="${position.y}" />
                <one:Size width="200" height="200" />
                <one:Data>${imageData}</one:Data>
            </one:Image>
        `;
        
        let pageXml = '';
        this.app.invoke('GetPageContent', pageId, pageXml);
        const updatedXml = pageXml + imageXml;
        this.app.invoke('UpdatePageContent', updatedXml);
    }
    
    /**
     * Create formatted text XML
     * @param {string} text - Text content
     * @param {Object} options - Formatting options {bold, italic, fontSize}
     * @returns {string} XML string
     */
    static createFormattedText(text, options = {}) {
        const bold = options.bold ? '<one:T><![CDATA[' + text + ']]></one:T>' : text;
        return `<one:OE>${bold}</one:OE>`;
    }
    
    /**
     * Create table XML
     * @param {number} rows - Number of rows
     * @param {number} cols - Number of columns
     * @returns {string} XML string
     */
    static createTable(rows, cols) {
        let xml = '<one:Table>';
        for (let r = 0; r < rows; r++) {
            xml += '<one:Row>';
            for (let c = 0; c < cols; c++) {
                xml += '<one:Cell><one:OEChildren></one:OEChildren></one:Cell>';
            }
            xml += '</one:Row>';
        }
        xml += '</one:Table>';
        return xml;
    }
    
    /**
     * Create list XML
     * @param {Array<string>} items - List items
     * @param {boolean} ordered - True for numbered list
     * @returns {string} XML string
     */
    static createList(items, ordered = false) {
        const listType = ordered ? 'number' : 'bullet';
        let xml = '<one:OEChildren>';
        items.forEach(item => {
            xml += `<one:OE quickStyleIndex="${listType}"><one:T><![CDATA[${item}]]></one:T></one:OE>`;
        });
        xml += '</one:OEChildren>';
        return xml;
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

module.exports = {
    OneNoteConnector,
};
