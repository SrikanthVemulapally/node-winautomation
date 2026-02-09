/**
 * @fileoverview Microsoft Visio Connector
 * Provides API for Visio diagram automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * Visio save formats
 */
const VisSaveAsFileType = {
    visOpenXML: 0,
    visSaveAsDefault: 0,
    visSaveAsVDX: 1,
    visSaveAsVSX: 2,
    visSaveAsVTX: 3,
    visSaveAsVDW: 4,
    visSaveAsVSW: 5,
};

/**
 * @class VisioConnector
 * @description API for Microsoft Visio automation
 */
class VisioConnector {
    constructor() {
        this.app = new COMObject('Visio.Application');
    }

    /**
     * Set Visio application visibility
     * @param {boolean} visible - True to show Visio window
     */
    setVisible(visible) {
        this.app.setProperty('Visible', visible);
    }

    /**
     * Get Visio application visibility
     * @returns {boolean} Visibility status
     */
    getVisible() {
        return this.app.getProperty('Visible');
    }

    /**
     * Get Visio version
     * @returns {string} Visio version
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Create new document
     * @param {string} template - Template path (optional)
     * @returns {VisioDocument} New document wrapper
     */
    addDocument(template = '') {
        const documents = this.app.getProperty('Documents');
        const doc = template ? documents.invoke('Add', template) : documents.invoke('Add');
        return new VisioDocument(doc);
    }

    /**
     * Open existing document
     * @param {string} path - Full path to document
     * @returns {VisioDocument} Opened document wrapper
     */
    openDocument(path) {
        const documents = this.app.getProperty('Documents');
        const doc = documents.invoke('Open', path);
        return new VisioDocument(doc);
    }

    /**
     * Get active document
     * @returns {VisioDocument} Active document wrapper
     */
    getActiveDocument() {
        const doc = this.app.getProperty('ActiveDocument');
        return new VisioDocument(doc);
    }

    /**
     * Get active page
     * @returns {VisioPage} Active page wrapper
     */
    getActivePage() {
        const page = this.app.getProperty('ActivePage');
        return new VisioPage(page);
    }

    /**
     * Quit Visio application
     */
    quit() {
        this.app.invoke('Quit');
    }

    /**
     * Open stencil file
     * @param {string} path - Full path to stencil file
     * @returns {VisioStencil} Stencil wrapper
     */
    openStencil(path) {
        const documents = this.app.getProperty('Documents');
        const stencil = documents.invoke('OpenEx', path, 4); // visOpenRO (read-only)
        return new VisioStencil(stencil);
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

/**
 * @class VisioDocument
 * @description Wrapper for Visio Document
 */
class VisioDocument {
    constructor(comObject) {
        this.document = comObject;
    }

    /**
     * Get document name
     * @returns {string} Document name
     */
    getName() {
        return this.document.getProperty('Name');
    }

    /**
     * Get full document path
     * @returns {string} Full path
     */
    getFullName() {
        return this.document.getProperty('FullName');
    }

    /**
     * Get document path
     * @returns {string} Directory path
     */
    getPath() {
        return this.document.getProperty('Path');
    }

    /**
     * Get saved status
     * @returns {boolean} True if saved
     */
    getSaved() {
        return this.document.getProperty('Saved');
    }

    /**
     * Save document
     */
    save() {
        this.document.invoke('Save');
    }

    /**
     * Save document with new name
     * @param {string} path - Full path for saved document
     */
    saveAs(path) {
        this.document.invoke('SaveAs', path);
    }

    /**
     * Close document
     */
    close() {
        this.document.invoke('Close');
    }

    /**
     * Export document
     * @param {string} path - Export file path
     */
    export(path) {
        const pages = this.getPages();
        const page = pages.invoke('Item', 1);
        page.invoke('Export', path);
    }

    /**
     * Get pages collection
     * @returns {COMObject} Pages collection
     */
    getPages() {
        return this.document.getProperty('Pages');
    }

    /**
     * Get page count
     * @returns {number} Number of pages
     */
    getPageCount() {
        const pages = this.getPages();
        return pages.getProperty('Count');
    }

    /**
     * Get specific page
     * @param {number} index - Page index (1-based)
     * @returns {VisioPage} Page wrapper
     */
    getPage(index) {
        const pages = this.getPages();
        const page = pages.invoke('Item', index);
        return new VisioPage(page);
    }

    /**
     * Add new page
     * @returns {VisioPage} New page wrapper
     */
    addPage() {
        const pages = this.getPages();
        const page = pages.invoke('Add');
        return new VisioPage(page);
    }

    /**
     * Release COM object
     */
    release() {
        this.document.release();
    }
}

/**
 * @class VisioPage
 * @description Wrapper for Visio Page
 */
class VisioPage {
    constructor(comObject) {
        this.page = comObject;
    }

    /**
     * Get page name
     * @returns {string} Page name
     */
    getName() {
        return this.page.getProperty('Name');
    }

    /**
     * Set page name
     * @param {string} name - Page name
     */
    setName(name) {
        this.page.setProperty('Name', name);
    }

    /**
     * Get shapes collection
     * @returns {COMObject} Shapes collection
     */
    getShapes() {
        return this.page.getProperty('Shapes');
    }

    /**
     * Get shape count
     * @returns {number} Number of shapes
     */
    getShapeCount() {
        const shapes = this.getShapes();
        return shapes.getProperty('Count');
    }

    /**
     * Get specific shape
     * @param {number} index - Shape index (1-based)
     * @returns {VisioShape} Shape wrapper
     */
    getShape(index) {
        const shapes = this.getShapes();
        const shape = shapes.invoke('Item', index);
        return new VisioShape(shape);
    }

    /**
     * Drop shape from stencil
     * @param {COMObject} master - Master shape from stencil
     * @param {number} x - X position
     * @param {number} y - Y position
     * @returns {VisioShape} New shape wrapper
     */
    dropShape(master, x, y) {
        const shape = this.page.invoke('Drop', master, x, y);
        return new VisioShape(shape);
    }

    /**
     * Draw rectangle
     * @param {number} x1 - Start X
     * @param {number} y1 - Start Y
     * @param {number} x2 - End X
     * @param {number} y2 - End Y
     * @returns {VisioShape} New shape wrapper
     */
    drawRectangle(x1, y1, x2, y2) {
        const shape = this.page.invoke('DrawRectangle', x1, y1, x2, y2);
        return new VisioShape(shape);
    }

    /**
     * Draw line
     * @param {number} x1 - Start X
     * @param {number} y1 - Start Y
     * @param {number} x2 - End X
     * @param {number} y2 - End Y
     * @returns {VisioShape} New shape wrapper
     */
    drawLine(x1, y1, x2, y2) {
        const shape = this.page.invoke('DrawLine', x1, y1, x2, y2);
        return new VisioShape(shape);
    }

    /**
     * Draw oval
     * @param {number} x1 - Start X
     * @param {number} y1 - Start Y
     * @param {number} x2 - End X
     * @param {number} y2 - End Y
     * @returns {VisioShape} New shape wrapper
     */
    drawOval(x1, y1, x2, y2) {
        const shape = this.page.invoke('DrawOval', x1, y1, x2, y2);
        return new VisioShape(shape);
    }
    
    /**
     * Drop master shape from stencil
     * @param {COMObject} master - Master shape from stencil
     * @param {number} x - X position
     * @param {number} y - Y position
     * @returns {VisioShape} New shape wrapper
     */
    dropMaster(master, x, y) {
        const shape = this.page.invoke('Drop', master, x, y);
        return new VisioShape(shape);
    }
    
    /**
     * Connect two shapes with a connector
     * @param {VisioShape} fromShape - Source shape
     * @param {VisioShape} toShape - Target shape
     * @returns {VisioShape} Connector shape
     */
    connectShapes(fromShape, toShape) {
        try {
            // Create a connector line
            const connector = this.page.invoke('DrawLine', 0, 0, 1, 1);
            
            // Glue connector to shapes
            const beginX = connector.invoke('CellsU', 'BeginX');
            const endX = connector.invoke('CellsU', 'EndX');
            const fromPinX = fromShape.shape.invoke('CellsU', 'PinX');
            const toPinX = toShape.shape.invoke('CellsU', 'PinX');
            
            beginX.invoke('GlueTo', fromPinX);
            endX.invoke('GlueTo', toPinX);
            
            return new VisioShape(connector);
        } catch (e) {
            throw new Error('Failed to connect shapes: ' + e.message);
        }
    }

    /**
     * Delete page
     */
    delete() {
        this.page.invoke('Delete');
    }

    /**
     * Release COM object
     */
    release() {
        this.page.release();
    }
}

/**
 * @class VisioShape
 * @description Wrapper for Visio Shape
 */
class VisioShape {
    constructor(comObject) {
        this.shape = comObject;
    }

    /**
     * Get shape name
     * @returns {string} Shape name
     */
    getName() {
        return this.shape.getProperty('Name');
    }

    /**
     * Set shape name
     * @param {string} name - Shape name
     */
    setName(name) {
        this.shape.setProperty('Name', name);
    }

    /**
     * Get shape text
     * @returns {string} Shape text
     */
    getText() {
        return this.shape.getProperty('Text');
    }

    /**
     * Set shape text
     * @param {string} text - Shape text
     */
    setText(text) {
        this.shape.setProperty('Text', text);
    }

    /**
     * Delete shape
     */
    delete() {
        this.shape.invoke('Delete');
    }

    /**
     * Release COM object
     */
    release() {
        this.shape.release();
    }
}

/**
 * @class VisioStencil
 * @description Wrapper for Visio Stencil
 */
class VisioStencil {
    constructor(comObject) {
        this.stencil = comObject;
    }
    
    /**
     * Get masters collection
     * @returns {COMObject} Masters collection
     */
    getMasters() {
        return this.stencil.getProperty('Masters');
    }
    
    /**
     * Get master by name
     * @param {string} name - Master name
     * @returns {COMObject} Master shape
     */
    getMaster(name) {
        const masters = this.getMasters();
        return masters.invoke('Item', name);
    }
    
    /**
     * Get master by index
     * @param {number} index - Master index (1-based)
     * @returns {COMObject} Master shape
     */
    getMasterByIndex(index) {
        const masters = this.getMasters();
        return masters.invoke('Item', index);
    }
    
    /**
     * Get master count
     * @returns {number} Number of masters
     */
    getMasterCount() {
        const masters = this.getMasters();
        return masters.getProperty('Count');
    }
    
    /**
     * Close stencil
     */
    close() {
        this.stencil.invoke('Close');
    }
    
    /**
     * Release COM object
     */
    release() {
        this.stencil.release();
    }
}

module.exports = {
    VisioConnector,
    VisioDocument,
    VisioPage,
    VisioShape,
    VisioStencil,
    VisSaveAsFileType,
};
