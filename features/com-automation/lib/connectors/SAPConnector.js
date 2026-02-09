/**
 * @fileoverview SAP GUI Connector
 * Provides API for SAP GUI Scripting automation
 * Note: Requires SAP GUI with scripting enabled
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * @class SAPConnector
 * @description API for SAP GUI Scripting automation
 */
class SAPConnector {
    constructor() {
        try {
            this.engine = new COMObject('Sapgui.ScriptingCtrl.1');
        } catch (e) {
            throw new Error('SAP GUI not installed or scripting not enabled: ' + e.message);
        }
    }

    /**
     * Get SAP GUI application
     * @returns {SAPApplication} SAP application wrapper
     */
    getApplication() {
        const app = this.engine.invoke('GetScriptingEngine');
        return new SAPApplication(app);
    }

    /**
     * Release COM object
     */
    release() {
        this.engine.release();
    }
}

/**
 * @class SAPApplication
 * @description Wrapper for SAP GUI Application
 */
class SAPApplication {
    constructor(comObject) {
        this.app = comObject;
    }

    /**
     * Get connections collection
     * @returns {COMObject} Connections collection
     */
    getConnections() {
        return this.app.getProperty('Connections');
    }

    /**
     * Get connection count
     * @returns {number} Number of connections
     */
    getConnectionCount() {
        const connections = this.getConnections();
        return connections.getProperty('Count');
    }

    /**
     * Get specific connection
     * @param {number} index - Connection index (0-based)
     * @returns {SAPConnection} Connection wrapper
     */
    getConnection(index) {
        const connections = this.getConnections();
        const conn = connections.invoke('Item', index);
        return new SAPConnection(conn);
    }

    /**
     * Open connection
     * @param {string} description - Connection description
     * @param {boolean} sync - Synchronous connection
     * @returns {SAPConnection} New connection wrapper
     */
    openConnection(description, sync = true) {
        const conn = this.app.invoke('OpenConnection', description, sync);
        return new SAPConnection(conn);
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

/**
 * @class SAPConnection
 * @description Wrapper for SAP Connection
 */
class SAPConnection {
    constructor(comObject) {
        this.connection = comObject;
    }

    /**
     * Get connection description
     * @returns {string} Connection description
     */
    getDescription() {
        return this.connection.getProperty('Description');
    }

    /**
     * Get sessions collection
     * @returns {COMObject} Sessions collection
     */
    getSessions() {
        return this.connection.getProperty('Sessions');
    }

    /**
     * Get session count
     * @returns {number} Number of sessions
     */
    getSessionCount() {
        const sessions = this.getSessions();
        return sessions.getProperty('Count');
    }

    /**
     * Get specific session
     * @param {number} index - Session index (0-based)
     * @returns {SAPSession} Session wrapper
     */
    getSession(index) {
        const sessions = this.getSessions();
        const session = sessions.invoke('Item', index);
        return new SAPSession(session);
    }

    /**
     * Close connection
     */
    close() {
        this.connection.invoke('CloseConnection');
    }

    /**
     * Release COM object
     */
    release() {
        this.connection.release();
    }
}

/**
 * @class SAPSession
 * @description Wrapper for SAP Session
 */
class SAPSession {
    constructor(comObject) {
        this.session = comObject;
    }

    /**
     * Get session info
     * @returns {COMObject} Session info object
     */
    getInfo() {
        return this.session.getProperty('Info');
    }

    /**
     * Get transaction code
     * @returns {string} Current transaction code
     */
    getTransaction() {
        const info = this.getInfo();
        return info.getProperty('Transaction');
    }

    /**
     * Start transaction
     * @param {string} tcode - Transaction code
     */
    startTransaction(tcode) {
        this.session.invoke('StartTransaction', tcode);
    }

    /**
     * End transaction
     */
    endTransaction() {
        this.session.invoke('EndTransaction');
    }

    /**
     * Find element by ID
     * @param {string} id - Element ID
     * @returns {COMObject} Element object
     */
    findById(id) {
        return this.session.invoke('FindById', id);
    }

    /**
     * Set field value
     * @param {string} id - Field ID
     * @param {string} value - Value to set
     */
    setFieldValue(id, value) {
        const field = this.findById(id);
        field.setProperty('Text', value);
    }

    /**
     * Get field value
     * @param {string} id - Field ID
     * @returns {string} Field value
     */
    getFieldValue(id) {
        const field = this.findById(id);
        return field.getProperty('Text');
    }

    /**
     * Press button
     * @param {string} id - Button ID
     */
    pressButton(id) {
        const button = this.findById(id);
        button.invoke('Press');
    }

    /**
     * Select menu item
     * @param {string} id - Menu item ID
     */
    selectMenuItem(id) {
        const menu = this.findById(id);
        menu.invoke('Select');
    }

    /**
     * Send VKey (function key)
     * @param {number} key - VKey code (0=Enter, 3=Back, 12=Save, 15=Exit, etc.)
     */
    sendVKey(key) {
        this.session.invoke('SendVKey', key);
    }

    /**
     * Create session (open new window)
     */
    createSession() {
        this.session.invoke('CreateSession');
    }

    /**
     * Lock/unlock session
     * @param {boolean} lock - True to lock, false to unlock
     */
    lockSession(lock) {
        this.session.setProperty('LockSessionUI', lock);
    }
    
    /**
     * Get status bar text
     * @returns {string} Status bar message
     */
    getStatusBarText() {
        try {
            const statusbar = this.findById('wnd[0]/sbar');
            return statusbar.getProperty('Text');
        } catch (e) {
            return '';
        }
    }
    
    /**
     * Get status bar message type
     * @returns {string} Message type (S=Success, E=Error, W=Warning, I=Info, A=Abort)
     */
    getStatusBarType() {
        try {
            const statusbar = this.findById('wnd[0]/sbar');
            return statusbar.getProperty('MessageType');
        } catch (e) {
            return '';
        }
    }
    
    /**
     * Take screenshot
     * @param {string} filename - Output filename
     */
    takeScreenshot(filename) {
        try {
            const shell = this.findById('wnd[0]');
            shell.invoke('HardCopy', filename, 1); // 1 = BMP format
        } catch (e) {
            throw new Error('Failed to take screenshot: ' + e.message);
        }
    }
    
    /**
     * Get grid view
     * @param {string} id - Grid control ID
     * @returns {SAPGridView} Grid view wrapper
     */
    getGridView(id) {
        const grid = this.findById(id);
        return new SAPGridView(grid);
    }

    /**
     * Release COM object
     */
    release() {
        this.session.release();
    }
}

/**
 * @class SAPGridView
 * @description Wrapper for SAP Grid Control
 */
class SAPGridView {
    constructor(comObject) {
        this.grid = comObject;
    }
    
    /**
     * Get row count
     * @returns {number} Number of rows
     */
    getRowCount() {
        return this.grid.getProperty('RowCount');
    }
    
    /**
     * Get column count
     * @returns {number} Number of columns
     */
    getColumnCount() {
        return this.grid.getProperty('ColumnCount');
    }
    
    /**
     * Get cell value
     * @param {number} row - Row index (0-based)
     * @param {string} column - Column name
     * @returns {string} Cell value
     */
    getCellValue(row, column) {
        return this.grid.invoke('GetCellValue', row, column);
    }
    
    /**
     * Set cell value
     * @param {number} row - Row index (0-based)
     * @param {string} column - Column name
     * @param {string} value - Value to set
     */
    setCellValue(row, column, value) {
        this.grid.invoke('ModifyCell', row, column, value);
    }
    
    /**
     * Select row
     * @param {number} row - Row index (0-based)
     */
    selectRow(row) {
        this.grid.setProperty('SelectedRows', row.toString());
    }
    
    /**
     * Double click cell
     * @param {number} row - Row index (0-based)
     * @param {string} column - Column name
     */
    doubleClickCell(row, column) {
        this.grid.invoke('DoubleClick', row, column);
    }
    
    /**
     * Get selected rows
     * @returns {string} Selected rows
     */
    getSelectedRows() {
        return this.grid.getProperty('SelectedRows');
    }
    
    /**
     * Press toolbar button
     * @param {string} buttonId - Button ID
     */
    pressButton(buttonId) {
        this.grid.invoke('PressButton', buttonId);
    }
    
    /**
     * Release COM object
     */
    release() {
        this.grid.release();
    }
}

/**
 * SAP VKey codes
 */
const SAPVKey = {
    Enter: 0,
    F1: 1,
    F2: 2,
    F3: 3,      // Back
    F4: 4,
    F5: 5,
    F6: 6,
    F7: 7,
    F8: 8,
    F9: 9,
    F10: 10,
    F11: 11,
    F12: 12,    // Save
    F15: 15,    // Exit
    F16: 16,
    F17: 17,
    F18: 18,
    F19: 19,
    F20: 20,
    F21: 21,
    F22: 22,
    F23: 23,
    F24: 24,
};

module.exports = {
    SAPConnector,
    SAPApplication,
    SAPConnection,
    SAPSession,
    SAPGridView,
    SAPVKey,
};
