/**
 * @fileoverview Microsoft Access Connector
 * Provides API for Access database automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * @class AccessConnector
 * @description API for Microsoft Access automation
 */
class AccessConnector {
    constructor() {
        this.app = new COMObject('Access.Application');
    }

    /**
     * Set Access application visibility
     * @param {boolean} visible - True to show Access window
     */
    setVisible(visible) {
        this.app.setProperty('Visible', visible);
    }

    /**
     * Get Access application visibility
     * @returns {boolean} Visibility status
     */
    getVisible() {
        return this.app.getProperty('Visible');
    }

    /**
     * Get Access version
     * @returns {string} Access version
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Open database
     * @param {string} path - Full path to database file
     */
    openDatabase(path) {
        this.app.invoke('OpenCurrentDatabase', path);
    }

    /**
     * Close current database
     */
    closeDatabase() {
        this.app.invoke('CloseCurrentDatabase');
    }

    /**
     * Get current database path
     * @returns {string} Database path
     */
    getCurrentDatabasePath() {
        const currentDb = this.app.getProperty('CurrentDb');
        return currentDb.getProperty('Name');
    }

    /**
     * Execute SQL query
     * @param {string} sql - SQL query string
     */
    runSQL(sql) {
        this.app.invoke('DoCmd').invoke('RunSQL', sql);
    }

    /**
     * Open table
     * @param {string} tableName - Table name
     */
    openTable(tableName) {
        this.app.invoke('DoCmd').invoke('OpenTable', tableName);
    }

    /**
     * Open query
     * @param {string} queryName - Query name
     */
    openQuery(queryName) {
        this.app.invoke('DoCmd').invoke('OpenQuery', queryName);
    }

    /**
     * Open form
     * @param {string} formName - Form name
     */
    openForm(formName) {
        this.app.invoke('DoCmd').invoke('OpenForm', formName);
    }

    /**
     * Open report
     * @param {string} reportName - Report name
     */
    openReport(reportName) {
        this.app.invoke('DoCmd').invoke('OpenReport', reportName);
    }

    /**
     * Close object
     * @param {number} objectType - Object type (0=Table, 1=Query, 2=Form, 3=Report)
     * @param {string} objectName - Object name
     */
    closeObject(objectType, objectName) {
        this.app.invoke('DoCmd').invoke('Close', objectType, objectName);
    }

    /**
     * Quit Access application
     */
    quit() {
        this.app.invoke('Quit');
    }

    /**
     * Execute query and return recordset
     * @param {string} sql - SQL query
     * @returns {AccessRecordset} Recordset wrapper
     */
    executeQuery(sql) {
        const db = this.app.invoke('CurrentDb');
        const recordset = db.invoke('OpenRecordset', sql);
        return new AccessRecordset(recordset);
    }
    
    /**
     * Get records from table
     * @param {string} tableName - Table name
     * @returns {AccessRecordset} Recordset wrapper
     */
    getRecords(tableName) {
        const db = this.app.invoke('CurrentDb');
        const recordset = db.invoke('OpenRecordset', tableName);
        return new AccessRecordset(recordset);
    }
    
    /**
     * Find records matching criteria
     * @param {string} tableName - Table name
     * @param {string} criteria - WHERE clause criteria
     * @returns {AccessRecordset} Recordset wrapper
     */
    findRecords(tableName, criteria) {
        const sql = `SELECT * FROM ${tableName} WHERE ${criteria}`;
        return this.executeQuery(sql);
    }

    /**
     * Release COM object
     */
    release() {
        this.app.release();
    }
}

/**
 * @class AccessRecordset
 * @description Wrapper for Access Recordset (DAO/ADO)
 */
class AccessRecordset {
    constructor(comObject) {
        this.rs = comObject;
    }
    
    /**
     * Move to first record
     */
    moveFirst() {
        this.rs.invoke('MoveFirst');
    }
    
    /**
     * Move to last record
     */
    moveLast() {
        this.rs.invoke('MoveLast');
    }
    
    /**
     * Move to next record
     */
    moveNext() {
        this.rs.invoke('MoveNext');
    }
    
    /**
     * Move to previous record
     */
    movePrevious() {
        this.rs.invoke('MovePrevious');
    }
    
    /**
     * Get field value
     * @param {string} fieldName - Field name
     * @returns {*} Field value
     */
    getFieldValue(fieldName) {
        const fields = this.rs.getProperty('Fields');
        const field = fields.invoke('Item', fieldName);
        return field.getProperty('Value');
    }
    
    /**
     * Set field value
     * @param {string} fieldName - Field name
     * @param {*} value - Value to set
     */
    setFieldValue(fieldName, value) {
        const fields = this.rs.getProperty('Fields');
        const field = fields.invoke('Item', fieldName);
        field.setProperty('Value', value);
    }
    
    /**
     * Add new record
     */
    addNew() {
        this.rs.invoke('AddNew');
    }
    
    /**
     * Update current record
     */
    update() {
        this.rs.invoke('Update');
    }
    
    /**
     * Delete current record
     */
    delete() {
        this.rs.invoke('Delete');
    }
    
    /**
     * Get record count
     * @returns {number} Number of records
     */
    getRecordCount() {
        return this.rs.getProperty('RecordCount');
    }
    
    /**
     * Check if at end of file
     * @returns {boolean} True if EOF
     */
    isEOF() {
        return this.rs.getProperty('EOF');
    }
    
    /**
     * Check if at beginning of file
     * @returns {boolean} True if BOF
     */
    isBOF() {
        return this.rs.getProperty('BOF');
    }
    
    /**
     * Find first record matching criteria
     * @param {string} criteria - Search criteria
     */
    findFirst(criteria) {
        this.rs.invoke('FindFirst', criteria);
    }
    
    /**
     * Find next record matching criteria
     * @param {string} criteria - Search criteria
     */
    findNext(criteria) {
        this.rs.invoke('FindNext', criteria);
    }
    
    /**
     * Close recordset
     */
    close() {
        this.rs.invoke('Close');
    }
    
    /**
     * Release COM object
     */
    release() {
        this.rs.release();
    }
}

module.exports = {
    AccessConnector,
    AccessRecordset,
};
