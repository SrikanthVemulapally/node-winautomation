/**
 * @fileoverview Outlook Rule Class
 */

/**
 * @class OutlookRule
 * @description Wrapper for Outlook Rule
 */
class OutlookRule {
    constructor(comObject) {
        this.rule = comObject;
    }
    
    getName() {
        return this.rule.getProperty('Name');
    }
    
    setName(name) {
        this.rule.setProperty('Name', name);
    }
    
    getEnabled() {
        return this.rule.getProperty('Enabled');
    }
    
    setEnabled(enabled) {
        this.rule.setProperty('Enabled', enabled);
    }
    
    getConditions() {
        return this.rule.getProperty('Conditions');
    }
    
    getActions() {
        return this.rule.getProperty('Actions');
    }
    
    getExceptions() {
        return this.rule.getProperty('Exceptions');
    }
    
    execute(showProgress = false, folder = null) {
        if (folder) {
            this.rule.invoke('Execute', showProgress, folder.folder);
        } else {
            this.rule.invoke('Execute', showProgress);
        }
    }
    
    release() {
        this.rule.release();
    }
}

module.exports = { OutlookRule };
