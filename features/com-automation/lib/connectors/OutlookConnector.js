/**
 * @fileoverview Microsoft Outlook Connector
 * Provides a high-level, type-safe API for Outlook automation
 */

const { COMObject } = require('../../../../build/Release/Automation.node');

/**
 * Outlook item types
 */
const OlItemType = {
    olMailItem: 0,
    olAppointmentItem: 1,
    olContactItem: 2,
    olTaskItem: 3,
    olJournalItem: 4,
    olNoteItem: 5,
    olPostItem: 6,
    olDistributionListItem: 7
};

/**
 * Outlook default folders
 */
const OlDefaultFolders = {
    olFolderDeletedItems: 3,
    olFolderOutbox: 4,
    olFolderSentMail: 5,
    olFolderInbox: 6,
    olFolderCalendar: 9,
    olFolderContacts: 10,
    olFolderJournal: 11,
    olFolderNotes: 12,
    olFolderTasks: 13,
    olFolderDrafts: 16
};

/**
 * Importance levels
 */
const OlImportance = {
    olImportanceLow: 0,
    olImportanceNormal: 1,
    olImportanceHigh: 2
};

/**
 * @class OutlookConnector
 * @description High-level API for Microsoft Outlook automation
 */
class OutlookConnector {
    constructor() {
        this.app = new COMObject('Outlook.Application');
        this.namespace = null;
        this.eventCallbacks = new Map();
    }

    /**
     * Get the MAPI namespace
     * @returns {COMObject} Namespace object
     */
    getNamespace() {
        if (!this.namespace) {
            this.namespace = this.app.invoke('GetNamespace', 'MAPI');
        }
        return this.namespace;
    }

    /**
     * Get Outlook version
     * @returns {string} Version string
     */
    getVersion() {
        return this.app.getProperty('Version');
    }

    /**
     * Create a new mail item
     * @returns {MailItem} Mail item wrapper
     */
    createMailItem() {
        const item = this.app.invoke('CreateItem', OlItemType.olMailItem);
        return new MailItem(item);
    }

    /**
     * Create a new appointment
     * @returns {AppointmentItem} Appointment item wrapper
     */
    createAppointmentItem() {
        const item = this.app.invoke('CreateItem', OlItemType.olAppointmentItem);
        return new AppointmentItem(item);
    }

    /**
     * Create a new contact
     * @returns {ContactItem} Contact item wrapper
     */
    createContactItem() {
        const item = this.app.invoke('CreateItem', OlItemType.olContactItem);
        return new ContactItem(item);
    }

    /**
     * Create a new task
     * @returns {TaskItem} Task item wrapper
     */
    createTaskItem() {
        const item = this.app.invoke('CreateItem', OlItemType.olTaskItem);
        return new TaskItem(item);
    }

    /**
     * Get a default folder
     * @param {number} folderType - Folder type constant
     * @returns {Folder} Folder wrapper
     */
    getDefaultFolder(folderType) {
        const ns = this.getNamespace();
        const folder = ns.invoke('GetDefaultFolder', folderType);
        return new Folder(folder);
    }

    /**
     * Get inbox folder
     * @returns {Folder} Inbox folder
     */
    getInbox() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderInbox);
    }

    /**
     * Get sent mail folder
     * @returns {Folder} Sent mail folder
     */
    getSentMail() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderSentMail);
    }

    /**
     * Get drafts folder
     * @returns {Folder} Drafts folder
     */
    getDrafts() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderDrafts);
    }

    /**
     * Get calendar folder
     * @returns {Folder} Calendar folder
     */
    getCalendar() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderCalendar);
    }

    /**
     * Get contacts folder
     * @returns {Folder} Contacts folder
     */
    getContacts() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderContacts);
    }

    /**
     * Get tasks folder
     * @returns {Folder} Tasks folder
     */
    getTasks() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderTasks);
    }

    /**
     * Subscribe to new mail event
     * @param {Function} callback - Callback function(entryId)
     */
    onNewMail(callback) {
        const iid = '{0006304E-0000-0000-C000-000000000046}';
        this.app.adviseEvent(iid, (event) => {
            if (event.dispId === 64181) { // NewMailEx
                const entryId = event.args[0];
                callback(entryId);
            }
        });
        this.eventCallbacks.set('newmail', iid);
    }

    /**
     * Subscribe to item send event
     * @param {Function} callback - Callback function(item, cancel)
     */
    onItemSend(callback) {
        const iid = '{0006304E-0000-0000-C000-000000000046}';
        this.app.adviseEvent(iid, (event) => {
            if (event.dispId === 61445) { // ItemSend
                callback(event.args[0], event.args[1]);
            }
        });
        this.eventCallbacks.set('itemsend', iid);
    }

    /**
     * Unsubscribe from events
     */
    removeAllEvents() {
        for (const iid of this.eventCallbacks.values()) {
            try {
                this.app.unadviseEvent(iid);
            } catch (e) {
                // Ignore errors during cleanup
            }
        }
        this.eventCallbacks.clear();
    }

    /**
     * Quit Outlook
     */
    quit() {
        this.removeAllEvents();
        this.app.invoke('Quit');
    }

    /**
     * Release resources
     */
    release() {
        this.removeAllEvents();
        if (this.namespace) {
            this.namespace.release();
            this.namespace = null;
        }
        this.app.release();
    }
}

/**
 * @class MailItem
 * @description Wrapper for Outlook MailItem
 */
class MailItem {
    constructor(comObject) {
        this.item = comObject;
    }

    getSubject() { return this.item.getProperty('Subject'); }
    setSubject(value) { this.item.setProperty('Subject', value); }

    getBody() { return this.item.getProperty('Body'); }
    setBody(value) { this.item.setProperty('Body', value); }

    getHTMLBody() { return this.item.getProperty('HTMLBody'); }
    setHTMLBody(value) { this.item.setProperty('HTMLBody', value); }

    getTo() { return this.item.getProperty('To'); }
    setTo(value) { this.item.setProperty('To', value); }

    getCC() { return this.item.getProperty('CC'); }
    setCC(value) { this.item.setProperty('CC', value); }

    getBCC() { return this.item.getProperty('BCC'); }
    setBCC(value) { this.item.setProperty('BCC', value); }

    getImportance() { return this.item.getProperty('Importance'); }
    setImportance(value) { this.item.setProperty('Importance', value); }

    getSentOn() { return this.item.getProperty('SentOn'); }
    getReceivedTime() { return this.item.getProperty('ReceivedTime'); }

    getUnRead() { return this.item.getProperty('UnRead'); }
    setUnRead(value) { this.item.setProperty('UnRead', value); }

    getSize() { return this.item.getProperty('Size'); }

    send() { this.item.invoke('Send'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    save() { this.item.invoke('Save'); }
    close(saveMode = 0) { this.item.invoke('Close', saveMode); }
    
    reply() { return new MailItem(this.item.invoke('Reply')); }
    replyAll() { return new MailItem(this.item.invoke('ReplyAll')); }
    forward() { return new MailItem(this.item.invoke('Forward')); }

    getAttachments() {
        return this.item.getProperty('Attachments');
    }

    addAttachment(path, type = 1, position = 1, displayName = '') {
        const attachments = this.getAttachments();
        return attachments.invoke('Add', path, type, position, displayName);
    }

    release() { this.item.release(); }
}

/**
 * @class AppointmentItem
 * @description Wrapper for Outlook AppointmentItem
 */
class AppointmentItem {
    constructor(comObject) {
        this.item = comObject;
    }

    getSubject() { return this.item.getProperty('Subject'); }
    setSubject(value) { this.item.setProperty('Subject', value); }

    getLocation() { return this.item.getProperty('Location'); }
    setLocation(value) { this.item.setProperty('Location', value); }

    getStart() { return this.item.getProperty('Start'); }
    setStart(value) { this.item.setProperty('Start', value); }

    getEnd() { return this.item.getProperty('End'); }
    setEnd(value) { this.item.setProperty('End', value); }

    getBody() { return this.item.getProperty('Body'); }
    setBody(value) { this.item.setProperty('Body', value); }

    getAllDayEvent() { return this.item.getProperty('AllDayEvent'); }
    setAllDayEvent(value) { this.item.setProperty('AllDayEvent', value); }

    getReminderSet() { return this.item.getProperty('ReminderSet'); }
    setReminderSet(value) { this.item.setProperty('ReminderSet', value); }

    getReminderMinutesBeforeStart() { return this.item.getProperty('ReminderMinutesBeforeStart'); }
    setReminderMinutesBeforeStart(value) { this.item.setProperty('ReminderMinutesBeforeStart', value); }

    getBusyStatus() { return this.item.getProperty('BusyStatus'); }
    setBusyStatus(value) { this.item.setProperty('BusyStatus', value); }

    save() { this.item.invoke('Save'); }
    send() { this.item.invoke('Send'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    
    release() { this.item.release(); }
}

/**
 * @class ContactItem
 * @description Wrapper for Outlook ContactItem
 */
class ContactItem {
    constructor(comObject) {
        this.item = comObject;
    }

    getFirstName() { return this.item.getProperty('FirstName'); }
    setFirstName(value) { this.item.setProperty('FirstName', value); }

    getLastName() { return this.item.getProperty('LastName'); }
    setLastName(value) { this.item.setProperty('LastName', value); }

    getFullName() { return this.item.getProperty('FullName'); }
    setFullName(value) { this.item.setProperty('FullName', value); }

    getEmail1Address() { return this.item.getProperty('Email1Address'); }
    setEmail1Address(value) { this.item.setProperty('Email1Address', value); }

    getBusinessTelephoneNumber() { return this.item.getProperty('BusinessTelephoneNumber'); }
    setBusinessTelephoneNumber(value) { this.item.setProperty('BusinessTelephoneNumber', value); }

    getMobileTelephoneNumber() { return this.item.getProperty('MobileTelephoneNumber'); }
    setMobileTelephoneNumber(value) { this.item.setProperty('MobileTelephoneNumber', value); }

    getCompanyName() { return this.item.getProperty('CompanyName'); }
    setCompanyName(value) { this.item.setProperty('CompanyName', value); }

    getJobTitle() { return this.item.getProperty('JobTitle'); }
    setJobTitle(value) { this.item.setProperty('JobTitle', value); }

    save() { this.item.invoke('Save'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    delete() { this.item.invoke('Delete'); }
    
    release() { this.item.release(); }
}

/**
 * @class TaskItem
 * @description Wrapper for Outlook TaskItem
 */
class TaskItem {
    constructor(comObject) {
        this.item = comObject;
    }

    getSubject() { return this.item.getProperty('Subject'); }
    setSubject(value) { this.item.setProperty('Subject', value); }

    getBody() { return this.item.getProperty('Body'); }
    setBody(value) { this.item.setProperty('Body', value); }

    getStartDate() { return this.item.getProperty('StartDate'); }
    setStartDate(value) { this.item.setProperty('StartDate', value); }

    getDueDate() { return this.item.getProperty('DueDate'); }
    setDueDate(value) { this.item.setProperty('DueDate', value); }

    getStatus() { return this.item.getProperty('Status'); }
    setStatus(value) { this.item.setProperty('Status', value); }

    getPriority() { return this.item.getProperty('Priority'); }
    setPriority(value) { this.item.setProperty('Priority', value); }

    getPercentComplete() { return this.item.getProperty('PercentComplete'); }
    setPercentComplete(value) { this.item.setProperty('PercentComplete', value); }

    getComplete() { return this.item.getProperty('Complete'); }
    setComplete(value) { this.item.setProperty('Complete', value); }

    save() { this.item.invoke('Save'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    
    release() { this.item.release(); }
}

/**
 * @class Folder
 * @description Wrapper for Outlook Folder
 */
class Folder {
    constructor(comObject) {
        this.folder = comObject;
    }

    getName() { return this.folder.getProperty('Name'); }
    getUnReadItemCount() { return this.folder.getProperty('UnReadItemCount'); }

    getItems() {
        return this.folder.getProperty('Items');
    }

    getItem(index) {
        const items = this.getItems();
        return items.invoke('Item', index);
    }

    getCount() {
        const items = this.getItems();
        return items.getProperty('Count');
    }

    release() { this.folder.release(); }
}

module.exports = {
    OutlookConnector,
    MailItem,
    AppointmentItem,
    ContactItem,
    TaskItem,
    Folder,
    OlItemType,
    OlDefaultFolders,
    OlImportance
};
