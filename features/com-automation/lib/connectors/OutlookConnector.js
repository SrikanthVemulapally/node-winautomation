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
     * Get deleted items folder
     * @returns {Folder} Deleted items folder
     */
    getDeletedItems() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
    }

    /**
     * Get outbox folder
     * @returns {Folder} Outbox folder
     */
    getOutbox() {
        return this.getDefaultFolder(OlDefaultFolders.olFolderOutbox);
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
     * Get folder by path
     * @param {string} path - Folder path (e.g., "Inbox\\Subfolder")
     * @returns {Folder} Folder wrapper
     */
    getFolderByPath(path) {
        const ns = this.getNamespace();
        const folder = ns.invoke('GetFolderFromID', path);
        return new Folder(folder);
    }

    /**
     * Create a new folder
     * @param {string} name - Folder name
     * @param {Folder} parentFolder - Parent folder
     * @returns {Folder} New folder
     */
    createFolder(name, parentFolder) {
        const folders = parentFolder.folder.getProperty('Folders');
        const newFolder = folders.invoke('Add', name);
        return new Folder(newFolder);
    }

    /**
     * Empty deleted items folder
     */
    emptyDeletedItems() {
        const ns = this.getNamespace();
        const deletedItems = this.getDeletedItems();
        const items = deletedItems.getItems();
        const count = items.getProperty('Count');
        for (let i = count; i >= 1; i--) {
            const item = items.invoke('Item', i);
            item.invoke('Delete');
        }
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
     * Get rules collection
     * @returns {COMObject} Rules collection
     */
    getRules() {
        const session = this.getNamespace();
        const store = session.getProperty('DefaultStore');
        return store.getProperty('Rules');
    }
    
    /**
     * Create new rule
     * @param {string} name - Rule name
     * @param {number} ruleType - Rule type (0=Receive, 1=Send)
     * @returns {OutlookRule} New rule wrapper
     */
    createRule(name, ruleType = 0) {
        const rules = this.getRules();
        const rule = rules.invoke('Create', name, ruleType);
        return new OutlookRule(rule);
    }
    
    /**
     * Get rule by name
     * @param {string} name - Rule name
     * @returns {OutlookRule} Rule wrapper
     */
    getRule(name) {
        const rules = this.getRules();
        const rule = rules.invoke('Item', name);
        return new OutlookRule(rule);
    }
    
    /**
     * Save rules
     */
    saveRules() {
        const rules = this.getRules();
        rules.invoke('Save');
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

    getSenderName() { return this.item.getProperty('SenderName'); }
    getSenderEmailAddress() { return this.item.getProperty('SenderEmailAddress'); }

    getCategories() { return this.item.getProperty('Categories'); }
    setCategories(value) { this.item.setProperty('Categories', value); }
    
    /**
     * Add category to item
     * @param {string} category - Category name
     */
    addCategory(category) {
        const current = this.getCategories();
        if (current && current.length > 0) {
            this.setCategories(current + '; ' + category);
        } else {
            this.setCategories(category);
        }
    }
    
    /**
     * Remove category from item
     * @param {string} category - Category name
     */
    removeCategory(category) {
        const current = this.getCategories();
        if (current) {
            const categories = current.split(';').map(c => c.trim());
            const filtered = categories.filter(c => c !== category);
            this.setCategories(filtered.join('; '));
        }
    }
    
    /**
     * Clear all categories
     */
    clearCategories() {
        this.setCategories('');
    }

    getFlagStatus() { return this.item.getProperty('FlagStatus'); }
    setFlagStatus(value) { this.item.setProperty('FlagStatus', value); }
    
    /**
     * Set voting options for mail
     * @param {string} options - Voting options (e.g., 'Approve;Reject')
     */
    setVotingOptions(options) {
        this.item.setProperty('VotingOptions', options);
    }
    
    /**
     * Get voting options
     * @returns {string} Voting options
     */
    getVotingOptions() {
        return this.item.getProperty('VotingOptions');
    }
    
    /**
     * Get voting response
     * @returns {string} Voting response
     */
    getVotingResponse() {
        return this.item.getProperty('VotingResponse');
    }

    send() { this.item.invoke('Send'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    save() { this.item.invoke('Save'); }
    close(saveMode = 0) { this.item.invoke('Close', saveMode); }
    delete() { this.item.invoke('Delete'); }
    
    reply() { return new MailItem(this.item.invoke('Reply')); }
    replyAll() { return new MailItem(this.item.invoke('ReplyAll')); }
    forward() { return new MailItem(this.item.invoke('Forward')); }

    /**
     * Move to folder
     * @param {Folder} folder - Destination folder
     */
    move(folder) {
        this.item.invoke('Move', folder.folder);
    }

    /**
     * Copy to folder
     * @param {Folder} folder - Destination folder
     * @returns {MailItem} Copied item
     */
    copy(folder) {
        const copied = this.item.invoke('Copy');
        const copiedItem = new MailItem(copied);
        copiedItem.move(folder);
        return copiedItem;
    }

    /**
     * Mark as read
     */
    markAsRead() {
        this.setUnRead(false);
    }

    /**
     * Mark as unread
     */
    markAsUnread() {
        this.setUnRead(true);
    }

    getAttachments() {
        return this.item.getProperty('Attachments');
    }

    addAttachment(path, type = 1, position = 1, displayName = '') {
        const attachments = this.getAttachments();
        if (!attachments) {
            throw new Error('Failed to get attachments collection');
        }
        return attachments.invoke('Add', path, type, position, displayName);
    }

    /**
     * Get attachment count
     * @returns {number} Number of attachments
     */
    getAttachmentCount() {
        const attachments = this.getAttachments();
        return attachments ? attachments.getProperty('Count') : 0;
    }

    /**
     * Save attachment to file
     * @param {number} index - Attachment index (1-based)
     * @param {string} path - Save path
     */
    saveAttachment(index, path) {
        const attachments = this.getAttachments();
        const attachment = attachments.invoke('Item', index);
        attachment.invoke('SaveAsFile', path);
    }

    /**
     * Remove attachment
     * @param {number} index - Attachment index (1-based)
     */
    removeAttachment(index) {
        const attachments = this.getAttachments();
        const attachment = attachments.invoke('Item', index);
        attachment.invoke('Delete');
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

    /**
     * Get required attendees
     * @returns {string} Required attendees
     */
    getRequiredAttendees() { return this.item.getProperty('RequiredAttendees'); }
    setRequiredAttendees(value) { this.item.setProperty('RequiredAttendees', value); }

    /**
     * Get optional attendees
     * @returns {string} Optional attendees
     */
    getOptionalAttendees() { return this.item.getProperty('OptionalAttendees'); }
    setOptionalAttendees(value) { this.item.setProperty('OptionalAttendees', value); }

    /**
     * Get meeting status
     * @returns {number} Meeting status
     */
    getMeetingStatus() { return this.item.getProperty('MeetingStatus'); }
    setMeetingStatus(value) { this.item.setProperty('MeetingStatus', value); }

    save() { this.item.invoke('Save'); }
    send() { this.item.invoke('Send'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    delete() { this.item.invoke('Delete'); }
    
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

    getEmail2Address() { return this.item.getProperty('Email2Address'); }
    setEmail2Address(value) { this.item.setProperty('Email2Address', value); }

    getEmail3Address() { return this.item.getProperty('Email3Address'); }
    setEmail3Address(value) { this.item.setProperty('Email3Address', value); }

    getHomeTelephoneNumber() { return this.item.getProperty('HomeTelephoneNumber'); }
    setHomeTelephoneNumber(value) { this.item.setProperty('HomeTelephoneNumber', value); }

    getBusinessAddress() { return this.item.getProperty('BusinessAddress'); }
    setBusinessAddress(value) { this.item.setProperty('BusinessAddress', value); }

    getHomeAddress() { return this.item.getProperty('HomeAddress'); }
    setHomeAddress(value) { this.item.setProperty('HomeAddress', value); }

    getWebPage() { return this.item.getProperty('WebPage'); }
    setWebPage(value) { this.item.setProperty('WebPage', value); }

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

    getOwner() { return this.item.getProperty('Owner'); }
    setOwner(value) { this.item.setProperty('Owner', value); }

    getActualWork() { return this.item.getProperty('ActualWork'); }
    setActualWork(value) { this.item.setProperty('ActualWork', value); }

    getTotalWork() { return this.item.getProperty('TotalWork'); }
    setTotalWork(value) { this.item.setProperty('TotalWork', value); }

    save() { this.item.invoke('Save'); }
    display(modal = false) { this.item.invoke('Display', modal); }
    delete() { this.item.invoke('Delete'); }

    /**
     * Mark task as complete
     */
    markComplete() {
        this.setComplete(true);
        this.setPercentComplete(100);
        this.save();
    }
    
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
        if (!items) {
            throw new Error('Failed to get items collection');
        }
        return items.invoke('Item', index);
    }

    getCount() {
        const items = this.getItems();
        if (!items) {
            throw new Error('Failed to get items collection');
        }
        return items.getProperty('Count');
    }

    /**
     * Get folders collection
     * @returns {COMObject} Folders collection
     */
    getFolders() {
        return this.folder.getProperty('Folders');
    }

    /**
     * Get subfolder by name
     * @param {string} name - Folder name
     * @returns {Folder} Subfolder
     */
    getSubfolder(name) {
        const folders = this.getFolders();
        const subfolder = folders.invoke('Item', name);
        return new Folder(subfolder);
    }

    /**
     * Delete folder
     */
    delete() {
        this.folder.invoke('Delete');
    }

    /**
     * Search items by subject
     * @param {string} subject - Subject to search for
     * @returns {Array} Array of matching items
     */
    searchBySubject(subject) {
        const items = this.getItems();
        const filter = `[Subject] = "${subject}"`;
        const results = items.invoke('Restrict', filter);
        return results;
    }

    /**
     * Get unread items
     * @returns {COMObject} Filtered items collection
     */
    getUnreadItems() {
        const items = this.getItems();
        const filter = '[UnRead] = true';
        return items.invoke('Restrict', filter);
    }

    /**
     * Get items by date range
     * @param {Date} startDate - Start date
     * @param {Date} endDate - End date
     * @returns {COMObject} Filtered items collection
     */
    getItemsByDateRange(startDate, endDate) {
        const items = this.getItems();
        const start = startDate.toISOString();
        const end = endDate.toISOString();
        const filter = `[ReceivedTime] >= "${start}" AND [ReceivedTime] <= "${end}"`;
        return items.invoke('Restrict', filter);
    }

    /**
     * Sort items
     * @param {string} property - Property to sort by
     * @param {boolean} descending - Sort descending
     */
    sortItems(property, descending = false) {
        const items = this.getItems();
        items.invoke('Sort', `[${property}]`, descending);
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
