/**
 * Comprehensive Outlook Connector Test Suite
 * Tests all enhanced methods matching UiPath capabilities
 */

const { OutlookConnector, OlDefaultFolders, OlImportance } = require('../index.js');

console.log('='.repeat(70));
console.log('COMPREHENSIVE OUTLOOK CONNECTOR TEST SUITE');
console.log('='.repeat(70));

let testsPassed = 0;
let testsFailed = 0;
const testResults = [];

function logTest(testName, passed, error = null) {
    if (passed) {
        console.log(`✓ ${testName}`);
        testsPassed++;
        testResults.push({ test: testName, status: 'PASS' });
    } else {
        console.log(`✗ ${testName}: ${error}`);
        testsFailed++;
        testResults.push({ test: testName, status: 'FAIL', error: error });
    }
}

async function testMailOperations() {
    console.log('\n--- Testing Mail Operations ---');
    const outlook = new OutlookConnector();
    
    try {
        // Test createMailItem
        try {
            const mail = outlook.createMailItem();
            logTest('createMailItem', mail !== null);
            mail.release();
        } catch (e) {
            logTest('createMailItem', false, e.message);
        }
        
        // Test mail properties
        try {
            const mail = outlook.createMailItem();
            mail.setSubject('Test Subject');
            mail.setBody('Test Body');
            mail.setTo('test@example.com');
            mail.setCC('cc@example.com');
            mail.setBCC('bcc@example.com');
            mail.setImportance(OlImportance.olImportanceHigh);
            logTest('Mail properties (set)', true);
            mail.release();
        } catch (e) {
            logTest('Mail properties (set)', false, e.message);
        }
        
        // Test reply/forward
        try {
            const mail = outlook.createMailItem();
            mail.setSubject('Original');
            mail.setBody('Original body');
            // Note: reply/forward require a sent/received item, so we just test the methods exist
            logTest('Mail reply/forward methods', typeof mail.reply === 'function' && typeof mail.forward === 'function');
            mail.release();
        } catch (e) {
            logTest('Mail reply/forward methods', false, e.message);
        }
        
        // Test mark as read/unread
        try {
            const mail = outlook.createMailItem();
            mail.markAsUnread();
            mail.markAsRead();
            logTest('Mark as read/unread', true);
            mail.release();
        } catch (e) {
            logTest('Mark as read/unread', false, e.message);
        }
        
        // Test categories and flags
        try {
            const mail = outlook.createMailItem();
            mail.setCategories('Important, Work');
            mail.setFlagStatus(2); // Flagged
            logTest('Categories and flags', true);
            mail.release();
        } catch (e) {
            logTest('Categories and flags', false, e.message);
        }
        
    } finally {
        outlook.release();
    }
}

async function testFolderOperations() {
    console.log('\n--- Testing Folder Operations ---');
    const outlook = new OutlookConnector();
    
    try {
        // Test getInbox
        try {
            const inbox = outlook.getInbox();
            logTest('getInbox', inbox !== null);
        } catch (e) {
            logTest('getInbox', false, e.message);
        }
        
        // Test getSentMail
        try {
            const sentMail = outlook.getSentMail();
            logTest('getSentMail', sentMail !== null);
        } catch (e) {
            logTest('getSentMail', false, e.message);
        }
        
        // Test getDrafts
        try {
            const drafts = outlook.getDrafts();
            logTest('getDrafts', drafts !== null);
        } catch (e) {
            logTest('getDrafts', false, e.message);
        }
        
        // Test getDeletedItems
        try {
            const deleted = outlook.getDeletedItems();
            logTest('getDeletedItems', deleted !== null);
        } catch (e) {
            logTest('getDeletedItems', false, e.message);
        }
        
        // Test getOutbox
        try {
            const outbox = outlook.getOutbox();
            logTest('getOutbox', outbox !== null);
        } catch (e) {
            logTest('getOutbox', false, e.message);
        }
        
        // Test getContacts
        try {
            const contacts = outlook.getContacts();
            logTest('getContacts', contacts !== null);
        } catch (e) {
            logTest('getContacts', false, e.message);
        }
        
        // Test getTasks
        try {
            const tasks = outlook.getTasks();
            logTest('getTasks', tasks !== null);
        } catch (e) {
            logTest('getTasks', false, e.message);
        }
        
        // Test getCalendar
        try {
            const calendar = outlook.getCalendar();
            logTest('getCalendar', calendar !== null);
        } catch (e) {
            logTest('getCalendar', false, e.message);
        }
        
        // Test folder item count
        try {
            const inbox = outlook.getInbox();
            const count = inbox.getCount();
            logTest('Folder getCount', typeof count === 'number');
        } catch (e) {
            logTest('Folder getCount', false, e.message);
        }
        
        // Test folder unread count
        try {
            const inbox = outlook.getInbox();
            const unreadCount = inbox.getUnReadItemCount();
            logTest('Folder getUnReadItemCount', typeof unreadCount === 'number');
        } catch (e) {
            logTest('Folder getUnReadItemCount', false, e.message);
        }
        
    } finally {
        outlook.release();
    }
}

async function testAppointmentOperations() {
    console.log('\n--- Testing Appointment Operations ---');
    const outlook = new OutlookConnector();
    
    try {
        // Test createAppointmentItem
        try {
            const appointment = outlook.createAppointmentItem();
            logTest('createAppointmentItem', appointment !== null);
            appointment.release();
        } catch (e) {
            logTest('createAppointmentItem', false, e.message);
        }
        
        // Test appointment properties
        try {
            const appointment = outlook.createAppointmentItem();
            appointment.setSubject('Meeting');
            appointment.setLocation('Conference Room');
            appointment.setBody('Meeting agenda');
            
            const start = new Date();
            start.setHours(14, 0, 0, 0);
            appointment.setStart(start);
            
            const end = new Date();
            end.setHours(15, 0, 0, 0);
            appointment.setEnd(end);
            
            appointment.setReminderSet(true);
            appointment.setReminderMinutesBeforeStart(15);
            appointment.setAllDayEvent(false);
            
            logTest('Appointment properties', true);
            appointment.release();
        } catch (e) {
            logTest('Appointment properties', false, e.message);
        }
        
        // Test meeting attendees
        try {
            const appointment = outlook.createAppointmentItem();
            appointment.setRequiredAttendees('user1@example.com; user2@example.com');
            appointment.setOptionalAttendees('user3@example.com');
            logTest('Meeting attendees', true);
            appointment.release();
        } catch (e) {
            logTest('Meeting attendees', false, e.message);
        }
        
    } finally {
        outlook.release();
    }
}

async function testContactOperations() {
    console.log('\n--- Testing Contact Operations ---');
    const outlook = new OutlookConnector();
    
    try {
        // Test createContactItem
        try {
            const contact = outlook.createContactItem();
            logTest('createContactItem', contact !== null);
            contact.release();
        } catch (e) {
            logTest('createContactItem', false, e.message);
        }
        
        // Test contact properties
        try {
            const contact = outlook.createContactItem();
            contact.setFirstName('John');
            contact.setLastName('Doe');
            contact.setFullName('John Doe');
            contact.setEmail1Address('john@example.com');
            contact.setEmail2Address('john2@example.com');
            contact.setEmail3Address('john3@example.com');
            contact.setBusinessTelephoneNumber('555-1234');
            contact.setMobileTelephoneNumber('555-5678');
            contact.setHomeTelephoneNumber('555-9012');
            contact.setCompanyName('Acme Corp');
            contact.setJobTitle('Manager');
            contact.setWebPage('https://example.com');
            logTest('Contact properties', true);
            contact.release();
        } catch (e) {
            logTest('Contact properties', false, e.message);
        }
        
    } finally {
        outlook.release();
    }
}

async function testTaskOperations() {
    console.log('\n--- Testing Task Operations ---');
    const outlook = new OutlookConnector();
    
    try {
        // Test createTaskItem
        try {
            const task = outlook.createTaskItem();
            logTest('createTaskItem', task !== null);
            task.release();
        } catch (e) {
            logTest('createTaskItem', false, e.message);
        }
        
        // Test task properties
        try {
            const task = outlook.createTaskItem();
            task.setSubject('Complete Report');
            task.setBody('Quarterly report');
            
            const startDate = new Date();
            task.setStartDate(startDate);
            
            const dueDate = new Date();
            dueDate.setDate(dueDate.getDate() + 7);
            task.setDueDate(dueDate);
            
            task.setStatus(1); // In progress
            task.setPriority(2); // High
            task.setPercentComplete(50);
            task.setOwner('John Doe');
            task.setTotalWork(40);
            task.setActualWork(20);
            
            logTest('Task properties', true);
            task.release();
        } catch (e) {
            logTest('Task properties', false, e.message);
        }
        
        // Test markComplete
        try {
            const task = outlook.createTaskItem();
            task.setSubject('Test Task');
            task.markComplete();
            logTest('Task markComplete', true);
            task.release();
        } catch (e) {
            logTest('Task markComplete', false, e.message);
        }
        
    } finally {
        outlook.release();
    }
}

async function testAttachmentOperations() {
    console.log('\n--- Testing Attachment Operations ---');
    const outlook = new OutlookConnector();
    
    try {
        // Test getAttachmentCount
        try {
            const mail = outlook.createMailItem();
            const count = mail.getAttachmentCount();
            logTest('getAttachmentCount', count === 0);
            mail.release();
        } catch (e) {
            logTest('getAttachmentCount', false, e.message);
        }
        
        // Note: addAttachment, saveAttachment, removeAttachment require actual files
        // These are tested in the examples
        
    } finally {
        outlook.release();
    }
}

async function runAllTests() {
    try {
        await testMailOperations();
        await testFolderOperations();
        await testAppointmentOperations();
        await testContactOperations();
        await testTaskOperations();
        await testAttachmentOperations();
        
        console.log('\n' + '='.repeat(70));
        console.log('TEST SUMMARY');
        console.log('='.repeat(70));
        console.log(`Total Tests: ${testsPassed + testsFailed}`);
        console.log(`✓ Passed: ${testsPassed}`);
        console.log(`✗ Failed: ${testsFailed}`);
        console.log(`Success Rate: ${((testsPassed / (testsPassed + testsFailed)) * 100).toFixed(1)}%`);
        
        if (testsFailed > 0) {
            console.log('\nFailed Tests:');
            testResults.filter(r => r.status === 'FAIL').forEach(r => {
                console.log(`  - ${r.test}: ${r.error}`);
            });
        }
        
        console.log('='.repeat(70));
        
    } catch (error) {
        console.error('\n❌ Test suite error:', error.message);
        console.error(error.stack);
    }
}

runAllTests();
